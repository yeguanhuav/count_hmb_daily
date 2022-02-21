############################## 引导退单并在别处重新下单的情况统计 ##################################
library(glue)
#library(openxlsx)
library(scales)
library(writexl)
library(tidyverse)

count_lured_refund_orders <- function(
  city,
  df_orders,
  self_or_com_or_all = 'all',
  start_time = NA,
  end_time = Sys.time(),
  col_order_item_id = 'order_item_id',
  col_date_created = 'date_created',
  col_policy_order_status = 'policy_order_status',
  col_name = 'name',
  col_id_type = 'id_type',
  col_id_no = 'id_no',
  col_birthday = 'birthday',
  col_sex = 'sex',
  col_company = 'company',
  col_channel_cls1 = 'channel_cls1',
  col_channel_cls2 = 'channel_cls2',
  col_channel_cls3 = 'channel_cls3',
  save_result = FALSE
) {
  sop_dir <- getwd()
  
  # Rename columns ----
  df_orders <- df_orders %>% 
    rename(
      col_order_item_id = !!col_order_item_id,
      col_date_created = !!col_date_created,
      col_policy_order_status = !!col_policy_order_status,
      col_company = !!col_company,
      col_channel_cls1 = !!col_channel_cls1,
      col_channel_cls2 = !!col_channel_cls2,
      col_channel_cls3 = !!col_channel_cls3
    ) %>% 
    unite(
      !!col_name, !!col_id_type, !!col_id_no, !!col_birthday, !!col_sex, 
      col = 'col_insurant_id', sep = '_'
    ) %>% 
    # 去掉多余的列，减少内存使用
    select(contains('col_'))
  
  # 开始日期和结束日期
  if(is.na(start_time)) { start_time <- min(df_orders[['col_date_created']]) }
  start_date <- as.Date(start_time, tz = Sys.timezone())
  end_date <- as.Date(end_time, tz = Sys.timezone())
  
  # refund_status
  if(self_or_com_or_all == 'self') {
    order_type <- '个人'
    refund_status <- '04'
  } else if(self_or_com_or_all == 'com') {
    order_type <- '企业'
    refund_status <- '05'
  } else if(self_or_com_or_all == 'all') {
    order_type <- '个人和企业'
    refund_status <- c('04', '05')
  } else {
    stop(paste0(" 'self_or_com_or_all' should be one of c('self', 'com', 'all') "))
  }
  
  no_refund_orders <- df_orders %>% filter(!col_policy_order_status %in% refund_status)
  
  refund_orders <- df_orders %>% filter(col_policy_order_status %in% refund_status) %>% 
    group_by(col_insurant_id) %>% 
    arrange(col_date_created) %>% 
    slice(1) %>% 
    ungroup()
  
  # 创建excel workbook
  # wb <- createWorkbook()
  # pct <- createStyle(numFmt = "0.0%")
  output_list <- list()
  
  # Sheet1: 按【保司】统计总投保量、引导退单量和占比、总退单量、被引导退单量和占比 ----
  # 按公司统计总退单量
  refund_orders_company <- refund_orders %>% 
    group_by(col_company) %>% 
    mutate(total_refund = length(col_order_item_id)) %>% 
    ungroup()
  
  # 所有退单中，被引导退单的订单
  lured_refund_orders1 <- refund_orders_company %>% 
    # 找到既有有效订单又有退单的客户
    filter(col_insurant_id %in% no_refund_orders$col_insurant_id) %>% 
    # 添加上有效订单的信息
    left_join(
      no_refund_orders %>% 
        select(col_insurant_id, col_company2 = col_company, 
               col_date_created2 = col_date_created) %>% 
        distinct()
    ) %>% 
    # 找到有效订单购买时间比退单时间晚的订单（即有效订单购买时间比退单时间大）
    filter(col_date_created2 >= col_date_created) %>% 
    # 找到退单和有效订单的保司不一样的订单（即先退单后购买且保司不一样的订单）
    filter(col_company != col_company2)
  
  # 每个公司被引导退单的订单
  lured_refund_orders_company <- lured_refund_orders1 %>% 
    group_by(col_company) %>% 
    mutate(lured_refund = length(col_order_item_id)) %>% 
    ungroup() %>% 
    select(col_company, total_refund, lured_refund) %>% 
    distinct() %>% 
    # 计算占比
    mutate(lured_refund_percent = round(lured_refund / total_refund, 4))
  
  # 每个公司通过引导退单获得的订单
  lured_orders_company <- lured_refund_orders1 %>% 
    group_by(col_company2) %>% 
    mutate(lured_orders = length(col_order_item_id)) %>% 
    ungroup() %>% 
    select(col_company = col_company2, lured_orders) %>% 
    distinct() %>% 
    # 添加总投保量
    left_join(
      df_orders %>% 
        group_by(col_company) %>% 
        mutate(total_orders = length(col_order_item_id)) %>% 
        ungroup() %>% 
        select(col_company, total_orders) %>% 
        distinct()
    ) %>% 
    # 计算占比
    mutate(lured_orders_percent = round(lured_orders / total_orders, 4))
  
  sheet1 <- full_join(lured_refund_orders_company, lured_orders_company) %>% 
    arrange(desc(total_orders)) %>% 
    rename(保司 = col_company, 
             总投保量 = total_orders, 引导退单量 = lured_orders, 
             引导退单量占比 = lured_orders_percent,
             总退单量 = total_refund, 被引导退单量 = lured_refund, 
             被引导退单量占比 = lured_refund_percent) %>% 
    select(保司, 总投保量, 引导退单量, 引导退单量占比, 总退单量, 被引导退单量, 被引导退单量占比)
  
  # Sheet 1
  # sht1 <- addWorksheet(wb, sheetName = paste0('按公司统计引导退单量-', order_type, '订单'))
  # # Add data to the worksheet we just created
  # writeData(wb, sheet = sht1, x = sheet1)
  # # Add the percent style to the desired cells
  # addStyle(
  #   wb = wb, sheet = sht1, style = pct, 
  #   cols = (1:ncol(sheet1))[str_detect(colnames(sheet1), '占比|率')],
  #   rows = 1:(nrow(sheet1) + 1),
  #   gridExpand = T
  # )
  output_list[[paste0('按公司统计引导退单量-', order_type, '订单')]] <- sheet1
  
  # Sheet2: 按【各级架构】统计总投保量，引导退单量及占比 ----
  # 保司排序
  company_order <- sheet1$保司[sheet1$保司 != '线上平台'] %>% na.omit()
  
  # 按三级渠道统计总退单量
  refund_orders_channel <- refund_orders %>% 
    group_by(col_channel_cls3) %>% 
    mutate(total_refund = length(col_order_item_id)) %>% 
    ungroup()
  
  # 所有退单中，被引导退单的订单
  lured_refund_orders2 <- refund_orders_channel %>% 
    # 找到既有购买又有退单的客户
    filter(col_insurant_id %in% no_refund_orders$col_insurant_id) %>% 
    # 添加上购买的订单
    left_join(
      no_refund_orders %>% 
        select(col_insurant_id, col_channel_cls3_2 = col_channel_cls3, 
               col_date_created2 = col_date_created) %>% 
        distinct()
    ) %>% 
    # 找到购买时间比退单时间大的订单，既购买时间比退单时间晚
    filter(col_date_created2 >= col_date_created) %>% 
    # 找到退单和购买的三级渠道不一样的订单，既先退单后购买，且三级渠道不一样
    filter(col_channel_cls3 != col_channel_cls3_2)
  
  # 每个三级渠道被引导退单的订单
  lured_refund_orders_channel <- lured_refund_orders2 %>% 
    group_by(col_channel_cls3) %>% 
    mutate(lured_refund = length(col_order_item_id)) %>% 
    ungroup() %>% 
    select(col_company, col_channel_cls1, col_channel_cls2, col_channel_cls3, 
           total_refund, lured_refund) %>% 
    distinct() %>% 
    # 计算占比
    mutate(lured_refund_percent = round(lured_refund / total_refund, 4))
  
  # 每个三级渠道通过引导退单获得的订单
  lured_orders_channel <- lured_refund_orders2 %>% 
    group_by(col_channel_cls3_2) %>% 
    mutate(lured_orders = length(col_order_item_id)) %>% 
    ungroup() %>% 
    select(col_channel_cls3 = col_channel_cls3_2, lured_orders) %>% 
    distinct() %>% 
    # 添加总投保量
    left_join(
      df_orders %>% 
        group_by(col_channel_cls3) %>% 
        mutate(total_orders = length(col_order_item_id)) %>% 
        ungroup() %>% 
        select(col_channel_cls3, total_orders) %>% 
        distinct()
    ) %>% 
    # 计算占比
    mutate(lured_orders_percent = round(lured_orders / total_orders, 4))
  
  sheet2 <- full_join(lured_refund_orders_channel, lured_orders_channel) %>% 
    rename(
      保司 = col_company, 一级渠道 = col_channel_cls1, 二级渠道 = col_channel_cls2, 
      三级渠道 = col_channel_cls3, 总投保量 = total_orders, 引导退单量 = lured_orders, 
      引导退单量占比 = lured_orders_percent
    ) %>% 
    select(保司, 一级渠道, 二级渠道, 三级渠道, 总投保量, 引导退单量, 引导退单量占比) %>% 
    distinct() %>% 
    # 去掉线上平台的三级架构
    filter(保司 != '线上平台') %>% 
    # 将没有总投保量的值替换为0
    mutate(总投保量 = replace_na(总投保量, 0)) %>% 
    # 去掉总投保量为0的三级架构
    filter(总投保量 != 0) %>% 
    mutate(保司 = factor(保司, levels = company_order)) %>% 
    group_by(保司, 一级渠道, 二级渠道) %>% 
    arrange(desc(总投保量), .by_group = TRUE) %>% 
    ungroup()
  
  # Sheet 2
  # sht2 <- addWorksheet(wb, sheetName = paste0('按三级渠道统计引导退单量-', order_type, '订单'))
  # # Add data to the worksheet we just created
  # writeData(wb, sheet = sht2, x = sheet2)
  # # Add the percent style to the desired cells
  # addStyle(
  #   wb = wb, sheet = sht2, style = pct, 
  #   cols = (1:ncol(sheet2))[str_detect(colnames(sheet2), '占比|率')],
  #   rows = 1:(nrow(sheet2) + 1),
  #   gridExpand = T
  # )
  output_list[[paste0('按三级渠道统计引导退单量-', order_type, '订单')]] <- sheet2
  
  # if(save_result) {
  #   # 创建文件夹
  #   if( !dir.exists(paste0(sop_dir, '/', city))) {
  #     dir.create(paste0(sop_dir, '/', city))
  #   }
  #   if( !dir.exists(paste0(sop_dir, '/', city, '/', as.Date(end_time, tz = Sys.timezone()))) ) {
  #     dir.create(paste0(sop_dir, '/', city, '/', as.Date(end_time, tz = Sys.timezone())))
  #   }
  #   # 保存Excel
  #   saveWorkbook(
  #     wb, 
  #     file = paste0(
  #       sop_dir, '/', city, '/', end_date, '/',
  #       '引导退单量统计-', order_type, '订单-截止至',
  #       lubridate::year(end_time), '年', 
  #       lubridate::month(end_time), '月', 
  #       lubridate::day(end_time), '日', 
  #       lubridate::hour(end_time), '时',
  #       '.xlsx'
  #     ), 
  #     overwrite = T
  #   )
  # }
  return(output_list)
}
