###################################### 每日投保量分组统计 #########################################
library(glue)
library(scales)
library(writexl)
#library(openxlsx)
library(lubridate)
library(tidyverse)

count_daily_orders <- function(
  city,
  df_orders,
  self_or_com_or_all = 'all', # or 'self' or 'com'
  start_time = NA,
  end_time = Sys.time(),
  col_order_item_id = 'order_item_id',
  col_date_created = 'date_created',
  col_date_updated = 'date_updated',
  col_policy_order_status = 'policy_order_status',
  col_company = 'company',
  col_channel_cls1 = 'channel_cls1',
  col_channel_cls2 = 'channel_cls2',
  col_channel_cls3 = 'channel_cls3',
  col_applicant_id_no = 'applicant_id_no',
  col_open_id = 'open_id',
  col_com_id = 'com_id',
  col_automatic_deduction = 'automatic_deduction',
  sop_datetime = NA,
  no_refund = FALSE, # developing, not usable
  count_automatic_deduction = TRUE,
  plot_orders_and_users = FALSE,
  plot_user_insured_number = FALSE,
  plot_width = 1200,
  plot_height = 900,
  font = NA,
  save_result = FALSE
) {
  sop_dir <- getwd()
  if(is.na(sop_datetime)) {
    sop_datetime <- paste0(
      lubridate::year(end_time), '年', 
      lubridate::month(end_time), '月', 
      lubridate::day(end_time), '日', 
      lubridate::hour(end_time), '时'
    )
  }
  
  # Rename columns ----
  df_orders <- df_orders %>% 
    rename(
      col_order_item_id = !!col_order_item_id,
      col_date_created = !!col_date_created,
      col_company = !!col_company,
      col_channel_cls1 = !!col_channel_cls1,
      col_channel_cls2 = !!col_channel_cls2,
      col_channel_cls3 = !!col_channel_cls3
    ) %>% 
    mutate(
      col_date_created = as.Date(col_date_created, tz = Sys.timezone())
    )
  # col_open_id
  if(self_or_com_or_all == 'self') {
    
    df_orders <- df_orders %>% 
      rename(
        col_open_id = !!col_open_id, 
        col_applicant_id_no = !!col_applicant_id_no,
        col_automatic_deduction = !!col_automatic_deduction
      )
    
  } else if(self_or_com_or_all == 'com') {
    
    df_orders <- df_orders %>% 
      rename(col_open_id = !!col_com_id) %>% # 公司当做投保用户
      mutate(col_applicant_id_no = col_open_id) # 公司当做投保人
    count_automatic_deduction <- FALSE # 企业订单默认没有自动扣费
    
  } else if(self_or_com_or_all == 'all') {
    
    df_orders$col_open_id <- ifelse(
      !is.na(df_orders[[col_com_id]]),
      df_orders[[col_com_id]], # 如果com_id不是空就用com_id
      df_orders[[col_open_id]] # 其他情况一律用open_id
    )
    df_orders$col_applicant_id_no <- ifelse(
      !is.na(df_orders[[col_com_id]]),
      df_orders[[col_com_id]], # 如果com_id不是空就用com_id
      df_orders[[col_applicant_id_no]] # 其他情况一律用col_applicant_id_no
    )
    
  }
  # col_date_updated
  if(no_refund) {
    df_orders <- df_orders %>% mutate(col_date_updated = col_date_created)
  } else {
    df_orders <- df_orders %>% 
      rename(
        col_date_updated = !!col_date_updated,
        col_policy_order_status = !!col_policy_order_status
      ) %>% 
      mutate(
        col_date_updated = as.Date(col_date_updated, tz = Sys.timezone())
      )
  }
  # 去掉多余的列，减少内存使用
  df_orders <- df_orders %>% select(contains('col_'))
  # 开始日期和结束日期
  if(is.na(start_time)) { start_time <- min(df_orders[['col_date_created']]) }
  start_date <- as.Date(start_time, tz = Sys.timezone())
  end_date <- as.Date(end_time, tz = Sys.timezone())
  # 按input日期筛选订单（注意input日期格式）
  df_orders <- df_orders %>% filter(col_date_created >= start_date & col_date_created <= end_date)
  
  # refund_status ----
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
  # 有效订单
  no_refund_orders <- df_orders %>% filter(!col_policy_order_status %in% !!refund_status)
  # 退保订单
  refund_orders <- df_orders %>% 
    # 统一日期
    mutate(日期 = col_date_updated) %>% 
    # 只保留退保订单
    filter(col_policy_order_status %in% refund_status) %>% 
    # 计算每日总退保量
    group_by(日期) %>% 
    mutate(总退保量 = length(col_order_item_id)) %>% 
    ungroup() %>% 
    # 计算保司的每日退保量
    group_by(日期, col_company) %>% 
    mutate(保司退保量 = length(col_order_item_id)) %>% 
    ungroup() %>% 
    # 计算一级渠道的每日退保量
    group_by(日期, col_company, col_channel_cls1) %>% 
    mutate(一级渠道退保量 = length(col_order_item_id)) %>% 
    ungroup() %>% 
    # 计算二级渠道的每日退保量
    group_by(日期, col_company, col_channel_cls1, col_channel_cls2) %>% 
    mutate(二级渠道退保量 = length(col_order_item_id)) %>% 
    ungroup()
  
  # # openxlsx
  # # https://stackoverflow.com/a/57279302/10341233
  # # Formatting percentages to excel in R
  # # https://stackoverflow.com/a/48066298/10341233
  # # Create an Excel workbook object and add a worksheet
  # wb <- createWorkbook()
  # # Create a percent style
  # pct <- createStyle(numFmt = "0.0%")
  
  # writexl ----
  output_list <- list()
  
  # Sheet1,2,3: 计算每日净投保量 ----
  df_orders <- df_orders %>% mutate(日期 = col_date_created)
  
  # Sheet1：每日总量
  # 总每日净投保量
  count_by_col_date <- df_orders %>% 
    group_by(日期) %>% 
    mutate(
      总投保用户 = length(unique(col_open_id)),
      总投保人 = length(unique(col_applicant_id_no)),
      总投保量 = length(col_order_item_id)
    ) %>% 
    ungroup() %>% 
    select(日期, 总投保用户, 总投保人, 总投保量) %>% 
    distinct() %>% 
    full_join(refund_orders %>% select(日期, 总退保量) %>% distinct(), by = c('日期' = '日期')) %>% 
    # 替换NA为0
    mutate(across(where(is.numeric), ~ replace_na(., 0))) %>% 
    mutate(总净投保量 = 总投保量 - 总退保量) %>% 
    distinct() %>% 
    arrange(日期)
  # # Sheet1
  # sht1 <- addWorksheet(wb, sheetName = paste0('每日投保量统计-', order_type, '订单'))
  # # Add data to the worksheet we just created
  # writeData(wb, sheet = sht1, x = count_by_col_date)
  # # Add the percent style to the desired cells
  # addStyle(
  #   wb = wb, sheet = sht1, style = pct, 
  #   cols = (1:ncol(count_by_col_date))[
  #     str_detect(colnames(count_by_col_date), '占比|率')
  #   ],
  #   rows = 1:(nrow(count_by_col_date) + 1),
  #   gridExpand = T
  # )
  # Last row
  last_row <- tibble(
    日期 = '总计',
    总投保用户 = length(unique(df_orders$col_open_id)),
    总投保人 = length(unique(df_orders$col_applicant_id_no)),
    总投保量 = length(df_orders$col_order_item_id),
    总退保量 = length(refund_orders$col_order_item_id),
    总净投保量 = 总投保量 - 总退保量
  )
  count_by_col_date <- count_by_col_date %>% 
    mutate(日期 = as.character(日期)) %>% 
    bind_rows(last_row)
  # 如果是个人订单则统计开通自动扣费的人数（无退单）
  if(self_or_com_or_all == 'self') {
    automatic_deduction <- no_refund_orders %>% 
      group_by(col_date_created) %>% 
      mutate(开通自动扣费 = sum(col_automatic_deduction)) %>% 
      ungroup() %>% 
      select(日期 = col_date_created, 开通自动扣费) %>% 
      distinct() %>% 
      arrange(日期)
    last_row <- tibble(日期 = '总计', 开通自动扣费 = sum(no_refund_orders$col_automatic_deduction))
    automatic_deduction <- automatic_deduction %>% 
      mutate(日期 = as.character(日期)) %>% 
      bind_rows(last_row)
    # 添加到每日投保量统计
    count_by_col_date <- count_by_col_date %>% 
      full_join(automatic_deduction) %>% 
      mutate(开通自动扣费 = replace_na(开通自动扣费, 0))
  }
  # Sheet1
  output_list[[paste0('每日投保量统计-', order_type, '订单')]] <- count_by_col_date
  
  # Sheet2：每家保司每日投保量
  # 按保司分组计算每日净投保量
  count_by_col_company <- df_orders %>% 
    # 计算每家保司的每日净投保量
    group_by(日期, col_company) %>% 
    mutate(
      保司投保用户 = length(unique(col_open_id)),
      保司投保人 = length(unique(col_applicant_id_no)),
      保司投保量 = length(col_order_item_id)
    ) %>% 
    ungroup() %>% 
    select(日期, 保险公司 = col_company, 保司投保用户, 保司投保人, 保司投保量) %>% 
    distinct() %>% 
    full_join(
      refund_orders %>% select(日期, col_company, 保司退保量) %>% distinct(), 
      by = c('日期' = '日期', '保险公司' = 'col_company')
    ) %>% 
    # 替换NA为0
    mutate(across(where(is.numeric), ~ replace_na(., 0))) %>% 
    mutate(保司净投保量 = 保司投保量 - 保司退保量) %>% 
    distinct()
  # 保司按投保量排序
  company_order <- count_by_col_company %>% 
    select(保险公司, 保司投保量) %>% 
    group_by(保险公司) %>% 
    mutate(total_orders = sum(保司投保量)) %>% 
    ungroup() %>% 
    select(-保司投保量) %>% 
    distinct() %>% 
    arrange(desc(total_orders)) %>% 
    .$保险公司 %>% 
    na.omit()
  # 每家保司的last row
  count_by_col_company_output <- count_by_col_company[1,] %>% .[-1,] %>% 
    mutate(日期 = as.character(日期))
  for(i in company_order) {
    tmp_count_by_col_company <- count_by_col_company %>% 
      filter(保险公司 == !!i) %>% 
      arrange(日期)
    tmp_df_orders <- filter(df_orders, col_company == !!i)
    tmp_refund_orders <- filter(refund_orders, col_company == !!i)
    # last row
    last_row <- tibble(
      日期 = '合计',
      保险公司 = i,
      保司投保用户 = length(unique(tmp_df_orders$col_open_id)),
      保司投保人 = length(unique(tmp_df_orders$col_applicant_id_no)),
      保司投保量 = length(tmp_df_orders$col_order_item_id),
      保司退保量 = length(tmp_refund_orders$col_order_item_id),
      保司净投保量 = 保司投保量 - 保司退保量
    )
    tmp_count_by_col_company <- tmp_count_by_col_company %>% 
      mutate(日期 = as.character(日期)) %>% 
      bind_rows(last_row)
    count_by_col_company_output <- bind_rows(count_by_col_company_output, tmp_count_by_col_company)
  }
  # # 每家保司的每日投保量（宽表格）
  # count_by_col_company_output <- count_by_col_company[1,0] %>% .[-1,]
  # for(i in company_order) {
  #  count_by_col_company_tmp <- count_by_col_company %>%
  #    select(-总投保用户, -总投保人, -总投保量, -总净投保量) %>%
  #    filter(保险公司 == !!i) %>%
  #    pivot_wider(names_from = 保险公司,
  #                values_from = c(保司投保用户, 保司投保人, 保司投保量, 保司净投保量),
  #                names_glue = "{保险公司}{.value}")
  #  count_by_col_company_output <- full_join(count_by_col_company, count_by_col_company_tmp)
  # }
  # 每家保司的每日投保量（长表格）
  count_by_col_company_output <- count_by_col_company_output %>% 
    mutate(保险公司 = factor(保险公司, levels = company_order)) %>% 
    arrange(保险公司, 日期)
  # 如果是个人订单则添加自动扣费统计
  if(self_or_com_or_all == 'self') {
    # 按保司分组计算每日开通自动扣费人数（无退单）
    automatic_deduction <- no_refund_orders %>% 
      group_by(col_date_created, col_company) %>% 
      mutate(开通自动扣费 = sum(col_automatic_deduction)) %>% 
      ungroup() %>% 
      select(日期 = col_date_created, 保险公司 = col_company, 开通自动扣费) %>% 
      distinct()
    # 每家保司的last row
    automatic_deduction_output <- automatic_deduction[1,] %>% .[-1,] %>% 
      mutate(日期 = as.character(日期))
    for(i in company_order) {
      tmp_automatic_deduction <- automatic_deduction %>% 
        filter(保险公司 == !!i) %>% 
        arrange(日期)
      tmp_no_refund_orders <- filter(no_refund_orders, col_company == !!i)
      # last row
      last_row <- tibble(
        日期 = '合计',
        保险公司 = i,
        开通自动扣费 = sum(tmp_no_refund_orders$col_automatic_deduction)
      )
      tmp_automatic_deduction <- tmp_automatic_deduction %>% 
        mutate(日期 = as.character(日期)) %>% 
        bind_rows(last_row)
      automatic_deduction_output <- bind_rows(automatic_deduction_output, tmp_automatic_deduction)
    }
    # 添加到按保司统计-个人订单中
    count_by_col_company_output <- count_by_col_company_output %>% 
      full_join(automatic_deduction_output) %>% 
      mutate(开通自动扣费 = replace_na(开通自动扣费, 0))
  }
  # # Sheet2
  # sht2 <- addWorksheet(wb, sheetName = paste0('按保司统计-', order_type, '订单'))
  # # Add data to the worksheet we just created
  # writeData(wb, sheet = sht2, x = count_by_col_company_output)
  # # Add the percent style to the desired cells
  # addStyle(
  #   wb = wb, sheet = sht2, style = pct, 
  #   cols = (1:ncol(count_by_col_company_output))[
  #     str_detect(colnames(count_by_col_company_output), '占比|率')
  #   ],
  #   rows = 1:(nrow(count_by_col_company_output) + 1),
  #   gridExpand = T
  # )
  # Sheet2
  output_list[[paste0('按保司统计-', order_type, '订单')]] <- count_by_col_company_output
  
  # Sheet3：每家保司每个一级渠道每日投保量
  # 按保司一级渠道分组计算每日净投保量
  count_by_col_channel_cls1 <- df_orders %>% 
    # 计算每家保司的每日净投保量
    group_by(日期, col_company, col_channel_cls1) %>% 
    mutate(
      一级渠道投保用户 = length(unique(col_open_id)),
      一级渠道投保人 = length(unique(col_applicant_id_no)),
      一级渠道投保量 = length(col_order_item_id)
    ) %>% 
    ungroup() %>% 
    select(
      日期, 保险公司 = col_company, 一级渠道 = col_channel_cls1, 一级渠道投保用户, 一级渠道投保人, 
      一级渠道投保量
    ) %>% 
    distinct() %>% 
    full_join(
      refund_orders %>% select(日期, col_company, col_channel_cls1, 一级渠道退保量) %>% distinct(), 
      by = c('日期' = '日期', '保险公司' = 'col_company', '一级渠道' = 'col_channel_cls1')
    ) %>% 
    # 替换NA为0
    mutate(across(where(is.numeric), ~ replace_na(., 0))) %>% 
    mutate(一级渠道净投保量 = 一级渠道投保量 - 一级渠道退保量) %>% 
    distinct()
  # 每家保司每个一级渠道的last row
  count_by_col_channel_cls1_output <- count_by_col_channel_cls1[1,] %>% .[-1,] %>% 
    mutate(日期 = as.character(日期))
  for(i in company_order) {
    all_channel_cls1 <- df_orders %>% 
      filter(col_company == !!i) %>% 
      .$col_channel_cls1 %>% 
      unique() %>% 
      na.omit()
    for (j in all_channel_cls1) {
      tmp_count_by_col_channel_cls1 <- count_by_col_channel_cls1 %>% 
        filter(保险公司 == !!i, 一级渠道 == !!j) %>% 
        arrange(日期)
      tmp_df_orders <- filter(df_orders, col_company == !!i, col_channel_cls1 == !!j)
      tmp_refund_orders <- filter(refund_orders, col_company == !!i, col_channel_cls1 == !!j)
      # last row
      last_row <- tibble(
        日期 = '合计',
        保险公司 = i,
        一级渠道 = j,
        一级渠道投保用户 = length(unique(tmp_df_orders$col_open_id)),
        一级渠道投保人 = length(unique(tmp_df_orders$col_applicant_id_no)),
        一级渠道投保量 = length(tmp_df_orders$col_order_item_id),
        一级渠道退保量 = length(tmp_refund_orders$col_order_item_id),
        一级渠道净投保量 = 一级渠道投保量 - 一级渠道退保量
      )
      tmp_count_by_col_channel_cls1 <- tmp_count_by_col_channel_cls1 %>% 
        mutate(日期 = as.character(日期)) %>% 
        bind_rows(last_row)
      count_by_col_channel_cls1_output <- bind_rows(
        count_by_col_channel_cls1_output, tmp_count_by_col_channel_cls1
      )
    }
  }
  count_by_col_channel_cls1_output <- count_by_col_channel_cls1_output %>% 
    mutate(保险公司 = factor(保险公司, levels = company_order)) %>% 
    arrange(保险公司, 一级渠道, 日期)
  # # Sheet3
  # sht3 <- addWorksheet(wb, sheetName = paste0('按一级渠道统计-', order_type, '订单'))
  # # Add data to the worksheet we just created
  # writeData(wb, sheet = sht3, x = count_by_col_channel_cls1_output)
  # # Add the percent style to the desired cells
  # addStyle(
  #   wb = wb, sheet = sht3, style = pct, 
  #   cols = (1:ncol(count_by_col_channel_cls1_output))[
  #     str_detect(colnames(count_by_col_channel_cls1_output), '占比|率')
  #   ],
  #   rows = 1:(nrow(count_by_col_channel_cls1_output) + 1),
  #   gridExpand = T
  # )
  # Sheet3
  output_list[[paste0('按一级渠道统计-', order_type, '订单')]] <- count_by_col_channel_cls1_output
  
  # Sheet4：每家保司每个一级渠道每日投保量
  # 按保司一级渠道分组计算每日净投保量
  count_by_col_channel_cls2 <- df_orders %>% 
    # 计算每家保司的每日净投保量
    group_by(日期, col_company, col_channel_cls1, col_channel_cls2) %>% 
    mutate(
      二级渠道投保用户 = length(unique(col_open_id)),
      二级渠道投保人 = length(unique(col_applicant_id_no)),
      二级渠道投保量 = length(col_order_item_id)
    ) %>% 
    ungroup() %>% 
    select(
      日期, 保险公司 = col_company, 一级渠道 = col_channel_cls1, 二级渠道 = col_channel_cls2, 
      二级渠道投保用户, 二级渠道投保人, 二级渠道投保量
    ) %>% 
    distinct() %>% 
    full_join(
      refund_orders %>% 
        select(日期, col_company, col_channel_cls1, col_channel_cls2, 二级渠道退保量) %>% 
        distinct(), 
      by = c('日期' = '日期', '保险公司' = 'col_company', '一级渠道' = 'col_channel_cls1',
             '二级渠道' = 'col_channel_cls2')
    ) %>% 
    # 替换NA为0
    mutate(across(where(is.numeric), ~ replace_na(., 0))) %>% 
    mutate(二级渠道净投保量 = 二级渠道投保量 - 二级渠道退保量) %>% 
    distinct()
  # 每家保司每个二级渠道的last row
  count_by_col_channel_cls2_output <- count_by_col_channel_cls2[1,] %>% .[-1,] %>% 
    mutate(日期 = as.character(日期))
  for(i in company_order) {
    all_channel_cls1 <- df_orders %>% 
      filter(col_company == !!i) %>% 
      .$col_channel_cls1 %>% 
      unique() %>% 
      na.omit()
    for(j in all_channel_cls1) {
      all_channel_cls2 <- df_orders %>% 
        filter(col_company == !!i, col_channel_cls1 == !!j) %>% 
        .$col_channel_cls2 %>% 
        unique() %>% 
        na.omit()
      for(k in all_channel_cls2) {
        tmp_count_by_col_channel_cls2 <- count_by_col_channel_cls2 %>% 
          filter(保险公司 == !!i, 一级渠道 == !!j, 二级渠道 == !!k) %>% 
          arrange(日期)
        tmp_df_orders <- df_orders %>% 
          filter(col_company == !!i, col_channel_cls1 == !!j, col_channel_cls2 == !!k)
        tmp_refund_orders <- refund_orders %>% 
          filter(col_company == !!i, col_channel_cls1 == !!j, col_channel_cls2 == !!k)
        # last row
        last_row <- tibble(
          日期 = '合计',
          保险公司 = i,
          一级渠道 = j,
          二级渠道 = k,
          二级渠道投保用户 = length(unique(tmp_df_orders$col_open_id)),
          二级渠道投保人 = length(unique(tmp_df_orders$col_applicant_id_no)),
          二级渠道投保量 = length(tmp_df_orders$col_order_item_id),
          二级渠道退保量 = length(tmp_refund_orders$col_order_item_id),
          二级渠道净投保量 = 二级渠道投保量 - 二级渠道退保量
        )
        tmp_count_by_col_channel_cls2 <- tmp_count_by_col_channel_cls2 %>% 
          mutate(日期 = as.character(日期)) %>% 
          bind_rows(last_row)
        count_by_col_channel_cls2_output <- bind_rows(
          count_by_col_channel_cls2_output, tmp_count_by_col_channel_cls2
        )
      }
    }
  }
  count_by_col_channel_cls2_output <- count_by_col_channel_cls2_output %>% 
    mutate(保险公司 = factor(保险公司, levels = company_order)) %>% 
    arrange(保险公司, 一级渠道, 二级渠道, 日期)
  # Sheet4
  output_list[[paste0('按二级渠道统计-', order_type, '订单')]] <- count_by_col_channel_cls2_output
  
  # 输出Excel ----
  # if(save_result) {
  #   # 创建文件夹
  #   if( !dir.exists(paste0(sop_dir, '/', city)) ) {
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
  #       '每日投保量统计-', order_type, '订单-截止至',
  #       lubridate::year(end_time), '年',
  #       lubridate::month(end_time), '月',
  #       lubridate::day(end_time), '日',
  #       lubridate::hour(end_time), '时',
  #       '.xlsx'
  #     ),
  #     overwrite = T
  #   )
  #   write_xlsx(
  #     output_list,
  #     path = paste0(
  #       sop_dir, '/', city, '/', end_date, '/',
  #       '每日投保量统计-', order_type, '订单-截止至',
  #       lubridate::year(end_time), '年',
  #       lubridate::month(end_time), '月',
  #       lubridate::day(end_time), '日',
  #       lubridate::hour(end_time), '时',
  #       '.xlsx'
  #     )
  #   )
  # }
  
  # 参保量与用户数画图 ----
  if(plot_orders_and_users) {
    orders_and_users <- tibble(
      x = c('参保量', '用户量'), 
      y = c(
        length(unique(no_refund_orders$col_order_item_id)), 
        length(unique(no_refund_orders$col_open_id))
      )
    )
    p <- ggplot(data = orders_and_users, aes(x = x, y = y)) + 
      geom_bar(stat = "identity", fill = '#003366') + 
      theme_classic() +
      theme(
        plot.title = element_text(size = 25),# hjust = 0.5), # title居中
        axis.text.x = element_text(size = 20),
        axis.text.y = element_text(size = 20),
        axis.title = element_text(size = 25),
        legend.title = element_blank(),
        legend.text = element_text(size = 20),
        legend.position = 'top'
      ) + 
      labs(
        title = paste0(
          city, '参保量与用户数总览-', order_type, '订单-截止至', 
          lubridate::year(end_time), '年', 
          lubridate::month(end_time), '月', 
          lubridate::day(end_time), '日', 
          lubridate::hour(end_time), '时'
        ), 
        x = '', y = '人数'
      ) +
      geom_text(aes(label = y), vjust = 1.6, color = "white", size = 10) + 
      # 添加客单比
      annotate(
        geom = "text", 
        x = max(orders_and_users$x), y = max(orders_and_users$y), 
        label = glue("客单比：{round(orders_and_users[1,2] / orders_and_users[2,2], 2)}"),
        color = '#003366', size = 10
      ) +
      scale_y_continuous(
        breaks = seq(
          from = 0, 
          to = max(orders_and_users$y), 
          by = signif(max(orders_and_users$y) / 10, 1)
        )
      )
    # 添加字体
    if(!is.na(font)) {
      p <- p + theme(text = element_text(family = font))
    }
    png(
      filename = glue(
        "{sop_dir}/{city}/{sop_datetime}/{city}参保量与用户数总览-{order_type}订单-{sop_datetime}",
        ".png"
      ), 
      width = plot_width, height = plot_height
    )
    grid::grid.draw(p)
    dev.off()
  }
  
  # 用户投保人数分布画图 ----
  if(plot_user_insured_number) {
    user_insured_number <- table(no_refund_orders$col_open_id) %>% as.data.frame()
    colnames(user_insured_number) <- c('col_open_id', 'Freq')
    user_insured_number <- user_insured_number %>% 
      mutate(
        x = case_when(
          Freq == 1 ~ '为1人投保',
          Freq == 2 ~ '为2人投保',
          Freq == 3 ~ '为3人投保',
          Freq == 4 ~ '为4人投保',
          Freq == 5 ~ '为5人投保',
          Freq == 6 ~ '为6人投保',
          Freq == 7 ~ '为7人投保',
          Freq == 8 ~ '为8人投保',
          Freq == 9 ~ '为9人投保',
          TRUE ~ '为>9人投保'
        )
      ) %>% 
      group_by(x) %>% 
      mutate(y = length(col_open_id)) %>% 
      ungroup() %>% 
      select(x, y) %>% distinct() %>% 
      mutate(
        x = factor(
          x, 
          levels = c(
            '为1人投保', '为2人投保', '为3人投保', '为4人投保', '为5人投保', '为6人投保',
            '为7人投保', '为8人投保', '为9人投保', '为>9人投保'
          )
        )
      )
    p <- ggplot(data = user_insured_number, aes(x = x, y = y)) + 
      geom_bar(stat = "identity", fill = '#CC6633') + 
      theme_classic() +
      theme(
        plot.title = element_text(size = 25),# hjust = 0.5), # title居中
        axis.text.x = element_text(size = 20),
        axis.text.y = element_text(size = 20),
        axis.title = element_text(size = 25),
        legend.title = element_blank(),
        legend.text = element_text(size = 20),
        legend.position = 'top'
      ) + 
      labs(
        title = paste0(
          city, '用户投保人数分布-', order_type, '订单-截止至', 
          lubridate::year(end_time), '年', 
          lubridate::month(end_time), '月', 
          lubridate::day(end_time), '日', 
          lubridate::hour(end_time), '时'
        ), 
        x = '', y = '用户数'
      ) +
      geom_text(aes(label = y), vjust = -0.5, color = "black", size = 8) +
      scale_y_continuous(
        breaks = seq(
          from = 0, 
          to = max(user_insured_number$y), 
          by = signif(max(user_insured_number$y) / 10, 1)
        )
      )
    # 添加字体
    if(!is.na(font)) {
      p <- p + theme(text = element_text(family = font))
    }
    png(
      filename = glue(
        "{sop_dir}/{city}/{sop_datetime}/{city}用户投保人数分布-{order_type}订单-{sop_datetime}",
        ".png"
      ), 
      width = plot_width, height = plot_height
    )
    grid::grid.draw(p)
    dev.off()
  }
  
  return(output_list)
}
