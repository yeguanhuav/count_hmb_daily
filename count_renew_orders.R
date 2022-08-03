######################################## 续保订单统计 #############################################
library(bit64)
library(glue)
#library(openxlsx)
library(scales)
library(writexl)
library(lubridate)
library(tidyverse)

count_renew_orders <- function(
  city,
  start_time = NA,
  end_time = Sys.time(),
  # 旧个人订单（不含退单）
  df_orders_self_old, 
  col_id_no_self_old = 'id_no',
  col_automatic_deduction_self_old = 'automatic_deduction', # 该订单是否签约了自动扣费
  col_company_self_old = 'company',
  col_deducted_self_old = NA, # 该订单是否因自动扣费生成了新订单（automatic_deduction_is_success）
  col_medical_insure_flag = NA, # 预约医保个账划扣的标签
  col_pay_way = NA, # 订单支付方式
  # 新个人订单（含退单）
  df_orders_self, 
  col_order_item_id_self = 'order_item_id',
  col_policy_order_status_self = 'policy_order_status',
  col_id_no_self = 'id_no',
  col_sku_code = 'sku_code',
  col_company_self = 'company',
  col_is_automatic_deduction_self = NA, # 该订单是否由自动扣费产生（is_automatic_deduction）
  # 旧企业订单（不含退单）
  df_orders_com_old = NULL, 
  col_company_com_old = 'company',
  col_id_no_com_old = 'id_no',
  # 新企业订单（含退单）
  df_orders_com = NULL, 
  col_order_item_id_com = 'order_item_id',
  col_policy_order_status_com = 'policy_order_status',
  col_company_com = 'company',
  col_id_no_com = 'id_no',
  count_self_renew = TRUE,
  count_com_renew = TRUE,
  count_automatic_deduction = TRUE,
  save_result = FALSE
) {
  sop_dir <- getwd()
  
  # Rename columns ----
  # 旧个人订单
  df_orders_self_old <- df_orders_self_old %>% 
    rename(
      col_id_no_self_old = !!col_id_no_self_old,
      col_automatic_deduction_self_old = !!col_automatic_deduction_self_old,
      col_company_self_old = !!col_company_self_old
    )
  # 标记发起过自动扣费的订单
  if(is.na(col_deducted_self_old)) {
    # 如果没有发起过自动扣费，设为0方便后续统计
    df_orders_self_old <- df_orders_self_old %>% mutate(col_deducted_self_old = 0)
  } else {
    # 如果已经发起过自动扣费，则会标记该订单是否进行过自动扣费
    df_orders_self_old <- df_orders_self_old %>% 
      rename(col_deducted_self_old = !!col_deducted_self_old)
  }
  # 预约医保个账划扣的标签
  if(!is.na(col_medical_insure_flag)) {
    df_orders_self_old <- df_orders_self_old %>% 
      rename(col_medical_insure_flag = !!col_medical_insure_flag)
  }
  # 订单支付方式
  if(!is.na(col_pay_way)) {
    df_orders_self_old <- df_orders_self_old %>% 
      rename(col_pay_way = !!col_pay_way)
  }
  # 去掉多余的列，减少内存使用
  df_orders_self_old <- df_orders_self_old %>% select(contains('col_'))
  
  # 旧企业订单（旧企业订单今年可能在个人订单购买）
  if(is.null(df_orders_com_old) | nrow(df_orders_com_old) == 0) {
    warning(
      glue("Cannot find any old company orders, count_com_renew function has been skipped." )
    )
    count_com_renew <- FALSE
    df_orders_com_old <- tibble(col_company_com_old = 'NA', col_id_no_com_old = 'NA')
  } else {
    df_orders_com_old <- df_orders_com_old %>% 
      rename(
        col_company_com_old = !!col_company_com_old,
        col_id_no_com_old = !!col_id_no_com_old
      ) %>% 
      # 去掉多余的列，减少内存使用
      select(contains('col_'))
  }
  
  # 新个人订单
  df_orders_self <- df_orders_self %>% 
    rename(
      col_order_item_id_self = !!col_order_item_id_self,
      col_policy_order_status_self = !!col_policy_order_status_self,
      col_id_no_self = !!col_id_no_self,
      col_company_self = !!col_company_self
    )
  # 标记自动扣费成功的订单
  if(is.na(col_is_automatic_deduction_self)) {
    # 如果没有发起过自动扣费，设为0方便后续统计
    df_orders_self <- df_orders_self %>% mutate(col_is_automatic_deduction_self = 0)
  } else {
    # 如果已经发起过自动扣费，则有标记显示该订单是否为自动扣费产生的订单
    df_orders_self <- df_orders_self %>% 
      rename(col_is_automatic_deduction_self = !!col_is_automatic_deduction_self)
  }
  # 去掉多余的列，减少内存使用
  df_orders_self <- df_orders_self %>% select(contains('col_'))
  # 自动扣费订单中的退单
  deducted_and_refund <- df_orders_self %>% 
    filter(col_is_automatic_deduction_self == 1) %>% 
    filter(col_policy_order_status_self %in% c('04'))
  # 去掉退保的个人订单
  df_orders_self <- df_orders_self %>% 
    filter(!col_policy_order_status_self %in% c('04'))
  
  # 新企业订单
  if(is.null(df_orders_com) | nrow(df_orders_com) == 0) {
    warning(glue("Cannot find any new company orders."))
    df_orders_com <- tibble(
      col_order_item_id_com = 'NA',
      col_policy_order_status_com = 'NA',
      col_company_com = 'NA',
      col_id_no_com = 'NA'
    )
  } else {
    # 最新的企业订单
    df_orders_com <- df_orders_com %>% 
      rename(
        col_order_item_id_com = !!col_order_item_id_com,
        col_policy_order_status_com = !!col_policy_order_status_com,
        col_company_com = !!col_company_com,
        col_id_no_com = !!col_id_no_com
      ) %>% 
      # 去掉多余的列，减少内存使用
      select(contains('col_')) %>% 
      # 去掉退保订单
      filter(!col_policy_order_status_com %in% c('05'))
  }
  
  #df_all_orders <- bind_rows(df_orders_self, df_orders_com)
  
  # 开始日期和结束日期
  start_date <- as_date(start_time)
  end_date <- as_date(end_time)
  
  # # 创建excel workbook
  # wb <- createWorkbook()
  # pct <- createStyle(numFmt = "0.0%")
  output_list <- list()
  
  # 自动扣费统计 ----
  if(count_automatic_deduction) {
    
    # 今年自动扣费成功且未退款的订单
    ad_orders <- filter(df_orders_self, col_is_automatic_deduction_self == 1)
    
    # 去掉自动扣费成功且未退款的订单
    df_orders_self_left <- df_orders_self %>% 
      filter(!col_order_item_id_self %in% ad_orders$col_order_item_id_self)
    
    # 旧个人订单中自动扣费订单添加标签
    self_automatic_deduction_orders <- df_orders_self_old %>% 
      # 筛选旧订单中签约自动扣费的订单
      filter(col_automatic_deduction_self_old == 1) %>% 
      # 今年企业已购买
      mutate(
        is_renew_com = ifelse(col_id_no_self_old %in% df_orders_com$col_id_no_com, 1, 0)
      ) %>% 
      # 今年个人已购买
      mutate(
        is_renew_self = case_when(
          # 没有企业已购买订单 且 有个人已购买订单
          (is_renew_com == 0) & (col_id_no_self_old %in% df_orders_self_left$col_id_no_self) ~ 1,
          TRUE ~ 0
        )
      ) %>% 
      # 自动扣费成功且未退款的订单（这个结果会比最终成功扣费的少，因为有成功扣费后解约的）
      mutate(
        is_ad = ifelse(col_id_no_self_old %in% ad_orders$col_id_no_self, 1, 0)
      )
    
    # 统计表格
    self_automatic_deduction_orders_sum <- self_automatic_deduction_orders %>% 
      group_by(col_company_self_old) %>% 
      mutate(
        # 开通自动扣费人数
        去年开通自动扣费人数 = length(col_id_no_self_old),
        # 去年签了自动扣费且今年又在企业订单购买了的人数
        今年企业订单已购买人数 = sum(is_renew_com),
        # 去年签了自动扣费且今年又在个人订单购买了的人数
        今年个人订单已购买人数 = sum(is_renew_self),
        # 今年自动扣费成功且无退款的订单
        自动扣费成功且无退款人数 = sum(is_ad)
      ) %>% 
      ungroup() %>% 
      select(
        保险公司 = col_company_self_old,
        去年开通自动扣费人数,
        今年企业订单已购买人数,
        今年个人订单已购买人数,
        自动扣费成功且无退款人数
      ) %>% 
      distinct()
    
    # 统计自动扣费后退款且未购买的订单
    self_automatic_deduction_orders <- self_automatic_deduction_orders %>% 
      filter(
        is_renew_com == 0 & is_renew_self == 0 & is_ad == 0
      ) %>% 
      mutate(
        is_ad_refund = ifelse(col_id_no_self_old %in% deducted_and_refund$col_id_no_self, 1, 0)
      ) %>% 
      group_by(col_company_self_old) %>% 
      mutate(自动扣费后退款且未购买人数 = sum(is_ad_refund))
    # 备份剩余自动扣费名单
    self_automatic_deduction_orders_bak <- self_automatic_deduction_orders %>% 
      filter(is_ad_refund == 0)
    # 统计自动扣费后退款且未购买的订单
    self_automatic_deduction_orders <- self_automatic_deduction_orders %>% 
      select(保险公司 = col_company_self_old, 自动扣费后退款且未购买人数) %>% 
      distinct()
    self_automatic_deduction_orders_sum <- self_automatic_deduction_orders_sum %>% 
      left_join(self_automatic_deduction_orders) %>% 
      mutate(
        剩余自动扣费人数 = 去年开通自动扣费人数 - (
          今年企业订单已购买人数 + 今年个人订单已购买人数 + 
            自动扣费成功且无退款人数 + 自动扣费后退款且未购买人数
        )
      ) %>% 
      mutate(剩余人数占比 = round(剩余自动扣费人数 / 去年开通自动扣费人数, 3))
    
    # 如果有个账预约标记和支付方式，则对自动扣费剩余人数进一步分类，统计自费、个账成功、个账失败
    # 三种客户剩余人数
    if(!is.na(col_medical_insure_flag) & !is.na(col_pay_way)) {
      self_automatic_deduction_orders_bak <- self_automatic_deduction_orders_bak %>% 
        mutate(
          # 标记没有预约医保个账划扣，自费购买的客户
          paid_by_their_own = case_when(
            (col_medical_insure_flag == 0) & (col_pay_way != '医保个账') ~ 1,
            TRUE ~ 0
          ),
          # 标记预约个账划扣，扣费失败，自费购买的客户
          medical_insure_failed = case_when(
            (col_medical_insure_flag == 1) & (col_pay_way != '医保个账') ~ 1,
            TRUE ~ 0
          ),
          # 标记预约个账划扣，扣费成功的客户
          medical_insure_succeeded = case_when(
            (col_medical_insure_flag == 1) & (col_pay_way == '医保个账') ~ 1,
            TRUE ~ 0
          )
        ) %>% 
        group_by(col_company_self_old) %>% 
        mutate(
          未预约个账自费购买的剩余人数 = sum(paid_by_their_own),
          个账购买失败自费购买的剩余人数 = sum(medical_insure_failed),
          个账购买成功的剩余人数 = sum(medical_insure_succeeded)
        ) %>% 
        select(
          保险公司 = col_company_self_old, 
          未预约个账自费购买的剩余人数,
          个账购买失败自费购买的剩余人数,
          个账购买成功的剩余人数
        ) %>% 
        distinct()
      # 加到前面的统计表上
      self_automatic_deduction_orders_sum <- self_automatic_deduction_orders_sum %>% 
        left_join(self_automatic_deduction_orders_bak)
    }
    # 合计
    last_row <- tibble(
      保险公司 = '合计', 
      去年开通自动扣费人数 = sum(self_automatic_deduction_orders_sum$去年开通自动扣费人数), 
      今年个人订单已购买人数 = sum(self_automatic_deduction_orders_sum$今年个人订单已购买人数), 
      今年企业订单已购买人数 = sum(self_automatic_deduction_orders_sum$今年企业订单已购买人数), 
      自动扣费成功且无退款人数 = sum(self_automatic_deduction_orders_sum$自动扣费成功且无退款人数),
      自动扣费后退款且未购买人数 = sum(
        self_automatic_deduction_orders_sum$自动扣费后退款且未购买人数
      ),
      剩余自动扣费人数 = sum(self_automatic_deduction_orders_sum$剩余自动扣费人数), 
      剩余人数占比 = round(
        sum(self_automatic_deduction_orders_sum$剩余自动扣费人数) / 
          sum(self_automatic_deduction_orders_sum$去年开通自动扣费人数), 
        3
      )
    )
    # 如果有个账预约标记和支付方式，则有合计2
    if(!is.na(col_medical_insure_flag) & !is.na(col_pay_way)) {
      last_row2 <- tibble(
        保险公司 = '合计', 
        未预约个账自费购买的剩余人数 = sum(
          self_automatic_deduction_orders_bak$未预约个账自费购买的剩余人数
        ),
        个账购买失败自费购买的剩余人数 = sum(
          self_automatic_deduction_orders_bak$个账购买失败自费购买的剩余人数
        ),
        个账购买成功的剩余人数 = sum(
          self_automatic_deduction_orders_bak$个账购买成功的剩余人数
        )
      )
      last_row <- last_row %>% left_join(last_row2)
    }
    # 最终表格
    self_automatic_deduction_orders_sum <- bind_rows(self_automatic_deduction_orders_sum, last_row)
    
    # Sheet 3
    # sht3 <- addWorksheet(wb, sheetName = paste0('自动扣费剩余量'))
    # # Add data to the worksheet we just created
    # writeData(wb, sheet = sht3, x = self_automatic_deduction_orders_sum)
    # # Add the percent style to the desired cells
    # addStyle(
    #   wb = wb, sheet = sht3, style = pct, 
    #   cols = (1:ncol(self_automatic_deduction_orders_sum))[
    #     str_detect(colnames(self_automatic_deduction_orders_sum), '占比|率')
    #   ],
    #   rows = 1:(nrow(self_automatic_deduction_orders_sum) + 1),
    #   gridExpand = T
    # )
    # 将含有“占比”或“率”的列转换成百分比
    column_to_scale <- colnames(self_automatic_deduction_orders_sum)[
      str_detect(colnames(self_automatic_deduction_orders_sum), '占比|率')
    ]
    for (column in column_to_scale) {
      self_automatic_deduction_orders_sum[[column]] <- scales::label_percent(
        accuracy = 0.1, big.mark = ""
      )(
        as.numeric(self_automatic_deduction_orders_sum[[column]])
      ) %>% 
        str_replace('Inf', '')
    }
    output_list[['自动扣费剩余量']] <- self_automatic_deduction_orders_sum
    
  }
  
  # 个人订单续保统计 ----
  if(count_self_renew) {
    self_renew_orders <- df_orders_self_old %>% 
      mutate(
        is_renew = ifelse(
          (col_id_no_self_old %in% df_orders_self$col_id_no_self) | 
            (col_id_no_self_old %in% df_orders_com$col_id_no_com), 
          1, 
          0
        )
      ) %>% 
      group_by(col_company_self_old) %>% 
      mutate(
        去年个人订单总数 = length(col_id_no_self_old),
        个人订单续保数 = sum(is_renew)
      ) %>% 
      ungroup() %>% 
      select(保险公司 = col_company_self_old, 去年个人订单总数, 个人订单续保数) %>% 
      distinct() %>% 
      mutate(个人订单续保率 = round(个人订单续保数 / 去年个人订单总数, 3))
    last_row <- tibble(
      保险公司 = '合计', 
      去年个人订单总数 = sum(na.omit(self_renew_orders$去年个人订单总数)), 
      个人订单续保数 = sum(na.omit(self_renew_orders$个人订单续保数)),
      个人订单续保率 = round(
        na.omit(sum(self_renew_orders$个人订单续保数)) / 
          na.omit(sum(self_renew_orders$去年个人订单总数)), 
        3
      )
    )
    self_renew_orders <- bind_rows(self_renew_orders, last_row)
    # Sheet 1
    # sht1 <- addWorksheet(wb, sheetName = paste0('个人订单续保量'))
    # # Add data to the worksheet we just created
    # writeData(wb, sheet = sht1, x = self_renew_orders)
    # # Add the percent style to the desired cells
    # addStyle(
    #   wb = wb, sheet = sht1, style = pct, 
    #   cols = (1:ncol(self_renew_orders))[str_detect(colnames(self_renew_orders), '占比|率')],
    #   rows = 1:(nrow(self_renew_orders) + 1),
    #   gridExpand = T
    # )
  }
  
  # 企业订单续保统计 ----
  if(count_com_renew) {
    # 企业订单的续保统计
    com_renew_orders <- df_orders_com_old %>% 
      mutate(
        is_renew = ifelse(
          (col_id_no_com_old %in% df_orders_com$col_id_no_com) | 
            (col_id_no_com_old %in% df_orders_self$col_id_no_self), 
          1, 
          0
        )
      ) %>% 
      group_by(col_company_com_old) %>% 
      mutate(
        去年企业订单总数 = length(col_id_no_com_old),
        企业订单续保数 = sum(is_renew)
      ) %>% 
      ungroup() %>% 
      select(保险公司 = col_company_com_old, 去年企业订单总数, 企业订单续保数) %>% 
      distinct() %>% 
      mutate(企业订单续保率 = round(企业订单续保数 / 去年企业订单总数, 3))
    last_row <- tibble(
      保险公司 = '合计', 
      去年企业订单总数 = sum(na.omit(com_renew_orders$去年企业订单总数)), 
      企业订单续保数 = sum(na.omit(com_renew_orders$企业订单续保数)), 
      企业订单续保率 = round(
        na.omit(sum(com_renew_orders$企业订单续保数)) / 
          na.omit(sum(com_renew_orders$去年企业订单总数)), 
        3
      )
    )
    com_renew_orders <- bind_rows(com_renew_orders, last_row)
    # Sheet 2
    # sht2 <- addWorksheet(wb, sheetName = paste0('企业订单续保量'))
    # # Add data to the worksheet we just created
    # writeData(wb, sheet = sht2, x = com_renew_orders)
    # # Add the percent style to the desired cells
    # addStyle(
    #   wb = wb, sheet = sht2, style = pct, 
    #   cols = (1:ncol(com_renew_orders))[str_detect(colnames(com_renew_orders), '占比|率')],
    #   rows = 1:(nrow(com_renew_orders) + 1),
    #   gridExpand = T
    # )
  } else {
    com_renew_orders <- tibble(
      保险公司 = c(unique(df_orders_self$col_company_self), '合计'),
      去年企业订单总数 = 'NA',
      企业订单续保数 = 'NA',
      企业订单续保率 = 'NA'
    )
  }
  
  renew_orders <- full_join(self_renew_orders, com_renew_orders)
  renew_orders[renew_orders == 'NA'] <- NA
  column_to_scale <- colnames(renew_orders)[str_detect(colnames(renew_orders), '占比|率')]
  for(column in column_to_scale) {
    renew_orders[[column]] <- scales::label_percent(
      accuracy = 0.1, big.mark = ""
    )(
      as.numeric(renew_orders[[column]])
    ) %>% 
      str_replace('Inf', '')
  }
  output_list[['续保情况']] <- renew_orders

  # if(save_result) {
  #   # 创建文件夹
  #   if( !dir.exists(paste0(sop_dir, '/', city)) ) {
  #     dir.create(paste0(sop_dir, '/', city))
  #   }
  #   if( !dir.exists(paste0(sop_dir, '/', city, '/', as_date(end_time))) ) {
  #     dir.create(paste0(sop_dir, '/', city, '/', as_date(end_time)))
  #   }
  #   # 保存Excel
  #   # saveWorkbook(
  #   #   wb,
  #   #   file = paste0(
  #   #     sop_dir, '/', city, '/', end_date, '/',
  #   #     city, '自动扣费剩余量-截止至',
  #   #     lubridate::year(end_time), '年',
  #   #     lubridate::month(end_time), '月',
  #   #     lubridate::day(end_time), '日',
  #   #     lubridate::hour(end_time), '时',
  #   #     '.xlsx'
  #   #   ),
  #   #   overwrite = T
  #   # )
  #   write_xlsx(
  #     output_list,
  #     path = paste0(
  #       sop_dir, '/', city, '/', end_date, '/',
  #       city, '自动扣费剩余量-截止至',
  #       lubridate::year(end_time), '年',
  #       lubridate::month(end_time), '月',
  #       lubridate::day(end_time), '日',
  #       lubridate::hour(end_time), '时',
  #       '.xlsx'
  #     )
  #   )
  # }
  return(output_list)
}
