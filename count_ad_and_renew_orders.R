######################################## 续保订单统计 #############################################
library(bit64)
library(RMariaDB)
library(glue)
#library(openxlsx)
library(scales)
library(writexl)
library(lubridate)
library(tidyverse)

count_ad_and_renew_orders <- function(
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
  count_automatic_deduction = TRUE,
  count_ad_detail = FALSE,
  ad_id = NA,
  count_self_renew_last_year = TRUE,
  count_com_renew_last_year = TRUE,
  count_self_renew_this_year = TRUE,
  count_com_renew_this_year = TRUE,
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
  # 如果df_orders_com_old不为NULL -> 如果df_orders_com_old行数不为0 -> 处理旧企业订单
  if(!is.null(df_orders_com_old)) {
    if(nrow(df_orders_com_old) != 0) {
      df_orders_com_old <- df_orders_com_old %>% 
        rename(
          col_company_com_old = !!col_company_com_old,
          col_id_no_com_old = !!col_id_no_com_old
        ) %>% 
        # 去掉多余的列，减少内存使用
        select(contains('col_'))
    } else {
      warning(
        glue("Cannot find any company orders from last year, count_com_renew_last_year has been ",
             "disabled." )
      )
      count_com_renew_last_year <- FALSE
      df_orders_com_old <- tibble(col_company_com_old = 'NA', col_id_no_com_old = 'NA')
    }
  } else {
    warning(
      glue("Cannot find any company orders from last year, count_com_renew_last_year has been ",
           "disabled." )
    )
    count_com_renew_last_year <- FALSE
    df_orders_com_old <- tibble(col_company_com_old = 'NA', col_id_no_com_old = 'NA')
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
  # 如果df_orders_com不为NULL -> 如果df_orders_com行数不为0 -> 处理企业订单
  if(!is.null(df_orders_com)) {
    if(nrow(df_orders_com) != 0) {
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
    } else {
      warning(
        glue("Cannot find any new company orders this year, count_com_renew_this_year has been ",
             "disabled.")
      )
      count_com_renew_this_year <- FALSE
      df_orders_com <- tibble(
        col_order_item_id_com = 'NA',
        col_policy_order_status_com = 'NA',
        col_company_com = 'NA',
        col_id_no_com = 'NA'
      )
    }
  } else {
    warning(
      glue("Cannot find any new company orders this year, count_com_renew_this_year has been ",
           "disabled.")
    )
    count_com_renew_this_year <- FALSE
    df_orders_com <- tibble(
      col_order_item_id_com = 'NA',
      col_policy_order_status_com = 'NA',
      col_company_com = 'NA',
      col_id_no_com = 'NA'
    )
  }
  
  #df_all_orders <- bind_rows(df_orders_self, df_orders_com)
  
  # 开始日期和结束日期
  start_date <- as.Date(start_time, tz = Sys.timezone())
  end_date <- as.Date(end_time, tz = Sys.timezone())
  
  # # 创建excel workbook
  # wb <- createWorkbook()
  # pct <- createStyle(numFmt = "0.0%")
  output_list <- list()
  
  # Sheet 1：自动扣费统计 ----
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
  
  # Sheet 2：自动扣费的所有异常情况及转化率统计 ----
  if(count_ad_detail) {
    # 数据库
    product <- dbConnect(
      drv = RMariaDB::MariaDB(), 
      user = 'product',
      password = 'hmb2020',
      host = '10.3.130.16',
      port = 3306,
      dbname = 'trddb',
      groups = 'product'
    )
    
    # 自动扣费任务ID
    auto_deduction_base_config <- dbGetQuery(
      product,
      glue(
        "
SELECT
    CAST( id AS CHAR ) id, 
    project_name,
    CAST( advance_deduction_notice_time AS CHAR ) advance_deduction_notice_time,
    CAST( start_deduction_time AS CHAR ) start_deduction_time
FROM
    auto_deduction_base_config
        "
      )
    ) %>% 
      filter(str_detect(project_name, '测试|模拟', negate = T))
    # 如果ad_id的length为1且为NA，则表示没有输入ad_id，则需要根据city获取自动扣费的任务ID
    if(length(ad_id) == 1) {
      if(is.na(ad_id)) {
        if(is.na(city)) {
          stop('"ad_id" is empty, please input "city" to query "ad_id", or input "ad_id" directly.')
        }
        auto_deduction_base_config <- auto_deduction_base_config %>% 
          filter(str_detect(project_name, !!city))
        ad_id = auto_deduction_base_config$id
      } else {
        ad_id = as.character(ad_id)
      }
    }
    all_ad_id <- paste0("'", paste0(ad_id, collapse = "', '"), "'")
    
    # 每一次自动扣费的各项数据统计
    count_ad_query <- dbGetQuery(
      product,
      glue(
        "
# 扣费成功人数
(
	SELECT
		a2.auto_deduction_base_config_id 扣费任务ID,
		'Step 4：扣费结果' AS 扣费阶段,
		'扣费成功' AS 数据类型,
		SUM(a2.deduction_pay_success_total) 人数
	FROM
		auto_deduction_customer_info a2
	WHERE
		a2.auto_deduction_base_config_id IN ({all_ad_id})
	GROUP BY
		a2.auto_deduction_base_config_id
)
UNION
# 开发：筛选出来的数据
# 思派向腾讯发起扣费的基数
(
	SELECT
		a1.auto_deduction_base_config_id 扣费任务ID,
		'Step 1：数据清洗' AS 扣费阶段,
		'扣费基数' AS 数据类型,
		SUM(a1.send_msg_total) 人数
	FROM
		auto_deduction_customer_info a1
	WHERE
		a1.auto_deduction_base_config_id IN ({all_ad_id})
	GROUP BY
		a1.auto_deduction_base_config_id
)
UNION
# 每次自动扣费异常统计
(
	SELECT
		aa.auto_deduction_base_id 扣费任务ID,
		aa.step 扣费阶段,
		aa.exception_describe 数据类型,
		COUNT(DISTINCT aa.order_item_id) 人数
	FROM
		(
			SELECT
				order_item_id,
				auto_deduction_base_id,
				CASE
					WHEN step = 1 THEN 'Step 1：数据清洗'
					WHEN step = 2 THEN 'Step 2：消息发送'
					WHEN step = 3 THEN 'Step 3：发起扣费'
					WHEN step = 4 THEN 'Step 4：扣费结果'
				END step,
				CASE
					WHEN step = 1 AND exception_describe = '商户号错误' THEN
						'签约关系不存在-商户号错误'
					WHEN step = 4 AND exception_describe IN (
						'当前用户账户已被限制交易', 
						'当前银行卡有安全设置无法完成支付，如有疑问，请联系银行确认', 
						'你的支付账户处于冻结状态，已开启保护模式。可以点击“查看解决方法”，申请解除冻结', 
						'当前用户账户暂时无法完成交易，请用户检查账户后再试'
					) THEN
						'账号异常/冻结'
					WHEN step = 4 AND exception_describe IN (
						'银行拒绝该交易，请联系银行客服',
						'您的银行卡号有误，请核对后再试',
						'用户银行卡暂不可用',
						'103 您的银行卡已被冻结，请尝试其他卡',
						'银行卡剩余额度不足，请选择其他支付方式继续支付',
						'该卡已被银行系统锁定，无法继续使用，请携带相关证件前往银行柜台解锁。如有疑问，可联系银行客服确认',
						'你的银行卡绑定状态不存在或者已过期，请解绑后重新绑定或者使用其他卡支付',
						'该银行卡已注销，请尝试其他卡，如有疑问请联系银行客服',
						'可用额度不足，请使用零钱或其他银行卡支付',
						'您的银行卡状态异常，可能未激活、挂失、没收、过期、销户，请咨询银行',
						'当前用户支付方式暂不可用，请用户检查后再试',
						'暂无可用的支付方式,请绑定其它银行卡完成支付',
						'102 银行卡可用余额不足（如信用卡则为可透支额度不足），请核实后再试'
					) THEN
						'余额不足/卡号冻结等支付方式异常'
					WHEN step = 4 AND exception_describe LIKE '%银行%' THEN
						'余额不足/卡号冻结等支付方式异常'
					WHEN step = 4 AND exception_describe = '交易金额或次数超出限制，请检查后再试' THEN
						'交易金额或次数超出限制'
					WHEN step = 4 AND exception_describe = '签约协议不存在，请检查传入的签约协议是否正确' THEN
						'解绑自动扣费-签约协议不存在'
					ELSE
						exception_describe
				END exception_describe
			FROM
				automatic_deduction_exception
			WHERE
				auto_deduction_base_id IN ({all_ad_id})
		) AS aa
	GROUP BY
		aa.auto_deduction_base_id,
		aa.step,
		aa.exception_describe
)
ORDER BY
	扣费任务ID,
	扣费阶段,
	数据类型
        "
      )
    )
    # 更新有数据的扣费任务ID
    ad_id <- as.character(intersect(ad_id, unique(count_ad_query$扣费任务ID)))
    # 转换成宽表格
    count_ad_query <- pivot_wider(count_ad_query, names_from = 扣费任务ID, values_from = 人数)
    
    # 中间数据
    count_ad <- tibble(
      扣费阶段 = c(
        replicate(3, 'Step 1：数据清洗'), 
        'Step 2：消息发送',
        'Step 3：发起扣费',
        replicate(6, 'Step 4：扣费结果')
      ),
      数据类型 = c(
        '扣费基数',
        '白名单',
        '签约关系不存在-商户号错误',
        '有重复下单',
        '有重复下单',
        '交易金额或次数超出限制',
        '余额不足/卡号冻结等支付方式异常',
        '扣费成功',
        '未成功签约',
        '解绑自动扣费-签约协议不存在',
        '账号异常/冻结'
      )
    )
    # 将query数据合并到中间表格，没有值的设为0
    count_ad <- count_ad %>% 
      left_join(count_ad_query) %>% 
      mutate(across(c(!!ad_id), ~ replace_na(., 0)))
    
    # 最终输出表格
    count_ad_output <- tibble(
      扣费阶段 = c(
        replicate(3, '数据清洗'), 
        replicate(2, '消息发送'), 
        replicate(3, '发起扣费'), 
        replicate(3, '发起阶段转化率'), 
        '扣费成功',
        replicate(6, '扣费异常'), 
        replicate(3, '实际扣费转化率')
      ),
      情况 = c(
        '[1] 白名单', 
        '[2] 自动扣费基数', 
        '[3] 签约关系不存在（商户号错误）',
        '消息发送时间',
        '[4] 有重复下单',
        '发起扣费时间',
        '[5] 有重复下单', 
        '[6] 最终发起微信扣费的人数',
        '[6]/[2] 从扣费名单筛选到发起扣费的转化率',
        '[3]/[2] 发起扣费前的解约率', 
        '([4]+[5])/[2] 发起扣费前的自主下单率',
        '[7] 微信扣费成功的人数',
        '[8] 交易金额或次数超出限制',
        '[9] 余额不足/卡号冻结等支付方式异常',
        '[10] 未签约成功（找不到签约信息）',
        '[11] 解绑自动扣费（签约协议不存在）',
        '[12] 账号异常/冻结',
        '[13] 同一天多次扣费任务的重复账号',
        '[7]/[6] 实际扣费的成功率',
        '([8]+[9]+[12]+[13])/[6] 实际扣费时因账户异常的失败率',
        '([10]+[11])/[6] 实际扣费的解约率'
      )
    )
    
    # 填入最后输出表中的各项数量及计算转化率
    for(id in ad_id) {
      # Assigned data type must be compatible with existing data.
      count_ad_output[[id]] <- 0
      # [1] 白名单
      count_ad_output[
        count_ad_output$扣费阶段 == '数据清洗' & count_ad_output$情况 == '[1] 白名单', id
      ] <- count_ad[
        count_ad$扣费阶段 == 'Step 1：数据清洗' & count_ad$数据类型 == '白名单', id
      ]
      # [2] 自动扣费基数
      count_ad_output[
        count_ad_output$扣费阶段 == '数据清洗' & count_ad_output$情况 == '[2] 自动扣费基数', id
      ] <- (
        count_ad[
          count_ad$扣费阶段 == 'Step 1：数据清洗' & count_ad$数据类型 == '扣费基数', id
        ] + count_ad[
          count_ad$扣费阶段 == 'Step 1：数据清洗' & 
            count_ad$数据类型 == '签约关系不存在-商户号错误', id
        ]
      )
      # [3] 签约关系不存在
      count_ad_output[
        count_ad_output$扣费阶段 == '数据清洗' & 
          count_ad_output$情况 == '[3] 签约关系不存在（商户号错误）', id
      ] <- count_ad[
        count_ad$扣费阶段 == 'Step 1：数据清洗' & 
          count_ad$数据类型 == '签约关系不存在-商户号错误', id
      ]
      # 消息发送时间（先占个位）
      count_ad_output[
        count_ad_output$扣费阶段 == '消息发送' & count_ad_output$情况 == '消息发送时间', id
      ] <- 0
      # [4] 有重复下单
      count_ad_output[
        count_ad_output$扣费阶段 == '消息发送' & count_ad_output$情况 == '[4] 有重复下单', id
      ] <- count_ad[
        count_ad$扣费阶段 == 'Step 2：消息发送' & count_ad$数据类型 == '有重复下单', id
      ]
      # 发起扣费时间（先占个位）
      count_ad_output[
        count_ad_output$扣费阶段 == '发起扣费' & count_ad_output$情况 == '发起扣费时间', id
      ] <- 0
      # [5] 有重复下单
      count_ad_output[
        count_ad_output$扣费阶段 == '发起扣费' & count_ad_output$情况 == '[5] 有重复下单', id
      ] <- count_ad[
        count_ad$扣费阶段 == 'Step 3：发起扣费' & count_ad$数据类型 == '有重复下单', id
      ]
      # [6] 最终发起微信扣费的人数
      count_ad_output[
        count_ad_output$扣费阶段 == '发起扣费' & 
          count_ad_output$情况 == '[6] 最终发起微信扣费的人数', id
      ] <- (
        count_ad[
          count_ad$扣费阶段 == 'Step 1：数据清洗' & count_ad$数据类型 == '扣费基数', id
        ] - count_ad[
          count_ad$扣费阶段 == 'Step 2：消息发送' & count_ad$数据类型 == '有重复下单', id
        ] - count_ad[
          count_ad$扣费阶段 == 'Step 3：发起扣费' & count_ad$数据类型 == '有重复下单', id
        ]
      )
      # [6]/[2] 从扣费名单筛选到发起扣费的转化率
      count_ad_output[
        count_ad_output$扣费阶段 == '发起阶段转化率' & 
          count_ad_output$情况 == '[6]/[2] 从扣费名单筛选到发起扣费的转化率', id
      ] <- (
        count_ad_output[
          count_ad_output$扣费阶段 == '发起扣费' & 
            count_ad_output$情况 == '[6] 最终发起微信扣费的人数', id
        ] / count_ad_output[
          count_ad_output$扣费阶段 == '数据清洗' & 
            count_ad_output$情况 == '[2] 自动扣费基数', id
        ]
      )
      # [3]/[2] 发起扣费前的解约率
      count_ad_output[
        count_ad_output$扣费阶段 == '发起阶段转化率' & 
          count_ad_output$情况 == '[3]/[2] 发起扣费前的解约率', id
      ] <- (
        count_ad_output[
          count_ad_output$扣费阶段 == '数据清洗' & 
            count_ad_output$情况 == '[3] 签约关系不存在（商户号错误）', id
        ] / count_ad_output[
          count_ad_output$扣费阶段 == '数据清洗' & 
            count_ad_output$情况 == '[2] 自动扣费基数', id
        ]
      )
      # ([4]+[5])/[2] 发起扣费前的自主下单率
      count_ad_output[
        count_ad_output$扣费阶段 == '发起阶段转化率' & 
          count_ad_output$情况 == '([4]+[5])/[2] 发起扣费前的自主下单率', id
      ] <- (
        (
          count_ad_output[
            count_ad_output$扣费阶段 == '消息发送' & 
              count_ad_output$情况 == '[4] 有重复下单', id
          ] + count_ad_output[
            count_ad_output$扣费阶段 == '发起扣费' & 
              count_ad_output$情况 == '[5] 有重复下单', id
          ]
        ) / count_ad_output[
          count_ad_output$扣费阶段 == '数据清洗' & 
            count_ad_output$情况 == '[2] 自动扣费基数', id
        ]
      )
      # [7] 微信扣费成功的人数
      count_ad_output[
        count_ad_output$扣费阶段 == '扣费成功' & 
          count_ad_output$情况 == '[7] 微信扣费成功的人数', id
      ] <- count_ad[
        count_ad$扣费阶段 == 'Step 4：扣费结果' & 
          count_ad$数据类型 == '扣费成功', id
      ]
      # [8] 交易金额或次数超出限制
      count_ad_output[
        count_ad_output$扣费阶段 == '扣费异常' & 
          count_ad_output$情况 == '[8] 交易金额或次数超出限制', id
      ] <- count_ad[
        count_ad$扣费阶段 == 'Step 4：扣费结果' & 
          count_ad$数据类型 == '交易金额或次数超出限制', id
      ]
      # [9] 余额不足/卡号冻结等支付方式异常
      count_ad_output[
        count_ad_output$扣费阶段 == '扣费异常' & 
          count_ad_output$情况 == '[9] 余额不足/卡号冻结等支付方式异常', id
      ] <- count_ad[
        count_ad$扣费阶段 == 'Step 4：扣费结果' & 
          count_ad$数据类型 == '余额不足/卡号冻结等支付方式异常', id
      ]
      # [10] 未签约成功（找不到签约信息）
      count_ad_output[
        count_ad_output$扣费阶段 == '扣费异常' & 
          count_ad_output$情况 == '[10] 未签约成功（找不到签约信息）', id
      ] <- count_ad[
        count_ad$扣费阶段 == 'Step 4：扣费结果' & 
          count_ad$数据类型 == '未成功签约', id
      ]
      # [11] 解绑自动扣费（签约协议不存在）
      count_ad_output[
        count_ad_output$扣费阶段 == '扣费异常' & 
          count_ad_output$情况 == '[11] 解绑自动扣费（签约协议不存在）', id
      ] <- count_ad[
        count_ad$扣费阶段 == 'Step 4：扣费结果' & 
          count_ad$数据类型 == '解绑自动扣费-签约协议不存在', id
      ]
      # [12] 账号异常/冻结
      count_ad_output[
        count_ad_output$扣费阶段 == '扣费异常' & 
          count_ad_output$情况 == '[12] 账号异常/冻结', id
      ] <- count_ad[
        count_ad$扣费阶段 == 'Step 4：扣费结果' & 
          count_ad$数据类型 == '账号异常/冻结', id
      ]
      # [13] 同一天多次扣费任务的重复账号
      count_ad_output[
        count_ad_output$扣费阶段 == '扣费异常' & 
          count_ad_output$情况 == '[13] 同一天多次扣费任务的重复账号', id
      ] <- count_ad_output[
        count_ad_output$扣费阶段 == '发起扣费' & 
          count_ad_output$情况 == '[6] 最终发起微信扣费的人数', id
      ] - (
        count_ad_output[
          count_ad_output$扣费阶段 == '扣费成功' & 
            count_ad_output$情况 == '[7] 微信扣费成功的人数', id
        ] + count_ad_output[
          count_ad_output$扣费阶段 == '扣费异常' & 
            count_ad_output$情况 == '[8] 交易金额或次数超出限制', id
        ] + count_ad_output[
          count_ad_output$扣费阶段 == '扣费异常' & 
            count_ad_output$情况 == '[9] 余额不足/卡号冻结等支付方式异常', id
        ] + count_ad_output[
          count_ad_output$扣费阶段 == '扣费异常' & 
            count_ad_output$情况 == '[10] 未签约成功（找不到签约信息）', id
        ] + count_ad_output[
          count_ad_output$扣费阶段 == '扣费异常' & 
            count_ad_output$情况 == '[11] 解绑自动扣费（签约协议不存在）', id
        ] + count_ad_output[
          count_ad_output$扣费阶段 == '扣费异常' & 
            count_ad_output$情况 == '[12] 账号异常/冻结', id
        ]
      )
      # [7]/[6] 实际扣费的成功率
      count_ad_output[
        count_ad_output$扣费阶段 == '实际扣费转化率' & 
          count_ad_output$情况 == '[7]/[6] 实际扣费的成功率', id
      ] <- (
        count_ad_output[
          count_ad_output$扣费阶段 == '扣费成功' & 
            count_ad_output$情况 == '[7] 微信扣费成功的人数', id
        ] / count_ad_output[
          count_ad_output$扣费阶段 == '发起扣费' & 
            count_ad_output$情况 == '[6] 最终发起微信扣费的人数', id
        ]
      )
      # ([8]+[9]+[12]+[13])/[6] 实际扣费时因账户异常的失败率
      count_ad_output[
        count_ad_output$扣费阶段 == '实际扣费转化率' & 
          count_ad_output$情况 == '([8]+[9]+[12]+[13])/[6] 实际扣费时因账户异常的失败率', id
      ] <- (
        (
          count_ad_output[
            count_ad_output$扣费阶段 == '扣费异常' & 
              count_ad_output$情况 == '[8] 交易金额或次数超出限制', id
          ] + count_ad_output[
            count_ad_output$扣费阶段 == '扣费异常' & 
              count_ad_output$情况 == '[9] 余额不足/卡号冻结等支付方式异常', id
          ] + count_ad_output[
            count_ad_output$扣费阶段 == '扣费异常' & 
              count_ad_output$情况 == '[12] 账号异常/冻结', id
          ] + count_ad_output[
            count_ad_output$扣费阶段 == '扣费异常' & 
              count_ad_output$情况 == '[13] 同一天多次扣费任务的重复账号', id
          ]
        ) / count_ad_output[
          count_ad_output$扣费阶段 == '发起扣费' & 
            count_ad_output$情况 == '[6] 最终发起微信扣费的人数', id
        ]
      )
      # ([10]+[11])/[6] 实际扣费的解约率
      count_ad_output[
        count_ad_output$扣费阶段 == '实际扣费转化率' & 
          count_ad_output$情况 == '([10]+[11])/[6] 实际扣费的解约率', id
      ] <- (
        (
          count_ad_output[
            count_ad_output$扣费阶段 == '扣费异常' & 
              count_ad_output$情况 == '[10] 未签约成功（找不到签约信息）', id
          ] + count_ad_output[
            count_ad_output$扣费阶段 == '扣费异常' & 
              count_ad_output$情况 == '[11] 解绑自动扣费（签约协议不存在）', id
          ]
        ) / count_ad_output[
          count_ad_output$扣费阶段 == '发起扣费' & 
            count_ad_output$情况 == '[6] 最终发起微信扣费的人数', id
        ]
      )
      # Assigned data type must be compatible with existing data.
      count_ad_output[[id]] <- as.character(count_ad_output[[id]])
      # 消息发送时间
      count_ad_output[
        count_ad_output$扣费阶段 == '消息发送' & count_ad_output$情况 == '消息发送时间', id
      ] <- auto_deduction_base_config[
        auto_deduction_base_config$id == id, 'advance_deduction_notice_time'
      ]
      # 发起扣费时间
      count_ad_output[
        count_ad_output$扣费阶段 == '发起扣费' & count_ad_output$情况 == '发起扣费时间', id
      ] <- auto_deduction_base_config[
        auto_deduction_base_config$id == id, 'start_deduction_time'
      ]
    }
    # count_ad_output <- count_ad_output %>% 
    #   mutate(
    #     across(
    #       # mutate(across(where())): https://www.tidyverse.org/blog/2020/04/dplyr-1-0-0-colwise/
    #       where(is.numeric), ~ round(., 4)
    #     )
    #   )
    
    # 将小数转换成百分比（因为output表是固定的，所以直接找到转化率的行，从第3列开始，转换小数为百分比）
    count_ad_output <- as.data.frame(count_ad_output)
    for(
      i in 3:ncol(
        count_ad_output[count_ad_output$扣费阶段 %in% c('发起阶段转化率', '实际扣费转化率'), ]
      )
    ) {
      count_ad_output[count_ad_output$扣费阶段 %in% c('发起阶段转化率', '实际扣费转化率'), ][[i]] <- 
        scales::label_percent(
          accuracy = 0.1, big.mark = ""
        )(
          as.numeric(
            count_ad_output[
              count_ad_output$扣费阶段 %in% c('发起阶段转化率', '实际扣费转化率'), 
            ][[i]]
          )
        ) %>% 
        str_replace('Inf', '')
    }
    output_list[['自动扣费各项明细统计']] <- count_ad_output
    
    # 断开数据库
    dbDisconnect(product)
    rm(product)
  }
  
  # Sheet 3：续保情况-去年 ----
  # 去年个人订单续保统计
  if(count_self_renew_last_year) {
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
  
  # 企业订单续保统计
  if(count_com_renew_last_year) {
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
      保险公司 = c(unique(df_orders_self_old$col_company_self), '合计'),
      去年企业订单总数 = 0,
      企业订单续保数 = 0,
      企业订单续保率 = 0
    )
  }
  
  renew_orders <- full_join(self_renew_orders, com_renew_orders)
  renew_orders <- renew_orders %>% 
    mutate(across(.cols = where(is.numeric), .fns = ~ replace_na(., 0)))
  # 百分比格式转换
  column_to_scale <- colnames(renew_orders)[str_detect(colnames(renew_orders), '占比|率')]
  for(column in column_to_scale) {
    renew_orders[[column]] <- scales::label_percent(
      accuracy = 0.1, big.mark = ""
    )(
      as.numeric(renew_orders[[column]])
    ) %>% 
      str_replace('Inf', '')
  }
  output_list[['续保情况-去年']] <- renew_orders
  
  # Sheet 4：续保情况-今年 ----
  if(count_self_renew_this_year) {
    self_renew_orders_this_year <- df_orders_self %>% 
      mutate(
        is_renew = ifelse(
          (col_id_no_self %in% df_orders_self_old$col_id_no_self_old) | 
            (col_id_no_self %in% df_orders_com_old$col_id_no_com_old), 
          1, 
          0
        )
      ) %>% 
      group_by(col_company_self) %>% 
      mutate(
        今年个人订单总数 = length(col_id_no_self),
        个人订单续保数 = sum(is_renew)
      ) %>% 
      ungroup() %>% 
      select(保险公司 = col_company_self, 今年个人订单总数, 个人订单续保数) %>% 
      distinct() %>% 
      mutate(个人订单续保占比 = round(个人订单续保数 / 今年个人订单总数, 3))
    last_row <- tibble(
      保险公司 = '合计', 
      今年个人订单总数 = sum(na.omit(self_renew_orders_this_year$今年个人订单总数)), 
      个人订单续保数 = sum(na.omit(self_renew_orders_this_year$个人订单续保数)),
      个人订单续保占比 = round(
        na.omit(sum(self_renew_orders_this_year$个人订单续保数)) / 
          na.omit(sum(self_renew_orders_this_year$今年个人订单总数)), 
        3
      )
    )
    self_renew_orders_this_year <- bind_rows(self_renew_orders_this_year, last_row)
  }
  
  if(count_com_renew_this_year) {
    # 企业订单的续保统计
    com_renew_orders_this_year <- df_orders_com %>% 
      mutate(
        is_renew = ifelse(
          (col_id_no_com %in% df_orders_com_old$col_id_no_com_old) | 
            (col_id_no_com %in% df_orders_self_old$col_id_no_self_old), 
          1, 
          0
        )
      ) %>% 
      group_by(col_company_com) %>% 
      mutate(
        今年企业订单总数 = length(col_id_no_com),
        企业订单续保数 = sum(is_renew)
      ) %>% 
      ungroup() %>% 
      select(保险公司 = col_company_com, 今年企业订单总数, 企业订单续保数) %>% 
      distinct() %>% 
      mutate(企业订单续保占比 = round(企业订单续保数 / 今年企业订单总数, 3))
    last_row <- tibble(
      保险公司 = '合计', 
      今年企业订单总数 = sum(na.omit(com_renew_orders_this_year$今年企业订单总数)), 
      企业订单续保数 = sum(na.omit(com_renew_orders_this_year$企业订单续保数)), 
      企业订单续保占比 = round(
        na.omit(sum(com_renew_orders_this_year$企业订单续保数)) / 
          na.omit(sum(com_renew_orders_this_year$今年企业订单总数)), 
        3
      )
    )
    com_renew_orders_this_year <- bind_rows(com_renew_orders_this_year, last_row)
  } else {
    com_renew_orders_this_year <- tibble(
      保险公司 = c(unique(df_orders_self$col_company_self), '合计'),
      今年企业订单总数 = 0,
      企业订单续保数 = 0,
      企业订单续保占比 = 0
    )
  }
  
  renew_orders_this_year <- full_join(self_renew_orders_this_year, com_renew_orders_this_year)
  renew_orders_this_year <- renew_orders_this_year %>% 
    mutate(across(.cols = where(is.numeric), .fns = ~ replace_na(., 0)))
  # 百分比格式转换
  column_to_scale <- colnames(renew_orders_this_year)[
    str_detect(colnames(renew_orders_this_year), '占比|率')
  ]
  for(column in column_to_scale) {
    renew_orders_this_year[[column]] <- scales::label_percent(
      accuracy = 0.1, big.mark = ""
    )(
      as.numeric(renew_orders_this_year[[column]])
    ) %>% 
      str_replace('Inf', '')
  }
  output_list[['续保情况-今年']] <- renew_orders_this_year
  
  # if(save_result) {
  #   # 创建文件夹
  #   if( !dir.exists(paste0(sop_dir, '/', city)) ) {
  #     dir.create(paste0(sop_dir, '/', city))
  #   }
  #   if( !dir.exists(paste0(sop_dir, '/', city, '/', as.Date(end_time, tz = Sys.timezone()))) ) {
  #     dir.create(paste0(sop_dir, '/', city, '/', as.Date(end_time, tz = Sys.timezone())))
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
