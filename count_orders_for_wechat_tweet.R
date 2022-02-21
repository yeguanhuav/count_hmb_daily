###################################### 公众号推文所需数据统计 ######################################
#library(ggrepel)
library(glue)
#library(openxlsx)
library(scales)
library(writexl)
library(lubridate)
library(tidyverse)

count_orders_for_wechat_tweet <- function(
  city,
  df_orders,
  district_info,
  group_district_orders_by_id_no = TRUE, # 用【身份证前6位】分组进行区县统计
  group_district_orders_by_sale_channel = FALSE, # 用【销售渠道】分组进行区县统计
  drop_refunds = TRUE,
  col_policy_order_status = 'policy_order_status',
  have_relationship = FALSE,
  col_relation_ship = 'relation_ship',
  start_time = NA,
  end_time = Sys.time(),
  col_order_item_id = 'order_item_id',
  col_date_created = 'date_created',
  col_date_updated = 'date_updated',
  col_name = 'name',
  col_id_type = 'id_type',
  col_id_no = 'id_no',
  col_sex = 'sex',
  col_birthday = 'birthday',
  col_group = 'group',
  col_channel_cls1 = 'channel_cls1',
  col_company = 'company',
  online_company = '线上平台|思派',
  col_customer_id = 'customer_id',
  col_automatic_deduction = 'automatic_deduction',
  col_medical_insure_flag = NA, # 预约医保个账扣缴
  col_pay_way = NA, # 支付方式
  #plot_district_pie_chart = FALSE,
  #legend_position = 'right',
  #font = NA,
  #save_district = FALSE,
  save_all = TRUE,
  save_result = FALSE
) {
  sop_dir <- getwd()
  
  # 先规定用什么方法划分区县订单
  if(
    (group_district_orders_by_id_no & group_district_orders_by_sale_channel) |
    (!group_district_orders_by_id_no & !group_district_orders_by_sale_channel)
  ) {
    stop(
      glue(
        "group_district_orders_by_id_no & group_district_orders_by_sale_channel are both TRUE or ",
        "both FALSE, please set one of them as TRUE and the other as FALSE."
      )
    )
  }
  
  # 保留所需的列 ----
  df_orders <- df_orders %>% 
    rename(
      col_order_item_id = !!col_order_item_id,
      col_date_created = !!col_date_created,
      col_date_updated = !!col_date_updated,
      col_group = !!col_group,
      col_policy_order_status = !!col_policy_order_status,
      col_customer_id = !!col_customer_id,
      col_name = !!col_name,
      col_id_type = !!col_id_type,
      col_id_no = !!col_id_no,
      col_birthday = !!col_birthday,
      col_sex = !!col_sex,
      col_channel_cls1 = !!col_channel_cls1,
      col_company = !!col_company,
      col_automatic_deduction = !!col_automatic_deduction
    ) %>% 
    mutate(
      col_datetime_created = as.POSIXct(col_date_created, tz = Sys.timezone()),
      col_datetime_updated = as.POSIXct(col_date_updated, tz = Sys.timezone())
    ) %>% 
    # 转换日期格式
    mutate(col_date_created = as.Date(col_date_created, tz = Sys.timezone()))
  if(have_relationship) {
    df_orders <- df_orders %>% rename(col_relation_ship = !!col_relation_ship)
  }
  # 预约医保个账扣缴
  if(!is.na(col_medical_insure_flag)) {
    df_orders <- df_orders %>% rename(col_medical_insure_flag = !!col_medical_insure_flag)
  }
  # 支付方式
  if(!is.na(col_pay_way)) {
    df_orders <- df_orders %>% rename(col_pay_way = !!col_pay_way)
  }
  # 去掉多余的列，减少内存使用
  df_orders <- df_orders %>% select(contains('col_'))
  
  # 开始日期和结束日期
  if(is.na(start_time)) { start_time <- min(df_orders[['col_date_created']]) }
  # 按input日期筛选订单（注意input日期格式）
  df_orders <- df_orders %>% 
    filter(col_datetime_created >= as.POSIXct(start_time) & 
             col_datetime_created <= as.POSIXct(end_time))
  
  # 去掉退单 ----
  #order_type <- '个人和企业'
  refund_status <- c('04', '05')
  if(drop_refunds) {
    df_orders <- df_orders %>% filter(!col_policy_order_status %in% !!refund_status)
  }
  
  # 统计数据 ----
  # # Create an Excel workbook object and add a worksheet
  # df_orders_wb <- createWorkbook()
  # # Create a percent style
  # pct = createStyle(numFmt = "0.0%")
  # sht1 = addWorksheet(df_orders_wb, '投保数据统计')
  output_list <- list()
  
  df_orders <- df_orders %>% 
    # 转换日期格式
    mutate(
      col_date_created = as.Date(col_date_created, tz = Sys.timezone()),
      col_birthday = as.Date(col_birthday)
    ) %>% 
    # 生成信息，注意顺序
    mutate(
      age = as.numeric(col_date_created - col_birthday), # 参保时的年龄
      district = str_replace(col_id_no, '(\\d\\d\\d\\d\\d\\d).*', '\\1'), # 地区
      birthday_year = year(col_birthday), # 生肖
      birthday_month_day = format(col_birthday, "%m-%d"), # 星座
      star_sign = case_when(
        birthday_month_day >= format(as.Date('2021-01-21'), '%m-%d') & 
          birthday_month_day <= format(as.Date('2021-02-19'), '%m-%d') ~ '水瓶座',
        birthday_month_day >= format(as.Date('2021-02-20'), '%m-%d') & 
          birthday_month_day <= format(as.Date('2021-03-20'), '%m-%d') ~ '双鱼座',
        birthday_month_day >= format(as.Date('2021-03-21'), '%m-%d') & 
          birthday_month_day <= format(as.Date('2021-04-20'), '%m-%d') ~ '白羊座',
        birthday_month_day >= format(as.Date('2021-04-21'), '%m-%d') & 
          birthday_month_day <= format(as.Date('2021-05-21'), '%m-%d') ~ '金牛座',
        birthday_month_day >= format(as.Date('2021-05-22'), '%m-%d') & 
          birthday_month_day <= format(as.Date('2021-06-21'), '%m-%d') ~ '双子座',
        birthday_month_day >= format(as.Date('2021-06-22'), '%m-%d') & 
          birthday_month_day <= format(as.Date('2021-07-23'), '%m-%d') ~ '巨蟹座',
        birthday_month_day >= format(as.Date('2021-07-24'), '%m-%d') & 
          birthday_month_day <= format(as.Date('2021-08-23'), '%m-%d') ~ '狮子座',
        birthday_month_day >= format(as.Date('2021-08-24'), '%m-%d') & 
          birthday_month_day <= format(as.Date('2021-09-23'), '%m-%d') ~ '处女座',
        birthday_month_day >= format(as.Date('2021-09-24'), '%m-%d') & 
          birthday_month_day <= format(as.Date('2021-10-23'), '%m-%d') ~ '天秤座',
        birthday_month_day >= format(as.Date('2021-10-24'), '%m-%d') & 
          birthday_month_day <= format(as.Date('2021-11-22'), '%m-%d') ~ '天蝎座',
        birthday_month_day >= format(as.Date('2021-11-23'), '%m-%d') & 
          birthday_month_day <= format(as.Date('2021-12-22'), '%m-%d') ~ '射手座',
        TRUE ~ '摩羯座'
      ),
      zodiac = case_when(
        (birthday_year %% 12) == 0 ~ "属猴",
        (birthday_year %% 12) == 1 ~ "属鸡",
        (birthday_year %% 12) == 2 ~ "属狗",
        (birthday_year %% 12) == 3 ~ "属猪",
        (birthday_year %% 12) == 4 ~ "属鼠",
        (birthday_year %% 12) == 5 ~ "属牛",
        (birthday_year %% 12) == 6 ~ "属虎",
        (birthday_year %% 12) == 7 ~ "属兔",
        (birthday_year %% 12) == 8 ~ "属龙",
        (birthday_year %% 12) == 9 ~ "属蛇",
        (birthday_year %% 12) == 10 ~ "属马",
        (birthday_year %% 12) == 11 ~ "属羊"
      )
    )
  
  # Sheet 1: ----
  output_table <- tibble(
    分组 = c(
      '净投保量', 
      '个人净投保量', 
      '企业净投保量',
      '线上平台',
      '线下保司',
      #'线上平台自动扣费签约',
      #'线下保司自动扣费签约',
      '男性', 
      "女性",
      '和项目同一天生日', 
      '生日当天购买', 
      '港澳台同胞', 
      '国际友人',
      #'为自己投保', 
      #'为家人投保',
      #'年龄最小参保人', 
      #'年龄最大参保人',
      '0岁', 
      '1-20岁', 
      '21-40岁', 
      '41-60岁', 
      '61-80岁', 
      '81-100岁',
      '>100岁',
      '80岁及以上', 
      '100岁及以上'
    ),
    人数 = 0.1, 占比 = 1
  )
  
  # 个人、企业的投保量 ----
  output_table[output_table$分组 == '个人净投保量', colnames(output_table) == '人数'] <- 
    nrow(filter(df_orders, col_group == 'self'))
  output_table[output_table$分组 == '个人净投保量', colnames(output_table) == '占比'] <- 
    round(nrow(filter(df_orders, col_group == 'self')) / nrow(df_orders), 3)
  
  output_table[output_table$分组 == '企业净投保量', colnames(output_table) == '人数'] <- 
    nrow(filter(df_orders, col_group == 'com'))
  output_table[output_table$分组 == '企业净投保量', colnames(output_table) == '占比'] <- 
    round(nrow(filter(df_orders, col_group == 'com')) / nrow(df_orders), 3)
  
  # 总投保量数和占比 ----
  output_table[output_table$分组 == '净投保量', colnames(output_table) == '人数'] <- 
    length(df_orders$col_order_item_id)
  output_table[output_table$分组 == '净投保量', colnames(output_table) == '占比'] <- 
    round(length(df_orders$col_order_item_id) / length(df_orders$col_order_item_id), 3)
  
  # 线上平台和线下保司的人数和占比 ----
  output_table[output_table$分组 == '线上平台', colnames(output_table) == '人数'] <- 
    nrow(filter(df_orders, str_detect(col_company, !!online_company)))
  output_table[output_table$分组 == '线上平台', colnames(output_table) == '占比'] <- 
    round(
      nrow(filter(df_orders, str_detect(col_company, !!online_company))) / 
        nrow(df_orders), 
      3
    )
  
  output_table[output_table$分组 == '线下保司', colnames(output_table) == '人数'] <- 
    (nrow(df_orders) - nrow(filter(df_orders, str_detect(col_company, !!online_company))))
  output_table[output_table$分组 == '线下保司', colnames(output_table) == '占比'] <- 
    round(
      (nrow(df_orders) - nrow(filter(df_orders, str_detect(col_company, !!online_company)))) / 
        nrow(df_orders), 
      3
    )
  
  # # 线上平台和线下保司的自动扣费签约人数和占比 ----
  # output_table[output_table$分组 == '线上平台自动扣费签约', colnames(output_table) == '人数'] <- 
  #   nrow(
  #     df_orders %>% 
  #       filter(col_group == 'self') %>% 
  #       filter(str_detect(col_company, !!online_company)) %>% 
  #       filter(col_automatic_deduction == 1)
  #   )
  # output_table[output_table$分组 == '线上平台自动扣费签约', colnames(output_table) == '占比'] <- 
  #   round(
  #     nrow(
  #       df_orders %>% 
  #         filter(col_group == 'self') %>% 
  #         filter(str_detect(col_company, !!online_company)) %>% 
  #         filter(col_automatic_deduction == 1)
  #     ) / nrow(df_orders), 
  #     3
  #   )
  # 
  # output_table[output_table$分组 == '线下保司自动扣费签约', colnames(output_table) == '人数'] <- 
  #   nrow(
  #     df_orders %>% 
  #       filter(col_group == 'self') %>% 
  #       filter(str_detect(col_company, !!online_company, negate = T)) %>% 
  #       filter(col_automatic_deduction == 1)
  #   )
  # output_table[output_table$分组 == '线下保司自动扣费签约', colnames(output_table) == '占比'] <- 
  #   round(
  #     nrow(
  #       df_orders %>% 
  #         filter(col_group == 'self') %>% 
  #         filter(str_detect(col_company, !!online_company, negate = T)) %>% 
  #         filter(col_automatic_deduction == 1)
  #     ) / nrow(df_orders), 
  #     3
  #   )
  
  # 男女参保人数和占比 ----
  # 男：
  output_table[output_table$分组 == "男性", colnames(output_table) == '人数'] <- 
    length(which(df_orders$col_sex == '男'))
  output_table[output_table$分组 == "男性", colnames(output_table) == '占比'] <- 
    round(length(which(df_orders$col_sex == '男')) / length(df_orders$col_order_item_id), 3)
  # 女：
  output_table[output_table$分组 == "女性", colnames(output_table) == '人数'] <- 
    length(which(df_orders$col_sex == '女'))
  output_table[output_table$分组 == "女性", colnames(output_table) == '占比'] <- 
    round(length(which(df_orders$col_sex == '女')) / length(df_orders$col_order_item_id), 3)
  
  # 港澳台同胞及国际友人 ----
  output_table[output_table$分组 == '港澳台同胞', colnames(output_table) == '人数'] <- 
    nrow(filter(df_orders, col_id_type %in% c('台湾居民来往大陆通行证', '港澳居民来往内地通行证')))
  output_table[output_table$分组 == '港澳台同胞', colnames(output_table) == '占比'] <- 
    round(
      nrow(
        filter(df_orders, col_id_type %in% c('台湾居民来往大陆通行证', '港澳居民来往内地通行证'))
      ) / nrow(df_orders), 
      3
    )
  output_table[output_table$分组 == '国际友人', colnames(output_table) == '人数'] <- 
    nrow(filter(df_orders, col_id_type %in% c('护照')))
  output_table[output_table$分组 == '国际友人', colnames(output_table) == '占比'] <- 
    round(nrow(filter(df_orders, col_id_type %in% c('护照'))) / nrow(df_orders), 3)
  
  # 和项目同一天生日及生日当天购买 ----
  start_month_day <- format(as.Date(start_time, tz = Sys.timezone()), "%m-%d")
  # 项目上线当天生日的参保用户
  output_table[
    output_table$分组 == "和项目同一天生日", colnames(output_table) == '人数'
  ] <- nrow(df_orders %>% filter(birthday_month_day == start_month_day))
  output_table[
    output_table$分组 == "和项目同一天生日", colnames(output_table) == '占比'
  ] <- nrow(df_orders %>% filter(birthday_month_day == start_month_day)) / nrow(df_orders)
  # 生日当天购买的用户（根据身份证取生日去计算）
  df_orders <- df_orders %>% mutate(
    id_no_birthday = case_when(
      col_id_type == '身份证' & nchar(col_id_no) == 18 ~ 
        str_replace(col_id_no, '..........(..)(..).*', '\\1-\\2'),
      col_id_type == '身份证' & nchar(col_id_no) == 15 ~
        str_replace(col_id_no, '........(..)(..).*', '\\1-\\2'),
      TRUE ~ str_replace(col_birthday, '.....(.*)', '\\1')
    )
  )
  df_orders <- df_orders %>% mutate(col_date_created_month_day = format(col_date_created, "%m-%d"))
  output_table[
    output_table$分组 == "生日当天购买", colnames(output_table) == '人数'
  ] <- nrow(df_orders %>% filter(col_date_created_month_day == id_no_birthday))
  output_table[
    output_table$分组 == "生日当天购买", colnames(output_table) == '占比'
  ] <- nrow(df_orders %>% filter(col_date_created_month_day == id_no_birthday)) / nrow(df_orders)
  
  # 各年龄段参保人数 ----
  # 0岁
  output_table[output_table$分组 == '0岁', colnames(output_table) == '人数'] <- 
    nrow(filter(df_orders, age < 365))
  output_table[output_table$分组 == '0岁', colnames(output_table) == '占比'] <- 
    round(nrow(filter(df_orders, age < 365)) / nrow(df_orders), 3)
  
  # 1-20岁
  output_table[output_table$分组 == '1-20岁', colnames(output_table) == '人数'] <- 
    nrow(filter(df_orders, age >= 365 & age < 365*21))
  output_table[output_table$分组 == '1-20岁', colnames(output_table) == '占比'] <- 
    round(nrow(filter(df_orders, age >= 365 & age < 365*21)) / nrow(df_orders), 3)
  
  # 21-40岁
  output_table[output_table$分组 == '21-40岁', colnames(output_table) == '人数'] <- 
    nrow(filter(df_orders, age >= 365*21 & age < 365*41))
  output_table[output_table$分组 == '21-40岁', colnames(output_table) == '占比'] <- 
    round(nrow(filter(df_orders, age >= 365*21 & age < 365*41)) / nrow(df_orders), 3)
  
  # 41-60岁
  output_table[output_table$分组 == '41-60岁', colnames(output_table) == '人数'] <- 
    nrow(filter(df_orders, age >= 365*41 & age < 365*61))
  output_table[output_table$分组 == '41-60岁', colnames(output_table) == '占比'] <- 
    round(nrow(filter(df_orders, age >= 365*41 & age < 365*61)) / nrow(df_orders), 3)
  
  # 61-80岁
  output_table[output_table$分组 == '61-80岁', colnames(output_table) == '人数'] <- 
    nrow(filter(df_orders, age >= 365*61 & age < 365*81))
  output_table[output_table$分组 == '61-80岁', colnames(output_table) == '占比'] <- 
    round(nrow(filter(df_orders, age >= 365*61 & age < 365*81)) / nrow(df_orders), 3)
  
  # 81-100岁
  output_table[output_table$分组 == '81-100岁', colnames(output_table) == '人数'] <- 
    nrow(filter(df_orders, age >= 365*81 & age < 365*101))
  output_table[output_table$分组 == '81-100岁', colnames(output_table) == '占比'] <- 
    round(nrow(filter(df_orders, age >= 365*81 & age < 365*101)) / nrow(df_orders), 3)
  
  # >100岁
  output_table[output_table$分组 == '>100岁', colnames(output_table) == '人数'] <- 
    nrow(filter(df_orders, age >= 365*101))
  output_table[output_table$分组 == '>100岁', colnames(output_table) == '占比'] <- 
    round(nrow(filter(df_orders, age >= 365*101)) / nrow(df_orders), 3)
  
  # 80岁及以上
  output_table[output_table$分组 == '80岁及以上', colnames(output_table) == '人数'] <- 
    nrow(filter(df_orders, age >= 365*80))
  output_table[output_table$分组 == '80岁及以上', colnames(output_table) == '占比'] <- 
    round(nrow(filter(df_orders, age >= 365*80)) / nrow(df_orders), 3)
  
  # 100岁及以上有多少
  output_table[output_table$分组 == '100岁及以上', colnames(output_table) == '人数'] <- 
    nrow(filter(df_orders, age >= 365*100))
  output_table[output_table$分组 == '100岁及以上', colnames(output_table) == '占比'] <- 
    round(nrow(filter(df_orders, age >= 365*100)) / nrow(df_orders), 3)
  
  # 星座 ----
  star_sign <- tibble(
    '分组' = c(
      "水瓶座",
      "双鱼座",
      "白羊座",
      "金牛座",
      "双子座",
      "巨蟹座",
      "狮子座",
      "处女座",
      "天秤座",
      "天蝎座",
      "射手座",
      "摩羯座"
    ),
    '人数' = 0.1, '占比' = 1
  )
  star_sign[star_sign$分组 == '水瓶座', colnames(star_sign) == '人数'] <- 
    nrow(filter(df_orders, star_sign == '水瓶座'))
  star_sign[star_sign$分组 == '水瓶座', colnames(star_sign) == '占比'] <- 
    round(nrow(filter(df_orders, star_sign == '水瓶座')) / nrow(df_orders), 3)
  star_sign[star_sign$分组 == '双鱼座', colnames(star_sign) == '人数'] <- 
    nrow(filter(df_orders, star_sign == '双鱼座'))
  star_sign[star_sign$分组 == '双鱼座', colnames(star_sign) == '占比'] <- 
    round(nrow(filter(df_orders, star_sign == '双鱼座')) / nrow(df_orders), 3)
  star_sign[star_sign$分组 == '白羊座', colnames(star_sign) == '人数'] <- 
    nrow(filter(df_orders, star_sign == '白羊座'))
  star_sign[star_sign$分组 == '白羊座', colnames(star_sign) == '占比'] <- 
    round(nrow(filter(df_orders, star_sign == '白羊座')) / nrow(df_orders), 3)
  star_sign[star_sign$分组 == '金牛座', colnames(star_sign) == '人数'] <- 
    nrow(filter(df_orders, star_sign == '金牛座'))
  star_sign[star_sign$分组 == '金牛座', colnames(star_sign) == '占比'] <- 
    round(nrow(filter(df_orders, star_sign == '金牛座')) / nrow(df_orders), 3)
  star_sign[star_sign$分组 == '双子座', colnames(star_sign) == '人数'] <- 
    nrow(filter(df_orders, star_sign == '双子座'))
  star_sign[star_sign$分组 == '双子座', colnames(star_sign) == '占比'] <- 
    round(nrow(filter(df_orders, star_sign == '双子座')) / nrow(df_orders), 3)
  star_sign[star_sign$分组 == '巨蟹座', colnames(star_sign) == '人数'] <- 
    nrow(filter(df_orders, star_sign == '巨蟹座'))
  star_sign[star_sign$分组 == '巨蟹座', colnames(star_sign) == '占比'] <- 
    round(nrow(filter(df_orders, star_sign == '巨蟹座')) / nrow(df_orders), 3)
  star_sign[star_sign$分组 == '狮子座', colnames(star_sign) == '人数'] <- 
    nrow(filter(df_orders, star_sign == '狮子座'))
  star_sign[star_sign$分组 == '狮子座', colnames(star_sign) == '占比'] <- 
    round(nrow(filter(df_orders, star_sign == '狮子座')) / nrow(df_orders), 3)
  star_sign[star_sign$分组 == '处女座', colnames(star_sign) == '人数'] <- 
    nrow(filter(df_orders, star_sign == '处女座'))
  star_sign[star_sign$分组 == '处女座', colnames(star_sign) == '占比'] <- 
    round(nrow(filter(df_orders, star_sign == '处女座')) / nrow(df_orders), 3)
  star_sign[star_sign$分组 == '天秤座', colnames(star_sign) == '人数'] <- 
    nrow(filter(df_orders, star_sign == '天秤座'))
  star_sign[star_sign$分组 == '天秤座', colnames(star_sign) == '占比'] <- 
    round(nrow(filter(df_orders, star_sign == '天秤座')) / nrow(df_orders), 3)
  star_sign[star_sign$分组 == '天蝎座', colnames(star_sign) == '人数'] <- 
    nrow(filter(df_orders, star_sign == '天蝎座'))
  star_sign[star_sign$分组 == '天蝎座', colnames(star_sign) == '占比'] <- 
    round(nrow(filter(df_orders, star_sign == '天蝎座')) / nrow(df_orders), 3)
  star_sign[star_sign$分组 == '射手座', colnames(star_sign) == '人数'] <- 
    nrow(filter(df_orders, star_sign == '射手座'))
  star_sign[star_sign$分组 == '射手座', colnames(star_sign) == '占比'] <- 
    round(nrow(filter(df_orders, star_sign == '射手座')) / nrow(df_orders), 3)
  star_sign[star_sign$分组 == '摩羯座', colnames(star_sign) == '人数'] <- 
    nrow(filter(df_orders, star_sign == '摩羯座'))
  star_sign[star_sign$分组 == '摩羯座', colnames(star_sign) == '占比'] <- 
    round(nrow(filter(df_orders, star_sign == '摩羯座')) / nrow(df_orders), 3)
  if(save_all) {
    star_sign <- star_sign %>% arrange(desc(人数))
  } else {
    # 取人数前三
    star_sign <- star_sign %>% slice_max(order_by = 人数, n = 3)
  }
  # 添加到输出表格
  output_table <- bind_rows(output_table, star_sign)
  
  # 生肖 ----
  zodiac <- tibble(
    '分组' = c(
      "属鼠",
      "属牛",
      "属虎",
      "属兔",
      "属龙",
      "属蛇",
      "属马",
      "属羊",
      "属猴",
      "属鸡",
      "属狗",
      "属猪"
    ),
    '人数' = 0.1, '占比' = 1
  )
  zodiac[zodiac$分组 == '属猴', colnames(zodiac) == '人数'] <- 
    nrow(filter(df_orders, zodiac == '属猴'))
  zodiac[zodiac$分组 == '属猴', colnames(zodiac) == '占比'] <- 
    round(nrow(filter(df_orders, zodiac == '属猴')) / nrow(df_orders), 3)
  zodiac[zodiac$分组 == '属鸡', colnames(zodiac) == '人数'] <- 
    nrow(filter(df_orders, zodiac == '属鸡'))
  zodiac[zodiac$分组 == '属鸡', colnames(zodiac) == '占比'] <- 
    round(nrow(filter(df_orders, zodiac == '属鸡')) / nrow(df_orders), 3)
  zodiac[zodiac$分组 == '属狗', colnames(zodiac) == '人数'] <- 
    nrow(filter(df_orders, zodiac == '属狗'))
  zodiac[zodiac$分组 == '属狗', colnames(zodiac) == '占比'] <- 
    round(nrow(filter(df_orders, zodiac == '属狗')) / nrow(df_orders), 3)
  zodiac[zodiac$分组 == '属猪', colnames(zodiac) == '人数'] <- 
    nrow(filter(df_orders, zodiac == '属猪'))
  zodiac[zodiac$分组 == '属猪', colnames(zodiac) == '占比'] <- 
    round(nrow(filter(df_orders, zodiac == '属猪')) / nrow(df_orders), 3)
  zodiac[zodiac$分组 == '属鼠', colnames(zodiac) == '人数'] <- 
    nrow(filter(df_orders, zodiac == '属鼠'))
  zodiac[zodiac$分组 == '属鼠', colnames(zodiac) == '占比'] <- 
    round(nrow(filter(df_orders, zodiac == '属鼠')) / nrow(df_orders), 3)
  zodiac[zodiac$分组 == '属牛', colnames(zodiac) == '人数'] <- 
    nrow(filter(df_orders, zodiac == '属牛'))
  zodiac[zodiac$分组 == '属牛', colnames(zodiac) == '占比'] <- 
    round(nrow(filter(df_orders, zodiac == '属牛')) / nrow(df_orders), 3)
  zodiac[zodiac$分组 == '属虎', colnames(zodiac) == '人数'] <- 
    nrow(filter(df_orders, zodiac == '属虎'))
  zodiac[zodiac$分组 == '属虎', colnames(zodiac) == '占比'] <- 
    round(nrow(filter(df_orders, zodiac == '属虎')) / nrow(df_orders), 3)
  zodiac[zodiac$分组 == '属兔', colnames(zodiac) == '人数'] <- 
    nrow(filter(df_orders, zodiac == '属兔'))
  zodiac[zodiac$分组 == '属兔', colnames(zodiac) == '占比'] <- 
    round(nrow(filter(df_orders, zodiac == '属兔')) / nrow(df_orders), 3)
  zodiac[zodiac$分组 == '属龙', colnames(zodiac) == '人数'] <- 
    nrow(filter(df_orders, zodiac == '属龙'))
  zodiac[zodiac$分组 == '属龙', colnames(zodiac) == '占比'] <- 
    round(nrow(filter(df_orders, zodiac == '属龙')) / nrow(df_orders), 3)
  zodiac[zodiac$分组 == '属蛇', colnames(zodiac) == '人数'] <- 
    nrow(filter(df_orders, zodiac == '属蛇'))
  zodiac[zodiac$分组 == '属蛇', colnames(zodiac) == '占比'] <- 
    round(nrow(filter(df_orders, zodiac == '属蛇')) / nrow(df_orders), 3)
  zodiac[zodiac$分组 == '属马', colnames(zodiac) == '人数'] <- 
    nrow(filter(df_orders, zodiac == '属马'))
  zodiac[zodiac$分组 == '属马', colnames(zodiac) == '占比'] <- 
    round(nrow(filter(df_orders, zodiac == '属马')) / nrow(df_orders), 3)
  zodiac[zodiac$分组 == '属羊', colnames(zodiac) == '人数'] <- 
    nrow(filter(df_orders, zodiac == '属羊'))
  zodiac[zodiac$分组 == '属羊', colnames(zodiac) == '占比'] <- 
    round(nrow(filter(df_orders, zodiac == '属羊')) / nrow(df_orders), 3)
  if(save_all) {
    zodiac <- zodiac %>% arrange(desc(人数))
  } else {
    # 取人数前三
    zodiac <- zodiac %>% slice_max(order_by = 人数, n = 3)
  }
  # 添加到输出表格
  output_table <- bind_rows(output_table, zodiac)
  
  # 各区参保量 ----
  # 添加地区
  district_orders <- tibble('分组' = 'a', '人数' = 0, '占比' = 1) %>% .[-1, ]
  # 根据身份证前6位区分区域
  if(group_district_orders_by_id_no) {
    for (i in names(district_info)) {
      district_orders_tmp <- tibble('分组' = i, '人数' = 0, '占比' = 1)
      district_orders <- bind_rows(district_orders, district_orders_tmp)
    }
    for(i in names(district_info)) {
      district_code <- district_info[i] %>% str_split(", ", simplify = T) %>% as.vector()
      for(j in district_code) {
        district_orders[district_orders$分组 == i, colnames(district_orders) == '人数'] <- 
          district_orders[district_orders$分组 == i, colnames(district_orders) == '人数'] + 
          length(which(df_orders$district == j)) # 将找到的人数相加
      }
    }
    # 计算占比的时候去掉其他地区的订单
    # 不能去掉其他地区，去掉会不严谨（BY明东老师）
    if(F) {
      district_orders <- district_orders %>% 
        mutate(占比 = round(人数 / sum(district_orders$人数), 3))
    }
    district_orders <- district_orders %>% 
      # 添加其他地区
      bind_rows(
        tibble(分组 = '其他地区', 人数 = nrow(df_orders) - sum(district_orders$人数), 占比 = 1)
      ) %>% 
      # 计算占比
      mutate(占比 = round(人数 / sum(人数), 3))
  }
  # 根据一级架构名称区分区域
  if(group_district_orders_by_sale_channel) {
    df_orders_tmp <- df_orders
    for(i in 1:length(district_info)) {
      # 区域名称
      district_name <- district_info[i]
      # 填入区域统计表格
      district_orders[i, 1] <- district_name # 分组
      # 查询所有架构，只要包含区域名称就算到区域订单
      district_orders_tmp <- df_orders_tmp %>% 
        filter_at(vars(contains('channel')), any_vars(str_detect(., !!district_name)))
      if(nrow(district_orders_tmp) == 0) {
        district_orders[i, 2] <- 0 # 人数
        district_orders[i, 3] <- 0 # 占比
      } else {
        district_orders[i, 2] <- nrow(district_orders_tmp) # 人数
        district_orders[i, 3] <- round(nrow(district_orders_tmp) / nrow(df_orders), 3) # 占比
        df_orders_tmp <- df_orders_tmp %>% 
          filter(!col_order_item_id %in% district_orders_tmp$col_order_item_id)
      }
    }
    district_orders[length(district_info) + 1, 1] <- '市区' # 分组
    if(nrow(df_orders_tmp) == 0) {
      district_orders[length(district_info) + 1, 2] <- 0 # 人数
      district_orders[length(district_info) + 1, 3] <- 0 # 占比
    } else {
      district_orders[length(district_info) + 1, 2] <- nrow(df_orders_tmp) # 人数
      district_orders[length(district_info) + 1, 3] <- round(nrow(df_orders_tmp) / nrow(df_orders), 
                                                             3) # 占比
    }
  }
  district_orders <- district_orders %>% arrange(desc(人数))
  
  if(!save_all) {
    # 如果不需要保存所有地区的数据，取其他地区以外的人数前三名的数据
    district_orders <- district_orders %>% 
      filter(分组 != '其他地区') %>% 
      slice_max(order_by = 人数, n = 3)
  }
  
  # 添加到输出表格
  output_table <- bind_rows(output_table, district_orders)
  
  # 医保个账划扣 ----
  if(!is.na(col_medical_insure_flag)) {
    medical_table <- tibble(分组 = '预约医保个张扣缴', 人数 = 0.1, 占比 = 1)
    medical_table[medical_table$分组 == '预约医保个张扣缴', colnames(medical_table) == '人数'] <- 
      nrow(filter(df_orders, col_medical_insure_flag == 1))
    medical_table[medical_table$分组 == '预约医保个张扣缴', colnames(medical_table) == '占比'] <- 
      round(nrow(filter(df_orders, col_medical_insure_flag == 1)) / nrow(df_orders), 3)
    output_table <- bind_rows(output_table, medical_table)
  }
  
  # 支付方式 ----
  if(!is.na(col_pay_way)) {
    payway_table <- df_orders %>% 
      select(col_order_item_id, col_pay_way) %>% 
      group_by(col_pay_way) %>% 
      mutate(人数 = length(col_order_item_id)) %>% 
      ungroup() %>% 
      select(-col_order_item_id) %>% 
      distinct() %>% 
      mutate(占比 = 人数/sum(人数)) %>% 
      rename(分组 = col_pay_way)
    output_table <- bind_rows(output_table, payway_table)
  }
  
  # 为自己投保和为家人投保（人数为空值，放倒数第二） ----
  relationship_table <- tibble(分组 = c('为自己投保', '为家人投保'), 人数 = 0.1, 占比 = 1)
  output_table <- bind_rows(output_table, relationship_table)
  output_table <- as.data.frame(output_table) # to add chr to num col.
  # 企业订单没有相关信息，只用个人订单计算
  if(have_relationship) {
    # 个单模式有记录投保关系
    self_orders <- filter(df_orders, col_group == 'self')
    # 为自己参保
    output_table[output_table$分组 == "为自己投保", colnames(output_table) == '人数'] <- NA
    output_table[output_table$分组 == "为自己投保", colnames(output_table) == '占比'] <- 
      round(
        length(which(self_orders$col_relation_ship == '本人')) / 
          length(self_orders$col_order_item_id), 
        3
      )
    # 为家人参保
    output_table[output_table$分组 == "为家人投保", colnames(output_table) == '人数'] <- NA
    output_table[output_table$分组 == "为家人投保", colnames(output_table) == '占比'] <- 
      round(
        length(which(self_orders$col_relation_ship != '本人')) / 
          length(self_orders$col_order_item_id), 
        3
      )
  } else {
    # 团单模式没有记录投保关系
    # 将custom_id的freq >= 2的当做是为自己和亲人参保，
    # custom_id的freq = 1的当作为自己投保，忽略其他情况
    x <- df_orders$col_customer_id %>% table() %>% as.data.frame()
    # 为家人参保
    output_table[output_table$分组 == "为家人投保", colnames(output_table) == '人数'] <- NA
    output_table[output_table$分组 == "为家人投保", colnames(output_table) == '占比'] <- 
      round(length(which(x$Freq >= 2)) / length(x$.), 3)
    # 为自己参保
    output_table[output_table$分组 == "为自己投保", colnames(output_table) == '人数'] <- NA
    output_table[output_table$分组 == "为自己投保", colnames(output_table) == '占比'] <- 
      round(length(which(x$Freq == 1)) / length(x$.), 3)
  }
  
  # 年龄最大和最小参保人（占比为空值，放倒数第一） ----
  age_table <- tibble(分组 = c('年龄最小参保人', '年龄最大参保人'), 人数 = 0.1, 占比 = 1)
  output_table <- bind_rows(output_table, age_table)
  output_table <- as.data.frame(output_table) # to add chr to num col.
  # 年龄最小参保人
  min_age <- min(na.omit(df_orders$age))
  age_min <- df_orders %>% 
    filter(age == min_age) %>% # 找到参保年龄最小的被保人
    arrange(desc(col_date_created)) %>%  # 如果有不止一人，筛选参保日期最晚的
    .[1, ]
  # 填入表格
  output_table[output_table$分组 == "年龄最小参保人", colnames(output_table) == '人数'] <- 
    paste0(
      "参保年龄：出生第", age_min$age, "天；出生日期：", age_min$col_birthday, "；",
      "性别：", ifelse(age_min$col_sex == '女', '女性', '男性')
    )
  output_table[output_table$分组 == "年龄最小参保人", colnames(output_table) == '占比'] <- NA
  
  # 年龄最大参保人
  max_age <- max(na.omit(df_orders$age))
  age_max <- df_orders %>% 
    filter(age == max_age) %>% # 找到参保年龄最大的被保人
    arrange(col_date_created) %>%  # 如果有不止一人，筛选参保日期最早的
    .[1, ]
  
  # 填入表格
  output_table[output_table$分组 == "年龄最大参保人", colnames(output_table) == '人数'] <- 
    paste0(
      "参保年龄：", round(age_max$age / 365), "岁；出生日期：", age_max$col_birthday, "；",
      "性别：", ifelse(age_max$col_sex == '女', '女性', '男性')
    )
  output_table[output_table$分组 == "年龄最大参保人", colnames(output_table) == '占比'] <- NA
  
  # Sheet 2: 参保用户姓氏数量 ----
  df_name <- df_orders %>% 
    select(col_name) %>% 
    mutate(
      col_first_name = str_replace(col_name, "(.).*", "\\1"),
      col_total_count = length(col_name)
    ) %>% 
    group_by(col_first_name) %>% 
    mutate(col_first_name_count = length(col_first_name)) %>% 
    mutate(col_percent = round(col_first_name_count / col_total_count, 3)) %>% 
    ungroup() %>% 
    arrange(desc(col_first_name_count)) %>% 
    select(姓氏 = col_first_name, 数量 = col_first_name_count, 占比 = col_percent) %>% 
    distinct()
  # # Add sheet
  # sht_name = addWorksheet(df_orders_wb, '投保用户姓氏统计')
  # # Write data
  # writeData(df_orders_wb, sheet = sht_name, x = df_name)
  # # Add the percent style to the desired cells
  # addStyle(
  #   wb = df_orders_wb, sheet = sht_name, style = pct, 
  #   cols = (1:ncol(df_name))[str_detect(colnames(df_name), '占比')],
  #   rows = 1:(nrow(df_name) + 1),
  #   gridExpand = T
  # )
  column_to_scale <- colnames(df_name)[str_detect(colnames(df_name), '占比|率')]
  for(column in column_to_scale) {
    df_name[[column]] <- scales::label_percent(
      accuracy = 0.1, big.mark = ""
    )(
      as.numeric(df_name[[column]])
    ) %>% 
      str_replace('Inf', '')
  }
  
  # Sheet 3: 每小时平均参保数据 ----
  df_hour <- df_orders %>% 
    select(col_datetime_created) %>% 
    mutate(
      col_date = as.Date(col_datetime_created, tz = Sys.timezone()),
      col_hour = str_replace(col_datetime_created, "....-..-..\\s(..)\\:.*", "\\1"),
      col_total_count = length(col_datetime_created)
    )
  # 按日期统计每小时投保量，找到最大最小值
  df_date_hour <- df_hour %>% 
    group_by(col_date, col_hour) %>% 
    mutate(col_date_hour_count = length(col_datetime_created)) %>% 
    ungroup() %>% 
    select(col_date, col_hour, col_date_hour_count) %>% 
    distinct()
  df_date_hour_min <- df_date_hour %>% 
    group_by(col_hour) %>% 
    arrange(col_date_hour_count) %>% 
    slice(n = 1) %>% 
    select(小时 = col_hour, 最小值 = col_date_hour_count, 最小值日期 = col_date)
  df_date_hour_max <- df_date_hour %>% 
    group_by(col_hour) %>% 
    arrange(desc(col_date_hour_count)) %>% 
    slice(n = 1) %>% 
    select(小时 = col_hour, 最大值 = col_date_hour_count, 最大值日期 = col_date)
  # 统计销售期全长每小时平均投保量
  df_hour <- df_hour %>% 
    group_by(col_hour) %>% 
    mutate(col_hour_count = length(col_hour)) %>% 
    mutate(col_percent = round(col_hour_count / col_total_count, 3)) %>% 
    ungroup() %>% 
    arrange(col_hour) %>% 
    select(小时 = col_hour, 数量 = col_hour_count, 占比 = col_percent) %>% 
    distinct() %>% 
    # 加上最大值最小值
    left_join(df_date_hour_max) %>% 
    left_join(df_date_hour_min)
  # # Add sheet
  # sht_hour = addWorksheet(df_orders_wb, '每小时平均参保数据')
  # # Write data
  # writeData(df_orders_wb, sheet = sht_hour, x = df_hour)
  # # Add the percent style to the desired cells
  # addStyle(
  #   wb = df_orders_wb, sheet = sht_hour, style = pct, 
  #   cols = (1:ncol(df_hour))[str_detect(colnames(df_hour), '占比')],
  #   rows = 1:(nrow(df_hour) + 1),
  #   gridExpand = T
  # )
  column_to_scale <- colnames(df_hour)[str_detect(colnames(df_hour), '占比|率')]
  for(column in column_to_scale) {
    df_hour[[column]] <- scales::label_percent(
      accuracy = 0.1, big.mark = ""
    )(
      as.numeric(df_hour[[column]])
    ) %>% 
      str_replace('Inf', '')
  }
  
  # 输出excel ----
  # # Formatting percentages to excel in R: https://stackoverflow.com/a/48066298/10341233
  # # Add data to the worksheet we just created
  # writeData(df_orders_wb, sheet = sht1, x = output_table)
  # # Add the percent style to the desired cells
  # addStyle(
  #   wb = df_orders_wb, sheet = sht1, style = pct, 
  #   cols = (1:ncol(output_table))[str_detect(colnames(output_table), '占比')],
  #   rows = 1:(nrow(output_table) + 1),
  #   gridExpand = T
  # )
  column_to_scale <- colnames(output_table)[str_detect(colnames(output_table), '占比|率')]
  for(column in column_to_scale) {
    output_table[[column]] <- scales::label_percent(
      accuracy = 0.1, big.mark = ""
    )(
      as.numeric(output_table[[column]])
    ) %>% 
      str_replace('Inf', '')
  }
  # if(save_result) {
  #   # 创建文件夹
  #   if( !dir.exists(paste0(sop_dir, '/', city)) ) {
  #     dir.create(paste0(sop_dir, '/', city))
  #   }
  #   if( !dir.exists(paste0(sop_dir, '/', city, '/', as.Date(end_time, tz = Sys.timezone()))) ) {
  #     dir.create(paste0(sop_dir, '/', city, '/', as.Date(end_time, tz = Sys.timezone())))
  #   }
  #   # 输出excel
  #   saveWorkbook(
  #     df_orders_wb, 
  #     file = paste0(
  #       sop_dir, '/', city, '/', as.Date(end_time, tz = Sys.timezone()), '/', 
  #       city, '投保数据统计-截止至', 
  #       lubridate::year(end_time), '年', 
  #       lubridate::month(end_time), '月', 
  #       lubridate::day(end_time), '日', 
  #       lubridate::hour(end_time), '时.xlsx'
  #     ), 
  #     overwrite = T
  #   )
  # }
  output_list[['投保数据统计']] <- output_table
  output_list[['每小时平均参保数据']] <- df_hour
  output_list[['投保用户姓氏统计']] <- df_name
  output_list[['区域参保量']] <- district_orders %>% 
    mutate(
      占比 = scales::label_percent(
        accuracy = 0.1, big.mark = ""
      )(
        as.numeric(district_orders[['占比']])
      )
    )
  return(output_list)
  
  # 单独保存区域订单统计表
  if(F) {
    # 创建文件夹
    if( !dir.exists(paste0(sop_dir, '/', city)) ) {
      dir.create(paste0(sop_dir, '/', city))
    }
    if( !dir.exists(paste0(sop_dir, '/', city, '/', as.Date(end_time, tz = Sys.timezone()))) ) {
      dir.create(paste0(sop_dir, '/', city, '/', as.Date(end_time, tz = Sys.timezone())))
    }
    # Sheet1
    colnames(district_orders)[colnames(district_orders) == '分组'] <- '区域'
    wb <- createWorkbook()
    sht1 = addWorksheet(wb, '区域投保量统计')
    writeData(wb, sheet = sht1, x = district_orders)
    pct = createStyle(numFmt = "0.0%")
    addStyle(
      wb = wb, sheet = sht1, style = pct, 
      cols = (1:ncol(district_orders))[str_detect(colnames(district_orders), '占比')],
      rows = 1:(nrow(district_orders) + 1),
      gridExpand = T
    )
    # Write excel
    if(length(names(district_info)) != 0) {
      method_name <- '身份证前六位匹配'
    } else {
      method_name <- '架构名称与区域匹配'
    }
    saveWorkbook(
      wb, 
      file = paste0(
        sop_dir, '/', city, '/', as.Date(end_time, tz = Sys.timezone()), '/', 
        '区域投保量统计-', method_name, '-截止至', 
        lubridate::year(end_time), '年', 
        lubridate::month(end_time), '月', 
        lubridate::day(end_time), '日', 
        lubridate::hour(end_time), '时',
        '.xlsx'
      ), 
      overwrite = T
    )
  }
}
