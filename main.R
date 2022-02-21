############################ Main script of Standard Operating Procedure ##########################
# 一、获取输入参数 ----
# Get option from optparse
if(T) {
  library(optparse)
  opt = optparse::parse_args(
    OptionParser(
      option_list = list(
        # 需要统计数据的项目
        make_option(
          c("--city-list"), type = "character", default = NA,
          help = paste0(
            "Required. City list, input multiple cities in format like: --city-list '广州','惠州'."
          )
        ),
        # 对应的产品保障年度
        make_option(
          c("--year-list"), type = "character", default = NA,
          help = paste0(
            "Optional. Project insurance period that pairs with city list, used in filtering ",
            "sku_code. Input format if has multiple years: --year-list '2021','2021'. If no ",
            "inputs, year-list will take current year as project insurance period for every city ",
            "automatically."
          )
        ),
        # 产品销售起始日期（可精确到时分秒）
        make_option(
          c("--start-time-list"), type = "character", default = NA,
          help = paste0(
            "Optional. Project selling start time that pairs with city list, used in filtering ",
            "orders by date_created, start time will be included in final orders. ",
            "It can take multiple inputs including empty value and be accurate to seconds in ", 
            "format like: --start-time-list '2020-11-01 00:00:00',''. ",
            "If no inputs, start-time will take values stored in this script automatically."
          )
        ),
        # 产品销售结束日期（可精确到时分秒）
        make_option(
          c("--end-time-list"), type = "character", default = NA,
          help = paste0(
            "Optional. Project selling end time that pairs with city list, used in filtering ",
            "orders by date_created, end time will NOT be included in final orders. ",
            "It can take multiple inputs including empty value and be accurate to seconds in ",
            "format like: --end-time-list '2021-02-01 00:00:00',''. ",
            "If no inputs, end-time will take values stored in this script automatically."
          )
        ),
        # 固定小时
        make_option(
          c("--hr"), type = 'character', default = NA,
          help = paste0(
            "Optional. Set end_time to a fixed hour. This is useful when the same end_time is ",
            "required in running multiple cities' routine SOP job, since they run in order. ",
            "Input 2 digits number such as '12' and end_time will be set to xxxx-xx-xx 12:00:00."
          )
        ),
        # 是否提取退单的原因
        make_option(
          c("--refund-status"), type = "logical", default = FALSE
        ),
        # 是否需要特殊处理（例如提取某个城市去年的订单）
        make_option(
          c("--special-processing"), type = "logical", default = FALSE
        ),
        # 是否保存RData
        make_option(
          c("--save-query"), type = "logical", default = FALSE
        ),
        # 是否有额外需求（例如单独发送保司出单情况）
        make_option(
          c("--extra-requirement"), type = "logical", default = FALSE
        ),
        # 有额外需求的城市列表
        make_option(
          c("--extra-requirement-city-list"), type = "character", default = NA,
          help = paste0(
            "Optional. City list that have extra requirement, input multiple cities in format ",
            "like: --city-list '广州','惠州'."
          )
        ),
        # 是否发送邮件
        make_option(
          c("--send-sop-email"), type = "logical", default = FALSE
        ),
        # 是否生成标准数据报告PDF
        make_option(
          c("--standard-report"), type = "logical", default = FALSE
        ),
        # 是否为测试run
        make_option(
          c("-t", "--test"), type = "logical", default = TRUE
        )
      )
    )
  )
}
#saveRDS(opt, 'opt.RDS')
# 判断参数是否输入正确
if(T) {
  # 1.城市
  # 首先拆分输入的城市列表
  city_list <- as.character(stringr::str_split(opt$`city-list`, ',', simplify = TRUE))
  # # 如果城市列表含有NA，则报错
  # if(length(which(is.na(opt$`city-list`))) != 0) {
  #   stop("City list contains NA, please check again.")
  # }
  
  # 2.年份
  # 首先拆分输入的年份
  year_list <- as.character(stringr::str_split(opt$`year-list`, ',', simplify = TRUE))
  # 如果年份列表只有一个且为NA或空，即为没有输入year-list
  # 直接取今年当作项目的保障年度year_list
  if(length(year_list) == 1) {
    if(is.na(year_list) | year_list == '') {
      year_list <- replicate(length(city_list), lubridate::year(Sys.Date()))
    }
  } else {
    # 如果年份有空值，则直接取今年当作项目的保障年度
    for(y in 1:length(year_list)) {
      if(year_list[y] == '') { year_list[y] <- lubridate::year(Sys.Date()) }
    }
    rm(y)
  }
  # 如果年份的数量和城市不一样则报错
  if(length(year_list) != length(city_list)) {
    stop(
      paste0(
        "The number of year_list is not equal to the number of city_list, please check again."
      )
    )
  }
  
  # 3.销售起始日期
  # 首先拆分输入的销售起始日期
  start_times <- as.character(stringr::str_split(opt$`start-time-list`, ',', simplify = TRUE))
  # 如果销售起始日期只有一个且为NA或空，即为没有输入start-time，所以直接用2.项目信息里的起始日期
  if(length(start_times) == 1) {
    if(is.na(start_times) | start_times == '') { start_times <- NA }
  } else {
    # 如果输入的销售起始日期为空，则转换成NA
    for(st in 1:length(start_times)) { if(start_times[st] == '') { start_times[st] <- NA } }
    rm(st)
    # 如果输入了销售起始日期，但数量和城市不一样，报错
    if(length(start_times) != length(city_list)) {
      stop(
        paste0(
          "The number of start_times is not equal to the number of city_list, please check again."
        )
      )
    }
  }
  
  # 4.销售结束日期
  # 首先拆分输入的销售结束日期
  end_times <- as.character(stringr::str_split(opt$`end-time-list`, ',', simplify = TRUE))
  # 如果销售结束日期只有一个且为NA或空，即为没有输入end-time，所以直接用2.项目信息里的结束日期
  if(length(end_times) == 1) {
    if(is.na(end_times) | end_times == '') { end_times <- NA }
  } else {
    # 如果输入的销售结束日期为空，则转换成NA
    for(et in 1:length(end_times)) { if(end_times[et] == '') { end_times[et] <- NA } }
    rm(et)
    # 如果输入了销售结束日期，但数量和城市不一样，报错
    if(length(end_times) != length(city_list)) {
      stop(
        paste0(
          "The number of end_times is not equal to the number of city_list, please check again."
        )
      )
    }
  }
  
  # 5.有额外需求的城市列表
  # 该列表是为了在同时运行多个项目的时候控制哪些项目运行额外需求的模块
  if(is.na(opt$`extra-requirement-city-list`) | opt$`extra-requirement-city-list` == '') {
    # 没有指定额外需求的城市，则设为所有统计数据的城市
    extra_requirement_city_list <- city_list
  } else {
    # 有指定额外需求的城市
    extra_requirement_city_list <- as.character(
      stringr::str_split(opt$`extra-requirement-city-list`, ',', simplify = TRUE)
    )
  }
  
  # # 6.邮件附件列表
  # email_attachments <- NULL
  
  # 7.如果某个项目的SOP压缩文件没有生成，将保存于此
  missing_files <- NULL
}

# 二、环境 ----
# 加载包，指定工作路径
if(T) {
  # Load packages
  library(bit64) # MySQL取出的数字是interger64格式，需要手动加载bit64包，否则会乱码
  library(rmarkdown)
  library(rticles)
  library(RMariaDB)
  library(emayili)
  library(logr)
  library(zip)
  library(glue)
  library(ggsci)
  library(scales)
  library(readxl)
  library(writexl)
  library(openxlsx)
  library(lubridate)
  library(tidyverse)
  
  # Working directory
  if(dir.exists('/data/users/yeguanhua/SOP')) {
    sop_dir <- '/data/users/yeguanhua/SOP'
  } else {
    sop_dir <- getwd()
  }
  setwd(sop_dir)
}
# log
if(T) {
  # If working directory is in server, open log file
  if(sop_dir == '/data/users/yeguanhua/SOP') {
    logr::log_open(
      file_name = paste0(
        'log_', 
        lubridate::year(Sys.time()), '年', 
        lubridate::month(Sys.time()), '月', 
        lubridate::day(Sys.time()), '日', 
        lubridate::hour(Sys.time()), '时'
      )
    )
  }
}
# 通用信息提取
if(T) {
  # 数据库
  product <- dbConnect(
    drv = RMariaDB::MariaDB(), 
    user = '***',
    password = '***',
    host = '***',
    port = 123,
    dbname = '***',
    groups = '***'
  )
  
  # 渠道信息
  channel_cls123 <- dbGetQuery(
    product, 
    glue(
      "
SELECT 
    channel_encode sale_channel_id,
    # 提取SQL日期到R时需要转成字符串，否则time-zone会混乱
    CAST( gmt_created AS CHAR ) channel_date_created,
		IF(
				`level` = 3,
				SUBSTRING_INDEX( ancestor_name, ',', 1 ),
				IF(
						`level` = 2,
						ancestor_name,
						channel_name
				)
		) channel_cls1,
		IF(
				`level` = 3,
				SUBSTRING_INDEX( ancestor_name, ',', -1 ),
				channel_name
		) channel_cls2,
		# 去掉三级渠道井号前的字符
		CASE
		    WHEN (LOCATE('#', channel_name) > 0) THEN 
				    SUBSTR(channel_name, LOCATE('#', channel_name) + 1)
		    ELSE 
				    channel_name
    END channel_cls3
FROM 
    glk_channel
WHERE
    (
    ancestor_name NOT LIKE '%思派%药房%' 
    AND ancestor_name NOT LIKE '%健1宝%' 
    AND ancestor_name NOT LIKE '%派服务%' 
    AND ancestor_name NOT LIKE '%远通%'
    )
    OR ancestor_name IS NULL
# 不要限定时间，因为代理人会在开售前就注册，而且总量不大直接全部取
        "
    )
  ) %>% 
    # 去掉time-zone
    mutate(channel_date_created = str_replace(channel_date_created, '(.*)\\sCST', '\\1')) %>% 
    # 转换日期格式
    mutate(channel_date_created = as.POSIXct(channel_date_created)) %>% 
    # 去掉测试渠道
    filter(across(.cols = everything(), .fns = ~ str_detect(., "测试", negate = T)))
  
  # 企业信息
  company_info <- dbGetQuery(
    product,
    glue(
      "
SELECT
    com_id, 
    business_name,
    CASE
        WHEN business_type = '1' THEN '企业'
        WHEN business_type = '2' THEN '事业单位'
        WHEN business_type = '3' THEN '基金会'
        WHEN business_type = '0' THEN '社会组织'
        WHEN business_type = '8' THEN '办事处'
        WHEN business_type = '9' THEN '个体户'
        WHEN business_type = '10' THEN '行业协会'
        WHEN business_type = '99' THEN '其他'
    END business_type,
    credit_code credit_code,
    name business_contact,
    phone_num contact_tel,
    mail contact_mail,
    address business_address
FROM 
    company_info
# 企业信息不能限定时间，如果不是第一年的项目那企业信息可能之前就注册过了
        "
    )
  )
  
  # 退单原因
  if(opt$`refund-status`) {
    refund <- dbGetQuery(
      product,
      glue(
        "
SELECT
    pk_order_item order_item_id,
    CAST( MAX( date_updated ) AS CHAR ) date_updated_refund_remark,
    CAST( remark AS CHAR ) remark
FROM 
    order_trade_record
WHERE 
    `event` = '07' # 07为撤契，08为取消
    #AND date_created >= '{start_time}' 
    #AND date_created < '{end_time}' 
    AND remark NOT LIKE '%测试%'
    AND remark NOT LIKE '%周文怡要求退款%'
GROUP BY 
    pk_order_item
          "
      )
    ) %>% 
      # 去掉time-zone
      mutate(
        date_updated_refund_remark = str_replace(date_updated_refund_remark, '(.*)\\sCST', '\\1')
      ) %>% 
      # 转换日期格式
      mutate(
        date_updated_refund_remark = as.POSIXct(date_updated_refund_remark)
      )
    # refund <- refund %>% 
    #  # 取更新时间最新的退款原因，否则用left_join()会导致个人订单出现重复
    #  group_by(order_item_id) %>% 
    #  #slice_max(order_by = date_updated_refund_remark, n = 1) %>% 
    #  # slice无法根据日期排序所以分两步做
    #  arrange(date_updated_refund_remark, .by_group = T) %>% 
    #  slice_tail(n = 1) %>% 
    #  ungroup()
    # 退款原因分成七大类
    refund$remark <- ifelse(str_detect(refund$remark, '用户发起'), '用户发起', refund$remark)
    refund$remark <- ifelse(str_detect(refund$remark, '错误'), '信息输入错误', refund$remark)
    refund$remark <- ifelse(
      str_detect(refund$remark, '不符合参保条件|没有.*[医保|参保|户口|社保]'),
      '没有本地医保，无法购买',
      refund$remark
    )
    refund$remark <- ifelse(
      str_detect(refund$remark, '不懂|[不]?理解|[不]?知道|[不]?喜欢'),
      '看不懂产品内容，不喜欢',
      refund$remark
    )
    refund$remark <- ifelse(
      str_detect(
        refund$remark, 
        '[岁|穗][岁|穗][康|保]|其[它|他][产品|保险]|别[家|的]|[类似|同类]'
      ),
      '别家的产品更适合我',
      refund$remark
    )
    refund$remark <- ifelse(
      str_detect(refund$remark, glue('国外|不在[{city}|中国|国内]')),
      '其他',
      refund$remark
    )
    refund$remark <- ifelse(
      str_detect(
        refund$remark, 
        '死|去世|逝|过世|病故|身故|生故|不在.*人世|人不在|[已经]?走了'
      ),
      '被保险人去世',
      refund$remark
    )
    refund$remark <- ifelse(
      str_detect(
        refund$remark, 
        paste0(
          '用户发起|',
          '信息输入错误|',
          '没有本地医保，无法购买|',
          '看不懂产品内容，不喜欢|',
          '别家的产品更适合我|',
          '被保险人去世'
        ),
        negate = TRUE
      ),
      '其他',
      refund$remark
    )
    refund <- refund %>% select(order_item_id, refund_remark = remark)
  }
  
  # 清除SQL数据，防止后续爆内存
  dbDisconnect(product)
  rm(product)
}

# 对每个城市做数据统计
for(c in 1:length(city_list)) {
  city <- city_list[c]
  year <- year_list[c]
  print(
    glue(
      "
-------------------------------------------------------
 Starting {city}{year} SOP project at {Sys.time()}.
-------------------------------------------------------
      "
    )
  )
  
  # 三、项目信息 ----
  if(T) {
    # 个人skucode
    company_self <- read_xlsx(
      path = paste0(sop_dir, '/', 'skucode_all.xlsx'), 
      col_types = c('text', 'text', 'text', 'text', 'text', 'text')
    ) %>% 
      filter(city == !!city) %>% 
      filter(year == !!year) %>% 
      filter(group == '个人') %>% 
      select(sku_code, company)
    if(nrow(company_self) == 0) {
      stop(
        paste0("Can't find any company with SELF sku_code of ", city, year, ", ",
               "please check skucode_all.xlsx file.")
      )
    }
    sku_code_self <- paste0("'", paste0(company_self$sku_code, collapse = "', '"), "'")
    
    # 企业skucode
    company_com <- read_xlsx(
      path = paste0(sop_dir, '/', 'skucode_all.xlsx'), 
      col_types = c('text', 'text', 'text', 'text', 'text', 'text')
    ) %>% 
      filter(city == !!city) %>% 
      filter(year == !!year) %>% 
      filter(group == '企业') %>% 
      select(sku_code, company)
    if(nrow(company_com) == 0) {
      warning(
        paste0("Can't find any company with COM sku_code of ", city, year, ", ",
               "please check skucode_all.xlsx file.")
      )
    }
    sku_code_com <- paste0("'", paste0(company_com$sku_code, collapse = "', '"), "'")
    
    # 一级渠道信息
    df_channel_cls1 <- read_xlsx(
      path = paste0(sop_dir, '/', 'channel_cls1.xlsx'),
      col_types = c('text', 'text', 'text', 'text')
    ) %>% 
      filter(city == !!city, year == !!year)
    
    if(city == '徐州') {
      project_name <- '惠徐保'
      if(year == '2021') {
        start_time <- '2020-10-27 00:00:00'
        end_time <- '2021-02-01 00:00:00'
        online_sku_code <- c('00000013', '00000133')
        ended <- TRUE
      }
      if(year == '2022') {
        start_time <- '2021-10-20 00:00:00'
        end_time <- '2022-02-01 00:00:00'
        online_sku_code <- c('200001582', '200001583', '20001719', '200001693')
        ended <- FALSE
      }
      price <- "'6900'"
      have_relationship <- TRUE
      district <- c(
        # 用“, ”分开同一区中不同的身份证区号
        # 江苏省徐州市、江苏省徐州市鼓楼区、江苏省徐州市云龙区、江苏省徐州市泉山区合并为市区
        '江苏省徐州市市区' = '320300, 320301, 320302, 320303, 320311', 
        '江苏省徐州市贾汪区' = '320305', 
        '江苏省徐州市铜山区' = '320312, 320323, 320304', 
        '江苏省徐州市丰县' = '320321', 
        '江苏省徐州市沛县' = '320322', 
        '江苏省徐州市睢宁县' = '320324', 
        '江苏省徐州市新沂市' = '320381, 320326',
        '江苏省徐州市邳州市' = '320382, 320325'
      )
      
    }
    
    # 如果查询已结束的项目，则忽略online_sku_code
    if(ended) { online_sku_code <- c() }
    
    # 判断是否有外部输入的起始日期
    if(!is.na(start_times)) {
      start_time_tmp <- start_times[c]
      # 判断与city匹配的起始日期是否为NA
      if(!is.na(start_time_tmp)) {
        # 如果不为NA则用外部输入的日期
        start_time <- start_time_tmp
      }
      rm(start_time_tmp)
    }
    
    # 判断是否有外部输入的结束日期
    if(!is.na(end_times)) {
      end_time_tmp <- end_times[c]
      # 判断与city匹配的结束日期是否为NA
      if(!is.na(end_time_tmp)) {
        # 如果不为NA则用外部输入的日期
        end_time <- end_time_tmp
      }
      rm(end_time_tmp)
    }
    if(!is.na(opt$hr) & opt$hr != '') {
      end_time <- str_replace(end_time, "(....-..-..\\s).*", glue("\\1{opt$hr}:00:00"))
    }
    
    # 文件名的时间
    sop_datetime <- paste0(
      lubridate::year(end_time), '年', 
      lubridate::month(end_time), '月', 
      lubridate::day(end_time), '日', 
      lubridate::hour(end_time), '时'
    )
    
    # 如果目标目录为空则以city为名称创建第一层文件夹
    if( !dir.exists(glue("{sop_dir}/{city}")) ) {
      dir.create(glue("{sop_dir}/{city}"))
    }
    
    # 如果目标目录为空则以sop_datetime为名称创建第二层文件夹，所有数据文件保存在此
    if( !dir.exists(glue("{sop_dir}/{city}/{sop_datetime}")) ) {
      dir.create(glue("{sop_dir}/{city}/{sop_datetime}"))
    }
    proj_dir <- glue("{sop_dir}/{city}/{sop_datetime}")
    
    # 如果需要发送SOP邮件，提前将第6步生成的压缩文件提前添加到附件列表
    # if(opt$`send-sop-email`) {
    #   email_attachments <- c(
    #     email_attachments,
    #     glue("{sop_dir}/{city}/{sop_datetime}/{city}项目统计数据-截止至{sop_datetime}.zip")
    #   )
    # }
  }
  
  print(
    glue(
      "
-------------------------------------------------------
 The sop_datetime of {city}{year} is {sop_datetime}.
-------------------------------------------------------
      "
    )
  )
  
  # 四、生成总订单及信息 ----
  if(T) {
    start.time <- Sys.time()
    # 4.1 数据库 ----
    product <- dbConnect(
      drv = RMariaDB::MariaDB(), 
      user = '***',
      password = '***',
      host = '***',
      port = 123,
      dbname = '***',
      groups = '***'
    )
    
    # 4.2 个人订单 ----
    # 基础表格 (order_item) ----
    self_orders <- dbGetQuery(
      product, 
      glue(
        "
SELECT 
    id order_item_id, 
    CAST( date_created AS CHAR ) date_created, # 必须先转成字符串，否则取到R中time-zone会混乱
    CAST( date_updated AS CHAR ) date_updated, # 必须先转成字符串，否则取到R中time-zone会混乱
    sku_code, 
    SUBSTR( sale_channel_id, 1, 16 ) sale_channel_id, # 渠道码只有16位
    CASE
		    WHEN automatic_deduction = 1 THEN 1 ELSE 0 # 是否签约自动扣费
    END automatic_deduction, 
    CASE
		    WHEN automatic_deduction_is_success = 1 THEN 1 ELSE 0 # 自动扣费是否成功
    END automatic_deduction_is_success, 
    CASE
		    WHEN is_automatic_deduction = 1 THEN 1 ELSE 0 # 是否为自动扣费生成的订单
    END is_automatic_deduction, 
    customer_id,
    agent_no, # 平安底下还有一个四级渠道，为agent_no
    pk_main_order,
    sub_order_no,
    CASE 
        WHEN medical_insure_flag = 1 THEN 1 ELSE 0 # 转换成数字方便后续统计
    END medical_insure_flag,
    CAST( payment_amount AS CHAR ) payment_amount,
    CASE
        WHEN payway = '00' THEN '微信'
        WHEN payway = '02' THEN '支付宝'
        # 医保个账：07-徐州2021，14-芜湖2021，15-江门2021，16-徐州2022
        WHEN payway IN ('07','14','15','16') THEN '医保个账'
        ELSE '其他'
    END pay_way,
    CASE
        WHEN refund_status >= '01' THEN '04' ELSE '00'
    END policy_order_status,
    'self' AS `group`
FROM 
    order_item 
WHERE 
    order_status != '02'
    AND payment_status = '01'
    AND payment_amount IN ({price})
    AND date_created >= '{start_time}'
    AND date_created < '{end_time}'
    AND sku_code IN ({sku_code_self})
        "
      )
    ) %>% 
      # 去掉time-zone
      mutate(
        date_created = str_replace(date_created, '(.*)\\sCST', '\\1'),
        date_updated = str_replace(date_updated, '(.*)\\sCST', '\\1')
      ) %>% 
      # 转换日期格式
      mutate(
        date_created = as.POSIXct(date_created), 
        date_updated = as.POSIXct(date_updated)
      )
    # 添加被保险人信息 (order_service_obj) ----
    self_orders <- self_orders %>% 
      left_join(
        # order_service_obj
        dbGetQuery(
          product, 
          glue(
            "
SELECT 
    pk_order_item order_item_id, 
    name,
    CASE
        WHEN id_type = 1 THEN '身份证'
        WHEN id_type = 2 THEN '护照'
        WHEN id_type = 7 THEN '港澳居民来往内地通行证'
        WHEN id_type = 8 THEN '台湾居民来往大陆通行证'
    END id_type,
    UPPER( id_no ) id_no, 
    CASE  
        WHEN sex IS NULL 
        AND LENGTH( id_no ) = 18 
        AND MOD ( substr( id_no, 17, 1 ), 2 ) = 1 THEN 
            '男' 
        WHEN sex IS NULL 
        AND LENGTH( id_no ) = 18 
        AND MOD ( substr( id_no, 17, 1 ), 2 ) = 0 THEN 
            '女' 
        ELSE 
            REPLACE ( REPLACE ( sex, 'F', '男' ), 'W', '女' ) 
    END sex,
    CASE  
        WHEN birthday = '' 
        AND LENGTH( id_no ) = 18 THEN 
            STR_TO_DATE( substr( id_no, 7, 8 ), '%Y%m%d' ) 
        ELSE
            birthday 
    END birthday,
    tel,
    CASE
        WHEN relation_ship = 1 THEN '本人'
        WHEN relation_ship = 2 THEN '父母'
        WHEN relation_ship = 3 THEN '配偶'
        WHEN relation_ship = 4 THEN '子女'
        ELSE '其他'
    END relation_ship
FROM 
    order_service_obj
WHERE 
    date_created >= '{start_time}' 
    AND date_created < '{end_time}' # 限定时间，否则提取太慢！
            "
          )
        )
      )
    # 添加投保人信息 (order_insurance_applicant) ----
    self_orders <- self_orders %>% 
      left_join(
        # order_insurance_applicant
        dbGetQuery(
          product,
          glue(
            "
SELECT 
    pk_main_order,
    open_id,
    name applicant_name,
    tel applicant_tel,
    CASE
        WHEN id_type = 1 THEN '身份证'
        WHEN id_type = 2 THEN '护照'
        WHEN id_type = 7 THEN '港澳居民来往内地通行证'
        WHEN id_type = 8 THEN '台湾居民来往大陆通行证'
    END applicant_id_type,
    UPPER( id_no ) applicant_id_no,
    CASE  
        WHEN sex IS NULL 
        AND LENGTH( id_no ) = 18 
        AND MOD ( substr( id_no, 17, 1 ), 2 ) = 1 THEN 
            '男' 
        WHEN sex IS NULL 
        AND LENGTH( id_no ) = 18 
        AND MOD ( substr( id_no, 17, 1 ), 2 ) = 0 THEN 
            '女' 
        ELSE 
            REPLACE ( REPLACE ( sex, 'F', '男' ), 'W', '女' ) 
    END applicant_sex,
    CASE  
        WHEN birthday = '' 
        AND LENGTH( id_no ) = 18 THEN 
            STR_TO_DATE( substr( id_no, 7, 8 ), '%Y%m%d' ) 
        ELSE
            birthday 
    END applicant_birthday
FROM 
    order_insurance_applicant
WHERE 
    date_created >= '{start_time}' 
    AND date_created < '{end_time}' # 限定时间，否则提取太慢！
            "
          )
        )
      ) %>% 
      # 去掉测试单
      filter(str_detect(name, '测试', negate = T))
    # 添加订单承保状态 (policy_order) ----
    order_policy <- tryCatch(
      # expr
      { order_policy },
      # error condition
      error = function(e) {
        print("Cannot find order_policy, set order_policy to FALSE.")
        FALSE
      }
    )
    if(order_policy) {
      self_orders <- self_orders %>% 
        left_join(
          # policy_order
          dbGetQuery(
            product, 
            glue(
              "
SELECT 
    pk_order_item order_item_id, 
    policy_company, # 承保保司
    policy_no, # 承保单号
    CAST( effective_date AS CHAR ) effective_date, # 保单起期
    CAST( expiry_date AS CHAR ) expiry_date # 保单止期
FROM 
    policy_order 
WHERE 
    date_created >= '{start_time}' 
    AND date_created < '{end_time}' # 限定时间，否则提取太慢！
              "
            )
          ) %>% 
            # 去掉time-zone
            mutate(
              effective_date = str_replace(effective_date, '(.*)\\sCST', '\\1'),
              expiry_date = str_replace(expiry_date, '(.*)\\sCST', '\\1')
            ) %>% 
            # 转换日期格式
            mutate(
              effective_date = as.POSIXct(effective_date), 
              expiry_date = as.POSIXct(expiry_date)
            )
        )
    }
    # 其他筛选 ----
    # 把没有sale_channel_id的订单挑出来再添加渠道信息，否则会造成订单重复！！！
    self_orders_no_channel <- filter(self_orders, is.na(sale_channel_id))
    self_orders <- filter(self_orders, !order_item_id %in% self_orders_no_channel$order_item_id)
    # 个人订单添加渠道信息
    self_orders <- self_orders %>% left_join(channel_cls123)
    # 对于匹配不上渠道的sale_channel_id的处理
    if(F) {
      # 剩下有sale_channel_id的，找不到一级机构的订单，就用sale_channel_id的前10/12/16位尝试
      for(i in c(10, 12, 16)) {
        self_orders_tmp <- self_orders
        if(nrow(self_orders_tmp[which(is.na(self_orders_tmp$channel_cls1)), ]) != 0) {
          self_no_channel1 <- self_orders_tmp[which(is.na(self_orders_tmp$channel_cls1)), ]
          ## 1.1.1-用sale_channel_id的前i位
          self_head_tmp <- self_orders_tmp %>% 
            # 取出没有一级机构的订单
            filter(order_item_id %in% self_no_channel1$order_item_id) %>% 
            select(-c(channel_cls1, channel_cls2, channel_cls3)) %>% 
            # 取sale_channel_id的前i位
            mutate(sale_channel_id = str_sub(sale_channel_id, end = i)) %>% 
            # 与channel_cls123的sale_channel_id前i位做匹配
            left_join(channel_cls123 %>% mutate(sale_channel_id = str_sub(sale_channel_id, end = i)))
          # 如果left_join前后行数一致，说明join key没有重复。
          # 如果X表中的key在Y中有多个，join完行数会变多。
          if(length(self_no_channel1$order_item_id) != length(self_head_tmp$order_item_id)) {
            warning(paste0("个人订单一级机构ID前", i, "位有重复！"))
          }
          self_head_tmp <- filter(self_head_tmp, !is.na(channel_cls1))
          self_orders <- self_orders %>% 
            filter(!order_item_id %in% self_head_tmp$order_item_id) %>% 
            bind_rows(self_head_tmp) %>% 
            arrange(desc(date_created))
        }
      }
      # 清除SQL数据，防止后续爆内存
      rm(self_orders_tmp, self_head_tmp, self_no_channel1)
      gc()
    }
    # 把没有sale_channel_id的订单加回去
    self_orders <- bind_rows(self_orders, self_orders_no_channel)
    # 为平安添加四级渠道
    self_orders <- self_orders %>% 
      # agent_no前面加上“_”
      mutate(agent_no = str_replace(agent_no, "(.*)", "_\\1")) %>% 
      # 如果agent_no是NA则替换成“_”
      mutate(agent_no = replace_na(agent_no, '_')) %>% 
      # 合并agent_no与channel_cls3
      mutate(channel_cls3 = str_c(channel_cls3, agent_no, sep = "")) %>% 
      # 去掉channel_cls3最后的“_”
      mutate(channel_cls3 = str_replace(channel_cls3, "(.*)_$", "\\1")) %>% 
      select(-agent_no)
    # 添加保险公司
    self_orders <- self_orders %>% left_join(company_self)
    # 添加退单原因
    if(opt$`refund-status`) { self_orders <- self_orders %>% left_join(refund) }
    # 根据skucode标记线上平台的订单
    self_orders[self_orders$sku_code %in% online_sku_code, 'company'] <- '线上平台'
    # 去掉不知道哪里产生的重复订单
    self_orders <- self_orders %>% distinct()
    # 检查订单参保人身份证是否有重复订单
    if(length(which(duplicated(self_orders$id_no))) != 0) {
      warning(paste0(city, "个人订单的参保人身份证有重复"))
    }
    
    # 4.3 企业订单 ----
    # 企业订单所需表格
    # 判断有无输入unpaid_com，如果没有则设为FALSE
    unpaid_com <- tryCatch(
      # expr
      { unpaid_com },
      # error condition
      error = function(e) {
        print("Cannot find unpaid_com, set unpaid_com to FALSE.")
        FALSE
      }
    )
    if(unpaid_com) {
      # 00初始化、01支付中、02已支付
      payment_status <- "(payment_status = '00' OR payment_status = '01' OR payment_status = '02')"
    } else {
      # 02已支付、03待退款、04退款中、05已退款、06部分退款
      payment_status <- "payment_status = '02'"
    }
    
    # 基础表格 (com_order_item) ----
    com_orders <- dbGetQuery(
        product,
        glue::glue(
          "
SELECT 
    id order_item_id, 
    CAST( date_created AS CHAR ) date_created, # 必须先转成字符串，否则取到R中time-zone会混乱
    CAST( date_updated AS CHAR ) date_updated, # 必须先转成字符串，否则取到R中time-zone会混乱
    product_id sku_code, 
    SUBSTR( sale_channel_id, 1, 16 ) sale_channel_id, # 渠道码只有16位
    policy_status policy_order_status,
    payment_status,
    CAST( amount AS CHAR ) payment_amount,
    com_id,
    order_id,
    'com' AS `group`
FROM 
    com_order_item 
WHERE 
    order_status != '02' # 不是已取消
    AND {payment_status}
    AND amount IN ({price})
    AND product_id IN ({sku_code_com})
    AND date_created >= '{start_time}'
    AND date_created < '{end_time}'
    #AND policy_status != '05' # 非退单
          "
        )
      ) %>% 
      # 去掉time-zone
      mutate(
        date_created = str_replace(date_created, '(.*)\\sCST', '\\1'),
        date_updated = str_replace(date_updated, '(.*)\\sCST', '\\1')
      ) %>% 
      # 转换日期格式
      mutate(
        date_created = as.POSIXct(date_created), 
        date_updated = as.POSIXct(date_updated)
      )
    # 添加被保险人信息 (com_order_item_info) ----
    com_orders <- com_orders %>% 
      left_join(
        dbGetQuery(
          product,
          glue(
            "
SELECT 
    order_item_id, 
    name,
    CASE
        WHEN id_type = 1 THEN '身份证'
        WHEN id_type = 2 THEN '护照'
        WHEN id_type = 3 THEN '港澳居民来往内地通行证'
        WHEN id_type = 4 THEN '军官证'
        WHEN id_type = 5 THEN '台湾居民来往大陆通行证'
        WHEN id_type = 6 THEN '出生证'
        WHEN id_type = 7 THEN '外国人永久居留证' 
    END id_type,
    UPPER( id_no ) id_no, 
    CASE 
        WHEN gender = 'M' THEN '男' ELSE '女'
    END sex, 
    birthday birthday,
    tel
FROM 
    com_order_item_info
WHERE 
    date_created >= '{start_time}' 
    AND date_created < '{end_time}'
            "
          )
        )
      )
    # 添加支付时间 (com_order_payment) ----
    com_orders <- com_orders %>% 
      left_join(
        dbGetQuery(
          product,
          glue(
            "
SELECT 
    order_id, 
    CAST( date_updated AS CHAR ) date_updated_pay,
    CASE
        WHEN pay_way = '00' THEN '微信支付'
        WHEN pay_way = '05' THEN '银联支付'
        WHEN pay_way IN ('09', '10') THEN '对公转账'
        WHEN pay_way IN ('11', '12') THEN '医保个账支付'
        ELSE '其他支付'
    END pay_way
FROM 
    com_order_payment
WHERE 
    payment_status = '02' # 筛选支付成功的记录，否则同一个订单会有多条记录
    AND date_created >= '{start_time}' 
    AND date_created < '{end_time}'
            "
          )
        ) %>% 
          # 去掉time-zone
          mutate(
            date_updated_pay = str_replace(date_updated_pay, '(.*)\\sCST', '\\1')
          ) %>% 
          # 转换日期格式
          mutate(
            date_updated_pay = as.POSIXct(date_updated_pay)
          )
      )
    # 添加主订单编号 (com_order) ----
    com_orders <- com_orders %>% 
      left_join(
        dbGetQuery(
          product, 
          glue(
            "
SELECT 
    order_no main_order_id, 
    id order_id 
FROM 
    com_order
WHERE 
    date_created >= '{start_time}'
    AND date_created <= '{end_time}'
            "
          )
        )
      )
    # 添加公司信息 (com_order_policy) ----
    com_orders <- com_orders %>% 
      left_join(
        dbGetQuery(
          product,
          glue(
            "
SELECT
    order_id,
    insurance_company_name policy_company, # 承保保司
    policy_no, # 承保单号
    CAST( effect_start_date AS CHAR ) effective_date, # 保单起期
    CAST( effect_end_date AS CHAR ) expiry_date # 保单止期
FROM
    com_order_policy
WHERE
    date_created >= '{start_time}'
    AND date_created <= '{end_time}'
            "
          )
        ) %>% 
          # 去掉time-zone
          mutate(
            effective_date = str_replace(effective_date, '(.*)\\sCST', '\\1'),
            expiry_date = str_replace(expiry_date, '(.*)\\sCST', '\\1')
          ) %>% 
          # 转换日期格式
          mutate(
            effective_date = as.POSIXct(effective_date), 
            expiry_date = as.POSIXct(expiry_date)
          )
      )
    # 其他筛选 ----
    # 把没有sale_channel_id的订单挑出来再添加渠道信息，否则会造成订单重复！！！
    com_orders_no_channel <- filter(com_orders, is.na(sale_channel_id))
    com_orders <- filter(com_orders, !order_item_id %in% com_orders_no_channel$order_item_id)
    # 个人订单添加渠道信息
    com_orders <- com_orders %>% left_join(channel_cls123)
    # 对于匹配不上渠道的sale_channel_id的处理
    if(F) {
      # 剩下有sale_channel_id的订单，找不到一级机构的，就用sale_channel_id的前10/12/16位尝试
      for(i in c(10, 12, 16)) {
        com_orders_tmp <- com_orders
        if(nrow(com_orders_tmp[which(is.na(com_orders_tmp$channel_cls1)), ]) != 0) {
          com_no_channel1 <- com_orders_tmp[which(is.na(com_orders_tmp$channel_cls1)), ]
          ## 1.1.1-用sale_channel_id的前i位
          com_head_tmp <- com_orders_tmp %>% 
            # 取出没有一级机构的订单
            filter(order_item_id %in% com_no_channel1$order_item_id) %>% 
            select(-c(channel_cls1, channel_cls2, channel_cls3)) %>% 
            # 取sale_channel_id的前i位
            mutate(sale_channel_id = str_sub(sale_channel_id, end = i)) %>% 
            # 与channel_cls123的sale_channel_id前i位做匹配
            left_join(channel_cls123 %>% mutate(sale_channel_id = str_sub(sale_channel_id, end = i)))
          # 如果left_join前后行数一致，说明join key没有重复。
          # 如果X表中的key在Y中有多个，join完行数会变多。
          if(length(com_no_channel1$order_item_id) != length(com_head_tmp$order_item_id)) {
            warning(paste0("企业订单一级机构ID前", i, "位有重复！"))
          }
          com_head_tmp <- filter(com_head_tmp, !is.na(channel_cls1))
          com_orders <- com_orders %>% 
            filter(!order_item_id %in% com_head_tmp$order_item_id) %>% 
            bind_rows(com_head_tmp) %>% 
            arrange(desc(date_created))
        }
      }
      # 清除SQL数据，防止后续爆内存
      rm(com_orders_tmp, com_head_tmp, com_no_channel1)
      gc()
    }
    # 把没有sale_channel_id的订单加回去
    com_orders <- bind_rows(com_orders, com_orders_no_channel)
    # 添加保险公司
    com_orders <- com_orders %>% left_join(company_com)
    # 根据sku_code标记线上平台的订单
    com_orders[com_orders$sku_code %in% online_sku_code, 'company'] <- '线上平台'
    # 用企业支付日期作为企业订单的日期
    com_orders <- com_orders %>% 
      rename(date_created_origin = date_created) %>% 
      rename(date_created = date_updated_pay) %>% 
      distinct()
    # 检查订单参保人身份证是否有重复订单
    if(length(which(duplicated(com_orders$id_no))) != 0) {
      warning(paste0(city, "企业订单的参保人身份证有重复"))
    }
    
    # 4.4 特殊处理 ----
    if(opt$`special-processing`) {
      
      save_list_special_processing <- NULL
      
      # # 成都
      # if(city == '成都') {
      #   if(year == '2021') { source('成都/special_processing_chengdu.R') }
      # }
      
      # # 惠州
      # if(city == '惠州') {
      #   if(year == '2021') { source('惠州/special_processing_huizhou.R') }
      # }
      
      # # 苏州
      # if(city == '苏州') {
      #   if(year == '2022') { source('苏州/special_processing_suzhou.R') }
      # }
      
      # # 芜湖
      # if(city == '芜湖') {
      #   if(year == '2021') {
      #     # 将皖事通作为单独的保司进行统计
      #     wanshitong <- self_orders[
      #       str_detect(self_orders$channel_cls1, '皖事通'), 
      #       c('order_item_id', 'company', 'channel_cls1')
      #     ] %>% 
      #       na.omit()
      #     self_orders$company <- ifelse(
      #       self_orders$order_item_id %in% wanshitong$order_item_id, 
      #       '皖事通',
      #       self_orders$company
      #     )
      #     rm(wanshitong)
      #     gc()
      #   }
      # }
      
      # 徐州
      if(city == '徐州') {
        if(year == '2022') {
          # 将去年的中国太保拆成太保财险和太保寿险
          self_orders <- self_orders %>% 
            mutate(
              company = case_when(
                company == '中国太保' & str_detect(channel_cls1, '产险') ~ '太保财险',
                company == '中国太保' & str_detect(channel_cls1, '产险', negate = T) ~ '太保寿险',
                TRUE ~ company
              )
            )
          com_orders <- com_orders %>% 
            mutate(
              company = case_when(
                company == '中国太保' & str_detect(channel_cls1, '产险') ~ '太保财险',
                company == '中国太保' & str_detect(channel_cls1, '产险', negate = T) ~ '太保寿险',
                TRUE ~ company
              )
            )
        }
      }
      
      # # 宜宾
      # if(city == '宜宾') {
      #   if(year == '2022') {
      #     ### 宜宾2022年度自动扣费为一次扣所有保司，所以需要对自动扣费订单区分归属
      #     # 筛选自动扣费订单，去掉归属保司
      #     ad_orders <- filter(self_orders, sku_code == '200001946') %>% select(-company)
      #     # 2021年度所有订单
      #     yibin2021_orders <- bind_rows(
      #       readRDS('宜宾/yibin2021_self_orders.RDS'), readRDS('宜宾/yibin2021_com_orders.RDS')
      #     ) %>% 
      #       select(-order_item_id)
      #     # 为自动扣费订单添加2021年度的保司归属
      #     ad_orders <- ad_orders %>% left_join(yibin2021_orders)
      #     self_orders <- self_orders %>% 
      #       filter(!order_item_id %in% ad_orders$order_item_id) %>% 
      #       bind_rows(ad_orders) %>% 
      #       mutate(company = str_replace(company, '线上平台', '思派平台'))
      #     rm(ad_orders, yibin2021_orders)
      #     gc()
      #   }
      # }
      
    }
    
    # 4.5 保存 ----
    # 保存
    save_list <- NULL
    # 如果有重复订单，保存
    # 企业
    if( length( which( duplicated( com_orders$order_item_id ) ) ) != 0 ) {
      com_orders_dup <- com_orders %>% 
        filter(
          order_item_id %in% 
            (com_orders[duplicated(com_orders$order_item_id), ][['order_item_id']])
        )
      if(opt$`save-query`) {
        save_list <- c(save_list, 'com_orders_dup')
      }
      #if(sop_dir == '/data/users/yeguanhua/SOP') { logr::log_print(com_orders_dup) }
    }
    # 个人
    if( length( which( duplicated( self_orders$order_item_id ) ) ) != 0 ) {
      self_orders_dup <- self_orders %>% 
        filter(
          order_item_id %in% 
            (self_orders[duplicated(self_orders$order_item_id), ][['order_item_id']])
        )
      if(opt$`save-query`) {
        save_list <- c(save_list, 'self_orders_dup')
      }
      #if(sop_dir == '/data/users/yeguanhua/SOP') { logr::log_print(self_orders_dup) }
    }
    if(opt$`save-query`) {
      save_list <- c(
        save_list, 'city', 'year', 'start_time', 'end_time', 'sop_dir', 'sop_datetime', 
        'self_orders', 'com_orders', 'channel_cls123', 'district', 'have_relationship', 'ended'
      )
      if(opt$`special-processing`) {
        save_list <- c(save_list, save_list_special_processing)
      }
      print(
        glue(
          "
-------------------------------------------------------
         Saving query data to local disk.
-------------------------------------------------------
          "
        )
      )
      save(
        list = save_list,
        file = glue(
          "{sop_dir}/{city}/{sop_datetime}/{city}个人和企业订单-截止至{sop_datetime}.RData"
        )
      )
    }
    
    # 4.6 关闭数据库 ----
    dbDisconnect(product)
    rm(product)
    # 计算数据库取数所需时间
    end.time <- Sys.time()
    time.taken <- end.time - start.time
    if(sop_dir == '/data/users/yeguanhua/SOP') {
      logr::log_print(paste0('The SQL query time of ', city, ' is: '))
      logr::log_print(time.taken)
    }
  }
  
  # 五、统计数据 ----
  if(T) {
    
    print(
      glue(
        "
-------------------------------------------------------
           Start to count SOP data.
-------------------------------------------------------
        "
      )
    )
    
    # all_orders ----
    if(T) {
      if(nrow(com_orders) != 0) {
        all_orders <- bind_rows(self_orders, com_orders)
      } else {
        all_orders <- self_orders
      }
      # 标记自动扣费产生的订单
      all_orders$is_automatic_deduction <- ifelse(
        is.na(all_orders$is_automatic_deduction), '0', all_orders$is_automatic_deduction
      )
    }
    
    # 年龄性别统计
    source(paste0(sop_dir, '/', 'count_age_sex.R'), encoding = 'UTF-8', echo = TRUE)
    # 每日投保量分组统计
    source(paste0(sop_dir, '/', 'count_daily_orders.R'), encoding = 'UTF-8', echo = TRUE)
    # 引导退单情况
    source(paste0(sop_dir, '/', 'count_lured_refund_orders.R'), encoding = 'UTF-8', echo = TRUE)
    # 个人
    if(nrow(self_orders) != 0) {
      count_age_sex(
        df_orders = self_orders, self_or_com_or_all = 'self', city = city, age_gap = 1, 
        col_group = 'company', start_time = start_time, end_time = end_time
      ) %>% 
        write_xlsx(
          .,
          glue(
            "{sop_dir}/{city}/{sop_datetime}/{city}个人年龄性别统计-截止至{sop_datetime}.xlsx"
          )
        )
      count_daily_orders(
        df_orders = self_orders, self_or_com_or_all = 'self', 
        city = city, start_time = start_time, end_time = end_time
      ) %>% 
        write_xlsx(
          .,
          glue(
            "{sop_dir}/{city}/{sop_datetime}/{city}个人每日投保量统计-截止至{sop_datetime}.xlsx"
          )
        )
      # count_lured_refund_orders(
      #   df_orders = self_orders, self_or_com_or_all = 'self', city = city, 
      #   start_time = start_time, end_time = end_time
      # )
    }
    # 企业（注意日期用的是下单日期还是支付日期）
    if(nrow(com_orders) != 0) {
      count_age_sex(
        df_orders = com_orders, self_or_com_or_all = 'com', city = city, 
        col_group = 'company', col_date_created = 'date_created_origin', 
        age_gap = 1, start_time = start_time, end_time = end_time
      ) %>% 
        write_xlsx(
          .,
          glue(
            "{sop_dir}/{city}/{sop_datetime}/{city}企业年龄性别统计-截止至{sop_datetime}.xlsx"
          )
        )
      count_daily_orders(
        df_orders = com_orders, self_or_com_or_all = 'com', city = city, 
        col_date_created = 'date_created_origin', start_time = start_time, end_time = end_time
      ) %>% 
        write_xlsx(
          .,
          glue(
            "{sop_dir}/{city}/{sop_datetime}/{city}企业每日投保量统计-截止至{sop_datetime}.xlsx"
          )
        )
      # 如果企业订单有退单
      # if(nrow(filter(com_orders, policy_order_status == '05')) != 0) {
      #   count_lured_refund_orders(
      #     df_orders = com_orders, self_or_com_or_all = 'com', city = city, 
      #     start_time = start_time, end_time = end_time
      #   )
      # }
    }
    # 个人和企业
    if(nrow(self_orders) != 0 & nrow(com_orders) != 0) {
      count_age_sex(
        df_orders = all_orders, self_or_com_or_all = 'all', city = city, 
        age_gap = 1, col_group = 'company', start_time = start_time, end_time = end_time
      ) %>% 
        write_xlsx(
          .,
          glue(
            "{sop_dir}/{city}/{sop_datetime}/{city}个人和企业年龄性别统计-",
            "截止至{sop_datetime}.xlsx"
          )
        )
      count_daily_orders(
        df_orders = all_orders, self_or_com_or_all = 'all', 
        city = city, start_time = start_time, end_time = end_time
      ) %>% 
        write_xlsx(
          .,
          glue(
            "{sop_dir}/{city}/{sop_datetime}/{city}个人和企业每日投保量统计-",
            "截止至{sop_datetime}.xlsx"
          )
        )
      # count_lured_refund_orders(
      #   df_orders = all_orders, self_or_com_or_all = 'all', city = city, 
      #   start_time = start_time, end_time = end_time
      # )
    }
    
    # 保司出单情况分析
    source(paste0(sop_dir, '/', 'count_orders_by_agent.R'), encoding = 'UTF-8', echo = TRUE)
    # 导入字体
    #library(extrafont)
    #font_import()
    #fonts()
    # 判断有无df_channel_cls1
    # if(is.null(df_channel_cls1)) {
    #   nrow_df_channel_cls1 <- 0
    # } else if(is.na(df_channel_cls1)) {
    #   nrow_df_channel_cls1 <- 0
    # } else if(df_channel_cls1 == '') {
    #   nrow_df_channel_cls1 <- 0
    # }
    nrow_df_channel_cls1 <- tryCatch(
      # expr
      { nrow(df_channel_cls1) },
      # error condition
      error = function(e) {
        print("Cannot find df_channel_cls1, set nrow_df_channel_cls1 to 0.")
        0
      }
    )
    if(nrow_df_channel_cls1 > 0) {
      # 如果有df_channel_cls1，则可以统计代理人出单
      count_orders_by_agent(
        df_orders = filter(all_orders, !company %in% c('线上平台')),
        city = city, 
        start_time = start_time, 
        end_time = end_time, 
        ended = ended,
        group_dates = F,
        count_agent = T,
        col_is_automatic_deduction = 'is_automatic_deduction',
        df_channel_cls1 = df_channel_cls1, 
        df_channel_cls123 = channel_cls123,
        agent_below_avg = TRUE,
        sop_datetime = sop_datetime
      ) %>% 
        write_xlsx(
          .,
          glue(
            "{sop_dir}/{city}/{sop_datetime}/{city}保司出单情况-截止至{sop_datetime}.xlsx"
          )
        )
    } else if(nrow_df_channel_cls1 == 0) {
      # 如果没有，则只统计常规结果
      count_orders_by_agent(
        df_orders = filter(all_orders, !company %in% c('线上平台')),
        city = city, 
        start_time = start_time, 
        end_time = end_time, 
        ended = ended,
        group_dates = T
        #font = 'PingFang SC Semibold' # for MacOS
      ) %>% 
        write_xlsx(
          .,
          glue(
            "{sop_dir}/{city}/{sop_datetime}/{city}保司出单情况-截止至{sop_datetime}.xlsx"
          )
        )
    } else {
      warning(
        paste0("Data format of df_channel_cls1 is not right, skipping count_orders_by_agent.")
      )
    }
    
    # 公众号推文所需数据
    source(
      paste0(sop_dir, '/', 'count_orders_for_wechat_tweet.R'), encoding = 'UTF-8', echo = TRUE
    )
    count_orders_for_wechat_tweet(
      df_orders = all_orders, district_info = district, city = city, save_all = T, 
      have_relationship = have_relationship, col_medical_insure_flag = 'medical_insure_flag', 
      col_pay_way = 'pay_way', start_time = start_time, end_time = end_time
    ) %>% 
      write_xlsx(
        .,
        glue(
          "{sop_dir}/{city}/{sop_datetime}/{city}公众号推文数据统计-截止至{sop_datetime}.xlsx"
        )
      )
    
  }
  
  # 六、额外需求 ----
  # 统计自动扣费和续保情况或发送保司出单情况给保司等
  if(opt$`extra-requirement`) {
    if(city %in% extra_requirement_city_list) {
      # # 成都
      # if(city == '成都') {
      #   if(year == '2021') { source('成都/extra_requirement_chengdu.R') }
      # }
      # # 临沂
      # if(city == '临沂') {
      #   if(year == '2021') { source('临沂/extra_requirement_linyi.R') }
      # }
      # # 随州
      # if(city == '随州') {
      #   if(year == '2021') { source('随州/extra_requirement_suizhou.R') }
      # }
      # # 苏州
      # if(city == '苏州') {
      #   if(year == '2022') { source('苏州/extra_requirement_suzhou.R') }
      # }
      # 徐州
      if(city == '徐州') {
        if(year == '2022') { source('徐州/extra_requirement_xuzhou.R') }
      }
      # # 韶关
      # if(city == '韶关') {
      #   if(year == '2022') { source('韶关/extra_requirement_shaoguan.R') }
      # }
      # # 广州
      # if(city == '广州') {
      #   if(year == '2022') { source('广州/extra_requirement_guangzhou.R') }
      # }
      # # 宜宾
      # if(city == '宜宾') {
      #   if(year == '2022') { source('宜宾/extra_requirement_yibin.R') }
      # }
      # # 惠州
      # if(city == '惠州') {
      #   if(year == '2022') { source('惠州/extra_requirement_huizhou.R') }
      # }
      # # 吉林
      # if(city == '吉林') {
      #   if(year == '2022') { source('吉林/extra_requirement_jilin.R') }
      # }
    }
  }
  # 在所有数据都生成后，再生成标准数据汇报PDF
  if(opt$`standard-report`) {
    rmarkdown::render(
      'standard_report.Rmd', 
      output_file = glue(
        "{sop_dir}/{city}/{sop_datetime}/{city}项目数据汇报-截止至{sop_datetime}.pdf"
      ),
      params = list(
        work_dir = sop_dir,
        city = city,
        project_name = project_name,
        start_date = start_time,
        end_date = end_time,
        sop_datetime = sop_datetime
      )
    )
  }
  
  # 七、压缩文件和发送邮件 ----
  if(opt$`send-sop-email`) {
    # 压缩文件
    # 表格所在的路径
    #proj_dir <- paste0(sop_dir, '/', city, '/', sop_datetime)
    # 移动到表格的路径
    setwd(proj_dir)
    zip::zip(
      zipfile = glue("{city}项目统计数据-截止至{sop_datetime}.zip"),
      files = list.files()[
        str_detect(
          list.files(), 
          glue(
            "{paste0(sop_datetime, '.xlsx')}|{paste0(sop_datetime, '.pdf')}|",
            "低于累计人均出单量的代理人名单"
          )
        )
      ]
    )
    # # 检查附件列表
    # if(is.null(email_attachments)) {
    #   # 如果没有附件，就不发送邮件
    #   stop(paste0("Can't find any attached SOP files."))
    # } else {
    #   # 如果有附件列表，检查每个项目的压缩文件是否存在
    #   # 只能检查每个项目的压缩文件是否存在，无法检查压缩文件内的表格是否齐全
    #   for(attachment in email_attachments) {
    #     if(file.exists(attachment)) {
    #       next
    #     } else {
    #       # 记录压缩文件缺失的项目
    #       missing_files <- c(missing_files, attachment)
    #     }
    #   }
    #   # 如果有任何项目的压缩文件不存在，就不发送邮件
    #   if(length(missing_files) != 0) {
    #     stop(
    #       glue(
    #         "The below SOP files are missing, please check again: ",
    #         "{paste0(missing_files, collapse = ', ')}."
    #       )
    #     )
    #   }
    # }
    # 如果项目的SOP压缩文件不存在，则跳过该项目，否则发送SOP邮件
    if( !file.exists(glue("{city}项目统计数据-截止至{sop_datetime}.zip")) ) {
      warnings(glue("Cannot find SOP zipfile in {city}, email didnot sent."))
      missing_files <- c(missing_files, glue("{city}项目统计数据-截止至{sop_datetime}.zip"))
      next
    } else if( file.exists(glue("{city}项目统计数据-截止至{sop_datetime}.zip")) ) {
      # 收件人列表
      if(opt$test) {
        receivers <- 'xxx@email.com'
      } else {
        # 防止因没有设置收件人而中断任务
        receivers <- 'xxx@email.com'
      }
      ### emayili
      # 邮件服务器
      smtp <- emayili::server(
        host = "smtp.exmail.qq.com",
        port = 465,
        username = "xxx@email.com",
        password = "your_passwd"
      )
      # 邮件内容
      email <- emayili::envelope() %>% 
        from("xxx@email.com") %>% 
        to(receivers) %>% 
        subject(enc2utf8(glue("{city}项目统计数据汇总-截止至{sop_datetime}"))) %>% 
        text(
          glue(
            "
            Dear all，
    
            附件中是截止至{sop_datetime}{city}项目的各项统计数据，请查收。
            仅供内部使用，谢谢！
    
            Best regards,
            叶冠华
            "
          )
        ) %>% 
        attachment(path = glue("{city}项目统计数据-截止至{sop_datetime}.zip"))
      # 发送邮件
      smtp(email, verbose = FALSE)
    }
    
    # 回到SOP路径
    setwd(sop_dir)
  }
  
  # 八、清理内存 ----
  # 先清理内存，再进行下一个循环，防止任务被kill
  rm(
    list = ls()[
      str_detect(
        ls(), 
        # 需要保留的变量
        paste0(
          c('opt', 'sop_dir', 'city_list', 'year_list', 'start_times', 'end_times', 'refund', 
            'extra_requirement_city_list', 'missing_files', 'channel_cls123', 'company_info'),
          collapse = "|"
        ),
        negate = TRUE
      )
    ]
  )
  gc()
}
