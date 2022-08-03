#### 统计自动扣费相关数据 ####
if(T) {
  library(bit64, quietly = T) # automatic_deduction字段是interger64格式，需要手动加载bit64包，否则会乱码
  library(glue)
  library(RMariaDB)
  library(readxl)
  library(writexl)
  library(tidyverse)
  city <- '惠州'
}

# 一、筛选自动扣费名单 ----
# 1. 目前的自动扣费系统只能根据保司和日期筛选订单，其余筛选需要通过白名单操作
# 2. 白名单不需要添加个人订单已购买的，因为发起自动扣费前开发会去掉
# 3. 给每一批自动扣费名单加上第几批的标签

if(city == '苏州') {
  # 是否用一天多次自动扣费任务
  same_day_multiple_jobs <- FALSE
  
  # 是否去掉扣费失败的订单
  remove_ad_failed <- TRUE
  # 发起过的自动扣费任务ID
  ad_job_id <- "'91', '92', '93'"
  
  rm_500 <- F
  # 订单数大于多少单则金额大于500元
  order_num_over500 <- 5
  
  # 是否为医保个账预约划扣的项目
  medical_insurance_flag = F
  
  rm_com_orders <- FALSE
  
  # 身故人员名单
  passed_away <- read_xlsx('苏州/2022自动扣费/身故人员名单.xlsx')
  
  # 2020年紫金、紫鼎两家公司的参保人从个人订单投保，应当作企业订单处理，加到白名单不进行自动扣费
  exclude_insurant <- read_xlsx('苏州/2022自动扣费/苏康保紫金紫鼎名单.xlsx')
  
  # # 2021.9.25第一次自动扣费-分三批
  # load('苏州/2021-09-25/苏州个人和企业订单-截止至2021年2月1日0时.RData')
  # self_orders <- filter(self_orders, policy_order_status != '04')
  # # 任务分组
  # # 第一批：12.31以后的所有保司订单
  # # 第二批：12.1-12.31的所有保司订单
  # # 第三批：所有思派订单
  # ad_orders <- self_orders %>% 
  #   mutate(date_created = as_date(date_created)) %>% 
  #   mutate(
  #     ad_group = case_when(
  #       sku_code != '00000341' & date_created > as_date('2020-12-31') ~ 
  #         '第一批',
  #       sku_code != '00000341' & date_created <= as_date('2020-12-31') ~ 
  #         '第二批',
  #       sku_code == '00000341' ~ 
  #         '第三批'
  #     )
  #   )
  
  # 2021.10.8 ~ 2021.10.9第二、三次自动扣费-思派、保司
  load("~/SOP/苏州/2021-10-07/suzhou2021_orders_20211007.RData")
  load("~/SOP/苏州/2021-10-07/苏州个人和企业订单-截止至2021年10月7日2时.RData")
  # 筛选非续保的客户
  ad_orders <- filter(suzhou2021_self_orders, !id_no %in% self_orders$id_no)
  # 2021.10.8第二次自动扣费-思派
  ad_orders <- ad_orders %>% filter(sku_code == '00000341') %>% mutate(ad_group = '第四批')
  # 2021.10.9第三次自动扣费-保司
  #ad_orders <- ad_orders %>% filter(sku_code != '00000341') %>% mutate(ad_group = '第五批')
}

if(city == '韶关') {
  # 是否找出订单金额大于500的名单
  rm_500 <- F
  # 订单数大于多少单则金额大于500元
  order_num_over500 <- 8
  
  # 去年签约自动扣费但今年在企业投保
  rm_com_orders <- FALSE
  
  # 是否用一天多次自动扣费任务
  same_day_multiple_jobs <- F
  
  # 是否去掉扣费失败的订单
  remove_ad_failed <- FALSE
  
  # 身故人员名单
  
  
  # 是否为医保个账预约划扣的项目
  medical_insurance_flag = FALSE
  
  # 韶关2021自动扣费订单
  load("~/SOP/韶关/2021年2月1日0时/韶关个人和企业订单-截止至2021年2月1日0时.RData")
  
  # 已出白名单
  # 第一轮
  if(F) {
    # 第一批：人保
    ad_orders_origin <- self_orders %>% 
      mutate(date_created = as_date(date_created)) %>%
      mutate(
        ad_group = case_when(
          sku_code == '00000353' ~ '第一批',
          TRUE ~ 'drop'
        )
      ) %>% 
      filter(ad_group != 'drop')
    # 第二批：国寿
    ad_orders_origin <- self_orders %>% 
      mutate(date_created = as_date(date_created)) %>%
      mutate(
        ad_group = case_when(
          sku_code == '00000354' ~ '第二批',
          TRUE ~ 'drop'
        )
      ) %>% 
      filter(ad_group != 'drop')
    # 第三批：平安
    ad_orders_origin <- self_orders %>% 
      mutate(date_created = as_date(date_created)) %>%
      mutate(
        ad_group = case_when(
          sku_code %in% c('00000355','00000356') ~ '第三批',
          TRUE ~ 'drop'
        )
      ) %>% 
      filter(ad_group != 'drop')
    # 第四批：思派
    ad_orders_origin <- self_orders %>% 
      mutate(date_created = as_date(date_created)) %>%
      mutate(
        ad_group = case_when(
          sku_code == '00000329' ~ '第四批',
          TRUE ~ 'drop'
        )
      ) %>% 
      filter(ad_group != 'drop')
  }
}

if(city == '徐州') {
  # 是否去掉扣费金额大于500元的订单
  rm_500 <- FALSE
  # 订单数大于多少单则金额大于500元
  order_num_over500 <- 7
  
  # 是否用一天多次自动扣费任务
  same_day_multiple_jobs <- FALSE
  
  # 是否去掉扣费失败的订单
  remove_ad_failed <- FALSE
  # 发起过的自动扣费任务ID
  ad_job_id <- "'104','105','106','107','108','109','101','111'"
  
  # 是否去掉今年的企业订单
  rm_com_orders <- TRUE
  # 徐州2022企业订单
  load("徐州/2022自动扣费/徐州个人和企业订单-截止至2022年1月26日17时.RData")
  
  # 徐州身故人员及高客诉风险的名单
  exclude_insurant <- read_xlsx('徐州/2022自动扣费/exclude_insurant.xlsx')
  
  # 是否需要匹配用户中心手机号，发送预扣费短信
  find_phone_num <- FALSE
  if(find_phone_num) {
    customer_id_phone <- read_excel(
      "徐州/2022自动扣费/customer_id_phone.xlsx", col_types = c('text', 'text')
    ) %>% 
      rename(customer_id = id)
    ad_orders <- ad_orders %>% 
      filter(!sub_order_no %in% white_list_output$sub_order_no) %>% 
      left_join(customer_id_phone) %>% 
      select(customer_id, phone) %>% 
      distinct()
    write_xlsx(ad_orders, glue("徐州/2022自动扣费/徐州自动扣费名单-手机号.xlsx"))
  }
  
  # 徐州2021自动扣费订单
  xuzhou2021_self_orders_20210501 <- readRDS(
    "徐州/2021年度/徐州2021个人订单-截止日期2021.05.01-提取日期2021.10.28-自动扣费开关关闭.RDS"
  ) %>% 
    filter(policy_order_status != '04')
  
  # 是否为医保个账预约划扣的项目
  medical_insurance_flag = FALSE
  # 医保白名单（即为有医保个账的人，不进行自动扣费）
  medical_insurance_white_list <- read_csv('徐州/2022自动扣费/职工医保白名单.csv', col_types = 'c')
  
  # 第一轮自动扣费（不在医保白名单的客户）（已出白名单、已扣费）
  if(F) {
    # 第一批：人保财险
    ad_orders_origin <- xuzhou2021_self_orders_20210501 %>%
      filter(company == '人保财险') %>%
      mutate(ad_group = '第一批')
    # 第二批：思派
    ad_orders_origin <- xuzhou2021_self_orders_20210501 %>%
      filter(company == '线上平台') %>%
      mutate(ad_group = '第二批')
    # 第三批：东吴人寿
    ad_orders_origin <- xuzhou2021_self_orders_20210501 %>%
      filter(company == '东吴人寿') %>%
      mutate(ad_group = '第三批')
    # 第四批：中国人寿
    ad_orders_origin <- xuzhou2021_self_orders_20210501 %>%
      filter(company == '中国人寿') %>%
      mutate(ad_group = '第四批')
    # 第五批：泰康养老
    ad_orders_origin <- xuzhou2021_self_orders_20210501 %>%
      filter(company == '泰康养老') %>%
      mutate(ad_group = '第五批')
    # 第六批：国寿财险
    ad_orders_origin <- xuzhou2021_self_orders_20210501 %>%
      filter(company == '国寿财险') %>%
      mutate(ad_group = '第六批')
    # 第七批：中国太保
    ad_orders_origin <- xuzhou2021_self_orders_20210501 %>%
      filter(company == '中国太保') %>%
      mutate(ad_group = '第七批')
    # 第八批：平安养老
    ad_orders_origin <- xuzhou2021_self_orders_20210501 %>%
      filter(company == '平安养老') %>%
      mutate(ad_group = '第八批')
  }
  
  # 第二轮自动扣费（去年微信支付，今年在医保白名单的客户）
  if(F) {
    # 第九批：线上平台
    ad_orders_origin <- xuzhou2021_self_orders_20210501 %>%
      filter(company == '线上平台') %>%
      mutate(ad_group = '第九批')
    
    # 第十批：线上平台，第十一批：中国人寿
    ad_orders_origin <- xuzhou2021_self_orders_20210501 %>% 
      filter(company %in% c('人保财险', '中国人寿')) %>% 
      mutate(date_created = as_date(date_created)) %>%
      mutate(
        ad_group = case_when(
          company == '人保财险' ~ '第十批',
          company == '中国人寿' ~ '第十一批'
        )
      )
    
    # 第十二批：国寿财险，第十三批：东吴人寿，第十四批：泰康养老，
    # 第十五批：平安养老，第十六批：中国太保
    ad_orders_origin <- xuzhou2021_self_orders_20210501 %>% 
      filter(company %in% c('国寿财险', '东吴人寿', '泰康养老', '平安养老', '中国太保')) %>% 
      mutate(date_created = as_date(date_created)) %>%
      mutate(
        ad_group = case_when(
          company == '国寿财险' ~ '第十二批',
          company == '东吴人寿' ~ '第十三批',
          company == '泰康养老' ~ '第十四批',
          company == '平安养老' ~ '第十五批',
          company == '中国太保' ~ '第十六批'
        )
      )
  }
  
  # 最后一次自动扣费
  if(T) {
    self_orders <- filter(self_orders, policy_order_status != '04')
    # 最后一批
    ad_orders_origin <- xuzhou2021_self_orders_20210501 %>% 
      # 去掉个人订单已购买
      filter(!id_no %in% self_orders$id_no) %>% 
      mutate(ad_group = '最后一批')
  }
}

if(city == '广州') {
  # 广州2021自动扣费订单（自动扣费已关闭）
  ad_orders_origin <- readRDS("广州/2022自动扣费/guangzhou2021_ad_orders.RDS") %>% 
    mutate(ad_group = '第一批')
  
  # 广州【无】身故人员及高客诉风险的名单
  
  # 是否去掉扣费金额大于500元的订单
  rm_500 <- TRUE
  # 订单数大于(>)多少单则金额大于500元
  order_num_over500 <- 10
  
  # 是否用一天多次自动扣费任务
  same_day_multiple_jobs <- F
  
  # 是否去掉扣费失败的订单
  remove_ad_failed <- F
  
  # 是否去掉今年的企业订单
  rm_com_orders <- TRUE
  # 最新的广州2022订单
  load("广州/2021年12月16日15时/广州个人和企业订单-截止至2021年12月16日15时.RData")
  
  # 是否需要匹配用户中心手机号，发送预扣费短信
  find_phone_num <- F
  
  # 是否为医保个账预约划扣的项目
  medical_insurance_flag = F
}

if(city == '惠州') {
  # 惠州2021自动扣费订单（自动扣费已关闭）
  ad_orders_origin <- readRDS("惠州/2022自动扣费/huizhou2021_ad_orders.RDS")
  # 第一批
  #ad_orders_origin <- filter(ad_orders_origin, sku_code %in% c('00000419', '200001068')) %>% 
  #  mutate(ad_group = '第一批')
  # 第二批
  #ad_orders_origin <- filter(ad_orders_origin, sku_code %in% c('00000415')) %>% 
  #    mutate(ad_group = '第二批')
  # 第三批
  #ad_orders_origin <- filter(ad_orders_origin, sku_code %in% c('00000416')) %>% 
  #  mutate(ad_group = '第三批')
  # 第四批
  ad_orders_origin <- filter(ad_orders_origin, sku_code %in% c('200001068')) %>% 
    mutate(ad_group = '4')
  
  # 惠州【无】身故人员及高客诉风险的名单
  
  # 是否去掉扣费金额大于500元的订单
  rm_500 <- F
  # 订单数大于(>)多少单则金额大于500元
  #order_num_over500 <- 10
  # 因为第4批扣的人没有openID，所以不处理
  
  # 是否用一天多次自动扣费任务
  same_day_multiple_jobs <- F
  
  # 是否去掉扣费失败的订单
  remove_ad_failed <- F
  
  # 是否去掉今年的企业订单
  rm_com_orders <- TRUE
  # 最新的惠州2022订单
  load("惠州/2022年度/惠州个人和企业订单-截止至2022年5月7日15时.RData")
  
  # 是否需要匹配用户中心手机号，发送预扣费短信
  find_phone_num <- F
  
  # 是否为医保个账预约划扣的项目
  medical_insurance_flag = F
}

if(city == '宜宾') {
  # 宜宾所有订单一次性扣完
  # 宜宾2021自动扣费订单
  ad_orders_origin <- readRDS("宜宾/2022年度-自动扣费/yibin2021_ad_orders.RDS") %>% 
    mutate(ad_group = '第一批')
  
  # 宜宾【无】身故人员及高客诉风险的名单
  
  # 是否去掉扣费金额大于500元的订单
  rm_500 <- TRUE
  # 订单数大于(>)多少单则金额大于500元
  order_num_over500 <- 8
  
  # 是否用一天多次自动扣费任务
  same_day_multiple_jobs <- F
  
  # 是否去掉扣费失败的订单
  remove_ad_failed <- F
  
  # 是否去掉今年的企业订单
  rm_com_orders <- TRUE
  # 最新的宜宾2022订单
  load("宜宾/2021年12月26日12时/宜宾个人和企业订单-截止至2021年12月26日12时.RData")
  
  # 是否需要匹配用户中心手机号，发送预扣费短信
  find_phone_num <- F
  
  # 是否为医保个账预约划扣的项目
  medical_insurance_flag = F
}

if(city == '昆明') {
  # 昆明2021自动扣费订单
  # 第一天
  if(F) {
    ad_orders_origin <- readRDS("昆明/2022年度/自动扣费/kunming2021_ad_orders.RDS") %>% 
      mutate(
        ad_group = case_when(
          company == '人保财' ~ '第一批',
          company == '诚泰' ~ '第二批',
          company == '国寿' ~ '第三批',
          TRUE ~ 'NULL'
        )
      ) %>% 
      filter(ad_group != 'NULL')
  }
  # 第二天
  if(F) {
    ad_orders_origin <- readRDS("昆明/2022年度/自动扣费/kunming2021_ad_orders.RDS") %>% 
      mutate(
        ad_group = case_when(
          company == '平安养老' ~ '4',
          company == '大地财险' ~ '5',
          company == '太保财' ~ '6',
          TRUE ~ 'NULL'
        )
      ) %>% 
      filter(ad_group != 'NULL') %>% 
      arrange(ad_group)
  }
  # 第三天
  if(T) {
    ad_orders_origin <- readRDS("昆明/2022年度/自动扣费/kunming2021_ad_orders.RDS") %>% 
      mutate(
        ad_group = case_when(
          company == '思派平台' ~ '7',
          company == '太保寿-新' ~ '8',
          TRUE ~ 'NULL'
        )
      ) %>% 
      filter(ad_group != 'NULL') %>% 
      arrange(ad_group)
  }
  
  # 昆明【无】身故人员及高客诉风险的名单
  
  # 是否去掉扣费金额大于500元的订单
  rm_500 <- TRUE
  # 订单数大于(>)多少单则金额大于500元
  order_num_over500 <- 7
  
  # 是否用一天多次自动扣费任务
  same_day_multiple_jobs <- T
  
  # 是否去掉扣费失败的订单
  remove_ad_failed <- F
  
  # 是否去掉今年的企业订单
  rm_com_orders <- TRUE
  # 最新的昆明2022订单
  load("昆明/2022年3月17日18时/昆明个人和企业订单-截止至2022年3月17日18时.RData")
  
  # 是否需要匹配用户中心手机号，发送预扣费短信
  find_phone_num <- F
  
  # 是否为医保个账预约划扣的项目
  medical_insurance_flag = F
}

# 筛选开通自动扣费的订单 ----
ad_orders <- ad_orders_origin %>% 
  filter(automatic_deduction == 1) %>% 
  select(
    ad_group, order_item_id, pk_main_order, sub_order_no, company, sku_code, sale_channel_id, 
    customer_id, open_id, applicant_name, applicant_id_type, applicant_id_no, applicant_tel, name, 
    id_type, id_no, sex, birthday, tel
  )

# 去掉openID为空的用户
# P.S: 如果是H5端（如手机UC浏览器）扫码下单的订单，在订单库无法记录openID，但签约成功后用户中心会
# 收到微信返回的openID，这部分客户数量不多并且无法获取准确的openID，所以不做处理
# no_open_id <- ad_orders_origin %>% 
#   filter(is.na(open_id) | open_id == '') %>% 
#   mutate(unique_id = str_c(name, id_type, id_no), reason = '无openID')
# ad_orders <- filter(ad_orders_origin, !sub_order_no %in% no_open_id$sub_order_no)
# P.S: 不能去掉，否则其他情况会有遗漏，只有在处理open_id时才去掉

# 二、白名单（筛选顺序按优先级从高到低） ----

# 添加身故人员及其他需要排除的人员到白名单 ----
exclude_orders <- ad_orders[1,] %>% .[-1,]
# 苏州的“身故人员及其他”包括1.身故人员，2.投诉高风险人员，3.紫金、紫鼎公司订单
# 判断有无身故人员名单
nrow_passed_away <- tryCatch(
  # expr
  { nrow(passed_away) },
  # error condition
  error = function(e) {
    print("Couldn't find passed_away table, set nrow_passed_away to 0.")
    0
  }
)
if(nrow_passed_away > 0) {
  passed_away_orders <- filter(ad_orders, id_no %in% passed_away$id_no)
  exclude_orders <- bind_rows(exclude_orders, passed_away_orders) %>% 
    mutate(unique_id = str_c(name, id_type, id_no), reason = '身故人员及其他')
} else {
  passed_away_orders <- ad_orders[1,] %>% .[-1,]
}
# 判断有无其他需要排除的人员名单
nrow_exclude_insurant <- tryCatch(
  # expr
  { nrow(exclude_insurant) },
  # error condition
  error = function(e) {
    print("Couldn't find exclude_insurant table, set nrow_exclude_insurant to 0.")
    0
  }
)
if(nrow_exclude_insurant > 0) {
  exclude_insurant_orders <- filter(ad_orders, id_no %in% exclude_insurant$id_no)
  exclude_orders <- bind_rows(exclude_orders, exclude_insurant_orders) %>% 
    mutate(unique_id = str_c(name, id_type, id_no), reason = '身故人员及其他')
} else {
  exclude_insurant_orders <- ad_orders[1,] %>% .[-1,]
}

# 去年签约自动扣费但今年在企业投保 ----
if(rm_com_orders) {
  in_com_orders <- ad_orders %>% 
    filter(id_no %in% com_orders$id_no) %>% 
    mutate(unique_id = str_c(name, id_type, id_no), reason = '重复下单')
  # 去年签约自动扣费但今年在个人投保
  if(F) {
    in_self_orders <- ad_orders %>% 
      filter(id_no %in% self_orders$id_no) %>% 
      mutate(unique_id = str_c(name, id_type, id_no), reason = '重复下单')
    in_com_orders <- bind_rows(in_com_orders, in_self_orders)
  }
}

# 订单总金额超过500元 ----
# P.S: 自动扣费时开发会根据openID去掉超过500元的并发送扣费失败通知，可以不做处理
if(rm_500 | find_phone_num) {
  # 同一任务同一openID订单总金额超过500元无法扣费，加入白名单
  over500 <- ad_orders %>%
    ## 筛选openID不为空的订单
    #filter(!is.na(open_id), open_id != "") %>%
    # 如果openID为空则用customerID代替
    mutate(open_id = ifelse(open_id=='', customer_id, open_id)) %>% 
    group_by(ad_group, open_id) %>%
    mutate(count_order = length(sub_order_no)) %>%
    ungroup() %>%
    filter(count_order > !!order_num_over500) %>%
    mutate(unique_id = str_c(name, id_type, id_no), reason = '总额超500元') %>%
    select(-count_order)
}

# 去除自动扣费失败的人 ----
if(remove_ad_failed) {
  product <- dbConnect(
    drv = RMariaDB::MariaDB(), 
    user = 'product',
    password = 'hmb2020',
    host = '10.3.130.16',
    port = 3306,
    dbname = 'trddb',
    groups = 'product'
  )
  # 发起过扣费的用户
  ad_failed <- dbGetQuery(
    product, 
    glue(
      "
SELECT 
    old_order_item_id,
    deal_status
FROM 
    temporary_automatic_deduction_order
WHERE 
    auto_deduction_base_config_id IN ({ad_job_id})
    AND deal_status IN ('31','41')
        # 0初始化；5推消息时检测到白名单；6发起扣费时白名单；10数据清洗成功；11数据清洗失败 
        # 20发送消息成功；21发送消息失败；30发起扣费成功；31发起扣费失败；40扣费成功回调处理 
        # 41扣费失败回调处理
      "
    )
  )
  # 将扣费失败的用户加入白名单
  ad_failed_orders <- ad_orders %>% 
    filter(order_item_id %in% ad_failed$old_order_item_id) %>% 
    mutate(unique_id = str_c(name, id_type, id_no))
  # # 去掉扣费失败的用户
  # ad_orders <- ad_orders %>% 
  #   filter(!order_item_id %in% ad_failed_orders$order_item_id)
  dbDisconnect(product)
  rm(product)
}

# 系统已优化，不需要去除这些人
if(F) {
  # 同一任务同一customerID下有多个openID ----
  # P.S: 扣费时开发会根据customerID找到当前绑定的openID，微信根据openID扣费，如果一个customerID下有多
  # 笔订单不管openID是什么都会一起扣费，这个逻辑不对，因为不一定都是同一个人下的单，可能是在别人手机
  # 上登录了自己的customerID然后帮别人下单，应该按签约时的openID去扣款，所以要把这些人找出来。
  dup_openid_in_customerid <- ad_orders %>%
    # 筛选openID不为空的订单
    filter(!is.na(open_id), open_id != "") %>% 
    group_by(ad_group, customer_id) %>% 
    mutate(open_id_num = length(unique(open_id))) %>% 
    ungroup() %>% 
    select(customer_id, open_id_num) %>% 
    distinct() %>% 
    filter(open_id_num >= 2)
  dup_openid_in_customerid_orders <- ad_orders %>% 
    filter(customer_id %in% dup_openid_in_customerid$customer_id) %>% 
    mutate(unique_id = str_c(name, id_type, id_no), reason = '多个openID')
  
  # 同一任务同一openID下有多个customerID ----
  # P.S: 和3.同理，应该按签约时的openID分开扣
  dup_customerid_in_openid <- ad_orders %>%
    # 筛选openID不为空的订单
    filter(!is.na(open_id), open_id != "") %>% 
    group_by(ad_group, open_id) %>% 
    mutate(customer_id_num = length(unique(customer_id))) %>% 
    ungroup() %>% 
    select(open_id, customer_id_num) %>% 
    distinct() %>% 
    filter(customer_id_num >= 2)
  dup_customerid_in_openid_orders <- ad_orders %>% 
    filter(open_id %in% dup_customerid_in_openid$open_id) %>% 
    mutate(unique_id = str_c(name, id_type, id_no), reason = '多个customerID')
  
  # 同一任务同一customerID下有多个投保人（以身份证区别）的订单数 ----
  # P.S: 如果同个customerID下有多笔主订单，自动扣费时会随机取一个投保人，最终落库的订单信息会出错，
  # 可能引起客诉，需要找出来加入白名单
  dup_applicant_in_customerid <- ad_orders %>% 
    group_by(ad_group, customer_id) %>% 
    mutate(applicant_num = length(unique(applicant_id_no))) %>% 
    ungroup() %>% 
    filter(applicant_num >= 2)
  dup_applicant_in_customerid_orders <- ad_orders %>% 
    filter(customer_id %in% dup_applicant_in_customerid$customer_id) %>% 
    mutate(unique_id = str_c(name, id_type, id_no), reason = '多个投保人')
}

# 同一天的不同扣费任务之间的微信账户重复 ----
if(same_day_multiple_jobs) {
  # 不同扣费任务之间的openID重复：
  # 如果同一天发起多次任务，微信无法在同一天对一个微信账号进行多次扣费
  # 所以要根据ad_group找出重复的账号加入白名单
  dup_wechat_in_jobs <- tibble(open_id = 'a', ad_group = 'b') %>% .[-1,]
  for (i in unique(ad_orders$ad_group)) {
    dup_wechat_tmp <- filter(ad_orders, ad_group == !!i) %>% 
      # 如果openID为空则用customerID代替
      mutate(open_id = ifelse(open_id=='', customer_id, open_id)) %>% 
      select(open_id, ad_group) %>% 
      distinct()
    dup_wechat_in_jobs <- bind_rows(dup_wechat_in_jobs, dup_wechat_tmp)
  }
  dup_wechat_in_jobs <- dup_wechat_in_jobs %>% 
    ## 筛选openID不为空的订单
    #filter(!is.na(open_id), open_id != "") %>% 
    group_by(open_id) %>% 
    mutate(group_num = length(unique(ad_group))) %>% 
    ungroup() %>% 
    filter(group_num >= 2) %>% 
    arrange(open_id)
  dup_openid_in_jobs_orders <- ad_orders %>% 
    filter(open_id %in% dup_wechat_in_jobs$open_id) %>% 
    mutate(unique_id = str_c(name, id_type, id_no), reason = '不同任务的openID重复')
  
  # 系统已优化，不需要去除这些人
  # # 不同扣费任务之间的customerID重复：
  # # 如果同一天发起多次任务，微信无法在同一天对一个微信账号进行多次扣费，所以要找出重复的加入白名单
  # dup_wechat_in_jobs <- tibble(customer_id = 'a', ad_group = 'b') %>% .[-1,]
  # for (i in unique(ad_orders$ad_group)) {
  #   dup_wechat_tmp <- filter(ad_orders, ad_group == !!i) %>% 
  #     select(customer_id, ad_group) %>% 
  #     distinct()
  #   dup_wechat_in_jobs <- bind_rows(dup_wechat_in_jobs, dup_wechat_tmp)
  # }
  # dup_wechat_in_jobs <- dup_wechat_in_jobs %>% 
  #   # 筛选openID不为空的订单
  #   filter(!is.na(customer_id), customer_id != "") %>% 
  #   group_by(customer_id) %>% 
  #   mutate(group_num = length(unique(ad_group))) %>% 
  #   ungroup() %>% 
  #   filter(group_num >= 2) %>% 
  #   arrange(customer_id)
  # dup_customerid_in_jobs_orders <- ad_orders %>% 
  #   filter(customer_id %in% dup_wechat_in_jobs$customer_id) %>% 
  #   mutate(unique_id = str_c(name, id_type, id_no), reason = '不同任务的customerID有重复')
}

# 医保个账预约划扣的项目 ----
if(medical_insurance_flag) {
  # 如果投保人在医保白名单内，则该投保人下的所有订单加入白名单
  medical_insurance_orders <- ad_orders %>% 
    filter(applicant_id_no %in% medical_insurance_white_list$id_no) %>% 
    mutate(unique_id = str_c(name, id_type, id_no), reason = '投保人在医保白名单中')
  # 提取因预约个账进医保白名单的名单
  if(F) {
    z <- xuzhou2021_self_orders_20210501 %>%
      filter(sub_order_no %in% medical_insurance_orders$sub_order_no) %>%
      arrange(date_created) %>% 
      select(
        一级架构 = company,
        二级架构 = channel_cls1,
        三级架构 = channel_cls2,
        代理人 = channel_cls3,
        投保人姓名 = applicant_name,
        投保人证件类型 = applicant_id_type,
        投保人证件号 = applicant_id_no,
        投保人联系方式 = applicant_tel,
        被保险人姓名 = name,
        被保险人证件类型 = id_type,
        被保险人证件号 = id_no
      )
    write_xlsx(z, glue('{city}/{city}签约自动扣费的个账白名单-{ad_orders$ad_group[1]}.xlsx'))
  }
}

# 最终白名单 ----
white_list <- bind_rows(
  #dup_openid_in_customerid_orders, 
  #dup_customerid_in_openid_orders,
  #dup_applicant_in_customerid_orders, 
  exclude_orders
)
if(rm_500 | find_phone_num) {
  white_list <- bind_rows(white_list, over500)
}
if(rm_com_orders) {
  white_list <- bind_rows(white_list, in_com_orders)
}
if(same_day_multiple_jobs) {
  white_list <- bind_rows(white_list, dup_openid_in_jobs_orders)
}
if(remove_ad_failed) {
  white_list <- bind_rows(white_list, ad_failed_orders)
}
if(medical_insurance_flag) {
  white_list <- bind_rows(white_list, medical_insurance_orders)
}
white_list <- white_list %>% select(unique_id) %>% distinct()


white_list_output <- ad_orders %>% 
  mutate(unique_id = str_c(name, id_type, id_no)) %>% 
  filter(unique_id %in% white_list$unique_id)
white_list_output[['身故人员及其他']] <- ifelse(
  white_list_output$unique_id %in% exclude_orders$unique_id, 1, 0
)
if(rm_500 | find_phone_num) {
  white_list_output[['同一任务同一openID总额超500元']] <- ifelse(
    white_list_output$unique_id %in% over500$unique_id, 1, 0
  )
}
if(rm_com_orders) {
  white_list_output[['重复下单']] <- ifelse(
    white_list_output$unique_id %in% in_com_orders$unique_id, 1, 0
  )
}
if(remove_ad_failed) {
  white_list_output[['发起过扣费且失败的用户']] <- ifelse(
    white_list_output$unique_id %in% ad_failed_orders$unique_id, 1, 0
  )
}
# 系统已优化，不需要去除这些人
# white_list_output[['同一任务同一customerID有多个openID']] <- ifelse(
#   white_list_output$unique_id %in% dup_openid_in_customerid_orders$unique_id, 1, 0
# )
# white_list_output[['同一任务同一openID有多个customerID']] <- ifelse(
#   white_list_output$unique_id %in% dup_customerid_in_openid_orders$unique_id, 1, 0
# )
# white_list_output[['同一任务同一customerID有多个投保人']] <- ifelse(
#   white_list_output$unique_id %in% dup_applicant_in_customerid_orders$unique_id, 1, 0
# )
if(same_day_multiple_jobs) {
  # 系统已优化，不需要去除这些人
  # white_list_output[['同一天不同任务的customerID有重复']] <- ifelse(
  #   white_list_output$unique_id %in% dup_customerid_in_jobs_orders$unique_id, 1, 0
  # )
  white_list_output[['同一天不同任务的openID有重复']] <- ifelse(
    white_list_output$unique_id %in% dup_openid_in_jobs_orders$unique_id, 1, 0
  )
}
if(medical_insurance_flag) {
  white_list_output[['投保人在医保白名单中']] <- ifelse(
    white_list_output$unique_id %in% medical_insurance_orders$unique_id, 1, 0
  )
}

write_xlsx(
  white_list_output %>% 
    mutate(参保人出生日期 = str_replace_all(birthday, '-', '')) %>% 
    select(
      参保人姓名 = name,
      参保人证件类型 = id_type,
      参保人证件号 = id_no,
      参保人性别 = sex,
      参保人出生日期
    ),
  glue("{city}/{city}自动扣费白名单-",
       "第{paste0(unique(ad_orders$ad_group), collapse='、')}批-{Sys.Date()}.xlsx")
)

write_xlsx(
  white_list_output, 
  glue("{city}/{city}自动扣费白名单raw-",
       "第{paste0(unique(ad_orders$ad_group), collapse='、')}批-{Sys.Date()}.xlsx")
)
