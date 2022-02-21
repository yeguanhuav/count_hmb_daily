######################################### 年龄性别统计 #############################################
library(glue)
#library(openxlsx)
library(scales)
library(writexl)
library(lubridate)
library(tidyverse)
options(scipen = 200) # scipen表示在200个数字以内都不使用科学计数法

count_age_sex <- function(
  city,
  df_orders,
  self_or_com_or_all = 'all', # 'self', 'com', 'all'
  start_time = NA,
  end_time = Sys.time(),
  age_gap = 1,
  col_order_item_id = 'order_item_id',
  col_date_created = 'date_created',
  col_birthday = 'birthday',
  col_group = 'company',
  col_policy_order_status = 'policy_order_status',
  drop_refunds = TRUE,
  col_sex = 'sex',
  col_id_no = 'id_no',
  plot_age_sex = FALSE,
  plot_width = 1600,
  plot_height = 900,
  legend_position = 'top',
  font = NA,
  save_result = FALSE
) {
  sop_dir <- getwd()
  
  # Rename columns ----
  df_orders <- df_orders %>% 
    rename(
      col_order_item_id = !!col_order_item_id,
      col_date_created = !!col_date_created,
      col_birthday = !!col_birthday,
      col_group = !!col_group,
      col_policy_order_status = !!col_policy_order_status,
      col_sex = !!col_sex,
      col_id_no = !!col_id_no
    ) %>% 
    mutate(
      col_date_created = as.Date(col_date_created, tz = Sys.timezone())
    ) %>% 
    # 去掉多余的列，减少内存使用
    select(
      col_order_item_id,
      col_date_created,
      col_birthday,
      col_group,
      col_policy_order_status,
      col_sex,
      col_id_no
    )
  
  # 开始日期和结束日期
  if(is.na(start_time)) { start_time <- min(df_orders[['col_date_created']]) }
  start_date <- as.Date(start_time, tz = Sys.timezone())
  end_date <- as.Date(end_time, tz = Sys.timezone())
  # 按input日期筛选订单（注意input日期格式）
  df_orders <- df_orders %>% filter(col_date_created >= start_date & col_date_created <= end_date)
  
  # 去掉退单
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
  if(drop_refunds) {
    no_refund_orders <- df_orders %>% filter(!col_policy_order_status %in% !!refund_status)
  } else {
    no_refund_orders <- df_orders
  }
  
  no_refund_orders <- no_refund_orders %>% 
    # 转换日期格式
    mutate(
      col_date_created = as.Date(col_date_created, tz = Sys.timezone()),
      col_birthday = as.Date(col_birthday, tz = Sys.timezone())
    ) %>% 
    # 计算每日总订单
    group_by(col_date_created) %>% 
    mutate(总投保量 = length(col_order_item_id)) %>% 
    ungroup()
  
  # 年龄分组----
  no_refund_orders <- no_refund_orders %>% 
    # 年龄计算以【参保日期】为准，这样年龄就是固定的，不会变
    mutate(
      ## trunc(x): https://www.dummies.com/programming/r/how-to-round-off-numbers-in-r/
      #col_age = trunc(
      #  # Get date difference in year: https://stackoverflow.com/a/48810322/10341233
      #  lubridate::time_length(
      #    x = difftime(.$col_date_created, .$col_birthday), unit = 'years'
      #  )
      #)
      col_age = as.numeric(col_date_created - col_birthday) # 参保时的年龄
    )
  
  # 生成单独的年龄表再join总表
  df_age <- tibble('col_order_item_id' = 'a', 'age_group' = 'b') %>% .[-1,]
  # 0岁
  # 准备level后面进行排序
  age_level <- c('0')
  tmp_age_orders <- no_refund_orders %>% 
    filter(col_age < 365) %>% 
    mutate(age_group = '0')
  df_age <- bind_rows(df_age, select(tmp_age_orders, col_order_item_id, age_group))
  # 1-100岁按age_gap分成不同年龄组
  age_group <- c(seq(1, 100, by = age_gap))
  for(i in 1:length(age_group)) {
    # seq()不包含最大值，所以要在最后一组手动加上100岁
    if(i == length(age_group)) {
      tmp_age_orders <- no_refund_orders %>% 
        filter(col_age >= age_group[i]*365 & col_age <= 100 * 365) %>% 
        mutate(age_group = paste0(age_group[i], '-100'))
      age_level <- c(age_level, paste0(age_group[i], '-100'))
    } else {
      tmp_age_orders <- no_refund_orders %>% 
        filter(col_age >= age_group[i]*365 & col_age < age_group[i+1] * 365) %>% 
        mutate(age_group = paste0(age_group[i], '-', (age_group[i+1] - 1)))
      age_level <- c(age_level, paste0(age_group[i], '-', (age_group[i+1] - 1)))
    }
    # output of for loop
    df_age <- bind_rows(df_age, select(tmp_age_orders, col_order_item_id, age_group))
  }
  # 100岁以上
  # 准备level后面进行排序
  age_level <- c(age_level, '>100')
  tmp_age_orders <- no_refund_orders %>% 
    filter(col_age > 100*365) %>% 
    mutate(age_group = '>100')
  df_age <- bind_rows(df_age, select(tmp_age_orders, col_order_item_id, age_group))
  
  # 在总表添加年龄分组信息
  no_refund_orders <- no_refund_orders %>% left_join(df_age)
  
  # 总的年龄分布统计 ----
  # 按age_group分组进行统计
  df_age1 <- no_refund_orders %>% 
    group_by(age_group) %>% 
    mutate(
      男性人数 = length(which(col_sex == '男')),
      女性人数 = length(which(col_sex == '女')),
      人数 = length(col_id_no)
    ) %>% 
    ungroup() %>% 
    # 取出age_group, 男, 女, total并去重
    select(年龄 = age_group, 人数, 男性人数, 女性人数) %>% 
    distinct()
  # last row
  last_row <- tibble(
    年龄 = '总数', 
    人数 = sum(df_age1$人数), 
    女性人数 = sum(df_age1$女性人数), 
    男性人数 = sum(df_age1$男性人数) 
  )
  df_age1 <- bind_rows(df_age1, last_row) %>% 
    mutate(
      男性占比 = round(男性人数 / 人数, 2), 
      女性占比 = round(女性人数 / 人数, 2), 
    ) %>% 
    mutate(年龄 = factor(年龄, levels = c(age_level, '总数'))) %>% 
    arrange(年龄) %>% 
    # 计算每个年龄段人数占总人数的比例
    mutate(不同年龄段人数占比 = round(人数 / (sum(人数) / 2), 2)) %>% 
    # 列排序
    select(年龄, 人数, 不同年龄段人数占比, everything())
  
  # 按保司分组的年龄统计 ----
  # 按age_group、col_group分组进行统计
  df_age2 <- no_refund_orders %>% 
    mutate(col_group = replace_na(col_group, 'NA')) %>% 
    group_by(col_group, age_group) %>% 
    mutate(
      男性人数 = length(which(col_sex == '男')),
      女性人数 = length(which(col_sex == '女')),
      人数 = length(col_id_no)
    ) %>% 
    ungroup() %>% 
    # 取出age_group, 男, 女, total并去重
    select(分组 = col_group, 年龄 = age_group, 人数, 男性人数, 女性人数) %>% 
    distinct()
  df_age22 <- tibble(
    分组 = 'a', 年龄 = 'a', 人数 = 0.1, 不同年龄段人数占比 = 1,
    女性人数 = 0.1, 男性人数 = 0.1, 男性占比 = 1, 女性占比 = 1
  ) %>% .[-1,]
  
  # 保司按投保量排序
  company_order <- df_age2 %>% select(分组, 人数) %>% 
    group_by(分组) %>% 
    mutate(total_orders = sum(人数)) %>% 
    ungroup() %>% 
    select(-人数) %>% 
    distinct() %>% 
    arrange(desc(total_orders)) %>% 
    .$分组 %>% 
    na.omit()
  
  for(i in company_order) {
    tmp_df_age2 <-  filter(df_age2, 分组 == !!i)
    # last row
    last_row <- tibble(
      分组 = i,
      年龄 = paste0(i, '人数'), 
      人数 =  sum(tmp_df_age2$人数), 
      女性人数 = sum(tmp_df_age2$女性人数), 
      男性人数 = sum(tmp_df_age2$男性人数)
    )
    tmp_df_age2 <- bind_rows(tmp_df_age2, last_row) %>% 
      mutate(
        男性占比 = round(男性人数 / 人数, 2), 
        女性占比 = round(女性人数 / 人数, 2), 
      ) %>% 
      mutate(年龄 = factor(年龄, levels = c(age_level, paste0(i, '人数')))) %>% 
      arrange(年龄) %>% 
      # 计算每个年龄段人数占总人数的比例
      mutate(不同年龄段人数占比 = round(人数 / (sum(人数) / 2), 2))
    df_age22 <- bind_rows(df_age22, tmp_df_age2)
  }
  
  # 输出Excel ----
  # # Use openxlsx
  # # https://stackoverflow.com/a/57279302/10341233
  # # Formatting percentages to excel in R
  # # https://stackoverflow.com/a/48066298/10341233
  # # Create an Excel workbook object and add a worksheet
  # wb <- createWorkbook()
  # # Create a percent style
  # pct = createStyle(numFmt = "0.0%")
  # # Sheets
  # sht1 = addWorksheet(wb, sheetName = glue("年龄性别统计-{order_type}订单"))
  # sht2 = addWorksheet(wb, sheetName = glue("年龄性别分组统计-{order_type}订单"))
  # # Add data to the worksheet we just created
  # writeData(wb, sheet = sht1, x = df_age1)
  # writeData(wb, sheet = sht2, x = df_age22)
  # # Add the percent style to the desired cells
  # addStyle(
  #   wb = wb, sheet = sht1, style = pct, 
  #   cols = (1:ncol(df_age1))[str_detect(colnames(df_age1), '占比|率')],
  #   rows = 1:(nrow(df_age1) + 1),
  #   gridExpand = T
  # )
  # addStyle(
  #   wb = wb, sheet = sht2, style = pct, 
  #   cols = (1:ncol(df_age22))[str_detect(colnames(df_age22), '占比|率')],
  #   rows = 1:(nrow(df_age22) + 1),
  #   gridExpand = T
  # )
  output_list <- list()
  
  column_to_scale <- colnames(df_age1)[str_detect(colnames(df_age1), '占比|率')]
  for(column in column_to_scale) {
    df_age1[[column]] <- scales::label_percent(
      accuracy = 0.1, big.mark = ""
    )(
      as.numeric(df_age1[[column]])
    ) %>% 
      str_replace('Inf', '')
  }
  output_list[[glue("年龄性别统计-{order_type}订单")]] <- df_age1
  
  column_to_scale <- colnames(df_age22)[str_detect(colnames(df_age22), '占比|率')]
  for(column in column_to_scale) {
    df_age22[[column]] <- scales::label_percent(
      accuracy = 0.1, big.mark = ""
    )(
      as.numeric(df_age22[[column]])
    ) %>% 
      str_replace('Inf', '')
  }
  output_list[[glue("年龄性别分组统计-{order_type}订单")]] <- df_age22
  
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
  #       '年龄性别统计-', order_type, '订单-截止至',
  #       lubridate::year(end_time), '年', 
  #       lubridate::month(end_time), '月', 
  #       lubridate::day(end_time), '日', 
  #       lubridate::hour(end_time), '时',
  #       '.xlsx'
  #     ), 
  #     overwrite = T
  #   )
  # }
  
  # 画图 ----
  if(plot_age_sex) {
    df_age_sex <- df_age1 %>% 
      filter(年龄 != '总数') %>% 
      select(年龄, 男性人数, 女性人数) %>% 
      pivot_longer(cols = c('男性人数', '女性人数'), names_to = 'sex_group', values_to = 'y') %>% 
      rename(x = 年龄, 性别 = sex_group) %>% 
      mutate(性别 = str_replace(性别, '人数', '')) %>% 
      mutate(性别 = factor(性别, levels = c('男性', '女性')))
    p <- ggplot(data = df_age_sex, aes(x = x, y = y, fill = 性别)) + 
      geom_bar(stat = "identity", position = position_dodge()) + 
      theme_classic() +
      theme(
        plot.title = element_text(size = 25),# hjust = 0.5), # title居中
        axis.text.x = element_text(size = 20),
        axis.text.y = element_text(size = 20),
        axis.title = element_text(size = 25),
        legend.title = element_blank(),
        legend.text = element_text(size = 20),
        legend.position = legend_position
      ) + 
      labs(
        title = paste0(
          '参保人年龄性别分布-', order_type, '订单-截止至', 
          lubridate::year(end_time), '年', 
          lubridate::month(end_time), '月', 
          lubridate::day(end_time), '日', 
          lubridate::hour(end_time), '时'
        ), 
        x = '年龄', y = '人数'
      ) +
      scale_y_continuous(
        breaks = seq(
          from = 0, 
          to = max(df_age_sex$y), 
          by = signif(max(df_age_sex$y) / 10, 1)
        )
      ) + 
      scale_fill_manual(values = c('#003366', '#CC6633'))
    # 添加字体
    if(!is.na(font)) {
      p <- p + theme(text = element_text(family = font))
    }
    png(
      filename = paste0(
        sop_dir, '/', city, '/', end_date, '/',
        '年龄性别分布', '-', order_type, '订单', '-', end_date,
        '.png'
      ), 
      width = plot_width, height = plot_height
    )
    grid::grid.draw(p)
    dev.off()
  }
  
  return(output_list)
}
