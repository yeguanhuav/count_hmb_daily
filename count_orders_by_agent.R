###################################### 代理人投保情况分析 #########################################
library(grid)
library(ggsci)
library(glue)
#library(openxlsx)
library(scales)
library(writexl)
library(lubridate)
library(tidyverse)

count_orders_by_agent <- function(
  city,
  df_orders,
  start_time = NA,
  end_time = Sys.time(),
  ended = TRUE,
  group_dates = FALSE,
  company_order = NULL, # 格式如c('company1', 'company2', 'company3')
  col_order_item_id = 'order_item_id',
  col_date_created = 'date_created',
  col_date_updated = 'date_updated',
  col_policy_order_status = 'policy_order_status',
  col_company = 'company',
  col_channel_cls3 = 'channel_cls3',
  col_channel_id = 'sale_channel_id',
  col_is_automatic_deduction = NA,
  count_agent = FALSE,
  agent_below_avg = FALSE,
  fixed_avg_order_number = NA,
  # 由保司提供的一级渠道匹配sku_code的表
  df_channel_cls1 = NULL,
  col_channel_cls1_channel_cls1 = 'channel_cls1',
  col_company_channel_cls1 = 'company',
  # 数据库中一二三级渠道的表
  df_channel_cls123 = NULL,
  col_channel_id_channel_cls123 = 'sale_channel_id',
  col_channel_date_created_cls123 = 'channel_date_created',
  col_channel_cls1_channel_cls123 = 'channel_cls1',
  col_channel_cls2_channel_cls123 = 'channel_cls2',
  col_channel_cls3_channel_cls123 = 'channel_cls3',
  sop_datetime = NA,
  # 作图参数
  save_plot = FALSE,
  font = NA, # For MacOS: font = 'PingFang SC Semibold'
  plot_width = 1600,
  plot_height = 900,
  legend_position = 'right',
  annotate_avg_order = FALSE,
  annotation_text_size = 8,
  save_result = FALSE
) {
  sop_dir <- getwd()
  # 如果没有sop_datetime则根据end_time生成sop_datetime
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
      col_date_updated = !!col_date_updated,
      col_policy_order_status = !!col_policy_order_status,
      col_company = !!col_company,
      col_channel_cls3 = !!col_channel_cls3,
      col_channel_id = !!col_channel_id
    ) %>% 
    mutate(
      col_date_created = as.Date(col_date_created, tz = Sys.timezone()),
      col_date_updated = as.Date(col_date_updated, tz = Sys.timezone())
    )
  # 如果含有自动扣费的订单，要去掉
  if(!is.na(col_is_automatic_deduction)) {
    df_orders <- df_orders %>% 
      rename(col_is_automatic_deduction = !!col_is_automatic_deduction) %>% 
      filter(col_is_automatic_deduction == '0')
  }
  # 去掉多余的列，减少内存使用
  df_orders <- df_orders %>% select(contains('col_'))
  
  # 开始日期和结束日期
  if(is.na(start_time)) { start_time <- min(df_orders[['col_date_created']]) }
  start_date <- as.Date(start_time, tz = Sys.timezone())
  end_date <- as.Date(end_time, tz = Sys.timezone())
  
  # 统计退单量 ----
  #order_type <- '个人和企业'
  refund_status <- c('04', '05')
  # 筛选退保订单
  refund_orders <- df_orders %>% 
    # 日期
    mutate(日期 = col_date_updated) %>% 
    # 只保留退保订单
    filter(col_policy_order_status %in% refund_status) %>% 
    # 计算每日总退保量
    group_by(日期) %>% 
    mutate(总退保量 = length(col_order_item_id)) %>% 
    ungroup() %>% 
    # 计算保司的每日退保量
    group_by(日期, col_company) %>% 
    mutate(退保量 = length(col_order_item_id)) %>% 
    ungroup()
  
  # 去掉线上平台的订单 ----
  #df_orders <- df_orders %>% filter(!str_detect(col_company, '线上平台'))
  
  # 保司排序 ----
  if(is.null(company_order)) {
    # 如果没有提供保司顺序，就按投保量排序
    company_order <- df_orders %>% 
      group_by(col_company) %>% 
      mutate(total_orders = length(col_order_item_id)) %>% 
      ungroup() %>% 
      select(col_company, total_orders) %>% 
      distinct() %>% 
      arrange(desc(total_orders)) %>% 
      .$col_company %>% 
      na.omit()
  } else {
    if(
      # 如果提供了保司排序顺序，检查是否含有所有保司
      length(
        which(
          company_order %in% unique(df_orders$col_company)
        )
      ) != length(company_order) | 
      # 原表格中的保司是否都在保司排序的名单中
      length(
        which(
          unique(df_orders$col_company) %in% company_order
        )
      ) != length(unique(df_orders$col_company))
    ) {
      stop(glue("保司排序名单(company_order)和原表格(df_orders)中的保司名单不同，请检查。"))
    }
  }
  
  # 分公司和不同时间段统计 ----
  df_plot <- data.frame(
    row.names = c("日期", "累计投保量", "累计退保量", 
                  "累计出单代理人人数", "累计人均出单量", "公司")
  ) %>% 
    t() %>% as.data.frame() %>% 
    # 不同格式的数据无法combine，所以根据等下要添加的数据转换column格式
    mutate(
      日期 = as.character(日期), 
      累计投保量 = as.numeric(累计投保量), 
      累计退保量 = as.numeric(累计退保量), 
      #累计净投保量 = as.numeric(累计净投保量), 
      累计出单代理人人数 = as.numeric(累计出单代理人人数),
      累计人均出单量 = as.numeric(累计人均出单量), 
      公司 = as.character(公司)
    )
  # 如果是统计每日的数据，添加每日出单代理人人数
  df_plot_daily <- df_plot[1,] %>% .[-1,] %>% 
    rename(
      每日投保量 = 累计投保量, 
      每日退保量 = 累计退保量, 
      #每日净投保量 = 累计净投保量, 
      每日出单代理人人数 = 累计出单代理人人数,
      每日人均出单量 = 累计人均出单量
    )
  
  if(group_dates) {
    # named vector
    group_by_date <- c(
      '前3天' = start_date + 3, # 第4天
      '前5天' = start_date + 5, # 第6天
      '前10天' = start_date + 10, # 第11天
      '前20天' = start_date + 20, # 第21天
      '前30天' = start_date + 31, # 第31天
      '投保期全长' = end_date
    )
    if(ended) {
      group_by_date <- c(
        group_by_date[-length(group_by_date)], # 去掉“投保期全长”
        '倒数10天' = end_date - 10, # 倒数第11天
        '倒数5天' = end_date - 5, # 倒数第6天
        '投保期全长' = end_date
      )
    }
  } else if(!group_dates) {
    # start_date为起始日期
    group_by_date <- c(start_date)
    # end_date的前一天为终止日期
    for(i in 1:(as.numeric(difftime(end_date, start_date)))) {
      group_by_date <- c(group_by_date, as.character(start_date + i))
    }
    group_by_date <- setNames(object = group_by_date, nm = group_by_date)
  } else {
    stop(paste0("Argument 'group_dates' should be either TRUE or FALSE."))
  }
  
  for(i in company_order) {
    # 退单
    refund_orders_tmp <- filter(refund_orders, col_company == !!i)
    # 投保
    df_orders_tmp <- filter(df_orders, col_company == !!i)
    for(j in 1:length(group_by_date)) {
      # 按日期累计的退单
      refund_orders_by_date_tmp <- refund_orders_tmp %>% 
        # 累计时要加上等于号，才能把当天包含进去
        filter(col_date_updated <= as.Date(group_by_date[j]))
      # 按日期累计的投保
      df_orders_by_date_tmp <- df_orders_tmp %>% 
        # 累计时要加上等于号，才能把当天包含进去
        filter(col_date_created <= as.Date(group_by_date[j])) %>% 
        mutate(
          # 添加时间分组
          date_group = names(group_by_date)[j],
          # 累计投保量
          total_order = length(col_order_item_id),
          # 累计退保量
          total_refund = nrow(refund_orders_by_date_tmp),
          # 累计净投保量
          #order_count = length(col_order_item_id) - nrow(refund_orders_by_date_tmp), 
          # 累计出单代理人人数（根据总投保量计算而非净投保量，含退单）
          agent_count = length(unique(col_channel_id))
        ) %>% 
        mutate(
          # 每天的人均出单数（含退单）
          avg_orders_by_agent = round(total_order / agent_count, digits = 1)
        ) %>% 
        select(
          公司 = col_company, 日期 = date_group, 
          累计投保量 = total_order, 
          累计退保量 = total_refund, 
          #累计净投保量 = order_count, 
          累计出单代理人人数 = agent_count, 
          累计人均出单量 = avg_orders_by_agent, 
        ) %>% 
        distinct()
      df_plot <- bind_rows(df_plot, df_orders_by_date_tmp) %>% arrange(日期)
      # 如果是统计每日的数据，添加每日出单代理人人数
      if(!group_dates) {
        # 退单
        refund_orders_by_date_tmp <- refund_orders_tmp %>% 
          filter(col_date_updated == as.Date(group_by_date[j]))
        # 投保
        df_orders_by_date_tmp <- df_orders_tmp %>% 
          filter(col_date_created == as.character(group_by_date[j])) %>% 
          mutate(
            # 添加时间分组
            date_group = names(group_by_date)[j],
            # 每日投保量
            total_order = length(col_order_item_id),
            # 每日退保量
            total_refund = nrow(refund_orders_by_date_tmp), 
            # 每天的净投保量
            #order_count = length(col_order_item_id) - nrow(refund_orders_by_date_tmp), 
            # 每天出单的代理人员（根据总投保量计算而非净投保量，含退单）
            agent_count = length(unique(col_channel_id))
          ) %>% 
          mutate(
            # 每天的人均出单数（含退单）
            avg_orders_by_agent = round(total_order / agent_count, digits = 3)
          ) %>% 
          select(
            公司 = col_company, 日期 = date_group, 
            每日投保量 = total_order, 
            每日退保量 = total_refund, 
            #每日净投保量 = order_count, 
            每日出单代理人人数 = agent_count, 
            每日人均出单量 = avg_orders_by_agent, 
          ) %>% 
          distinct()
        df_plot_daily <- bind_rows(df_plot_daily, df_orders_by_date_tmp) %>% arrange(日期)
      }
    }
  }
  # 如果是统计每日的数据，添加每日出单代理人人数
  if(!group_dates) { df_plot <- left_join(df_plot_daily, df_plot) }
  # 保存原表格
  df_save <- df_plot
  
  if(save_plot) {
    # Dual y-axis (bar + line) ----
    transf_fact <- (max(df_plot$累计出单代理人人数)) / (max(df_plot$累计人均出单量))
    df_plot <- df_plot %>% 
      # 保司顺序排序
      mutate(公司 = factor(公司, levels = company_order)) %>% 
      select(日期, 公司, 累计出单代理人人数, 累计人均出单量) %>% 
      # 将右侧Y轴数值转换成和左侧Y轴数值相近的值
      mutate(累计人均出单量 = 累计人均出单量 * transf_fact) %>% 
      # 转换成长表格，将左右Y轴所需的值合到一个column
      # one column specifying variable, one column specifying value
      pivot_longer(
        cols = c(累计出单代理人人数, 累计人均出单量), names_to = 'y_var', values_to = 'y_val'
      ) %>% 
      # 最终作图的表格需要将每行为单个样本的表格变为统计数值之后的表格
      distinct() %>% 
      mutate(Group = str_c(公司, y_var))
    if(group_dates) {
      if(!ended) {
        df_plot <- df_plot %>% 
          mutate(
            日期 = factor(
              日期, 
              levels = c('前3天', '前5天', '前10天', '前20天', '前30天', '投保期全长')
            )
          )
      } else {
        df_plot <- df_plot %>% 
          mutate(
            日期 = factor(
              日期, 
              levels = c('前3天', '前5天', '前10天', '前20天', '前30天', '倒数5天', '倒数10天', 
                         '投保期全长')
            )
          )
      }
    } else if(!group_dates) {
      # named vector需要用as.character转换成普通的vector
      #df_plot <- df_plot %>% mutate(日期 = factor(日期, levels = as.character(group_by_date)))
      df_plot <- df_plot %>% mutate(日期 = as.Date(日期, tz = Sys.timezone()))
    }
    p <- ggplot(data = df_plot, aes(x = 日期, y = y_val)) +
      # bar chart on left y-axis
      geom_col(
        data = . %>% filter(y_var == '累计出单代理人人数'), aes(group = y_var, fill = y_var)
      ) +
      # lines and point on right y-axis
      geom_line(
        data = . %>% filter(y_var == '累计人均出单量'), size = 1, aes(group = y_var, color = y_var)
      ) +
      geom_point(
        data = . %>% filter(y_var == '累计人均出单量'), size = 2, aes(group = y_var, color = y_var)
      ) +
      scale_y_continuous(
        # 1.手动修改左侧y-axis数值
        breaks = seq(
          0, 
          # 左侧Y轴的最大值
          max(filter(df_plot, y_var == '累计出单代理人人数') %>% .$y_val), 
          # 用signif()以第一位数字取整
          by = signif(
            # 左侧Y轴的最大值
            max(filter(df_plot, y_var == '累计出单代理人人数') %>% .$y_val) / 5, 
            1
          )
        ),
        # 2.添加右侧y-axis，数值为百分比
        sec.axis = sec_axis(
          breaks = seq(
            0, 
            # 右Y轴的最大值
            max((filter(df_plot, y_var == '累计人均出单量') %>% .$y_val) / transf_fact), 
            # 用signif()以第一位数字取整
            by = signif(
              # 右侧Y轴的最大值
              max((filter(df_plot, y_var == '累计人均出单量') %>% .$y_val)) / transf_fact / 5, 
              1
            )
          ),
          trans = ~ . / transf_fact, # 需要除以转换因子，变回百分比
          name = "累计人均出单量"
        )
      ) +
      theme_bw() + 
      theme(
        axis.text.x = element_text(size = 25, angle = 90),
        axis.text.y = element_text(size = 25),
        axis.title = element_text(size = 30),
        legend.text = element_text(size = 25),
        legend.title = element_blank(),
        legend.position = legend_position,
        plot.title = element_text(size = 30, hjust = 0.5), # center plot title
        # facet的title大小
        strip.text.x = element_text(size = 15)
      ) +
      # color为描边的颜色，用于line和plot
      #scale_color_manual(values = c(rep("black", length(unique(df_plot$公司))))) +
      scale_color_manual(values = c('#003366')) +
      # fill为填充的颜色，用于bar
      #scale_fill_manual(
      #  values = pal_igv()(length(unique(df_plot$公司))) # pal_igv()(51) or pal_ucscgb()(26)
      #) +
      scale_fill_manual(values = c('#CC6633')) + 
      # 分公司
      facet_wrap(vars(公司), scales = 'fixed') +
      labs(
        title = paste0(
          city, '保司累计出单代理人人数及累计人均出单量-截止至',
          lubridate::year(end_time), '年', 
          lubridate::month(end_time), '月', 
          lubridate::day(end_time), '日', 
          lubridate::hour(end_time), '时'
        ), 
        x = '', y = '累计出单代理人人数'
      )
    # 在图上添加每天的累计人均出单量
    if(annotate_avg_order) {
      # 准备annotation的表格
      df_text <- df_plot %>% 
        filter(y_var == '累计人均出单量') %>% 
        mutate(annotation = y_val / transf_fact)
      # 添加累计人均出单量
      p <- p + 
        geom_text(
          data = df_text, 
          mapping = aes(x = 日期, y = y_val, label = annotation),
          hjust = 0.5, # 水平
          vjust = 1.5, # 垂直
          size = annotation_text_size
        )
    }
    # 如果日期没有分组，则对X轴的日期进行合并，以避免日期太多看不清
    if(!group_dates) {
      p <- p +
        # 对X轴日期进行调整
        scale_x_date(
          date_breaks = "5 day", date_labels =  "%m月%d日", 
          limits = c(min(df_orders$col_date_created), max(df_orders$col_date_created))
        )
    }
    # 添加字体
    if(!is.na(font)) {
      p <- p + theme(text = element_text(family = font))
    }
    # 保存图片
    if(dir.exists(glue("{sop_dir}/{city}/{sop_datetime}"))) {
      png(
        filename = glue("{sop_dir}/{city}/{sop_datetime}/{city}保司出单情况统计-{end_date}.png"), 
        width = plot_width, height = plot_height
      )
      grid.draw(p)
      dev.off()
    } else {
      png(
        filename = glue("{city}保司出单情况统计-{end_date}.png"), 
        width = plot_width, height = plot_height
      )
      grid.draw(p)
      dev.off()
    }
  }
  
  # 未出单代理人统计 ----
  if(count_agent) {
    if( (!is.null(df_channel_cls1)) & (!is.null(df_channel_cls123)) ) {
      # Rename columns in df_channel_cls1
      df_channel_cls1 <- df_channel_cls1 %>% 
        rename(
          col_channel_cls1_channel_cls1 = !!col_channel_cls1_channel_cls1,
          col_company_channel_cls1 = !!col_company_channel_cls1
        ) %>% 
        # 去掉多余的列，减少内存使用
        select(col_channel_cls1_channel_cls1, col_company_channel_cls1)
      
      # Rename columns in df_channel_cls123
      df_channel_cls123 <- df_channel_cls123 %>% 
        rename(
          col_channel_id_channel_cls123 = !!col_channel_id_channel_cls123,
          col_channel_date_created_cls123 = !!col_channel_date_created_cls123,
          col_channel_cls1_channel_cls123 = !!col_channel_cls1_channel_cls123,
          col_channel_cls2_channel_cls123 = !!col_channel_cls2_channel_cls123,
          col_channel_cls3_channel_cls123 = !!col_channel_cls3_channel_cls123
        ) %>% 
        mutate(
          col_channel_date_created_cls123 = as.Date(col_channel_date_created_cls123, 
                                                    tz = Sys.timezone())
        ) %>% 
        # 去掉多余的列，减少内存使用
        select(
          col_channel_id_channel_cls123,
          col_channel_date_created_cls123,
          col_channel_cls1_channel_cls123,
          col_channel_cls2_channel_cls123,
          col_channel_cls3_channel_cls123
        )
      
      # 添加代理人的统计，只统计三级渠道
      all_agent <- inner_join(
        # inner join【数据库中的代理人表格】和【保司提供的一级渠道表格】
        df_channel_cls123, df_channel_cls1,
        by = c('col_channel_cls1_channel_cls123' = 'col_channel_cls1_channel_cls1')
      ) %>%
        # 去掉二级渠道和三级渠道相同的代理
        filter(col_channel_cls2_channel_cls123 != col_channel_cls3_channel_cls123) %>% 
        distinct()
      
      # 按时间分组统计
      if(group_dates) {
        df_save <- df_save %>% 
          mutate(累计注册代理人人数 = 0.1, 低于累计人均出单量的代理人人数 = 0.1)
      } else {
        df_save <- df_save %>% 
          mutate(
            每日注册代理人人数 = 0.1, 累计注册代理人人数 = 0.1, 低于累计人均出单量的代理人人数 = 0.1
          )
      }
      
      for(i in company_order) {
        print(i)
        for(j in 1:length(group_by_date)) {
          # df_save中对应的日期
          tmp_date_name <- names(group_by_date[j])
          #print(tmp_date_name)
          # 累计人均出单量
          tmp_avg_orders_count <- df_save[
            df_save[['公司']] == i & df_save[['日期']] == tmp_date_name, '累计人均出单量'
          ]
          # 如果没有人均出单量，需要设为0，否则会报错
          if(length(tmp_avg_orders_count) == 0) {
            tmp_avg_orders_count <- 0
          } else if(is.na(tmp_avg_orders_count) | is.null(tmp_avg_orders_count)) {
            warning(glue("The average order count of {i} in {j} is missing."))
            tmp_avg_orders_count <- 0
          }
          # 对应日期的累计注册代理人人数
          tmp_all_agent <- all_agent %>% 
            filter(col_company_channel_cls1 == !!i & 
                     # 累计时要加上等于号，才能把当天包含进去
                     col_channel_date_created_cls123 <= as.Date(group_by_date[j]))
          # 不要将人数进行相加，如果同一company有多个sku_code需要分开统计，如sku_code1对应
          # companyA-1，sku_code2对应companyA-2
          df_save[
            df_save[['公司']] == i & df_save[['日期']] == tmp_date_name, '累计注册代理人人数'
          ] <- nrow(tmp_all_agent)
          # 统计大于累计人均出单量的代理人人数
          # 首先找到该日期的总出单量
          tmp_agent_orders <- df_orders %>% 
            filter(col_company == !!i & 
                     # 累计时加上等于号，才能把当天包含进去
                     col_date_created <= as.Date(group_by_date[j]))
          # 再统计每个代理人出单量
          tmp_agent_orders <- tmp_agent_orders %>% 
            # 加上代理人信息
            inner_join(
              tmp_all_agent, by = c('col_channel_id' = 'col_channel_id_channel_cls123')
            ) %>% 
            # 统计每个代理人出多少单
            group_by(col_channel_id) %>% 
            mutate(count_orders_by_agent = length(col_order_item_id)) %>% 
            ungroup()
          # 大于累计人均出单量的代理人
          tmp_agent_orders <- tmp_agent_orders %>% 
            filter(count_orders_by_agent >= !!tmp_avg_orders_count) %>% 
            select(col_channel_id, count_orders_by_agent) %>% 
            distinct()
          # 保存低于累计人均出单量的代理人
          if(agent_below_avg) {
            if(j == length(group_by_date)) {
              # 统计【已出单代理人的】出单量
              tmp_all_agent_orders <- df_orders %>% 
                # 所有日期的总出单量
                filter(col_company == !!i & col_date_created <= end_date) %>% 
                # 加上代理人信息
                left_join(tmp_all_agent, 
                           by = c('col_channel_id' = 'col_channel_id_channel_cls123')) %>% 
                # 统计每个代理人累计出多少单
                group_by(col_channel_id) %>% 
                mutate(count_orders_by_agent = length(col_order_item_id)) %>% 
                ungroup() %>% 
                select(col_channel_id, count_orders_by_agent) %>% 
                distinct()
              # 如果没有输入人均出单量则计算【所有代理人的】人均出单量，否则直接用输入的人均出单量
              if(is.na(fixed_avg_order_number)) {
                avg_orders <- round(
                  sum(tmp_all_agent_orders$count_orders_by_agent) / nrow(tmp_all_agent),
                  1
                )
              } else {
                avg_orders <- fixed_avg_order_number
              }
              ### 【所有代理人中】低于累计人均出单量的代理人
              # 先找到高于人均出单量的代理人
              agents_above_avg <- tmp_all_agent_orders %>% 
                filter(count_orders_by_agent >= avg_orders)
              # 在所有代理人中去掉高于人均出单量的代理人并将剩下的代理人加上出单量
              agents_below_avg <- tmp_all_agent %>% 
                filter(
                  !col_channel_id_channel_cls123 %in% agents_above_avg$col_channel_id
                ) %>% 
                left_join(
                  tmp_all_agent_orders, by = c('col_channel_id_channel_cls123' = 'col_channel_id')
                ) %>% 
                mutate(count_orders_by_agent = replace_na(count_orders_by_agent, 0)) %>% 
                arrange(desc(count_orders_by_agent)) %>% 
                select(
                  保司 = col_company_channel_cls1, 
                  一级渠道 = col_channel_cls1_channel_cls123,
                  二级渠道 = col_channel_cls2_channel_cls123,
                  三级渠道 = col_channel_cls3_channel_cls123,
                  渠道编码 = col_channel_id_channel_cls123,
                  累计出单量 = count_orders_by_agent
                )
              # 保存名单
              # 如果有sop_datetime则保存在sop_datetime下，否则按日期保存
              if(is.na(sop_datetime)) {
                # 创建文件夹
                if(
                  !dir.exists(
                    paste0(sop_dir, '/', city, '/', end_date, '/低于累计人均出单量的代理人名单')
                  )
                ) {
                  dir.create(
                    paste0(sop_dir, '/', city, '/', end_date, '/低于累计人均出单量的代理人名单')
                  )
                }
                write_xlsx(
                  agents_below_avg, 
                  path = paste0(
                    sop_dir, '/', city, '/', end_date, '/', '低于累计人均出单量的代理人名单/',
                    i, '低于累计人均出单量的代理人名单-截止至', 
                    lubridate::year(end_time), '年', 
                    lubridate::month(end_time), '月', 
                    lubridate::day(end_time), '日', 
                    lubridate::hour(end_time), '时.xlsx'
                  )
                )
              } else {
                # 创建文件夹
                if(
                  !dir.exists(
                    glue("{sop_dir}/{city}/{sop_datetime}/低于累计人均出单量的代理人名单")
                  )
                ) {
                  dir.create(
                    glue("{sop_dir}/{city}/{sop_datetime}/低于累计人均出单量的代理人名单")
                  )
                }
                write_xlsx(
                  agents_below_avg, 
                  path = glue(
                    "{sop_dir}/{city}/{sop_datetime}/低于累计人均出单量的代理人名单/",
                    "{i}低于累计人均出单量的代理人名单-截止至{sop_datetime}.xlsx"
                  )
                )
              }
            }
          }
          # 用总注册人数减大于累计人均出单量的代理人人数
          df_save[
            df_save[['公司']] == i & df_save[['日期']] == tmp_date_name, 
            '低于累计人均出单量的代理人人数'
          ] <- nrow(tmp_all_agent) - nrow(tmp_agent_orders)
          # 如果统计每天的数据，添加每日注册代理人人数
          if(!group_dates) {
            # 对应日期的累计注册代理人人数
            tmp_all_agent <- all_agent %>% 
              filter(col_company_channel_cls1 == !!i & 
                       # named vector需要用as.character转换成普通的vector
                       col_channel_date_created_cls123 == as.character(group_by_date[j]))
            # 将人数进行相加，因为一个company可能有多个sku_code
            df_save[
              df_save[['公司']] == i & df_save[['日期']] == tmp_date_name, '每日注册代理人人数'
            ] <- nrow(tmp_all_agent)
          }
        }
      }
      # 计算累计出单率
      df_save <- df_save %>% mutate(累计出单率 = round(累计出单代理人人数 / 累计注册代理人人数, 3))
      # 列排序
      if(group_dates) {
        df_save <- df_save %>% 
          select(
            公司, 日期, 
            累计投保量, 累计退保量, 累计注册代理人人数, 累计出单代理人人数, 
            累计出单率, 累计人均出单量, 低于累计人均出单量的代理人人数
          )
      } else if(!group_dates) {
        df_save <- df_save %>% 
          select(
            公司, 日期, 
            每日投保量, 每日退保量, 每日注册代理人人数, 每日出单代理人人数, 
            累计投保量, 累计退保量, 累计注册代理人人数, 累计出单代理人人数, 
            累计出单率, 累计人均出单量, 低于累计人均出单量的代理人人数
          )
      }
      if(!is.na(fixed_avg_order_number)) {
        colnames(df_save) <- str_replace(
          colnames(df_save), 
          "低于累计人均出单量的代理人人数", 
          glue("低于{fixed_avg_order_number}单的代理人人数")
        )
      }
    } else {
      stop(
        glue(
          "
          Either df_channel_cls1 or df_channel_cls123 is empty, please check again.
          "
        )
      )
    }
  }
  
  # 保存结果 ----
  # 创建文件夹
  if(save_result | save_plot) {
    if( !dir.exists(paste0(sop_dir, '/', city)) ) {
      dir.create(paste0(sop_dir, '/', city))
    }
    if( !dir.exists(paste0(sop_dir, '/', city, '/', as.Date(end_time, tz = Sys.timezone()))) ) {
      dir.create(paste0(sop_dir, '/', city, '/', as.Date(end_time, tz = Sys.timezone())))
    }
  }
  # 保存表格
  if(T) {
    # 给日期排序
    if(group_dates) {
      if(!ended) {
        df_save <- df_save %>% 
          select(公司, everything()) %>% 
          mutate(
            日期 = factor(
              日期, 
              levels = c('前3天', '前5天', '前10天', '前20天', '前30天', '投保期全长')
            )
          )
      } else {
        df_save <- df_save %>% 
          select(公司, everything()) %>% 
          mutate(
            日期 = factor(
              日期, 
              levels = c('前3天', '前5天', '前10天', '前20天', '前30天', '倒数5天', '倒数10天', 
                         '投保期全长')
            )
          )
      }
    } else if(!group_dates) {
      df_save <- df_save %>% 
        select(公司, everything()) %>% 
        mutate(日期 = factor(日期, levels = as.character(group_by_date)))
    }
    # 保司排序
    df_save <- df_save %>% mutate(公司 = factor(公司, levels = company_order))
    
    # # Write to excel using openxlsx
    # wb <- createWorkbook()
    # # Create a percent style
    # pct <- createStyle(numFmt = "0.0%")
    # # Sheet name
    # sht1 <- addWorksheet(wb, sheetName = glue("保司出单情况统计"))
    # writeData(wb, sheet = sht1, x = df_save %>% group_by(公司) %>% arrange(日期, .by_group = T))
    # # Add the percent style to the desired cells
    # addStyle(
    #   wb = wb, sheet = sht1, style = pct, 
    #   cols = (1:ncol(df_save))[str_detect(colnames(df_save), '占比|率')],
    #   rows = 1:(nrow(df_save) + 1),
    #   gridExpand = T
    # )
    # # Save excel
    # saveWorkbook(
    #   wb,
    #   file = paste0(
    #     sop_dir, '/', city, '/', end_date, '/', '保司出单情况统计-',
    #     #order_type, '订单-',
    #     '截止至',
    #     lubridate::year(end_time), '年',
    #     lubridate::month(end_time), '月',
    #     lubridate::day(end_time), '日',
    #     lubridate::hour(end_time), '时',
    #     '.xlsx'
    #   ),
    #   overwrite = T
    # )
    
    # Write to excel using writexl
    # The way to set SheetName in write_xlsx(): https://github.com/ropensci/writexl/issues/18
    # 将含有“占比”或“率”的列转换成百分比
    column_to_scale <- colnames(df_save)[str_detect(colnames(df_save), '占比|率')]
    for(column in column_to_scale) {
      df_save[[column]] <- scales::label_percent(
        accuracy = 0.1, big.mark = ""
      )(
        as.numeric(df_save[[column]])
      ) %>% 
        str_replace('Inf', '')
    }
    # write_xlsx(
    #   x = list(保司出单情况统计 = (df_save %>% group_by(公司) %>% arrange(日期, .by_group = T))),
    #   path = paste0(
    #     sop_dir, '/', city, '/', end_date, '/', 
    #     '保司出单情况统计-截止至',
    #     lubridate::year(end_time), '年', lubridate::month(end_time), '月',
    #     lubridate::day(end_time), '日', lubridate::hour(end_time), '时.xlsx'
    #   )
    # )
  }
  output_list <- list(
    保司出单情况统计 = (df_save %>% group_by(公司) %>% arrange(日期, .by_group = T))
  )
  return(output_list)
}
