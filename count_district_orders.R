count_district_orders <- function(
  df_orders,
  col_order_item_id = 'order_item_id',
  col_date_created = 'date_created',
  col_date_updated = 'date_updated',
  col_policy_order_status = 'policy_order_status',
  col_id_no = 'id_no',
  replace_id_no = 2,
  district_list,
  col_district = 'district',
  col_district_code = 'district_code',
  col_group = NA,
  group_order = NULL,
  col_price = NA,
  display_date = NA,
  city,
  start_time = NA,
  end_time = Sys.time(),
  sop_datetime = NA
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
  district_list <- district_list %>% 
    select(col_district = !!col_district, col_district_code = !!col_district_code)
  df_orders <- df_orders %>% 
    rename(
      col_order_item_id = !!col_order_item_id,
      col_date_created = !!col_date_created,
      col_date_updated = !!col_date_updated,
      col_policy_order_status = !!col_policy_order_status,
      col_id_no = !!col_id_no
    ) %>% 
    mutate(
      col_date_created = as_date(col_date_created),
      col_date_updated = as_date(col_date_updated)
    )
  # 添加分组
  if(!is.na(col_group)) {
    df_orders <- df_orders %>% rename(col_group = !!col_group)
    if(!is.null(group_order)) {
      if(length(unique(group_order)) != length(unique(df_orders$col_group))) {
        stop(
          glue(
            "Number of unique(group_order) is not equal to number of unique(col_group), ",
            "please check again."
          )
        )
      }
    }
  }
  # 价格列
  if(!is.na(col_price)) {
    df_orders <- df_orders %>% rename(col_price = !!col_price)
    df_orders$col_price <- as.numeric(df_orders$col_price) / 100
  }
  # 去掉多余的列，减少内存使用
  df_orders <- df_orders %>% select(contains('col_'))
  # 开始日期和结束日期
  if(is.na(start_time)) { start_time <- min(df_orders[['col_date_created']]) }
  start_date <- as_date(start_time)
  end_date <- as_date(end_time)
  # 按input日期筛选订单（注意input日期格式）
  df_orders <- df_orders %>% filter(col_date_created >= start_date & col_date_created <= end_date)
  
  # 添加地区
  df_orders <- df_orders %>% 
    # 提取身份证前2位或前6位，用于统计地区
    mutate(
      col_id_no_2 = str_replace(
        col_id_no, paste0("(", paste0(rep('.', replace_id_no), collapse = ''), ").*"), '\\1'
      )
    ) %>% 
    # 添加地区信息
    left_join(district_list, by = c('col_id_no_2' = 'col_district_code')) %>% 
    # 匹配不是的归到其他
    mutate(col_district = replace_na(col_district, '其他'))
  
  # 找到所有退单
  refund_orders <- df_orders %>% filter(col_policy_order_status %in% c('04', '05'))
  
  # 有效订单和退单分别根据下单日期和退单日期添加date
  df_orders <- df_orders %>% 
    # 添加年度_一年中的第几周（yr_week）的信息
    mutate(
      yr = str_replace(col_date_created, '(....).*', '\\1'), 
      week_date = week(col_date_created)
    ) %>% 
    unite(yr, week_date, col = 'yr_week', sep = '_') %>% 
    # 添加月份信息
    mutate(yr_mon = format(col_date_created, format = '%Y-%m'))
  refund_orders <- refund_orders %>% 
    # 添加年度_一年中的第几周（yr_week）的信息
    mutate(
      yr = str_replace(col_date_updated, '(....).*', '\\1'), 
      week_date = week(col_date_updated)
    ) %>% 
    unite(yr, week_date, col = 'yr_week', sep = '_') %>% 
    # 添加月份信息
    mutate(yr_mon = format(col_date_updated, format = '%Y-%m'))
  
  # output file
  output_list <- list()
  
  # Sheet1 ----
  order_group <- c('col_district', 'col_date_created')
  if(!is.na(col_group)) { order_group <- c(order_group, 'col_group') }
  if(!is.na(col_price)) { order_group <- c(order_group, 'col_price') }
  district_order_daily <- df_orders %>% 
    # 根据display_date筛选订单
    filter(col_date_created >= as_date(display_date)) %>% 
    group_by(across(all_of(order_group))) %>% 
    mutate(orders_daily = length(unique(col_order_item_id))) %>% 
    ungroup() %>% 
    select(!!order_group, orders_daily) %>% 
    distinct()
  refund_order_group <- c('col_district', 'col_date_updated')
  if(!is.na(col_group)) { refund_order_group <- c(refund_order_group, 'col_group') }
  if(!is.na(col_price)) { refund_order_group <- c(refund_order_group, 'col_price') }
  district_refund_order_daily <- refund_orders %>% 
    # 根据display_date筛选订单
    filter(col_date_created >= as_date(display_date)) %>% 
    group_by(across(all_of(refund_order_group))) %>% 
    mutate(refund_orders_daily = length(unique(col_order_item_id))) %>% 
    ungroup() %>% 
    select(!!refund_order_group, refund_orders_daily) %>% 
    rename(col_date_created = col_date_updated) %>% 
    distinct()
  district_order_daily <- 
    full_join(district_order_daily, district_refund_order_daily) %>% 
    mutate(
      orders_daily = replace_na(orders_daily, 0), 
      refund_orders_daily = replace_na(refund_orders_daily, 0)
    ) %>% 
    # 净投保量
    mutate(net_orders = orders_daily - refund_orders_daily) %>% 
    arrange(desc(col_date_created), col_district)
  # 计算保费
  if(!is.na(col_price)) {
    # 按照不同价格计算保费
    district_order_daily$col_price <- 
      district_order_daily$col_price * district_order_daily$net_orders
    # 然后根据省份、日期、产品相加
    district_order_daily <- district_order_daily %>% 
      group_by(col_district, col_date_created, col_group) %>% 
      summarize(
        col_price = sum(col_price), orders_daily = sum(orders_daily), 
        refund_orders_daily = sum(refund_orders_daily), net_orders = sum(net_orders)
      ) %>% 
      ungroup()
  }
  # # 如果display_date不为空，则只展示该日期以后的数据
  # if(!is.na(display_date)) {
  #   district_order_daily <- district_order_daily %>%
  #     filter(col_date_created >= as_date(display_date))
  # }
  # 合计
  order_group2 <- order_group[order_group != 'col_date_created']
  if(!is.na(col_price)) {
    order_group2 <- order_group2[order_group2 != 'col_price']
    district_order_daily_output <- district_order_daily %>% 
      group_by(across(all_of(order_group2))) %>% 
      summarise(total_net_orders = sum(net_orders), total_premiums = sum(col_price)) %>% 
      ungroup()
  } else {
    district_order_daily_output <- district_order_daily %>% 
      group_by(across(all_of(order_group2))) %>% 
      summarise(total_net_orders = sum(net_orders)) %>% 
      ungroup()
  }
  # 如果有分组，需要把累计净投保量分组展示，并统计总净投保量和保费
  if(!is.na(col_group)) {
    district_order_daily_output <- district_order_daily_output %>% 
      pivot_wider(
        names_from = col_group, values_from = c(total_net_orders, total_premiums),
        names_glue = "{col_group}_{.value}"
      )
    colnames(district_order_daily_output) <- str_replace(
      colnames(district_order_daily_output), "_total_net_orders", "累计净投保量"
    )
    colnames(district_order_daily_output) <- str_replace(
      colnames(district_order_daily_output), "_total_premiums", "累计保费"
    )
    district_order_daily_output[is.na(district_order_daily_output)] <- 0
    # 计算累计净投保量和累计保费
    total_net_orders <- district_order_daily_output['col_district']
    total_net_orders$累计净投保量 <- 0
    total_premiums <- district_order_daily_output['col_district']
    total_premiums$累计保费 <- 0
    for(grp in unique(district_order_daily$col_group)) {
      total_net_orders$累计净投保量 <- 
        total_net_orders$累计净投保量 + district_order_daily_output[[glue("{grp}累计净投保量")]]
      total_premiums$累计保费 <- 
        total_premiums$累计保费 + district_order_daily_output[[glue("{grp}累计保费")]]
      district_order_daily_output[glue("{grp}累计保费")] <- NULL
    }
    district_order_daily_output <- total_net_orders %>% 
      left_join(total_premiums) %>% 
      left_join(district_order_daily_output)
  }
  # 转置成宽表格
  # 日期从大到小遍历（从近到远）
  if(!is.na(col_group)) {
    # 提取每组的每日的净投保量，转置成宽表格，合并到output
    for(grp in unique(district_order_daily$col_group)) {
      tmp_district_order_daily <- district_order_daily %>% filter(col_group == !!grp)
      for(d in sort(unique(tmp_district_order_daily$col_date_created), decreasing = TRUE)) {
        tmp2_district_order_daily <- tmp_district_order_daily %>% 
          filter(col_date_created == !!d) %>% 
          # 只取需要的列进行转置，转置后剩两列
          select(col_district, col_date_created, col_group, net_orders) %>% 
          pivot_wider(
            names_from = c(col_date_created, col_group), 
            values_from = net_orders, 
            names_glue = "{col_group}{col_date_created}净投保量"
          )
        district_order_daily_output <- left_join(
          district_order_daily_output, tmp2_district_order_daily
        )
      }
    }
  } else {
    for(d in sort(unique(district_order_daily$col_date_created), decreasing = TRUE)) {
      tmp_district_order_daily <- district_order_daily %>% 
        filter(col_date_created == !!d) %>% 
        # 只取需要的列进行转置，转置后剩两列
        select(col_district, col_date_created, net_orders) %>% 
        pivot_wider(names_from = col_date_created, values_from = net_orders)
      district_order_daily_output <- left_join(
        district_order_daily_output, tmp_district_order_daily
      )
    }
    district_order_daily_output <- district_order_daily_output %>% 
      rename(累计净投保量 = total_net_orders)
  }
  district_order_daily_output <- district_order_daily_output %>% 
    # 重命名
    rename(地区 = col_district) %>% 
    # 替换净投保量的NA为0
    mutate(across(contains('净投保量'), ~replace_na(.x, 0)))
  # 如果有分组信息，需要添加所有分组的累计净投保量
  if(!is.na(col_group)) {
    # 准备列排序
    col_order <- c('地区', '累计净投保量')
    # 如果有输入保费，则计算累计保费
    if(!is.na(col_price)) { col_order <- c(col_order, '累计保费') }
    # 列排序
    for(i in group_order) { col_order <- c(col_order, glue("{i}累计净投保量")) }
    for(i in group_order) {
      tmp_col_order <- colnames(district_order_daily_output) %>% 
        .[(!. %in% col_order)] %>% 
        .[(str_detect(., i))]
      col_order <- c(col_order, tmp_col_order)
    }
    district_order_daily_output <- district_order_daily_output %>% select(!!col_order)
  }
  # 行排序，如果有保费则按照保费，否则按照累计净投保量
  district_order_daily_output <- district_order_daily_output %>% arrange(desc(累计净投保量))
  if(!is.na(col_price)) {
    district_order_daily_output <- district_order_daily_output %>% arrange(desc(累计保费))
  }
  display_month <- str_replace(display_date, '.....(..).*', '\\1')
  output_list[[glue("分版本-{display_month}月")]] <- district_order_daily_output
  
  # Sheet2 ----
  order_group <- c('col_district', 'col_date_created')
  if(!is.na(col_group)) { order_group <- c(order_group, 'col_group') }
  if(!is.na(col_price)) { order_group <- c(order_group, 'col_price') }
  district_order_daily <- df_orders %>% 
    # # 根据display_date筛选订单
    # filter(col_date_created >= as_date(display_date)) %>% 
    group_by(across(all_of(order_group))) %>% 
    mutate(orders_daily = length(unique(col_order_item_id))) %>% 
    ungroup() %>% 
    select(!!order_group, orders_daily) %>% 
    distinct()
  refund_order_group <- c('col_district', 'col_date_updated')
  if(!is.na(col_group)) { refund_order_group <- c(refund_order_group, 'col_group') }
  if(!is.na(col_price)) { refund_order_group <- c(refund_order_group, 'col_price') }
  district_refund_order_daily <- refund_orders %>% 
    # # 根据display_date筛选订单
    # filter(col_date_created >= as_date(display_date)) %>% 
    group_by(across(all_of(refund_order_group))) %>% 
    mutate(refund_orders_daily = length(unique(col_order_item_id))) %>% 
    ungroup() %>% 
    select(!!refund_order_group, refund_orders_daily) %>% 
    rename(col_date_created = col_date_updated) %>% 
    distinct()
  district_order_daily <- 
    full_join(district_order_daily, district_refund_order_daily) %>% 
    mutate(
      orders_daily = replace_na(orders_daily, 0), 
      refund_orders_daily = replace_na(refund_orders_daily, 0)
    ) %>% 
    # 净投保量
    mutate(net_orders = orders_daily - refund_orders_daily) %>% 
    arrange(desc(col_date_created), col_district)
  # 计算保费
  if(!is.na(col_price)) {
    # 按照不同价格计算保费
    district_order_daily$col_price <- 
      district_order_daily$col_price * district_order_daily$net_orders
    # 然后根据省份、日期、产品相加
    district_order_daily <- district_order_daily %>% 
      group_by(col_district, col_date_created, col_group) %>% 
      summarize(
        col_price = sum(col_price), orders_daily = sum(orders_daily), 
        refund_orders_daily = sum(refund_orders_daily), net_orders = sum(net_orders)
      ) %>% 
      ungroup()
  }
  # # 如果display_date不为空，则只展示该日期以后的数据
  # if(!is.na(display_date)) {
  #   district_order_daily <- district_order_daily %>%
  #     filter(col_date_created >= as_date(display_date))
  # }
  # 合计
  order_group2 <- order_group[order_group != 'col_date_created']
  if(!is.na(col_price)) {
    order_group2 <- order_group2[order_group2 != 'col_price']
    district_order_daily_output <- district_order_daily %>% 
      group_by(across(all_of(order_group2))) %>% 
      summarise(total_net_orders = sum(net_orders), total_premiums = sum(col_price)) %>% 
      ungroup()
  } else {
    district_order_daily_output <- district_order_daily %>% 
      group_by(across(all_of(order_group2))) %>% 
      summarise(total_net_orders = sum(net_orders)) %>% 
      ungroup()
  }
  # 如果有分组，需要把累计净投保量分组展示，并统计总净投保量和保费
  if(!is.na(col_group)) {
    district_order_daily_output <- district_order_daily_output %>% 
      pivot_wider(
        names_from = col_group, values_from = c(total_net_orders, total_premiums),
        names_glue = "{col_group}_{.value}"
      )
    colnames(district_order_daily_output) <- str_replace(
      colnames(district_order_daily_output), "_total_net_orders", "累计净投保量"
    )
    colnames(district_order_daily_output) <- str_replace(
      colnames(district_order_daily_output), "_total_premiums", "累计保费"
    )
    district_order_daily_output[is.na(district_order_daily_output)] <- 0
    # 计算累计净投保量和累计保费
    total_net_orders <- district_order_daily_output['col_district']
    total_net_orders$累计净投保量 <- 0
    total_premiums <- district_order_daily_output['col_district']
    total_premiums$累计保费 <- 0
    for(grp in unique(district_order_daily$col_group)) {
      total_net_orders$累计净投保量 <- 
        total_net_orders$累计净投保量 + district_order_daily_output[[glue("{grp}累计净投保量")]]
      total_premiums$累计保费 <- 
        total_premiums$累计保费 + district_order_daily_output[[glue("{grp}累计保费")]]
      district_order_daily_output[glue("{grp}累计保费")] <- NULL
    }
    district_order_daily_output <- total_net_orders %>% 
      left_join(total_premiums) %>% 
      left_join(district_order_daily_output)
  }
  # 转置成宽表格
  # 日期从大到小遍历（从近到远）
  if(!is.na(col_group)) {
    # 提取每组的每日的净投保量，转置成宽表格，合并到output
    for(grp in unique(district_order_daily$col_group)) {
      tmp_district_order_daily <- district_order_daily %>% filter(col_group == !!grp)
      for(d in sort(unique(tmp_district_order_daily$col_date_created), decreasing = TRUE)) {
        tmp2_district_order_daily <- tmp_district_order_daily %>% 
          filter(col_date_created == !!d) %>% 
          # 只取需要的列进行转置，转置后剩两列
          select(col_district, col_date_created, col_group, net_orders) %>% 
          pivot_wider(
            names_from = c(col_date_created, col_group), 
            values_from = net_orders, 
            names_glue = "{col_group}{col_date_created}净投保量"
          )
        district_order_daily_output <- left_join(
          district_order_daily_output, tmp2_district_order_daily
        )
      }
    }
  } else {
    for(d in sort(unique(district_order_daily$col_date_created), decreasing = TRUE)) {
      tmp_district_order_daily <- district_order_daily %>% 
        filter(col_date_created == !!d) %>% 
        # 只取需要的列进行转置，转置后剩两列
        select(col_district, col_date_created, net_orders) %>% 
        pivot_wider(names_from = col_date_created, values_from = net_orders)
      district_order_daily_output <- left_join(
        district_order_daily_output, tmp_district_order_daily
      )
    }
    district_order_daily_output <- district_order_daily_output %>% 
      rename(累计净投保量 = total_net_orders)
  }
  district_order_daily_output <- district_order_daily_output %>% 
    # 重命名
    rename(地区 = col_district) %>% 
    # 替换净投保量的NA为0
    mutate(across(contains('净投保量'), ~replace_na(.x, 0)))
  # 如果有分组信息，需要添加所有分组的累计净投保量
  if(!is.na(col_group)) {
    # 准备列排序
    col_order <- c('地区', '累计净投保量')
    # 如果有输入保费，则计算累计保费
    if(!is.na(col_price)) { col_order <- c(col_order, '累计保费') }
    # 列排序
    for(i in group_order) { col_order <- c(col_order, glue("{i}累计净投保量")) }
    for(i in group_order) {
      tmp_col_order <- colnames(district_order_daily_output) %>% 
        .[(!. %in% col_order)] %>% 
        .[(str_detect(., i))]
      col_order <- c(col_order, tmp_col_order)
    }
    district_order_daily_output <- district_order_daily_output %>% select(!!col_order)
  }
  # 行排序，如果有保费则按照保费，否则按照累计净投保量
  district_order_daily_output <- district_order_daily_output %>% arrange(desc(累计净投保量))
  if(!is.na(col_price)) {
    district_order_daily_output <- district_order_daily_output %>% arrange(desc(累计保费))
  }
  output_list[[glue("分版本-投保期全长")]] <- district_order_daily_output
  
  # 按日期统计 ----
  area_order_daily <- df_orders %>% 
    group_by(col_district, col_date_created) %>% 
    mutate(orders_daily = length(unique(col_order_item_id))) %>% 
    ungroup() %>% 
    select(col_date_created, col_district, orders_daily) %>% 
    distinct()
  area_refund_order_daily <- refund_orders %>% 
    group_by(col_district, col_date_updated) %>% 
    mutate(refund_orders_daily = length(unique(col_order_item_id))) %>% 
    ungroup() %>% 
    select(col_date_updated, col_district, refund_orders_daily) %>% 
    distinct()
  area_order_daily <- full_join(
    area_order_daily, area_refund_order_daily,
    by = c('col_date_created' = 'col_date_updated', "col_district" = 'col_district')
  ) %>% 
    mutate(
      orders_daily = replace_na(orders_daily, 0), 
      refund_orders_daily = replace_na(refund_orders_daily, 0)
    ) %>% 
    # 净投保量
    mutate(net_orders = orders_daily - refund_orders_daily) %>% 
    # 净投保量日占比
    group_by(col_date_created) %>% 
    mutate(net_orders_ratio = round(net_orders / sum(net_orders), 3)) %>% 
    ungroup() %>% 
    arrange(desc(col_date_created), col_district)
  # 准备转置成宽表格
  area_order_daily <- area_order_daily %>% 
    select(地区= col_district, 日期 = col_date_created, 净投保量 = net_orders, 
             净投保量占比 = net_orders_ratio) %>% 
    mutate(日期 = format(日期, format = "%m月%d日"))
  area_order_daily_output <- area_order_daily %>% 
    # 添加合计
    group_by(地区) %>% 
    summarise(合计净投保量 = sum(净投保量)) %>% 
    ungroup() %>% 
    mutate(合计净投保量占比 = round(合计净投保量/sum(合计净投保量), 3))
  # 日期从大到小遍历（从近到远）
  for(d in sort(unique(area_order_daily$日期), decreasing = TRUE)) {
    tmp_area_order_daily <- filter(area_order_daily, 日期 == !!d) %>% 
      mutate(日期 = str_replace_all(日期, '0', '')) %>% 
      pivot_wider(
        names_from = 日期, values_from = c(净投保量, 净投保量占比), names_glue = "{日期}{.value}"
      )
    area_order_daily_output <- left_join(
      area_order_daily_output, tmp_area_order_daily, by = c('地区' = '地区')
    )
  }
  # 替换NA为0
  area_order_daily_output <- area_order_daily_output %>% 
    mutate(across(contains('净投保量'), ~replace_na(.x, 0)))
  # 转换ratio为百分比形式
  column_to_scale <- colnames(area_order_daily_output)[
    str_detect(colnames(area_order_daily_output), '占比')
  ]
  for(column in column_to_scale) {
    area_order_daily_output[[column]] <- scales::label_percent(
      accuracy = 0.1, big.mark = ""
    )(
      as.numeric(area_order_daily_output[[column]])
    ) %>% 
      str_replace('Inf', '')
  }
  output_list[['按日统计']] <- area_order_daily_output %>% arrange(desc(合计净投保量))
  
  # 按周统计 ----
  district_order_weekly <- df_orders %>% 
    group_by(col_district, yr_week) %>% 
    mutate(orders_weekly = length(unique(col_order_item_id))) %>% 
    ungroup() %>% 
    select(yr_week, col_district, orders_weekly) %>% 
    distinct()
  district_refund_order_weekly <- refund_orders %>% 
    group_by(col_district, yr_week) %>% 
    mutate(refund_orders_weekly = length(unique(col_order_item_id))) %>% 
    ungroup() %>% 
    select(yr_week, col_district, refund_orders_weekly) %>% 
    distinct()
  district_order_weekly <- full_join(district_order_weekly, district_refund_order_weekly) %>% 
    # 替换NA为0
    mutate(
      orders_weekly = replace_na(orders_weekly, 0), 
      refund_orders_weekly = replace_na(refund_orders_weekly, 0)
    ) %>% 
    # 净投保量
    mutate(net_orders = orders_weekly - refund_orders_weekly) %>% 
    # 净投保量占比
    group_by(yr_week) %>% 
    mutate(net_orders_ratio = round(net_orders / sum(net_orders), 3)) %>% 
    ungroup() %>% 
    arrange(desc(yr_week), col_district)
  # 对星期排序，即最终表格显示上线后的第几周
  week_order <- select(district_order_weekly, yr_week) %>% distinct() %>% arrange(yr_week) %>% 
    mutate(week_order = 1:n())
  district_order_weekly <- district_order_weekly %>% left_join(week_order)
  # 准备转置成宽表格
  district_order_weekly <- district_order_weekly %>% 
    select(
      地区= col_district, 周 = week_order, 净投保量 = net_orders, 净投保量占比 = net_orders_ratio
    )
  district_order_weekly_output <- district_order_weekly %>% 
    # 添加合计
    group_by(地区) %>% 
    summarise(合计净投保量 = sum(净投保量)) %>% 
    ungroup() %>% 
    mutate(合计净投保量占比 = round(合计净投保量/sum(合计净投保量), 3))
  # 星期从大到小遍历（从近到远）
  for(w in sort(unique(district_order_weekly$周), decreasing = TRUE)) {
    tmp_district_order_weekly <- filter(district_order_weekly, 周 == !!w) %>% 
      mutate(周 = str_replace(周, '(.*)', '第\\1周')) %>% 
      pivot_wider(
        names_from = 周, values_from = c(净投保量, 净投保量占比), names_glue = "{周}{.value}"
      )
    district_order_weekly_output <- left_join(
      district_order_weekly_output, tmp_district_order_weekly, by = c('地区' = '地区')
    )
  }
  # 替换NA为0
  district_order_weekly_output <- district_order_weekly_output %>% 
    mutate(across(contains('净投保量'), ~replace_na(.x, 0)))
  # 转换ratio为百分比形式
  column_to_scale <- colnames(district_order_weekly_output)[
    str_detect(colnames(district_order_weekly_output), '占比')
  ]
  for(column in column_to_scale) {
    district_order_weekly_output[[column]] <- scales::label_percent(
      accuracy = 0.1, big.mark = ""
    )(
      as.numeric(district_order_weekly_output[[column]])
    ) %>% 
      str_replace('Inf', '')
  }
  output_list[['按周统计']] <- district_order_weekly_output %>% arrange(desc(合计净投保量))
  
  # 按月统计 ----
  district_order_monthly <- df_orders %>% 
    group_by(col_district, yr_mon) %>% 
    mutate(orders_monthly = length(unique(col_order_item_id))) %>% 
    ungroup() %>% 
    select(yr_mon, col_district, orders_monthly) %>% 
    distinct()
  district_refund_order_monthly <- refund_orders %>% 
    group_by(col_district, yr_mon) %>% 
    mutate(refund_orders_monthly = length(unique(col_order_item_id))) %>% 
    ungroup() %>% 
    select(yr_mon, col_district, refund_orders_monthly) %>% 
    distinct()
  district_order_monthly <- full_join(district_order_monthly, district_refund_order_monthly) %>% 
    # 替换NA为0
    mutate(
      orders_monthly = replace_na(orders_monthly, 0), 
      refund_orders_monthly = replace_na(refund_orders_monthly, 0)
    ) %>% 
    # 净投保量
    mutate(net_orders = orders_monthly - refund_orders_monthly) %>%
    # 净投保量占比
    group_by(yr_mon) %>% 
    mutate(net_orders_ratio = round(net_orders / sum(net_orders), 3)) %>% 
    ungroup() %>% 
    arrange(yr_mon, col_district)
  # 准备转置成宽表格
  district_order_monthly <- district_order_monthly %>% 
    select(地区= col_district, 月份 = yr_mon, 净投保量 = net_orders, 
             净投保量占比 = net_orders_ratio) %>% 
    mutate(月份 = format(月份, format = "%Y年%m月"))
  district_order_monthly_output <- district_order_monthly %>% 
    # 添加合计
    group_by(地区) %>% 
    summarise(合计净投保量 = sum(净投保量)) %>% 
    ungroup() %>% 
    mutate(合计净投保量占比 = round(合计净投保量/sum(合计净投保量), 3))
  # 月份从大到小遍历（从近到远）
  for(m in sort(unique(district_order_monthly$月份), decreasing = TRUE)) {
    tmp_district_order_monthly <- filter(district_order_monthly, 月份 == !!m) %>% 
      mutate(月份 = str_replace(月份, '(....)-(..)', '\\1年\\2月')) %>% 
      pivot_wider(
        names_from = 月份, values_from = c(净投保量, 净投保量占比), names_glue = "{月份}{.value}"
      )
    district_order_monthly_output <- left_join(
      district_order_monthly_output, tmp_district_order_monthly, by = c('地区' = '地区')
    )
  }
  # 替换NA为0
  district_order_monthly_output <- district_order_monthly_output %>% 
    mutate(across(contains('净投保量'), ~replace_na(.x, 0)))
  # 转换ratio为百分比形式
  column_to_scale <- colnames(district_order_monthly_output)[
    str_detect(colnames(district_order_monthly_output), '占比')
  ]
  for(column in column_to_scale) {
    district_order_monthly_output[[column]] <- scales::label_percent(
      accuracy = 0.1, big.mark = ""
    )(
      as.numeric(district_order_monthly_output[[column]])
    ) %>% 
      str_replace('Inf', '')
  }
  output_list[['按月统计']] <- district_order_monthly_output %>% arrange(desc(合计净投保量))
  
  return(output_list)
}