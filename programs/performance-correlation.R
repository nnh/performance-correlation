# performance-correlation
# Mariko Ohtsuka
# 2021/07/02
# ------ libraries ------
library(tidyverse)
library(readxl)
library(here)
# ------ functions ------
ReadExcel <- function(input_excel){
  temp_ds <- read_excel(here('input', input_excel), sheet=kTargetSheetName, skip=2, col_names=F)
  colnames(temp_ds) <- c('main_sub_column', 'main_column', 'main_category', 'sub_column', 'sub_category', 'v6', 'v7', 'v8', 'v9', 'v10', 'v11', 'v12', 'v13', 'v14', 'v15', 'v16', 'sum_column')
  temp_ds$facility_code <- str_sub(input_excel, 1, 4) %>% as.numeric()
  temp_ds$filename <- input_excel
  return(temp_ds)
}
EditData3 <- function(input_ds){
  # 列：01がん、02小児、03ゲノム医療、00空欄、計
  temp_ds <- input_ds %>% select(c('main_column', 'main_category', 'facility_code', 'filename')) %>% distinct() %>% na.omit()
  ds_data3 <- temp_ds %>% select(c('main_column', 'main_category'))
  sub_categories <-  input_ds %>% select(c('sub_column', 'sub_category')) %>% distinct()
  sub_categories$sub_category <- ifelse(sub_categories$sub_column == '00', '空欄', sub_categories$sub_category)
  sub_categories <- sub_categories %>% na.omit()
  for (i in 1:nrow(sub_categories)){
    temp_sub_column <- sub_categories[i, "sub_column"] %>% as.character()
    temp_sub_categories <- sub_categories[i, "sub_category"] %>% as.character()
    temp_sum <- input_ds %>% filter(sub_column == temp_sub_column) %>% select(!!temp_sub_categories:='sum_column')
    ds_data3 <- ds_data3 %>% bind_cols(temp_sum)
  }
  ds_data3$計 <- ds_data3$がん + ds_data3$小児 + ds_data3$ゲノム医療 + ds_data3$空欄
  ds_data3 <- ds_data3 %>% left_join(temp_ds, by=c('main_column', 'main_category'))
  temp_sum <- c('計', NA, sum(ds_data3$がん), sum(ds_data3$小児), sum(ds_data3$ゲノム医療), sum(ds_data3$空欄), sum(ds_data3$計), NA, NA)
  ds_data3 <- ds_data3 %>% rbind(temp_sum)
  return(ds_data3)
}
EditOutputDS <- function(input_ds){
  input_colnames <- c(c('main_column', '計'))
  output_colname <- input_ds %>% select('facility_code') %>% distinct() %>% na.omit() %>% as.character()
  output_colnames <- c(input_colnames[1], output_colname)
  temp_ds <- input_ds %>% select(all_of(input_colnames))
  temp_ds$計 <- temp_ds$計 %>% as.character()
  output_ds <- data.frame(input_colnames[1], output_colname)
  colnames(output_ds) <- input_colnames
  output_ds <- output_ds %>% rbind(temp_ds)
  return(output_ds)
}
# ------ constant ------
kTargetSheetName <- 'データ２（記入不要）全分野'
kOutputFileName <- 'research_performance2018.csv'
# ------ main ------
filenames <- list.files(here('input')) %>% str_extract('^[0-9].*xlsx$') %>% na.omit()
excelfiles <- map(filenames, ReadExcel)
data3_list <- map(excelfiles, EditData3)
output_data3_list <- map(data3_list, EditOutputDS)
output_ds <- output_data3_list[[1]]
for (i in 2:length(output_data3_list)){
  temp_colname <- str_c('V', i)
  temp_ds <- output_data3_list[[i]] %>% rename(!!temp_colname:='計')
  output_ds <- left_join(output_ds, temp_ds, by='main_column')
}
write_excel_csv(output_ds, here('output', kOutputFileName), col_names=F)
