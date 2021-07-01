library(tidyverse)
library(readxl)
ReadExcel <- function(input_excel){
  temp_ds <- read_excel(str_c(kTargetPath, input_excel), sheet=kTargetSheetName, skip=2, col_names=F)
  colnames(temp_ds) <- c('main_sub_col', 'main_col', 'main_category', 'sub_col', 'sub_category', 'v6', 'v7', 'v8', 'v9', 'v10', 'v11', 'v12', 'v13', 'v14', 'v15', 'v16', 'sum_col')
  temp_ds$facility_name <- str_sub(input_excel, 1, 4) %>% as.numeric()
  temp_ds$filename <- input_excel
  return(temp_ds)
}
EditData3 <- function(input_ds){
  # 列：01がん、02小児、03ゲノム医療、00空欄、計
  temp_ds <- input_ds %>% select(c('main_col', 'main_category', 'facility_name', 'filename')) %>% distinct() %>% na.omit()
  ds_data3 <- temp_ds %>% select(c('main_col', 'main_category'))
  sub_categories <-  input_ds %>% select(c('sub_col', 'sub_category')) %>% distinct()
  sub_categories$sub_category <- ifelse(sub_categories$sub_col == '00', '空欄', sub_categories$sub_category)
  sub_categories <- sub_categories %>% na.omit()
  for (i in 1:nrow(sub_categories)){
    temp_sub_col <- sub_categories[i, "sub_col"] %>% as.character()
    temp_sub_cat <- sub_categories[i, "sub_category"] %>% as.character()
    temp_sum <- input_ds %>% filter(sub_col == temp_sub_col) %>% select(!!temp_sub_cat:='sum_col')

    ds_data3 <- ds_data3 %>% bind_cols(temp_sum)
  }
  ds_data3$計 <- ds_data3$がん + ds_data3$小児 + ds_data3$ゲノム医療 + ds_data3$空欄
  ds_data3 <- ds_data3 %>% left_join(temp_ds, by=c('main_col', 'main_category'))
  return(ds_data3)
}
kTargetPath <- '/Users/mariko/Downloads/2018/'
kTargetSheetName='データ２（記入不要）全分野'
filenames <- list.files(kTargetPath) %>% str_extract('^[0-9].*xlsx$') %>% na.omit()
excelfiles <- map(filenames, ReadExcel)
data3_list <- map(excelfiles, EditData3)

