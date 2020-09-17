  
  ####### PerriconeMD dashboard ########
  
  # Read Libraries
  #install.packages("rmarkdown") 
  #install.packages("knitr") 
  #devtools::install_github("tidyverse/readxl")
  library("xlsx")
  library("formattable")
  library("openxlsx")
  library("readxl")
  library("reshape")
  library("tidyr")
  options(scipen=999)
  library("lubridate")
  library("rsconnect")
  library("shiny")
  library("stringi")
  library("highcharter")
  library("shinythemes")
  library("shinydashboard")
  library("dplyr")
  library("DT")
  library("knitr")
  library("purrr")

  
  ####################################
  ######## Start loading data ########
  ####################################

  setwd("C:/Users/ema/Desktop/projects/Perricone Dashboard")
  
  ## 1. load BOM list
  BOM_list <-as.data.frame(read_excel("SmartList BOM Extract.xlsx", sheet=1,col_types='text'))[,c(1,2,9,10,11,13)]
  
  ## 2. S&OP forecast data 
  setwd("C:/Users/ema/Desktop/projects/Forecast_Inventory_Report")
  
  SKU_added <-as.data.frame(unique(read.xlsx("Forecast & Inventory Planning Report (Master).xlsm", sheet="Perricone",startRow=4)[,4]))
  names(SKU_added)='Sku'
  new_forecastSKU_list=as.vector(SKU_added[,1])
  
  ## 6. MTD cosumption channel breakdown --> to help readjust the forecast of this month
  MTD_consumption_shipped <-read.xlsx("shift_MTD_consumption_breakdown.xlsx", sheet = 'Shipped',startRow = 2)[,c(1:5)]

  
  
  
  ####################################
  ######## Manipulate Data ###########
  ####################################
  
  ## data manipulation to get the summarised Sku consumption info (This Month)
  
  #MTD_consumption_total<-rbind(MTD_consumption_shipped,MTD_consumption_processed)
  
  MTD_consumption_total<-MTD_consumption_shipped
  
  #### MTD consumption overall level 
  
  MTD_consumption_group <-group_by(MTD_consumption_total,Sku)
  MTD_consumption_summ <-summarise(MTD_consumption_group,MTD_consumed=sum(`QTY`))
  colnames(MTD_consumption_summ)[1]='Item #'
  
  ### MTD consumption Channel level
  
  MTD_consumption_total_group <- group_by(MTD_consumption_total,Sku, Channel)
  MTD_consumption_total_summ<-summarise(MTD_consumption_total_group,MTD_consumed=sum(QTY))
  colnames(MTD_consumption_total_summ)[1]='Item #'
 
  
  
  #############################################################################################################
  ########## Section 2 --> capture SKU's individual & Kit consumption of this month ###########################
  #############################################################################################################
  
  ###### For Overall -->  SKU consumption ######
  
  
  ## Sku as individual FG consumption
  
  df_FG_consumption <- MTD_consumption_summ[,c('Item #','MTD_consumed')]
  colnames(df_FG_consumption)<-c('Sku','MTD_consumed')
  
  
  df_consumption_group <-group_by(df_FG_consumption, Sku)
  df_consumption_summ <-summarise(df_consumption_group,MTD_consumed=sum(MTD_consumed))
  df_consumption_combine=as.data.frame(df_consumption_summ)
  
  #### get rid of some one-off components being added into S&OP forecast
  df_consumption_combine=df_consumption_combine[which(df_consumption_combine$Sku%in%new_forecastSKU_list),]
  df_consumption_combine=df_consumption_combine[order(df_consumption_combine$Sku),]
  
  #### Remove zero MTD shipment & combine shipped & pending for shipment
  df_consumption_combine=df_consumption_combine[which(df_consumption_combine$MTD_consumed!=0),]
  
  ###### For Channel Breakdown -->  SKU consumption ######
  
  ## Sku as Kit component consumption
  df_channel_comp_consumption <-data.frame()
  
  for (i in new_forecastSKU_list){
    df<-BOM_list[which(BOM_list$CMPTITNM_C==i),c(1,2,5)]
    
    renew_df<-MTD_consumption_total_summ[which(MTD_consumption_total_summ$`Item #`%in%df$Item_Number_FGI),]
    renew1_df<-merge(df,renew_df,by.x = 'Item_Number_FGI', by.y = 'Item #', all.y = TRUE)
    renew1_df[is.na(renew1_df)]<-0
    renew1_df$MTD_consumed<-as.numeric(renew1_df$CMPITQTY_C)*renew1_df$MTD_consumed
    renew1_df <-renew1_df[,-c(1,2,3)]
    
    if(length(rownames(renew1_df))==0){
      renew1_df$Sku=as.list(NULL)
    }
    else{
      renew1_df$Sku =i
      renew1_df$Channel=paste(renew1_df$Channel,"(Kit)",sep = " ")
    }
    df_channel_comp_consumption=rbind(df_channel_comp_consumption,renew1_df)
    
  }
  
  
  df_channel_comp_consumption$concat=paste(df_channel_comp_consumption$Sku,df_channel_comp_consumption$Channel,sep="")
  df_channel_comp_consumption=df_channel_comp_consumption[,c('concat','Sku',"Channel",'MTD_consumed')]
  
  ## Sku as individual FG consumption
  
  MTD_consumption_total_summ$Channel=paste(MTD_consumption_total_summ$Channel,"(FG)",sep = " ")
  MTD_consumption_total_summ$concat=paste(MTD_consumption_total_summ$`Item #`,MTD_consumption_total_summ$Channel,sep="")
  
  df_channel_FG_consumption <- MTD_consumption_total_summ[,c('concat','Item #',"Channel",'MTD_consumed')]
  colnames(df_channel_FG_consumption)<-c('concat','Sku',"Channel",'MTD_consumed')
  
  ## combine these 2
  df_channel_consumption <-rbind(df_channel_FG_consumption,df_channel_comp_consumption)
  df_channel_consumption_group <-group_by(df_channel_consumption, concat)
  df_channel_consumption_summ <-summarise(df_channel_consumption_group,Sku=first(Sku),MTD_consumed=sum(MTD_consumed))
  df_channel_consumption_combine=as.data.frame(df_channel_consumption_summ)
  
  #### get rid of some one-off components being added into S&OP forecast
  df_channel_consumption_combine=df_channel_consumption_combine[which(df_channel_consumption_combine$Sku%in%new_forecastSKU_list),]
  df_channel_consumption_combine=df_channel_consumption_combine[order(df_channel_consumption_combine$Sku),]
  
  #### Remove zero MTD shipment & combine shipped & pending for shipment
  df_channel_consumption_combine=df_channel_consumption_combine[which(df_channel_consumption_combine$MTD_consumed!=0),]
  
  #############################################################################################################
  #############################################################################################################
  
  
  ###############################################################################################################################
  ###### Section 3 --> capture On Hand & On Order Kit units that  need to be reconciled & put back to Component Supply  #########
  ###############################################################################################################################
  
  
  #############################################################################################################
  ################################ Section 4--> exporting file ################################################
  #############################################################################################################
  

  xlsx::write.xlsx(df_consumption_combine,"masterplan_raw.xlsx",sheetName='Consumption', row.names=FALSE)
  xlsx::write.xlsx(df_channel_consumption_combine,"masterplan_raw.xlsx",sheetName='Consumption_channel',append = TRUE, row.names=FALSE)

  #xlsx::write.xlsx(OO_reconcile,"masterplan_raw.xlsx",sheetName='reconcile_OO',append = TRUE, row.names=FALSE)
  ########################################################################################################
  ######################### AFTER GENERATING THE FILE ####################################################
  ###################### NEED TO CHANGE SKU format #######################################################
  ##################### ALT + A + E ---> FINISH ##########################################################
  ########################################################################################################
  
  
  
  
  
  
  
  
  
