  
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
  
  ### Check upon the columns ########
  ### April 2020 in S&OP forecast file
  channel_start_col=27 
  company_start_col=20
  
  
  
  ####################################
  ######## Start loading data ########
  ####################################
  
  setwd("C:/Users/ema/Desktop/projects/Forecast_Inventory_Report")
  
  ## 1. load BOM list
  BOM_list <-as.data.frame(read_excel("SmartList BOM Extract.xlsx", sheet=1,col_types='text'))[,c(1,2,9,10,11,13)]
  
  ## 2. S&OP forecast data 
  
  channel_forecast <-read.xlsx("S&OP Roll up - (most recent).xlsx", sheet = 'Channel Forecast',startRow = 1)[,c(1:5,channel_start_col:35,37:48)]
  company_forecast <-read.xlsx("S&OP Roll up - (most recent).xlsx", sheet = 'Company Forecast',startRow = 1)[,c(1:3,company_start_col:28,30:41)]
  
  ## 3. get the inventory report
  
  OH_available <-read.xlsx('OH_inventory_report.xlsx',sheet = 'Avai Inventory',startRow = 4)
  colnames(OH_available)=c("Sku","Description","OH available")
  
  OH_reserve <-read.xlsx('OH_inventory_report.xlsx',sheet = 'Reserve Inventory',startRow = 4)
  colnames(OH_reserve)=c("Sku","Description","OH reserve")
  
  OH_total=merge(OH_available,OH_reserve,all.x = TRUE)
  OH_total$`OH reserve`[is.na(OH_total$`OH reserve`)]=0
  
  ## 4. get BIUB information
  
  OH_BIUB <-read.xlsx('OH_inventory_report.xlsx',sheet = 'BIUB',startRow = 2)
  OH_BIUB <-OH_BIUB[,c(1,ncol(OH_BIUB))]
  colnames(OH_BIUB)=c("Sku","BIUB_info")
  
  
  ## 5. get the open PO report
  open_PO <-read.xlsx('Open_PO_report.xlsx',sheet = 'Pivot PO',startRow = 1)
  colnames(open_PO)=c("Sku","Month","PO number",'QTY')
  
  ## 6. MTD cosumption channel breakdown --> to help readjust the forecast of this month
  MTD_consumption_shipped <-read.xlsx("MTD_consumption_breakdown.xlsx", sheet = 'Shipped',startRow = 2)[,c(1:5)]
  MTD_consumption_processed <-read.xlsx("MTD_consumption_breakdown.xlsx", sheet = 'Processed to ship',startRow = 2)[,c(1:5)]
  
  
  
  ####################################
  ######## Manipulate Data ###########
  ####################################
  
  ## data manipulation to get the summarised Sku consumption info (This Month)
  MTD_consumption_total<-rbind(MTD_consumption_shipped,MTD_consumption_processed)
  
  #### MTD consumption overall level 
  MTD_consumption=MTD_consumption_total[,c("Sku",'QTY','rolling_index')]
  MTD_consumption_group <-group_by(MTD_consumption,Sku)
  MTD_consumption_summ <-summarise(MTD_consumption_group,MTD_consumed=sum(`QTY`),time_period=first(`rolling_index`))
  colnames(MTD_consumption_summ)[1]='Item #'
  
  ### MTD consumption Channel level
  
  MTD_consumption_total_group <- group_by(MTD_consumption_total,Sku, Channel)
  MTD_consumption_total_summ<-summarise(MTD_consumption_total_group,MTD_consumed=sum(QTY),time_period=first(`rolling_index`))
  colnames(MTD_consumption_total_summ)[1]='Item #'
  
  
  ### combine to get the whole Sku list need to check forecast/consumption
  WIP_SKU<-unique(BOM_list[which(BOM_list$Item_Class_Code_CI=='INTERMED'),c(3)])
  FG_SKU <-as.vector(unique(company_forecast[,1]))
  forecasted_SKU<-unique(c(WIP_SKU,FG_SKU))
  ## get rid of some one-off components being added into S&OP forecast
  forecasted_SKU <-forecasted_SKU[which(forecasted_SKU%in%BOM_list$Item_Number_FGI)]
  
  new_forecastSKU_list <-unique(c(forecasted_SKU,as.vector(company_forecast$Sku)))
  
  ## massage open PO SKU information
  open_PO$Sku <- gsub(pattern = "MF",replacement = "",open_PO$Sku)
  open_PO$Sku <- gsub(pattern = "KA",replacement = "",open_PO$Sku)
  
  #############################################################################################################
  ### Section 1 --> capture and categorize SKUs' FG demand from each channel & Kits demand from each channel### 
  #############################################################################################################
  
  ### Sku as kit component forecast
  df_kits=data.frame()
  
  for (i in new_forecastSKU_list){
    df<-BOM_list[which(BOM_list$CMPTITNM_C==i),c(1,2,5)]
    new_df <-channel_forecast[which(channel_forecast$Sku%in%df$Item_Number_FGI),]
    new1_df <-merge(df, new_df,by.x = "Item_Number_FGI", by.y = "Sku",all.y=TRUE)
    new1_df[is.na(new1_df)]<-0
    new1_df[,c(8:ncol(new1_df))]<-as.numeric(new1_df$CMPITQTY_C)*new1_df[,c(8:ncol(new1_df))]
    new2_df <-group_by(new1_df,Channel)[,-c(1:3,5:7)]
    new3_df<-summarise_all(new2_df,funs(sum))
    
    if(length(rownames(new3_df))==0){
      new3_df$Sku=as.list(NULL)
    }
    else{
      new3_df$Sku =i
      new3_df$Channel=paste(new3_df$Channel,"(Kit)",sep = " ")
      
    }
    
    df_kits=rbind(df_kits,new3_df)
    
  }
  df_kits=df_kits[,c(ncol(df_kits),2:ncol(df_kits)-1)]
  
  ## Sku as individual FG forecast 
  channel_mutate <-channel_forecast[,-c(2,4,5)]
  channel_mutate1 <-cbind(channel_mutate[,c("Sku","Channel")],channel_mutate[,c(3:ncol(channel_mutate))])
  channel_mutate1[is.na(channel_mutate1)]<-0
  channel_group<-group_by(channel_mutate1,Sku, Channel)
  df_FG <-summarise_all(channel_group,funs(sum))
  df_FG$Channel<-paste(df_FG$Channel,"(FG)", sep=" ")
  
  # Combine these 2
  df_combine <-rbind(as.data.frame(df_kits),as.data.frame(df_FG))
  col_names <-colnames(df_combine)[3:ncol(df_combine)]
  trimmed_names <- gsub(pattern = "\\.",replacement = " ",col_names)
  colnames(df_combine)[3:ncol(df_combine)]=trimmed_names
  
  #### get rid of some one-off components being added into S&OP forecast
  df_forecast_combine=df_combine[which(df_combine$Sku%in%new_forecastSKU_list),]
  df_forecast_combine=df_forecast_combine[order(df_forecast_combine$Sku),]
  df_forecast_combine$concat=paste(df_forecast_combine$Sku,df_forecast_combine$Channel,sep="")
  df_forecast_combine=df_forecast_combine[,c(ncol(df_forecast_combine),2:ncol(df_forecast_combine)-1)]
  
  #############################################################################################################
  ########## Section 2 --> capture SKU's individual & Kit consumption of this month ###########################
  #############################################################################################################
  
  ###### For Overall -->  SKU consumption ######
  
  
  ## Sku as individual FG consumption
  
  df_FG_consumption <- MTD_consumption_summ[,c('time_period','Item #','MTD_consumed')]
  colnames(df_FG_consumption)<-c('time_period','Sku','MTD_consumed')
  
  
  df_consumption_group <-group_by(df_FG_consumption, Sku)
  df_consumption_summ <-summarise(df_consumption_group,MTD_consumed=sum(MTD_consumed), time_period=first(time_period))
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
  df_channel_comp_consumption=df_channel_comp_consumption[,c('concat','Sku',"Channel",'MTD_consumed','time_period')]
  
  ## Sku as individual FG consumption
  
  MTD_consumption_total_summ$Channel=paste(MTD_consumption_total_summ$Channel,"(FG)",sep = " ")
  MTD_consumption_total_summ$concat=paste(MTD_consumption_total_summ$`Item #`,MTD_consumption_total_summ$Channel,sep="")
  
  df_channel_FG_consumption <- MTD_consumption_total_summ[,c('concat','Item #',"Channel",'MTD_consumed','time_period')]
  colnames(df_channel_FG_consumption)<-c('concat','Sku',"Channel",'MTD_consumed','time_period')
  
  ## combine these 2
  df_channel_consumption <-rbind(df_channel_FG_consumption,df_channel_comp_consumption)
  df_channel_consumption_group <-group_by(df_channel_consumption, concat)
  df_channel_consumption_summ <-summarise(df_channel_consumption_group,Sku=first(Sku),MTD_consumed=sum(MTD_consumed), MTD_consumed_timeframe=first(time_period))
  df_channel_consumption_combine=as.data.frame(df_channel_consumption_summ)
  
  #### get rid of some one-off components being added into S&OP forecast
  df_channel_consumption_combine=df_channel_consumption_combine[which(df_channel_consumption_combine$Sku%in%new_forecastSKU_list),]
  df_channel_consumption_combine=df_channel_consumption_combine[order(df_channel_consumption_combine$Sku),]
  
  #### Remove zero MTD shipment & combine shipped & pending for shipment
  df_channel_consumption_combine=df_channel_consumption_combine[which(df_channel_consumption_combine$MTD_consumed!=0),]
  
  #############################################################################################################
  #############################################################################################################
  
  
  #### Total SKUs need to be on the Forecast_Inventory_Report
  SKU_on_report <-as.data.frame(unique(df_forecast_combine$Sku))
  names(SKU_on_report)='Sku'
  
  
  SKU_added <-as.data.frame(unique(read.xlsx("Forecast & Inventory Planning Report.xlsm", sheet="Perricone",startRow=4)[,4]))
  names(SKU_added)='Sku'
  
  mutual_SKUs =merge(SKU_added,SKU_on_report)
  
  ### to filter out SKUs that need to be added
  SKU_to_add=subset(SKU_on_report, !(Sku%in%mutual_SKUs$Sku))
  
  
  ###############################################################################################################################
  ###### Section 3 --> capture On Hand & On Order Kit units that  need to be reconciled & put back to Component Supply  #########
  ###############################################################################################################################
  
  
  
  
  ##################################################
  ####  1. get inventory OH QTY from Kits ##########
  ##################################################
  
  names(OH_available)[3]='QTY'
  names(OH_reserve)[3]='QTY'
  OH_rbind=rbind(OH_available,OH_reserve)
  OH_Kit_reconcile = summarise(group_by(OH_rbind,Sku),QTY=sum(QTY))
  
  
  df_Kit_Inv_Reconcile <-data.frame()
  
  for (i in new_forecastSKU_list){
    df<-BOM_list[which(BOM_list$CMPTITNM_C==i),c(1,5)]
    
    renew_df<-OH_Kit_reconcile[which(OH_Kit_reconcile$Sku%in%df$Item_Number_FGI),]
    
    renew1_df<-merge(df,renew_df,by.x = 'Item_Number_FGI', by.y = 'Sku', all.y = TRUE)
    renew1_df[is.na(renew1_df)]<-0
    renew1_df$OH_units_reconcile<-as.numeric(renew1_df$CMPITQTY_C)*renew1_df$QTY
    renew1_df <-renew1_df[,-c(2,3)]
    
    if(length(rownames(renew1_df))==0){
      renew1_df$Sku=as.list(NULL)
    }
    else{
      renew1_df$Sku =i
      
      renew1_df$Item_Number_FGI=paste(renew1_df$Item_Number_FGI,"'s OH inventory",sep = " ")
    }
    df_Kit_Inv_Reconcile=rbind(df_Kit_Inv_Reconcile,renew1_df)
    
  }
  
  df_Kit_Inv_Reconcile=df_Kit_Inv_Reconcile[which(df_Kit_Inv_Reconcile$OH_units_reconcile!=0),]
  
  df_Kit_Inv_Reconcile_summ=summarise(group_by(df_Kit_Inv_Reconcile,Sku),QTY=sum(OH_units_reconcile),Note=paste(`Item_Number_FGI`,collapse = "; "))

  ##################################################
  ########  2. get open PO QTY from Kits ###########
  ##################################################
  
  OO_Kit_reconcile=summarise(group_by(open_PO,Sku),QTY=sum(QTY))
  df_Kit_OO_Reconcile <-data.frame()
  
  for (i in new_forecastSKU_list){
    df<-BOM_list[which(BOM_list$CMPTITNM_C==i),c(1,5)]
    
    renew_df<-OO_Kit_reconcile[which(OO_Kit_reconcile$Sku%in%df$Item_Number_FGI),]
    
    renew1_df<-merge(df,renew_df,by.x = 'Item_Number_FGI', by.y = 'Sku', all.y = TRUE)
    renew1_df[is.na(renew1_df)]<-0
    renew1_df$OO_units_reconcile<-as.numeric(renew1_df$CMPITQTY_C)*renew1_df$QTY
    renew1_df <-renew1_df[,-c(2,3)]
    
    if(length(rownames(renew1_df))==0){
      renew1_df$Sku=as.list(NULL)
    }
    else{
      renew1_df$Sku =i
      
      renew1_df$Item_Number_FGI=paste(renew1_df$Item_Number_FGI,"'s open PO",sep = " ")
    }
    df_Kit_OO_Reconcile=rbind(df_Kit_OO_Reconcile,renew1_df)
    
  }
  
  df_Kit_OO_Reconcile=df_Kit_OO_Reconcile[which(df_Kit_OO_Reconcile$OO_units_reconcile!=0),]
  
  df_Kit_OO_Reconcile_summ=summarise(group_by(df_Kit_OO_Reconcile,Sku),QTY=sum(OO_units_reconcile),Note=paste(`Item_Number_FGI`,collapse = "; "))
  
  ##### wrap up the data
  OH_reconcile=as.data.frame(df_Kit_Inv_Reconcile_summ[which(df_Kit_Inv_Reconcile_summ$Sku%in%as.vector(SKU_added[,1])),])
  OO_reconcile=as.data.frame(df_Kit_OO_Reconcile_summ[which(df_Kit_OO_Reconcile_summ$Sku%in%as.vector(SKU_added[,1])),])
  
  #############################################################################################################
  ################################ Section 4--> exporting file ################################################
  #############################################################################################################
  
  xlsx::write.xlsx(SKU_to_add,"masterplan_raw.xlsx",sheetName='SKU_to_add', row.names=FALSE)
  xlsx::write.xlsx(df_forecast_combine,"masterplan_raw.xlsx",sheetName='Forecast', append = TRUE,row.names=FALSE)
  xlsx::write.xlsx(open_PO,"masterplan_raw.xlsx",sheetName='open_PO',append = TRUE, row.names=FALSE)
  xlsx::write.xlsx(df_consumption_combine,"masterplan_raw.xlsx",sheetName='Consumption',append = TRUE, row.names=FALSE)
  xlsx::write.xlsx(df_channel_consumption_combine,"masterplan_raw.xlsx",sheetName='Consumption_channel',append = TRUE, row.names=FALSE)
  xlsx::write.xlsx(OH_total,"masterplan_raw.xlsx",sheetName='OH_total',append = TRUE, row.names=FALSE)
  xlsx::write.xlsx(OH_BIUB,"masterplan_raw.xlsx",sheetName='BIUB',append = TRUE, row.names=FALSE)
  xlsx::write.xlsx(OH_reconcile,"masterplan_raw.xlsx",sheetName='reconcile_OH',append = TRUE, row.names=FALSE)
  xlsx::write.xlsx(OO_reconcile,"masterplan_raw.xlsx",sheetName='reconcile_OO',append = TRUE, row.names=FALSE)
  ########################################################################################################
  ######################### AFTER GENERATING THE FILE ####################################################
  ###################### NEED TO CHANGE SKU format #######################################################
  ##################### ALT + A + E ---> FINISH ##########################################################
  ########################################################################################################
  
  
  
  
  
  
  
  
  
