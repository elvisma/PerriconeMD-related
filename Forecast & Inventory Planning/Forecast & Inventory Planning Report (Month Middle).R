### Update MPS #####

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

channel_start_col=19 #Sep col 



####################################
######## Start loading data ########
####################################

setwd("C:/Users/ema/Desktop/projects/Perricone Dashboard")

## 1. load BOM list
BOM_list <-as.data.frame(read_excel("SmartList BOM Extract.xlsx", sheet=1,col_types='text'))[,c(1,2,9,10,11,13)]

## 2. S&OP forecast data 
setwd("O:/Supply Chain/Demand Planning/Forecast Master")

channel_forecast <-read.xlsx("Master Forecast.xlsm", sheet = 'Channel Forecast',startRow = 1)[,c(1:5,channel_start_col:22,24:35)]
channel_forecast=channel_forecast[-c(1),]

#change the "character" into "Numeric"



channel_forecast[,c(6:ncol(channel_forecast))]=sapply(channel_forecast[,c(6:ncol(channel_forecast))], as.numeric)
channel_forecast[,c(6:ncol(channel_forecast))][is.na(channel_forecast[,c(6:ncol(channel_forecast))])]=0
channel_forecast$total_forecast=rowSums(channel_forecast[,c(6:ncol(channel_forecast))])
channel_forecast=channel_forecast[which(channel_forecast$total_forecast>0),-c(ncol(channel_forecast))]

## 3. get the inventory report

setwd("C:/Users/ema/Desktop/projects/Forecast_Inventory_Report")

OH_available <-read.xlsx('OH_inventory_report.xlsx',sheet = 'Avai Inventory',startRow = 4)
colnames(OH_available)=c("Sku","Description","OH available")

OH_reserve <-read.xlsx('OH_inventory_report.xlsx',sheet = 'Reserve Inventory',startRow = 4)
colnames(OH_reserve)=c("Sku","Description","OH reserve")

OH_total=merge(OH_available,OH_reserve,all.x = TRUE)
OH_total$`OH reserve`[is.na(OH_total$`OH reserve`)]=0

## 4. get BIUB information

avai_BIUB <-read.xlsx('OH_inventory_report.xlsx',sheet = 'Avai BIUB',startRow = 4)
reserv_BIUB <-read.xlsx("OH_inventory_report.xlsx", sheet='Reserve BIUB', startRow=4)


avai_BIUB <-avai_BIUB[,c(1,ncol(avai_BIUB)-1,ncol(avai_BIUB))]
colnames(avai_BIUB)=c("Sku","avai_BIUB","aging?")

reserv_BIUB <-reserv_BIUB[,c(1,ncol(reserv_BIUB))]
colnames(reserv_BIUB)=c("Sku","reserv_BIUB")

OH_BIUB=merge(avai_BIUB,reserv_BIUB,all= TRUE)

OH_BIUB$reserv_BIUB[is.na(OH_BIUB$reserv_BIUB)]="no BIUB information"
OH_BIUB$`aging?`[is.na(OH_BIUB$`aging?`)]='fine shelf life'

## 5. get the open PO report
open_PO <-read.xlsx('Open_PO_report.xlsx',sheet = 'Pivot PO',startRow = 1)
colnames(open_PO)=c("Sku","Month","PO number",'QTY')



### SKUs on the MPS report
SKU_added <-as.data.frame(unique(read.xlsx("Forecast & Inventory Planning Report (Master).xlsm", sheet="Perricone",startRow=4)[,4]))
names(SKU_added)='Sku'

### combine to get the whole Sku list need to check forecast/consumption
WIP_SKU<-unique(BOM_list[which(BOM_list$Item_Class_Code_CI=='INTERMED'),c(3)])
FG_SKU <-as.vector(unique(c(channel_forecast$Sku,as.vector(SKU_added$Sku))))

forecasted_SKU<-unique(c(WIP_SKU,FG_SKU))
## get rid of some one-off components being added into S&OP forecast
forecasted_SKU <-forecasted_SKU[which(forecasted_SKU%in%BOM_list$Item_Number_FGI)]

new_forecastSKU_list <-unique(c(forecasted_SKU,as.vector(channel_forecast$Sku)))



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

df_forecast_combine$total_forecast=rowSums(df_forecast_combine[,c(4:ncol(df_forecast_combine))])
df_forecast_combine=df_forecast_combine[which(df_forecast_combine$total_forecast>0),-c(ncol(df_forecast_combine))]
#############################################################################################################
########## Section 2 --> capture SKU's individual & Kit consumption of this month ###########################
#############################################################################################################


shift_mtd_consumption_shipped <-read.xlsx("shift_MTD_consumption_breakdown.xlsx", sheet="Shipped",startRow = 2)[,c(1:5)]
shift_mtd_consumption_processed <-read.xlsx("shift_MTD_consumption_breakdown.xlsx", sheet="Processed",startRow = 2)[,c(1:5)]

shift_mtd_consumption<-rbind(shift_mtd_consumption_shipped,shift_mtd_consumption_processed)




##### SKU consumption by channel #########
shift_mtd_consumption_group<-group_by(shift_mtd_consumption,Channel, Indicator, Sku)
shift_mtd_consumption_summ <-summarise(shift_mtd_consumption_group,consump_QTY=sum(QTY))
shift_mtd_consumption_summ$note<-paste(shift_mtd_consumption_summ$Indicator, shift_mtd_consumption_summ$consump_QTY,sep = ": ")
shift_mtd_consumption_summ$bool<-startsWith(shift_mtd_consumption_summ$note,"MTD")
shift_mtd_consumption_summ$MTD_consump<-as.numeric(shift_mtd_consumption_summ$bool*shift_mtd_consumption_summ$consump_QTY)

shift_mtd_consumption_summ_new=shift_mtd_consumption_summ[,c("Sku","Channel","MTD_consump","note")]
shift_mtd_consumption_summ_new_group=group_by(shift_mtd_consumption_summ_new,Sku, Channel)

p1 <-function(v){
  if(startsWith(v,"MTD")) " " else v
}

shift_mtd_consumption_summ_new_group$note=sapply(shift_mtd_consumption_summ_new_group$note,p1)


p <- function(v) {
  Reduce(f=paste, x = paste(v,",  ",sep = " "))
}



# concatenate all notes together
shift_mtd_consumption_summ_new_summ=summarise(shift_mtd_consumption_summ_new_group,MTD_consumed=sum(MTD_consump),Note=p(as.character(note)))

# remove "," from the end of the strings
shift_mtd_consumption_summ_new_summ$Note=substr(shift_mtd_consumption_summ_new_summ$Note,1,nchar(shift_mtd_consumption_summ_new_summ$Note)-1)

shift_mtd_consumption_summ_new_summ$Note=trimws(gsub(",","",shift_mtd_consumption_summ_new_summ$Note),which=c("left"))

names(shift_mtd_consumption_summ_new_summ)[1]="Item #"


df_channel_comp_consumption <-data.frame()

for (i in new_forecastSKU_list){
  df<-BOM_list[which(BOM_list$CMPTITNM_C==i),c(1,2,5)]
  
  renew_df<-shift_mtd_consumption_summ_new_summ[which(shift_mtd_consumption_summ_new_summ$`Item #`%in%df$Item_Number_FGI),]
  renew1_df<-merge(df,renew_df,by.x = 'Item_Number_FGI', by.y = 'Item #', all.y = TRUE)
  renew1_df[is.na(renew1_df)]<-0
  renew1_df$MTD_consumed<-as.numeric(renew1_df$CMPITQTY_C)*renew1_df$MTD_consumed
  
  
  if(length(rownames(renew1_df))==0){
    renew1_df$Sku=as.list(NULL)
  }
  else{
    renew1_df$Note<-paste("from",renew1_df$Item_Number_FGI,sep = " ")
    renew1_df <-renew1_df[,-c(1,2,3)]
    renew1_df$Sku =i
    renew1_df$Channel=paste(renew1_df$Channel,"(Kit)",sep = " ")
  }
  df_channel_comp_consumption=rbind(df_channel_comp_consumption,renew1_df)
  
}

df_channel_comp_consumption$concat=paste(df_channel_comp_consumption$Sku,df_channel_comp_consumption$Channel,sep="")
df_channel_comp_consumption=df_channel_comp_consumption[,c('concat','Sku',"Channel",'MTD_consumed','Note')]

shift_mtd_consumption_summ_new_summ$Channel=paste(shift_mtd_consumption_summ_new_summ$Channel,"(FG)",sep = " ")
shift_mtd_consumption_summ_new_summ$concat=paste(shift_mtd_consumption_summ_new_summ$`Item #`,shift_mtd_consumption_summ_new_summ$Channel,sep="")
df_channel_FG_consumption <- shift_mtd_consumption_summ_new_summ[,c('concat','Item #',"Channel",'MTD_consumed','Note')]
colnames(df_channel_FG_consumption)<-c('concat','Sku',"Channel",'MTD_consumed','Note')


## combine these 2
df_channel_consumption <-rbind(df_channel_FG_consumption,df_channel_comp_consumption)
df_channel_consumption_group <-group_by(df_channel_consumption, concat)
df_channel_consumption_summ <-summarise(df_channel_consumption_group,Sku=first(Sku),MTD_consumed=sum(MTD_consumed), Note=p(as.character(Note)))
df_channel_consumption_combine=as.data.frame(df_channel_consumption_summ)
df_channel_consumption_combine$Note=gsub(",","",df_channel_consumption_combine$Note)
df_channel_consumption_combine$Note[startsWith(df_channel_consumption_combine$Note," ")]="no shifted orders"

#### get rid of some one-off components being added into S&OP forecast
df_channel_consumption_combine=df_channel_consumption_combine[which(df_channel_consumption_combine$Sku%in%new_forecastSKU_list),]
df_channel_consumption_combine=df_channel_consumption_combine[order(df_channel_consumption_combine$Sku),]


##### SKU consumptoin overall ########

df_consumption_combine=as.data.frame(summarise(group_by(shift_mtd_consumption_summ,Sku),MTD_consumed=sum(MTD_consump)))


#### Total SKUs need to be on the Forecast_Inventory_Report
SKU_on_report <-as.data.frame(unique(df_forecast_combine$Sku))
names(SKU_on_report)='Sku'




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
OH_Kit_reconcile_summ = summarise(group_by(OH_rbind,Sku),QTY=sum(QTY))
OH_Kit_reconcile=OH_Kit_reconcile_summ[which(OH_Kit_reconcile_summ$Sku%in%new_forecastSKU_list),]



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
    
    renew1_df$FGI_details=paste(renew1_df$Item_Number_FGI,"'s OH inventory",sep = " ")
  }
  df_Kit_Inv_Reconcile=rbind(df_Kit_Inv_Reconcile,renew1_df)
  
}

df_Kit_Inv_Reconcile=df_Kit_Inv_Reconcile[which(df_Kit_Inv_Reconcile$OH_units_reconcile!=0),]

###### remove those Kits items that does not have active forecasts (in order to not over counting kitted inventory)#### 

df_Kit_Inv_Reconcile=df_Kit_Inv_Reconcile[which(df_Kit_Inv_Reconcile$Item_Number_FGI%in%as.vector(unique(channel_forecast$Sku))),]


df_Kit_Inv_Reconcile_summ=summarise(group_by(df_Kit_Inv_Reconcile,Sku),QTY=sum(OH_units_reconcile),Note=paste(`FGI_details`,collapse = "; "))


##### wrap up the data
OH_reconcile=as.data.frame(df_Kit_Inv_Reconcile_summ[which(df_Kit_Inv_Reconcile_summ$Sku%in%as.vector(SKU_added[,1])),])

##################################################
########  2. get open PO QTY from Kits ###########
##################################################

OO_Kit_reconcile=summarise(group_by(open_PO,Sku, Month),QTY=sum(QTY))
df_Kit_OO_Reconcile <-data.frame()

for (i in new_forecastSKU_list){
  df<-BOM_list[which(BOM_list$CMPTITNM_C==i),c(1,5)]
  
  renew_df<-OO_Kit_reconcile[which(OO_Kit_reconcile$Sku%in%df$Item_Number_FGI),]
  
  renew1_df<-merge(df,renew_df,by.x = 'Item_Number_FGI', by.y = 'Sku', all.y = TRUE)
  renew1_df[is.na(renew1_df)]<-0
  renew1_df$OO_units_reconcile<-as.numeric(renew1_df$CMPITQTY_C)*renew1_df$QTY
  renew1_df <-renew1_df[,-c(2,4)]
  
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

df_Kit_OO_Reconcile_summ=summarise(group_by(df_Kit_OO_Reconcile,Sku, Month),QTY=sum(OO_units_reconcile),Note=paste(`Item_Number_FGI`,collapse = "; "))

##### wrap up the data
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









