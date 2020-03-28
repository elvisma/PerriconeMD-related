
####### PerriconeMD dashboard ########

# Read Libraries
#instAll Franchises.packages("devtools") 
#devtools::install_github("tidyverse/readxl")
#library("xlsx")

library("formattable")
library("openxlsx")
library("readxl")
#library("reshape")
library("tidyr")
options(scipen=999)
library("lubridate")
library("rsconnect")
library("shiny")
#library("stringi")
library("highcharter")
library("shinythemes")
library("shinydashboard")
library("dplyr")
library("DT")

setwd("C:/Users/ema/Desktop/projects/Perricone Dashboard")

######Profit & Margin, Pie Chart breakdown for Acct/Franchise, bar column for Top SKUs for different date ranges #############
financial_PMD <-read_excel("financial raw data.xlsx", sheet='PMD.com (S&OP actuals)', skip=0)
financial_PMD_gather <- gather(financial_PMD,"Month", "QTY", c(10:ncol(financial_PMD)))
financial_PMD_gather$COGS[which(is.na(financial_PMD_gather$COGS))]<-0
financial_PMD_gather$`Selling Price`[which(is.na(financial_PMD_gather$`Selling Price`))]<-0
financial_PMD_gather$`ExtendedCost` <-as.numeric(financial_PMD_gather$COGS)*financial_PMD_gather$QTY
financial_PMD_gather$`ExtendedRevenue`<-as.numeric(financial_PMD_gather$`Selling Price`)*financial_PMD_gather$QTY

colnames(financial_PMD_gather)<-gsub(pattern='\\s+',replacement="", colnames(financial_PMD_gather))

cols <-c('Account','Sku','Description','Franchise','Month','QTY','ExtendedCost','ExtendedRevenue')
financial_PMD_ready <-financial_PMD_gather[,cols]


financial_adrian<-read_excel("financial raw data.xlsx", sheet='Data Log (Adrian)', skip=0,range = cell_cols(c("A:I")))
colnames(financial_adrian)<-gsub(pattern='\\s+',replacement="", colnames(financial_adrian))
financial_adrian_ready <- financial_adrian[which(financial_adrian$SOPtype=='Invoice'&!is.na(financial_adrian$Franchise)),cols]
financial_combine <- rbind(financial_adrian_ready,financial_PMD_ready)
financial_combine <-as.data.frame(financial_combine)
financial_combine$QTY[which(is.na(financial_combine$QTY))] <-0
financial_combine$ExtendedCost[which(is.na(financial_combine$ExtendedCost))]<-0
financial_combine$ExtendedRevenue[which(is.na(financial_combine$ExtendedRevenue))] <-0
financial_combine$Month<-parse_date_time(financial_combine$Month,"my")




###
account_variable4
franchise_variable4
daterange_variable3 [last month, last 3 month, last 6 months, last 12 month --> best to be select bar (time frame slider bar)]

!!!SKU bars visualize Revenues ONLY
!!!Franchise/Account Pies visualize Revenues & QTY 

#############################################################################################
###############################GOOD#########################################################
acctpie_function<-reactive({
if(input$franchise_variable4=='All Franchises'){
  if(input$acct_variable4=='Company'){
    financial_group <-group_by(financial_combine,Account,Month)
    financial_summ1 <-summarise(financial_group,QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
    
  }
  
  else {
    financial_group <-group_by(financial_combine,Account,Month)
    financial_summ1 <-summarise(financial_group,QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
    financial_summ1 <-financial_summ1[which(financial_summ1$Account==input$acct_variable4),]
    #financial_summ1 <-financial_summ1[which(financial_summ1$Account=='APAC+LA'),]
  }}
  
else{
  if(input$acct_variable4=='Company'){
    financial_combine<-financial_combine[which(financial_combine$Franchise==input$franchise_variable4),]
    financial_group <-group_by(financial_combine,Account,Month)
    financial_summ1 <-summarise(financial_group,QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
    
  }
  
  else {
    financial_combine<-financial_combine[which(financial_combine$Franchise==input$franchise_variable4),]
    financial_group <-group_by(financial_combine,Account,Month)
    financial_summ1 <-summarise(financial_group,QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
    financial_summ1 <-financial_summ1[which(financial_summ1$Account==input$acct_variable4),]
    #financial_summ1 <-financial_summ1[which(financial_summ1$Account=='APAC+LA'),]
  }}
  

  
  #financial_summ1$Month<-parse_date_time(financial_summ1$Month,"my")
  financial_summ1<-financial_summ1[order(financial_summ1$Month),]
  financial_summ1 <-financial_summ1[which(financial_summ1$Month>=input$daterange_variable3[1]&financial_summ1$Month<=input$daterange_variable3[2]),]
  #financial_summ1 <-financial_summ1[which(financial_summ1$Month>=as.Date('2019-10-01')&financial_summ1$Month<=as.Date('2020-01-01')),]
  
  
  return(financial_summ1)
  
})



franpie_function<-reactive({
if(input$acct_variable4=='Company'){
  if(input$franchise_variable4=='All Franchises'){
    financial_group <-group_by(financial_combine,Franchise,Month)
    financial_summ2 <-summarise(financial_group,QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
    
  }
  
  else {
    financial_group <-group_by(financial_combine,Franchise,Month)
    financial_summ2 <-summarise(financial_group,QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
    financial_summ2 <-financial_summ2[which(financial_summ2$Franchise==input$franchise_variable4),]
    #financial_summ2 <-financial_summ2[which(financial_summ2$Franchise=='Acne'),]
  }}
  
else {
  if(input$franchise_variable4=='All Franchises'){
      financial_combine<-financial_combine[which(financial_combine$Account==input$acct_variable4),]
      financial_group <-group_by(financial_combine,Franchise,Month)
      financial_summ2 <-summarise(financial_group,QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
      
    }
    
  else {
      financial_combine<-financial_combine[which(financial_combine$Account==input$acct_variable4),]
      financial_group <-group_by(financial_combine,Franchise,Month)
      financial_summ2 <-summarise(financial_group,QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
      financial_summ2 <-financial_summ2[which(financial_summ2$Franchise==input$franchise_variable4),]
      #financial_summ2 <-financial_summ2[which(financial_summ2$Franchise=='Acne'),]
    }} 
  
  #financial_summ2$Month<-parse_date_time(financial_summ2$Month,"my")
  financial_summ2<-financial_summ2[order(financial_summ2$Month),]
  financial_summ2 <-financial_summ2[which(financial_summ2$Month>=input$daterange_variable3[1]&financial_summ2$Month<=input$daterange_variable3[2]),]
  #financial_summ2 <-financial_summ2[which(financial_summ2$Month>=as.Date('2019-10-01')&financial_summ2$Month<=as.Date('2020-01-01')),]
  
  return(financial_summ2)
  
})




skubar_function <-reactive({
  if(input$acct_variable4=='Company'&&input$franchise_variable4=='All Franchises'){
    financial_group <-group_by(financial_combine,Sku,Month)
    financial_summ3 <-summarise(financial_group,QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
    
  }
  else if(input$acct_variable4!='Company'&&input$franchise_variable4=='All Franchises'){
    financial_group <-group_by(financial_combine,Sku,Month,Account)
    financial_summ3 <-summarise(financial_group,QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
    financial_summ3 <-financial_summ3[which(financial_summ3$Account==input$acct_variable4),-c(3)]
  }
  
  else if(input$acct_variable4=='Company'&&input$franchise_variable4!='All Franchises'){
    financial_group <-group_by(financial_combine,Sku,Month,Franchise)
    financial_summ3 <-summarise(financial_group,QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
    financial_summ3 <-financial_summ3[which(financial_summ3$Franchise==input$franchise_variable4),-c(3)]
  }
  
  else {
    financial_group <-group_by(financial_combine,Sku,Month,Account,Franchise)
    financial_summ3 <-summarise(financial_group,QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
    financial_summ3 <-financial_summ3[which(financial_summ3$Franchise==input$franchise_variable4&financial_summ3$Account==input$acct_variable4),-c(3,4)]
  }
  #financial_summ3$Month<-parse_date_time(financial_summ3$Month,"my")
  financial_summ3<-financial_summ3[order(financial_summ3$Month),]
  financial_summ3 <-financial_summ3[which(financial_summ3$Month>=input$daterange_variable3[1]&financial_summ3$Month<=input$daterange_variable3[2]),]
  #financial_summ3 <-financial_summ3[which(financial_summ3$Month>=as.Date('2019-10-01')&financial_summ3$Month<=as.Date('2020-01-01')),]
  
  return(financial_summ3)
})


#############################################################################################
###############################GOOD#########################################################

skubar_function() --> financial_summ3
franpie_function() --> financial_summ2
acctpie_function() --> financial_summ1

financial_summ3[order(-financial_summ3$Revenues),][1:10,1]



output$acct_dollarpie  acct_qtypie
   franchise_dollarpie franchise_qtypie
   topBars
   topTables


#### testing date format here #####
   
   output$top10GGbar <- renderHighchart({
     highchart() %>%
       hc_add_series_df(data = Top10G()[1:10,], type = "bar", x = SKU, y = Total) %>%
       hc_xAxis(list(categories = Top10G()[1:10,1])) %>%
       hc_legend(enabled = FALSE) %>%
       hc_tooltip(valueDecimals = 2, valuePrefix = "$")
   })
   
   
   output$ZinusInv_ClassPie <-renderHighchart({
     ################ Set Color ########################
     colors <- revalue(ZinusInv_Class$Class, c("A" = "#32CD32","B" = "#ADFF2F","C" = "#98FB98","NA" = "#FF1493", "NB" = "#FF69B4", "NC" = "#FFB6C1",
                                               "New" = "#E44A12", "Discontinued" = "#A9A9A9"))
     highchart() %>% 
       hc_chart(type = "pie") %>% 
       hc_add_series_labels_values(labels = ZinusInv_Class$Class, values = ZinusInv_Class$`Last Week`,colors = colors)%>% 
       hc_tooltip(pointFormat = paste('${point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
       hc_title(text = "INVENTORY DISTRIBUTION BY CLASS (LAST WEEK)")
     #ggplot(data = Zinus_ClassData(), aes(x = "", y = Percentage, fill = Class )) + 
     # geom_bar(stat = "identity", position = position_fill()) +
     #geom_text(aes(label = Label), position = position_fill(vjust = 0.5)) +
     #coord_polar(theta = "y") +
     #facet_wrap(~List )  +
     #theme(
     # axis.title.x = element_blank(),
     #axis.title.y = element_blank()) + theme(legend.position='bottom') + guides(fill=guide_legend(nrow=2,byrow=TRUE))
   })
   output$ZinusSales_ClassPie <-renderHighchart({
     colors <- revalue(ZinusSales_Class$Class, c("A" = "#32CD32","B" = "#ADFF2F","C" = "#98FB98","NA" = "#FF1493", "NB" = "#FF69B4", "NC" = "#FFB6C1",
                                                 "New" = "#E44A12", "Discontinued" = "#A9A9A9" ))
     highchart() %>% 
       hc_chart(type = "pie") %>% 
       hc_add_series_labels_values(labels = ZinusSales_Class$Class, values = ZinusSales_Class$`Last Week`,colors = colors)%>% 
       hc_tooltip(pointFormat = paste('${point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
       hc_title(text = "SALES DISTRIBUTION BY CLASS (LAST WEEK)")
     
   })
   output$ZinusInv_CategoryPie<-renderHighchart({
     highchart() %>% 
       hc_chart(type = "pie") %>% 
       hc_add_series_labels_values(labels = ZinusInv_Category$Category, values = ZinusInv_Category$`Last Week`)%>% 
       hc_tooltip(pointFormat = paste('${point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
       hc_title(text = "INVENTORY DISTRIBUTION BY CATEGORY (LAST WEEK)")
   }) 
   
   output$ZinusSales_CategoryPie <-renderHighchart({
     highchart() %>% 
       hc_chart(type = "pie") %>% 
       hc_add_series_labels_values(labels = ZinusSales_Category$Category, values = ZinusSales_Category$`Last Week`)%>% 
       hc_tooltip(pointFormat = paste('${point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
       hc_title(text = "SALES DISTRIBUTION BY CATEGORY (LAST WEEK)")
     
   })
   
   


test_df=as.Date(unique(financial_combine$Month))





library(shiny)

#### Pay Attention!!!

#### there are 2 elements of input$dateRange ---> input$dateRange[1]  &  input$dateRange[2]
#### testing  --->  df[which(df$month>=input$dateRange[1]&df$month<=input$dateRange[2]),]
#### testing  --->  group_by()
#### testing  --->  summarise()


ui <- basicPage(dateRangeMonthsInput('dateRange',label = "select range : ",format = "mm/yyyy",start = test_df[length(test_df)-3], end=test_df[length(test_df)],min=head(test_df,1),max=tail(test_df,1),startview = "year",separator = " - "),
                textOutput("SliderText")
)
server <- shinyServer(function(input, output, session){
  
  Dates <- reactiveValues()
  observe({
    Dates$SelectedDates <- c(as.character(format(input$dateRange[1],format = "%m/%d/%Y")),as.character(format(input$dateRange[2],format = "%m/%d/%Y")))
    #Dates$DateRangeLength <- length(input$dateRange)
  })
  output$SliderText <- renderText({Dates$SelectedDates})
  #output$SliderText <- renderText({Dates$DateRangeLength})
})
shinyApp(ui = ui, server = server)




###############################################################################
###############################################################################


https://stackoverflow.com/questions/31152960/display-only-months-in-daterangeinput-or-dateinput-for-a-shiny-app-r-programmin/38974106


###########################################################################################
###################### customize date Range Months Input ##################################
###########################################################################################

dateRangeMonthsInput <- function(inputId, label, start = NULL, end = NULL,
                                 min = NULL, max = NULL, format = "yyyy-mm-dd", startview = "month",
                                 minviewmode="months", # added manually
                                 weekstart = 0, language = "en", separator = " to ", width = NULL) {
  
  # If start and end are date objects, convert to a string with yyyy-mm-dd format
  # Same for min and max
  if (inherits(start, "Date"))  start <- format(start, "%Y-%m-%d")
  if (inherits(end,   "Date"))  end   <- format(end,   "%Y-%m-%d")
  if (inherits(min,   "Date"))  min   <- format(min,   "%Y-%m-%d")
  if (inherits(max,   "Date"))  max   <- format(max,   "%Y-%m-%d")
  
  htmltools::attachDependencies(
    div(id = inputId,
        class = "shiny-date-range-input form-group shiny-input-container",
        style = if (!is.null(width)) paste0("width: ", validateCssUnit(width), ";"),
        
        controlLabel(inputId, label),
        # input-daterange class is needed for dropdown behavior
        div(class = "input-daterange input-group",
            tags$input(
              class = "input-sm form-control",
              type = "text",
              `data-date-language` = language,
              `data-date-weekstart` = weekstart,
              `data-date-format` = format,
              `data-date-start-view` = startview,
              `data-date-min-view-mode` = minviewmode, # added manually
              `data-min-date` = min,
              `data-max-date` = max,
              `data-initial-date` = start
            ),
            span(class = "input-group-addon", separator),
            tags$input(
              class = "input-sm form-control",
              type = "text",
              `data-date-language` = language,
              `data-date-weekstart` = weekstart,
              `data-date-format` = format,
              `data-date-start-view` = startview,
              `data-date-min-view-mode` = minviewmode, # added manually
              `data-min-date` = min,
              `data-max-date` = max,
              `data-initial-date` = end
            )
        )
    ),
    datePickerDependency
  )
}

`%AND%` <- function(x, y) {
  if (!is.null(x) && !is.na(x))
    if (!is.null(y) && !is.na(y))
      return(y)
  return(NULL)
}

controlLabel <- function(controlName, label) {
  label %AND% tags$label(class = "control-label", `for` = controlName, label)
}

# the datePickerDependency is taken from https://github.com/rstudio/shiny/blob/master/R/input-date.R
datePickerDependency <- htmltools::htmlDependency(
  "bootstrap-datepicker", "1.6.4", c(href = "shared/datepicker"),
  script = "js/bootstrap-datepicker.min.js",
  stylesheet = "css/bootstrap-datepicker3.min.css",
  # Need to enable noConflict mode. See #1346.
  head = "<script>
 (function() {
 var datepicker = $.fn.datepicker.noConflict();
 $.fn.bsDatepicker = datepicker;
 })();
 </script>")


###########################################################################################
###################### customize date Range Months Input ##################################
###########################################################################################


B2B_list <- c("B2B customers","Int'l"="APAC+LA","Amazon","Bloomingdales","Costco","Dillard's","EC Scott","L&T"="Lord & Taylor","Macy's","NM"="Neiman Marcus","Nordstrom","Sephora","TSC","UK","ULTA"="Ulta","Zulily","NCO","Retail - Other")

 (fill rate chart )
franchise_variable5
acct_variable5
checksku_variable3
sku_variable3
sku_testing3  --> uiOutPut 
daterange_variable4
downloadData4
fillQTY_chart
fillrate_chart
fillrate_table


(top cut table)
daterange_variable5
skucount_variable2
downloadData5
top_cuts


output$sku_testing2 <- renderUI({
  if(!input$checksku_variable2)
    return()
  else
    selectInput("sku_variable2",paste("      ","2). Type or select SKU:",sep="   "),
                choices=sku_testing2_function(),
                selected=NULL, multiple=FALSE)
})


sku_testing2_function <-reactive({
  if(input$franchise_variable2=='All Franchises'){sku_list2<-unique(combine_data$Sku)}
  else {sku_list2<-unique(combine_data[which(combine_data$Franchise==input$franchise_variable2),1])}
  return(sku_list2)
})





observe({
  if(input$selectall2 == 0) return(NULL) 
  else if (input$selectall2%%2 == 0)
  {
    updateCheckboxGroupInput(session,"acct_variable5",choices=B2B_list, inline=TRUE)
  }
  else
  {
    updateCheckboxGroupInput(session,"acct_variable5",choices=B2B_list,selected=B2B_list, inline=TRUE)
  }
})

observe({
  if(input$selectall3 == 0) return(NULL) 
  else if (input$selectall3%%2 == 0)
  {
    updateCheckboxGroupInput(session,"franchise_variable5",choices=franchise_list2, inline=TRUE)
  }
  else
  {
    updateCheckboxGroupInput(session,"franchise_variable5",choices=franchise_list2,selected=franchise_list2, inline=TRUE)
  }
})