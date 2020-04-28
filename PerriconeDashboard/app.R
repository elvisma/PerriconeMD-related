
####### PerriconeMD dashboard ########

# Read Libraries
#install.packages("rmarkdown") 
#install.packages("knitr") 
#devtools::install_github("tidyverse/readxl")
#library("xlsx")

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

#setwd("C:/Users/ema/Desktop/projects/Perricone Dashboard")
SOP_publish_date="data source: S&OP forecasting publications (updated @ 4/21/2020) Lag0: M-0 forecast, Lag1: M-1 forecast, Lag2: M-2 forecast"
IDC_update_date='data source: IDC shipment reports (updated @ 4/22/2020)'
financial_report_date='data source: financial monthly reports (updated @ 4/7/2020)'
FR_update_date="data source: B2B Order Fulfillment Reports (updated @ 4/1/2020)"

lag_clarification_date='Lag0 forecast: April S&OP forecast;  Lag1 forecast: March S&OP forecast'

date_ranges=c("Apr 2020 (in process)","Mar 2020","Feb 2020","Jan 2020")
fill_date_ranges=c("Mar 2020","Feb 2020","Jan 2020")


############################## THRESHOLD #################################
BOM_list <-as.data.frame(read_excel("SmartList BOM Extract.xlsx", sheet=1,col_types='text'))[,c(1,2,9,10,11,13)]
SKU_list <-unique(BOM_list$Item_Number_FGI)
franchise_list <- c("Cold Plasma","HP Classics"="High Potency Classics","No Makeup"="No Makeup Skincare","VC Ester"="Vitamin C Ester","Supplements","Neuropeptide","Essential Fx"="Essential Fx Acyl Glutathione","Hypoallergenic","Mixed Franchise","Acne","Masks","Hypo CBD"="Hypoallergenic CBD","No:Rinse","Heritage","Intensive Pore","Pre:Empt","Mens","H2 EE"="H2 Elemental Energy","Re:Firm","Thio:Plex","Sample")
franchise_list2 <-franchise_list
franchise_vector <-append(franchise_list,"All Franchises", after=0)
acct_vector <- c("Company","Int'l"="APAC+LA","Amazon","Bloomingdales","Costco","Dillard's","EC Scott","GR"="Guthy Renker","L&T"="Lord & Taylor","Macy's","NM"="Neiman Marcus","Nordstrom","Other.com","PMD.com","QVC","Sephora","TSC","UK","ULTA"="Ulta","Zulily",'Retail - Other')
B2B_list <- c("B2B customers","Int'l"="APAC+LA","Amazon","Bloomingdales","Costco","Dillard's","EC Scott","L&T"="Lord & Taylor","Macy's","NM"="Neiman Marcus","Nordstrom","Sephora","TSC","UK","ULTA"="Ulta","Zulily","NCO","Retail - Other")
###########################      getting fill rate data     ###########################


###########################      SECTION 1     ###########################
### Archive different lags forecast and actuals, capture top variance SKUs from month to month ########


archive_data <-read_excel("S&OP raw data.xlsx", sheet='pivot',col_types=c("text","text","text","text","text","numeric","numeric","numeric","numeric"), skip = 1)
colnames(archive_data)<-gsub(pattern='\\s+',replacement="", colnames(archive_data))

## lag 0 ###
future_data_lag0 <-read_excel("S&OP raw data.xlsx", sheet='lag 0',skip = 0)
future_data_lag0[,c(6:ncol(future_data_lag0))][is.na(future_data_lag0[,c(6:ncol(future_data_lag0))])] <-0
future_data_lag0_new <-gather(future_data_lag0, "Month","FCST(lag0)", c(6:ncol(future_data_lag0)))
colnames(future_data_lag0_new)<-gsub(pattern='\\s+',replacement="", colnames(future_data_lag0_new))
future_data_lag0_new$ACTUAL<-0
future_data_lag0_mutate <- future_data_lag0_new[,c("Sku","Description","Franchise","Account","Month","ACTUAL","FCST(lag0)")]
future_data_lag0_group<-group_by(future_data_lag0_mutate,Sku,Description, Franchise, Account, Month)
future_data_lag0_summ <-summarise(future_data_lag0_group,ACTUAL=sum(ACTUAL), `FCST(lag0)`=sum(`FCST(lag0)`))

## lag 1 ##

future_data_lag1 <-read_excel("S&OP raw data.xlsx", sheet='lag 1',skip = 0)
future_data_lag1[,c(6:ncol(future_data_lag1))][is.na(future_data_lag1[,c(6:ncol(future_data_lag1))])] <-0
future_data_lag1_new <-gather(future_data_lag1, "Month","FCST(lag1)", c(6:ncol(future_data_lag1)))
colnames(future_data_lag1_new)<-gsub(pattern='\\s+',replacement="", colnames(future_data_lag1_new))
future_data_lag1_new$ACTUAL<-0
future_data_lag1_mutate <- future_data_lag1_new[,c("Sku","Description","Franchise","Account","Month","ACTUAL","FCST(lag1)")]
future_data_lag1_group<-group_by(future_data_lag1_mutate,Sku,Description, Franchise, Account, Month)
future_data_lag1_summ <-summarise(future_data_lag1_group,ACTUAL=sum(ACTUAL), `FCST(lag1)`=sum(`FCST(lag1)`))
## lag 2 ##


future_data_lag2 <-read_excel("S&OP raw data.xlsx", sheet='lag 2',skip = 0)
future_data_lag2[,c(6:ncol(future_data_lag2))][is.na(future_data_lag2[,c(6:ncol(future_data_lag2))])] <-0
future_data_lag2_new <-gather(future_data_lag2, "Month","FCST(lag2)", c(6:ncol(future_data_lag2)))
colnames(future_data_lag2_new)<-gsub(pattern='\\s+',replacement="", colnames(future_data_lag2_new))
future_data_lag2_new$ACTUAL<-0
future_data_lag2_mutate <- future_data_lag2_new[,c("Sku","Description","Franchise","Account","Month","ACTUAL","FCST(lag2)")]
future_data_lag2_group<-group_by(future_data_lag2_mutate,Sku,Description, Franchise, Account, Month)
future_data_lag2_summ <-summarise(future_data_lag2_group,ACTUAL=sum(ACTUAL), `FCST(lag2)`=sum(`FCST(lag2)`))

# combine #
fcst_data_beta <-merge(future_data_lag0_summ,future_data_lag1_summ,all = TRUE)
fcst_data <-merge(fcst_data_beta,future_data_lag2_summ,all = TRUE)
combine_data <- rbind(archive_data, fcst_data)
combine_data<-as.data.frame(combine_data)
combine_data[is.na(combine_data)]<-0

######################      SECTION 3       ########################
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

#for the input dateRange_variable3
test_df=as.Date(unique(financial_combine$Month))

################################### SECTION 4   ###########################################
##############################   getting fill rate data   #################################
## need to specify column types, otherwise generate some NAs
fill_rate_data_raw <- read_excel("FillRate raw data.xlsx", sheet="Sheet1",col_types=c("text","text","text","text","text","text","text","text","text","numeric","numeric","numeric","text"),skip=0)

fill_rate_data <-fill_rate_data_raw[,c('Item #',"Franchise","forecasted_account",'Date Range','item_description',"Order QTY","Fulfill QTY")]
fill_rate_data$`Item #`<-as.character(fill_rate_data$`Item #`)
fill_rate_data<-fill_rate_data[which(!is.na(fill_rate_data$`Item #`)),]

### combine all accounts
groupAll_fill_rate_data<-group_by(fill_rate_data,`Item #`,Franchise, `Date Range`)
sumAll_fill_rate_data <-summarise(groupAll_fill_rate_data,item_description=first(`item_description`),`Order QTY`=sum(`Order QTY`),`Fulfill QTY`=sum(`Fulfill QTY`))
sumAll_fill_rate_data$`forecasted_account`="B2B customers"
sumAll_fill_rate_data=sumAll_fill_rate_data[,c('Item #',"Franchise","forecasted_account",'Date Range','item_description',"Order QTY","Fulfill QTY")]
combine_FR_data <- rbind(as.data.frame(fill_rate_data),as.data.frame(sumAll_fill_rate_data))
#########################################################################

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

##### Build UI ########
ui <- navbarPage("Perricone Dashboard", theme = shinytheme("flatly"),
                 ##################################################################################
                 ######################### TESTING START ##########################################
                 ##################################################################################
                 tabPanel("Monthly Scorecard",
                          fluidPage(
                            sidebarLayout(
                              sidebarPanel(width =3,
                                           
                                           
                                           div(style="display: inline-block;",radioButtons("acct_variable1",h4(p(strong("Choose Account:"))),inline=TRUE,
                                                                                           choices=list("Company","Int'l"="APAC+LA","Amazon","Bloomingdales","Costco","Dillard's","EC Scott","GR"="Guthy Renker","L&T"="Lord & Taylor","Macy's","NM"="Neiman Marcus","Nordstrom","Other.com","PMD.com","QVC","Sephora","TSC","UK","ULTA"="Ulta","Zulily","NCO","Retail - Other")
                                                                                           
                                           )),
                                           hr(),
                                           h4(p(strong("Check by Franchise"))),
                                           actionLink("selectall","Select/Unselect All Franchises"), 
                                           div(style="display: inline-block;",checkboxGroupInput("franchise_variable1"," ",inline=TRUE,franchise_list,
                                                                                                 franchise_list[1:length(franchise_list)-1]
                                                                                                 
                                           )),
                                           
                                           
                                           h4(p(strong("Check by SKU"))),
                                           
                                           
                                           checkboxInput("checksku_variable1",label = p(strong("1). check to disregard Franchise Selection")), value=FALSE),
                                           uiOutput("sku_testing"),
                                           #uiOutput("franchise_drop"),
                                           selectInput("daterange_variable1",h4(p(strong("Date Range:"))),
                                                       date_ranges,
                                                       selected = date_ranges[1]
                                           )
                              ),
                              mainPanel(width = 9,
                                        h3("MTD Shipment Run Rate"),
                                        h5(IDC_update_date),
                                        br(),
                                        tabsetPanel(
                                          tabPanel("MTD runrate", fluidRow(splitLayout(cellWidths = c("36%", "64%"),highchartOutput("MTDpie", height = "480px"),highchartOutput("MTDchart", height = "480px")), br(),fluidRow(column(width=10,dataTableOutput("MTDtable",height='380px')))))
                                          #tabPanel("MTD runrate",fluidRow(highchartOutput("MTDchart",height = "480px")), br(),fluidRow(column(width=11,dataTableOutput("MTDtable",height='380px'))))
                                        ))),
                            hr(),
                            br(),
                            ######################################### starting from here for fill rate testing ############################################                
                            sidebarLayout(
                              sidebarPanel(width =3,
                                           h4(p(strong("Choose Account:"))),
                                           
                                           actionLink("selectall2","Select/Unselect All Accounts"), 
                                           div(style="display: inline-block;",checkboxGroupInput("acct_variable5"," ",inline=TRUE,B2B_list,
                                                                                                 c("Int'l"="APAC+LA",'Amazon','Costco',"Sephora","UK","ULTA"='Ulta','Retail-Other')        
                                           )),
                                           hr(),
                                           h4(p(strong("Check by Franchise"))),
                                           actionLink("selectall3","Select/Unselect All Franchises"), 
                                           div(style="display: inline-block;",checkboxGroupInput("franchise_variable5"," ",inline=TRUE,franchise_list,
                                                                                                 franchise_list[1:length(franchise_list)-1]
                                                                                                 
                                           )),
                                           
                                           
                                           h4(p(strong("Check by SKU"))),
                                           
                                           
                                           checkboxInput("checksku_variable3",label = p(strong("1). check to disregard Franchise Selection")), value=FALSE),
                                           uiOutput("sku_testing3"),
                                           #uiOutput("franchise_drop"),
                                           selectInput("daterange_variable4",h4(p(strong("Date Range:"))),
                                                       fill_date_ranges,
                                                       selected = fill_date_ranges[1]
                                           )
                              ),
                              mainPanel(width = 9,
                                        h3("Monthly Orders Fill Rate"),
                                        h5(FR_update_date),
                                        br(),
                                        tabsetPanel(
                                          ### bug reason --> cannot put same output under different ui layout
                                          
                                          tabPanel("Fill Rate",fluidRow(highchartOutput("fillrate_chart",height = "480px")), br(),fluidRow(column(width=11,dataTableOutput("fillrate_table",height='380px'))))
                                        ))),
                            hr(),
                            br(),
                            #######################################################################################################################################################
                            sidebarLayout(
                              sidebarPanel(width = 3,
                                           
                                           selectInput("acct_variable6",h4(p(strong("Choose Account:"))),
                                                       B2B_list,
                                                       selected = B2B_list[1]
                                           ),
                                           radioButtons("skucount_variable2",label=h5(p(strong("Choose SKU QTY:"))),
                                                        choices=list("Top 5"=5,"Top 10"=10,"Top 15"=15,"Top 20"=20),
                                                        selected = 5
                                           ),
                                           selectInput("daterange_variable5",h4(p(strong("Date Range:"))),
                                                       fill_date_ranges,
                                                       selected = fill_date_ranges[1]
                                           )
                              ),
                              
                              mainPanel(width=9,
                                        h3("Monthly Top Cuts"),
                                        br(),
                                        tabsetPanel(
                                          tabPanel("Top Cuts", dataTableOutput("top_cuts", height = "580px"))
                                          
                                        )
                              )
                            ),
                            hr(),
                            
                            
                            
                            br(),
                            hr(),
                            hr(),br()
                            
                            ######################################### ending from here for fill rate testing ############################################ 
                            
                          )),
                 
                 tabPanel("Sales Overview",
                          fluidPage( 
                            sidebarLayout(
                              sidebarPanel(width =2,
                                           
                                           div(style="display: inline-block;",radioButtons("acct_variable3",h4(p(strong("Choose Account:"))),inline=TRUE,
                                                                                           choices=acct_vector
                                                                                           
                                           )),
                                           hr(),
                                           
                                           div(style="display: inline-block;",radioButtons("franchise_variable3", h4(p(strong("Choose Franchise"))),inline=TRUE,
                                                                                           choices=franchise_vector
                                                                                           
                                           ))
                                           
                                           
                              ),
                              
                              mainPanel(width=10,
                                        h3("Costs v.s. Revenues"),
                                        h5(financial_report_date),
                                        br(),
                                        tabsetPanel(
                                          tabPanel("Profits by $", highchartOutput("profit", height = "580px"),fluidRow(column(width=10,dataTableOutput("profittable",height='380px')))),
                                          tabPanel("Margin by %",highchartOutput("margin", height = "580px"),fluidRow(column(width=10,dataTableOutput("margintable",height='380px'))))
                                        )
                              )
                            ),
                            hr(),
                            sidebarLayout(
                              sidebarPanel(width =2,
                                           div(style='display: inline-block;',dateRangeMonthsInput('daterange_variable3',label =h4(p(strong("Select Months"))),format = "mm/yyyy",start = test_df[length(test_df)-3], end=test_df[length(test_df)],min=head(test_df,1),max=tail(test_df,1),startview = "year",separator = " to ")),
                                           hr(),
                                           
                                           div(style="display: inline-block;",radioButtons("acct_variable4",h4(p(strong("Choose Account:"))),inline=TRUE,
                                                                                           choices=acct_vector
                                                                                           
                                           )),
                                           hr(),
                                           
                                           div(style="display: inline-block;",radioButtons("franchise_variable4", h4(p(strong("Choose Franchise"))),inline=TRUE,
                                                                                           choices=franchise_vector
                                                                                           
                                           ))
                                           
                                           
                              ),
                              
                              mainPanel(width=10,
                                        h3("Sell-in Data breakdown"),
                                        #h3(paste(format(input$daterange_variable3[1],'%b %y'),'to',format(input$daterange_variable3[2],"%b %y"),'Sell-in Data',sep=' ')),
                                        h5(financial_report_date),
                                        br(),
                                        tabsetPanel(
                                          tabPanel("By Dollars",  splitLayout(cellWidths = c("50%", "50%"),highchartOutput("acct_dollarpie", height = "380px"),highchartOutput("franchise_dollarpie", height = "380px"))),
                                          tabPanel("By Units", splitLayout(cellWidths = c("50%", "50%"),highchartOutput("acct_qtypie", height = "380px"),highchartOutput("franchise_qtypie", height = "380px")))
                                        ),
                                        br(),
                                        #h5(paste('Top 5 performers in',input$acct_variable4, input$franchise_variable4,sep=' ')),
                                        fluidRow(column(width=10,highchartOutput("topBars",height='280px'))),
                                        fluidRow(column(width=10,dataTableOutput("topTables",height='480px')))
                              )
                            ),
                            hr(),
                            
                            
                            
                            br(),
                            hr(),
                            hr(),br()
                            ### Newly Added ended ###
                          )),
                 ##################################################################################
                 ######################### TESTING END   ##########################################
                 ##################################################################################
                 
                 
                 tabPanel("Forecasting Performance",
                          fluidPage( 
                            
                            # br(),
                            sidebarLayout(
                              sidebarPanel(width =3,
                                           
                                           div(style="display: inline-block;",radioButtons("acct_variable2",h4(p(strong("Choose Account:"))),inline=TRUE,
                                                                                           choices=list("Company","Int'l"="APAC+LA","Amazon","Bloomingdales","Costco","Dillard's","EC Scott","GR"="Guthy Renker","L&T"="Lord & Taylor","Macy's","NM"="Neiman Marcus","Nordstrom","Other.com","PMD.com","QVC","Sephora","TSC","UK","ULTA"="Ulta","Zulily")
                                                                                           
                                           )),
                                           hr(),
                                           
                                           div(style="display: inline-block;",radioButtons("franchise_variable2", h4(p(strong("Choose Franchise"))),inline=TRUE,
                                                                                           choices=list("All Franchises","HP Classics"="High Potency Classics","Heritage","No Makeup"="No Makeup Skincare","No:Rinse","Intensive Pore","Pre:Empt","Masks","H2 EE"="H2 Elemental Energy","VC Ester"="Vitamin C Ester","Supplements","Neuropeptide","Cold Plasma","Essential Fx"="Essential Fx Acyl Glutathione","Re:Firm","Thio:Plex","Hypoallergenic","Mixed Franchise","Acne","Hypo CBD"="Hypoallergenic CBD")
                                                                                           
                                           )),
                                           
                                           
                                           h4(p(strong("Check SKU"))),
                                           
                                           checkboxInput("checksku_variable2",label = p(strong("1). dive into SKU level performance")), value=FALSE),
                                           
                                           uiOutput("sku_testing2"),
                                           
                                           downloadButton("downloadData2","Download")
                              ),
                              mainPanel(width = 9,
                                        h3("Actuals v.s. Forecasts"),
                                        h5(SOP_publish_date),
                                        br(),
                                        tabsetPanel(
                                          tabPanel("Performance Graph", highchartOutput("tschart", height = "580px")),
                                          tabPanel("Data Table",dataTableOutput("tstable", height = "580px")),
                                          tabPanel("BOM List",splitLayout(cellWidths = c("47%",'6%', "47%"),dataTableOutput("sku_contain", height = "380px"),dataTableOutput("blank", height = "380px"),dataTableOutput("sku_belong", height = "380px")))
                                        )
                                        
                              )),
                            hr(),
                            ### Newly Added ###
                            
                            br(),
                            sidebarLayout(
                              sidebarPanel(width = 3,
                                           radioButtons("daterange_variable2",label=h5(p(strong("Choose Rolling Months:"))),
                                                        choices=list("Next 4 months"=4,"Next 6 months"=6,"Next 12 months"=12)
                                           ),
                                           radioButtons("skucount_variable1",label=h5(p(strong("Choose SKU QTY:"))),
                                                        choices=list("Top 5"=5,"Top 10"=10,"Top 15"=15,"Top 20"=20),
                                                        selected = 10
                                           ),
                                           downloadButton("downloadData3","Download")
                              ),
                              
                              mainPanel(width=9,
                                        h3("MoM Top forecasting variances"),
                                        h5(lag_clarification_date),
                                        br(),
                                        tabsetPanel(
                                          tabPanel("Top SKUs Up", dataTableOutput("up10", height = "580px")),
                                          tabPanel("Top SKUs Down",dataTableOutput("down10", height = "580px"))
                                        )
                              )
                            )
                            
                            
                            
                            
                            
                          ))
                
                 
                 
                 
)



###### Build Server #####
server <- function(input, output, session){
  
  ### testing to include most variance SKUs, for Supply Planning purpose ####
  
  observe({
    if(input$selectall == 0) return(NULL) 
    else if (input$selectall%%2 == 0)
    {
      updateCheckboxGroupInput(session,"franchise_variable1",choices=franchise_list, inline=TRUE)
    }
    else
    {
      updateCheckboxGroupInput(session,"franchise_variable1",choices=franchise_list,selected=franchise_list, inline=TRUE)
    }
  })
  
  
  output$sku_testing <- renderUI({
    
    if((input$checksku_variable1)==FALSE)
      return()
    else
      
    {selectInput("sku_variable1",paste("      ","2). Type or select SKU:",sep="   "),
                 choices=SKU_list,
                 selected=NULL, multiple=FALSE)
      #updateCheckboxGroupInput(session,"franchise_variable1","Choose Franchise(s):",choices=franchise_list)
      #input$franchise_variable1==FALSE
      
    }
  })
  
  observe({
    if(!input$checksku_variable1) return()
    
    else
    {
      updateCheckboxGroupInput(session,"franchise_variable1",choices=franchise_list, inline=TRUE)
    }
  })
  output$sku_testing2 <- renderUI({
    if(!input$checksku_variable2)
      return()
    else
      selectInput("sku_variable2",paste("      ","2). Type or select SKU:",sep="   "),
                  choices=sku_testing2_function(),
                  selected=NULL, multiple=FALSE)
  })
  
  
  output$MTDchart <-renderHighchart(
    highchart()%>%
      
      hc_title(text=paste("Rolling shipment for",input$acct_variable1,sep=' '))%>%
      #{if(!input$checksku_variable1) hc_title(text="Shipment Running Log (Franchises)") else .}
      hc_subtitle(text=if(!input$checksku_variable1) paste("Selected",length(input$franchise_variable1),"Franchise(s)",sep=' ') else paste(input$sku_variable1,description_function()$description[1],sep=" "))%>%
      
      #hc_title(text=paste("Shipment Running Log:", input$sku_variable1,sep = " "))%>%
      hc_tooltip(crosshairs=TRUE, valueDecimals = 2, borderWidth = 5, sort = TRUE, table = TRUE)%>%
      hc_xAxis(list(categories=MTD_group_function()$`Rolling Dates`))%>%
      hc_yAxis_multiples(
        
        list(title=list(text="Shipped UNITS")),
        list(title=list(text="MTD runrate"),opposite=TRUE)
        
      ) %>% 
      hc_plotOptions(column=list(color="#EAC46D",
                                 dataLabels=list(enabled=FALSE),
                                 stacking="normal",
                                 enableMouseTracking = TRUE
      ))%>%
      hc_plotOptions(area=list(color="#B0C4DE",
                               dataLabels=list(enabled=FALSE),
                               stacking="normal",
                               enableMouseTracking = TRUE
      ))%>%
      
      hc_plotOptions(spline=list(color="#F4A460",
                                 dataLabels=list(enabled=FALSE),
                                 stacking="normal",
                                 enableMouseTracking = TRUE
      ))%>%
      
      hc_add_series(data = MTD_group_function()$`FCST line`, type = "area",  name="Monthly Forecast Line",yAxis=0) %>%
      hc_add_series(data = MTD_group_function()$`MTD ship`, type = "column", name="MTD shipment",yAxis=0) %>%
      hc_add_series(data = MTD_group_function()$runrate,type='spline', name="MTD runrate",yAxis=1)%>%
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = TRUE) %>%
      hc_exporting(enabled = TRUE,fallbackToExportServer = F,
                   filename = "MTD-runrate")%>%
      #hc_legend(enabled = TRUE)%>%
      hc_add_theme(hc_theme_google()) 
  )
  
  output$MTDpie <-renderHighchart({
    highchart() %>% 
      hc_chart(type = "pie") %>% 
      hc_add_series_labels_values(labels = MTD_pie_function()$Account, values = MTD_pie_function()$QTY)%>% 
      hc_tooltip(pointFormat = paste('{point.y} <br/><b>{point.percentage:.0f}%</b>')) %>%
      hc_title(text=paste(MTD_pie_function()$MTD[1],"Shipment breakdown",sep=' '))%>%
      hc_subtitle(text=if(!input$checksku_variable1) paste("Selected",length(input$franchise_variable1),"Franchise(s)",sep=' ') else input$sku_variable1)
    })  
  
  output$MTDtable <-DT::renderDataTable(
    # DT::datatable(MTD_group_function(),options = list(searching=FALSE),rownames= FALSE)%>%
    
    DT::datatable(cbind(MTD_group_function(),`SKU/Franchise`= if(!(input$checksku_variable1)) paste("Selected",length(input$franchise_variable1),"Franchise(s)",sep=' ') else input$sku_variable1,Account=input$acct_variable1),options = list( columnDefs = list(list(className = 'dt-center', targets = 0:4))),rownames= FALSE)%>%
      
      formatRound(c("MTD ship","FCST line"), digits = 0, interval = 3, mark = ",", 
                  dec.mark = getOption("OutDec"))%>%
      formatPercentage(c("runrate"),0)
  )
  
  output$profittable <- DT::renderDataTable(
    DT::datatable(convert_function()[c(4,5,1,3,6),],options = list( columnDefs = list(list(className = 'dt-center', targets = 0:4))))
    
    
  )
  
  output$margintable <- DT::renderDataTable(
    DT::datatable(convert_function()[c(4,5,1,7),],options = list( columnDefs = list(list(className = 'dt-center', targets = 0:4))))
    
  )
  
  output$up10 <- DT::renderDataTable(
    DT::datatable(up10_function(),options = list( columnDefs = list(list(className = 'dt-center', targets = 0:4))))%>%
      formatRound(c(3,4,5), digits = 0, interval = 3, mark = ",", 
                  dec.mark = getOption("OutDec"))
  )
  
  output$down10 <- DT::renderDataTable(
    DT::datatable(down10_function(),options = list( columnDefs = list(list(className = 'dt-center', targets = 0:4))))%>%
      formatRound(c(3,4,5), digits = 0, interval = 3, mark = ",", 
                  dec.mark = getOption("OutDec"))
  )
  
  up10_function <-reactive({
    combine_data2 <-combine_data
    combine_data2$Month<-parse_date_time(combine_data2$Month,"my")
    basic2_group <-group_by(combine_data2,Month,Sku)
    basic2_summ<-summarise(basic2_group, Description=first(`Description`),ACTUAL=sum(`ACTUAL`),`FCST(lag0)`=sum(`FCST(lag0)`),`FCST(lag1)`=sum(`FCST(lag1)`),`FCST(lag2)`=sum(`FCST(lag2)`))
    topsku <- basic2_summ[which(basic2_summ$Month>Sys.Date()&basic2_summ$Month<Sys.Date()+as.numeric(input$daterange_variable2)*30),]
    topsku_group <-group_by(topsku,Sku)
    topsku_summ <-summarise(topsku_group, Description=first(`Description`),ACTUAL=sum(`ACTUAL`),`FCST(lag0)`=sum(`FCST(lag0)`),`FCST(lag1)`=sum(`FCST(lag1)`),`FCST(lag2)`=sum(`FCST(lag2)`))
    topsku_summ$variance <-topsku_summ$`FCST(lag0)`-topsku_summ$`FCST(lag1)`
    topsku_summ$abs_variance <-abs(topsku_summ$variance)
    topsku_summ<-topsku_summ[order(-topsku_summ$variance),]
    top_10_jump <- head(topsku_summ,as.numeric(input$skucount_variable1))
    
    return_top_list=c()
    for (i in top_10_jump$Sku){
      getacct_data_group<-group_by(combine_data2[which(combine_data2$Sku==i&combine_data2$Month>Sys.Date()&combine_data2$Month<Sys.Date()+as.numeric(input$daterange_variable2)*30),],Account)
      getacct_data_summ<-summarise(getacct_data_group,`FCST(lag0)`=sum(`FCST(lag0)`),`FCST(lag1)`=sum(`FCST(lag1)`))
      getacct_data_summ$abs_variance <-abs(getacct_data_summ$`FCST(lag0)`- getacct_data_summ$`FCST(lag1)`)
      getacct_data_summ<-getacct_data_summ[order(-getacct_data_summ$abs_variance),]
      return_value<-getacct_data_summ$Account[1]
      return_top_list[i]=return_value
    }
    top_10_jump$`driven by` <-return_top_list
    jump_10_present <-top_10_jump[,c(1,2,4,5,7,9)]
    return(jump_10_present)
    
  })
  
  down10_function <-reactive({
    combine_data2 <-combine_data
    combine_data2$Month<-parse_date_time(combine_data2$Month,"my")
    basic2_group <-group_by(combine_data2,Month,Sku)
    basic2_summ<-summarise(basic2_group,Description=first(`Description`), ACTUAL=sum(`ACTUAL`),`FCST(lag0)`=sum(`FCST(lag0)`),`FCST(lag1)`=sum(`FCST(lag1)`),`FCST(lag2)`=sum(`FCST(lag2)`))
    topsku <- basic2_summ[which(basic2_summ$Month>Sys.Date()&basic2_summ$Month<Sys.Date()+as.numeric(input$daterange_variable2)*30),]
    topsku_group <-group_by(topsku,Sku)
    topsku_summ <-summarise(topsku_group,Description=first(`Description`), ACTUAL=sum(`ACTUAL`),`FCST(lag0)`=sum(`FCST(lag0)`),`FCST(lag1)`=sum(`FCST(lag1)`),`FCST(lag2)`=sum(`FCST(lag2)`))
    topsku_summ$variance <-topsku_summ$`FCST(lag0)`-topsku_summ$`FCST(lag1)`
    topsku_summ$abs_variance <-abs(topsku_summ$variance)
    topsku_summ<-topsku_summ[order(-topsku_summ$variance),]
    top_10_drop <- tail(topsku_summ,as.numeric(input$skucount_variable1))
    top_10_drop <-top_10_drop[order(top_10_drop$variance),]
    
    return_bottom_list=c()
    for (i in top_10_drop$Sku){
      getacct_data_group<-group_by(combine_data2[which(combine_data2$Sku==i&combine_data2$Month>Sys.Date()&combine_data2$Month<Sys.Date()+as.numeric(input$daterange_variable2)*30),],Account)
      getacct_data_summ<-summarise(getacct_data_group,`FCST(lag0)`=sum(`FCST(lag0)`),`FCST(lag1)`=sum(`FCST(lag1)`))
      getacct_data_summ$abs_variance <-abs(getacct_data_summ$`FCST(lag0)`- getacct_data_summ$`FCST(lag1)`)
      getacct_data_summ<-getacct_data_summ[order(-getacct_data_summ$abs_variance),]
      return_value<-getacct_data_summ$Account[1]
      return_bottom_list[i]=return_value
    }
    top_10_drop$`driven by`<-return_bottom_list
    drop_10_present <-top_10_drop[,c(1,2,4,5,7,9)]
    return(drop_10_present)
  })
  
  output$tschart<-renderHighchart(
    highchart(type = "stock") %>%
      hc_title(text=paste("forecasting time series performance",input$acct_variable2,sep=': '))%>%
      hc_subtitle(text=if(!input$checksku_variable2) input$franchise_variable2 else paste(input$sku_variable2,combine_data[which(combine_data$Sku==input$sku_variable2),2][1],sep=" "))%>%
      #hc_tooltip(valueDecimals = 0, sort = FALSE) %>%
      hc_tooltip(valueDecimals = 0, borderWidth = 3, sort = FALSE, table = TRUE)%>%
      hc_yAxis_multiples(
        list(opposite = FALSE),
        list()
      ) %>%  
      hc_add_series_times_values(ts_function()$Month, ts_function()$`FCST(lag0)`, name = "S&OP Forecasts (lag 0)",color="#3498DB") %>%
      hc_add_series_times_values(ts_function()$Month, ts_function()$`FCST(lag1)`, name = "S&OP Forecasts (lag 1)",color="#85C1E9") %>%
      hc_add_series_times_values(ts_function()$Month, ts_function()$`FCST(lag2)`, name = "S&OP Forecasts (lag 2)",color="#AED6F1") %>%
      hc_add_series_times_values(ts_function()$Month, ts_function()$ACTUAL, name = "S&OP Actuals",color="#DC7633") %>%
      
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = TRUE) %>%
      hc_rangeSelector(selected = 8) %>%
      hc_exporting(enabled = TRUE,fallbackToExportServer = F,
                   filename = "forecasting-performance-chart")%>%
      hc_add_theme(hc_theme_gridlight())
  )
  
  output$tstable<-DT::renderDataTable(
    DT::datatable(tstable_function(),options = list(columnDefs = list(list(className = 'dt-center', targets = 0:4))),rownames= FALSE)%>%
      #DT::datatable(cbind(ts_function(),SKU= if(is.null(input$sku_variable2))" " else input$sku_variable2,Account=input$acct_variable2),options = list(searching=FALSE),rownames= FALSE)%>% 
      formatRound(c("ACTUAL","FCST(lag0)","FCST(lag1)","FCST(lag2)"), digits = 0, interval = 3, mark = ",", 
                  dec.mark = getOption("OutDec"))
  )
  
  
  ########################################################################################
  ########################################################################################
  ########################################################################################
  
  output$acct_dollarpie<-renderHighchart({
    highchart() %>% 
      hc_chart(type = "pie") %>% 
      hc_add_series_labels_values(labels = acctpie_ranking_dollar()$Account, values = acctpie_ranking_dollar()$Revenues)%>% 
      hc_tooltip(pointFormat = paste('${point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
      hc_exporting(enabled = TRUE,fallbackToExportServer = F,
                   filename = "account-dollar-pie-chart")%>%
      hc_title(text = paste(format(input$daterange_variable3[1],'%b %y'),'to',format(input$daterange_variable3[2],"%b %y"),input$acct_variable4,'$ breakdown',sep=' '))
    #paste(format(input$daterange_variable3[1],'%b %y'),'to',format(input$daterange_variable3[2],"%b %y"),'Sell-in Data',sep=' ')
  })
  
  output$acct_qtypie<-renderHighchart({
    highchart() %>% 
      hc_chart(type = "pie") %>% 
      hc_add_series_labels_values(labels = acctpie_ranking_qty()$Account, values = acctpie_ranking_qty()$QTY)%>% 
      hc_tooltip(pointFormat = paste('{point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
      hc_exporting(enabled = TRUE,fallbackToExportServer = F,
                   filename = "account-qty-pie-chart")%>%
      hc_title(text = paste(format(input$daterange_variable3[1],'%b %y'),'to',format(input$daterange_variable3[2],"%b %y"),input$acct_variable4,'QTY breakdown',sep=' '))
  })
  
  output$franchise_dollarpie<-renderHighchart({
    highchart() %>% 
      hc_chart(type = "pie") %>% 
      hc_add_series_labels_values(labels = franpie_ranking_dollar()$Franchise, values = franpie_ranking_dollar()$Revenues)%>% 
      hc_tooltip(pointFormat = paste('${point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
      hc_exporting(enabled = TRUE,fallbackToExportServer = F,
                   filename = "franchise-dollar-pie-chart")%>%
      hc_title(text = paste(format(input$daterange_variable3[1],'%b %y'),'to',format(input$daterange_variable3[2],"%b %y"),input$franchise_variable4,'$ breakdown',sep=' '))
  })
  
  output$franchise_qtypie <-renderHighchart({
    highchart() %>% 
      hc_chart(type = "pie") %>% 
      hc_add_series_labels_values(labels = franpie_ranking_qty()$Franchise, values = franpie_ranking_qty()$QTY)%>% 
      hc_tooltip(pointFormat = paste('{point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
      hc_exporting(enabled = TRUE,fallbackToExportServer = F,
                   filename = "franchise-qty-pie-chart")%>%
      hc_title(text = paste(format(input$daterange_variable3[1],'%b %y'),'to',format(input$daterange_variable3[2],"%b %y"),input$franchise_variable4,'QTY breakdown',sep=' '))
  })
  
  ##############################################################################################
  ##############################################################################################
  
  output$topBars<-renderHighchart({
    
    
    highchart()%>%
      
      hc_title(text=paste('Top performers within',input$franchise_variable4, ',', input$acct_variable4,sep=' '))%>%
      #hc_xAxis(list(categories = ComparisonFunction()$Combine)) %>%
      hc_yAxis(title=list(text="SKU Revenuues"))%>%
      
      hc_xAxis(list(categories = if(nrow(skubar_ranking_function())<5) skubar_ranking_function()$Sku else skubar_ranking_function()[1:5,]$Sku)) %>%
      
      hc_tooltip(valueDecimals = 2, valuePrefix = "$")%>%
      
      hc_add_series(data =if(nrow(skubar_ranking_function())<5) skubar_ranking_function()$`Revenues` else skubar_ranking_function()[1:5,]$`Revenues`, type = "bar",name="Selected Months' Revenues") %>%
      
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = FALSE) %>%
      hc_exporting(enabled = TRUE,fallbackToExportServer = F,
                   filename = "top-performers-bars")%>%
      hc_legend(enabled = TRUE)
  })
  
  output$topTables<-DT::renderDataTable(
    DT::datatable(cbind(skubar_ranking_function(),`Account`=input$acct_variable4,`Date Range`=paste(format(input$daterange_variable3[1],'%b %y'),format(input$daterange_variable3[2],"%b %y"),sep='-')),options = list( columnDefs = list(list(className = 'dt-center', targets = 0:4)),pageLength=5),rownames= FALSE)%>%
      
      
      formatRound(c("QTY"), digits = 0, interval = 3, mark = ",", 
                  dec.mark = getOption("OutDec"))%>%
      formatCurrency(c("Revenues"),currency = '$',digits = 0,mark = ',',interval = 3)
    
  )
  
  
  ########################################################################################
  ########################################################################################format(as.Date(financial_summ$Month),"%b %y")
  ########################################################################################
  
  
  
  
  
  
  
  tstable_function <-reactive({
    
    #if(nrow(ts_function())==0) return()
    #tstable_df<-ts_function()
    validate(
      need(nrow(ts_function())!=0, "NO S&OP DATA"
      ))
    
    tstable_df<-ts_function()
    if(!(input$checksku_variable2)){
      tstable_df$`SKU/Franchise`=input$franchise_variable2 
      tstable_df$Account=input$acct_variable2
      
    }
    else{
      
      tstable_df$`SKU/Franchise`=input$sku_variable2 
      tstable_df$Account=input$acct_variable2
    }
    tstable_df$Month <- format(as.Date(tstable_df$Month),"%b %y")
    return(tstable_df)
  })
  
  ts_function <-reactive({
    
    if(input$checksku_variable2){
      if(input$acct_variable2=='Company'){
        
        basic_group <-group_by(combine_data,Month,Sku)
        basic_summ<-summarise(basic_group, ACTUAL=sum(`ACTUAL`),`FCST(lag0)`=sum(`FCST(lag0)`),`FCST(lag1)`=sum(`FCST(lag1)`),`FCST(lag2)`=sum(`FCST(lag2)`))
        basic_summ <-basic_summ[which(basic_summ$Sku==input$sku_variable2),c(1,3,4:ncol(basic_summ))]
      }
      else{
        basic_group <-group_by(combine_data,Month,Account,Sku)
        basic_summ<-summarise(basic_group, ACTUAL=sum(`ACTUAL`),`FCST(lag0)`=sum(`FCST(lag0)`),`FCST(lag1)`=sum(`FCST(lag1)`),`FCST(lag2)`=sum(`FCST(lag2)`))
        basic_summ <-basic_summ[which(basic_summ$Sku==input$sku_variable2&basic_summ$Account==input$acct_variable2),c(1,4,5:ncol(basic_summ))]
      }
    }
    else {
      if(input$acct_variable2=='Company'&&input$franchise_variable2=='All Franchises'){
        basic_group <-group_by(combine_data,Month)
        basic_summ<-summarise(basic_group, ACTUAL=sum(`ACTUAL`),`FCST(lag0)`=sum(`FCST(lag0)`),`FCST(lag1)`=sum(`FCST(lag1)`),`FCST(lag2)`=sum(`FCST(lag2)`))
      }
      else if(input$acct_variable2!='Company'&&input$franchise_variable2=='All Franchises'){
        basic_group <-group_by(combine_data,Month,Account)
        basic_summ<-summarise(basic_group, ACTUAL=sum(`ACTUAL`),`FCST(lag0)`=sum(`FCST(lag0)`),`FCST(lag1)`=sum(`FCST(lag1)`),`FCST(lag2)`=sum(`FCST(lag2)`))
        basic_summ <-basic_summ[which(basic_summ$Account==input$acct_variable2),c(1,3,4:ncol(basic_summ))]
      }
      else if(input$acct_variable2=='Company'&&input$franchise_variable2!='All Franchises'){
        basic_group <-group_by(combine_data,Month,Franchise)
        basic_summ<-summarise(basic_group, ACTUAL=sum(`ACTUAL`),`FCST(lag0)`=sum(`FCST(lag0)`),`FCST(lag1)`=sum(`FCST(lag1)`),`FCST(lag2)`=sum(`FCST(lag2)`))
        basic_summ <-basic_summ[which(basic_summ$Franchise==input$franchise_variable2),c(1,3,4:ncol(basic_summ))]
      }
      else {
        basic_group <-group_by(combine_data,Month,Account,Franchise)
        basic_summ<-summarise(basic_group, ACTUAL=sum(`ACTUAL`),`FCST(lag0)`=sum(`FCST(lag0)`),`FCST(lag1)`=sum(`FCST(lag1)`),`FCST(lag2)`=sum(`FCST(lag2)`))
        basic_summ <-basic_summ[which(basic_summ$Franchise==input$franchise_variable2&basic_summ$Account==input$acct_variable2),c(1,4,5:ncol(basic_summ))]
      }
    }
    
    basic_summ$Month<-parse_date_time(basic_summ$Month,"my")
    basic_summ<-basic_summ[order(basic_summ$Month),]
    
    return(basic_summ)
    
  })
  #### STARTING FROM HERE NEW
  ########################################
  output$profit<- renderHighchart(
    highchart()%>%
      
      hc_title(text=paste(input$acct_variable3, "'s Average COGS,ASP,QTY, Revenues & Profits", sep = ''))%>%
      
      hc_subtitle(text=input$franchise_variable3)%>%
      
      hc_tooltip(crosshairs=TRUE, valueDecimals = 1, borderWidth = 3, sort = FALSE, table = TRUE)%>%
      hc_xAxis(list(categories=CR_function()$Month))%>%
      hc_yAxis_multiples(
        
        list(title=list(text="Costs v.s. Revenues")),
        list(title=list(text="ASP v.s. Average COGS"),opposite=TRUE)
        
      ) %>% 
      hc_plotOptions(column=list(color="#EAC46D",
                                 dataLabels=list(enabled=FALSE),
                                 
                                 enableMouseTracking = TRUE
      ))%>%
      
      
      hc_plotOptions(spline=list(color="#F4A460",
                                 dataLabels=list(enabled=FALSE),
                                 stacking="normal",
                                 enableMouseTracking = TRUE
      ))%>%
      
      hc_add_series(data = CR_function()$ASP, type = "spline",  name="ASP",yAxis=1,color="#92C8E0") %>%
      hc_add_series(data = CR_function()$`Average COGS`, type = "spline",  name="Average COGS",yAxis=1,color="#E2A5AD") %>%
      
      hc_add_series(data = CR_function()$Costs, type = "column", name="Monthly Costs",yAxis=0,color="#546C8C") %>%
      hc_add_series(data = CR_function()$Revenues,type='column', name="Monthly Revenues",yAxis=0,color="#194568")%>%
      hc_add_series(data = CR_function()$profit,type='column', name="Monthly Profits",yAxis=0,color="#A19E50")%>%
      
      
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = FALSE) %>%
      hc_legend(enabled = TRUE)%>%
      hc_exporting(enabled = TRUE,fallbackToExportServer = F,
                   filename = "sales-chart-profit")%>%
      hc_add_theme(hc_theme_gridlight())
    #hc_add_theme(hc_theme_google()) 
  )
  
  output$margin<-renderHighchart(
    highchart()%>%
      
      hc_title(text=paste(input$acct_variable3, "'s Average COGS,ASP,QTY, & Margins", sep = ''))%>%
      
      hc_subtitle(text=input$franchise_variable3)%>%
      
      hc_tooltip(crosshairs=TRUE, valueDecimals = 1, borderWidth = 3, sort = FALSE, table = TRUE)%>%
      hc_xAxis(list(categories=CR_function()$Month))%>%
      hc_yAxis_multiples(
        
        list(title=list(text="Margins over months")),
        list(title=list(text="ASP v.s. Average COGS"),opposite=TRUE)
        
      ) %>% 
      hc_plotOptions(column=list(color="#EAC46D",
                                 dataLabels=list(enabled=FALSE),
                                 stacking="normal",
                                 enableMouseTracking = TRUE
      ))%>%
      
      
      hc_plotOptions(spline=list(color="#F4A460",
                                 dataLabels=list(enabled=FALSE),
                                 stacking="normal",
                                 enableMouseTracking = TRUE
      ))%>%
      
      hc_add_series(data = CR_function()$ASP, type = "spline",  name="ASP",yAxis=1,color="#92C8E0") %>%
      hc_add_series(data = CR_function()$`Average COGS`, type = "spline",  name="Average COGS",yAxis=1,color="#E2A5AD") %>%
      hc_add_series(data = CR_function()$margin,type='column', name="margin ratio",yAxis=0,color="#A19E50")%>%
      hc_add_series(data = CR_function()$cost_rate, type = "column",name='cost to revenue ratio',yAxis=0,color="#546C8C") %>%
      #hc_add_series(data = CR_function()$Revenues,type='column', name="Monthly Revenues",yAxis=0,color="#194568")%>%
      
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = FALSE) %>%
      hc_legend(enabled = TRUE)%>%
      hc_exporting(enabled = TRUE,fallbackToExportServer = F,
                   filename = "sales-chart-margin")%>%
      hc_add_theme(hc_theme_gridlight())
    #hc_add_theme(hc_theme_google()) 
  )
  ########################################################
  ## CR stands for cost and revenue
  CR_function<-reactive({
    if(input$acct_variable3=='Company'&&input$franchise_variable3=='All Franchises'){
      financial_group <-group_by(financial_combine,Month)
      financial_summ <-summarise(financial_group,QTY=sum(QTY),Costs=sum(`ExtendedCost`),Revenues=sum(`ExtendedRevenue`))
      
    }
    else if(input$acct_variable3!='Company'&&input$franchise_variable3=='All Franchises'){
      financial_group <-group_by(financial_combine,Month,Account)
      financial_summ <-summarise(financial_group,QTY=sum(QTY),Costs=sum(`ExtendedCost`),Revenues=sum(`ExtendedRevenue`))
      financial_summ <-financial_summ[which(financial_summ$Account==input$acct_variable3),-c(2)]
    }
    
    else if(input$acct_variable3=='Company'&&input$franchise_variable3!='All Franchises'){
      financial_group <-group_by(financial_combine,Month,Franchise)
      financial_summ <-summarise(financial_group,QTY=sum(QTY),Costs=sum(`ExtendedCost`),Revenues=sum(`ExtendedRevenue`))
      financial_summ <-financial_summ[which(financial_summ$Franchise==input$franchise_variable3),-c(2)]
    }
    
    else {
      financial_group <-group_by(financial_combine,Month,Account,Franchise)
      financial_summ <-summarise(financial_group,QTY=sum(QTY),Costs=sum(`ExtendedCost`),Revenues=sum(`ExtendedRevenue`))
      financial_summ <-financial_summ[which(financial_summ$Franchise==input$franchise_variable3&financial_summ$Account==input$acct_variable3),-c(2,3)]
    }
    financial_summ$ASP <-financial_summ$Revenues/financial_summ$QTY
    financial_summ$`Average COGS` <-financial_summ$Costs/financial_summ$QTY
    financial_summ$profit <-financial_summ$Revenues-financial_summ$Costs
    financial_summ$margin <-financial_summ$profit/financial_summ$Revenues
    financial_summ$cost_rate <-financial_summ$Costs/financial_summ$Revenues
    
    financial_summ$Costs<-currency(financial_summ$Costs,'$ ',digits = 0)
    financial_summ$Revenues<-currency(financial_summ$Revenues,'$ ',digits = 0)
    financial_summ$ASP<-currency(financial_summ$ASP,'$ ',digits = 1)
    financial_summ$`Average COGS`<-currency(financial_summ$`Average COGS`,'$ ',digits = 1)
    financial_summ$profit<-currency(financial_summ$profit,'$ ',digits = 0)
    financial_summ$margin<-percent(financial_summ$margin)
    financial_summ$cost_rate<-percent(financial_summ$cost_rate)
    financial_summ$QTY<-accounting(financial_summ$QTY,digits = 0, big.mark = ',')
    
    #financial_summ$Month<-parse_date_time(financial_summ$Month,"my")
    financial_summ<-financial_summ[order(financial_summ$Month),]
    financial_summ$Month <- format(as.Date(financial_summ$Month),"%b %y")
    
    return(financial_summ)
    
  })
  
  # show case under graphs
  convert_function <-reactive({
    CR_df <-as.data.frame(CR_function())
    convert_df <-as.data.frame(t(CR_df))
    colnames(convert_df) <-CR_df$Month
    convert_ready<-convert_df[-1,]
    return(convert_ready)
  })
  
  
  
  
  sku_testing2_function <-reactive({
    if(input$franchise_variable2=='All Franchises'){sku_list2<-SKU_list}
    else {sku_list2<-unique(combine_data[which(combine_data$Franchise==input$franchise_variable2),1])}
    return(sku_list2)
  })
  
  
  
  
  output$downloadData2 <-downloadHandler(
    ###### this function only works in the browser window !!!!!!!!! #############
    filename = function() {
      paste(input$acct_variable2, "Actual v.s. Forecast performance.csv", sep = " ")
    },
    
    
    content = function(file) {
      write.csv(tstable_function(),file,row.names = FALSE)
    }
  )
  
  
  output$downloadData3 <-downloadHandler(
    ###### this function only works in the browser window !!!!!!!!! #############
    filename = function() {
      paste("Top",input$skucount_variable1,"SKUs",input$daterange_variable2, "MoM forecasting variances.csv", sep = " ")
    },
    
    
    content = function(file) {
      write.csv(as.data.frame(rbind(up10_function(),down10_function())),file,row.names = FALSE)
    }
  )
  
  
  MTD_group_function <-reactive({
    ship_runrate <- read_excel('ship runrate raw data.xlsx', sheet = input$daterange_variable1, col_types=c("text","text","text","text","text","numeric","numeric"),skip=0)
    #runrate_group<-group_by(ship_runrate,rolling_index,forecasted_account
    
    
    
    if(input$acct_variable1=='Company')
    {
      if(input$checksku_variable1){
        checking_sku <-group_by(ship_runrate,`rolling_index`,`Item #`)
        group_summ<-summarise(checking_sku,description=first(item_description),shipTotal=sum(`Ship Qty`),fcstTotal=sum(`Monthly Forecast`))
        group_DF <-group_summ[which(group_summ$`Item #`==input$sku_variable1),]
      }
      
      else{
        testing_group <-group_by(ship_runrate,`rolling_index`,`Franchise`)
        group_summ <-summarise(testing_group,shipTotal=sum(`Ship Qty`),fcstTotal=sum(`Monthly Forecast`))
        group_DF <- group_summ[which(group_summ$Franchise%in%input$franchise_variable1),]
        
      }
    }
    
    
    
    else
    {
      if(input$checksku_variable1){
        checking_sku <-group_by(ship_runrate,`rolling_index`,`forecasted_account`,`Item #`)
        group_summ<-summarise(checking_sku,description=first(item_description),shipTotal=sum(`Ship Qty`),fcstTotal=sum(`Monthly Forecast`))
        group_DF <-group_summ[which(group_summ$`Item #`==input$sku_variable1&group_summ$`forecasted_account`==input$acct_variable1),]
        
        
      }
      else{
        testing_group <-group_by(ship_runrate,`rolling_index`,`forecasted_account`,`Franchise`)
        group_summ <-summarise(testing_group,shipTotal=sum(`Ship Qty`),fcstTotal=sum(`Monthly Forecast`))
        group_DF <- group_summ[which(group_summ$`forecasted_account`==input$acct_variable1&group_summ$Franchise%in%input$franchise_variable1),]
      }  
      
    }
    
    return_DF<-group_by(group_DF,`rolling_index`)
    return_DF<-summarise(return_DF,`MTD ship`=sum(shipTotal),`FCST line`=sum(fcstTotal))
    return_DF$runrate=return_DF$`MTD ship`/return_DF$`FCST line`
    names(return_DF)[names(return_DF) == "rolling_index"] <- "Rolling Dates"
    validate(
      need(nrow(return_DF)!=0, "PLEASE INPUT PROPER SKU/FRANCHISE(S)")
    )
    return(return_DF)
    
  })
  
  
  
  description_function <-reactive({
    ship_runrate <- read_excel('ship runrate raw data.xlsx', sheet = input$daterange_variable1, col_types=c("text","text","text","text","text","numeric","numeric"),skip=0)
    #runrate_group<-group_by(ship_runrate,rolling_index,forecasted_account
    
    
    
    if(input$acct_variable1=='Company')
    {
      
      checking_sku <-group_by(ship_runrate,`rolling_index`,`Item #`)
      group_summ<-summarise(checking_sku,description=first(item_description),shipTotal=sum(`Ship Qty`),fcstTotal=sum(`Monthly Forecast`))
      group_DF <-group_summ[which(group_summ$`Item #`==input$sku_variable1),]
      if(input$checksku_variable1)
        group_DF=group_DF
      
      else
        group_DF=group_DF[0,]
    }
    
    
    
    else
    {
      checking_sku <-group_by(ship_runrate,`rolling_index`,`forecasted_account`,`Item #`)
      group_summ<-summarise(checking_sku,description=first(item_description),shipTotal=sum(`Ship Qty`),fcstTotal=sum(`Monthly Forecast`))
      group_DF <-group_summ[which(group_summ$`Item #`==input$sku_variable1&group_summ$`forecasted_account`==input$acct_variable1),]
      if(input$checksku_variable1)
        group_DF=group_DF
      
      else
        group_DF=group_DF[0,]
      
    }
    
    
    
    return_DF<-group_by(group_DF,`rolling_index`)
    return_DF<-summarise(return_DF,description=first(description),`MTD ship`=sum(shipTotal),`FCST line`=sum(fcstTotal))
    names(return_DF)[names(return_DF) == "rolling_index"] <- "Rolling Dates"
    return(return_DF)
    
  })
  
  
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
    financial_summ1<-group_by(financial_summ1,Account)
    financial_summ1<-summarise(financial_summ1,QTY=sum(`QTY`),Revenues=sum(`Revenues`))
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
    financial_summ2<-group_by(financial_summ2,Franchise)
    financial_summ2<-summarise(financial_summ2,QTY=sum(`QTY`),Revenues=sum(`Revenues`))
    #financial_summ2 <-financial_summ2[which(financial_summ2$Month>=as.Date('2019-10-01')&financial_summ2$Month<=as.Date('2020-01-01')),]
    
    return(financial_summ2)
    
  })
  
  
  
  
  skubar_function <-reactive({
    if(input$acct_variable4=='Company'&&input$franchise_variable4=='All Franchises'){
      financial_group <-group_by(financial_combine,Sku,Month)
      financial_summ3 <-summarise(financial_group,Description=first(`Description`),Franchise=first(`Franchise`),QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
      
    }
    else if(input$acct_variable4!='Company'&&input$franchise_variable4=='All Franchises'){
      financial_group <-group_by(financial_combine,Sku,Month,Account)
      financial_summ3 <-summarise(financial_group,Description=first(`Description`),Franchise=first(`Franchise`),QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
      financial_summ3 <-financial_summ3[which(financial_summ3$Account==input$acct_variable4),-c(3)]
    }
    
    else if(input$acct_variable4=='Company'&&input$franchise_variable4!='All Franchises'){
      financial_group <-group_by(financial_combine,Sku,Month,Franchise)
      financial_summ3 <-summarise(financial_group,Description=first(`Description`),QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
      financial_summ3 <-financial_summ3[which(financial_summ3$Franchise==input$franchise_variable4),-c(3)]
    }
    
    else {
      financial_group <-group_by(financial_combine,Sku,Month,Account,Franchise)
      financial_summ3 <-summarise(financial_group,Description=first(`Description`),QTY=sum(QTY),Revenues=sum(`ExtendedRevenue`))
      financial_summ3 <-financial_summ3[which(financial_summ3$Franchise==input$franchise_variable4&financial_summ3$Account==input$acct_variable4),-c(3,4)]
    }
    validate(
      need(nrow(financial_summ3)!=0, "NO SELL-IN DATA"
      ))
    
    financial_summ3<-financial_summ3[order(financial_summ3$Month),]
    financial_summ3 <-financial_summ3[which(financial_summ3$Month>=as.Date(input$daterange_variable3[1])&financial_summ3$Month<=as.Date(input$daterange_variable3[2])),]
    validate(
      need(nrow(financial_summ3)!=0, "WRONG DATE RANGE"
      ))
    
    financial_summ3<-group_by(financial_summ3,Sku)
    
    if('Franchise'%in%colnames(financial_summ3)){financial_summ3<-summarise(financial_summ3,Description=first(`Description`),QTY=sum(`QTY`),Revenues=sum(`Revenues`),Franchise=first(`Franchise`))}
    else{financial_summ3<-cbind(summarise(financial_summ3,Description=first(`Description`),QTY=sum(`QTY`),Revenues=sum(`Revenues`)),Franchise=input$franchise_variable4)}
    
    
    return(financial_summ3)
  })
  
  skubar_ranking_function<-reactive({
    
    df<-skubar_function()[order(-skubar_function()$`Revenues`),]
    return(df)
  })
  
  acctpie_ranking_dollar <-reactive({
    df<-acctpie_function()[order(-acctpie_function()$`Revenues`),]
    return(df)
  })
  
  franpie_ranking_dollar <- reactive({
    df<-franpie_function()[order(-franpie_function()$`Revenues`),]
    return(df)
  })
  
  acctpie_ranking_qty <-reactive({
    df<-acctpie_function()[order(-acctpie_function()$`QTY`),]
    return(df)
  })
  
  franpie_ranking_qty <- reactive({
    df<-franpie_function()[order(-franpie_function()$`QTY`),]
    return(df)
  })
  
  
  output$sku_contain<-renderDataTable(
    DT::datatable(sku_contain_function(),options = list( columnDefs = list(list(className = 'dt-center'))),rownames= FALSE)%>%
      DT::formatStyle(columns = colnames(.), fontSize = '50%')
  )
  output$sku_belong<-renderDataTable(
    DT::datatable(sku_belong_function(),options = list( columnDefs = list(list(className = 'dt-center'))),rownames= FALSE)%>%
      DT::formatStyle(columns = colnames(.), fontSize = '50%')
  )
  
  sku_contain_function <-reactive({
    validate(
      need(input$checksku_variable2, "PLEASE SELECT SKU TO CHECK BOM")
    )
    
    df<-BOM_list[which(BOM_list$Item_Number_FGI==as.character(input$sku_variable2)),c(3:6)]
    df<-df[order(df[,4]),-c(4)]
    df<-map_df(df,rev)
    colnames(df)=c('SKU consists of','Description','Units Included')
    return(df)
  })
  
  sku_belong_function <-reactive({
    validate(
      need(input$checksku_variable2, "PLEASE SELECT SKU TO CHECK BOM")
    )
    
    df<-BOM_list[which(BOM_list$CMPTITNM_C==as.character(input$sku_variable2)),c(1,2,5)]
    colnames(df)=c('SKU kitted under','Description','Units Needed')
    return(df)
  })
  
  
  ################## starting from here, is the Fill Rate Testing  outputs ##########################
  
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
  
  output$sku_testing3 <-renderUI({
    if(!input$checksku_variable3)
      return()
    else
      selectInput("sku_variable3",paste("      ","2). Type or select SKU:",sep="   "),
                  choices=SKU_list,
                  selected=NULL, multiple=FALSE)
  })
  observe({
    if(!input$checksku_variable3) return()
    
    else
    {
      updateCheckboxGroupInput(session,"franchise_variable5",choices=franchise_list2, inline=TRUE)
    }
  })
  
  ########################### copy from here #######################
  output$fillrate_chart <-renderHighchart(
    highchart()%>%
      
      hc_title(text="Ordered v.s. Fulfilled / Fill Rate")%>%
      #{if(!input$checksku_variable1) hc_title(text="Shipment Running Log (Franchises)") else .}
      hc_subtitle(text=if(!input$checksku_variable3) paste("Selected",length(input$franchise_variable5),"Franchise(s)",sep=' ') else paste(input$sku_variable3,FR_function()$description[1],sep=" "))%>%
      
      #hc_title(text=paste("Shipment Running Log:", input$sku_variable1,sep = " "))%>%
      hc_tooltip(crosshairs=TRUE, valueDecimals = 2, borderWidth = 5, sort = TRUE, table = TRUE)%>%
      hc_xAxis(list(categories=FR_function()$forecasted_account))%>%
      hc_yAxis_multiples(
        
        list(title=list(text="Orders v.s. Fulfillment")),
        list(title=list(text="Monthly Fill Rate"),opposite=TRUE)
        
      ) %>% 
      hc_plotOptions(column=list(
        dataLabels=list(enabled=FALSE),
        
        enableMouseTracking = TRUE
      ))%>%
      
      
      hc_plotOptions(spline=list(color="#F4A460",
                                 dataLabels=list(enabled=FALSE),
                                 stacking="normal",
                                 enableMouseTracking = TRUE
      ))%>%
      
      hc_add_series(data = FR_function()$`Order QTY`, type = "column",  name="Monthly Orders",yAxis=0,color="#6DBEEA") %>%
      hc_add_series(data = FR_function()$`Fulfill QTY`, type = "column", name="Monthly Fulfillment",yAxis=0,color="#32CD32") %>%
      hc_add_series(data = FR_function()$fillrate,type='spline', name="Monthly Fill Rate",yAxis=1)%>%
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = TRUE) %>%
      hc_exporting(enabled = TRUE,fallbackToExportServer = F,
                   filename = "Monthly-FillRate")%>%
      #hc_legend(enabled = TRUE)%>%
      hc_add_theme(hc_theme_google()) 
  )
  
  output$fillrate_table<-DT::renderDataTable(
    # DT::datatable(MTD_group_function(),options = list(searching=FALSE),rownames= FALSE)%>%
    
    DT::datatable(FR_function(),options = list( columnDefs = list(list(className = 'dt-center', targets = 0:4))),rownames= FALSE)%>%
      
      formatRound(c("Order QTY","Fulfill QTY"), digits = 0, interval = 3, mark = ",", 
                  dec.mark = getOption("OutDec"))%>%
      formatPercentage(c("fillrate"),0)
  )
  output$top_cuts <-DT::renderDataTable(
    DT::datatable(cut_function(),options = list( columnDefs = list(list(className = 'dt-center', targets = 0:4))),rownames= FALSE)%>%
      formatRound(c(3,4,5), digits = 0, interval = 3, mark = ",", 
                  dec.mark = getOption("OutDec"))
  )
  
  cut_function <-reactive({
    cut_data <-combine_FR_data[which(combine_FR_data$`Date Range`==input$daterange_variable5&combine_FR_data$forecasted_account==input$acct_variable6),]
    group_cut <-group_by(cut_data, `Item #`)
    summ_cut<-summarise(group_cut, description=first(item_description), `Order QTY`=sum(`Order QTY`),`Fulfill QTY`=sum(`Fulfill QTY`))
    summ_cut$`Cut QTY`=as.numeric(summ_cut$`Order QTY`)- as.numeric(summ_cut$`Fulfill QTY`)
    summ_cut=summ_cut[order(-summ_cut$`Cut QTY`),]
    return_cut <-head(summ_cut, as.numeric(input$skucount_variable2))
    return(return_cut)
  })
  
  FR_function <-reactive({
    
    
    
    if(input$checksku_variable3){
      checking_sku2 <-group_by(combine_FR_data,`Date Range`,`forecasted_account`,`Item #`)
      group_summ<-summarise(checking_sku2,description=first(item_description),`Order QTY`=sum(`Order QTY`),`Fulfill QTY`=sum(`Fulfill QTY`))
      group_DF <-group_summ[which(group_summ$`Item #`==input$sku_variable3&group_summ$`forecasted_account`%in%input$acct_variable5),]
      
      return_DF <-group_DF[which(group_DF$`Date Range`==input$daterange_variable4),]
      return_DF<-group_by(return_DF,`Date Range`,`Item #`,forecasted_account)
      return_DF<-summarise(return_DF,description=first(description),`Order QTY`=sum(`Order QTY`),`Fulfill QTY`=sum(`Fulfill QTY`))
      
      
      
    }
    else{
      checking_fran2 <-group_by(combine_FR_data,`Date Range`,`forecasted_account`,`Franchise`)
      group_summ <-summarise(checking_fran2,`Order QTY`=sum(`Order QTY`),`Fulfill QTY`=sum(`Fulfill QTY`))
      group_DF <- group_summ[which(group_summ$`forecasted_account`%in%input$acct_variable5&group_summ$Franchise%in%input$franchise_variable5),]
      
      return_DF <-group_DF[which(group_DF$`Date Range`==input$daterange_variable4),]
      
      
      return_DF<-group_by(return_DF,`Date Range`,forecasted_account)
      return_DF<-summarise(return_DF,`Order QTY`=sum(`Order QTY`),`Fulfill QTY`=sum(`Fulfill QTY`))
      
    }  
    
    
    
    return_DF$fillrate=return_DF$`Fulfill QTY`/return_DF$`Order QTY`
    
    validate(
      need(nrow(return_DF)!=0, "PLEASE INPUT PROPER DATA")
    )
    return(return_DF)
    
  })
  
  MTD_pie_function<-reactive({
    
    ship_runrate <- read_excel('ship runrate raw data.xlsx', sheet = input$daterange_variable1, col_types=c("text","text","text","text","text","numeric","numeric"),skip=0)
    #runrate_group<-group_by(ship_runrate,rolling_index,forecasted_account
    
    unique_index=unique(ship_runrate$rolling_index)
    MTD_index=unique_index[length(unique_index)]
    if(input$checksku_variable1){
      pie_df=ship_runrate[which(ship_runrate$rolling_index==MTD_index&ship_runrate$`Item #`==input$sku_variable1),]
    }
    else{
      pie_df=ship_runrate[which(ship_runrate$rolling_index==MTD_index&ship_runrate$Franchise%in%input$franchise_variable1),]
    }
    
    pie_group=group_by(pie_df,forecasted_account)
    pie_DF<-summarise(pie_group,`MTD ship`=sum(`Ship Qty`),`rolling_index`=first(`rolling_index`))
    pie_DF<-pie_DF[which(pie_DF$`MTD ship`>0),]
    pie_DF<-pie_DF[order(-pie_DF$`MTD ship`),]
    names(pie_DF)=c("Account","QTY","MTD")
    
    validate(
      need(nrow(pie_DF)!=0, "No MTD shipment information")
    )
    return(pie_DF)
  })
  
  
  
  
}


## Deploy App
shinyApp(ui, server)


