
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

setwd("C:/Users/ema/Desktop/projects/Perricone Dashboard")
SOP_publish_date="from S&OP forecasting publications (updated @ 2/21/2020)"
IDC_update_date='from IDC shipment reports (updated @ 3/9/2020)'
financial_report_date='from financial reports (updated @ 2/13/2020)'
lag_clarification_date='Feb S&OP for Lag0,  Jan S&OP for Lag1'

date_ranges=c("Mar 2020 (in process)","Feb 2020","Jan 2020","Dec 2019","Nov 2019")
fill_date_ranges=c("Feb 2020","Jan 2020","Dec 2019","Nov 2019")
##############################   getting fill rate data   #################################
## need to specify column types, otherwise generate some NAs
fill_rate_data_raw <- read_excel("FillRate raw data.xlsx", sheet="Sheet1",col_types=c("text","text","text","text","text","text","text","text","text","numeric","numeric","numeric","text"),skip=0)

fill_rate_data <-fill_rate_data_raw[,c('Item #',"Franchise","forecasted_account",'Date Range','item_description',"Order QTY","Fulfill QTY")]
fill_rate_data$`Item #`<-as.character(fill_rate_data$`Item #`)

### combine all accounts
groupAll_fill_rate_data<-group_by(fill_rate_data,`Item #`,Franchise, `Date Range`)
sumAll_fill_rate_data <-summarise(groupAll_fill_rate_data,item_description=first(`item_description`),`Order QTY`=sum(`Order QTY`),`Fulfill QTY`=sum(`Fulfill QTY`))
sumAll_fill_rate_data$`forecasted_account`="B2B customers"
sumAll_fill_rate_data=sumAll_fill_rate_data[,c('Item #',"Franchise","forecasted_account",'Date Range','item_description',"Order QTY","Fulfill QTY")]
combine_FR_data <- rbind(as.data.frame(fill_rate_data),as.data.frame(sumAll_fill_rate_data))
#########################################################################
##############################   getting fill rate data   #################################

BOM_list <-as.data.frame(read_excel("SmartList BOM Extract.xlsx", sheet=1,col_types='text'))[,c(1,2,9,10,11,13)]
SKU_list <-unique(BOM_list$Item_Number_FGI)
franchise_list <- c("Cold Plasma","HP Classics"="High Potency Classics","No Makeup"="No Makeup Skincare","VC Ester"="Vitamin C Ester","Supplements","Neuropeptide","Essential Fx"="Essential Fx Acyl Glutathione","Hypoallergenic","Mixed Franchise","Acne","Masks","Hypo CBD"="Hypoallergenic CBD","No:Rinse","Heritage","Intensive Pore","Pre:Empt","Mens","H2 EE"="H2 Elemental Energy","Re:Firm","Thio:Plex","Sample")
franchise_list2 <-franchise_list
franchise_vector <-append(franchise_list,"All Franchises", after=0)
acct_vector <- c("Company","Int'l"="APAC+LA","Amazon","Bloomingdales","Costco","Dillard's","EC Scott","GR"="Guthy Renker","L&T"="Lord & Taylor","Macy's","NM"="Neiman Marcus","Nordstrom","Other.com","PMD.com","QVC","Sephora","TSC","UK","ULTA"="Ulta","Zulily",'Retail - Other')

B2B_list <- c("B2B customers","Int'l"="APAC+LA","Amazon","Bloomingdales","Costco","Dillard's","EC Scott","L&T"="Lord & Taylor","Macy's","NM"="Neiman Marcus","Nordstrom","Sephora","TSC","UK","ULTA"="Ulta","Zulily","NCO","Retail - Other")


##### Build UI ########
ui <- navbarPage("Fill Rate Testing", theme = shinytheme("flatly"),
    
                
                 tabPanel("Monthly Scorecard",
                          fluidPage(
                           
                            ######################################### starting from here for fill rate testing ############################################                
                            sidebarLayout(
                              sidebarPanel(width =3,
                                           h4(p(strong("Choose Account:"))),
                                           
                                           actionLink("selectall2","Select/Unselect All Accounts"), 
                                           div(style="display: inline-block;",checkboxGroupInput("acct_variable5"," ",inline=TRUE,B2B_list,
                                                                                                 c('B2B customers','Amazon','Costco',"Macy's","Sephora","UK","ULTA"='Ulta','Retail-Other')        
                                           )),
                                           hr(),
                                           h4(p(strong("Check by Franchise"))),
                                           actionLink("selectall3","Select/Unselect All Franchises"), 
                                           div(style="display: inline-block;",checkboxGroupInput("franchise_variable5"," ",inline=TRUE,franchise_list,
                                                                                                 c("HP Classics"="High Potency Classics","No Makeup"="No Makeup Skincare","Intensive Pore","VC Ester"="Vitamin C Ester","Supplements","Neuropeptide","Cold Plasma","Essential Fx"="Essential Fx Acyl Glutathione","Hypoallergenic","Mixed Franchise","Acne","Hypo CBD"="Hypoallergenic CBD")
                                                                                                 
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
                                        h3("Monthly Fill Rate (B2B customers)"),
                                        h5(IDC_update_date),
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
                            )
                            
                            ######################################### ending from here for fill rate testing ############################################     
                          ))
                 
                 
                 
)



###### Build Server #####
server <- function(input, output, session){
  
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
      
      hc_add_series(data = FR_function()$`Order QTY`, type = "column",  name="Monthly Orders",yAxis=0,color="#89CFF0") %>%
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
      DT::datatable(cut_function(),options = list( columnDefs = list(list(className = 'dt-center', targets = 0:4))))%>%
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
  
  
  #combine_FR_data
  #acct_variable5
  #acct_variable6
  #franchise_variable5
  #checksku_variable3
  #daterange_variable4
  #daterange_variable5
  #sku_variable3
  #skucount_variable2
  
  
  ########################### copy ends here #########################
  
  ##### testing FR_functions #####
  #checkfran_list = c("High Potency Classics","No Makeup Skincare","Intensive Pore","VC Ester"="Vitamin C Ester","Supplements","Neuropeptide","Cold Plasma","Essential Fx Acyl Glutathione","Hypoallergenic","Mixed Franchise","Acne","Hypoallergenic CBD")
  
  
  #checkacct_list = c('B2B customers','Amazon','Costco',"Macy's","Sephora","UK","ULTA"='Ulta','Retail-Other')    
  #group_DF <- group_summ[which(group_summ$`forecasted_account`%in%checkacct_list&group_summ$Franchise%in%checkfran_list),]
  #group_DF <- group_summ[which(group_summ$`forecasted_account`%in%checkacct_list&group_summ$`Item #`==51080001),]
  #return_DF <-group_DF[which(group_DF$`Date Range`==fill_date_ranges[1]),]
  #cut_data <-combine_FR_data[which(combine_FR_data$`Date Range`=="Feb 2020"&combine_FR_data$forecasted_account=="Ulta"),]
  #return_cut=head(summ_cut,10)
  
}


## Deploy App
shinyApp(ui, server)


