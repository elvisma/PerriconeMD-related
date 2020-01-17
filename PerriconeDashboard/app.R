
####### PerriconeMD dashboard ########

# Read Libraries
#install.packages("devtools") 
#devtools::install_github("tidyverse/readxl")

library("readxl")
library("reshape")
options(scipen=999)

library("rsconnect")
library("shiny")
library("stringi")
library("highcharter")
library("shinythemes")
library("shinydashboard")
library("dplyr")
library("DT")



setwd("C:/Users/ema/Desktop/projects/Perricone Dashboard")
SKU_list <-unique(read_excel("SmartList BOM Extract.xlsx", sheet=1,col_types='text')$`Item_Number_FGI`)


##### Build UI ########
ui <- navbarPage("Perricone Dashboard", theme = shinytheme("flatly"),
                 tabPanel("Monthly Shipment Runrate",
                          fluidPage(
                            sidebarLayout(
                              sidebarPanel(width =3,
                                           
                                           div(style="display: inline-block;",radioButtons("acct_variable1",h4(p(strong("Choose Accounts:"))),inline=TRUE,
                                                              choices=list("Company","Int'l"="APAC+LA","Amazon","Bloomingdales","Costco","Dillard's","EC Scott","GR"="Guthy Renker","L&T"="Lord & Taylor","Macy's","NM"="Neiman Marcus","Nordstrom","Other.com","PMD.com","QVC","Sample","Sephora","TSC","UK","ULTA"="Ulta","Zulily","NCO/Liquidators",'Component transfer')
                                                              #c("Amazon"="AMZ","Costco","Walmart","Sam's Club","Wayfair","Overstock","Target","Homedepot","Zinus.com","MACY's.com"="Macys","CHEWY.com"="CHEWY.COM"),
                                           )),
                                           hr(),
                                           h4(p(strong("Check by Franchise"))),
                                           div(style="display: inline-block;",checkboxGroupInput("franchise_variable1","Select Franchises:",inline=TRUE,
                                                              c("HP Classics"="High Potency Classics","Heritage","No Makeup"="No Makeup Skincare","No:Rinse","Intensive Pore","Pre:Empt","Masks","H2 EE"="H2 Elemental Energy","VC Ester"="Vitamin C Ester","Supplements","Neuropetide","Cold Plasma","Essential Fx"="Essential Fx Acyl Glutathione","Re:Firm","Thio:Plex","Hypoallergenic","Mixed Franchise","Acne","Hypo CBD"="Hypoallergenic CBD","Component"),
                                                              #c("Class A"="A","Class B"="B","Class C"="C","TOTAL"),
                                                              selected = c("HP Classics"="High Potency Classics","Heritage","No Makeup"="No Makeup Skincare","No:Rinse","Intensive Pore","Pre:Empt","Masks","H2 EE"="H2 Elemental Energy","VC Ester"="Vitamin C Ester","Supplements","Neuropetide","Cold Plasma","Essential Fx"="Essential Fx Acyl Glutathione","Re:Firm","Thio:Plex","Hypoallergenic","Mixed Franchise","Acne","Hypo CBD"="Hypoallergenic CBD")
                                           )),
                                          
                                           
                                           h4(p(strong("Check by SKU"))),
                                           
                                           
                                           checkboxInput("checksku_variable1",label = p(strong("1). check to disregard Franchise Selection")), value=FALSE),
                                           uiOutput("sku_testing"),
                                           selectInput("daterange_variable1",h4(p(strong("Date Range:"))),
                                                       c("Jan 2020 (in process)","Dec 2019","Nov 2019"),
                                                       selected = c("Jan 2020 (in process)")
                                           ),
                                           
                                           #selectInput("sku_variable1",paste("      ","2). Type or select SKU:",sep="     "),
                                            #            choices=SKU_list,
                                            #          selected=NULL, multiple=FALSE
                        
                                           #),
                                           
                                           downloadButton("downloadData","Download")
                              ),
                              mainPanel(width = 8,
                                        h3("Rolling MTD shipment"),
                                        br(),
                                        tabsetPanel(
                                          tabPanel("Chart",highchartOutput("MTDchart",height = "680px")),
                                          tabPanel("Table",dataTableOutput("MTDtable",height='680px'))
                                        ))),
                            hr(),
                            ### Newly Added ###
                            h3("Forecast Accuracy & MAPE (under testing)"),
                            br(),
                            sidebarLayout(
                              sidebarPanel(width = 2,
                                           div(style="display: inline-block;",radioButtons("acct_variable2","Choose Accounts:",inline=TRUE,
                                                                                           choices=list("Company","Int'l"="APAC+LA","Amazon","Bloomingdales","Costco","Dillard's","EC Scott","GR"="Guthy Renker","L&T"="Lord & Taylor","Macy's","NM"="Neiman Marcus","Nordstrom","Other.com","PMD.com","QVC","Sample","Sephora","TSC","UK","Ulta","Zulily","Liquidator")
                                                                                           #c("Amazon"="AMZ","Costco","Walmart","Sam's Club","Wayfair","Overstock","Target","Homedepot","Zinus.com","MACY's.com"="Macys","CHEWY.com"="CHEWY.COM"),
                                           )),   
                                           
                              ),
                              
                              mainPanel(width=10,
                                        br(),
                                        highchartOutput("MAPEtracker",height = "480px")
                              )
                            ),
                            br(),
                            hr(),
                            hr(),br()
                            ### Newly Added ended ###
                          )),
                 tabPanel("Actuals v.s. Forecasts",
                          sidebarPanel(width =2,
                                       div(style="display: inline-block;",radioButtons("acct_variable3","Choose Accounts:",inline=TRUE,
                                                                                       choices=list("Company","Int'l"="APAC+LA","Amazon","Bloomingdales","Costco","Dillard's","EC Scott","GR"="Guthy Renker","L&T"="Lord & Taylor","Macy's","NM"="Neiman Marcus","Nordstrom","Other.com","PMD.com","QVC","Sample","Sephora","TSC","UK","Ulta","Zulily","Liquidator")
                                                                                       #c("Amazon"="AMZ","Costco","Walmart","Sam's Club","Wayfair","Overstock","Target","Homedepot","Zinus.com","MACY's.com"="Macys","CHEWY.com"="CHEWY.COM"),
                                       )),   
                                       checkboxGroupInput("variable4","Choose Classes:",
                                                          c("Class A"="A","Class B"="B","Class C"="C","TOTAL"),
                                                          selected = c("TOTAL")
                                       ),
                                       selectInput("dateRange2","Date Range Selection:",
                                                   c("Dec 2019","Nov 2019"),
                                                   selected = c("Dec 2019")
                                       )
                          ),
                          mainPanel(width = 10,
                                    h3("under testing"),
                                    br(),
                                    tabsetPanel(
                                      tabPanel("Accomplishment By $", highchartOutput("DollarChart", height = "480px")),
                                      tabPanel("Accomplishment Rate By %",highchartOutput("PercentChart", height = "480px"))
                                    )
                                    
                          ))
                 
                 
)



###### Build Server #####
server <- function(input, output){

  output$MTDchart <-renderHighchart(
    highchart()%>%
      
      hc_title(text=paste("Shipment Running Log for",input$acct_variable1,sep=' '))%>%
      #{if(!input$checksku_variable1) hc_title(text="Shipment Running Log (Franchises)") else .}
      hc_subtitle(text=if(!input$checksku_variable1) "Selected Franchises" else paste(input$sku_variable1,description_function()$description[1],sep=" "))%>%
      
      #hc_title(text=paste("Shipment Running Log:", input$sku_variable1,sep = " "))%>%
      hc_tooltip(crosshairs=TRUE, valueDecimals = 2, borderWidth = 5, sort = TRUE, table = TRUE)%>%
      hc_xAxis(list(categories=MTD_group_function()$`rolling_index`))%>%
      hc_yAxis_multiples(
        
        list(title=list(text="Shipped UNITS")),
        list(title=list(text="MTD runrate"),opposite=TRUE)
        
      ) %>% 
      hc_plotOptions(column=list(color="limegreen",
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
      
      hc_add_series(data = MTD_group_function()$fcstline, type = "area",  name="Monthly Forecast Line",yAxis=0) %>%
      hc_add_series(data = MTD_group_function()$MTDship, type = "column", name="MTD shipment",yAxis=0) %>%
      hc_add_series(data = MTD_group_function()$runrate,type='spline', name="MTD runrate",yAxis=1)%>%
      
      hc_legend(enabled = TRUE)%>%
      hc_add_theme(hc_theme_google()) 
  )
  
  output$sku_testing <- renderUI({
    #if(!input$checksku_variable1)
    #  selectInput("sku_variabletesting","Type or Select SKU:",choices=SKU_list)
    #else
    if((input$checksku_variable1)==FALSE)
      return()
    else
      
      selectInput("sku_variable1",paste("      ","2). Type or select SKU:",sep="   "),
                  choices=SKU_list,
                  selected=NULL, multiple=FALSE)
                  
      
      
  })
  output$MTDtable <-DT::renderDataTable(
   # DT::datatable(MTD_group_function(),options = list(searching=FALSE),rownames= FALSE)%>%
    
    DT::datatable(cbind(MTD_group_function(),SKU= if(is.null(input$sku_variable1))" " else input$sku_variable1,Account=input$acct_variable1),options = list(searching=FALSE),rownames= FALSE)%>%
     
    formatRound(c("MTDship","fcstline"), digits = 0, interval = 3, mark = ",", 
                dec.mark = getOption("OutDec"))%>%
    formatPercentage(c("runrate"),2)
      
  )
  output$downloadData <-downloadHandler(
    ###### this function only works in the browser window !!!!!!!!! #############
    filename = function() {
      paste(input$daterange_variable1,input$acct_variable1, "Shipment runrate.csv", sep = " ")
    },
    
    
    content = function(file) {
      write.csv(cbind(MTD_group_function(),SKU=if(is.null(input$sku_variable1))" " else input$sku_variable1,Account=input$acct_variable1),file,row.names = FALSE)
    }
  )
  MTD_group_function <-reactive({
    ship_runrate <- read_excel('ship runrate raw data (monthly).xlsx', sheet = input$daterange_variable1, col_types=c("text","text","text","text","text","numeric","numeric"),skip=0)
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
    return_DF<-summarise(return_DF,MTDship=sum(shipTotal),fcstline=sum(fcstTotal))
    return_DF$runrate=return_DF$MTDship/return_DF$fcstline
    validate(
      need(nrow(return_DF)!=0, "NO SELL IN DATA")
    )
    return(return_DF)
    
  })
  
  
  
  description_function <-reactive({
    ship_runrate <- read_excel('ship runrate raw data (monthly).xlsx', sheet = input$daterange_variable1, col_types=c("text","text","text","text","text","numeric","numeric"),skip=0)
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
    return_DF<-summarise(return_DF,description=first(description),MTDship=sum(shipTotal),fcstline=sum(fcstTotal))
    
    return(return_DF)
    
  })
  
 


  #### testing function #### 
  #ship_runrate_testing <- read_excel('ship runrate raw data (monthly).xlsx', sheet = "Dec 2019", col_types=c("text","text","text","text","text","numeric","numeric"),skip=0)
  #testing_sku <-group_by(ship_runrate_testing,`rolling_index`,`forecasted_account`,`Item #`)
  #group_summ <-summarise(testing_sku,description=first(item_description),shipTotal=sum(`Ship Qty`),fcstTotal=sum(`Monthly Forecast`))
  #group_DF <-group_summ[which(group_summ$`Item #`=="51080001"&group_summ$`forecasted_account`=="Ulta"),]
  #group_DF<-group_by(group_DF,rolling_index)
  #group_DF <-summarise(group_DF,shipTotal=sum(shipTotal),fcstTotal=sum(fcstTotal))
  #group_DF$runrate <-group_DF$shipTotal/group_DF$fcstTotal
}


## Deploy App
shinyApp(ui, server)


