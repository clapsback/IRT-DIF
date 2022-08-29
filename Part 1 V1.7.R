#1.2: Adds in better error detection for imported files
#1.2: Fixed item plot titles
#1.3: Added error condition for purification of LR and MH so the app doesn't crash when no original DIF was detected
#1.4: Minor bug fixes
#1.5: Minor spelling issues and formatting changes
#1.6: Overhaul of ui and addition of omission selection
#1.7: Integrated instructions into app
#####Prep#####
library(shiny)
library(irtoys)
library(readxl)
library(DT)
library(difR)
library(ltm)
library(ggplot2)
library(openxlsx)
library(shinyjs)
library(shinydisconnect)
library(subscore)


runApp(
  list(
    ui = fluidPage(useShinyjs(), disconnectMessage("Something went wrong.  If the problem persists, please feel free to contact David at schreurd@uwm.edu for additional help."),
                   #Title
                   titlePanel("Part 1: Finding Anchor Items Through Differential Item Functioning"),
                   sidebarLayout(
                     #####Prep#####
                     
                     #####UI Selection#####
                     sidebarPanel(
                       #Input File
                       fileInput('file1', 'Choose xlsx File:', accept = c(".xlsx")),
                       
                       selectInput("Grouping", "Select Grouping Heading (Column with Exam Indicator)", choices = "NULL"),
                       
                       numericInput('alpha', "Alpha Level to Run all DIF Analyses at:", 0.05, min = 0.00000000001, max = 0.1),
                       
                       checkboxInput("ItemPlots", "Would You Like Item Plots to be Generated?", FALSE), #Omission
                       
                       checkboxInput("FitInput", "Would You Like the IRT Model Fit to be Calculated?", FALSE), #Omission
                       
                       checkboxInput("Omission", "Are There any Columns in Your Uploaded Data You Would Like to be Omitted?", FALSE), #Omission
                       
                       uiOutput("ui"),
                       #checkboxGroupInput("Omits", "A", choices = "NULL"),  #Omit
                       
                       actionButton("Submit", "Run DIF Analysis"),
                       
                       actionButton("Refresh", "Reset to Run New Analysis"),
                       
                       fluidRow(column(12, "", disabled(downloadButton("report", "Generate Full DIF Report (May Take a Few Minutes)"))))
                       
                       
                     ),
                     #####UI Selection#####          
                     
                     #####Output Shown#####          
                     mainPanel(
                       h3(textOutput("Error", container = span)),
                       h3(textOutput("NoError", container = span)),
                       #tableOutput("DIFItems"),
                       tableOutput("DIFItems")
                       
                     ),
                   )
    ),
    #####Output Shown#####
    
    
    server = function(input, output, session) {
      
      #####Change Grouping Options to Data Headings#####
      observe({
        x <- input$file1
        
        # Can use character(0) to remove all choices
        if (is.null(x)){
          x <- c("Pending Upload")
          
          # Can also set the label and select items
          updateSelectInput(session, "Grouping", choices = x)}
        else{
          ImportedData <- read_excel(input$file1$datapath)
          as.data.frame(ImportedData) 
          x <- colnames(ImportedData)
          updateSelectInput(session, "Grouping", choices = x)
          
          
        }})
      #####Change Grouping Options to Data Headings#####
      
      
      
      #####Change Grouping Options to Data Headings 2#####
      observeEvent(input$Omission,{
        x <- input$Omission
        # Can use character(0) to remove all choices
        if (identical(input$Omission, FALSE) ){
          x <- c("Pending Upload")
          # Can also set the label and select items
          #updateSelectInput(session, "Grouping", choices = x)}
        }
        else{
          test_data <- reactive({
            inFile <- input$file1
            if (is.null(inFile))
              return(NULL)  #data.frame(V1=character(0), V2=integer(0), V3=integer(0)))
            ImportedData <- read_excel(input$file1$datapath)})
          
          # Trying to generate the check boxes
          output$ui <- renderUI(checkboxGroupInput('Omits', 'Please Select Any Columns You Would Like to Be Omitted From Analysis (e.g. Student Identifiers, Student Sex, Undesired Questions, Ext.)', colnames(test_data())))
        }
        
      }, ignoreInit = TRUE)
      #####Change Grouping Options to Data Headings 2#####
      
      
      #####Reset App#####
      observeEvent(input$Refresh, {
        refresh()
      })
      #####Reset App#####
      
      
      #####Run When Submit Button Pressed#####
      observeEvent(input$Submit, {
        
        Alpha=input$alpha
        
        ##No File Uploaded##
        x <- input$Grouping
        if (identical(x, "Pending Upload")){output$Error=renderText("Error: No File Selected for Analysis!")}
        else{
          
          ##Import Data##
          ItemAnalysisRightWrong <- read_excel(input$file1$datapath)
          UntouchedItemAnalysisRightWrong <- ItemAnalysisRightWrong
          ItemAnalysisRightWrong <- as.data.frame(ItemAnalysisRightWrong)
          
          ExcelGrouping <- ItemAnalysisRightWrong[input$Grouping]
          ExcelGrouping <- unname(unlist(ExcelGrouping))
          
          ItemAnalysisRightWrong = subset(ItemAnalysisRightWrong, select = !(names(ItemAnalysisRightWrong) %in% input$Omits))
          
          ExcelQuestions <- ItemAnalysisRightWrong
          ExcelQuestions[input$Grouping] <- NULL
          ##Import Data##
          
          
          ####Error Checking####
          ImportedData <- ExcelQuestions
          
          ErrorCheck=ImportedData
          ErrorCheck[is.na(ErrorCheck)] <- 2
          ErrorCheck[ErrorCheck == 0] <- 999
          ErrorCheck[ErrorCheck == 1] <- 999
          
          ErrorCheck <- tryCatch({ErrorCheck/999},
                                 error=function(cond) {return(rbind(rep(0,ncol(ErrorCheck)),rep(0,ncol(ErrorCheck))))})
          
          ErrorCheck=colSums(ErrorCheck)
          ErrorCheck[ErrorCheck == nrow(ImportedData)] <- 1
          
          x <- input$Grouping
          if (sum(ErrorCheck) != ncol(ImportedData)){output$Error=renderText("Error! All values besides headings in the excel document must be a '0' or a '1'.  Please check document for missing values or other values.")}
          ####Error Checking####
          else{
            output$NoError=renderText(paste0("An 'X' Indicates DIF Was Detected on That Question Based on Alpha of ", Alpha))
            
            ExcelScore <- rowSums(ExcelQuestions)
            
            
            ####MH####
            withProgress(message = 'Conducting Mantel–Haenszel', value = 0, {
              
              res.MH<- difMH(Data=ExcelQuestions, group=ExcelGrouping, focal.name = 0, alpha = Alpha)
              MHFullResults=cbind(res.MH[["MH"]],res.MH[["p.value"]],res.MH[["alphaMH"]])
              colnames(MHFullResults) <- c("MH Chi-Squared", "MH Significance", "MH Odds Ratio")
              rownames(MHFullResults) <- res.MH[["names"]]
              MHQuestionsShowingDIF=res.MH[["DIFitems"]]
              MHp <- res.MH[["p.value"]]
              #  write.xlsx(MHFullResults, file = paste0(HomePath, "/DIF_Results.xlsx"), sheetName = "MHFullResults", col.names = TRUE, row.names = TRUE, append = FALSE)
              #  write.xlsx(MHQuestionsShowingDIF, file = paste0(HomePath, "/DIF_Results.xlsx"), sheetName="MHQuestionsShowingDIF", append=TRUE)
              ####MH####
              
              
              ####MH Purified####
              
              res.MH=tryCatch({difMH(Data=ExcelQuestions, group=ExcelGrouping, focal.name = 0, purify = TRUE, nrIter=50, alpha = Alpha)},
                              error=function(cond) {return(difMH(Data=ExcelQuestions, group=ExcelGrouping, focal.name = 0, purify = FALSE, nrIter=50, alpha = Alpha))})
              
              
              
              MHPureFullResults=cbind(res.MH[["MH"]],res.MH[["p.value"]],res.MH[["alphaMH"]])
              colnames(MHPureFullResults) <- c("Pure MH Chi-Squared", "Pure MH Significance", "Pure Odds Ratio")
              rownames(MHPureFullResults) <- res.MH[["names"]]
              MHPureQuestionsShowingDIF=res.MH[["DIFitems"]]
              MHPurep <- res.MH[["p.value"]]
            })
            #  write.xlsx(MHPureFullResults, file = paste0(HomePath, "/DIF_Results.xlsx"), sheetName = "MHPureFullResults", append = TRUE)
            #  write.xlsx(MHPureQuestionsShowingDIF, file = paste0(HomePath, "/DIF_Results.xlsx"), sheetName="MHPureQuestionsShowingDIF", append=TRUE)
            ####MH Purified####
            
            
            ####LR####
            withProgress(message = 'Conducting Logistic Regression', value = 0, {
              ExcelQuestionsM=data.matrix(ExcelQuestions)
              #ExcelGrouping=as.factor(ExcelGrouping)
              LRUDIF=difGenLogistic(Data = ExcelQuestionsM, group = ExcelGrouping, focal.name = 0, type = "udif", alpha = Alpha)   #Uniform DIF
              LRNUDIF=difGenLogistic(Data = ExcelQuestionsM, group = ExcelGrouping, focal.name = 0, type = "nudif", alpha = Alpha)  #Non Uniform DIF
              
              LRUFullResults=cbind(LRUDIF[["p.value"]])
              LRUQuestionsShowingDIF=LRUDIF[["DIFitems"]]
              LRNUFullResults=cbind(LRNUDIF[["p.value"]])
              LRNUQuestionsShowingDIF=LRNUDIF[["DIFitems"]]
              colnames(LRUFullResults) <- c("LR Uniform Significance")
              rownames(LRUFullResults) <- res.MH[["names"]]
              colnames(LRNUFullResults) <- c("LR Nonuniform Significance")
              rownames(LRNUFullResults) <- res.MH[["names"]]
              
              #  write.xlsx(LRUFullResults, file = paste0(HomePath, "/DIF_Results.xlsx"), sheetName="LRUFullResults", append=TRUE)
              #  write.xlsx(LRUQuestionsShowingDIF, file = paste0(HomePath, "/DIF_Results.xlsx"), sheetName="LRUQuestionsShowingDIF", append=TRUE)
              #  write.xlsx(LRNUFullResults, file = paste0(HomePath, "/DIF_Results.xlsx"), sheetName="LRNUFullResults", append=TRUE)
              #  write.xlsx(LRNUQuestionsShowingDIF, file = paste0(HomePath, "/DIF_Results.xlsx"), sheetName="LRNUQuestionsShowingDIF", append=TRUE)
              ####LR####
              
              
              ####LR Pure####
              LRNUDIF=tryCatch({difGenLogistic(Data = ExcelQuestionsM, group = ExcelGrouping, focal.name = 0, type = "udif", purify = TRUE, nrIter=50, alpha = Alpha)},
                               error=function(cond) {return(difGenLogistic(Data = ExcelQuestionsM, group = ExcelGrouping, focal.name = 0, type = "udif", purify = FALSE, nrIter=50, alpha = Alpha))})
              
              LRNUDIF=tryCatch({difGenLogistic(Data = ExcelQuestionsM, group = ExcelGrouping, focal.name = 0, type = "nudif", purify = TRUE, nrIter=50, alpha = Alpha)},
                               error=function(cond) {return(difGenLogistic(Data = ExcelQuestionsM, group = ExcelGrouping, focal.name = 0, type = "nudif", purify = FALSE, nrIter=50, alpha = Alpha))})
              
              
              
              
              LRUPureFullResults=cbind(LRUDIF[["p.value"]])
              LRUPureQuestionsShowingDIF=LRUDIF[["DIFitems"]]
              LRNUPureFullResults=cbind(LRNUDIF[["p.value"]])
              LRNUPureQuestionsShowingDIF=LRNUDIF[["DIFitems"]]
              colnames(LRUPureFullResults) <- c("Pure LR Uniform Significance")
              rownames(LRUPureFullResults) <- res.MH[["names"]]
              colnames(LRNUPureFullResults) <- c("Pure LR Nonuniform Significance")
              rownames(LRNUPureFullResults) <- res.MH[["names"]]
            })
            #  write.xlsx(LRUPureFullResults, file = paste0(HomePath, "/DIF_Results.xlsx"), sheetName="LRUPureFullResults", append=TRUE)
            #  write.xlsx(LRUQuestionsShowingDIF, file = paste0(HomePath, "/DIF_Results.xlsx"), sheetName="LRUPureQuestionsShowingDIF", append=TRUE)
            #  write.xlsx(LRNUPureFullResults, file = paste0(HomePath, "/DIF_Results.xlsx"), sheetName="LRNUPureFullResults", append=TRUE)
            #  write.xlsx(LRNUQuestionsShowingDIF, file = paste0(HomePath, "/DIF_Results.xlsx"), sheetName="LRNUPureQuestionsShowingDIF", append=TRUE)
            ####LR Pure####
            
            ####IRT####
            withProgress(message = 'Conducting Item Response Theory', value = 0, {
              Numberofquestions=ncol(ExcelQuestions)        
              Lordp=tryCatch({Lord1=difLord(Data = ExcelQuestions, group = ExcelGrouping, focal.name = 0, model = "2PL", same.scale = TRUE, purify = FALSE, nrIter = 1, alpha = Alpha)
              Lord1[["p.value"]]},
              error=function(cond) {return(rep(NA,Numberofquestions))})
              
              IRTPar=tryCatch({Lord1[["itemParInit"]]},
                              error=function(cond) {return(rep(NA,Numberofquestions))})
              
            })
            ####IRT####
            
            
            ####Pure IRT####
            withProgress(message = 'Conducting Purification of Item Response Theory', value = 0, {
              LordPurep=tryCatch({Lord2=difLord(Data = ExcelQuestions, group = ExcelGrouping, focal.name = 0, model = "2PL", purify = TRUE, nrIter = 50, alpha = Alpha)
              Lord2[["p.value"]]},
              error=function(cond) {return(rep(NA,Numberofquestions))})
              
              IRTParPure=tryCatch({Lord2[["itemParInit"]]},
                                  error=function(cond) {return(rep(NA,Numberofquestions))})
              
            })
            ####Pure IRT####
            
            
            ####Results####
            withProgress(message = 'Generating Final Results', value = 0, {
              
              ##Preview 
              DIFItems=cbind(MHp, MHPurep, Lordp, LordPurep, LRUFullResults, LRUPureFullResults, LRNUFullResults, LRNUPureFullResults)
              DIFItems=replace(DIFItems, DIFItems>Alpha, 999)
              DIFItems=replace(DIFItems, DIFItems<Alpha, "X")
              DIFItems=replace(DIFItems, DIFItems==999, "")  
              
              QNumbers=colnames(ExcelQuestions)
              DIFItems=cbind(QNumbers, DIFItems)
              colnames(DIFItems) <- c("Question # Based on Excel Heading", "Mantel–Haenszel", "Purified Mantel–Haenszel", "Lord with 2-PL IRT", "Purified Lord with 2-PL IRT", "Logistic Regression Uniform DIF", "Purified Logistic Regression Uniform DIF", "Logistic Regression Non-Uniform DIF", "Purified Logistic Regression Non-Uniform DIF")
              #  DIFItems=qpcR:::cbind.na(MHQuestionsShowingDIF, MHPureQuestionsShowingDIF, LRUQuestionsShowingDIF, LRNUQuestionsShowingDIF, LRUPureQuestionsShowingDIF, LRNUPureQuestionsShowingDIF)
              #  write.xlsx(AllResults, file = paste0(HomePath, "/DIF_Results.xlsx"), sheetName="AllResults", append=TRUE)
              #  write.xlsx(ReducedResults, file = paste0(HomePath, "/DIF_Results.xlsx"), sheetName="ReducedResults", append=TRUE)
            })
              ####Results####
              
              output$DIFItems <- renderTable(DIFItems)
              
              toggleState("report")
              
              withProgress(message = 'Preparing Data for Export', value = 0, {
                ##Significance and Odds Ratios (Full Results)
                FullResults=cbind(MHFullResults, MHPureFullResults, Lordp, LordPurep, LRUFullResults, LRUPureFullResults, LRNUFullResults, LRNUPureFullResults)
                FullResults <- format(FullResults, digits=3, scientific = TRUE)
                FullResults=cbind(QNumbers, FullResults)
                ####Results####
              })
              
####Always Exported####
              output$report <- downloadHandler(
                filename = "DIF Report.xlsx",
                content = function(file) {
                  wb <- createWorkbook()
                  
                  addWorksheet(wb, "Result Preview")
                  writeData(wb, "Result Preview", DIFItems, startCol = 1, startRow = 1, rowNames = FALSE)
                  
                  addWorksheet(wb, "Numeric Results")
                  writeData(wb, "Numeric Results", FullResults, startCol = 1, startRow = 1, rowNames = FALSE)
                  ####Always Exported####
                  
                  ####IRT Parameters#####
                  withProgress(message = 'Extracting IRT Parameters', value = 0.33, {
                    Questions=c(QNumbers,QNumbers)
                    Group=c(rep(1, length(QNumbers)),rep(0, length(QNumbers)))
                    IRTPar=cbind(Questions, Group, IRTPar)
                    
                    addWorksheet(wb, "IRT Parameters")
                    writeData(wb, "IRT Parameters", IRTPar, startCol = 1, startRow = 1, rowNames = FALSE)
                  })
                  saveWorkbook(wb, file, overwrite = TRUE)
                  ####IRT Parameters#####
                  
                  ####Item Plots####
                  if (identical(input$ItemPlots, TRUE)) {
                    
                  withProgress(message = 'Generating Item Plots', value = 0, {
                    ScoreSUM=ExcelScore
                    ScoreSUM=scale(ScoreSUM)
                    
                    zscore=scale(ScoreSUM)
                    bins=zscore
                    bins[bins < -1.49] <- 99      #STUPID R CODE WONT LET ME CONVERT STRAIGHT TO TEXT, WILL CONVERT OVER LATER
                    bins[bins < -.99] <- 98
                    bins[bins < -.49] <- 97
                    bins[bins < -.24] <- 96
                    bins[bins < .01] <- 95
                    bins[bins < .26] <- 94
                    bins[bins < .51] <- 93
                    bins[bins < 1.01] <- 92
                    bins[bins < 1.51] <- 91
                    bins[bins < 20] <- 90
                    Merged=cbind(ExcelQuestions, ExcelGrouping, bins)
                    BinNames=rbind("-1.50 or Less", "-1.49 to -1.00", "-0.99 to -0.50", "-0.49 to -0.25", "-0.24 to 0.00", "0.01 to 0.25", "0.26 to 0.50", "0.51 to 1.00", "1.01 to 1.50", "1.51 or More","-1.50 or Less", "-1.49 to -1.00", "-0.99 to -0.50", "-0.49 to -0.25", "-0.24 to 0.00", "0.01 to 0.25", "0.26 to 0.50", "0.51 to 1.00", "1.01 to 1.50", "1.51 or More")
                    Grouping=rbind("Group 0","Group 0","Group 0","Group 0","Group 0","Group 0","Group 0","Group 0","Group 0","Group 0","Group 1","Group 1","Group 1","Group 1","Group 1","Group 1","Group 1","Group 1","Group 1","Group 1")  #0=Group 0, 1=Group 1
                    
                    BinColumn=Numberofquestions+2
                    GroupingColumn=Numberofquestions+1
                    
                    
                    for (i in 1:Numberofquestions) {
                      Bin991=sum(Merged[which(Merged[,BinColumn]=="99" & Merged[,GroupingColumn]==1 ),i])/sum(ifelse(Merged[,BinColumn]=="99" & Merged[,GroupingColumn]==1,1,0))
                      Bin990=sum(Merged[which(Merged[,BinColumn]=="99" & Merged[,GroupingColumn]==0 ),i])/sum(ifelse(Merged[,BinColumn]=="99" & Merged[,GroupingColumn]==0,1,0))
                      Bin981=sum(Merged[which(Merged[,BinColumn]=="98" & Merged[,GroupingColumn]==1 ),i])/sum(ifelse(Merged[,BinColumn]=="98" & Merged[,GroupingColumn]==1,1,0))
                      Bin980=sum(Merged[which(Merged[,BinColumn]=="98" & Merged[,GroupingColumn]==0 ),i])/sum(ifelse(Merged[,BinColumn]=="98" & Merged[,GroupingColumn]==0,1,0))
                      Bin971=sum(Merged[which(Merged[,BinColumn]=="97" & Merged[,GroupingColumn]==1 ),i])/sum(ifelse(Merged[,BinColumn]=="97" & Merged[,GroupingColumn]==1,1,0))
                      Bin970=sum(Merged[which(Merged[,BinColumn]=="97" & Merged[,GroupingColumn]==0 ),i])/sum(ifelse(Merged[,BinColumn]=="97" & Merged[,GroupingColumn]==0,1,0))
                      Bin961=sum(Merged[which(Merged[,BinColumn]=="96" & Merged[,GroupingColumn]==1 ),i])/sum(ifelse(Merged[,BinColumn]=="96" & Merged[,GroupingColumn]==1,1,0))
                      Bin960=sum(Merged[which(Merged[,BinColumn]=="96" & Merged[,GroupingColumn]==0 ),i])/sum(ifelse(Merged[,BinColumn]=="96" & Merged[,GroupingColumn]==0,1,0))
                      Bin951=sum(Merged[which(Merged[,BinColumn]=="95" & Merged[,GroupingColumn]==1 ),i])/sum(ifelse(Merged[,BinColumn]=="95" & Merged[,GroupingColumn]==1,1,0))
                      Bin950=sum(Merged[which(Merged[,BinColumn]=="95" & Merged[,GroupingColumn]==0 ),i])/sum(ifelse(Merged[,BinColumn]=="95" & Merged[,GroupingColumn]==0,1,0))
                      Bin941=sum(Merged[which(Merged[,BinColumn]=="94" & Merged[,GroupingColumn]==1 ),i])/sum(ifelse(Merged[,BinColumn]=="94" & Merged[,GroupingColumn]==1,1,0))
                      Bin940=sum(Merged[which(Merged[,BinColumn]=="94" & Merged[,GroupingColumn]==0 ),i])/sum(ifelse(Merged[,BinColumn]=="94" & Merged[,GroupingColumn]==0,1,0))
                      Bin931=sum(Merged[which(Merged[,BinColumn]=="93" & Merged[,GroupingColumn]==1 ),i])/sum(ifelse(Merged[,BinColumn]=="93" & Merged[,GroupingColumn]==1,1,0))
                      Bin930=sum(Merged[which(Merged[,BinColumn]=="93" & Merged[,GroupingColumn]==0 ),i])/sum(ifelse(Merged[,BinColumn]=="93" & Merged[,GroupingColumn]==0,1,0))
                      Bin921=sum(Merged[which(Merged[,BinColumn]=="92" & Merged[,GroupingColumn]==1 ),i])/sum(ifelse(Merged[,BinColumn]=="92" & Merged[,GroupingColumn]==1,1,0))
                      Bin920=sum(Merged[which(Merged[,BinColumn]=="92" & Merged[,GroupingColumn]==0 ),i])/sum(ifelse(Merged[,BinColumn]=="92" & Merged[,GroupingColumn]==0,1,0))
                      Bin911=sum(Merged[which(Merged[,BinColumn]=="91" & Merged[,GroupingColumn]==1 ),i])/sum(ifelse(Merged[,BinColumn]=="91" & Merged[,GroupingColumn]==1,1,0))
                      Bin910=sum(Merged[which(Merged[,BinColumn]=="91" & Merged[,GroupingColumn]==0 ),i])/sum(ifelse(Merged[,BinColumn]=="91" & Merged[,GroupingColumn]==0,1,0))
                      Bin901=sum(Merged[which(Merged[,BinColumn]=="90" & Merged[,GroupingColumn]==1 ),i])/sum(ifelse(Merged[,BinColumn]=="90" & Merged[,GroupingColumn]==1,1,0))
                      Bin900=sum(Merged[which(Merged[,BinColumn]=="90" & Merged[,GroupingColumn]==0 ),i])/sum(ifelse(Merged[,BinColumn]=="90" & Merged[,GroupingColumn]==0,1,0))
                      Percents1=rbind(Bin991,Bin981,Bin971,Bin961,Bin951,Bin941,Bin931,Bin921,Bin911,Bin901)
                      Percents0=rbind(Bin990,Bin980,Bin970,Bin960,Bin950,Bin940,Bin930,Bin920,Bin910,Bin900)
                      Percents=rbind(Percents0,Percents1)
                      PlotData=data.frame(cbind(Grouping,BinNames,Percents))
                      Groups=PlotData[,1]
                      PlotData[,3] <- as.numeric(as.character(PlotData[,3]))  #Whoever coded ggplot doesn't know to difference between categorical and continuous
                      plot<-ggplot(PlotData, aes(x=reorder(PlotData[,2],rbind(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20)), PlotData[,3], group=PlotData[,1])) +
                        geom_line(aes(linetype=Groups))+
                        geom_point()+
                        scale_linetype_manual(values=c("twodash", "solid"))
                      plot<-plot + labs(title=paste0(colnames(ExcelQuestions)[i], " Item Plot"), y="Percent Correct", x = "Z-Score of Test Score")
                      plot <- plot + ylim(0,1)
                      
                      XXX=tempdir()
                      TemporaryFile=paste0(XXX,"\\Q",i,".png")
                      png(TemporaryFile, height=500, width=1200)
                      print(plot)
                      dev.off()
                    }
                    
                    
                    sheet <- addWorksheet(wb, sheetName="Item Plots")
                    Row=-25
                    for (i in 1:Numberofquestions) {
                      Row=Row+26
                      TemporaryFile=paste0(XXX,"\\Q",i,".png")
                      insertImage(wb, sheet, TemporaryFile, width=3736.526946, height=1556.886228, startRow = Row, startCol = 1, units = "px")
                    }
                    
                    saveWorkbook(wb, file, overwrite = TRUE)
                  })}
                  ####Item Plots####

                    
                    ####IRT Model Fits####
                  if (identical(input$FitInput, TRUE)) {
                    withProgress(message = 'Assessing Model Fit', value = 0, {
                      
                      ItemAnalysisRightWrong=cbind(ExcelQuestions,ExcelGrouping)
                      colnames(ItemAnalysisRightWrong) = c(colnames(ExcelQuestions), input$Grouping)
                      
                      Group0=ItemAnalysisRightWrong[ItemAnalysisRightWrong[input$Grouping] == 0, ]
                      Group0[input$Grouping]=NULL
                      Group1=ItemAnalysisRightWrong[ItemAnalysisRightWrong[input$Grouping] == 1, ]
                      Group1[input$Grouping]=NULL
                      
                      Group0IRT=est(Group0, model = "2PL", engine = "ltm")
                      Group1IRT=est(Group1, model = "2PL", engine = "ltm")
                    })
                    
                    #Item Fit
                    withProgress(message = 'Assessing Model Fit', value = 0.33, {
                      
                      
                      E00="Item-model fit using the likelihood-ratio statistic"
                      Fits0=c("Likelihood-Ratio Statistic", "Degrees of Freedom", "P-Value")
                      E0 <- tryCatch({
                        for (i in 1:ncol(Group0)) {
                          Temp=itf(Group0, Group0IRT, i, do.plot = FALSE)
                          Fits0=rbind2(Fits0,Temp)
                        }
                        Temp2=nrow(Fits0[which(Fits0[,3]<=0.05),])
                        paste("There were", Temp2, "questions in group 0 with test item fit <0.05 based on the likelihood-ratio statistic (likelihood of significance inflated with large samples though so interpret carefully)")},
                        error=function(cond) {return("Error, cannot calculate item fit")})
                      
                      Fits1=c("Likelihood-Ratio Statistic", "Degrees of Freedom", "P-Value")
                      E1 <- tryCatch({
                        for (i in 1:ncol(Group1)) {
                          Temp=itf(Group1, Group1IRT, i, do.plot = FALSE)
                          Fits1=rbind2(Fits1,Temp)
                        }
                        Temp2=nrow(Fits1[which(Fits1[,3]<=0.05),])
                        paste("There were", Temp2, "questions in group 1 with test item fit <0.05 based on the likelihood-ratio statistic (likelihood of significance inflated with large samples though so interpret carefully)")},
                        error=function(cond) {return("Error, cannot calculate item fit")})
                      
                      E2=""
                    })
                    
                    
                    #Item Independence
                    withProgress(message = 'Assessing Model Fit', value = 0.66, {
                      Group0YenQ3=Yen.Q3(Group0, IRT.model = "2pl")
                      Group1YenQ3=Yen.Q3(Group1, IRT.model = "2pl")
                      G0YenQ3Abs=abs(Group0YenQ3[["Q3"]])
                      G0YenQ3AbsW=abs(Group0YenQ3[["Q3.weighted"]])
                      G1YenQ3Abs=abs(Group1YenQ3[["Q3"]])
                      G1YenQ3AbsW=abs(Group1YenQ3[["Q3.weighted"]])
                      
                      E3="Yens Q3 test for local independence (want values below 0.2)"
                      E4=paste("In group 0,there are", sum(G0YenQ3Abs[upper.tri(G0YenQ3Abs)] >= 0.2, na.rm = T), "Yen Q3 statistics which exceed the absolute value of 0.2.", sep = " ")
                      E5=paste("In group 0,there are", sum(G0YenQ3AbsW[upper.tri(G0YenQ3AbsW)] >= 0.2, na.rm = T), "Yen Q3 weighted statistics which exceed the absolute value of 0.2.", sep = " ")
                      E6=paste("In group 1,there are", sum(G1YenQ3Abs[upper.tri(G1YenQ3Abs)] >= 0.2, na.rm = T), "Yen Q3 statistics which exceed the absolute value of 0.2.", sep = " ")
                      E7=paste("In group 1,there are", sum(G1YenQ3AbsW[upper.tri(G1YenQ3AbsW)] >= 0.2, na.rm = T), "Yen Q3 weighted statistics which exceed the absolute value of 0.2.", sep = " ")
                      E8=""
                    })
                    
                    #Parameter Invariance (WIP)
                    withProgress(message = 'Assessing Model Fit', value = 1, {
                      E09="Parameter invariance based on splitting each set into two parts and evaluating the correlation between the item parameters"
                      set.seed(5)
                      Group0RNG=round(runif(nrow(Group0), 0, 1), 0)
                      Group0AB=cbind(Group0,Group0RNG)
                      Group0A=Group0AB[which(Group0AB[,"Group0RNG"]=="0"),]
                      Group0A=Group0A[, ! names(Group0A) %in% c("Group0RNG")]
                      Group0B=Group0AB[which(Group0AB[,"Group0RNG"]=="1"),]
                      Group0B=Group0B[, ! names(Group0B) %in% c("Group0RNG")]
                      
                      Group0AIRT=est(Group0A, model = "2PL", engine = "ltm")
                      Group0BIRT=est(Group0B, model = "2PL", engine = "ltm")
                      
                      Group0ADisCor=cor.test(Group0AIRT$est[,1], Group0BIRT$est[,1])
                      E9=paste("The p value for the discrimination correlation between randomly split groups within Group 0 is", format(Group0ADisCor[["p.value"]], digits=3, scientific = TRUE))
                      #plot(Group0AIRT$est[,1], Group0BIRT$est[,1], xlab = "Group 0 A Discrimination", ylab = "Group 0 B Discrimination")
                      Group0ADifCor=cor.test(Group0AIRT$est[,2], Group0BIRT$est[,2])
                      E10=paste("The p value for the difficulty correlation between randomly split groups within Group 0 is", format(Group0ADifCor[["p.value"]], digits=3, scientific = TRUE))
                      #plot(Group0AIRT$est[,2], Group0BIRT$est[,2], xlab = "Group 0 A Difficulty", ylab = "Group 0 B Difficulty")
                      
                      set.seed(6)
                      Group1RNG=round(runif(nrow(Group1), 0, 1), 0)
                      Group1AB=cbind(Group1,Group1RNG)
                      Group1A=Group1AB[which(Group1AB[,"Group1RNG"]=="0"),]
                      Group1A=Group1A[, ! names(Group1A) %in% c("Group1RNG")]
                      Group1B=Group1AB[which(Group1AB[,"Group1RNG"]=="1"),]
                      Group1B=Group1B[, ! names(Group1B) %in% c("Group1RNG")]
                      
                      Group1AIRT=est(Group1A, model = "2PL", engine = "ltm")
                      Group1BIRT=est(Group1B, model = "2PL", engine = "ltm")
                      
                      Group1ADisCor=cor.test(Group1AIRT$est[,1], Group1BIRT$est[,1])
                      E11=paste("The p value for the discrimination correlation between randomly split groups within Group 1 is", format(Group1ADisCor[["p.value"]], digits=3, scientific = TRUE))
                      #plot(Group1AIRT$est[,1], Group1BIRT$est[,1], xlab = "Group 1 A Discrimination", ylab = "Group 1 B Discrimination")
                      Group1ADifCor=cor.test(Group1AIRT$est[,2], Group1BIRT$est[,2])
                      E12=paste("The p value for the difficulty correlation between randomly split groups within Group 1 is", format(Group1ADifCor[["p.value"]], digits=3, scientific = TRUE))
                      #plot(Group1AIRT$est[,2], Group1BIRT$est[,2], xlab = "Group 1 A Difficulty", ylab = "Group 1 B Difficulty")
                    })
                    
                    #Combination of 
                    FitStats=c(E00, E0, E1, E2, E3, E4, E5, E6, E7, E8, E09, E9, E10, E11, E12)
                    FitStats=t(t(FitStats))
                    FitStats=as.data.frame(FitStats)
                    colnames(FitStats)=c("Fit Statistics for the Baseline IRT Model")
                    
                    addWorksheet(wb, "IRT Fit")
                    writeData(wb, "IRT Fit", FitStats, startCol = 1, startRow = 1, rowNames = FALSE)
                    saveWorkbook(wb, file, overwrite = TRUE)
                  } #IF CALCULATE FIT
                    ####IRT Model Fits####
                    
}) #DOWLOAD HANDLE
}}})}))
shinyApp(ui, server)
