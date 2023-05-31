library(shiny)
library(readxl)
library(DT)
library("writexl")
library("data.table")  
library(tidyverse)
library(readr)
library(googledrive)
library(googlesheets4)

options(gargle_oauth_cache = ".secrets",
        gargle_oauth_email = "cdeissds19@gmail.com"
)

googledrive::drive_auth()  
googlesheets4::gs4_auth()

a<-drive_ls("LISTADO DE CURSOS")

lcursos<-read_sheet(ss = a$id[1])

mujeres <- read_excel("mujeres.xlsx")
mujeres$Sexo <- "M"

options(shiny.maxRequestSize = 100*1024^2)
lcursos <- lcursos[,c(1:5)]


shinyUI(
  # Define UI for application that draws a histogram
  ui <- 
    # Application title
    navbarPage("Cursos Virtuales", tabPanel("AFE",fluidPage(
      titlePanel("Aplicativo de Formato Estandar"),
      sidebarLayout(
        sidebarPanel(
          fileInput('file1', 'Calificaciones',
                    accept = c(".xlsx")),
          fileInput('cert', 'Certificado',
                    accept = c(".xlsx")),
          # fileInput('ultm', 'Ultimo Acceso',
          #           accept = c(".xlsx")),
          fileInput('adjx', 'Ajustes',
                    accept = c(".xlsx")),
          selectInput("type","Tipo",choices = c("Informe Global",
                                                "Informe Modular",
                                                "Tabla"),selected = "Tabla"),
          selectInput("name","Nombre  Curso",choices = lcursos$ABREVIATURA),
          # selectInput(inputId = "inSelect", "Seleccione la Cohorte",
          #                 choices = c("")),
          # numericInput("cohorte","Cohorte",min = 0,max = NA, value = 5),
          dateInput("date1", "Fecha Inicio:", value = Sys.Date()),
          dateInput("date2", "Fecha Fin:", value = Sys.Date()),
          actionButton("go", "update"),
          downloadButton("downloadData", "Descargar Tabla")),
          mainPanel(
          dataTableOutput('my_output_data'))
      )
    )),
    tabPanel("Cursos Especiales",fluidPage(
      titlePanel("Aplicativo de Cursos Especiales"),
      sidebarLayout(
        sidebarPanel(
          fileInput('file_espx', 'Encuesta inicial',
                    accept = c(".txt")),
          fileInput('cert2', 'Constancia de participacion',
                    accept = c(".xlsx")),
          actionButton("go2", "update"),
          downloadButton("downloadData2", "Descargar Tabla")),
        mainPanel(dataTableOutput('table'))
    )))))
    
    server = function(input, output,session){

      certx <- reactive({
        inputFile <- input$cert
        if (is.null(inputFile))
          return()
        read_excel(inputFile$datapath)
      })
      
      # ultmx <- reactive({
      #   inputFile <- input$ultm
      #   if (is.null(inputFile))
      #     return()
      #   read_excel(inputFile$datapath, 
      #              col_types = c("skip", "skip", "skip", 
      #                            "text", "skip", "skip", "skip", "skip", 
      #                            "skip", "skip", "skip", "skip", "skip", 
      #                            "text")) 
      
      # })
      
      calx <- reactive({
        inputFile <- input$file1
        if (is.null(inputFile))
          return()
        read_excel(inputFile$datapath)
      })
      
      adjx <- reactive({
        inputFile <- input$adjx
        if (is.null(inputFile))
          return()
        read_excel(inputFile$datapath)
      })
      
      file_espx <- reactive({
        inputFile <- input$file_espx
        if (is.null(inputFile))
          return()
        file_espx<- read_delim(inputFile$datapath, 
                    delim = "\t", escape_double = FALSE, 
                    trim_ws = TRUE)

        file_espx$nombrecompleto <- paste(toupper(file_espx$`Nombre completo`), file_espx$Grupo)
        
        file_espx
        
      })
      
      cert2 <- reactive({
        inputFile <- input$cert2
        if (is.null(inputFile))
          return()
        #print(inputFile)
        cert2x<- read_excel(inputFile$datapath)
        
        cert2x$Grupo[is.na(cert2x$Grupo)]<-"(Personas que no están en ningún grupo)"
        
        cert2x$nombrecompleto<-paste(toupper(cert2x$Nombre),toupper(cert2x$`Apellido(s)`),cert2x$Grupo)
        
        
        
        cert2x$Certificado <- "SI"
        cert2x
        
      })
      
      mergedData2<-reactive({
        if (is.null(cert2()))
            return()
        df <- merge(file_espx(),cert2(),by = "nombrecompleto", all = TRUE)
        df
      })

      mergedData <- reactive({
        if (is.null(adjx()))
          return()
        calx <- calx()

        adjx <-adjx()
        
        df<- merge(calx,adjx,by.x="Número de ID",by.y = "DOCUMENTO",all.x = TRUE)

        df$`Número de ID`[df$`Número de ID` != df$DOCUMENTO & 
                            is.na(df$DOCUMENTO) == FALSE] <- df$DOCUMENTO[
                              df$`Número de ID` != df$DOCUMENTO & is.na(df$DOCUMENTO) == FALSE]
        
        df$`Número de ID` <- as.character(df$`Número de ID`)
        df$`INSTITUCIÓN`[is.na(df$`INSTITUCIÓN`)] <-"NINGUNO"
        df$`Institución`[is.na(df$`Institución`)] <-"NINGUNO"
        
        df$`Institución`[df$`Institución` != df$`INSTITUCIÓN`] <-  df$`INSTITUCIÓN`[df$`Institución` != df$`INSTITUCIÓN`]
        
        df$`CORREO ELECTRONICO`[is.na(df$`CORREO ELECTRONICO`)] <-"NINGUNO"
        df$`Dirección de correo`[is.na(df$`Dirección de correo`)] <-"NINGUNO"
        
        df$`Dirección de correo`[df$`Dirección de correo` != df$`CORREO ELECTRONICO`] <-df$`CORREO ELECTRONICO`[df$`Dirección de correo` != df$`CORREO ELECTRONICO`]
        
        df$Profesion[is.na(df$Profesion)] <-"NINGUNO"
        df$`PROFESIÓN`[is.na(df$`PROFESIÓN`)] <-"NINGUNO"
        
        df$Profesion[df$Profesion != df$`PROFESIÓN`] <- df$`PROFESIÓN`[df$Profesion != df$`PROFESIÓN`]
        
        df <- df[,colnames(calx())]
        
        cert <- certx()[,c(3,7,9)]
        cert$`Descargó_Certific` <- "NO"
        cert$`Descargó_Certific`[is.na(cert[,3]) == FALSE] <- "SI"
        colnames(cert)<-c("ID","Ultimo Acceso","Fecha De Descarga Certificado","Descargó_Certific")
        #certm<- merge(cert,adjx(),by.x="ID",by.y = "DOCUMENTO",all.y = TRUE)
        # z <- merge(y=ult,x=cert,by="ID",all=TRUE)
        # z$`Descargó_Certific`[z$`Descargó_Certific`!="SI"]<- "NO"
        colx <- c("ID","Fecha De Descarga Certificado","Ultimo Acceso","Descargó_Certific")
        z<-cert[,colx] 
        colnames(z)[1]<-"Número de ID"
        nc <- which(colnames(df) == "Última descarga de este curso")
        merge(x=df[,-nc],y = z,by = "Número de ID", all.x = TRUE)
        
      })
      
      # observe({
      #   updateSelectInput(session, "inSelect",
      #                     choices = unique(certx()$grupo)
      #   )})
      
      Cohorte <- reactive({
        basec<-certx()
      if (basec$idcurso[1] == "253" | basec$idcurso[1] == "255"  ){
          basec$grupo <- "Cohorte Permanente"
      }

     if (is.null(basec$grupo)[1] == FALSE){
          x1<-scan(text = str_to_title(str_split(basec$grupo[1],"-")[[1]][1]), what = "")
          if (length(x1) == 2| length(x1) == 3) {
          a<-paste(x1[1],x1[2])
          }
          if (length(x1) == 1){ 
          x1<-str_split(basec$grupo[1],"-")
          a<-paste(x1[[1]][1],x1[[1]][2])
          }
     }
        a
      })
    
      
      rv <- reactiveValues(my_data = NULL)
      
      observe({
        rv$my_data <- mergedData()
      })
      
      observeEvent(input$go, {
        
        rv$my_data$`Técnico` <- lcursos$`TÉCNICO A CARGO`[lcursos$ABREVIATURA == input$name]
        rv$my_data$`Nombre Del Curso` <- input$name
        rv$my_data$Cohorte <- Cohorte()
        rv$my_data$`Fecha Inicio` <- format(input$date1,'%d/%m/%Y')
        rv$my_data$`Fecha Fin` <- format(input$date2,'%d/%m/%Y')

        rv$my_data$`Módulos 1`<-NA
        rv$my_data$`Módulos 2`<-NA
        rv$my_data$`Módulos 3`<-NA
        rv$my_data$`Módulos 4`<-NA
        rv$my_data$`Módulos 5`<-NA
        rv$my_data$`Módulos 6`<-NA
        rv$my_data$`Módulos 7`<-NA
        rv$my_data$`Módulos 8`<-NA
        rv$my_data$`Módulos 9`<-NA
        rv$my_data$`Módulos 10`<-NA
        rv$my_data$`Módulos 11`<-NA
        apro <- lcursos$`NOTA DE APROBACIÓN`[lcursos$ABREVIATURA == input$name] %>% as.numeric()
        rv$my_data$`Nota Para Aprobar Curso` <- as.numeric(apro)

        colsx<-c("col1","col2","col3","col4",
                 "col5","col6","col7","col8",
                 "col9","col10","col11")
        
#        rv$my_data$`Descargó_Certific`[is.na(rv$my_data$`Descargó_Certific`)== TRUE]<-"NO" 
        
        numb1x<-which(colnames(rv$my_data) == "Observacion") + 1
        numb2x<-which(colnames(rv$my_data) == "Descargó_Certific") -3
        
        mm <-1
        for (m in numb1x:numb2x){
          rv$my_data[,m] <-gsub(pattern = "-","0",x = rv$my_data[,m])
          rv$my_data[,m]<-as.numeric(rv$my_data[,m])
          colnames(rv$my_data)[m]<-paste0("col",mm)
          mm <- mm + 1
        }

        rv$my_data$na_count <- apply(rv$my_data[,c(numb1x,numb2x)],
                                     1, function(x) sum(x == 0))
        
        for (i in 1:11){
          if (colsx[i] %in% colnames(rv$my_data)){
            if (is.numeric(rv$my_data[,colsx[i]]) == FALSE) {
              index<-rv$my_data[,colsx[i]] == "-"
              rv$my_data[index == TRUE, colsx[i]] <- ""
              rv$my_data[,colsx[i]] <- lapply(rv$my_data[,colsx[i]], as.numeric) %>% as.data.frame()
              rv$my_data[,paste("Módulos",i)]<-rv$my_data[,colsx[i]]
              number<-which(colsx[i] == colnames(rv$my_data))
              rv$my_data<- rv$my_data[,-number]
            } else{
            rv$my_data[,paste("Módulos",i)]<-rv$my_data[,colsx[i]]
            number<-which(colsx[i] == colnames(rv$my_data))
            rv$my_data<- rv$my_data[,-number]
            }
          }
        }

           number<-which("Módulos 1" == colnames(rv$my_data))
           number2 <- 10 + number
           rv$my_data$`Nota Mínima` <- apply(rv$my_data[,c(number:number2)], 1,
                                             FUN=min,na.rm = TRUE)
           
           rv$my_data$`Nota Mínima`[rv$my_data$`Nota Mínima` == Inf] <- 0

           rv$my_data$`Nota Promedio` <- rowMeans(rv$my_data[,number:number2],na.rm = TRUE)
           
              
           rv$my_data$`Módulos Perdidos` <- rowSums(rv$my_data[,number:number2] < as.numeric(apro),
                                                    na.rm =TRUE)

           rv$my_data$`Módulos Ganados` <- rowSums(is.na(rv$my_data[,number:number2]) != TRUE ) - rv$my_data$`Módulos Perdidos`

           colnames(rv$my_data)[2] <- "Nombre(S)"
           colnames(rv$my_data)[3] <- "Apellido(S)"
           colnames(rv$my_data)[1] <- "Número De Id"
           colnames(rv$my_data)[6] <- "Dirección De Correo"
           colnames(rv$my_data)[7] <- "Profesión"
           colnames(rv$my_data)[10] <- "Teléfono"
           colnames(rv$my_data)[11] <- "Observación"

           colnames(rv$my_data)[which(colnames(rv$my_data) == "Descargo_Certific")] <- "Descargo Certificado"
           colnames(rv$my_data)[which(colnames(rv$my_data) == "Descargó_Certific")] <- "Descargo Certificado"
           colnames(rv$my_data)[which(colnames(rv$my_data) =="Ultimo Acceso")] <- "Ultimo Acceso"
           #colnames(rv$my_data)[which(colnames(rv$my_data) =="UltimoAcceso")] <- "Ultimo Acceso"
           #colnames(rv$my_data)[which(colnames(rv$my_data) =="orden")] <- "Orden"
           
          
           rv$my_data$Aprobo <- ifelse(rv$my_data$`Módulos Perdidos` == 0, yes =  "SI",no = "NO")
           #rv$my_data$Aprobo[is.na(rv$my_data$Aprobo)] <-"NO"
           
          rv$my_data$`Ultimo Acceso`[is.na(rv$my_data$`Ultimo Acceso`) != TRUE] <- "Ingreso"
          
          rv$my_data$Orden[is.na(rv$my_data$`Ultimo Acceso`) == TRUE & is.na(rv$my_data$`Nota Promedio`) != TRUE] <- "No ha Ingresado"
          
          rv$my_data$Orden[rv$my_data$Aprobo == "NO" & is.na(rv$my_data$`Ultimo Acceso`) != TRUE ] <- "No Aprobo"
          
          rv$my_data$Orden[rv$my_data$`Nota Promedio` == 0 &
                             is.na(rv$my_data$`Ultimo Acceso`) != TRUE ] <- "Ingreso y No participo"
          
          rv$my_data$Orden[rv$my_data$Aprobo == "NO" & 
                             rv$my_data$`Nota Mínima` !=  0] <- "Presento todo y no Aprobo"
           
           rv$my_data$Orden[(rv$my_data$Aprobo == "SI") & 
                              (rv$my_data$`Descargo Certificado` == "NO")] <- "Aprobo y No Descargo"
           
           rv$my_data$Orden[(rv$my_data$Aprobo == "SI") &
                              is.na(rv$my_data$`Descargo Certificado` == TRUE)] <- "Aprobo y No Descargo"
           
           rv$my_data$Orden[rv$my_data$Aprobo == "SI" & 
                              rv$my_data$`Descargo Certificado` == "SI"] <- "Aprobo y Descargo"
           
            
           rv$my_data$`Profesional` <- lcursos$`PROFESIONAL A CARGO`[lcursos$ABREVIATURA == input$name]
          
           col_order <- c("Orden","Nombre Del Curso","Técnico","Cohorte","Profesional",
           "Nombre(S)","Apellido(S)","Número De Id","Institución",
           "Departamento","Dirección De Correo","Profesión","Grupo","Cargo",
           "Teléfono","Observación","Nota Para Aprobar Curso",
           "Módulos 1",	"Módulos 2",
           "Módulos 3",	"Módulos 4",	"Módulos 5",	"Módulos 6",
           "Módulos 7",	"Módulos 8",	"Módulos 9",	"Módulos 10",
           "Módulos 11",	"Nota Mínima",	"Nota Promedio",	"Módulos Perdidos",
           "Módulos Ganados",	"Descargo Certificado",
           "Ultimo Acceso", "Aprobo",
           		"Fecha Inicio",	"Fecha Fin","Sexo","Fecha De Descarga Certificado")
          
         rv$my_data$Sexo1 <- lapply(rv$my_data[,"Nombre(S)"], 
                                   function(x) ifelse(any(strsplit(x," ")[[1]]
                                    %in% mujeres$MUJERES),"M","H"))
         
         rv$my_data$Sexo<-as.character(rv$my_data$Sexo1)
         rv$my_data <- rv$my_data[,col_order]
         
        })
      
      output$my_output_data <- renderDataTable({
        
        if (input$type == "Informe Global"){
          a<-rv$my_data %>% count(Aprobo,Orden) %>% group_by(Aprobo)
          b <- sum(a[,"n"])
          a$Porcentaje <- round(a[,"n"]/b,4)*100
          c <- sum(a[a$Aprobo == "SI","n"])
          Aprobox <- data.frame("SI", "APROBADOS", c, round(c/b,4)*100)      
          names(Aprobox) <- c("Aprobo", "Orden", "n", "Porcentaje")  
          a<-rbind.data.frame(a,Aprobox) 
          
          nx<-nrow(rv$my_data[(is.na(rv$my_data$`Fecha De Descarga Certificado`) == FALSE),])
          d<-data.frame(Orden = "TOTAL",Aprobo = "DESCARGA CERTIFICADO",n = nx, Porcentaje = round(nx / b,4)*100)
          a <- rbind.data.frame(a,d)
        }
        
        if (input$type == "Informe Modular"){

          nn<-which(colnames(rv$my_data) == "Módulos 1")
          apro <- as.numeric(lcursos$`NOTA DE APROBACIÓN`[lcursos$ABREVIATURA == input$name])

          modx<-function(u){
          ma <- sum(ifelse(rv$my_data[,u] >= apro,1,0),na.rm = T)
          mb <- sum(ifelse(rv$my_data[,u] < apro & rv$my_data[,u] != 0 ,1,0),na.rm = T)
          mn <- sum(ifelse(is.na(rv$my_data[,u]) ==  T | rv$my_data[,u] == 0 ,1,0),na.rm = T)
          Modulo <- colnames(rv$my_data)[u]
          temp<-cbind.data.frame(Modulo,ma,mb,mn)
          colnames(temp)<-c("Modulo","Aprobados","No Aprobados","No Participa")
          return(temp)
          }
          
          a<-lapply(seq(nn,nn+10),modx) %>% rbindlist(.)

        }
        
        
        if (input$type == "Tabla"){
          a <-as.data.table(rv$my_data,keep.rownames = F)
        }
        a
        
        })
      

      output$downloadData <- downloadHandler(
        
        filename = function() {
          paste(paste("Reporte",input$name,input$cohorte), ".xlsx", sep = "")
        },
        content = function(file) {
          
          a<-rv$my_data %>% count(Aprobo,Orden) %>% group_by(Aprobo)
          b <- sum(a[,"n"])
          a <- a %>% 
            mutate(Porcentaje = round(n / b,4)*100)
          c <- sum(a[a$Aprobo == "SI","n"])
          Aprobox <- data.frame("SI", "APROBADOS", c, round(c/b,4)*100)
          names(Aprobox) <- c("Aprobo", "Orden", "n", "Porcentaje")
          Informe_Global<-rbind.data.frame(a,Aprobox)

          nx<-nrow(rv$my_data[is.na(rv$my_data$`Fecha De Descarga Certificado`) == FALSE,])
          d<-data.frame(Aprobo = "TOTAL",Orden = "DESCARGA CERTIFICADO",n = nx, Porcentaje = round(nx / b,4)*100)
          Informe_Global <- rbind.data.frame(Informe_Global,d)
          
          nn<-which(colnames(rv$my_data) == "Módulos 1")
          apro <- as.numeric(lcursos$`NOTA DE APROBACIÓN`[lcursos$ABREVIATURA == input$name])
          
          modx<-function(u){
            ma <- sum(ifelse(rv$my_data[,u] >= apro,1,0))
            mb <- sum(ifelse(rv$my_data[,u] < apro & rv$my_data[,u] != 0 ,1,0))
            mn <- sum(ifelse(is.na(rv$my_data[,u]) ==  T | rv$my_data[,u] == 0 ,1,0))
            Modulo <- colnames(rv$my_data)[u]
            temp<-cbind.data.frame(Modulo,ma,mb,mn)
            colnames(temp)<-c("Modulo","Aprobados","No Aprobados","No Participa")
            return(temp)
          }
          
          Informe_Modular<-lapply(seq(nn,nn+10),modx) %>% rbindlist(.)
          datos<-rv$my_data
          
          datos1<-datos[datos$Orden == "No ha Ingresado",]
          datos2<-datos[datos$Orden == "Aprobo y Descargo",]
          datos3<-datos[datos$Orden == "Aprobo y No Descargo",]
          datos4<-datos[datos$Orden == "Presento todo y no Aprobo",]
          datos5<-datos[datos$Orden == "Ingreso y No participo",]
          datos6<-datos[datos$Orden == "No Aprobo",]
          #datos7<-datos[datos$Orden == "Ingreso",]
          
          
          list_of_datasets <- list("Datos" = datos,"Informe Global" = Informe_Global, 
                                   "Informe_Modular" = Informe_Modular,
                                   "No ha Ingresado" = datos1,
                                   "Aprobo y descargo" = datos2,
                                   "Aprobo y no descargo" = datos3,
                                   "Presento todo y no aprobo" = datos4,
                                   "Ingreso y no participio" = datos5,
                                   "No aprobo" = datos6)
           #                        "Ingreso" = datos7)
          
          write_xlsx(list_of_datasets, file)
        })
      
      
      output$table <- renderDataTable({
        as.data.table(mergedData2())
        })
      
      
      output$downloadData2 <- downloadHandler(
        
        filename = function() {
          paste("Reporte Curso Especial.xlsx", sep = "")
        },
        content = function(file) {
          
          df<- mergedData2()
          
          df$duplicado <- "NO"
          df$duplicado[duplicated(df$nombrecompleto) == TRUE] <- "SI"
          
          df<-df[,c("Curso","Q01_Nombre(s)",
                    "Q00_Apellido(s)",
                    "Q02_Tipo de documento de identidad",
                    "Q03_Nro. de documento de identidad",
                    "Institución.x","Departamento.x",
                    "Q08_Correo electrónico","Teléfono","Certificado",
                    "Q10_Sexo","duplicado")]
          
          
          
          colnames(df)<-c("Curso","Nombre(s)","Apelllido(s)","Tipo de documento",
                          "ID","Institución","Departamento","Correo","Telefono",
                          "Certificado","Sexo","duplicado")
          
          df$Certificado[which(is.na(df$Certificado))]<-"NO"
          
          
          tabla <- df %>% group_by(Certificado,duplicado) %>% count()
          
          
      
        lista_files <- list("Datos" = df, "Resumen" = tabla)
          
          
          
          writexl::write_xlsx(lista_files, file)
    })
    }
    
shinyApp(ui = ui, server = server)






