if(!require(pacman)){
  
  install.packages("pacman")
}


pacman::p_load(readxl,
               tidyr,
               stringr,
               openxlsx,
               dplyr,
               janitor,
               lubridate,
               data.table)

library(data.table)
library(lubridate)
library(janitor)
library(openxlsx)
library(dplyr)
library(readxl)
library(tidyr)
library(stringr)

#  Functions --------

replace_na_for_numeric_columns <- function(df){
  
  for(j in seq_along(df)){
    data.table::set(df, i = which(is.na(df[[j]]) & is.numeric(df[[j]])), j = j, value = 0)
  }
  
}

today <- format(Sys.Date(), "%d_%b") 


start_time <- Sys.time()
#  Loading data ------
input_full <- read_excel("ignore/premiumguest/Input_26.07.xlsx") 
master_data <- read_excel("ignore/premiumguest/Datos Maestros Generales_26_jul.xlsx") 
master_data[master_data == "Vigo en Festas"] <- "Vigo En Festas"
master_data[master_data == "Discoteca Nexus L Escala"] <- "Discoteca NEXUS L Escala"

replace_na_for_numeric_columns(input_full)
replace_na_for_numeric_columns(master_data)

#  Selecting observation for testing ----- 
# input <- input_full[c(1:2, 25,26,27, 210),]
input <- input_full
#  Adjusting variables type -----

cols_char_md <- c("NEGOCIO_2")

master_data[cols_char_md] <- sapply(master_data[cols_char_md],as.character)

cols_char_input <- c("Nº Pedido","Promo Code", "UTM campaign...14",
                     "UTM campaign...45", "UTM medium", "Id KAM Negocio", "ID KAM WL")

input[cols_char_input] <- sapply(input[cols_char_input],as.character)

input <- input %>% 
  dplyr::mutate(negocio = tolower(Negocio),
                white_label = tolower(`White Label`))


master_data <- master_data %>% 
  dplyr::mutate(negocio = tolower(NEGOCIO_2),
                white_label = tolower(WL_1))



n_occur_negocio <- data.frame(table(master_data$negocio)) 

n_occur_negocio <-  n_occur_negocio %>% 
  dplyr::filter(Freq > 1) %>% 
  dplyr::rename(negocio = Var1,
                negocio_freq = Freq)

n_occur_white_label <- data.frame(table(master_data$white_label)) 

n_occur_white_label <-  n_occur_white_label %>% 
  dplyr::filter(Freq > 1) %>% 
  dplyr::rename(white_label = Var1,
                wl_freq = Freq)

master_data_check <- bind_rows(n_occur_negocio,n_occur_white_label)

get_output <- function(input){
  
  output <- input %>% 
    dplyr::left_join(master_data %>% 
                       dplyr::select(negocio,Grupo ),
                     by = c("negocio" )  ) %>% 
    dplyr::rename(GRUPO = Grupo) %>% 
    unique() %>% 
    dplyr::left_join(master_data %>% 
                       dplyr::select(white_label,`Es Canal de Venta` ) %>%  unique(),
                     by = c("white_label")  ) %>% 
    dplyr::mutate(`CANAL DE VENTA` = ifelse(`Es Canal de Venta` == "NO", "ORGANIZADOR","CANAL DE VENTA" )) %>% 
    dplyr::select(-`Es Canal de Venta`) %>% 
    dplyr::left_join(master_data %>% 
                       dplyr::select(white_label,`Es Canal de Venta` ) %>%  unique(),
                     by = c("white_label")  ) %>% 
    dplyr::mutate(`CANAL DE VENTA` = ifelse(`Es Canal de Venta` == "NO", "ORGANIZADOR","CANAL DE VENTA" )) %>% 
    dplyr::select(-`Es Canal de Venta`) %>% 
    dplyr::left_join(master_data %>% 
                       dplyr::select(negocio,Grupo )%>%  unique(),
                     by = c("negocio" )  ) %>% 
    dplyr::rename(`Negocio está en pestaña Datos Maestros Generales` = Grupo) %>% 
    unique() %>% 
    dplyr::left_join(master_data %>% 
                       dplyr::select(white_label, `Es Canal de Venta` )%>%  unique(),
                     by = c("white_label")  ) %>% 
    dplyr::rename(`WL está en pestaña Datos Maestros Generales` = `Es Canal de Venta` ) %>% 
    dplyr::left_join(master_data %>% 
                       dplyr::select(negocio, Liquidación_2 )%>%  unique(),
                     by = c("negocio")  ) %>% 
    dplyr::rename(PERIODICIDAD_LIQUIDACION = Liquidación_2 ) %>% 
    dplyr::left_join(master_data %>% 
                       dplyr::select(white_label, `TIPO CLIENTE` )%>%  unique(),
                     by = c("white_label")  ) %>% 
    dplyr::left_join(master_data %>% 
                       select(negocio,	`Comercial Negocio`)%>%  unique(),
                     by = c( "negocio")) %>%
    dplyr::left_join(master_data %>% 
                       dplyr::select(negocio,`Se liquida semanal?`)%>%  unique(),
                     by = c("negocio")) %>% 
    dplyr::rename(Remesa = `Se liquida semanal?`) %>%  
    dplyr::left_join(master_data %>% 
                       dplyr::select(negocio,AUTOFACTURA )%>%  unique(),
                     by = c("negocio")) %>% 
    dplyr::left_join(master_data %>%
                       dplyr::select(negocio,`Comercial WL`)%>%  unique(),
                     by = c( "negocio")) %>%
    # dplyr::rename(`Comercial WL`  = `Comercial Negocio.y`) %>%
    # dplyr::rename(`Comercial Negocio`  = `Comercial Negocio.x`) %>%
    unique()
  
  
  output <- output %>% 
    mutate(`Pagado banco >0` = ifelse(`Pagado Banco` > 0, "SI", "NO"),
           `FECHA COMPRA` = as.Date(format(`Fecha Compra`, format = "%m/%d/%Y"),"%m/%d/%Y"),
           `MES COMPRA` = match(months(as.POSIXlt(`FECHA COMPRA`, format="%m/%d/%Y")), month.name),
           `AÑO COMPRA` = format(as.POSIXlt(`FECHA COMPRA`, format="%m/%d/%Y"),format="%Y"),
           `FECHA EVENTO` = as.Date(format(`Fecha evento`, format = "%m/%d/%Y"),"%m/%d/%Y"),
           `MES EVENTO` = match(months(as.POSIXlt(`FECHA EVENTO`, format="%m/%d/%Y")), month.name),
           `AÑO EVENTO` = format(as.POSIXlt(`FECHA EVENTO`, format="%m/%d/%Y"),format="%Y"),
           DIA = format(as.POSIXlt(`FECHA EVENTO`, format="%m/%d/%Y"),format="%d"),
           SEMANA = paste("Semana",as.numeric(format(as.POSIXlt(`FECHA EVENTO`, format="%m/%d/%Y"),format="%V"))+1,sep="_"),
           `Ventas para Eventos fuera del mes de compra` = ifelse(`MES COMPRA` == `MES EVENTO`, "Evento mismo mes compra", "Compra Evento Futuro"),
           NEGOCIO_1 = Negocio,
           `GRUPO PROMOTOR` = case_when(
             substr(Negocio,1,8) == "Twenties" & `Grupo PR_FOR` == "YouBarcelona" ~ "YouBarcelona",
             `CANAL DE VENTA` == "CANAL DE VENTA" ~ "Red Premiumguest",
             
             TRUE ~ `Grupo PR_FOR`),
           Promotor_2 = ifelse(`GRUPO PROMOTOR` == "Red Premiumguest", "Red Premiumguest", Promotor),
           Recaudacion = `Pagado Banco` + Descuento,
           `Comisión Red de Promoción` = case_when(
             Negocio == "Carpe Diem Barcelona" & Tipo == "VIP" ~ round(((Recaudacion/1.1)*0.1)*1.21,3),
             Negocio == "Opium Malgrat" | Negocio == "Tropicana Malgrat" ~ round(((Recaudacion/1.1)*0.05)*1.21,3),
             Negocio == "Downtown Barcelona" | Ciudad == "Madrid" ~ 0,
             `CANAL DE VENTA` == "ORGANIZADOR" ~ 0,
             
             TRUE ~ round(((Recaudacion/1.1)*0.15)*1.21,3)),
           
           LIQUIDACIÓN = round((Recaudacion -`Comisión Red de Promoción`),3),
           `Comisión Promotores`= case_when(
             Negocio == "Opium Malgrat" | Negocio == "Tropicana Malgrat" ~  0,
             `Comisión Red de Promoción` == 0  ~ 0,
             TRUE ~ round((Recaudacion*0.1)- Descuento,3)),
           
           `PRECIO TICKET_1` = `Precio Ticket`,
           `Precio Ticket + Gasto gestión` = round((`Pagado Banco` + Fee),3),
           `Margen Promoción` = round((`Comisión Red de Promoción` * 0.3333333),3),
           `Total margen` = `Margen Promoción` + Fee,
           `PRECIO TICKET_2` = round((`Precio Ticket`/`Nº Tickets`),3),
           `NEGOCIO&CODE AUTH` = ifelse(is.na(Auth_code),paste0(Negocio, ""),paste0(Negocio, Auth_code) ),
           `DEVOLUCIONES CHECK` = "",
           `DIA DE LA SEMANA` = weekdays(as.Date(`Fecha Compra`, format = "%m/%d/%Y")),
           `DIA Compra` = format(as.POSIXlt(`Fecha Compra`, format="%m/%d/%Y"),format="%d"),
           `SEMANA Compra` =   paste("Semana",as.numeric(format(as.POSIXlt(`Fecha Compra`, format="%m/%d/%Y"),format="%V"))+1,sep="_"),
           `CLIENTE QUE VENDE` = "",
           `FECHA EVENTO BIEN` = `FECHA EVENTO`,
           `AÑO PORTFOLIO` = 1899,
           `Pagado?`  = "",
           AUTOFACTURA = ifelse(is.na(AUTOFACTURA), "Autofactura", AUTOFACTURA )
    )
  
  
  output$`Fórmula para ajuste Domingos en la semana_1` <- wday(as.POSIXlt(output$`FECHA EVENTO`, format="%m/%d/%Y"))-1
  output$`Fórmula para ajuste Domingos en la semana_2` <- wday(as.POSIXlt(output$`FECHA COMPRA`, format="%m/%d/%Y"))-1
  
  output <- output %>% replace_na(list(Grupo = "0"))
  
  output <- output %>% 
    dplyr::select( "#", "Tipo", "Cif", "Negocio", "Id Negocio", "Id KAM Negocio", "Ciudad", "White Label", "ID BUSINESS WL", "ID KAM WL",
                   "Promotor", "Grupo PR_FOR", "Grupo PROMO", "UTM campaign...14", "Promo Code", "Nº Tickets", "Asistencia", "Precio Ticket", 
                   "Descuento", "Pagado Banco", "Tipo Pago", "Fee", "Status", "Unique User", "New User...25", "Num Events", "Num Events Asist",
                   "Show Deleted", "Fecha Compra", "Comprador", "Contacto", "Teléfono", "Idioma", "Evento", "Fecha evento", "Fecha asistencia", 
                   "Servicio", "Auth_code", "UTM_Code", "Nº Pedido", "Registered", "New User...42", "UTM source", "UTM medium", "UTM campaign...45",
                   "Creador", "Pagado banco >0", "FECHA COMPRA", "MES COMPRA", "AÑO COMPRA", "FECHA EVENTO", "MES EVENTO", "AÑO EVENTO", "DIA",
                   "Fórmula para ajuste Domingos en la semana_1", "SEMANA", "Ventas para Eventos fuera del mes de compra", "GRUPO", "NEGOCIO_1",
                   
                   "GRUPO PROMOTOR", "Promotor_2",  "CANAL DE VENTA", "Recaudacion",
                   "Comisión Red de Promoción", "LIQUIDACIÓN", "AUTOFACTURA", "Comisión Promotores","PERIODICIDAD_LIQUIDACION", "Remesa", "PRECIO TICKET_1", 
                   "Precio Ticket + Gasto gestión", "Margen Promoción", "Total margen",
                   "Negocio está en pestaña Datos Maestros Generales", "WL está en pestaña Datos Maestros Generales", "PRECIO TICKET_2", "NEGOCIO&CODE AUTH",
                   "DEVOLUCIONES CHECK", "DIA DE LA SEMANA","DIA Compra","Fórmula para ajuste Domingos en la semana_2","SEMANA Compra", "CLIENTE QUE VENDE",  "TIPO CLIENTE",
                   "FECHA EVENTO BIEN", "AÑO PORTFOLIO", "Comercial Negocio", "Comercial WL" , "Pagado?"
    )
  
  
  cols_output_num <- c("AÑO EVENTO","AÑO COMPRA","DIA", "DIA Compra" )
  
  output[cols_output_num] <- sapply(output[cols_output_num],as.numeric)
  output
}

start_time <- Sys.time()

output <- get_output(input)



sprintf("ignore/output_%s.csv",today)

# write.csv(output, sprintf("ignore/output_%s.csv",today))

write.xlsx(output, file = "ignore/output_premiumguest.xlsx", colNames = TRUE)

write.xlsx(output, file="ignore/output_premiumguest.xlsx",
           sheetName="output", append=FALSE)

end_time <- Sys.time()


# Creating pivots -----

# TODO To think about first week of the year!!!
# current_year <- format(as.POSIXlt(Sys.Date(), format="%m/%d/%Y"),format="%Y")
# begining_current_year <- lubridate::floor_date(Sys.Date(),"year")
# begining_current_month <- lubridate::floor_date(Sys.Date(),"month")
# current_week <-  paste("Semana",as.numeric(format(as.POSIXlt(Sys.Date(), format="%m/%d/%Y"),format="%V"))+1,sep="_")
# previous_week <-  paste("Semana",as.numeric(format(as.POSIXlt(Sys.Date(), format="%m/%d/%Y"),format="%V")),sep="_")
# current_month <- match(months(as.POSIXlt(Sys.Date(), format="%m/%d/%Y")), month.name)

dummy_date <- "2022-07-18"
current_year <- "2022"
# current_week <-  "Semana_28"
previous_week <-  "Semana_28"
current_month <- 7
previous_month <- current_month-1


#Pivot_1: liquidacion_semanal_lunes -----

liquidacion_semanal_lunes_values <- output %>%
  dplyr::filter(`AÑO EVENTO` == current_year & SEMANA == previous_week) %>% 
  dplyr::group_by(Negocio,GRUPO, Ciudad,AUTOFACTURA, Remesa, PERIODICIDAD_LIQUIDACION) %>%
  dplyr::summarise(Recaudacion_sum= sum(Recaudacion),
                   Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`),
                   liquidacion_1_sum = sum(LIQUIDACIÓN)) %>% 
  # ungroup() %>% 
  dplyr::mutate(`AÑO EVENTO` = current_year,
                SEMANA = previous_week)

liquidacion_semanal_lunes_grand_total_sum <- output %>%
  dplyr::filter(`AÑO EVENTO` == current_year & SEMANA == previous_week) %>% 
  dplyr::group_by(Negocio,GRUPO, Ciudad,AUTOFACTURA, Remesa) %>% 
  plyr::summarise(Recaudacion_sum= sum(Recaudacion),
                  Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`),
                  liquidacion_1_sum = sum(LIQUIDACIÓN)) %>% 
  ungroup() %>% 
  dplyr::mutate(Negocio = "Grand Total",
                GRUPO = "",
                Ciudad = "",
                AUTOFACTURA = "",
                Remesa = "",
                `AÑO EVENTO` = "",
                SEMANA = ""
  )


remesa_total_sum <- output %>%
  dplyr::filter(`AÑO EVENTO` == current_year & SEMANA == previous_week & Remesa == "SI" ) %>% 
  dplyr::group_by(Negocio,GRUPO, Ciudad,AUTOFACTURA, Remesa) %>% 
  plyr::summarise(Remesa_sum = sum(LIQUIDACIÓN)) %>% 
  ungroup() 

liquidacion_semanal_lunes_grand_total_sum$Remesa <- as.character(remesa_total_sum$Remesa_sum)

liquidacion_semanal_lunes <- bind_rows(liquidacion_semanal_lunes_values, liquidacion_semanal_lunes_grand_total_sum)

# write.xlsx(liquidacion_semanal_lunes, file = "ignore/liquidacion_semanal_lunes.xlsx", colNames = TRUE)


#Pivot_2: liquidacion_mensual -----


liquidacion_mensual_values <- output %>%
  dplyr::filter(`AÑO EVENTO` == current_year & `MES EVENTO` == current_month) %>% 
  dplyr::group_by(Negocio,GRUPO, Ciudad,AUTOFACTURA, Remesa,PERIODICIDAD_LIQUIDACION ) %>% 
  dplyr::summarise(Recaudacion_sum= sum(Recaudacion),
                   Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`),
                   liquidacion_1_sum = sum(LIQUIDACIÓN)) %>% 
  ungroup() %>% 
  dplyr::mutate(`AÑO EVENTO` = current_year,
                `MES EVENTO` = current_month)

liquidacion_mensual_values$`MES EVENTO` <- as.character(liquidacion_mensual_values$`MES EVENTO`)

liquidacion_mensual_grand_total_sum <- output %>%
  dplyr::filter(`AÑO EVENTO` == current_year & `MES EVENTO` == current_month) %>% 
  dplyr::group_by(Negocio,GRUPO, Ciudad,AUTOFACTURA, Remesa) %>% 
  plyr::summarise(Recaudacion_sum= sum(Recaudacion),
                  Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`),
                  liquidacion_1_sum = sum(LIQUIDACIÓN)) %>% 
  ungroup() %>% 
  dplyr::mutate(Negocio = "Grand Total",
                GRUPO = "",
                Ciudad = "",
                AUTOFACTURA = "",
                Remesa = "",
                `AÑO EVENTO` = "",
                `MES EVENTO` = ""
  )


remesa_total_sum <- output %>%
  dplyr::filter(`AÑO EVENTO` == current_year & `MES EVENTO` == current_month & Remesa == "SI" ) %>% 
  dplyr::group_by(Negocio,GRUPO, Ciudad,AUTOFACTURA, Remesa) %>% 
  plyr::summarise(Remesa_sum = sum(LIQUIDACIÓN)) %>% 
  ungroup() 

liquidacion_mensual_grand_total_sum$Remesa <- as.character(remesa_total_sum$Remesa_sum)

liquidacion_mensual <- bind_rows(liquidacion_mensual_values, liquidacion_mensual_grand_total_sum)

# write.xlsx(liquidacion_mensual, file = "ignore/liquidacion_mensual.xlsx", colNames = TRUE)



# Pivot_3: cierre_facturacion -----

cierre_facturacion_values <- output %>%
  dplyr::mutate(evento_mismo_mes = (`Ventas para Eventos fuera del mes de compra` == "Evento mismo mes compra"),
                evento_futuro = (`Ventas para Eventos fuera del mes de compra` == "Compra Evento Futuro" )) %>% 
  dplyr::mutate(Recaudacion_mismo = ifelse(evento_mismo_mes, Recaudacion,0),
                Recaudacion_futuro = ifelse(evento_futuro, Recaudacion,0),
                Comision_Red_de_Promocion_mismo = ifelse(evento_mismo_mes, `Comisión Red de Promoción`,0),
                Comision_Red_de_Promocion_futuro = ifelse(evento_futuro, `Comisión Red de Promoción`,0)) %>% 
  dplyr::filter(`AÑO COMPRA` == current_year & `MES COMPRA` == current_month & AUTOFACTURA == "Autofactura") %>% 
  dplyr::group_by(Ciudad, Negocio, `Comercial Negocio`, `White Label`,`Comercial WL`,  `CANAL DE VENTA`,Tipo ) %>% 
  dplyr::summarise(Recaudacion_futuro_sum= sum(Recaudacion_futuro),
                   Comision_Red_de_Promocion_futuro_sum = sum(Comision_Red_de_Promocion_futuro),
                   
                   Recaudacion_mismo_sum = sum(Recaudacion_mismo),
                   Comision_Red_de_Promocion_mismo_sum= sum(Comision_Red_de_Promocion_mismo),
                   
                   Recaudacion_sum= sum(Recaudacion),
                   Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`)) %>% 
  ungroup() %>% 
  dplyr::mutate(`AÑO COMPRA` = current_year,
                `MES COMPRA` = current_month)

cierre_facturacion_values$`MES COMPRA` <- as.character(cierre_facturacion_values$`MES COMPRA`)

cierre_facturacion_grand_total <- output %>%
  dplyr::mutate(evento_mismo_mes = (`Ventas para Eventos fuera del mes de compra` == "Evento mismo mes compra"),
                evento_futuro = (`Ventas para Eventos fuera del mes de compra` == "Compra Evento Futuro" )) %>% 
  dplyr::mutate(Recaudacion_mismo = ifelse(evento_mismo_mes, Recaudacion,0),
                Recaudacion_futuro = ifelse(evento_futuro, Recaudacion,0),
                Comision_Red_de_Promocion_mismo = ifelse(evento_mismo_mes, `Comisión Red de Promoción`,0),
                Comision_Red_de_Promocion_futuro = ifelse(evento_futuro, `Comisión Red de Promoción`,0)) %>% 
  dplyr::filter(`AÑO COMPRA` == current_year & `MES COMPRA` == current_month & AUTOFACTURA == "Autofactura") %>% 
  dplyr::group_by(Ciudad, Negocio, `Comercial Negocio`, `White Label`,`Comercial WL`,  `CANAL DE VENTA`,Tipo ) %>% 
  plyr::summarise(Recaudacion_futuro_sum= sum(Recaudacion_futuro),
                  Comision_Red_de_Promocion_futuro_sum = sum(Comision_Red_de_Promocion_futuro),
                  
                  Recaudacion_mismo_sum = sum(Recaudacion_mismo),
                  Comision_Red_de_Promocion_mismo_sum= sum(Comision_Red_de_Promocion_mismo),
                  
                  Recaudacion_sum= sum(Recaudacion),
                  Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`)) %>% 
  ungroup() %>% 
  dplyr::mutate(Ciudad = "Grand Total")

cierre_facturacion <- bind_rows(cierre_facturacion_values, cierre_facturacion_grand_total)



# Pivot_4: cierre_provision  -----

cierre_provision_values <- output %>%
  dplyr::mutate(evento_mismo_mes = (`Ventas para Eventos fuera del mes de compra` == "Evento mismo mes compra"),
                evento_futuro = (`Ventas para Eventos fuera del mes de compra` == "Compra Evento Futuro" )) %>% 
  dplyr::mutate(Recaudacion_mismo = ifelse(evento_mismo_mes, Recaudacion,0),
                Recaudacion_futuro = ifelse(evento_futuro, Recaudacion,0),
                Comision_promotores_mismo = ifelse(evento_mismo_mes, `Comisión Promotores`,0),
                Comision_promotores_futuro = ifelse(evento_futuro, `Comisión Promotores`,0)) %>% 
  dplyr::filter(`AÑO COMPRA` == current_year & `MES COMPRA` == previous_month & AUTOFACTURA == "Autofactura") %>% 
  dplyr::group_by(Ciudad, Negocio, `Comercial Negocio`, `White Label`,`Comercial WL`,  `CANAL DE VENTA`,Tipo ) %>% 
  dplyr::summarise(Recaudacion_futuro_sum= sum(Recaudacion_futuro),
                   Comision_promotores_futuro_sum = sum(Comision_promotores_futuro),
                   
                   Recaudacion_mismo_sum = sum(Recaudacion_mismo),
                   Comision_promotores_mismo_sum= sum(Comision_promotores_mismo),
                   
                   Recaudacion_sum= sum(Recaudacion),
                   Comision_promotores_sum= sum(`Comisión Promotores`)) %>% 
  ungroup() %>% 
  dplyr::mutate(`AÑO COMPRA` = current_year,
                `MES COMPRA` = previous_month)

cierre_provision_values$`MES COMPRA` <- as.character(cierre_provision_values$`MES COMPRA`)

cierre_provision_grand_total <- output %>%
  dplyr::mutate(evento_mismo_mes = (`Ventas para Eventos fuera del mes de compra` == "Evento mismo mes compra"),
                evento_futuro = (`Ventas para Eventos fuera del mes de compra` == "Compra Evento Futuro" )) %>% 
  dplyr::mutate(Recaudacion_mismo = ifelse(evento_mismo_mes, Recaudacion,0),
                Recaudacion_futuro = ifelse(evento_futuro, Recaudacion,0),
                Comision_promotores_mismo = ifelse(evento_mismo_mes, `Comisión Promotores`,0),
                Comision_promotores_futuro = ifelse(evento_futuro, `Comisión Promotores`,0)) %>% 
  dplyr::filter(`AÑO COMPRA` == current_year & `MES COMPRA` == previous_month & AUTOFACTURA == "Autofactura") %>% 
  dplyr::group_by(Ciudad, Negocio, `Comercial Negocio`, `White Label`,`Comercial WL`,  `CANAL DE VENTA`,Tipo ) %>% 
  plyr::summarise(Recaudacion_futuro_sum= round(sum(Recaudacion_futuro),2),
                  Comision_promotores_futuro_sum = round(sum(Comision_promotores_futuro),2),
                  
                  Recaudacion_mismo_sum = round(sum(Recaudacion_mismo),2),
                  Comision_promotores_mismo_sum= round(sum(Comision_promotores_mismo),2),
                  
                  Recaudacion_sum= round(sum(Recaudacion),2),
                  Comision_promotores_sum= round(sum(`Comisión Promotores`),2)) %>% 
  ungroup() %>% 
  dplyr::mutate(Ciudad = "Grand Total")

cierre_provision <- bind_rows(cierre_provision_values, cierre_provision_grand_total)



# Pivot_5:cierre_anulacion_prov_ev_futur  -----


cierre_anulacion_prov_ev_futur_values <- output %>%
  dplyr::filter(`AÑO COMPRA` == current_year &
                  `MES COMPRA` == previous_month  &
                  AUTOFACTURA == "Autofactura" &
                  `Ventas para Eventos fuera del mes de compra` == "Compra Evento Futuro"
  ) %>%
  dplyr::group_by(Ciudad, Negocio, `Comercial Negocio`, `White Label`,`Comercial WL`,  `CANAL DE VENTA`,Tipo )%>% 
  dplyr::summarise(Recaudacion_sum= sum(Recaudacion),
                   Comision_promotores_sum = sum(`Comisión Promotores`)) %>% 
  ungroup() %>% 
  dplyr::mutate(`AÑO COMPRA` = current_year,
                `MES COMPRA` = previous_month)

cierre_anulacion_prov_ev_futur_values$`MES COMPRA` <- as.character(cierre_anulacion_prov_ev_futur_values$`MES COMPRA`)

cierre_anulacion_prov_ev_futur_values_ciudad <- cierre_anulacion_prov_ev_futur_values %>% 
  dplyr::group_by(Ciudad) %>% 
  dplyr::summarise(Recaudacion_sum= sum(Recaudacion_sum),
                   Comision_promotores_sum = sum(Comision_promotores_sum)) %>% 
  ungroup() 


cierre_anulacion_prov_ev_futur_grand_total_sum <- output %>%
  dplyr::filter(`AÑO COMPRA` == current_year & `MES COMPRA` == previous_month & AUTOFACTURA == "Autofactura" & `Ventas para Eventos fuera del mes de compra` == "Compra Evento Futuro") %>%
  dplyr::group_by(Ciudad, Negocio, `Comercial Negocio`, `White Label`,`Comercial WL`,  `CANAL DE VENTA`,Tipo )%>% 
  plyr::summarise(Recaudacion_sum= sum(Recaudacion),
                  Comision_promotores_sum = sum(`Comisión Promotores`)) %>% 
  ungroup() %>% 
  dplyr::mutate(Negocio = "Grand Total",
                Ciudad = "",
                `AÑO COMPRA` = "",
                `MES COMPRA` = ""
  )


cierre_anulacion_prov_ev_futur <- bind_rows(cierre_anulacion_prov_ev_futur_values, cierre_anulacion_prov_ev_futur_grand_total_sum)

# write.xlsx(cierre_anulacion_prov_ev_futur, file = "ignore/cierre_anulacion_prov_ev_futur.xlsx", colNames = TRUE)
# write.xlsx(output, file = "ignore/output.xlsx", colNames = TRUE)

# Pivot_6: reporting_clientes_varios  -----


reporting_clientes_varios_values <- output %>%
  dplyr::group_by(SEMANA,  `FECHA EVENTO`, DIA, Evento, Creador ) %>% 
  dplyr::summarise(Tickets_sum= sum(`Nº Tickets`),
                   Recaudacion_sum= sum(Recaudacion),
                   Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`),
                   LIQUIDACIÓN_sum = sum(LIQUIDACIÓN)) %>% 
  ungroup()

reporting_clientes_varios_values_semana <- output %>%
  dplyr::group_by(SEMANA ) %>% 
  dplyr::summarise(Tickets_sum= sum(`Nº Tickets`),
                   Recaudacion_sum= sum(Recaudacion),
                   Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`),
                   LIQUIDACIÓN_sum = sum(LIQUIDACIÓN)) %>% 
  ungroup()

reporting_clientes_varios_grand_total_sum <- output %>%
  dplyr::group_by(SEMANA,  `FECHA EVENTO`, DIA, Evento, Creador ) %>% 
  plyr::summarise(Tickets_sum= sum(`Nº Tickets`),
                  Recaudacion_sum= sum(Recaudacion),
                  Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`),
                  LIQUIDACIÓN_sum = sum(LIQUIDACIÓN)) %>% 
  ungroup() %>%  
  dplyr::mutate(SEMANA = "Grand Total",
  )



reporting_clientes_varios <- bind_rows(reporting_clientes_varios_values, reporting_clientes_varios_grand_total_sum)

# write.xlsx(reporting_clientes_varios, file = "ignore/reporting_clientes_varios.xlsx", colNames = TRUE)


#  Write all files together ----
#  option 1 -----
# https://www.rdocumentation.org/packages/openxlsx/versions/4.2.5/topics/write.xlsx
require(openxlsx)


## headerStyles

list_of_datasets <- list("master_data_check" = master_data_check,
                         "output" = output,
                         "liquidacion_semanal_lunes" = liquidacion_semanal_lunes,
                         "liquidacion_mensual"= liquidacion_mensual,
                         "cierre_facturacion"  = cierre_facturacion,
                         "cierre_provision"  = cierre_provision,
                         "cierre_anulacion_prov_ev_futur"  = cierre_anulacion_prov_ev_futur,
                         "cierre_anulacion_prov_ev_futur_values_ciudad"  = cierre_anulacion_prov_ev_futur_values_ciudad,
                         "reporting_clientes_varios"  = reporting_clientes_varios,
                         "reporting_clientes_varios_values_semana"  = reporting_clientes_varios_values_semana)

# openxlsx::write.xlsx(list_of_datasets, file = "ignore/output_premiumguest_simple.xlsx")

hs1 <- createStyle(fgFill = "#4F81BD", halign = "CENTER", textDecoration = "Bold",
                   border = "Bottom", fontColour = "white")
# testing xsls format ----
Sys.time()

openxlsx::write.xlsx(list_of_datasets, file = sprintf("ignore/output_premiumguest_%s.xlsx",today),
                     borders = "all",
                     borderColour = "#6B6C68",
                     colWidths  = "auto",
                     headerStyle = hs1,
                     tabColour = c("grey","black","dodgerblue", "cyan", "violet","violetred1", "yellow", "yellow","springgreen","springgreen"   ))

end_time <- Sys.time()

print(sprintf("Start time: %s,/n End time: %s ", start_time, end_time ))





