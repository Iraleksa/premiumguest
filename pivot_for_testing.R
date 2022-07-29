# Creating pivots -----

# TODO To think about first week of the year!!!
current_year <- format(as.POSIXlt(Sys.Date(), format="%m/%d/%Y"),format="%Y")
begining_current_year <- lubridate::floor_date(Sys.Date(),"year")
begining_current_month <- lubridate::floor_date(Sys.Date(),"month")
current_week <-  paste("Semana",as.numeric(format(as.POSIXlt(Sys.Date(), format="%m/%d/%Y"),format="%V"))+1,sep="_")
previous_week <-  paste("Semana",as.numeric(format(as.POSIXlt(Sys.Date(), format="%m/%d/%Y"),format="%V")),sep="_")
current_month <- match(months(as.POSIXlt(Sys.Date(), format="%m/%d/%Y")), month.name)

# dummy_date <- "2022-06-30"
# current_year <- "2022"
# current_week <-  "Semana_27"
# previous_week <-  "Semana_26"
# current_month <- 6

#Pivot_1: liquidacion_semanal_lunes -----

liquidacion_semanal_lunes_values <- output_vicent_renamed %>%
  dplyr::filter(`AÑO EVENTO` == current_year & SEMANA == previous_week) %>% 
  dplyr::group_by(Negocio,GRUPO, Ciudad,AUTOFACTURA, Remesa) %>% 
  dplyr::summarise(Recaudacion_sum= sum(Recaudación),
                   Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`),
                   liquidacion_1_sum = sum(LIQUIDACIÓN_1)) %>% 
  ungroup() %>% 
  dplyr::mutate(`AÑO EVENTO` = current_year,
                SEMANA = previous_week)

liquidacion_semanal_lunes_grand_total_sum <- output_vicent_renamed %>%
  dplyr::filter(`AÑO EVENTO` == current_year & SEMANA == previous_week) %>% 
  dplyr::group_by(Negocio,GRUPO, Ciudad,AUTOFACTURA, Remesa) %>% 
  plyr::summarise(Recaudacion_sum= sum(Recaudación),
                  Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`),
                  liquidacion_1_sum = sum(LIQUIDACIÓN_1)) %>% 
  ungroup() %>% 
  dplyr::mutate(Negocio = "Grand Total",
                GRUPO = "",
                Ciudad = "",
                AUTOFACTURA = "",
                Remesa = "",
                `AÑO EVENTO` = "",
                SEMANA = ""
  )


remesa_total_sum <- output_vicent_renamed %>%
  dplyr::filter(`AÑO EVENTO` == current_year & SEMANA == previous_week & Remesa == "SI" ) %>% 
  dplyr::group_by(Negocio,GRUPO, Ciudad,AUTOFACTURA, Remesa) %>% 
  plyr::summarise(Remesa_sum = sum(LIQUIDACIÓN_1)) %>% 
  ungroup() 

liquidacion_semanal_lunes_grand_total_sum$Remesa <- as.character(remesa_total_sum$Remesa_sum)

liquidacion_semanal_lunes <- bind_rows(liquidacion_semanal_lunes_values, liquidacion_semanal_lunes_grand_total_sum)

# write.xlsx(liquidacion_semanal_lunes, file = "ignore/liquidacion_semanal_lunes.xlsx", colNames = TRUE)


#Pivot_2: liquidacion_mensual -----


liquidacion_mensual_values <- output_vicent_renamed %>%
  dplyr::filter(`AÑO EVENTO` == current_year & `MES EVENTO` == current_month) %>% 
  dplyr::group_by(Negocio,GRUPO, Ciudad,AUTOFACTURA, Remesa) %>% 
  dplyr::summarise(Recaudacion_sum= sum(Recaudación),
                   Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`),
                   liquidacion_1_sum = sum(LIQUIDACIÓN_1)) %>% 
  ungroup() %>% 
  dplyr::mutate(`AÑO EVENTO` = current_year,
                `MES EVENTO` = current_month)

liquidacion_mensual_values$`MES EVENTO` <- as.character(liquidacion_mensual_values$`MES EVENTO`)

liquidacion_mensual_grand_total_sum <- output_vicent_renamed %>%
  dplyr::filter(`AÑO EVENTO` == current_year & `MES EVENTO` == current_month) %>% 
  dplyr::group_by(Negocio,GRUPO, Ciudad,AUTOFACTURA, Remesa) %>% 
  plyr::summarise(Recaudacion_sum= sum(Recaudación),
                  Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`),
                  liquidacion_1_sum = sum(LIQUIDACIÓN_1)) %>% 
  ungroup() %>% 
  dplyr::mutate(Negocio = "Grand Total",
                GRUPO = "",
                Ciudad = "",
                AUTOFACTURA = "",
                Remesa = "",
                `AÑO EVENTO` = "",
                `MES EVENTO` = ""
  )


remesa_total_sum <- output_vicent_renamed %>%
  dplyr::filter(`AÑO EVENTO` == current_year & `MES EVENTO` == current_month & Remesa == "SI" ) %>% 
  dplyr::group_by(Negocio,GRUPO, Ciudad,AUTOFACTURA, Remesa) %>% 
  plyr::summarise(Remesa_sum = sum(LIQUIDACIÓN_1)) %>% 
  ungroup() 

liquidacion_mensual_grand_total_sum$Remesa <- as.character(remesa_total_sum$Remesa_sum)

liquidacion_mensual <- bind_rows(liquidacion_mensual_values, liquidacion_mensual_grand_total_sum)

# write.xlsx(liquidacion_mensual, file = "ignore/liquidacion_mensual.xlsx", colNames = TRUE)



# Pivot_3: cierre_facturacion -----

cierre_facturacion_values <- output_vicent_renamed %>%
  dplyr::mutate(evento_mismo_mes = (`Ventas para Eventos fuera del mes de compra` == "Evento mismo mes compra"),
                evento_futuro = (`Ventas para Eventos fuera del mes de compra` == "Compra Evento Futuro" )) %>% 
  dplyr::mutate(Recaudacion_mismo = ifelse(evento_mismo_mes, Recaudación,0),
                Recaudacion_futuro = ifelse(evento_futuro, Recaudación,0),
                Comision_Red_de_Promocion_mismo = ifelse(evento_mismo_mes, `Comisión Red de Promoción`,0),
                Comision_Red_de_Promocion_futuro = ifelse(evento_futuro, `Comisión Red de Promoción`,0)) %>% 
  dplyr::filter(`AÑO COMPRA` == current_year & `MES COMPRA` == current_month & AUTOFACTURA == "Autofactura") %>% 
  dplyr::group_by(Ciudad, Negocio, `KAM Negocio`, `White Label`, `KAM WL`, `CANAL DE VENTA`,Tipo ) %>% 
  dplyr::summarise(Recaudacion_futuro_sum= sum(Recaudacion_futuro),
                   Comision_Red_de_Promocion_futuro_sum = sum(Comision_Red_de_Promocion_futuro),
                   
                   Recaudacion_mismo_sum = sum(Recaudacion_mismo),
                   Comision_Red_de_Promocion_mismo_sum= sum(Comision_Red_de_Promocion_mismo),
                   
                   Recaudacion_sum= sum(Recaudación),
                   Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`)) %>% 
  ungroup() %>% 
  dplyr::mutate(`AÑO COMPRA` = current_year,
                `MES COMPRA` = current_month)

cierre_facturacion_values$`MES COMPRA` <- as.character(cierre_facturacion_values$`MES COMPRA`)

cierre_facturacion_grand_total <- output_vicent_renamed %>%
  dplyr::mutate(evento_mismo_mes = (`Ventas para Eventos fuera del mes de compra` == "Evento mismo mes compra"),
                evento_futuro = (`Ventas para Eventos fuera del mes de compra` == "Compra Evento Futuro" )) %>% 
  dplyr::mutate(Recaudacion_mismo = ifelse(evento_mismo_mes, Recaudación,0),
                Recaudacion_futuro = ifelse(evento_futuro, Recaudación,0),
                Comision_Red_de_Promocion_mismo = ifelse(evento_mismo_mes, `Comisión Red de Promoción`,0),
                Comision_Red_de_Promocion_futuro = ifelse(evento_futuro, `Comisión Red de Promoción`,0)) %>% 
  dplyr::filter(`AÑO COMPRA` == current_year & `MES COMPRA` == current_month & AUTOFACTURA == "Autofactura") %>% 
  dplyr::group_by(Ciudad, Negocio, `KAM Negocio`, `White Label`, `KAM WL`, `CANAL DE VENTA`,Tipo ) %>% 
  plyr::summarise(Recaudacion_futuro_sum= sum(Recaudacion_futuro),
                  Comision_Red_de_Promocion_futuro_sum = sum(Comision_Red_de_Promocion_futuro),
                  
                  Recaudacion_mismo_sum = sum(Recaudacion_mismo),
                  Comision_Red_de_Promocion_mismo_sum= sum(Comision_Red_de_Promocion_mismo),
                  
                  Recaudacion_sum= sum(Recaudación),
                  Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`)) %>% 
  ungroup() %>% 
  dplyr::mutate(Ciudad = "Grand Total")

cierre_facturacion <- bind_rows(cierre_facturacion_values, cierre_facturacion_grand_total)



# Pivot_4: cierre_provision  -----

cierre_provision_values <- output_vicent_renamed %>%
  dplyr::mutate(evento_mismo_mes = (`Ventas para Eventos fuera del mes de compra` == "Evento mismo mes compra"),
                evento_futuro = (`Ventas para Eventos fuera del mes de compra` == "Compra Evento Futuro" )) %>% 
  dplyr::mutate(Recaudacion_mismo = ifelse(evento_mismo_mes, Recaudación,0),
                Recaudacion_futuro = ifelse(evento_futuro, Recaudación,0),
                Comision_promotores_mismo = ifelse(evento_mismo_mes, `Comisión Promotores`,0),
                Comision_promotores_futuro = ifelse(evento_futuro, `Comisión Promotores`,0)) %>% 
  dplyr::filter(`AÑO COMPRA` == current_year & `MES COMPRA` == current_month & AUTOFACTURA == "Autofactura") %>% 
  dplyr::group_by(Ciudad, Negocio, `KAM Negocio`, `White Label`, `KAM WL`, `CANAL DE VENTA`,Tipo ) %>% 
  dplyr::summarise(Recaudacion_futuro_sum= sum(Recaudacion_futuro),
                   Comision_promotores_futuro_sum = sum(Comision_promotores_futuro),
                   
                   Recaudacion_mismo_sum = sum(Recaudacion_mismo),
                   Comision_promotores_mismo_sum= sum(Comision_promotores_mismo),
                   
                   Recaudacion_sum= sum(Recaudación),
                   Comision_promotores_sum= sum(`Comisión Promotores`)) %>% 
  ungroup() %>% 
  dplyr::mutate(`AÑO COMPRA` = current_year,
                `MES COMPRA` = current_month)

cierre_provision_values$`MES COMPRA` <- as.character(cierre_provision_values$`MES COMPRA`)

cierre_provision_grand_total <- output_vicent_renamed %>%
  dplyr::mutate(evento_mismo_mes = (`Ventas para Eventos fuera del mes de compra` == "Evento mismo mes compra"),
                evento_futuro = (`Ventas para Eventos fuera del mes de compra` == "Compra Evento Futuro" )) %>% 
  dplyr::mutate(Recaudacion_mismo = ifelse(evento_mismo_mes, Recaudación,0),
                Recaudacion_futuro = ifelse(evento_futuro, Recaudación,0),
                Comision_promotores_mismo = ifelse(evento_mismo_mes, `Comisión Promotores`,0),
                Comision_promotores_futuro = ifelse(evento_futuro, `Comisión Promotores`,0)) %>% 
  dplyr::filter(`AÑO COMPRA` == current_year & `MES COMPRA` == current_month & AUTOFACTURA == "Autofactura") %>% 
  dplyr::group_by(Ciudad, Negocio, `KAM Negocio`, `White Label`, `KAM WL`, `CANAL DE VENTA`,Tipo ) %>% 
  plyr::summarise(Recaudacion_futuro_sum= round(sum(Recaudacion_futuro),2),
                  Comision_promotores_futuro_sum = round(sum(Comision_promotores_futuro),2),
                  
                  Recaudacion_mismo_sum = round(sum(Recaudacion_mismo),2),
                  Comision_promotores_mismo_sum= round(sum(Comision_promotores_mismo),2),
                  
                  Recaudacion_sum= round(sum(Recaudación),2),
                  Comision_promotores_sum= round(sum(`Comisión Promotores`),2)) %>% 
  ungroup() %>% 
  dplyr::mutate(Ciudad = "Grand Total")

cierre_provision <- bind_rows(cierre_provision_values, cierre_provision_grand_total)



# Pivot_5:cierre_anulacion_prov_ev_futur  -----


cierre_anulacion_prov_ev_futur_values <- output_vicent_renamed %>%
  dplyr::filter(`AÑO COMPRA` == current_year &
                  `MES COMPRA` == current_month &
                  AUTOFACTURA == "Autofactura" &
                  `Ventas para Eventos fuera del mes de compra` == "Compra Evento Futuro") %>%
  dplyr::group_by(Ciudad, Negocio, `KAM Negocio`, `White Label`, `KAM WL`, `CANAL DE VENTA`,Tipo ) %>% 
  dplyr::summarise(Recaudacion_sum= sum(Recaudación),
                   Comision_promotores_sum = sum(`Comisión Promotores`)) %>% 
  ungroup() %>% 
  dplyr::mutate(`AÑO COMPRA` = current_year,
                `MES COMPRA` = current_month)

cierre_anulacion_prov_ev_futur_values$`MES COMPRA` <- as.character(cierre_anulacion_prov_ev_futur_values$`MES COMPRA`)

cierre_anulacion_prov_ev_futur_values_ciudad <- cierre_anulacion_prov_ev_futur_values %>% 
  dplyr::group_by(Ciudad) %>% 
  dplyr::summarise(Recaudacion_sum= sum(Recaudacion_sum),
                   Comision_promotores_sum = sum(Comision_promotores_sum)) %>% 
  ungroup() 


cierre_anulacion_prov_ev_futur_grand_total_sum <- output_vicent_renamed %>%
  dplyr::filter(`AÑO COMPRA` == current_year & `MES COMPRA` == current_month & AUTOFACTURA == "Autofactura" & `Ventas para Eventos fuera del mes de compra` == "Compra Evento Futuro") %>%
  dplyr::group_by(Ciudad, Negocio, `KAM Negocio`, `White Label`, `KAM WL`, `CANAL DE VENTA`,Tipo) %>% 
  plyr::summarise(Recaudacion_sum= sum(Recaudación),
                  Comision_promotores_sum = sum(`Comisión Promotores`)) %>% 
  ungroup() %>% 
  dplyr::mutate(Negocio = "Grand Total",
                Ciudad = "",
                `AÑO COMPRA` = "",
                `MES COMPRA` = ""
  )


cierre_anulacion_prov_ev_futur <- bind_rows(cierre_anulacion_prov_ev_futur_values, cierre_anulacion_prov_ev_futur_grand_total_sum)

# write.xlsx(cierre_anulacion_prov_ev_futur, file = "ignore/cierre_anulacion_prov_ev_futur.xlsx", colNames = TRUE)
# write.xlsx(output_vicent_renamed, file = "ignore/output_vicent_renamed.xlsx", colNames = TRUE)

# Pivot_6: reporting_clientes_varios  -----


reporting_clientes_varios_values <- output_vicent_renamed %>%
  dplyr::group_by(SEMANA, Negocio, `FECHA EVENTO`, DIA, Evento, Creador ) %>% 
  dplyr::summarise(Tickets_sum= sum(`Nº Tickets`),
                   Recaudacion_sum= sum(Recaudación),
                   Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`)) %>% 
  ungroup()

reporting_clientes_varios_values_semana <- output_vicent_renamed %>%
  dplyr::group_by(SEMANA ) %>% 
  dplyr::summarise(Tickets_sum= sum(`Nº Tickets`),
                   Recaudacion_sum= sum(Recaudación),
                   Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`)) %>% 
  ungroup()

reporting_clientes_varios_grand_total_sum <- output_vicent_renamed %>%
  dplyr::group_by(SEMANA, Negocio, `FECHA EVENTO`, DIA, Evento, Creador ) %>% 
  plyr::summarise(Tickets_sum= sum(`Nº Tickets`),
                  Recaudacion_sum= sum(Recaudación),
                  Comision_Red_de_Promocion_sum= sum(`Comisión Red de Promoción`)) %>% 
  ungroup() %>%  
  dplyr::mutate(SEMANA = "Grand Total",
                Negocio = "", 
  )



reporting_clientes_varios <- bind_rows(reporting_clientes_varios_values, reporting_clientes_varios_grand_total_sum)

# write.xlsx(reporting_clientes_varios, file = "ignore/reporting_clientes_varios.xlsx", colNames = TRUE)


#  Write all files together ----
#  option 1 -----
# https://www.rdocumentation.org/packages/openxlsx/versions/4.2.5/topics/write.xlsx
require(openxlsx)


## headerStyles

list_of_datasets <- list("output_vicent_renamed" = output_vicent_renamed,
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

openxlsx::write.xlsx(list_of_datasets, file = "ignore/output_premiumguest_vicent.xlsx",
                     borders = "all",
                     borderColour = "#6B6C68",
                     colWidths  = "auto",
                     headerStyle = hs1,
                     tabColour = c("black","dodgerblue", "cyan", "violet","violetred1", "yellow", "yellow","springgreen","springgreen"   ))

Sys.time()