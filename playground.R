output_vicent <- read_excel("ignore/premiumguest/Datos semana 28 Irina.xlsx", sheet = "Hoja1") 

test <- readxl_example("ignore/premiumguest/Datos semana 28 Irina.xlsx")

library(readxl)
test <- read_excel("ignore/premiumguest/output_vicent.xlsx")

output_vicent <- output_vicent[,1:93]

output_vicent <- output_vicent[,c(1:91,93,92)]

# Starting to align databases ----
output_generated <- output %>% unique()

output_vicent_renamed <- output_vicent

# paste(colnames(output_generated), collapse='", "')

col_names = c("#", "Tipo", "Cif", "Negocio", "Id Negocio", "Id KAM Negocio", "Ciudad", "White Label",
             "ID BUSINESS WL", "ID KAM WL", "Promotor", "Grupo PR_FOR", "Grupo PROMO", "UTM campaign...14",
             "Promo Code", "Nº Tickets", "Asistencia", "Precio Ticket", "Descuento", "Recaudación", "Tipo Pago",
             "Fee", "Status", "Unique User", "New User...25", "Num Events", "Num Events Asist", "Show Deleted", "Fecha Compra",
             "Comprador", "Contacto", "Teléfono", "Idioma", "Evento", "Fecha evento", "Fecha asistencia", "Servicio", "Auth_code",
             "Bpremium_Code", "Nº Pedido", "Registered", "New User...42", "UTM source", "UTM medium", "UTM campaign...45", "Creador",
             "Pagado banco >0", "FECHA COMPRA", "MES COMPRA", "AÑO COMPRA", "FECHA EVENTO", "MES EVENTO", "AÑO EVENTO", "DIA", 
             "Fórmula para ajuste Domingos en la semana_1", "SEMANA", "Ventas para Eventos fuera del mes de compra", "GRUPO", "NEGOCIO_1",
             "KAM Negocio", "GRUPO PROMOTOR", "Promotor_2", "COMERCIAL BUSINESS", "COMERCIAL NEGOCIO", "KAM WL", "CANAL DE VENTA", 
             "Comisión Red de Promoción", "LIQUIDACIÓN_1", "AUTOFACTURA", "Comisión Promotores", "PRECIO TICKET_1", "Precio Ticket + Gasto gestión",
             "Margen Promoción", "Total margen", "Negocio está en pestaña Datos Maestros Generales", "WL está en pestaña Datos Maestros Generales", 
             "LOCKDOWN", "PRECIO TICKET_2", "NEGOCIO&CODE AUTH", "DEVOLUCIONES CHECK", "DIA DE LA SEMANA", "DIA Compra", 
             "Fórmula para ajuste Domingos en la semana_2", "SEMANA Compra", "CLIENTE QUE VENDE", "LIQUIDACIÓN_2", "TIPO CLIENTE",
             "FECHA EVENTO BIEN", "AÑO PORTFOLIO", "Comercial Negocio", "Comercial WL", "Pagado?", "Remesa") 


colnames(output_vicent_renamed) <- c(col_names)
compare_df_cols_same(output_vicent_renamed, output_generated)

result <- anti_join(output_generated, output_vicent_renamed)

column_name <-  "Comercial WL" 
identical(output_generated[[column_name]],output_vicent_renamed[[column_name]])



result <- anti_join(output_generated %>% 
                      dplyr::select("Negocio",`Id Negocio`,`Id KAM Negocio`,"Comercial WL"),
                    output_vicent_renamed %>% 
                      dplyr::select("Negocio",`Id Negocio`,`Id KAM Negocio`,"Comercial WL")) %>% 
  unique()

result_1 <- anti_join(output_generated %>% 
                      dplyr::select("Negocio",`Id Negocio`),
                    output_vicent_renamed %>% 
                      dplyr::select("Negocio",`Id Negocio`)) %>% 
  unique()


nrow(result)



# Aligning formats
cols_character <- c("CLIENTE QUE VENDE","DEVOLUCIONES CHECK", "Id KAM Negocio", "ID KAM WL", "Id Negocio", "LOCKDOWN",
                    "Promo Code", "Remesa", "UTM campaign...14", "UTM campaign...45", "UTM medium", "MES COMPRA", "MES EVENTO", "Nº Pedido")

output_vicent_renamed[cols_character] <- sapply(output_vicent_renamed[cols_character],as.character)
output_generated[cols_character] <- sapply(output_generated[cols_character],as.character)


cols_numeric <- c("Comisión Promotores", "Comisión Red de Promoción", "LIQUIDACIÓN_1", "Margen Promoción", "Total margen",
                  "Fórmula para ajuste Domingos en la semana_1", "Fórmula para ajuste Domingos en la semana_2")

output_vicent_renamed[cols_numeric] <- sapply(output_vicent_renamed[cols_numeric],as.numeric)
output_generated[cols_numeric] <- sapply(output_generated[cols_numeric],as.numeric)

cols_date <- c("FECHA COMPRA", "FECHA EVENTO", "FECHA EVENTO BIEN" )

# output_vicent_renamed[cols_date] <- sapply(output_vicent_renamed[cols_date],as.Date(format = "%m/%d/%Y"))
# output_generated[cols_date] <- sapply(output_generated[cols_date],as.Date(format = "%m/%d/%Y"))
# # output_vicent_renamed$`FECHA COMPRA`  <- as.Date(output_vicent_renamed$`FECHA COMPRA`, format = "%m/%d/%Y" )

output_vicent_renamed <- output_vicent_renamed %>% mutate_at(vars(`FECHA COMPRA`,`FECHA EVENTO`,`FECHA EVENTO BIEN`), as.Date, format= "%m/%d/%Y",origin = "1964-10-22")
output_generated <- output_generated %>% mutate_at(vars(`FECHA COMPRA`,`FECHA EVENTO`,`FECHA EVENTO BIEN`), as.Date, format= "%m/%d/%Y", origin = "1964-10-22")

library(janitor)
All.list <- list(output_vicent_renamed,output_generated)
compare_df_cols(All.list)

# install.packages("sqldf")
# require(sqldf)
# output_vicent_renamed
# result <- sqldf::sqldf('SELECT * FROM output_generated EXCEPT SELECT * FROM output_vicent_renamed')
# nrow(result)
# result <- output_generated[!(do.call(paste, output_generated) %in% do.call(paste, output_vicent_renamed)),]
# 
# check_master_data_wl <- master_data %>% 
#   select(WL_1, Liquidación) %>% unique() %>% 
#   group_by(WL_1) %>% 
#   count()
  
#  resaved versions-----


output_vicent_resaved <- read_excel("ignore/output_vicent_resaved.xlsx") 
output_generated_resaved <- read_excel("ignore/output_generated_resaved.xlsx")


All.list <- list(output_generated_resaved,output_vicent_resaved)
compare_df_cols(All.list)
