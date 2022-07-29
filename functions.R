#------  Packages ------#

if(!require(pacman)){
  
  install.packages("pacman")
}


pacman::p_load(logging,
               config,
               RJDBC,tidyverse
               )


library(RJDBC)
library(tidyverse)
library(openxlsx)
path_to_packages <- sprintf('%s/drivers', getwd())

environ_path <- sprintf('%s/.Renviron', getwd())
environ_vars <- readRenviron(environ_path)

pg_user <- Sys.getenv("USER")
pg_password <- Sys.getenv("PASSWORD")
pg_url <- Sys.getenv("URL")

# **** PATH  **** ----


#  connecting to MYSQL ------

drv <- JDBC("com.mysql.jdbc.Driver",
            "C:/Users/I0433554/OneDrive - Sanofi/Documents/dev/R/premiumguest/drivers/mysql-connector-java-5.1.44-bin.jar",
            identifier.quote="`")

pg_connection <- dbConnect(drv, "jdbc:mysql://premiumguestdb-read.c6gvkmamvc24.eu-west-1.rds.amazonaws.com:3306/bpremium",
                           pg_user, pg_password, 
                           useSSL = FALSE)

paste0()

sprintf()
sprintf("ignore/output_%s.csv",today)


a_test_table <- dbGetQuery(pg_connection, "select *
                                     from bpremium.a_test_table")
