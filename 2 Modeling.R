# Packages
library(tidyverse) # Main data manipulation package
library(httr) # Web requests (used for API)
library(jsonlite) # JSON (used for API)
library(glue) # String manipulation
library(zoo) # Date manipulation
library(lubridate) # Date manipulation
library(readxl) # Read excel files
library(openxlsx) # Write excel files
library(fpp3) # Forecast package

rm(list=ls())

source("1 API function.R")

# Put "yes" or "no" here to run the required functions in the script
# Forecast would provide forecasts for all scenarios and exported to the forecast excel file
forecast <- "yes"

# Prediction intervals required. Can put any number in the vector.
hilo_vec <- c(60, 80)

# Run accuracy test would run both test and training set tests for one step ahead forecasts
run_accuracy_test <- "yes"

# This is computationally intensive, so the option is provided. Automatically put to "no" unless required,
# since most decisions are taken based on test set accuracy and this is less important.
run_training_set_accuracy <- "no"

# Additional controls for the accuracy test:
# Where the first test set starts. Setting a default to 3 years from today's January (Jan 2019 in this case).
# Change as required to get a bigger/smaller experiment.
test_start <- 3

# Whether to take top n countries or until a specific percentage is fulfilled
# Commands: "top_n" or "perc"
method <- "perc"

# Specify the number of countries to take if method is "top_n
n_countries <- 20

# Specify the percentage to be fulfilled if method is "perc"
selected_perc <- 60

# Specify the number of most recent years to take. Either "all" or a number.
n_years <- 4

#### Models and Excel theme ####
headerstyle <- createStyle(fontName = "Roboto", fontSize = 10,
                   fgFill = "#0d1d4f", halign = "CENTER", textDecoration = "Bold", fontColour = "white",
                   border = c("top", "left", "right", "bottom"), borderStyle = "medium", borderColour = "black")

# Heading style
style1 <- createStyle(fontName = "Roboto", fontSize = 10, fontColour = "black", numFmt = "#,##0", textDecoration = "bold")

# Body style
style2 <- createStyle(fontName = "Roboto", fontSize = 10, fontColour = "black", numFmt = "#,##0")

# Body style: With decimals
style3 <- createStyle(fontName = "Roboto", fontSize = 10, fontColour = "black", numFmt = "#,##0.00")


## Forecast methods
Models_desc_df <- tibble("Models" = c("ARIMA", "ETS", "DHR", "COMB1","COMB2", "COMB3", "COMB4"),
                         "Description" = c("Autoregressive Integrated Moving Average models",
                                           "Exponential smoothing (Error-Trend-Seasonality models)",
                                           "Dynamic Harmonic Regression",
                                           "Combination of ARIMA and ETS",
                                           "Combination of ARIMA and DHR",
                                           "Combination of ETS and DHR",
                                           "Combination of ARIMA, ETS and DHR"),
                         "Learn_more" = c("https://otexts.com/fpp3/arima.html",
                                          "https://otexts.com/fpp3/expsmooth.html",
                                          "https://otexts.com/fpp3/dhr.html",
                                          "https://otexts.com/fpp3/combinations.html",
                                          "", "", ""))

Reconciliation_desc_df <- tibble("Reconciliation" = c("Bottom-up (BU)",
                                                      "Ordinary Least Squares (OLS)",
                                                      "Weighted Least Squares (WLS)",
                                                      "Minimum Trace (MinT)"),
                                 "Description" = c("Forecasts of bottom level series added to get top level.",
                                                   "Variances restricted to 1 and covariances to 0 in the var-cov matrix",
                                                   "No restriction on variance, covariances restricted to 0 in the var-cov matrix.",
                                                   "No restriction on variance, covariances shrunk towards zero in the var-cov matrix."),
                                 "Learn_more" = c("https://otexts.com/fpp3/hierarchical.html",
                                                  "", "", ""))

#### Data ####
meta_data <- read_csv("meta-data.csv", show_col_types = FALSE)

continents <- c(106, 121, 129, 138, 146, 152, 158, 167, 174, 179, 186, 197)
series_ids <- meta_data %>% filter(`Parent ID` %in% continents | `Series ID` %in% c(104, 150, 178, 202, 203)) %>%
  pull(`Series ID`)

query_file <- Query_db(series_id = series_ids)
df1 <- query_file[["Data"]] %>%
  mutate(name = str_remove(name, "Tourist arrivals from | tourist arrivals") %>% str_trim())

# # Check whether sum of the individual countries add up to total tourist arrivals
# full_join(df1 %>% filter(id == 104) %>% select(date, amount) %>% rename(total = amount),
#           df1 %>% filter(id != 104) %>% group_by(date) %>% summarise(sum = sum(amount)), by = "date") %>%
#   mutate(check = total - sum) %>%
#   filter(check > 0) %>%
#   nrow()

## Creating a hierarchy

# Number of years to take into consideration
if (n_years == "all") {
  
  selected_years <- df1 %>% pull(date) %>% year() %>% unique() 
  
} else {
  
  selected_years <- df1 %>% mutate(date = year(date)) %>% distinct(date) %>% arrange(date) %>% tail(n = n_years) %>% pull()
  
}

# Taking the major countries
if (method == "top_n") {
  
  df_countries <- df1 %>%
    filter(str_detect(date, paste0(selected_years, collapse = "|")) & id != 104) %>%
    group_by(id, name) %>%
    summarise(amount = sum(amount), .groups = "keep") %>%
    arrange(-amount) %>%
    head(n = n_countries)
  
} else if (method == "perc") {
  
  df2 <- df1 %>%
    filter(str_detect(date, paste0(selected_years, collapse ="|"))) %>%
    group_by(id, name) %>%
    summarise(amount = sum(amount), .groups = "drop")
  
  total <- df2 %>% filter(id == 104) %>% pull(amount)
  
  df_countries <- df2 %>%
    filter(id != 104) %>%
    mutate(perc = amount/total) %>%
    arrange(desc(perc)) %>%
    mutate(cumulative = cumsum(perc)) %>%
    filter(cumulative <= selected_perc/100)
  
}

dfx <- df1 %>% filter(name %in% c("Total", pull(df_countries, name)))

# Calculating Others to complete hierarchy
other_df <- full_join(dfx %>% filter(name == "Total") %>%
                        pivot_wider(id_cols = "date", names_from = "name", values_from = "amount"),
                      dfx %>% filter(name != "Total") %>%
                        group_by(date) %>%
                        summarise(Selected = sum(amount, na.rm = TRUE), .groups = "drop"),
                      by = "date") %>%
  mutate(Other = Total - Selected) %>%
  select(date, Other) %>%
  pivot_longer(cols = !date, names_to = "name", values_to = "amount")

# pivot_wider + pivot_longer ensures that all missing dates are filled with 0's
df_complete <- bind_rows(dfx, other_df) %>%
  pivot_wider(id_cols = date, names_from = name, values_from = amount, values_fill = 0) %>%
  pivot_longer(cols = !date, names_to = "name", values_to = "amount")

# Check to see if it all adds up
check_df <- full_join(df_complete %>% filter(name == "Total") %>% pivot_wider(id_cols = date, names_from = name, values_from = amount),
                      df_complete %>% filter(name != "Total") %>%
                        group_by(date) %>%
                        summarise(Subseries = sum(amount), .groups = "drop"), by = "date") %>%
  mutate(Difference = Total - Subseries) %>%
  filter(Difference > 0)

if (nrow(check_df) > 0) {
  
  print(check_df)
  
  stop("Issues in the hierarchy.")
  
} else {
  
  print("No issues in the hierarchy. Proceeding.")
  
}

## Preparing tsib
tsib <- df_complete %>% filter(name != "Total") %>%
  mutate(date = yearmonth(date)) %>%
  as_tsibble(index = date, key = name) %>%
  aggregate_key(name, amount = sum(amount))

date_limit <- tsib %>% distinct(date) %>% pull() %>%
  .[(length(.)-236)] %>%
  year() %>%
  paste("January", .) %>%
  yearmonth()

tsib <- tsib %>% filter(date >= date_limit)


#### Forecast ####

if (forecast == "yes") {
  
  print("Computing forecasts.")
  
  # Calculating H. Check the last month in the training set. If December, set forecast steps ahead to 12.
  # Otherwise, calculate the remaining months for the end of the year.
  train_end_date <- tsib %>% pull(date) %>% max() %>% as.Date() %>% as.yearmon()
  
  if (train_end_date == "Dec 2022") {
    
    H <- 12
    
  } else {
    
    year_end_date <- paste("Dec", year(train_end_date)) %>% as.yearmon()
    H <- ((year_end_date - train_end_date) * 12) %>% round(., digits = 0)
    
  }
  
  # Forecast
  
  # Checking for the Fourier K that minimizes AICc
  fit_dhr <- tsib %>% model("1" = ARIMA(amount ~ fourier(K = 1) + PDQ(0,0,0)),
                            "2" = ARIMA(amount ~ fourier(K = 2) + PDQ(0,0,0)),
                            "3" = ARIMA(amount ~ fourier(K = 3) + PDQ(0,0,0)),
                            "4" = ARIMA(amount ~ fourier(K = 4) + PDQ(0,0,0)),
                            "5" = ARIMA(amount ~ fourier(K = 5) + PDQ(0,0,0)),
                            "6" = ARIMA(amount ~ fourier(K = 6) + PDQ(0,0,0)))
  
  k <- glance(fit_dhr) %>%
    group_by(.model) %>%
    summarise(AICc = mean(AICc)) %>%
    arrange(AICc) %>%
    head(1) %>%
    pull(.model) %>%
    as.numeric()
  
  # Fitting the data for the models and doing the forecasts  
  fit <- tsib %>%
    model(ARIMA = ARIMA(amount),
          ETS = ETS(amount),
          DHR = ARIMA(amount ~ pdq() + PDQ(0,0,0) + fourier(K = k))) %>%
    mutate(COMB1 = (ARIMA + ETS)/2,
           COMB2 = (ARIMA + DHR)/2,
           COMB3 = (ETS + DHR)/2,
           COMB4 = (ARIMA + ETS + DHR)/3) %>%
    reconcile(ARIMA_bu = bottom_up(ARIMA),
              ARIMA_ols = min_trace(ARIMA, "ols"),
              ARIMA_wls = min_trace(ARIMA, "wls_var"),
              ARIMA_mint = min_trace(ARIMA, "mint_shrink"),
              ETS_bu = bottom_up(ETS),
              ETS_ols = min_trace(ETS, "ols"),
              ETS_wls = min_trace(ETS, "wls_var"),
              ETS_mint = min_trace(ETS, "mint_shrink"),
              DHR_bu = bottom_up(DHR),
              DHR_ols = min_trace(DHR, "ols"),
              DHR_wls = min_trace(DHR, "wls_var"),
              DHR_mint = min_trace(DHR, "mint_shrink"),
              COMB1_bu = bottom_up(COMB1),
              COMB1_ols = min_trace(COMB1, "ols"),
              COMB1_wls = min_trace(COMB1, "wls_var"),
              COMB1_mint = min_trace(COMB1, "mint_shrink"),
              COMB2_bu = bottom_up(COMB2),
              COMB2_ols = min_trace(COMB2, "ols"),
              COMB2_wls = min_trace(COMB2, "wls_var"),
              COMB2_mint = min_trace(COMB2, "mint_shrink"),
              COMB3_bu = bottom_up(COMB3),
              COMB3_ols = min_trace(COMB3, "ols"),
              COMB3_wls = min_trace(COMB3, "wls_var"),
              COMB3_mint = min_trace(COMB3, "mint_shrink"),
              COMB4_bu = bottom_up(COMB4),
              COMB4_ols = min_trace(COMB4, "ols"),
              COMB4_wls = min_trace(COMB4, "wls_var"),
              COMB4_mint = min_trace(COMB4, "mint_shrink"))
  
  fc <- forecast(fit, h = H, bootstap = TRUE) %>% hilo(hilo_vec)
  
  hilo_list <- list()
  
  for (pi in hilo_vec) {
    
    pi_colname <- paste0(pi, "%")
    
    hilo_dfx <- fc %>% as_tibble() %>%
      select(name, .model, date, all_of(pi_colname)) %>%
      rename(selected_pi = all_of(pi_colname))
    
    hilo_list[[glue("Upper {pi}")]] <- hilo_dfx %>%
      mutate("amount" = selected_pi$upper,
             "interval" = glue("Upper {pi}"),
             "order" = pi) %>%
      select(!selected_pi)
    
    hilo_list[[glue("Lower {pi}")]] <- hilo_dfx %>%
      mutate("amount" = selected_pi$lower,
             "interval" = glue("Lower {pi}"),
             order = -pi) %>%
      select(!selected_pi)
    
  }
  
  hilo_df <- bind_rows(hilo_list)
  
  # Combining actuals and mean forecasts
  models <- fc %>% pull(.model) %>% unique()
  
  fc_aggregates <- list()
  
  for (m in models) {
    
    fc_aggregates[[m]] <- bind_rows(tsib %>% as_tibble(),
                                    fc %>% as_tibble() %>%
                                      filter(.model == m) %>%
                                      select(date, name, .mean) %>%
                                      rename(amount = .mean)) %>%
      mutate(model = m, .before = name)
    
  }
  
  fc_combined <- bind_rows(fc_aggregates) %>%
    mutate(date = date %>% as.Date() %>% as.yearmon())
  
  # Taking adjustments to ensure no negative forecasts are there. Adjusted from the country total,
  # and adjusted by reducing it.
  for (m in models) {
    
    names <- fc_combined %>% filter(model == m & amount < 0) %>% distinct(name) %>% pull() %>% as.character()
    
    if (length(names) > 0) {
      
      start_date <- fc_combined %>% filter(model == m & is_aggregated(name)) %>% tail(H) %>% slice(1) %>% pull(date)
      
      for (n in names) {
        
        # Adjusting negative series
        issue_df <- fc_combined %>% filter(model == m & name == n) %>% tail(H)
        
        amount <- issue_df %>% pull(amount)
        initial_sum <- sum(amount)
        
        amount[amount < 0] <- 0
        modified_sum <- sum(amount)
        
        if (modified_sum == 0) {
          
          new_amount <- amount
          
        } else {
          
          new_amount <- amount/modified_sum * initial_sum
          new_amount[new_amount < 0] <- 0 
          
        }
        
        fc_combined <- fc_combined %>%
          mutate(amount = replace(amount, model == m & name == n & date >= start_date, new_amount))
      
      }
    
    }
    
    reconciled_models <- models[str_detect(models, "bu|ols|wls|mint")]
    
    if (m %in% reconciled_models) {
      
      # Adjusting totals for each month with the changes to the negative series
      dates <- fc_combined %>% filter(model == m & date >= start_date) %>% distinct(date) %>% pull()
      
      for (d in dates) {
        
        total <- fc_combined %>% filter(model == m & is_aggregated(name) & date == d) %>% pull(amount)
        
        subseries_df <- fc_combined %>% filter(model == m & !is_aggregated(name) & date == d)
        subseries_sum <- subseries_df %>% pull(amount) %>% sum()
        
        modified_df <- subseries_df %>% mutate(amount = amount/subseries_sum * total)
        
        fc_combined <- bind_rows(fc_combined %>% filter(!(model == m & !is_aggregated(name) & date == d)),
                                 modified_df)
        
      }
      
    }
    
  }
  
  # Annual forecast
  year <- fc_combined %>% pull(date) %>% year() %>% max()
  
  fc_annual <- fc_combined %>%
    filter(is_aggregated(name)) %>%
    mutate(date = year(date),
           name = "Total") %>%
    group_by(date, model) %>%
    summarise(amount = sum(amount), .groups = "drop") %>%
    filter(date %in% c((year-4):year)) %>%
    pivot_wider(id_cols = model, names_from = date, values_from = amount) %>%
    arrange(!!rlang::sym(as.character(year)))
  
  # Monthly forecast
  
  # Total
  fc_monthly <- fc_combined %>%
    filter(is_aggregated(name)) %>%
    mutate(name = "Total") %>%
    filter(date >= glue("Jan {year}") %>% as.yearmon()) %>%
    pivot_wider(id_cols = model, names_from = date, values_from = amount) %>%
    full_join(., fc_annual %>% select(model, !!rlang::sym(as.character(year))),
              by = "model") %>%
    arrange(!!rlang::sym(as.character(year)))
  
  
  ### Exporting results to Excel
  wb <- createWorkbook()
  
  ## Legend
  sh <- addWorksheet(wb, "Legend")
  
  start_row <- 2
  
  writeData(wb, sh, x = "Model information", startCol = 2, startRow = start_row)
  addStyle(wb, sh, style = style1, cols = 2, rows = start_row, stack = TRUE)
  
  start_row <- start_row + 1
  
  writeData(wb, sh, x = Models_desc_df, startCol = 2, startRow = start_row,
            borders = "columns", borderStyle = "medium", borderColour = "black",
            headerStyle = headerstyle)
  
  addStyle(wb, sh,
           style = style2,
           rows = (start_row + 1):(start_row + 1 + nrow(Models_desc_df)),
           cols = 3:(3 + ncol(Models_desc_df)),
           gridExpand = TRUE,
           stack = TRUE)
  
  start_row <- start_row + nrow(Models_desc_df) + 2
  
  writeData(wb, sh, x = "Reconciliation information", startCol = 2, startRow = start_row)
  addStyle(wb, sh, style = style1, cols = 2, rows = start_row, stack = TRUE)
  
  start_row <- start_row + 1
  
  writeData(wb, sh, x = Reconciliation_desc_df, startCol = 2, startRow = start_row,
            borders = "columns", borderStyle = "medium", borderColour = "black",
            headerStyle = headerstyle)
  
  addStyle(wb, sh,
           style = style2,
           rows = (start_row + 1):(start_row + 1 + nrow(Reconciliation_desc_df)),
           cols = 3:(3 + ncol(Reconciliation_desc_df)),
           gridExpand = TRUE,
           stack = TRUE)
  
  setColWidths(wb, sh, cols = (2:(2+max(ncol(Models_desc_df), ncol(Reconciliation_desc_df))-1)),
               widths = "auto", hidden = FALSE)
  
  ## Forecasts
  sh <- addWorksheet(wb, "FC Totals")
  
  # Annual
  start_row <- 2
  
  writeData(wb, sh, x = "Annual", startCol = 2, startRow = start_row)
  addStyle(wb, sh, style = style1, cols = 2, rows = start_row, stack = TRUE)
  
  start_row <- start_row + 1
  
  writeData(wb, sh, x = fc_annual, startCol = 2, startRow = start_row,
            borders = "columns", borderStyle = "medium", borderColour = "black",
            headerStyle = headerstyle)
  
  addStyle(wb, sh,
           style = style2,
           rows = (start_row + 1):(start_row + 1 + nrow(fc_annual)),
           cols = 3:(3 + ncol(fc_annual)),
           gridExpand = TRUE,
           stack = TRUE)
  
  # Monthly
  start_row <- start_row + 1 + nrow(fc_annual) + 2
  
  writeData(wb, sh, x = "Monthly", startCol = 2, startRow = start_row)
  
  addStyle(wb, sh, style = style1, cols = 2, rows = start_row, stack = TRUE)
  
  start_row <- start_row + 1
  
  writeData(wb, sh, x = fc_monthly, startCol = 2, startRow = start_row,
            borders = "columns", borderStyle = "medium", borderColour = "black",
            headerStyle = headerstyle)
  
  addStyle(wb, sh,
           style = style2,
           rows = (start_row + 1):(start_row + 1 + nrow(fc_monthly)),
           cols = 3:(3 + ncol(fc_monthly)),
           gridExpand = TRUE,
           stack = TRUE)
  
  # Monthly forecasts for each model
  for (m in models) {
    
    ## Preparing dataframe
    
    # Mean forecasts
    fc_modelx <- fc_combined %>%
      filter(model == m & date >= glue("Jan {year}") %>% as.yearmon()) %>%
      pivot_wider(id_cols = name, names_from = date, values_from = amount) %>%
      mutate(name = replace(name, is_aggregated(name), "Total")) %>%
      mutate(name = as.character(name))
    
    annual_total <- fc_modelx %>%
      pivot_longer(cols = !name, names_to = "date", values_to = "amount") %>%
      group_by(name) %>%
      summarise(Total = sum(amount)) %>%
      arrange(desc(Total))
    
    order <- bind_rows(filter(annual_total, name != "Other"),
                       filter(annual_total, name == "Other")) %>%
      pull(name)
    
    fc_model <- full_join(fc_modelx, annual_total, by = "name") %>%
      arrange(match(name, order))
    
    # Exporting to a new sheet
    sh <- addWorksheet(wb, m)
    
    start_row <- 2
    
    label <- glue("{m}: Mean forecasts") %>% as.character()
    
    writeData(wb, sh, x = label, startCol = 2, startRow = start_row)
    
    addStyle(wb, sh, style = style1, cols = 2, rows = start_row, stack = TRUE)
    
    start_row <- start_row + 1
    
    writeData(wb, sh, x = fc_model, startCol = 2, startRow = start_row,
              borders = "columns", borderStyle = "medium", borderColour = "black",
              headerStyle = headerstyle)
    
    addStyle(wb, sh,
             style = style2,
             rows = (start_row + 1):(start_row + 1 + nrow(fc_model)),
             cols = 3:(3 + ncol(fc_model)),
             gridExpand = TRUE,
             stack = TRUE)
    
    start_row <- start_row + 1 + nrow(fc_model)
    
    # Prediction intervals
    intervals <- hilo_df %>% distinct(interval, order) %>% arrange(desc(order)) %>% pull(interval)
    
    for (i in intervals) {
      
      # Preparing prediction interval dataframe
      fc_pi_dfx1 <- bind_rows(tsib %>% as_tibble(),
                             hilo_df %>%
                               filter(.model == m & interval == i) %>%
                               select(date, name, amount)) %>%
        mutate(date = as.Date(date) %>% as.yearmon()) %>%
        filter(date >= glue("Jan {year}") %>% as.yearmon())
      
      # Taking adjustments to ensure no negative forecasts are there. Adjusted from the country total,
      # and adjusted by reducing it.
      names <- fc_pi_dfx1 %>% filter(amount < 0) %>% distinct(name) %>% pull() %>% as.character()
      
      if (length(names) > 0) {
        
        start_date <- fc_pi_dfx1 %>% filter(is_aggregated(name)) %>% tail(H) %>% slice(1) %>% pull(date)
        
        for (n in names) {
          
          # Adjusting negative series
          issue_df <- fc_pi_dfx1 %>% filter(name == n) %>% tail(H)
          
          amount <- issue_df %>% pull(amount)
          initial_sum <- sum(amount)
          
          amount[amount < 0] <- 0
          modified_sum <- sum(amount)
          
          if (modified_sum == 0) {
            
            new_amount <- amount
            
          } else {
            
            new_amount <- amount/modified_sum * initial_sum
            new_amount[new_amount < 0] <- 0 
            
          }
          
          fc_pi_dfx1 <- fc_pi_dfx1 %>%
            mutate(amount = replace(amount, name == n & date >= start_date, new_amount))
          
        }
        
      }
      
      # Adjusting total for each month with the changes to the negative series
      reconciled_models <- models[str_detect(models, "bu|ols|wls|mint")]
      
      if (m %in% reconciled_models) {
        
        dates <- fc_pi_dfx1 %>% filter(date >= start_date) %>% distinct(date) %>% pull()
        
        for (d in dates) {
          
          total <- fc_pi_dfx1 %>% filter(is_aggregated(name) & date == d) %>% pull(amount)
          
          subseries_df <- fc_pi_dfx1 %>% filter(!is_aggregated(name) & date == d)
          subseries_sum <- subseries_df %>% pull(amount) %>% sum()
          
          modified_df <- subseries_df %>% mutate(amount = amount/subseries_sum * total)
          
          fc_pi_dfx1 <- bind_rows(fc_pi_dfx1 %>% filter(!(!is_aggregated(name) & date == d)),
                                  modified_df)
          
        }
        
      }
      
      fc_pi_dfx <- fc_pi_dfx1 %>%
        pivot_wider(id_cols = name, names_from = date, values_from = amount) %>%
        mutate(name = replace(name, is_aggregated(name), "Total")) %>%
        mutate(name = as.character(name))
      
      annual_total <- fc_pi_dfx %>%
        pivot_longer(cols = !name, names_to = "date", values_to = "amount") %>%
        group_by(name) %>%
        summarise(Total = sum(amount)) %>%
        arrange(desc(Total))
      
      order <- bind_rows(filter(annual_total, name != "Other"),
                         filter(annual_total, name == "Other")) %>%
        pull(name)
      
      fc_pi_df <- full_join(fc_pi_dfx, annual_total, by = "name") %>%
        arrange(match(name, order))
      
      # Exporting to Excel
      start_row <- start_row + 2
      
      label <- glue("{m}: {i}% forecasts") %>% as.character()
      
      writeData(wb, sh, x = label, startCol = 2, startRow = start_row)
      
      addStyle(wb, sh, style = style1, cols = 2, rows = start_row, stack = TRUE)
      
      start_row <- start_row + 1
      
      writeData(wb, sh, x = fc_pi_df, startCol = 2, startRow = start_row,
                borders = "columns", borderStyle = "medium", borderColour = "black",
                headerStyle = headerstyle)
      
      addStyle(wb, sh,
               style = style2,
               rows = (start_row + 1):(start_row + 1 + nrow(fc_pi_df)),
               cols = 3:(3 + ncol(fc_pi_df)),
               gridExpand = TRUE,
               stack = TRUE)
      
      start_row <- start_row + 1 + nrow(fc_pi_df)
      
    }
    
  }
  
  saveWorkbook(wb, file = "Forecast.xlsx", overwrite = TRUE)
  
  print("Forecasts completed.")
  
}

#### Cross validation test ####

if (run_accuracy_test == "yes") {
  
  print("Starting cross-validation.")
  
  # Cross validation tsibble - default one year ahead
  H <- 12
  
  max_date <- tsib %>% pull(date) %>% max() %>% as.Date() %>% as.yearmon()
  first_test <- paste("January", year(max_date) - test_start) %>% as.yearmon()
  window_number <- (first_test - (tsib %>% pull(date) %>% min() %>% as.Date() %>% as.yearmon())) * 12
  
  test_end <- max_date - H/12
  
  cv_tsib <- tsib %>%
    filter(date <= yearmonth(test_end)) %>%
    stretch_tsibble(.step = 1, .init = window_number) %>%
    relocate(date, .id)
  
  # Calculating a second H based on steps ahead until December for the year.
  # Calculation would be check the last month in the training set.
  # If December, set forecast steps ahead to 12.
  # Otherwise, calculate the remaining months for the end of the year.
  train_end_date <- tsib %>% pull(date) %>% max() %>% as.Date() %>% as.yearmon()

  if (train_end_date == "Dec 2022") {

    H_2 <- 12

  } else {

    year_end_date <- paste("Dec", year(train_end_date)) %>% as.yearmon()
    H_2 <- ((year_end_date - train_end_date) * 12) %>% round(., digits = 0)

  }
  
  # test_end <- max_date - H_2/12
  # 
  # cv_tsib2 <- tsib %>%
  #   filter(date <= yearmonth(test_end)) %>%
  #   stretch_tsibble(.step = 1, .init = window_number) %>%
  #   relocate(date, .id)
  
  # Checking for the Fourier K that minimizes AICc
  fit_dhr <- tsib %>% model("1" = ARIMA(amount ~ fourier(K = 1) + PDQ(0,0,0)),
                            "2" = ARIMA(amount ~ fourier(K = 2) + PDQ(0,0,0)),
                            "3" = ARIMA(amount ~ fourier(K = 3) + PDQ(0,0,0)),
                            "4" = ARIMA(amount ~ fourier(K = 4) + PDQ(0,0,0)),
                            "5" = ARIMA(amount ~ fourier(K = 5) + PDQ(0,0,0)),
                            "6" = ARIMA(amount ~ fourier(K = 6) + PDQ(0,0,0)))
  
  k <- glance(fit_dhr) %>%
    group_by(.model) %>%
    summarise(AICc = mean(AICc)) %>%
    arrange(AICc) %>%
    head(1) %>%
    pull(.model) %>%
    as.numeric()
  
  ## Accuracy tests
  fit <- cv_tsib %>%
    model(ARIMA = ARIMA(amount),
          ETS = ETS(amount),
          DHR = ARIMA(amount ~ pdq() + PDQ(0,0,0) + fourier(K = k))) %>%
    mutate(COMB1 = (ARIMA + ETS)/2,
           COMB2 = (ARIMA + DHR)/2,
           COMB3 = (ETS + DHR)/2,
           COMB4 = (ARIMA + ETS + DHR)/3) %>%
    reconcile(ARIMA_bu = bottom_up(ARIMA),
              ARIMA_ols = min_trace(ARIMA, "ols"),
              ARIMA_wls = min_trace(ARIMA, "wls_var"),
              ARIMA_mint = min_trace(ARIMA, "mint_shrink"),
              ETS_bu = bottom_up(ETS),
              ETS_ols = min_trace(ETS, "ols"),
              ETS_wls = min_trace(ETS, "wls_var"),
              ETS_mint = min_trace(ETS, "mint_shrink"),
              DHR_bu = bottom_up(DHR),
              DHR_ols = min_trace(DHR, "ols"),
              DHR_wls = min_trace(DHR, "wls_var"),
              DHR_mint = min_trace(DHR, "mint_shrink"),
              COMB1_bu = bottom_up(COMB1),
              COMB1_ols = min_trace(COMB1, "ols"),
              COMB1_wls = min_trace(COMB1, "wls_var"),
              COMB1_mint = min_trace(COMB1, "mint_shrink"),
              COMB2_bu = bottom_up(COMB2),
              COMB2_ols = min_trace(COMB2, "ols"),
              COMB2_wls = min_trace(COMB2, "wls_var"),
              COMB2_mint = min_trace(COMB2, "mint_shrink"),
              COMB3_bu = bottom_up(COMB3),
              COMB3_ols = min_trace(COMB3, "ols"),
              COMB3_wls = min_trace(COMB3, "wls_var"),
              COMB3_mint = min_trace(COMB3, "mint_shrink"),
              COMB4_bu = bottom_up(COMB4),
              COMB4_ols = min_trace(COMB4, "ols"),
              COMB4_wls = min_trace(COMB4, "wls_var"),
              COMB4_mint = min_trace(COMB4, "mint_shrink"))
  
  # Training set
  if (run_training_set_accuracy == "yes") {
    
    train_accuracy_list <- list()
    
    models <- names(fit) %>% .[3:length(.)]
    
    print("Training set accuracy")
    counter <- 0
    
    for (m in models) {
      
      counter <- counter + 1
      print(glue("{m} : {counter}/{length(models)}"))
      
      train_accuracy_list[[m]] <- fit %>% select(all_of(m)) %>% accuracy()
      
    }
    
    train_accuracy_df <- bind_rows(train_accuracy_list)
    
    train_accuracy_total <- train_accuracy_df %>%
      filter(is_aggregated(name)) %>%
      group_by(.model) %>%
      summarise(across(.cols = !c("name", ".id", ".type"), .fns = ~ mean(.x, na.rm = TRUE))) %>%
      mutate(series = "Total", .before = ".model") %>%
      arrange(RMSSE)
    
    train_accuracy_subseries <- train_accuracy_df %>%
      filter(!is_aggregated(name)) %>%
      group_by(.model) %>%
      summarise(across(.cols = !c("name", ".id", ".type"), .fns = ~ mean(.x, na.rm = TRUE))) %>%
      mutate(series = "Subseries", .before = ".model") %>%
      arrange(RMSSE)
    
    train_accuracy_all <- train_accuracy_df %>%
      group_by(.model) %>%
      summarise(across(.cols = !c("name", ".id", ".type"), .fns = ~ mean(.x, na.rm = TRUE))) %>%
      mutate(series = "All", .before = ".model") %>%
      arrange(RMSSE)
    
  }
  
  # Test set
  print("Test set accuracy")
  
  if (H_2 == 12) {
    
    fc <- forecast(fit, h = H)
    
    test_accuracy_df <- accuracy(fc, tsib)
    
    test_accuracy_total <- test_accuracy_df %>%
      filter(is_aggregated(name)) %>%
      mutate("series" = "Total",
             "H" = H,
             .before = ".model") %>%
      select(!c(name, .type)) %>%
      arrange(RMSSE)
    
    test_accuracy_subseries <- test_accuracy_df %>%
      filter(!is_aggregated(name)) %>%
      group_by(.model) %>%
      summarise(across(.cols = !c("name", ".type"), .fns = ~ mean(.x, na.rm = TRUE))) %>%
      mutate("series" = "Subseries",
             "H" = H, .before = ".model") %>%
      arrange(RMSSE)
    
    test_accuracy_all <- test_accuracy_df %>%
      group_by(.model) %>%
      summarise(across(.cols = !c("name", ".type"), .fns = ~ mean(.x, na.rm = TRUE))) %>%
      mutate("series" = "All",
             "H" = H, .before = ".model") %>%
      arrange(RMSSE)
    
  } else {
    
    fc_list <- list()
    
    fc_list[["1"]] <- forecast(fit, h = H)
    fc_list[["2"]] <- forecast(fit, h = H_2)
    
    test_accuracy_list <- list()
    
    for (f in names(fc_list)) {
      
      test_accuracy_df <- accuracy(fc_list[[f]], tsib)
      
      test_accuracy_list[["Total"]][[f]] <- test_accuracy_df %>%
        filter(is_aggregated(name)) %>%
        mutate(series = "Total",
               "H" = ifelse(f == "1", H, H_2),
               .before = ".model") %>%
        select(!c(name, .type))
      
      test_accuracy_list[["Subseries"]][[f]] <- test_accuracy_df %>%
        filter(!is_aggregated(name)) %>%
        group_by(.model) %>%
        summarise(across(.cols = !c("name", ".type"), .fns = ~ mean(.x, na.rm = TRUE))) %>%
        mutate(series = "Subseries",
               "H" = ifelse(f == "1", H, H_2),
               .before = ".model")
      
      test_accuracy_list[["All"]][[f]] <- test_accuracy_df %>%
        group_by(.model) %>%
        summarise(across(.cols = !c("name", ".type"), .fns = ~ mean(.x, na.rm = TRUE))) %>%
        mutate(series = "All",
               "H" = ifelse(f == "1", H, H_2),
               .before = ".model")
      
    }
    
    test_accuracy_total <- bind_rows(test_accuracy_list[["Total"]]) %>%
      arrange(H, RMSSE)
    
    test_accuracy_subseries <- bind_rows(test_accuracy_list[["Subseries"]]) %>%
      arrange(H, RMSSE)
    
    test_accuracy_all <- bind_rows(test_accuracy_list[["All"]]) %>%
      arrange(H, RMSSE)
    
  }
  
  
  ## Exporting results to Excel
  wb <- createWorkbook()
  
  # Test set
  sh <- addWorksheet(wb, "Test Accuracy")
  start_row <- 2
  
  writeData(wb, sh, x = "All series", startCol = 2, startRow = start_row)
  
  addStyle(wb, sh, style = style1, cols = 2, rows = start_row, stack = TRUE)
  
  start_row <- start_row + 1
  
  writeData(wb, sh, test_accuracy_all, startCol = 2, startRow = start_row,
            borders = "columns", borderStyle = "medium", borderColour = "black",
            headerStyle = headerstyle)
  
  addStyle(wb, sh,
           style = style3,
           rows = (start_row + 1):(start_row + 2 + nrow(test_accuracy_all)),
           cols = 2:(2 + ncol(test_accuracy_all)),
           gridExpand = TRUE,
           stack = TRUE)
  
  
  start_row <- start_row + 2 + nrow(test_accuracy_all) + 1
  
  writeData(wb, sh, x = "Total", startCol = 2, startRow = start_row)
  
  addStyle(wb, sh, style = style1, cols = 2, rows = start_row, stack = TRUE)
  
  start_row <- start_row + 1
  
  writeData(wb, sh, test_accuracy_total, startCol = 2, startRow = start_row,
            borders = "columns", borderStyle = "medium", borderColour = "black",
            headerStyle = headerstyle)
  
  addStyle(wb, sh,
           style = style3,
           rows = (start_row + 1):(start_row + 2 + nrow(test_accuracy_total)),
           cols = 2:(2 + ncol(test_accuracy_total)),
           gridExpand = TRUE,
           stack = TRUE)
  
  
  start_row <- start_row + 2 + nrow(test_accuracy_total) + 1
  
  writeData(wb, sh, x = "Subseries", startCol = 2, startRow = start_row)
  
  addStyle(wb, sh, style = style1, cols = 2, rows = start_row, stack = TRUE)
  
  start_row <- start_row + 1
  
  writeData(wb, sh, test_accuracy_subseries, startCol = 2, startRow = start_row,
            borders = "columns", borderStyle = "medium", borderColour = "black",
            headerStyle = headerstyle)
  
  addStyle(wb, sh,
           style = style3,
           rows = (start_row + 1):(start_row + 2 + nrow(test_accuracy_subseries)),
           cols = 2:(2 + ncol(test_accuracy_subseries)),
           gridExpand = TRUE,
           stack = TRUE)
  
  
  # Training set
  if (run_training_set_accuracy == "yes") {
    
    sh <- addWorksheet(wb, "Train Accuracy")
    start_row <- 2
    
    writeData(wb, sh, x = "All series", startCol = 2, startRow = start_row)
    
    addStyle(wb, sh, style = style1, cols = 2, rows = start_row, stack = TRUE)
    
    start_row <- start_row + 1
    
    writeData(wb, sh, train_accuracy_all, startCol = 2, startRow = start_row,
              borders = "columns", borderStyle = "medium", borderColour = "black",
              headerStyle = headerstyle)
    
    addStyle(wb, sh,
             style = style3,
             rows = (start_row + 1):(start_row + 2 + nrow(train_accuracy_all)),
             cols = 2:(2 + ncol(train_accuracy_all)),
             gridExpand = TRUE,
             stack = TRUE)
    
    
    start_row <- start_row + 2 + nrow(train_accuracy_all) + 1
    
    writeData(wb, sh, x = "Total", startCol = 2, startRow = start_row)
    
    addStyle(wb, sh, style = style1, cols = 2, rows = start_row, stack = TRUE)
    
    start_row <- start_row + 1
    
    writeData(wb, sh, train_accuracy_total, startCol = 2, startRow = start_row,
              borders = "columns", borderStyle = "medium", borderColour = "black",
              headerStyle = headerstyle)
    
    addStyle(wb, sh,
             style = style3,
             rows = (start_row + 1):(start_row + 2 + nrow(train_accuracy_total)),
             cols = 2:(2 + ncol(train_accuracy_total)),
             gridExpand = TRUE,
             stack = TRUE)
    
    
    start_row <- start_row + 2 + nrow(train_accuracy_total) + 1
    
    writeData(wb, sh, x = "Subseries", startCol = 2, startRow = start_row)
    
    addStyle(wb, sh, style = style1, cols = 2, rows = start_row, stack = TRUE)
    
    start_row <- start_row + 1
    
    writeData(wb, sh, train_accuracy_subseries, startCol = 2, startRow = start_row,
              borders = "columns", borderStyle = "medium", borderColour = "black",
              headerStyle = headerstyle)
    
    addStyle(wb, sh,
             style = style3,
             rows = (start_row + 1):(start_row + 2 + nrow(train_accuracy_subseries)),
             cols = 2:(2 + ncol(train_accuracy_subseries)),
             gridExpand = TRUE,
             stack = TRUE)
    
  }
  
  # Legend
  sh <- addWorksheet(wb, "Legend")
  
  start_row <- 2
  
  writeData(wb, sh, x = "Model information", startCol = 2, startRow = start_row)
  addStyle(wb, sh, style = style1, cols = 2, rows = start_row, stack = TRUE)
  
  start_row <- start_row + 1
  
  writeData(wb, sh, x = Models_desc_df, startCol = 2, startRow = start_row,
            borders = "columns", borderStyle = "medium", borderColour = "black",
            headerStyle = headerstyle)
  
  addStyle(wb, sh,
           style = style2,
           rows = (start_row + 1):(start_row + 1 + nrow(Models_desc_df)),
           cols = 3:(3 + ncol(Models_desc_df)),
           gridExpand = TRUE,
           stack = TRUE)
  
  start_row <- start_row + nrow(Models_desc_df) + 2
  
  writeData(wb, sh, x = "Reconciliation information", startCol = 2, startRow = start_row)
  addStyle(wb, sh, style = style1, cols = 2, rows = start_row, stack = TRUE)
  
  start_row <- start_row + 1
  
  writeData(wb, sh, x = Reconciliation_desc_df, startCol = 2, startRow = start_row,
            borders = "columns", borderStyle = "medium", borderColour = "black",
            headerStyle = headerstyle)
  
  addStyle(wb, sh,
           style = style2,
           rows = (start_row + 1):(start_row + 1 + nrow(Reconciliation_desc_df)),
           cols = 3:(3 + ncol(Reconciliation_desc_df)),
           gridExpand = TRUE,
           stack = TRUE)
  
  setColWidths(wb, sh, cols = (2:(2+max(ncol(Models_desc_df), ncol(Reconciliation_desc_df))-1)),
               widths = "auto", hidden = FALSE)
  
  saveWorkbook(wb, file = "Accuracy.xlsx", overwrite = TRUE)
 
  print("Cross-validation completed.")
   
}

print("Done!")
