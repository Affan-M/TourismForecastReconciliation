# Function for API call from the Database

Query_db <- function(series_id, frequency = NULL, include_incomplete = FALSE, force_convert = FALSE,
                     force_conversion_method = NULL) {
  
  ## Function description: Uses query database API to get the data, provides frequency conversion.
  ## In the case of incomplete data, it would also allow you to aggregate up to the most recent period,
  ## or to drop the incomplete period.
  
  ## Function arguments:
  # series_id:  Series IDs of the series you want. Easiest to check series IDs using Viya.
  #             Can be a vector or a single series ID.
  # 
  # frequency:  Specify the frequency you would like the data in. Options for frequency are "Quarterly" and "Annual".
  #             Default is NULL, in which case the data is provided at base website frequency. In the following
  #             cases, there would be a warning provided and a dataframe with base frequency would be provided:
  #               - The data is not convertible
  #               - The data is already in the frequency you want (i.e. available in Quarterly frequency
  #                 and you asked for it to be converted to Quarterly frequency)
  #               - The data is not available in a higher frequency (i.e. available in Quarterly frequency
  #                 and you asked for it to be converted to Monthly frequency)
  # 
  # include_incomplete: Specify whether you would like the aggregate of incomplete periods to be included or dropped.
  #                     By default, only complete quarters or years will be shown. In the case that you want to see
  #                     Year to Date data for example, you would have to specify TRUE for include_incomplete.
  #                     
  # force_convert:      Even though Query database has deemed it inconvertible, you can ask it to be force converted.
  #                     However, you need to provide a conversion method in this case. FALSE by default.
  #                     Not recommended unless certain.
  # 
  # force_coversion_method: Provides a conversion method to the function even though database sees the series as
  #                         inconvertible or has another method specified. Not recommended unless certain.
  #                         Default is NULL. Possible options for this argument are:
  #                           - "sum":  Aggregation by summing
  #                           - "avg":  Aggregation by finding the average
  #                           - "ep":   Aggregation by finding end of the period value
  
  
  # Setting the authorization header. Change bearer token to your API given by MMA database website.
  bearer_token <- api_key
  auth_headers <- c(Authorization = glue("Bearer {bearer_token}"))
  
  # Combining links with series IDs
  collapsed_series_id <- paste(series_id, collapse = ",")
  link <- glue("https://database.mma.gov.mv/api/series?ids={collapsed_series_id}") %>% as.character()
  
  # Requesting data
  data_list <- list()
  
  page = 1
  
  data_list[[page]] <- httr::GET(url = link, httr::add_headers(.headers = auth_headers), httr::config(ssl_verifypeer = FALSE)) %>%
    httr::content(as = "text", encoding = "UTF-8") %>%
    jsonlite::fromJSON()
  
  next_page <- data_list[[page]][["links"]][["next"]]
  
  while (!is.null(next_page)) {
    
    page = page + 1
    
    data_list[[page]] <- httr::GET(url = link, httr::add_headers(.headers = auth_headers), httr::config(ssl_verifypeer = FALSE)) %>%
      httr::content(as = "text", encoding = "UTF-8") %>%
      jsonlite::fromJSON()
    
    next_page <- data_list[[page]][["links"]][["next"]]
    
  }
  
  # Processing
  sorted_list <- list()
  
  for (p in 1:length(data_list)) {
    
    id <- data_list[[p]][["data"]][["id"]]
    
    for (i in id) {
      
      id_tibble <- data_list[[p]][["data"]] %>% filter(id == i)
      
      # Checks whether it is annual, monthly/quarterly or inconvertible
      sort_freq <- id_tibble %>% pull(frequency)
      sort_convertible <- id_tibble %>% pull(freq_convertible)
      
      if (sort_freq == "Annual") {
        
        id_tibble
        
        sorted_list[["Annual"]] <- c(sorted_list[["Annual"]], i)
        
      } else if (sort_freq != "Annual" & sort_convertible == FALSE) {
        
        sorted_list[["NotAnnual_Inconvertible"]] <- c(sorted_list[["NotAnnual_Inconvertible"]], i)
        
      } else if (sort_freq != "Annual" & sort_convertible == TRUE) {
        
        sorted_list[["NotAnnual_Convertible"]] <- c(sorted_list[["NotAnnual_Convertible"]], i)
        
      }
      
    }
    
  }
  
  processed_list <- list()
  meta_info <- list()
  
  for (s in 1:length(sorted_list)) {
    
    id <- sorted_list[[s]]
    
    for (i in 1:length(id)) {
      
      # Processing
      selected_id <- id[1]
      
      
    }
    
  }
  
  for (p in 1:length(data_list)) {
    
    id <- data_list[[p]][["data"]][["id"]]
    
    for (i in 1:length(id)) {
      
      # Processing
      selected_id <- id[i]
      name <- data_list[[p]][["data"]][["name"]][i] %>% str_squish()
      
      id_name <- glue("{selected_id} - {name}") %>% as.character()
      
      data_frequency <- data_list[[p]][["data"]][["frequency"]][i]
      
      if (data_frequency == "Monthly") {
        
        df <- data_list[[p]][["data"]][["data"]][[i]] %>%
          as_tibble() %>%
          mutate(date = as.yearmon(date),
                 name = name,
                 id = selected_id)
        
      } else if (data_frequency == "Quarterly") {
        
        df <- data_list[[p]][["data"]][["data"]][[i]] %>%
          as_tibble() %>%
          mutate(date = as.yearmon(date) %>% as.yearqtr(),
                 name = name,
                 id = selected_id)
        
      } else if (data_frequency == "Annual") {
        
        df <- data_list[[p]][["data"]][["data"]][[i]] %>%
          as_tibble() %>%
          mutate(date = as.yearmon(date) %>% year(),
                 name = name,
                 id = selected_id)
        
      }
      
      if (!is.null(frequency)) {
        
        ### Frequency conversion
        conversion_table <- tibble("Frequency" = c("Monthly", "Quarterly", "Annual"),
                                   "Frequency_numerical" = c(12, 4, 1))
        
        data_frequency_numerical <- conversion_table %>%
          filter(Frequency == data_list[[p]][["data"]][["frequency"]][i]) %>%
          pull(Frequency_numerical)
        
        frequency_required_numerical <- conversion_table %>%
          filter(Frequency == frequency) %>%
          pull(Frequency_numerical)
        
        # Checking if frequency required is lower than the frequency the data is in
        if (frequency_required_numerical < data_frequency_numerical) {
          
          ## Year to date check
          last_year <- df %>% pull(date) %>% year() %>% unique() %>% max() %>% as.character()
          check_dates <- df %>% filter(str_detect(date, last_year)) %>% pull(date)
          
          # Checking force_convert option
          if (force_convert == FALSE) {
            
            convertible <- data_list[[p]][["data"]][["freq_convertible"]][[i]] 
            
          } else if (force_convert == TRUE) {
            
            if (!is.null(force_conversion_method)) {
              
              if (force_conversion_method %in% c("sum", "avg", "ep")) {
                
                convertible <- TRUE 
                
              } else {
                
                stop("A valid force_conversion_method not provided. Please provide either \"sum\", \"avg\" or \"ep\" for force_conversion_method if you wish to force_convert.")
                
              }
              
            } else {
              
              stop("No force_conversion_method provided. Please provide either \"sum\", \"avg\" or \"ep\" as a method if you wish to force_convert.")
              
            }
            
          }
          
          # Checking if it is convertible
          if (convertible == TRUE) {
            
            # Checking force_conversion_method
            if (is.null(force_conversion_method)) {
              
              conversion_method <- data_list[[p]][["data"]][["conversion_method"]][i] 
              
            } else {
              
              conversion_method <- force_conversion_method
              
            }
            
            # Aggregating data to the target frequency Quarterly
            if (frequency == "Quarterly") {
              
              # Aggregation based on conversion method
              if (conversion_method == "sum") {
                
                converted_df <- df %>%
                  mutate(date = as.yearqtr(date)) %>%
                  group_by(date, name, id) %>%
                  summarise(amount = sum(amount), .groups = "keep") %>%
                  ungroup()
                
              } else if (conversion_method == "avg") {
                
                converted_df <- df %>%
                  mutate(date = as.yearqtr(date)) %>%
                  group_by(date, name, id) %>%
                  summarise(amount = mean(amount), .groups = "keep") %>%
                  ungroup()
                
              } else if (conversion_method == "ep") {
                
                df <- df %>% mutate(qtr = as.yearqtr(date))
                
                dates <- df %>% distinct(qtr) %>% pull()
                
                converted_data <- list()
                
                for (d in dates) {
                  
                  int_df <- df %>%
                    filter(qtr == d) %>%
                    arrange(date) %>%
                    tail(n = 1)
                  
                  d2 <- as.character(d)
                  
                  converted_data[[d2]] <- int_df
                  
                }
                
                converted_df <- bind_rows(converted_data) %>% select(!date) %>% rename(date = qtr) %>%
                  relocate(date)
                
              }
              
              
              if (length(check_dates) %% 3 != 0 & include_incomplete == FALSE) {
                
                drop_period <- converted_df %>% pull(date) %>% unique() %>% max() %>% as.character()
                converted_df <- converted_df %>% filter(!str_detect(date, drop_period))
                
              }
              
              # Aggregating data to the target frequency Annual
            } else if (frequency == "Annual") {
              
              # Aggregation based on conversion method
              if (conversion_method == "sum") {
                
                converted_df <- df %>%
                  mutate(date = year(date)) %>%
                  group_by(date, name, id) %>%
                  summarise(amount = sum(amount), .groups = "keep") %>%
                  ungroup()
                
              } else if (conversion_method == "avg") {
                
                converted_df <- df %>%
                  mutate(date = year(date)) %>%
                  group_by(date, name, id) %>%
                  summarise(amount = mean(amount), .groups = "keep") %>%
                  ungroup()
                
              } else if (conversion_method == "ep") {
                
                df <- df %>% mutate(year = year(date))
                
                dates <- df %>% distinct(year) %>% pull()
                
                converted_data <- list()
                
                for (d in dates) {
                  
                  converted_data[[d]] <- df %>%
                    filter(year == d) %>%
                    arrange(date) %>%
                    tail(n = 1)
                  
                }
                
                converted_df <- bind_rows(converted_data) %>% select(!date) %>% rename(date = year) %>%
                  relocate(date)
                
              }
              
              # Dropping the most recent period if incomplete and include_incomplete is not required (Monthly data)
              if (data_frequency == "Monthly" & length(check_dates) %% 12 != 0 & include_incomplete == FALSE) {
                
                drop_period <- converted_df %>% pull(date) %>% unique() %>% max() %>% as.character()
                converted_df <- converted_df %>% filter(!str_detect(date, drop_period))
                
              }
              
              # Dropping the most recent period if incomplete and include_incomplete is not required (Quarterly data)
              if (data_frequency == "Quarterly" & length(check_dates) %% 4 != 0 & include_incomplete == FALSE) {
                
                drop_period <- converted_df %>% pull(date) %>% unique() %>% max() %>% as.character()
                converted_df <- converted_df %>% filter(!str_detect(date, drop_period))
                
              }
              
            }
            
            df <- converted_df
            
          } else {
            
            print(glue("WARNING: Series ID: {id_name} is not convertible. Providing data without conversion in the resulting dataframe."))
            
          }
          
        } else {
          
          print(glue("WARNING: Series ID: {id_name} is either already {frequency} data or cannot be converted to a higher frequency. Providing data without conversion in the resulting dataframe."))
          
        }
        
      }
      
      processed_list[[id_name]] <- df %>% arrange(date)
      
      meta_info[[id_name]] <- tibble("ID" = selected_id,
                                     "Name" = name,
                                     "Original_Frequency" = data_frequency,
                                     "Conversion_Method" = ifelse(exists("conversion_method"),
                                                                  conversion_method,
                                                                  data_list[[p]][["data"]][["conversion_method"]][i]),
                                     "Converted_Frequency" = ifelse(!is.null(frequency), frequency, NA))
      
    }
    
  }
  
  final_df <- bind_rows(processed_list) %>%
    select(date, id, name, amount)
  
  meta_df <- bind_rows(meta_info)
  
  final_list <- list("Data" = final_df,
                     "Meta" = meta_df)
  
  return(final_list)
  
}
