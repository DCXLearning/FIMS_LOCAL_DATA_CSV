
download_excel_report <- function(
    input, output,
    filtered_reports_data,
    filtered_nat_data,
    filtered_aqu_table_data,
    filtered_processing_data,
    filtered_summary_data,
    summaryProductionChart,
    aquaProductionChart,
    filtered_aqu_data
) {
  output$download_excel_report <- downloadHandler(
    filename = function() {
      paste0("Fisheries_Report_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      # Show a modal dialog while generating the report
      showModal(modalDialog(
        title = "Generating Excel Report",
        "Please wait, your Excel report is being generated...",
        footer = NULL
      ))
      # Ensure modal is removed when done or if error
      on.exit(removeModal())
      
      # Create a new workbook
      wb <- createWorkbook()
      
      # --- TABLES ---
      
      if ("reportStatusTable" %in% input$dl_choices) {
        # Data preparation for "Total Number of Statistical Reports Submitted by Month"
        excel_data_report_status <- isolate(filtered_reports_data()) %>%
          group_by(province_en, MonthAbbr) %>%
          summarise(n = n(), .groups = 'drop') %>%
          pivot_wider(names_from = MonthAbbr, values_from = n, values_fill = 0)
        
        all_months_ordered <- c("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
        for(m in all_months_ordered){
          if(!m %in% names(excel_data_report_status)){
            excel_data_report_status[[m]] <- 0
          }
        }
        excel_data_report_status <- excel_data_report_status %>%
          select(province_en, all_of(all_months_ordered)) %>%
          mutate(Months = rowSums(across(all_of(all_months_ordered), ~ . > 0))) %>%
          rename(Province = province_en) %>%
          arrange(Province)
        
        provinces_submitting_row <- excel_data_report_status %>%
          summarise(across(all_of(all_months_ordered), ~ sum(. > 0, na.rm = TRUE))) %>%
          mutate(Province = "Provinces Submitting",
                 Months = NA)
        
        final_excel_report_status <- bind_rows(excel_data_report_status, provinces_submitting_row)
        
        addWorksheet(wb, "Report Status")
        writeData(wb, "Report Status", "Total Number of Statistical Reports Submitted by Month", startCol = 1, startRow = 1)
        writeData(wb, "Report Status", final_excel_report_status, startCol = 1, startRow = 2, rowNames = FALSE)
        addStyle(wb, sheet = "Report Status", style = createStyle(textDecoration = "bold"), rows = 1, cols = 1)
        setColWidths(wb, "Report Status", cols = 1:ncol(final_excel_report_status), widths = "auto")
      }
      
      if ("natPivotTable" %in% input$dl_choices) {
        # Data preparation for "Fish Catch by Quarter"
        excel_data_nat_pivot <- isolate(filtered_nat_data()) %>%
          filter(natfishcatch_typefishing != "3")
        if (nrow(excel_data_nat_pivot) > 0) {
          excel_nat_summary <- excel_data_nat_pivot %>%
            group_by(fish_name_kh, fish_name_en, Quarter) %>%
            summarise(Total = sum(catch, na.rm = TRUE), .groups = "drop") %>%
            pivot_wider(names_from = Quarter, values_from = Total, values_fill = 0)
          for (q in c("Q1", "Q2", "Q3", "Q4")) {
            if (!q %in% names(excel_nat_summary)) {
              excel_nat_summary[[q]] <- 0
            }
          }
          total_catch_nat <- sum(excel_nat_summary[c("Q1", "Q2", "Q3", "Q4")], na.rm = TRUE)
          excel_nat_summary <- excel_nat_summary %>%
            mutate(Total_Catch = Q1 + Q2 + Q3 + Q4) %>%
            mutate(`Contribution (%)` = if(total_catch_nat > 0) round((Total_Catch / total_catch_nat) * 100, 2) else 0) %>%
            arrange(desc(Total_Catch))
          grand_total_row_nat <- tibble(
            fish_name_kh = "Grand Total", fish_name_en = "",
            Q1 = sum(excel_nat_summary$Q1), Q2 = sum(excel_nat_summary$Q2),
            Q3 = sum(excel_nat_summary$Q3), Q4 = sum(excel_nat_summary$Q4),
            Total_Catch = total_catch_nat, `Contribution (%)` = 100.00
          )
          final_excel_nat_table <- bind_rows(excel_nat_summary, grand_total_row_nat) %>%
            rename(`Khmer Name` = fish_name_kh, `English Name` = fish_name_en)
          
          addWorksheet(wb, "Fish Catch by Quarter")
          writeData(wb, "Fish Catch by Quarter", "Fish Catch by Quarter", startCol = 1, startRow = 1)
          writeData(wb, "Fish Catch by Quarter", final_excel_nat_table, startCol = 1, startRow = 2, rowNames = FALSE)
          addStyle(wb, sheet = "Fish Catch by Quarter", style = createStyle(textDecoration = "bold"), rows = 1, cols = 1)
          setColWidths(wb, "Fish Catch by Quarter", cols = 1:ncol(final_excel_nat_table), widths = "auto")
        } else {
          addWorksheet(wb, "Fish Catch by Quarter")
          writeData(wb, "Fish Catch by Quarter", "No data available for this table with current filters.")
        }
      }
      
      if ("riceFieldTable" %in% input$dl_choices) {
        # Data preparation for "Fisheries Capture in Rice-Fields by Species"
        excel_data_rice_field <- isolate(filtered_nat_data()) %>% filter(natfishcatch_typefishing == "3")
        if (nrow(excel_data_rice_field) > 0) {
          excel_rice_summary <- excel_data_rice_field %>%
            group_by(fish_name_kh, fish_name_en, Quarter) %>%
            summarise(Quarter_Total = sum(catch, na.rm = TRUE), .groups = "drop") %>%
            pivot_wider(names_from = Quarter, values_from = Quarter_Total, values_fill = 0)
          for (q in c("Q1", "Q2", "Q3", "Q4")) {
            if (!q %in% names(excel_rice_summary)) {
              excel_rice_summary[[q]] <- 0
            }
          }
          total_production_rice <- sum(excel_rice_summary[c("Q1", "Q2", "Q3", "Q4")], na.rm = TRUE)
          excel_rice_summary <- excel_rice_summary %>%
            mutate(Total = Q1 + Q2 + Q3 + Q4) %>%
            mutate(`Contribution (%)` = if(total_production_rice > 0) round((Total / total_production_rice) * 100, 2) else 0) %>%
            arrange(desc(Total))
          grand_total_row_rice <- tibble(
            fish_name_kh = "Grand Total", fish_name_en = "",
            Q1 = sum(excel_rice_summary$Q1), Q2 = sum(excel_rice_summary$Q2),
            Q3 = sum(excel_rice_summary$Q3), Q4 = sum(excel_rice_summary$Q4),
            Total = total_production_rice, `Contribution (%)` = 100.00
          )
          final_excel_rice_table <- bind_rows(excel_rice_summary, grand_total_row_rice) %>%
            rename(`Khmer Name` = fish_name_kh, `English Name` = fish_name_en)
          
          addWorksheet(wb, "Rice-Field Catch by Species")
          writeData(wb, "Rice-Field Catch by Species", "Fisheries Capture in Rice-Fields by Species", startCol = 1, startRow = 1)
          writeData(wb, "Rice-Field Catch by Species", final_excel_rice_table, startCol = 1, startRow = 2, rowNames = FALSE)
          addStyle(wb, sheet = "Rice-Field Catch by Species", style = createStyle(textDecoration = "bold"), rows = 1, cols = 1)
          setColWidths(wb, "Rice-Field Catch by Species", cols = 1:ncol(final_excel_rice_table), widths = "auto")
        } else {
          addWorksheet(wb, "Rice-Field Catch by Species")
          writeData(wb, "Rice-Field Catch by Species", "No data available for this table with current filters.")
        }
      }
      
      if ("riceFieldProvinceTable" %in% input$dl_choices) {
        # Data preparation for "Freshwater Capture in Rice-Fields by Province"
        excel_data_rice_province <- isolate(filtered_nat_data()) %>% filter(natfishcatch_typefishing == "3")
        if (nrow(excel_data_rice_province) > 0) {
          excel_summary_rice_province <- excel_data_rice_province %>%
            group_by(province_kh, province_en) %>%
            summarise(`Production (MT)` = sum(catch, na.rm = TRUE), .groups = "drop")
          total_rice_catch_province <- sum(excel_summary_rice_province$`Production (MT)`, na.rm = TRUE)
          excel_summary_rice_province <- excel_summary_rice_province %>%
            mutate(`Contribution (%)` = if(total_rice_catch_province > 0) round((`Production (MT)` / total_rice_catch_province) * 100, 2) else 0) %>%
            arrange(desc(`Production (MT)`))
          grand_total_row_rice_province <- tibble(
            province_kh = "Grand Total", province_en = "",
            `Production (MT)` = total_rice_catch_province, `Contribution (%)` = 100.00
          )
          final_excel_rice_province_table <- bind_rows(excel_summary_rice_province, grand_total_row_rice_province) %>%
            rename(`Province (Khmer)` = province_kh, `Province (English)` = province_en)
          
          addWorksheet(wb, "Rice-Field Catch by Prov")
          writeData(wb, "Rice-Field Catch by Prov", "Freshwater Capture in Rice-Fields by Province", startCol = 1, startRow = 1)
          writeData(wb, "Rice-Field Catch by Prov", final_excel_rice_province_table, startCol = 1, startRow = 2, rowNames = FALSE)
          addStyle(wb, sheet = "Rice-Field Catch by Prov", style = createStyle(textDecoration = "bold"), rows = 1, cols = 1)
          setColWidths(wb, "Rice-Field Catch by Prov", cols = 1:ncol(final_excel_rice_province_table), widths = "auto")
        } else {
          addWorksheet(wb, "Rice-Field Catch by Prov")
          writeData(wb, "Rice-Field Catch by Prov", "No data available for this table with current filters.")
        }
      }
      
      if ("provinceProductionTable" %in% input$dl_choices) {
        # Data preparation for "Inland Production by Province"
        excel_data_province_prod <- isolate(filtered_nat_data()) %>% filter(natfishcatch_typefishing != "3")
        if (nrow(excel_data_province_prod) > 0) {
          excel_summary_province_prod <- excel_data_province_prod %>%
            group_by(province_kh, province_en) %>%
            summarise(`Total Catch` = sum(catch, na.rm = TRUE), .groups = "drop")
          total_catch_province_prod <- sum(excel_summary_province_prod$`Total Catch`, na.rm = TRUE)
          excel_summary_province_prod <- excel_summary_province_prod %>%
            mutate(`Contribution (%)` = if(total_catch_province_prod > 0) round((`Total Catch` / total_catch_province_prod) * 100, 2) else 0) %>%
            arrange(desc(`Total Catch`))
          grand_total_row_province_prod <- tibble(
            province_kh = "Grand Total", province_en = "",
            `Total Catch` = total_catch_province_prod, `Contribution (%)` = 100.00
          )
          final_excel_province_prod_table <- bind_rows(excel_summary_province_prod, grand_total_row_province_prod) %>%
            rename(`Province (Khmer)` = province_kh, `Province (English)` = province_en)
          
          addWorksheet(wb, "Inland Prod by Prov")
          writeData(wb, "Inland Prod by Prov", "Inland Production by Province", startCol = 1, startRow = 1)
          writeData(wb, "Inland Prod by Prov", final_excel_province_prod_table, startCol = 1, startRow = 2, rowNames = FALSE)
          addStyle(wb, sheet = "Inland Prod by Prov", style = createStyle(textDecoration = "bold"), rows = 1, cols = 1)
          setColWidths(wb, "Inland Prod by Prov", cols = 1:ncol(final_excel_province_prod_table), widths = "auto")
        } else {
          addWorksheet(wb, "Inland Prod by Prov")
          writeData(wb, "Inland Prod by Prov", "No data available for this table with current filters.")
        }
      }
      
      if ("aquaProvinceProductionTable" %in% input$dl_choices) {
        # Data preparation for "Aquaculture Production by Province"
        excel_data_aqua_province <- isolate(filtered_aqu_table_data())
        if (nrow(excel_data_aqua_province) > 0) {
          excel_summary_aqua_province <- excel_data_aqua_province %>%
            group_by(province_kh, province_en) %>%
            summarise(`Total Aquaculture Production (MT)` = sum(catch, na.rm = TRUE), .groups = "drop")
          total_aqu_prod_province <- sum(excel_summary_aqua_province$`Total Aquaculture Production (MT)`, na.rm = TRUE)
          excel_summary_aqua_province <- excel_summary_aqua_province %>%
            mutate(`Contribution (%)` = if(total_aqu_prod_province > 0) round((`Total Aquaculture Production (MT)` / total_aqu_prod_province) * 100, 2) else 0) %>%
            arrange(desc(`Total Aquaculture Production (MT)`))
          grand_total_row_aqua_province <- tibble(
            province_kh = "Grand Total", province_en = "",
            `Total Aquaculture Production (MT)` = total_aqu_prod_province, `Contribution (%)` = 100.00
          )
          final_excel_aqua_province_table <- bind_rows(excel_summary_aqua_province, grand_total_row_aqua_province) %>%
            rename(`Province (Khmer)` = province_kh, `Province (English)` = province_en)
          
          addWorksheet(wb, "Aqua Prod by Prov")
          writeData(wb, "Aqua Prod by Prov", "Aquaculture Production by Province", startCol = 1, startRow = 1)
          writeData(wb, "Aqua Prod by Prov", final_excel_aqua_province_table, startCol = 1, startRow = 2, rowNames = FALSE)
          addStyle(wb, sheet = "Aqua Prod by Prov", style = createStyle(textDecoration = "bold"), rows = 1, cols = 1)
          setColWidths(wb, "Aqua Prod by Prov", cols = 1:ncol(final_excel_aqua_province_table), widths = "auto")
        } else {
          addWorksheet(wb, "Aqua Prod by Prov")
          writeData(wb, "Aqua Prod by Prov", "No data available for this table with current filters.")
        }
      }
      
      if ("processingTable" %in% input$dl_choices) {
        # Data preparation for "Processed Fish Production by Quarter"
        excel_data_processing <- isolate(filtered_processing_data())
        if (nrow(excel_data_processing) > 0) {
          df_converted_excel <- excel_data_processing %>%
            mutate(
              processing_amount = case_when(
                processing_typeprocessing %in% c("11", "26") ~ processing_amount / 1000,
                TRUE ~ processing_amount
              ),
              proc_name_kh = case_when(
                processing_typeprocessing %in% c("11", "26") ~ paste0(proc_name_kh, " (ជាតោន)"),
                TRUE ~ proc_name_kh
              ),
              proc_name_en = case_when(
                processing_typeprocessing %in% c("11", "26") ~ paste0(proc_name_en, " (tons)"),
                TRUE ~ proc_name_en
              )
            )
          normal_df_excel <- df_converted_excel %>% filter(!processing_typeprocessing %in% c("11", "26"))
          special_df_excel <- df_converted_excel %>% filter(processing_typeprocessing %in% c("11", "26"))
          top_15_normal_excel <- tibble()
          total_normal_production_excel <- 0
          if(nrow(normal_df_excel) > 0) {
            normal_summary_excel <- normal_df_excel %>%
              group_by(proc_name_kh, proc_name_en, Quarter) %>%
              summarise(Quarter_Total = sum(processing_amount, na.rm = TRUE), .groups = "drop") %>%
              pivot_wider(names_from = Quarter, values_from = Quarter_Total, values_fill = 0)
            for (q in c("Q1", "Q2", "Q3", "Q4")) {
              if (!q %in% names(normal_summary_excel)) {
                normal_summary_excel[[q]] <- 0
              }
            }
            top_15_normal_excel <- normal_summary_excel %>%
              mutate(Total = Q1 + Q2 + Q3 + Q4) %>%
              arrange(desc(Total)) %>%
              slice_head(n = 15)
            total_normal_production_excel <- sum(top_15_normal_excel$Total, na.rm = TRUE)
            top_15_normal_excel <- top_15_normal_excel %>%
              mutate(`Contribution (%)` = if(total_normal_production_excel > 0) round((Total / total_normal_production_excel) * 100, 2) else 0)
          }
          special_summary_excel <- tibble()
          if (nrow(special_df_excel) > 0) {
            special_summary_excel <- special_df_excel %>%
              group_by(proc_name_kh, proc_name_en, Quarter) %>%
              summarise(Quarter_Total = sum(processing_amount, na.rm = TRUE), .groups = "drop") %>%
              pivot_wider(names_from = Quarter, values_from = Quarter_Total, values_fill = 0)
            for (q in c("Q1", "Q2", "Q3", "Q4")) {
              if (!q %in% names(special_summary_excel)) {
                special_summary_excel[[q]] <- 0
              }
            }
            special_summary_excel <- special_summary_excel %>%
              mutate(Total = Q1 + Q2 + Q3 + Q4, `Contribution (%)` = 0)
          }
          grand_total_row_processing <- tibble(
            proc_name_kh = "Grand Total", proc_name_en = "",
            Q1 = sum(top_15_normal_excel$Q1, na.rm = TRUE),
            Q2 = sum(top_15_normal_excel$Q2, na.rm = TRUE),
            Q3 = sum(top_15_normal_excel$Q3, na.rm = TRUE),
            Q4 = sum(top_15_normal_excel$Q4, na.rm = TRUE),
            Total = total_normal_production_excel,
            `Contribution (%)` = 100.00
          )
          final_excel_processing_table <- bind_rows(top_15_normal_excel, grand_total_row_processing, special_summary_excel) %>%
            rename(`Khmer Name` = proc_name_kh, `English Name` = proc_name_en)
          
          addWorksheet(wb, "Processed Fish Prod")
          writeData(wb, "Processed Fish Prod", "Processed Fish Production by Quarter", startCol = 1, startRow = 1)
          writeData(wb, "Processed Fish Prod", final_excel_processing_table, startCol = 1, startRow = 2, rowNames = FALSE)
          addStyle(wb, sheet = "Processed Fish Prod", style = createStyle(textDecoration = "bold"), rows = 1, cols = 1)
          setColWidths(wb, "Processed Fish Prod", cols = 1:ncol(final_excel_processing_table), widths = "auto")
        } else {
          addWorksheet(wb, "Processed Fish Prod")
          writeData(wb, "Processed Fish Prod", "No data available for this table with current filters.")
        }
      }
      
      # --- CHARTS ---
      
      if ("summaryProductionChart" %in% input$dl_choices) {
        # Data preparation for "Freshwater, Marine and Aquaculture Production Chart"
        excel_chart_data_summary <- isolate(filtered_summary_data())
        
        # Apply the same data processing as in the reactive chart for consistency
        if (nrow(excel_chart_data_summary) > 0) {
          # Ensure area_group is created and factored for consistency
          if (!"area_group" %in% names(excel_chart_data_summary)) {
            excel_chart_data_summary <- excel_chart_data_summary %>%
              mutate(area_group = case_when(
                # Adjust this mapping based on your actual data and how 'filtered_summary_data' is structured
                natfishcatch_typefishing == "1" ~ "Freshwater Capture",
                natfishcatch_typefishing == "2" ~ "Marine Capture",
                natfishcatch_typefishing == "4" ~ "Aquaculture",
                TRUE ~ "Other Category"
              ))
          }
          excel_chart_data_summary$area_group <- factor(excel_chart_data_summary$area_group,
                                                        levels = c("Freshwater Capture", "Marine Capture", "Aquaculture", "Total production"))
          
          # Further aggregation if 'plot_data' was already aggregated in the chart reactive
          # The provided 'summaryProductionChart' code takes `plot_data <- filtered_summary_data()`
          # and then uses `plot_data` directly in ggplot. If `filtered_summary_data()`
          # already gives you Year, catch, and a column to map to area_group,
          # then the above is sufficient. If you do further `group_by`/`summarise`
          # within the reactive, you need to replicate that here.
          # Example (if filtered_summary_data returns raw data and you group/summarize inside the chart reactive):
          excel_chart_data_summary <- excel_chart_data_summary %>%
            group_by(Year, area_group) %>%
            summarise(TotalCatch = sum(catch, na.rm = TRUE), .groups = 'drop')
          
          
          addWorksheet(wb, "Summary Prod Data")
          writeData(wb, "Summary Prod Data", "Freshwater, Marine and Aquaculture Production Data", startCol = 1, startRow = 1)
          writeData(wb, "Summary Prod Data", excel_chart_data_summary, startCol = 1, startRow = 2, rowNames = FALSE)
          addStyle(wb, sheet = "Summary Prod Data", style = createStyle(textDecoration = "bold"), rows = 1, cols = 1)
          setColWidths(wb, "Summary Prod Data", cols = 1:ncol(excel_chart_data_summary), widths = "auto")
        } else {
          addWorksheet(wb, "Summary Prod Data")
          writeData(wb, "Summary Prod Data", "No data available for this chart's table with current filters.")
        }
      }
      
      # if ("aquaProductionChart" %in% input$dl_choices) {
      #   # Data preparation for "Aquaculture and Fingerling Production Chart"
      #   df_aqua <- isolate(filtered_aqu_data()) # Get the initial data frame
      #   
      #   if (nrow(df_aqua) > 0) {
      #     excel_chart_data_aqua <- df_aqua %>%
      #       group_by(Year) %>%
      #       summarise(
      #         `Aquaculture Production (MT)` = sum(catch, na.rm = TRUE),
      #         `Fingerling Production (heads)` = sum(fingerling_production, na.rm = TRUE), # Ensure this column exists in filtered_aqu_data
      #         .groups = "drop"
      #       ) %>%
      #       arrange(Year) # Order by year for better table readability
      #     
      #     addWorksheet(wb, "Aquaculture Data")
      #     writeData(wb, "Aquaculture Data", "Aquaculture and Fingerling Production Data", startCol = 1, startRow = 1)
      #     writeData(wb, "Aquaculture Data", excel_chart_data_aqua, startCol = 1, startRow = 2, rowNames = FALSE)
      #     addStyle(wb, sheet = "Aquaculture Data", style = createStyle(textDecoration = "bold"), rows = 1, cols = 1)
      #     setColWidths(wb, "Aquaculture Data", cols = 1:ncol(excel_chart_data_aqua), widths = "auto")
      #   } else {
      #     addWorksheet(wb, "Aquaculture Data")
      #     writeData(wb, "Aquaculture Data", "No data available for this chart's table with current filters.")
      #   }
      # }
      
      # Save the workbook
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}