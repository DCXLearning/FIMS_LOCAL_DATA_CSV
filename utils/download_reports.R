library(openxlsx)    # For Excel downloads
library(flextable)   # For Word document formatting
library(officer)     # For Word document creation
library(ggplot2)     # For plots
library(scales)      # For comma formatting in plots
library(dplyr)       # For data manipulation
library(tidyr)       # For pivot_wider
library(tibble)      # For tibble

# download_word_report function
download_word_report <- function(
    input, output,
    filtered_reports_data,
    filtered_nat_data,
    filtered_aqu_table_data,
    filtered_processing_data,
    filtered_summary_data,
    summaryProductionChart,
    aquaProductionChart,
    filtered_aqu_data # Ensure this is passed and available
) {
  
  # --- DOWNLOAD HANDLER ---
  output$download_report <- downloadHandler(
    filename = function() {
      paste0("Fisheries_Report_", Sys.Date(), ".docx")
    },
    content = function(file) {
      # Show a modal dialog while generating the report
      showModal(modalDialog(
        title = "Generating Report",
        "Please wait, your report is being generated...",
        footer = NULL
      ))
      # Ensure modal is removed when done or if error
      on.exit(removeModal())
      
      doc <- read_docx()
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "", style = "Normal")
      # --- 1. Add the Cover Page ---
      # The logic from generate_cover_page_docx is integrated here.
      doc <- body_add_fpar(doc, fpar(
        ftext("KINGDOM OF CAMBODIA", prop = fp_text(font.size = 18,  color = "#1f497d")),
        fp_p = fp_par(text.align = "center")
      ))
      doc <- body_add_fpar(doc, fpar(
        ftext("National Religion King", prop = fp_text(font.size = 16, bold = TRUE, color = "#1f497d")),
        fp_p = fp_par(text.align = "center")
      ))
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "", style = "Normal")
      
      
      logo_path <- "img/Picture1.jpg" # Use forward slashes for file paths in R
      
      # Check if the logo path is a local file
      if (file.exists(logo_path)) {
        actual_logo_path <- logo_path
      } else {
        warning("Local logo file not found at: ", logo_path, ". Skipping logo insertion.")
        actual_logo_path <- NULL # Indicate failure to get logo
      }
      
      # Add a check for logo existence for robustness
      if (!is.null(actual_logo_path) && file.exists(actual_logo_path)) {
        # To align the image left, embed it within an fpar and set text.align = "left"
        doc <- body_add_fpar(doc, fpar(
          external_img(src = actual_logo_path, width = 1.2, height = 1.2), # Image as external_img
          fp_p = fp_par(text.align = "left") # Apply left alignment to the paragraph containing the image
        ))
      } else {
        # This is the modified placeholder line
        doc <- body_add_fpar(doc, fpar(
          ftext("--- [LOGO PLACEHOLDER] ---"),
          fp_p = fp_par(text.align = "left") # Apply left alignment
        ))
      }
      # Ministry and Administration - CORRECTED TO BE CENTERED
      # --- Ministry and Administration ---
      doc <- body_add_fpar(doc, fpar(
        ftext("Ministry of Agriculture Forestry and Fisheries", prop = fp_text(font.size = 12, color = "#1f497d")),
        fp_p = fp_par(text.align = "left")
      ))
      doc <- body_add_fpar(doc, fpar(
        ftext("                         Fisheries Administration", prop = fp_text(font.size = 12, color = "#1f497d")),
        fp_p = fp_par(text.align = "left") # Keep this as left
      ))
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "", style = "Normal")
      
      
      doc <- body_add_fpar(doc, fpar(
        ftext("Cambodia Programme for Sustainable and Inclusive Growth", prop = fp_text(font.size = 14, bold = TRUE, color = "#1f497d")),
        fp_p = fp_par(text.align = "center")
      ))
      doc <- body_add_fpar(doc, fpar(
        ftext("in the Fisheries Sector: Capture Component", prop = fp_text(font.size = 14, bold = TRUE, color = "#1f497d")),
        fp_p = fp_par(text.align = "center")
      ))
      
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "", style = "Normal")
      
      
      doc <- body_add_fpar(doc, fpar(
        ftext(paste0(format(Sys.Date(), "%Y"), " Annual Fisheries Statistical Report"), prop = fp_text(font.size = 16, bold = TRUE, color = "#1f497d")),
        fp_p = fp_par(text.align = "center")
      ))
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "", style = "Normal")
      
      doc <- body_add_fpar(doc, fpar(
        ftext(paste0("December ", format(Sys.Date(), "%Y")),
              prop = fp_text(font.size = 14, color = "#E76F51")),
        fp_p = fp_par(text.align = "center")
      ))
      
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_fpar(doc, fpar(
        ftext("Compiled by the ", prop = fp_text(font.size = 13, color = "#1f497d")),
        ftext("Fisheries Administration ", prop = fp_text(font.size = 13, color = "#1f497d")),
        fp_p = fp_par(text.align = "center")
      ))
      
      # --- Funding Information ---
      doc <- body_add_par(doc, "", style = "Normal") # Add some vertical space
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "", style = "Normal")
      
      # Apply color to "Funded by European Union"
      doc <- body_add_fpar(doc, fpar(
        ftext("Funded by European Union", prop = fp_text(color = "#1f497d")),
        fp_p = fp_par(text.align = "center")
      ))
      
      # Apply color to "ACA/2018/041-466 and ACA/2019/041-594"
      doc <- body_add_fpar(doc, fpar(
        ftext("ACA/2018/041-466 and ACA/2019/041-594", prop = fp_text(color = "#1f497d")),
        fp_p = fp_par(text.align = "center")
      ))
      
      # Add a page break after the cover page
      doc <- body_add_break(doc)
      
      # Add selected items to the report
      if ("reportStatusTable" %in% input$dl_choices) {
        doc <- body_add_par(doc, "Total Number of Statistical Reports Submitted by Month", style = "heading 1")
        # Ensure the data for flextable is not a DT proxy, but the underlying data
        # Note: DT's formatRound and column renaming are applied client-side.
        # For server-side flextable, we need to apply formatting to the raw data.
        # Here, I'm taking the already processed data from the reactive expression.
        ft_data_report_status <- isolate(filtered_reports_data()) %>%
          group_by(province_en, MonthAbbr) %>%
          summarise(n = n(), .groups = 'drop') %>%
          pivot_wider(names_from = MonthAbbr, values_from = n, values_fill = 0)
        
        all_months_ordered <- c("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
        # Ensure all months columns exist, fill with 0 if not
        for(m in all_months_ordered){
          if(!m %in% names(ft_data_report_status)){
            ft_data_report_status[[m]] <- 0
          }
        }
        ft_data_report_status <- ft_data_report_status %>%
          select(province_en, all_of(all_months_ordered)) %>%
          mutate(Months = rowSums(across(all_of(all_months_ordered), ~ . > 0))) %>%
          rename(Province = province_en) %>%
          arrange(Province)
        
        provinces_submitting_row <- ft_data_report_status %>%
          summarise(across(all_of(all_months_ordered), ~ sum(. > 0, na.rm = TRUE))) %>%
          mutate(Province = "Provinces Submitting",
                 Months = NA)
        
        final_ft_report_status <- bind_rows(ft_data_report_status, provinces_submitting_row)
        
        ft <- flextable(final_ft_report_status) %>%
          theme_box() %>%
          bold(part = "header") %>%
          align(align = "center", part = "all") %>%
          autofit() %>%
          set_table_properties(width = 1, layout = "autofit") # Added for better fitting
        doc <- body_add_flextable(doc, value = ft)
        doc <- body_add_break(doc)
      }
      
      if ("natPivotTable" %in% input$dl_choices) {
        doc <- body_add_par(doc, "Fish Catch by Quarter", style = "heading 1")
        # Re-generate the table data as it would appear in the DT, but directly for flextable
        ft_data_nat_pivot <- isolate(filtered_nat_data()) %>%
          filter(natfishcatch_typefishing != "3")
        if (nrow(ft_data_nat_pivot) > 0) {
          ft_nat_summary <- ft_data_nat_pivot %>%
            group_by(fish_name_kh, fish_name_en, Quarter) %>%
            summarise(Total = sum(catch, na.rm = TRUE), .groups = "drop") %>%
            pivot_wider(names_from = Quarter, values_from = Total, values_fill = 0)
          for (q in c("Q1", "Q2", "Q3", "Q4")) {
            if (!q %in% names(ft_nat_summary)) {
              ft_nat_summary[[q]] <- 0
            }
          }
          total_catch_nat <- sum(ft_nat_summary[c("Q1", "Q2", "Q3", "Q4")], na.rm = TRUE)
          ft_nat_summary <- ft_nat_summary %>%
            mutate(Total_Catch = Q1 + Q2 + Q3 + Q4) %>%
            mutate(`Contribution (%)` = if(total_catch_nat > 0) round((Total_Catch / total_catch_nat) * 100, 2) else 0) %>%
            arrange(desc(Total_Catch))
          grand_total_row_nat <- tibble(
            fish_name_kh = "Grand Total", fish_name_en = "",
            Q1 = sum(ft_nat_summary$Q1), Q2 = sum(ft_nat_summary$Q2),
            Q3 = sum(ft_nat_summary$Q3), Q4 = sum(ft_nat_summary$Q4),
            Total_Catch = total_catch_nat, `Contribution (%)` = 100.00
          )
          final_ft_nat_table <- bind_rows(ft_nat_summary, grand_total_row_nat) %>%
            rename(`Khmer Name` = fish_name_kh, `English Name` = fish_name_en)
          
          ft <- flextable(final_ft_nat_table) %>%
            theme_box() %>%
            bold(part = "header") %>%
            align(align = "center", part = "all") %>%
            autofit() %>%
            set_table_properties(width = 1, layout = "autofit") %>% # Added for better fitting
            colformat_num(j = c("Q1", "Q2", "Q3", "Q4", "Total_Catch"), digits = 0, big.mark = ",") %>%
            colformat_num(j = "Contribution (%)", digits = 2)
          doc <- body_add_flextable(doc, value = ft)
        } else {
          doc <- body_add_par(doc, "No data available for this table with current filters.", style = "Normal")
        }
        doc <- body_add_break(doc)
      }
      
      
      if ("riceFieldTable" %in% input$dl_choices) {
        doc <- body_add_par(doc, "Fisheries Capture in Rice-Fields by Species", style = "heading 1")
        ft_data_rice_field <- isolate(filtered_nat_data()) %>% filter(natfishcatch_typefishing == "3")
        if (nrow(ft_data_rice_field) > 0) {
          ft_rice_summary <- ft_data_rice_field %>%
            group_by(fish_name_kh, fish_name_en, Quarter) %>%
            summarise(Quarter_Total = sum(catch, na.rm = TRUE), .groups = "drop") %>%
            pivot_wider(names_from = Quarter, values_from = Quarter_Total, values_fill = 0)
          for (q in c("Q1", "Q2", "Q3", "Q4")) {
            if (!q %in% names(ft_rice_summary)) {
              ft_rice_summary[[q]] <- 0
            }
          }
          total_production_rice <- sum(ft_rice_summary[c("Q1", "Q2", "Q3", "Q4")], na.rm = TRUE)
          ft_rice_summary <- ft_rice_summary %>%
            mutate(Total = Q1 + Q2 + Q3 + Q4) %>%
            mutate(`Contribution (%)` = if(total_production_rice > 0) round((Total / total_production_rice) * 100, 2) else 0) %>%
            arrange(desc(Total))
          grand_total_row_rice <- tibble(
            fish_name_kh = "Grand Total", fish_name_en = "",
            Q1 = sum(ft_rice_summary$Q1), Q2 = sum(ft_rice_summary$Q2),
            Q3 = sum(ft_rice_summary$Q3), Q4 = sum(ft_rice_summary$Q4),
            Total = total_production_rice, `Contribution (%)` = 100.00
          )
          final_ft_rice_table <- bind_rows(ft_rice_summary, grand_total_row_rice) %>%
            rename(`Khmer Name` = fish_name_kh, `English Name` = fish_name_en)
          
          ft <- flextable(final_ft_rice_table) %>%
            theme_box() %>%
            bold(part = "header") %>%
            align(align = "center", part = "all") %>%
            autofit() %>%
            set_table_properties(width = 1, layout = "autofit") %>% # Added for better fitting
            colformat_num(j = c("Q1", "Q2", "Q3", "Q4", "Total"), digits = 0, big.mark = ",") %>%
            colformat_num(j = "Contribution (%)", digits = 2)
          doc <- body_add_flextable(doc, value = ft)
        } else {
          doc <- body_add_par(doc, "No data available for this table with current filters.", style = "Normal")
        }
        doc <- body_add_break(doc)
      }
      
      if ("riceFieldProvinceTable" %in% input$dl_choices) {
        doc <- body_add_par(doc, "Freshwater Capture in Rice-Fields by Province", style = "heading 1")
        ft_data_rice_province <- isolate(filtered_nat_data()) %>% filter(natfishcatch_typefishing == "3")
        if (nrow(ft_data_rice_province) > 0) {
          ft_summary_rice_province <- ft_data_rice_province %>%
            group_by(province_kh, province_en) %>%
            summarise(`Production (MT)` = sum(catch, na.rm = TRUE), .groups = "drop")
          total_rice_catch_province <- sum(ft_summary_rice_province$`Production (MT)`, na.rm = TRUE)
          ft_summary_rice_province <- ft_summary_rice_province %>%
            mutate(`Contribution (%)` = if(total_rice_catch_province > 0) round((`Production (MT)` / total_rice_catch_province) * 100, 2) else 0) %>%
            arrange(desc(`Production (MT)`))
          grand_total_row_rice_province <- tibble(
            province_kh = "Grand Total", province_en = "",
            `Production (MT)` = total_rice_catch_province, `Contribution (%)` = 100.00
          )
          final_ft_rice_province_table <- bind_rows(ft_summary_rice_province, grand_total_row_rice_province) %>%
            rename(`Province (Khmer)` = province_kh, `Province (English)` = province_en)
          
          ft <- flextable(final_ft_rice_province_table) %>%
            theme_box() %>%
            bold(part = "header") %>%
            align(align = "center", part = "all") %>%
            autofit() %>%
            set_table_properties(width = 1, layout = "autofit") %>% # Added for better fitting
            colformat_num(j = "Production (MT)", digits = 0, big.mark = ",") %>%
            colformat_num(j = "Contribution (%)", digits = 2)
          doc <- body_add_flextable(doc, value = ft)
        } else {
          doc <- body_add_par(doc, "No data available for this table with current filters.", style = "Normal")
        }
        doc <- body_add_break(doc)
      }
      
      if ("provinceProductionTable" %in% input$dl_choices) {
        doc <- body_add_par(doc, "Inland Production by Province", style = "heading 1")
        ft_data_province_prod <- isolate(filtered_nat_data()) %>% filter(natfishcatch_typefishing != "3")
        if (nrow(ft_data_province_prod) > 0) {
          ft_summary_province_prod <- ft_data_province_prod %>%
            group_by(province_kh, province_en) %>%
            summarise(`Total Catch` = sum(catch, na.rm = TRUE), .groups = "drop")
          total_catch_province_prod <- sum(ft_summary_province_prod$`Total Catch`, na.rm = TRUE)
          ft_summary_province_prod <- ft_summary_province_prod %>%
            mutate(`Contribution (%)` = if(total_catch_province_prod > 0) round((`Total Catch` / total_catch_province_prod) * 100, 2) else 0) %>%
            arrange(desc(`Total Catch`))
          grand_total_row_province_prod <- tibble(
            province_kh = "Grand Total", province_en = "",
            `Total Catch` = total_catch_province_prod, `Contribution (%)` = 100.00
          )
          final_ft_province_prod_table <- bind_rows(ft_summary_province_prod, grand_total_row_province_prod) %>%
            rename(`Province (Khmer)` = province_kh, `Province (English)` = province_en)
          
          ft <- flextable(final_ft_province_prod_table) %>%
            theme_box() %>%
            bold(part = "header") %>%
            align(align = "center", part = "all") %>%
            autofit() %>%
            set_table_properties(width = 1, layout = "autofit") %>% # Added for better fitting
            colformat_num(j = "Total Catch", digits = 0, big.mark = ",") %>%
            colformat_num(j = "Contribution (%)", digits = 2)
          doc <- body_add_flextable(doc, value = ft)
        } else {
          doc <- body_add_par(doc, "No data available for this table with current filters.", style = "Normal")
        }
        doc <- body_add_break(doc)
      }
      
      if ("aquaProvinceProductionTable" %in% input$dl_choices) {
        doc <- body_add_par(doc, "Aquaculture Production by Province", style = "heading 1")
        ft_data_aqua_province <- isolate(filtered_aqu_table_data())
        if (nrow(ft_data_aqua_province) > 0) {
          ft_summary_aqua_province <- ft_data_aqua_province %>%
            group_by(province_kh, province_en) %>%
            summarise(`Total Aquaculture Production (MT)` = sum(catch, na.rm = TRUE), .groups = "drop")
          total_aqu_prod_province <- sum(ft_summary_aqua_province$`Total Aquaculture Production (MT)`, na.rm = TRUE)
          ft_summary_aqua_province <- ft_summary_aqua_province %>%
            mutate(`Contribution (%)` = if(total_aqu_prod_province > 0) round((`Total Aquaculture Production (MT)` / total_aqu_prod_province) * 100, 2) else 0) %>%
            arrange(desc(`Total Aquaculture Production (MT)`))
          grand_total_row_aqua_province <- tibble(
            province_kh = "Grand Total", province_en = "",
            `Total Aquaculture Production (MT)` = total_aqu_prod_province, `Contribution (%)` = 100.00
          )
          final_ft_aqua_province_table <- bind_rows(ft_summary_aqua_province, grand_total_row_aqua_province) %>%
            rename(`Province (Khmer)` = province_kh, `Province (English)` = province_en)
          
          ft <- flextable(final_ft_aqua_province_table) %>%
            theme_box() %>%
            bold(part = "header") %>%
            align(align = "center", part = "all") %>%
            autofit() %>%
            set_table_properties(width = 1, layout = "autofit") %>% # Added for better fitting
            colformat_num(j = "Total Aquaculture Production (MT)", digits = 0, big.mark = ",") %>%
            colformat_num(j = "Contribution (%)", digits = 2)
          doc <- body_add_flextable(doc, value = ft)
        } else {
          doc <- body_add_par(doc, "No data available for this table with current filters.", style = "Normal")
        }
        doc <- body_add_break(doc)
      }
      
      if ("processingTable" %in% input$dl_choices) {
        doc <- body_add_par(doc, "Processed Fish Production by Quarter", style = "heading 1")
        ft_data_processing <- isolate(filtered_processing_data())
        if (nrow(ft_data_processing) > 0) {
          df_converted_ft <- ft_data_processing %>%
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
          normal_df_ft <- df_converted_ft %>% filter(!processing_typeprocessing %in% c("11", "26"))
          special_df_ft <- df_converted_ft %>% filter(processing_typeprocessing %in% c("11", "26"))
          top_15_normal_ft <- tibble()
          total_normal_production_ft <- 0
          if(nrow(normal_df_ft) > 0) {
            normal_summary_ft <- normal_df_ft %>%
              group_by(proc_name_kh, proc_name_en, Quarter) %>%
              summarise(Quarter_Total = sum(processing_amount, na.rm = TRUE), .groups = "drop") %>%
              pivot_wider(names_from = Quarter, values_from = Quarter_Total, values_fill = 0)
            for (q in c("Q1", "Q2", "Q3", "Q4")) {
              if (!q %in% names(normal_summary_ft)) {
                normal_summary_ft[[q]] <- 0
              }
            }
            top_15_normal_ft <- normal_summary_ft %>%
              mutate(Total = Q1 + Q2 + Q3 + Q4) %>%
              arrange(desc(Total)) %>%
              slice_head(n = 15)
            total_normal_production_ft <- sum(top_15_normal_ft$Total, na.rm = TRUE)
            top_15_normal_ft <- top_15_normal_ft %>%
              mutate(`Contribution (%)` = if(total_normal_production_ft > 0) round((Total / total_normal_production_ft) * 100, 2) else 0)
          }
          special_summary_ft <- tibble()
          if (nrow(special_df_ft) > 0) {
            special_summary_ft <- special_df_ft %>%
              group_by(proc_name_kh, proc_name_en, Quarter) %>%
              summarise(Quarter_Total = sum(processing_amount, na.rm = TRUE), .groups = "drop") %>%
              pivot_wider(names_from = Quarter, values_from = Quarter_Total, values_fill = 0)
            for (q in c("Q1", "Q2", "Q3", "Q4")) {
              if (!q %in% names(special_summary_ft)) {
                special_summary_ft[[q]] <- 0
              }
            }
            special_summary_ft <- special_summary_ft %>%
              mutate(Total = Q1 + Q2 + Q3 + Q4, `Contribution (%)` = 0)
          }
          grand_total_row_processing <- tibble(
            proc_name_kh = "Grand Total", proc_name_en = "",
            Q1 = sum(top_15_normal_ft$Q1, na.rm = TRUE),
            Q2 = sum(top_15_normal_ft$Q2, na.rm = TRUE),
            Q3 = sum(top_15_normal_ft$Q3, na.rm = TRUE),
            Q4 = sum(top_15_normal_ft$Q4, na.rm = TRUE),
            Total = total_normal_production_ft,
            `Contribution (%)` = 100.00
          )
          final_ft_processing_table <- bind_rows(top_15_normal_ft, grand_total_row_processing, special_summary_ft) %>%
            rename(`Khmer Name` = proc_name_kh, `English Name` = proc_name_en)
          
          ft <- flextable(final_ft_processing_table) %>%
            theme_box() %>%
            bold(part = "header") %>%
            align(align = "center", part = "all") %>%
            autofit() %>%
            set_table_properties(width = 1, layout = "autofit") %>% # Added for better fitting
            colformat_num(j = c("Q1", "Q2", "Q3", "Q4", "Total"), digits = 2, big.mark = ",") %>%
            colformat_num(j = "Contribution (%)", digits = 2)
          doc <- body_add_flextable(doc, value = ft)
        } else {
          doc <- body_add_par(doc, "No data available for this table with current filters.", style = "Normal")
        }
        doc <- body_add_break(doc)
      }
      
      
      if ("summaryProductionChart" %in% input$dl_choices) {
        doc <- body_add_par(doc, "Freshwater, Marine and Aquaculture Production Chart", style = "heading 1")
        # Ensure plot generation is isolated to capture current state
        plot_obj <- isolate(summaryProductionChart())
        if (!is.null(plot_obj) && !inherits(plot_obj, "ggplot_null")) { # Check if plot is not the "No data" placeholder
          doc <- body_add_gg(doc, value = plot_obj, style = "centered", width = 6, height = 5) # Adjusted width/height
        } else {
          doc <- body_add_par(doc, "No data available to generate this chart with current filters.", style = "Normal")
        }
        doc <- body_add_break(doc)
      }
      
      if ("aquaProductionChart" %in% input$dl_choices) {
        doc <- body_add_par(doc, "Aquaculture and Fingerling Production Chart", style = "heading 1")
        plot_obj <- isolate(aquaProductionChart())
        if (!is.null(plot_obj) && !inherits(plot_obj, "ggplot_null")) {
          doc <- body_add_gg(doc, value = plot_obj, style = "centered", width = 6, height = 5) # Adjusted width/height
        } else {
          doc <- body_add_par(doc, "No data available to generate this chart with current filters.", style = "Normal")
        }
        doc <- body_add_break(doc)
      }
      
      print(doc, target = file)
    }
  )
  
}