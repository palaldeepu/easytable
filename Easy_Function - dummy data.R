# Written by Dr. Deepu Palal, free to use and distribute with citation
library(readxl)
library(stats)
library(writexl)
library(officer)
library(flextable)
library(dplyr)
library(nortest)
library(DescTools)
options(nwarnings = 10000)

#df1 <- read_excel("C:/Users/user/Downloads/Laltopi_weight_classification.xlsx", 
#                  sheet = "Sheet1", guess_max = 100000)

output_location<-c("C:/Users/user/Downloads/Descriptives1.docx")
doc <- read_docx()
formatted_text <- fpar(ftext("Analysis", 
                             prop = fp_text(font.size = 24, 
                                            font.family = "Times New Roman", 
                                            bold = TRUE)))
doc <- body_add_fpar(doc, formatted_text, style = "centered")

doc <- body_add_par(doc, paste0("Date and Time: ",Sys.time()), style = "centered")
doc <- body_add_par(doc, " ", style = "centered")
tableno<-1

## Functions ####
calculate_frequency <- function(freqcolumns) {
  for (colno_colno_cat_freq_columns in freqcolumns)
  {
    results_cat_freq_columns <- data.frame(
      Category = character(0),
      Frequency = integer(0),
      Percentage = numeric(0),
      CI_95 = character(0)
    )
    frequency <- table(df1[[colno_colno_cat_freq_columns]])
    percent <- prop.table(frequency) * 100
    n <- sum(frequency)
    for (subcategory_colno_colno_cat_freq_columns in unique(df1[[colno_colno_cat_freq_columns]][!is.na(df1[[colno_colno_cat_freq_columns]]) & is.character(df1[[colno_colno_cat_freq_columns]])])) {
      category_frequency <- frequency[subcategory_colno_colno_cat_freq_columns]
      ci_95 <- binom.test(category_frequency, n, conf.level = 0.95)$conf.int
      results_cat_freq_columns <- rbind(
        results_cat_freq_columns,
        data.frame(
          Category = subcategory_colno_colno_cat_freq_columns,
          Frequency = category_frequency,
          Percentage = round(percent[subcategory_colno_colno_cat_freq_columns],2),
          CI_95 = paste0(round(ci_95[1]*100,2)," - ",round(ci_95[2]*100,2))
        )
      )
    }
    colnames(results_cat_freq_columns) <- c(names(df1)[[colno_colno_cat_freq_columns]],"Frequency","Percentage","CI_95")
    total_row <- data.frame(
      Category = "Total",
      Frequency = n,
      Percentage = 100,
      CI_95 = ""
    )
    colnames(total_row) <- c(names(df1)[[colno_colno_cat_freq_columns]],"Frequency","Percentage","CI_95")
    results_cat_freq_columns <- rbind(results_cat_freq_columns, total_row)
    
    tablename <- paste0("Table ", tableno, ": ", c(names(df1)[[colno_colno_cat_freq_columns]]))
    doc <- body_add_par(doc, tablename)
    doc <- body_add_par(doc, "")
    std_border = fp_border(color="gray")
    desc_table <- flextable(results_cat_freq_columns) 
    desc_table <- desc_table  %>% autofit() %>% 
      fit_to_width(20, inc = 1L, max_iter = 2, unit = "cm")  %>%
      border_remove() %>% 
      theme_booktabs() %>%
      vline(part = "all", j = 3, border = NULL) %>%
      flextable::align(align = "center", j = c(3:4), part = "all") %>% 
      flextable::align(align = "center", j = c(1:2), part = "header") %>%
      fontsize(i = 1, size = 10, part = "header") %>%   # adjust font size of header
      bold(i = 1, bold = TRUE, part = "header") %>%     # adjust bold face of header
      bold(j = 1, bold = TRUE, part = "body")  %>% 
      merge_v(j = 1)%>%
      merge_v(j = 2)%>%
      hline(part = "body", border = std_border)
    doc <- body_add_flextable(doc, value = desc_table)
    doc <- body_add_par(doc, "")
    tableno<<-tableno+1
  }
  rm(list= c("category_frequency","ci_95","colno_colno_cat_freq_columns","desc_table",                              
             "freqcolumns","frequency","percent","n", "results_cat_freq_columns","std_border","subcategory_colno_colno_cat_freq_columns","tablename",                               
             "total_row","typeoffacility"))
}
calculatedescriptives <- function (df_desc, colno, x, y) {
  frequency1<-nrow(df_desc[!is.na(df_desc[[colno]]),])
  print (length(unique(df_desc[[colno]][!is.na(df_desc[[colno]])])))
  if (length(unique(df_desc[[colno]][!is.na(df_desc[[colno]])])) <= 1) {
    normality1<-"N/A" }
  else {
    if (frequency1>3 && frequency1<5000){
      # normality <- shapiro.test(df_desc[[colno]])
      normality <- try(shapiro.test(df_desc[[colno]]), silent = TRUE)
      if (inherits(normality, "htest")) {
        normality1<-normality$p.value
        if (normality1>=0.05) {
          normality1<-""
        } else {normality1<-"*"}} else {
          normality1<-"**"
          
        }
      
    } else {
      if (frequency1>=5000){
        normality<-ad.test(df_desc[[colno]])
        normality1<-normality$p.value
        if (normality1>=0.05) {
          normality1<-""
        } else {normality1<-"*AD"}
      } else {
        normality1<-"**"}
    }
  }
  mean1<-mean(df_desc[[colno]], na.rm=TRUE)
  mean1<- round(mean1,2)
  sd1<-sd(df_desc[[colno]],na.rm=TRUE)
  sd1<- round(sd1,2)
  median1<-median(df_desc[[colno]],na.rm=TRUE)
  median1<- round(median1,2)
  model <- lm(df_desc[[colno]] ~ 1, data=df_desc)
  CI1<-confint(model, level=0.95)
  CI95 <- paste(round(CI1[[1]],2),"-",round(CI1[[2]],2))
  percentile25<-quantile(df_desc[[colno]],na.rm=TRUE,probs=0.25)
  percentile25<- round(percentile25,2)
  percentile75<-quantile(df_desc[[colno]],na.rm=TRUE,probs=0.75)
  percentile75<- round(percentile75,2)
  IQR <- paste(percentile25,"-",percentile75)
  min1<-min(df_desc[[colno]],na.rm=TRUE)
  min1<- round(min1,2)
  max1<-max(df_desc[[colno]],na.rm=TRUE)
  max1<- round(max1,2)
  colname <- names(df1)[[colno]]
  
  bootmed <- apply(matrix(sample(na.omit(df_desc[[colno]]), rep=TRUE, 10^4*length(na.omit(df_desc[[colno]]))), nrow=10^4), 1, median)
  Median_CI1 <- quantile(bootmed, c(0.025, 0.975))
  Median_CI95 <- paste(round(Median_CI1[[1]],2),"-",round(Median_CI1[[2]],2))
  descriptives<- list(colname,x,y,frequency1,mean1,sd1,median1,IQR,CI95,Median_CI95,min1,max1,normality1)
  listtemp <- c(descriptives)
  
}
descriptive_exec <- function(df1, desc_col_no, desc_category, desc_subcategory) {
  
  list1<-list()
  if (is.null(desc_col_no)) { } else {
    for (colno in desc_col_no)
    {
      col<-df1[[colno]]
      if (is.numeric(col)) {
        if (is.null(desc_category)) {
          df_desc<-df1
          x<-"NIL"
          y<-"Total"
          if (!all(is.na(df_desc[[colno]])))  
          {
            listtemp<- calculatedescriptives(df_desc, colno, x, y)
            list1<-append(list1, list(listtemp))
          }else {next}} else {
            desc_categoryname <- names(df1)[[desc_category]]
            for (x in unique(df1[[desc_category]]))
            {
              if (is.null(desc_subcategory)) {
                df_desc<-subset(df1, df1[[desc_category]] == x)
                y<-"NIL"
                if (!all(is.na(df_desc[[colno]])))  
                {
                  listtemp<- calculatedescriptives(df_desc,colno,x,y)
                  list1<-append(list1, list(listtemp))
                } else {next}} else {
                  desc_subcategoryname <- names(df1)[[desc_subcategory]]
                  for (y in unique(df1[[desc_subcategory]]))
                  {
                    df_desc<-subset(df1, df1[[desc_category]] == x & df1[[desc_subcategory]] == y)
                    if (!all(is.na(df_desc[[colno]])))  
                    {
                      listtemp<- calculatedescriptives(df_desc, colno, x, y)
                      list1<-append(list1, list(listtemp)) } else {next} }
                  
                  df_desc<-subset(df1, df1[[desc_category]] == x)
                  y<-"Total"
                  if (!all(is.na(df_desc[[colno]])))  
                  {
                    listtemp<- calculatedescriptives(df_desc, colno, x, y)
                    list1<-append(list1, list(listtemp))
                  } else {next}
                }
            }
            
            df_desc<-df1
            x<-"Total"
            listtemp<- calculatedescriptives(df_desc, colno,x,y)
            list1<-append(list1, list(listtemp))   
            
          }
        df_desc_temp <- data.frame(matrix(unlist(list1), ncol=13, byrow=T),stringsAsFactors=FALSE)
      }
      else {
        df_desc_temp <- data.frame(matrix(nrow = 0, ncol=13))
      }
      colnames(df_desc_temp) <- c("Variable","Category","Sub Category","Frequency","Mean", "SD", "Median", "IQR", "95% CI Mean", "95% CI Median", "Min","Max","Normality")
      
    }
    df_desc_merged <- cbind(df_desc_temp[, 1:2])
    df_desc_merged$Category <- paste0(df_desc_temp$Category, df_desc_temp$Normality)
    df_desc_merged$`Sub Category` <- paste0(df_desc_temp$`Sub Category`)
    if (all(df_desc_merged$`Sub Category` == "NIL")) {
      df_desc_merged <- df_desc_merged[, !names(df_desc_merged) %in% "Sub Category"]
    } else {
      df_desc_merged$`Sub Category` <- paste0(df_desc_temp$`Sub Category`, df_desc_temp$Normality)
    }
    df_desc_merged$N <- df_desc_temp$Frequency
    df_desc_merged$`Mean (SD)` <- paste0(df_desc_temp$Mean, " (", df_desc_temp$SD, ")")
    df_desc_merged$`Median (IQR)` <- paste0(df_desc_temp$Median, " (", df_desc_temp$IQR, ")")
    df_desc_merged$`95% CI Mean` <- df_desc_temp$`95% CI Mean`
    df_desc_merged$`95% CI Median` <- df_desc_temp$`95% CI Median`
    df_desc_merged$`Min-Max` <- paste0(df_desc_temp$Min, " - ", df_desc_temp$Max)
    n_col <- ncol(df_desc_merged)
    
    
    tabledesc <- paste0("Table ", tableno, ": Quantitative Data")
    #doc <- body_add_par(doc, tabledesc, style = "heading 2")
    std_border = fp_border(color="gray")
    desc_table <- flextable(df_desc_merged) 
    desc_table <- desc_table  %>% autofit() %>% 
      fit_to_width(20, inc = 1L, max_iter = 2, unit = "cm")  %>%
      border_remove() %>% 
      theme_booktabs() %>%
      vline(part = "all", j = (n_col-6), border = NULL) %>%
      flextable::align(align = "center", j = c(2:n_col), part = "all") %>% 
      flextable::align(align = "center", j = c(1:(n_col-6)), part = "header") %>%
      fontsize(i = 1, size = 10, part = "header") %>%   # adjust font size of header
      bold(i = 1, bold = TRUE, part = "header") %>%     # adjust bold face of header
      bold(j = 1, bold = TRUE, part = "body")  %>% 
      merge_v(j = 1)%>%
      merge_v(j = 2)%>%
      hline(j = ~`Variable`,part = "body") %>%
      hline(part = "body", border = std_border) %>% 
      set_caption(tabledesc) %>% 
      add_footer_row(values = c(" ","* Non-Gaussian distribution"),
                     colwidths = c(1,(n_col-1)), top = TRUE)
    doc <- body_add_flextable(doc, value = desc_table)
    tableno<-tableno+1
    
    rm(list= c("col","colno","desc_category","desc_categoryname","desc_col_no","desc_subcategory",
               "desc_subcategoryname","desc_table","df_desc","df_desc_temp","list1","listtemp",
               "std_border","tabledesc","x","y","calculateddescriptives"),envir = .GlobalEnv)
    
  }
}
association <- function(dependent, independent, tablepercent){
  if (is.null(dependent)) { } else {
    for (independentvar in independent) {
      independent_name <- names(df1)[[independentvar]]
      #doc <- body_add_par(doc, independent_name, style = "heading 1")
      for (dependentvar in dependent)
      {
        dependent_name <- names(df1)[[dependentvar]]
        tabledesc <- paste0("Table ", tableno, ": Association between ", independent_name, " and ", dependent_name)
        doc <- body_add_par(doc, "", style = "heading 2")
        
        cont_table <- table(df1[[independentvar]], df1[[dependentvar]])
        expected_counts <- chisq.test(cont_table)$expected
        prop_low_counts <- sum(expected_counts < 5) / length(expected_counts)
        extranote<-c("")
        if (prop_low_counts > 0.2 || any(chisq.test(cont_table)$expected < 1)) {
          fisher_test <- tryCatch(fisher.test(cont_table), error = function(e) NULL)
          if (is.null(fisher_test)) {
            fisher_test <- fisher.test(cont_table, simulate.p.value = TRUE, B=1e4)
            extranote <- c("Using Monte Carlo simulation")
          }
          pvalue<-round(fisher_test["p.value"][[1]],3)
          finalp<- paste0("Fisher exact P=",replace(pvalue,pvalue==0,"<0.001")," ",extranote)
        } else {
          chisq_test <- chisq.test(cont_table)  
          chisq<-chisq_test["statistic"][[1]][[1]]
          dfreedom<-chisq_test["parameter"][[1]][[1]]
          pvalue<-round(chisq_test["p.value"][[1]],3)
          finalp<- paste0("chisq(",dfreedom,")=",round(chisq,3),", P=",replace(pvalue,pvalue==0,"<0.001"))
        }
        
        cont_table_old <- cbind(cont_table, Total = rowSums(cont_table))
        cont_table_old <- rbind(cont_table_old, Total = colSums(cont_table_old))
        
        if (tablepercent=="row") {
          prop_table <- proportions(cont_table, margin = 1) * 100 # row percentages
        } else {
          if (tablepercent=="column") {
            prop_table <- proportions(cont_table, margin = 2) * 100 # row percentages
          }
        }
        
        col_pct <- rowSums(cont_table) / sum(cont_table) * 100  
        row_pct <- colSums(cont_table) / sum(cont_table) * 100
        prop_table <- rbind(prop_table, Total = row_pct)  
        prop_table <- cbind(prop_table, Total = col_pct)  
        prop_table[nrow(prop_table), ncol(prop_table)] <- 100
        
        
        table_print <- matrix(paste0(cont_table_old, " (", round(prop_table,2),")"), nrow = nrow(prop_table))
        rownames(table_print) <- rownames(cont_table_old)
        colnames(table_print) <- colnames(cont_table_old)
        
        
        table_print <- as.data.frame.matrix(table_print)
        table_print$Categories <- rownames(table_print)
        table_print <- table_print[, c(ncol(table_print), 1:(ncol(table_print)-1))]
        
        ##Printing Chisquare table using Flextable####
        nobrackets <- rep("N (%)", times = ncol(table_print)-1 )
        nobracketsfinal <- c("",nobrackets)
        
        chi_table <- flextable(table_print)
        chi_table <- chi_table  %>% autofit() %>% 
          fit_to_width(20, inc = 1L, max_iter = 20, unit = "cm")  %>% 
          add_header_row(top = FALSE,
                         values = nobracketsfinal) %>%
          border_remove() %>% 
          theme_booktabs() %>%
          vline(part = "all", j = 1, border = NULL) %>%   # at column 2 
          vline(part = "all", j = ncol(table_print)-1, border = NULL)  %>%
          hline(i = nrow(table_print)-1, border = NULL, part = "body")  %>%
          add_footer_row(values = c(finalp),
                         colwidths = ncol(table_print), top = TRUE) %>%
          fontsize(i = 1, size = 10, part = "footer") %>%
          hline(i = 1, part = "header", border = fp_border(color="transparent"))  %>%
          set_caption(tabledesc)
        chi_table
        
        doc <- body_add_flextable(doc, value = chi_table)
        #doc <- body_add_par(doc, finalp, style = "Normal")
        rm(list= c("cont_table", "expected_counts", "prop_low_counts", "finalp", "chisq_test", "chisq", "dfreedom", "pvalue", "cont_table_old", "col_pct", "row_pct", "prop_table", "table_print", "tabledesc", "chi_table","nobrackets", "nobracketsfinal"))
        #rm(list= c("cont_table", "expected_counts", "prop_low_counts", "fisher_test", "pvalue", "finalp", "chisq_test", "chisq", "dfreedom", "pvalue", "cont_table_old", "prop_table", "col_pct", "row_pct", "prop_table", "table_print", "tabledesc", "chi_table","nobrackets", "nobracketsfinal"))
        tableno<-tableno+1
      }
    }
  }
}
test_two_groups <- function(df1, desc_col_no, desc_category) {
  
  test_two_groups_result <- data.frame(
    Parameter = character(),
    Group1 = character(),
    Group2 = character(),
    Significance = character(),
    stringsAsFactors = FALSE
  )  
  
  
  df1[[desc_category]] <- as.factor(df1[[desc_category]])
  if (is.null(desc_col_no)) { } else {
    for (colno in desc_col_no)
    {
      col<-df1[[colno]]
      if (is.numeric(col)) {
        group_0 <- df1[[colno]][df1[[desc_category]] == levels(df1[[desc_category]])[1]]
        group_1 <- df1[[colno]][df1[[desc_category]] == levels(df1[[desc_category]])[2]]
        shapiro_test_group_0 <- shapiro.test(group_0)
        shapiro_test_group_1 <- shapiro.test(group_1)
        levene_test_check <- leveneTest(df1[[colno]] ~ df1[[desc_category]])
        if (shapiro_test_group_0$p.value >= 0.05 && shapiro_test_group_1$p.value >= 0.05) {
          
          if (levene_test_check$`Pr(>F)`[1] >= 0.05) {
            t_test_result <- t.test(df1[[colno]] ~ df1[[desc_category]], 
                                    var.equal = TRUE,  # Set to TRUE if variances are equal
                                    na.rm = TRUE)
            #Calculate effect size
            mean_0 <- mean(group_0, na.rm = TRUE)
            mean_1 <- mean(group_1, na.rm = TRUE)
            sd_0 <- sd(group_0, na.rm = TRUE)
            sd_1 <- sd(group_1, na.rm = TRUE)
            n_0 <- length(na.omit(group_0))
            n_1 <- length(na.omit(group_1))
            pooled_sd <- sqrt(((n_0 - 1) * sd_0^2 + (n_1 - 1) * sd_1^2) / (n_0 + n_1 - 2))
            cohen_d <- (mean_0 - mean_1) / pooled_sd
            
            Significance <- paste0("Independent Samples t test, t(",round(t_test_result$parameter[[1]],2),")=",round(t_test_result$statistic[[1]],2),", P=",round(t_test_result$p.value,3), ", Cohen's d=",round(cohen_d,2))
          } else {
            t_test_result <- t.test(df1[[colno]] ~ df1[[desc_category]], 
                                    var.equal = FALSE,  # Set to TRUE if variances are equal
                                    na.rm = TRUE)
            #Calculate effect size
            mean_0 <- mean(group_0, na.rm = TRUE)
            mean_1 <- mean(group_1, na.rm = TRUE)
            var_0 <- var(group_0, na.rm = TRUE)
            var_1 <- var(group_1, na.rm = TRUE)
            cohen_d <- (mean_0 - mean_1) / sqrt((var_0 + var_1) / 2)
            
            Significance <- paste0("Welch's t-test, t(",round(t_test_result$parameter[[1]],2),")=",round(t_test_result$statistic[[1]],2)," (Cont. Corrected), P=",round(t_test_result$p.value,3), ", Cohen's d=",round(cohen_d,2))
          }
        } else {
          
          mann_whitney_result <- wilcox.test(df1[[colno]] ~ df1[[desc_category]], 
                                             na.rm = TRUE)
          z_value <- qnorm(mann_whitney_result$p.value / 2) * sign(mann_whitney_result$statistic - (length(group_0) * length(group_1)) / 2)
          n_total <- length(na.omit(group_0)) + length(na.omit(group_1))
          effect_size_r <- z_value / sqrt(n_total)
          
          U <- mann_whitney_result$statistic
          n_0 <- length(na.omit(group_0))
          n_1 <- length(na.omit(group_1))
          rank_biserial_corr <- 1 - (2 * U) / (n_0 * n_1)
          
          Significance <- paste0("Mann Whitney Test",", Z=", round(z_value,2),", P=",round(mann_whitney_result$p.value,3),", Rank-biserial correlation=",round(rank_biserial_corr,2))
        }
      } else {
        Significance <- paste0("Non Numeric Data")
      }
      test_two_groups_temp <- data.frame(
        Parameter = colnames(df1)[colno],
        Group1 = levels(df1[[desc_category]])[1],
        Group2 = levels(df1[[desc_category]])[2],
        Significance = Significance,
        stringsAsFactors = FALSE
      )
      test_two_groups_result <- rbind(test_two_groups_result, test_two_groups_temp)
      
    }}
  #return(test_two_groups_result)
  
  n_col <- ncol(test_two_groups_result)
  tabledesc <- paste0("Table ", tableno, ": Statistical Analysis")
  std_border = fp_border(color="gray")
  desc_table <- flextable(test_two_groups_result) 
  desc_table <- desc_table  %>% autofit() %>% 
    fit_to_width(20, inc = 1L, max_iter = 2, unit = "cm")  %>%
    border_remove() %>% 
    theme_booktabs() %>%
    vline(part = "all", j = (n_col-3), border = NULL) %>%
    flextable::align(align = "center", j = c(2:n_col), part = "all") %>% 
    flextable::align(align = "center", j = c(1:(n_col-3)), part = "header") %>%
    fontsize(i = 1, size = 10, part = "header") %>%   # adjust font size of header
    bold(i = 1, bold = TRUE, part = "header") %>%     # adjust bold face of header
    bold(j = 1, bold = TRUE, part = "body")  %>% 
    merge_v(j = 2)%>%
    merge_v(j = 3)%>%
    # hline(j = 1,part = "body") %>%
    hline(part = "body", border = std_border) %>% 
    set_caption(tabledesc)
  doc <- body_add_flextable(doc, value = desc_table)
  tableno<-tableno+1
}

##Functions Commands####

#calculate_frequency (freqcolumns<- c())
#association (dependent<-c(), independent<-c(), tablepercent<-c("row"))
#descriptive_exec (df1, desc_col_no<-c(6), desc_category<-c(3), desc_subcategory<-c())

#Compute_variable <- "Hb"

#bootmed <- apply(matrix(sample(na.omit(df1[[Compute_variable]]), rep=TRUE, 10^4*length(na.omit(df1[[Compute_variable]]))), nrow=10^4), 1, median)
#quantile(bootmed, c(0.025, 0.975), na.rm = TRUE)

#sort(df1[[Compute_variable]])[qbinom(c(0.025,0.975), length(df1[[Compute_variable]]), 0.5)] #Less Accurate

#MedianCI(df1[[Compute_variable]], conf.level = 0.95, na.rm = TRUE, method = "exact",R = 10000) # Library Desctools
#quantile(df1[[Compute_variable]], probs = c(0.25, 0.75), na.rm = TRUE)


## Print Table ####
#print(doc, target = output_location)
