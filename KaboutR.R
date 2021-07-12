# functies inladen

# packages inladen
pkg <- c("tidyverse","expss","openxlsx","readxl")
new.pkg <- pkg[!(pkg %in% installed.packages())]
if (length(new.pkg)) {install.packages(new.pkg, repos = "http://cran.rstudio.com")}
invisible(capture.output(lapply(pkg, require, character.only = TRUE)))


#Stijl voor tabellen
style <- list()
style$fontSize <- 10L
style$total <- createStyle(fgFill = "#FFFFFF", border ="TopBottom", borderStyle = "thin", borderColour = "#77a2b5", halign = "right",fontSize=style$fontSize,fontColour="black",textDecoration="bold")
style$header <- createStyle(fgFill = "#77a2b5", halign = "left", wrapText =FALSE,fontColour="white",textDecoration="bold",fontSize=style$fontSize)

# style$main <- createStyle(halign = "right", numFmt = "0", border ="TopBottom", borderStyle = "thin", borderColour = "#77a2b5",fontSize=style$fontSize)
style$main <- createStyle(halign = "right", numFmt = "0.0%", border ="TopBottom", borderStyle = "thin", borderColour = "#77a2b5",fontSize=style$fontSize)

style$row <- createStyle(halign = "left",border ="TopBottom", borderStyle = "thin", borderColour = "#77a2b5",fontColour="black",textDecoration="bold",fontSize=style$fontSize)
style$oberonBlauw <- "#4f8197"
style$oberonGroen <- "#bdcc01"
style$labelbreedte <- 60
style$titelbreedte <- 60


OpenSurveydata_openxlsx <- function(fname=NULL){
  
  #Data inlezen
  if (is.null(fname)){fname <- file.choose()}
  df_xl <<- openxlsx::read.xlsx(fname, sheet = 1) %>%  select(-"_will_be_deleted_")
  codeplan <<- openxlsx::read.xlsx(fname, sheet = 2)
  names(codeplan) <- str_replace(names(codeplan),"[.]"," ")
  
  
  # data labelen
  for (i in 1:ncol(df_xl)){
    if (is.na(codeplan$Type[i]) || codeplan$Type[i] == 'MultipleChoice' ){
      var_lab(df_xl[i]) <<- paste0(codeplan$Question[i],"|",codeplan$`Choice./.Column`[i])
      
    } else {
      var_lab(df_xl[i]) <<- codeplan$Question[i]
      if (codeplan$Type[i]=="Matrix" & is.na(codeplan$Question[i])){
        var_lab(df_xl[i]) <<- codeplan$Question[i-1]
        codeplan$Question[i] <<- codeplan$Question[i-1]
      }  #toegevoegd  als matrix en Question="" dan overnemen van de vorige
    }
    
    x <- strsplit(codeplan$Values[i], split="[{:}]") %>% unlist %>% str_remove_all("\n") %>% 
      str_replace_all('#N/A', '999999') #tijdelijk. Kijken wat de echte waarde is.
    x <- x[-c(1,which(x==', '))]
    
    lab <- "\n"
    
    if (length(x)>0 & codeplan$Data.Type[i]=="Number") {
      
      for (n in 1:(length(x)/2) ){lab <- append(lab,paste(x[(n*2)-1],x[n*2], "\n"))}  
    }
    val_lab(df_xl[i]) <<- num_lab(lab)
    
  }
  
 
}


OpenSurveydata <- function(fname=NULL, filterexpressie=NULL){
  
  #Data inlezen
  if (is.null(fname)){fname <- file.choose()} #file dialog 
  
  
  
  
  
  
  df_xl <<- readxl::read_xlsx(fname, sheet = 1) %>%  select(-"_will_be_deleted_")
  codeplan <<- readxl::read_xlsx(fname, sheet = 2)
  
  if(missing(filterexpressie)){
    # doe niks
  } else {
    # lst <- as.list(match.call())
    df_xl <<- df_xl %>% filter(UQ(rlang::parse_quosure(filterexpressie)))
    
  }
  
  # data labelen
  for (i in 1:ncol(df_xl)){
    if (is.na(codeplan$Type[i]) || codeplan$Type[i] == 'MultipleChoice' ){
      var_lab(df_xl[i]) <<- paste0(codeplan$Question[i],"|",codeplan$`Choice / Column`[i])
      
    } else {
      var_lab(df_xl[i]) <<- codeplan$Question[i]
      if (codeplan$Type[i]=="Matrix" & is.na(codeplan$Question[i])){
        var_lab(df_xl[i]) <<- codeplan$Question[i-1]
        codeplan$Question[i] <<- codeplan$Question[i-1]
      }  #toegevoegd  als matrix en Question="" dan overnemen van de vorige
    }
    
    x <- strsplit(codeplan$Values[i], split="[{:}]") %>% unlist %>% str_remove_all("\n") %>% 
      str_replace_all('#N/A', '999999') #tijdelijk. Kijken wat de echte waarde is.
    x <- x[-c(1,which(x==', '))]
    
    lab <- "\n"
    
    if (length(x)>0 & codeplan$`Data Type`[i]=="Number") {
      
      for (n in 1:(length(x)/2) ){lab <- append(lab,paste(x[(n*2)-1],x[n*2], "\n"))}  
    }
    val_lab(df_xl[i]) <<- num_lab(lab)
    
  }
  
  
}


mrtabel <- function(df){
    data.frame(x = colSums(df, na.rm = T)/nrow(df)*100)%>%
    mutate(row_labels=sapply(df, var_lab)) %>% 
    # add_row(row_labels="|#N=" , x=nrow(df)) %>% #totaal N
    mutate('#%'= x,row_labels=str_replace(row_labels," / ","|")) %>% 
    select(-1) %>% 
    tibble::remove_rownames() %>% 
    arrange(desc(`#%`)) %>% 
    as.etable() 
  #mogelijkheid: strip alles wat gelijk is in het label en verplaats dat naar de titel
}


mrfiguur <- function(df){
  data.frame(y = colSums(df, na.rm = T)/nrow(df)) %>% 
    # data.frame(y = colSums(df, na.rm = T)/nrow(df)*100) %>%
    mutate(x=sapply(df, var_lab),x=str_remove_all(x, ".*[|]|, namelijk")) %>% 
    mutate(x=fct_reorder(str_wrap(x,style$labelbreedte), y)) %>%  #str_wrap verdeelt lange labels over meerdere regels
    ggplot(aes(x=x, y=y)) + 
    # geom_bar(stat="identity", width=.5, fill=style$oberonBlauw) +
    geom_col(fill=style$oberonBlauw, width=.5) +
    geom_text(aes(x = x, y = y, label = paste0(round(y,1), '%'), hjust = 1.2) , color="white")+
    coord_flip() +
    ylab("%") +  xlab("") +
    # labs(title= str_wrap(gsub("[|].*", "",var_lab(df[1])),style$titelbreedte) ) + #titel aan of uit
    theme_bw()
  }

matrixtabel <- function (cell_vars, col_vars = total(), row_vars = NULL, weight = NULL,subgroup = NULL, total_label = NULL, total_statistic = "u_cases", total_row_position = c("below", "above", "none")) {
  str_cell_vars = expss:::expr_to_character(substitute(cell_vars))
  str_col_vars = expss:::expr_to_character(substitute(col_vars))
  str_row_vars = expss:::expr_to_character(substitute(row_vars))
  cell_vars = expss:::test_for_null_and_make_list(cell_vars, str_cell_vars)
  col_vars = expss:::test_for_null_and_make_list(col_vars, str_col_vars)
  if (!is.null(row_vars)) 
    row_vars = expss:::test_for_null_and_make_list(row_vars, str_row_vars)
  expss:::multi_cro(cell_vars = col_vars, col_vars = cell_vars, row_vars = row_vars, 
                    weight = weight, subgroup = subgroup, total_label = total_label, 
                    total_statistic = total_statistic, total_row_position = "none", 
                    stat_type = "rpct")
}


# # bewerkt
# cro_cpct <- function (cell_vars, col_vars = total(), row_vars = NULL, weight = NULL, 
#                       subgroup = NULL, total_label = NULL, total_statistic = "u_cases", 
#                       total_row_position = c("below", "above", "none")) {
#   str_cell_vars = expss:::expr_to_character(substitute(cell_vars))
#   str_col_vars = expss:::expr_to_character(substitute(col_vars))
#   str_row_vars = expss:::expr_to_character(substitute(row_vars))
#   cell_vars = expss:::test_for_null_and_make_list(cell_vars, str_cell_vars)
#   col_vars = expss:::test_for_null_and_make_list(col_vars, str_col_vars)
#   if (!is.null(row_vars)) 
#     row_vars = expss:::test_for_null_and_make_list(row_vars, str_row_vars)
#   expss:::multi_cro(cell_vars = cell_vars, col_vars = col_vars, row_vars = row_vars, 
#             weight = weight, subgroup = subgroup, total_label = total_label, 
#             total_statistic = total_statistic, total_row_position = total_row_position, 
#             stat_type = "cpct")
# }




matrixfiguur <- function(df){
lapply(df, matrixtabel) %>% bind_rows() %>%  
  mutate(row_labels=item) %>% as.data.frame()  %>% 
  pivot_longer(!row_labels,  names_to = "kleur", values_to = "x") %>% 
  mutate(kleur=str_remove_all(kleur, ".*[|]")) %>% 
  ggplot(aes(x=x,y=row_labels, fill=kleur)) +
  geom_col()
  }

# str_wrap(x,30) %>%  as.factor()



  


rm("pkg","new.pkg")

KaboutR <- function(){
  
  algedaan <- ""
  list_of_tables <- list()
  # figuren <- list()
  rij <- 1
  
  wb = createWorkbook()
  figurensheet <- "Figuren"
  tabellensheet <- "Tabellen"
  addWorksheet(wb, tabellensheet , gridLines = FALSE)
  addWorksheet(wb, figurensheet , gridLines = FALSE)
  
  for (i in 1:nrow(codeplan)){
    
    if (is.na(codeplan$Type[i])) {
      # geen echte vraag, dus niks doen
    } else if (codeplan$Type[i]=='MultipleChoice' & codeplan$`Data Type`[i]=='Number') { 
      #MR_vraag-----
      q <- codeplan$`Question Variable`[i]
      if (q %in% algedaan ){
        # print("skip")
      } else {
        # Dataset samenstellen
        df <- df_xl %>% select(starts_with(paste0(q,"_"))  & !ends_with("_text"))
        n=nrow(df %>% drop_na())
        text <- paste0(q," - " , gsub("[|].*", "",var_lab(df[1]))," (N=",n ,") (Meerdere antwoorden mogelijk)")
        # figuur maken en wegschrijven
        hoogte=(length(df)/2.66666+2) %>% round()
        xl_write(text, wb, figurensheet, row=rij, col= 1) # Schijf naam naar excel
        
        fname <- paste0(tempdir(),"/", q,".jpg")
        ggsave(fname, mrfiguur(df), width = 16, dpi = 300, height = hoogte, units = "cm")
        insertImage(wb, figurensheet, file = fname, width = 16, height = hoogte, units ="cm" , startRow = rij+1, startCol = 2)
        rij=rij+hoogte*2+2
        
        
        list_of_tables[[q]] <- mrtabel(df) %>% #rename(!!quo_name(text) := 1) %>%  #variabele naam in de kop
          rename(' ' = 1) %>%
          mutate(across(2:last_col(), ~ (.x)/100)) %>%  
          set_caption(text)
        algedaan <- append(algedaan,q)
        print(q)
      }
    } else if ( (codeplan$Type[i] %in% c('Slider','SingleChoice','OpenQuestion')) | (codeplan$Type[i]=='MultipleChoice' & codeplan$`Data Type`[i]=='String') | (codeplan$Type[i]=='Matrix' & codeplan$`Data Type`[i]=='String')) {
      # Single, slider of open vraag----- 
      # q <- codeplan$`Question Variable`[i]
      q <- codeplan$Variable[i]
      df <- df_xl %>% select(q)
      n=nrow(df %>% drop_na())
      text <- paste0(q," - " , gsub("[|].*", "",codeplan$Question[i])," (N=", n,")")
      # qn <- paste0(q, " (N=", nrow(df %>% drop_na()),")")#correctie op aantal rijen: NA er uit 
      list_of_tables[[q]] <- cro_cpct(df, total_row_position ="none") %>% #rename(!!quo_name(text) := 1, '#%'=2) %>%   #variabele naam in de kop
        rename(' ' = 1, '#%'=2) %>%
        mutate(across(2:last_col(), ~ (.x)/100)) %>%  
        set_caption(text)
      print(q)
    } else if (codeplan$Type[i]=='Matrix' & codeplan$`Data Type`[i]=='Number') { 
      #matrix-----
      q <- codeplan$`Question Variable`[i]
      if (q %in% algedaan ){
        # print("skip")
      } else {
        
        item <- codeplan %>% filter(`Question Variable`==q & `Data Type`=="Number") %>% pull(Row) #Toegevoegd: & `Data Type`=="Number" 
        df <- df_xl %>% select(starts_with(paste0(q,"_"))  & !ends_with("_text"))
        n=nrow(df %>% drop_na())
        # qn <- paste0(q, " (N=", nrow(df),")")
        text <- paste0(q," - " , gsub("[|].*", "",codeplan$Question[i])," (N=", n,")")
        # text <- paste0(q," - " , gsub("[|].*", "",var_lab(df[1]))," (N=", nrow(df),")")
        
        
        hoogte=(length(df)/2.66666+2) %>% round()
        xl_write(text, wb, figurensheet, row=rij, col= 1) # Schijf naam naar excel
        fname <- paste0(tempdir(),"/", q,".jpg")
        ggsave(fname, mrfiguur(df), width = 16, dpi = 300, height = hoogte, units = "cm")
        insertImage(wb, figurensheet, file = fname, width = 16, height = hoogte, units ="cm" , startRow = rij+1, startCol = 2)
        rij=rij+hoogte*2+2
        
        
        list_of_tables[[q]] <- lapply(df, matrixtabel) %>% bind_rows() %>%  
          mutate(row_labels=item) %>% 
          # rename(!!quo_name(text) := 1)  %>% #variabele naam in de kop
          rename(' ' = 1)  %>% #variabele naam in de kop
          mutate(across(2:last_col(), ~ (.x)/100)) %>%  
          set_caption(text)
        algedaan <- append(algedaan,q)
        print(q)
      }
    }
    
    
  }
  
  list_of_tables <<- list_of_tables
  
  print("Output exporteren. Even geduld....")
  xl_write(c(list_of_tables), wb, tabellensheet, 
           borders = list(borderColour = "#77a2b5", borderStyle = "thin"),
           header_format = style$header,
           row_labels_format = style$row,
           total_format = style$total,
           total_row_labels_format = style$total,
           main_format = style$main,
           row_symbols_to_remove = c("#"),
           col_symbols_to_remove = "#",
           
  ) 
  
  saveWorkbook(wb, paste0("Output_",format(Sys.time(), "%y%m%d_%H%M"),".xlsx"), overwrite = F)
  
}
