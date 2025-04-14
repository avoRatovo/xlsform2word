utils::globalVariables(c(
  "type", "name", "group_order", "parent_group", "original_type"
))

#' Convert XLSForm to Word Document (French)
#'
#' This function reads a structured XLSForm Excel file and generates a formatted Word document in French.
#' @param path Path to XLSForm Excel file (.xls or .xlsx)
#' @param output Name of the output Word file
#' @return Returns Success if successfully generated
#' @export
# xlsToword Function - Complete and Robust Version
#' @importFrom dplyr filter mutate group_by summarise first arrange
#' @importFrom readxl read_excel excel_sheets
#' @importFrom stringr str_detect str_trim str_split str_match str_split_fixed
#' @importFrom officer read_docx body_add_par body_add_fpar fp_text ftext fpar fp_par
#' @importFrom stats na.omit
#' @importFrom utils install.packages
#' @importFrom rlang sym
xlsform2word_fr <- function(path, output = "xlsform_output.docx") {
  if (!requireNamespace("readxl", quietly = TRUE)) stop("Please install the 'readxl' package.")
  if (!requireNamespace("officer", quietly = TRUE)) stop("Please install the 'officer' package.")
  if (!requireNamespace("dplyr", quietly = TRUE)) stop("Please install the 'dplyr' package.")
  if (!requireNamespace("stringr", quietly = TRUE)) install.packages("stringr")


  `%>%` <- dplyr::`%>%`

  if (!file.exists(path)) stop("Le fichier specifie est introuvable")
  if (!tolower(tools::file_ext(path)) %in% c("xls", "xlsx")) stop("Le fichier doit etre au format Excel (xls ou xlsx)")

  sheets <- tryCatch({
    excel_sheets(path)
  }, error = function(e) stop("Impossible de lire les feuilles du fichier Excel", e$message))

  if (!"survey"%in% sheets) stop("La feuille 'survey' est manquante")
  survey <- tryCatch({
    read_excel(path, sheet = "survey")
  }, error = function(e) stop("Erreur lors de la lecture de la feuille 'survey'", e$message))

  form_title <- "Survey Form"
  if ("settings"%in% sheets) {
    settings <- tryCatch(read_excel(path, sheet = "settings"), error = function(e) {
      warning("Erreur lors de la lecture de la feuille 'settings'. Titre par defaut utilise")
      return(NULL)
    })
    if (!is.null(settings) && "form_title"%in% names(settings)) {
      titre_temp <- as.character(settings$form_title[1])
      if (!is.na(titre_temp) && str_trim(titre_temp) != "") {
        form_title <- titre_temp
        message("Titre d\u00e9tect\u00e9 ", form_title)
      } else {
        message("'form_title' est vide ou manquantTitre par d\u00e9faut utilise")
      }
    } else {
      message("Colonne 'form_title' non trouvee dans la feuille 'settings'Titre par defaut utilise.")
    }
  } else {
    message("Feuille 'settings' absente Titre par defaut utilise")
  }

  choices <- if ("choices"%in% sheets) {
    tryCatch({
      read_excel(path, sheet = "choices") %>%
        dplyr::mutate(list_name = tolower(str_trim(list_name)))
    }, error = function(e) {
      warning("Erreur lors de la lecture de la feuille 'choices'")
      data.frame()
    })
  } else data.frame()

  required_columns <- c("type", "name")
  if (!all(required_columns %in% names(survey))) {
    stop("Les colonnes essentielles 'type' et 'name' sont manquantes dans la feuille 'survey'")
  }

  get_label_column <- function(df) {
    colnames_lower <- tolower(names(df))
    french_label <- names(df)[stringr::str_detect(colnames_lower, "^label::french")]
    english_label <- names(df)[stringr::str_detect(colnames_lower, "^label::english")]
    generic_label <- if ("label"%in% colnames_lower) names(df)[colnames_lower == "label"] else character(0)

    label_col <- c(french_label, english_label, generic_label)
    if (length(label_col) > 0) {
      return(label_col[1])
    } else {
      stop("Aucune colonne de label trouvee dans la feuille 'survey'")
    }
  }

  label_col <- get_label_column(survey)
  message("Colonne utilise pour les libeles", label_col)

  # Filtrer uniquement les vraies questions (hors begin/end group/repeat)
  question_rows <- survey %>%
    dplyr::filter(!stringr::str_detect(tolower(type), "^begin[ _](group|repeat)$|^end[ _](group|repeat)$")) %>%
    dplyr::filter(!is.na(name))

  # V\u00e9rifier les doublons uniquement parmi les noms de vraies questions
  if (any(duplicated(question_rows$name))) {
    duplicated_names <- question_rows$name[duplicated(question_rows$name)]
    warning(paste0(
      "Il y a des noms de variables dupliques (hors groupes/r\u00e9p\u00e9titions) dans la feuille 'survey'\n",
      " Noms de variables dupliques ", paste(unique(duplicated_names), collapse = ", ")
    ))
  }

  survey <- survey %>%
    dplyr::filter(!is.na(type) & str_trim(type) != "") %>%
    dplyr::mutate(group_id = NA_character_, group_order = NA_integer_, parent_group = NA_character_, original_type = tolower(type))

  group_name <- NULL
  group_index <- 0
  group_stack <- list()
  total_rows <- nrow(survey)

  for (i in seq_len(total_rows)) {
    type_val <- tolower(survey$type[i])
    if (!is.na(type_val)) {
      if (stringr::str_detect(type_val, "^begin[ _]group$|^begin[ _]repeat$")) type_val <- "begin group"
      if (stringr::str_detect(type_val, "^end[ _]group$|^end[ _]repeat$")) type_val <- "end group"
    }
    survey$type[i] <- type_val

    if (!is.na(type_val) && type_val == "begin group") {
      group_index <- group_index + 1
      group_name <- ifelse(!is.na(survey[[label_col]][i]), survey[[label_col]][i], paste("Group", group_index))
      group_stack[[length(group_stack) + 1]] <- group_name
    }

    if (!is.null(group_name)) {
      survey$group_id[i] <- group_name
      survey$group_order[i] <- group_index
      if (length(group_stack) > 1) {
        survey$parent_group[i] <- group_stack[[length(group_stack) - 1]]
      }
    }

    if (!is.na(type_val) && type_val == "end group") {
      group_stack <- group_stack[-length(group_stack)]
      group_name <- if (length(group_stack) > 0) group_stack[[length(group_stack)]] else NULL
    }

    if (i %% 10 == 0 || i == total_rows) {
      cat(sprintf("Progression lecture : %.1f%%\n", 100 * i / total_rows))
    }
  }

  get_condition_fpar <- function(relevant_expr, element_type = "question") {
    if (is.na(relevant_expr)) return(NULL)
    pattern <- "\\$\\{(.+?)\\} ?= ?['\"](.+?)['\"]"
    match <- str_match(relevant_expr, pattern)
    if (nrow(match) == 0 || is.na(match[1, 2])) return(NULL)
    varname <- match[1, 2]
    value <- match[1, 3]
    var_row <- survey %>% dplyr::filter(name == varname)
    var_label <- if (nrow(var_row) > 0) var_row[[label_col]][1] else varname
    list_type <- tolower(str_trim(str_split(var_row$type[1], " ")[[1]][2]))
    value_label <- value
    if (!is.na(list_type) && list_type != ""&& "list_name"%in% names(choices)) {
      value_row <- choices %>% dplyr::filter(list_name == list_type, name == value)
      if (nrow(value_row) > 0) {
        if ("label::French"%in% names(value_row)) {
          value_label <- value_row[["label::French"]][1]
        } else if ("label"%in% names(value_row)) {
          value_label <- value_row[["label"]][1]
        }
      }
    }
    prefix <- if (element_type == "section") "Cette section s\u2019affiche si la question "else "Cette question s\u2019affiche si la question "
    return(fpar(
      ftext("Condition : ", fp_text(bold = TRUE)),
      ftext(paste0(prefix, var_label, "= ", value_label), fp_text(italic = TRUE))
    ))
  }

  doc <- officer::read_docx()
  title_text <- paste0("Questionnaire : ", form_title)
  big_blue_style <- fp_text(font.size = 24, bold = TRUE, color = "blue", font.family = "Calibri")
  doc <- body_add_fpar(doc, fpar(ftext(title_text, prop = big_blue_style), fp_p = fp_par(text.align = "center")))
  doc <- body_add_par(doc, value = "", style = "Normal")

  group_levels <- survey %>%
    dplyr::filter(!is.na(group_id)) %>%
    group_by(group_id) %>%
    summarise(order = min(group_order), parent = first(parent_group), original_type = original_type[1], .groups = "drop") %>%
    arrange(order)

  group_blocks <- lapply(group_levels$group_id, function(gid) {
    survey %>% dplyr::filter(group_id == gid)
  })

  prefix_map <- list()
  counter_map <- list()

  get_prefix <- function(group_id) {
    parent <- group_levels$parent[group_levels$group_id == group_id]
    parent <- parent[1]  # pour \u00e9viter les vecteurs longs
    if (is.na(parent)) {
      if (is.null(counter_map[["root"]])) counter_map[["root"]] <<- 1 else counter_map[["root"]] <<- counter_map[["root"]] + 1
      prefix_map[[group_id]] <<- as.character(counter_map[["root"]])
    } else {
      parent_prefix <- prefix_map[[parent]]
      if (is.null(counter_map[[parent_prefix]])) counter_map[[parent_prefix]] <<- 1 else counter_map[[parent_prefix]] <<- counter_map[[parent_prefix]] + 1
      prefix_map[[group_id]] <<- paste0(parent_prefix, ".", counter_map[[parent_prefix]])
    }
    return(prefix_map[[group_id]])
  }

  section_progress <- 1
  total_sections <- length(group_blocks)

  for (group in group_blocks) {
    group_id <- unique(na.omit(group$group_id))[1]
    group_label <- group_id
    group_prefix <- get_prefix(group_id)
    is_repeat <- any(stringr::str_detect(group$original_type, "begin[ _]repeat"))
    section_title <- paste0(if (is_repeat) "Section r\u00e9p\u00e9titive "else if (!is.na(group_levels$parent[group_levels$group_id == group_id])) "Sous-section "else "Section ", group_prefix, ": ", group_label)

    section_style <- fp_text(font.size = 16, bold = TRUE, color = "blue", underlined = TRUE, font.family = "Calibri")
    doc <- body_add_par(doc, value = "", style = "Normal")
    doc <- body_add_fpar(doc, fpar(ftext(section_title, prop = section_style)))

    if ("relevant"%in% names(group)) {
      relevant_text <- unique(na.omit(group$relevant))[1]
      if (!is.null(relevant_text)) {
        condition_fpar <- get_condition_fpar(relevant_text, element_type = "section")
        if (!is.null(condition_fpar)) {
          doc <- body_add_fpar(doc, condition_fpar)
          doc <- body_add_par(doc, "", style = "Normal")
        }
      }
    }

    questions <- group %>%
      dplyr::filter(!stringr::str_detect(type, "begin group|end group")) %>%
      dplyr::filter(!is.na(!!sym(label_col)))

    question_number <- 1
    for (i in seq_len(nrow(questions))) {
      question_type <- str_trim(questions$type[i])
      is_note <- stringr::str_detect(question_type, "note")
      is_calculate <- stringr::str_detect(question_type, "calculate")
      question_type_parts <- str_split_fixed(question_type, " ", 2)
      q_type_base <- tolower(question_type_parts[1])
      list_name <- if (ncol(question_type_parts) > 1) tolower(str_trim(question_type_parts[2])) else NA
      label <- questions[[label_col]][i]

      if (is_note) {
        doc <- body_add_fpar(doc, fpar(ftext(paste0("Note : ", label), prop = fp_text(italic = TRUE))))
        doc <- body_add_par(doc, "", style = "Normal")
      } else {
        full_q_number <- paste0(group_prefix, ".", question_number)
        doc <- body_add_par(doc, value = paste0(full_q_number, ". ", label), style = "Normal")
      }

      if (!is_note && stringr::str_detect(q_type_base, "text|integer|decimal")) {
        doc <- body_add_fpar(doc, fpar(ftext("R\u00e9ponse : [ins\u00e9rer votre r\u00e9ponse ici]", prop = fp_text(italic = TRUE))))
        doc <- body_add_par(doc, "", style = "Normal")
      }

      if (!is_note && q_type_base %in% c("select_one", "select_multiple")) {
        opt_icon <- if (q_type_base == "select_one") "( )"else "[ ]"
        msg <- if (q_type_base == "select_one") "Veuillez s\u00e9lectionner une seule option :"else "Veuillez s\u00e9lectionner une ou plusieurs options :"
        doc <- body_add_fpar(doc, fpar(ftext(msg, prop = fp_text(italic = TRUE))))

        if (!is.na(list_name)) {
          opts <- choices %>% dplyr::filter(list_name == !!list_name)

          # INSERTION ICI
          colnames_lower <- tolower(names(opts))
          label_choice_col <- names(opts)[stringr::str_detect(colnames_lower, "^label::french")]
          if (length(label_choice_col) == 0) label_choice_col <- names(opts)[stringr::str_detect(colnames_lower, "^label::english")]
          if (length(label_choice_col) == 0 && "label"%in% colnames_lower) label_choice_col <- names(opts)[colnames_lower == "label"]
          if (length(label_choice_col) > 0) {
            label_choice_col <- label_choice_col[1]
          } else {
            label_choice_col <- NULL
          }

          if (!is.null(label_choice_col) && nrow(opts) > 0) {
            for (opt in opts[[label_choice_col]]) {
              if (!is.na(opt)) {
                doc <- body_add_par(doc, value = paste0("    ", opt_icon, "", opt), style = "Normal")
              }
            }
          }
        }

        doc <- body_add_par(doc, "", style = "Normal")
      }

      if (!is_note && is_calculate) {
        doc <- body_add_fpar(doc, fpar(ftext("D\u00e9ja calcul\u00e9", prop = fp_text(italic = TRUE, color = "red"))))
      }

      if ("relevant"%in% names(questions) && !is.na(questions$relevant[i])) {
        condition_fpar <- get_condition_fpar(questions$relevant[i], element_type = "question")
        if (!is.null(condition_fpar)) {
          doc <- body_add_fpar(doc, condition_fpar)
          doc <- body_add_par(doc, "", style = "Normal")
        }
      }

      if (!is_note) {
        question_number <- question_number + 1
      }
    }

    cat(sprintf("Progression document : %.1f%% - %s\n", 100 * section_progress / total_sections, section_title))
    section_progress <- section_progress + 1
  }

  print(doc, target = output)
  cat("G\u00e9n\u00e9ration Word termin\u00e9e : ", output, "\n")
  return("Success")
}

#' Convert XLSForm to Word Document (English)
#'
#' This function reads a structured XLSForm Excel file and generates a formatted Word document in English.
#' @param path Path to XLSForm Excel file (.xls or .xlsx)
#' @param output Name of the output Word file
#' @return Returns Success if successfully generated
#' @export
# xlsToword Function - Complete and Robust Version
#' @importFrom dplyr filter mutate group_by summarise first arrange
#' @importFrom readxl read_excel excel_sheets
#' @importFrom stringr str_detect str_trim str_split str_match str_split_fixed
#' @importFrom officer read_docx body_add_par body_add_fpar fp_text ftext fpar fp_par
#' @importFrom stats na.omit
#' @importFrom utils install.packages
#' @importFrom rlang sym
xlsform2word_en <- function(path, output = "xlsform_output.docx") {
  if (!requireNamespace("readxl", quietly = TRUE)) stop("Please install the 'readxl' package.")
  if (!requireNamespace("officer", quietly = TRUE)) stop("Please install the 'officer' package.")
  if (!requireNamespace("dplyr", quietly = TRUE)) stop("Please install the 'dplyr' package.")
  if (!requireNamespace("stringr", quietly = TRUE)) install.packages("stringr")

  `%>%` <- dplyr::`%>%`

  if (!file.exists(path)) stop("The specified file was not found.")
  if (!tolower(tools::file_ext(path)) %in% c("xls", "xlsx")) stop("The file must be in Excel format (.xls or .xlsx).")

  sheets <- tryCatch({
    excel_sheets(path)
  }, error = function(e) stop("Unable to read the Excel sheet names: ", e$message))

  if (!"survey"%in% sheets) stop("The 'survey' sheet is missing.")
  survey <- tryCatch({
    read_excel(path, sheet = "survey")
  }, error = function(e) stop("Error reading the 'survey' sheet: ", e$message))

  form_title <- "Survey Form"
  if ("settings"%in% sheets) {
    settings <- tryCatch(read_excel(path, sheet = "settings"), error = function(e) {
      warning("Error reading the 'settings' sheet. Default title will be used.")
      return(NULL)
    })
    if (!is.null(settings) && "form_title"%in% names(settings)) {
      titre_temp <- as.character(settings$form_title[1])
      if (!is.na(titre_temp) && str_trim(titre_temp) != "") {
        form_title <- titre_temp
        message("Title detected: ", form_title)
      } else {
        message("'form_title' is empty or missing. Default title will be used.")
      }
    } else {
      message("'form_title' column not found in the 'settings' sheet. Default title will be used.")
    }
  } else {
    message("'settings' sheet is missing. Default title will be used.")
  }

  choices <- if ("choices"%in% sheets) {
    tryCatch({
      read_excel(path, sheet = "choices") %>%
        dplyr::mutate(list_name = tolower(str_trim(list_name)))
    }, error = function(e) {
      warning("Error reading the 'choices' sheet.")
      data.frame()
    })
  } else data.frame()

  required_columns <- c("type", "name")
  if (!all(required_columns %in% names(survey))) {
    stop("Essential columns 'type' and 'name' are missing in the 'survey' sheet.")
  }

  get_label_column <- function(df) {
    colnames_lower <- tolower(names(df))
    french_label <- names(df)[str_detect(colnames_lower, "^label::french")]
    english_label <- names(df)[str_detect(colnames_lower, "^label::english")]
    generic_label <- if ("label"%in% colnames_lower) names(df)[colnames_lower == "label"] else character(0)

    label_col <- c( english_label, generic_label,french_label)
    if (length(label_col) > 0) {
      return(label_col[1])
    } else {
      stop("No label column found in the 'survey' sheet.")
    }
  }

  label_col <- get_label_column(survey)
  message("Label column used: ", label_col)

  survey <- survey %>%
    dplyr::filter(!is.na(type) & str_trim(type) != "") %>%
    dplyr::mutate(group_id = NA_character_, group_order = NA_integer_, parent_group = NA_character_, original_type = tolower(type))

  group_name <- NULL
  group_index <- 0
  group_stack <- list()
  total_rows <- nrow(survey)
  cat("reprocessing the form...\n")

  for (i in seq_len(total_rows)) {
    type_val <- tolower(survey$type[i])
    if (!is.na(type_val)) {
      if (stringr::str_detect(type_val, "^begin[ _]group$|^begin[ _]repeat$")) type_val <- "begin group"
      if (stringr::str_detect(type_val, "^end[ _]group$|^end[ _]repeat$")) type_val <- "end group"
    }
    survey$type[i] <- type_val

    if (!is.na(type_val) && type_val == "begin group") {
      group_index <- group_index + 1
      group_name <- ifelse(!is.na(survey[[label_col]][i]), survey[[label_col]][i], paste("Group", group_index))
      group_stack[[length(group_stack) + 1]] <- group_name
    }

    if (!is.null(group_name)) {
      survey$group_id[i] <- group_name
      survey$group_order[i] <- group_index
      if (length(group_stack) > 1) {
        survey$parent_group[i] <- group_stack[[length(group_stack) - 1]]
      }
    }

    if (!is.na(type_val) && type_val == "end group") {
      group_stack <- group_stack[-length(group_stack)]
      group_name <- if (length(group_stack) > 0) group_stack[[length(group_stack)]] else NULL
    }

    if (i %% 10 == 0 || i == total_rows) {
      cat(sprintf("Reading progress: %.1f%%\n", 100 * i / total_rows))
    }
  }

  get_condition_fpar <- function(relevant_expr, element_type = "question") {
    if (is.na(relevant_expr)) return(NULL)
    pattern <- "\\$\\{(.+?)\\} ?= ?['\"](.+?)['\"]"
    match <- str_match(relevant_expr, pattern)
    if (nrow(match) == 0 || is.na(match[1, 2])) return(NULL)
    varname <- match[1, 2]
    value <- match[1, 3]
    var_row <- survey %>% dplyr::filter(name == varname)
    var_label <- if (nrow(var_row) > 0) var_row[[label_col]][1] else varname
    list_type <- tolower(str_trim(str_split(var_row$type[1], " ")[[1]][2]))
    value_label <- value
    if (!is.na(list_type) && list_type != ""&& "list_name"%in% names(choices)) {
      value_row <- choices %>% dplyr::filter(list_name == list_type, name == value)
      if (nrow(value_row) > 0) {
        if ("label::English"%in% names(value_row)) {
          value_label <- value_row[["label::English"]][1]
        } else if ("label"%in% names(value_row)) {
          value_label <- value_row[["label"]][1]
        }
      }
    }
    prefix <- if (element_type == "section") "This section is shown if question "else "This question is shown if question "
    return(fpar(
      ftext("Condition: ", fp_text(bold = TRUE)),
      ftext(paste0(prefix, var_label, "= ", value_label), fp_text(italic = TRUE))
    ))
  }

  doc <- officer::read_docx()
  title_text <- paste0("Questionnaire: ", form_title)
  big_blue_style <- fp_text(font.size = 24, bold = TRUE, color = "blue", font.family = "Calibri")
  doc <- body_add_fpar(doc, fpar(ftext(title_text, prop = big_blue_style), fp_p = fp_par(text.align = "center")))
  doc <- body_add_par(doc, value = "", style = "Normal")

  group_levels <- survey %>%
    dplyr::filter(!is.na(group_id)) %>%
    group_by(group_id) %>%
    summarise(order = min(group_order), parent = first(parent_group), original_type = original_type[1], .groups = "drop") %>%
    arrange(order)

  group_blocks <- lapply(group_levels$group_id, function(gid) {
    survey %>% dplyr::filter(group_id == gid)
  })

  prefix_map <- list()
  counter_map <- list()

  get_prefix <- function(group_id) {
    parent <- group_levels$parent[group_levels$group_id == group_id]
    parent <- parent[1]  # to avoid long vectors
    if (is.na(parent)) {
      if (is.null(counter_map[["root"]])) counter_map[["root"]] <<- 1 else counter_map[["root"]] <<- counter_map[["root"]] + 1
      prefix_map[[group_id]] <<- as.character(counter_map[["root"]])
    } else {
      parent_prefix <- prefix_map[[parent]]
      if (is.null(counter_map[[parent_prefix]])) counter_map[[parent_prefix]] <<- 1 else counter_map[[parent_prefix]] <<- counter_map[[parent_prefix]] + 1
      prefix_map[[group_id]] <<- paste0(parent_prefix, ".", counter_map[[parent_prefix]])
    }
    return(prefix_map[[group_id]])
  }

  section_progress <- 1
  total_sections <- length(group_blocks)

  for (group in group_blocks) {
    group_id <- unique(na.omit(group$group_id))[1]
    group_label <- group_id
    group_prefix <- get_prefix(group_id)
    is_repeat <- any(str_detect(group$original_type, "begin[ _]repeat"))
    section_title <- paste0(
      if (is_repeat) "Repeating Section "else if (!is.na(group_levels$parent[group_levels$group_id == group_id])) "Subsection "else "Section ",
      group_prefix, ": ", group_label
    )

    section_style <- fp_text(font.size = 16, bold = TRUE, color = "blue", underlined = TRUE, font.family = "Calibri")
    doc <- body_add_par(doc, value = "", style = "Normal")
    doc <- body_add_fpar(doc, fpar(ftext(section_title, prop = section_style)))

    if ("relevant"%in% names(group)) {
      relevant_text <- unique(na.omit(group$relevant))[1]
      if (!is.null(relevant_text)) {
        condition_fpar <- get_condition_fpar(relevant_text, element_type = "section")
        if (!is.null(condition_fpar)) {
          doc <- body_add_fpar(doc, condition_fpar)
          doc <- body_add_par(doc, "", style = "Normal")
        }
      }
    }

    questions <- group %>%
      dplyr::filter(!stringr::str_detect(type, "begin group|end group")) %>%
      dplyr::filter(!is.na(!!sym(label_col)))

    question_number <- 1
    for (i in seq_len(nrow(questions))) {
      question_type <- str_trim(questions$type[i])
      is_note <- stringr::str_detect(question_type, "note")
      is_calculate <- stringr::str_detect(question_type, "calculate")
      question_type_parts <- str_split_fixed(question_type, " ", 2)
      q_type_base <- tolower(question_type_parts[1])
      list_name <- if (ncol(question_type_parts) > 1) tolower(str_trim(question_type_parts[2])) else NA
      label <- questions[[label_col]][i]

      if (is_note) {
        doc <- body_add_fpar(doc, fpar(ftext(paste0("Note: ", label), prop = fp_text(italic = TRUE))))
        doc <- body_add_par(doc, "", style = "Normal")
      } else {
        full_q_number <- paste0(group_prefix, ".", question_number)
        doc <- body_add_par(doc, value = paste0(full_q_number, ". ", label), style = "Normal")
      }

      if (!is_note && stringr::str_detect(q_type_base, "text|integer|decimal")) {
        doc <- body_add_fpar(doc, fpar(ftext("Answer: [insert your answer here]", prop = fp_text(italic = TRUE))))
        doc <- body_add_par(doc, "", style = "Normal")
      }

      if (!is_note && q_type_base %in% c("select_one", "select_multiple")) {
        opt_icon <- if (q_type_base == "select_one") "( )"else "[ ]"
        msg <- if (q_type_base == "select_one") "Please select one option:"else "Please select one or more options:"
        doc <- body_add_fpar(doc, fpar(ftext(msg, prop = fp_text(italic = TRUE))))

        if (!is.na(list_name)) {
          opts <- choices %>% dplyr::filter(list_name == !!list_name)

          colnames_lower <- tolower(names(opts))
          label_choice_col <- names(opts)[stringr::str_detect(colnames_lower, "^label::english")]
          if (length(label_choice_col) == 0) label_choice_col <- names(opts)[stringr::str_detect(colnames_lower, "^label::english")]
          if (length(label_choice_col) == 0 && "label"%in% colnames_lower) label_choice_col <- names(opts)[colnames_lower == "label"]
          if (length(label_choice_col) > 0) {
            label_choice_col <- label_choice_col[1]
          } else {
            label_choice_col <- NULL
          }

          if (!is.null(label_choice_col) && nrow(opts) > 0) {
            for (opt in opts[[label_choice_col]]) {
              if (!is.na(opt)) {
                doc <- body_add_par(doc, value = paste0("    ", opt_icon, "", opt), style = "Normal")
              }
            }
          }
        }

        doc <- body_add_par(doc, "", style = "Normal")
      }

      if (!is_note && is_calculate) {
        doc <- body_add_fpar(doc, fpar(ftext("Already calculated", prop = fp_text(italic = TRUE, color = "red"))))
      }

      if ("relevant"%in% names(questions) && !is.na(questions$relevant[i])) {
        condition_fpar <- get_condition_fpar(questions$relevant[i], element_type = "question")
        if (!is.null(condition_fpar)) {
          doc <- body_add_fpar(doc, condition_fpar)
          doc <- body_add_par(doc, "", style = "Normal")
        }
      }

      if (!is_note) {
        question_number <- question_number + 1
      }
    }

    cat(sprintf("Document progress: %.1f%% - %s\n", 100 * section_progress / total_sections, section_title))
    section_progress <- section_progress + 1
  }

  print(doc, target = output)
  cat("Word document generation completed: ", output, "\n")
  return("Success")
}

# Remarques :
# - Les deux fonctions sont maintenant pretes pour etre export\u00e9es dans un package R.
# - Tu peux ajouter un fichier README et initier les tests si tu veux le soumettre sur GitHub ou CRAN plus tard.
