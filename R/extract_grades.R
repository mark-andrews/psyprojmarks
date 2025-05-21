# R/marksheet.R -------------------------------------------------------------
#' @title Extract grades and comments from a Word mark-sheet
#'
#' @description
#' Read a single NTU Psychology project mark-sheet (`.docx`) and return the
#' grades plus the free-text comments for each criterion (C1 – C6 and
#' *Overall*).
#'
#' @param path Path to the Word document marksheet.
#'
#' @return A tibble with columns
#' \describe{
#'   \item{`ID`}{Criterion ID (`C1` – `C6`, `Overall`).}
#'   \item{`label`}{Short criterion label, e.g. *Methods*, *Analysis*.}
#'   \item{`grade`}{The grade picked from the drop-down (e.g. `"1HIGH"`).}
#'   \item{`comment`}{Free-text comment typed by the marker (or `"No comment"`).}
#' }
#'
#' @examples
#' \dontrun{
#' grades <- extract_grades("UGPROJ25_343_1.docx")
#' print(grades)
#' }
#' @export
extract_grades <- function(path) {
  text <- extract_text(path)

  ## --- criterion grades -------------------------------------------------
  # C1, C2 ...C6 grades
  c_pattern <- "Grade\\s+for\\s+(?<description>.*?)\\s*:\\s*(?<grade>[0-9]+[A-Z]+)\\s*$"

  criteria_grades <- Filter(length, stringr::str_match_all(text, pattern = c_pattern))

  criteria_grades_df <-
    purrr::map(criteria_grades, as.data.frame) |>
    dplyr::bind_rows() |>
    dplyr::select(-1) |>
    tibble::as_tibble()

  # Overall grade -----------------------------------------------------------
  overall_pattern <- "Choose overall grade:\\s*(?<grade>[0-9]+[A-Z]+)\\s*$"
  overall_grade <- Filter(length, stringr::str_match_all(text, pattern = overall_pattern))

  overall_grade_df <- overall_grade |>
    as.data.frame() |>
    dplyr::transmute(description = "Overall", grade)

  dplyr::bind_rows(criteria_grades_df, overall_grade_df) |>
    dplyr::left_join(criteria_df, by = "description") |>
    dplyr::left_join(extract_comments(path), by = "criterion") |>
    dplyr::mutate(comment = tidyr::replace_na(comment, "No comment")) |>
    dplyr::select(criterion, grade, comment)
}

#' @title Join two markers' grade tables and flag agreement
#'
#' @param path_a Path to the first marker's Word document.
#' @param path_b Path to the second marker's Word document.
#'
#' @return A tibble with the grades/comments from both markers and a logical
#'   `agree` column showing whether the two grades match.
#'
#' @examples
#' \dontrun{
#' joined <- join_grades("marker_A.docx", "marker_B.docx")
#' joined
#' }
#' @export
join_grades <- function(path_a, path_b) {
  marker_a <- extract_marker(path_a)
  marker_b <- extract_marker(path_b)

  grades_df_A <- extract_grades(path_a)
  grades_df_B <- extract_grades(path_b)

  dplyr::full_join(grades_df_A, grades_df_B,
    by = "criterion",
    suffix = stringr::str_c("_", c(marker_a, marker_b))
  ) |>
    dplyr::mutate(agree = Reduce(
      `==`,
      dplyr::pick(dplyr::starts_with("grade_"))
    )) |>
    dplyr::relocate(dplyr::starts_with("grade"), agree, .after = criterion)
}



#' @title Write a joined-grades table to Excel
#'
#' @description
#' Saves the output of [\code{join_grades()}] to an `.xlsx` workbook with:
#'
#' * green ✔ ticks for agreements and red ✘ crosses for disagreements;
#' * wrapped, top-aligned comment cells;
#' * auto-fitted column widths plus a generous width for the comment columns.
#'
#' @param joined_grades_df A data frame returned by [\code{join_grades()}].
#' @param file Name of the Excel file to create (default `"results.xlsx"`).
#'
#' @return (Invisibly) the path of the written file.
#' @export
write_joined_grades <- function(joined_grades_df,
                                file = "results.xlsx") {
  tick <- "\u2714" # ✔
  cross <- "\u2718" # ✘

  df <- dplyr::mutate(
    joined_grades_df,
    agree = dplyr::if_else(agree, tick, cross)
  )

  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "Results")
  openxlsx::writeData(wb, "Results", df)

  ## --- coloured ticks / crosses ----------------------------------------
  style_tick <- openxlsx::createStyle(fontColour = "#00B050")
  style_cross <- openxlsx::createStyle(fontColour = "#C00000")

  agree_col <- match("agree", names(df))
  tick_rows <- which(df$agree == tick) + 1L # +1 header
  cross_rows <- which(df$agree == cross) + 1L

  openxlsx::addStyle(
    wb, "Results", style_tick,
    rows = tick_rows, cols = agree_col,
    gridExpand = TRUE, stack = TRUE
  )
  openxlsx::addStyle(
    wb, "Results", style_cross,
    rows = cross_rows, cols = agree_col,
    gridExpand = TRUE, stack = TRUE
  )

  ## --- sensible column widths ------------------------------------------
  pad <- 2
  targetWidth <- pad + purrr::map_dbl(names(df), stringr::str_length)
  openxlsx::setColWidths(
    wb, "Results",
    cols = seq_along(df),
    widths = targetWidth
  )

  ## --- wrap the comment_* columns --------------------------------------
  comment_cols <- which(stringr::str_detect(names(df), "^comment_"))

  wrap_style <- openxlsx::createStyle(
    wrapText = TRUE, valign = "top", halign = "left"
  )

  openxlsx::addStyle(
    wb, "Results", wrap_style,
    rows = 2:(nrow(df) + 1), cols = comment_cols,
    gridExpand = TRUE, stack = TRUE
  )

  openxlsx::setColWidths(
    wb, "Results",
    cols = comment_cols, widths = 50
  )

  openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
  invisible(normalizePath(file))
}

# -------------------------------------------------------------------------
# -------------------  internal helper functions  -------------------------

criteria_df <- readr::read_csv(
  "criterion,description,label
C1,Understanding  and evaluating existing research and theory,Understanding literature
C2,Formulating a research problem,Formulating research
C3,Designing / selecting and applying appropriate methods for obtaining data,Methods
C4,Designing and applying appropriate methods for data analysis,Analysis
C5,Understanding findings and drawing conclusions,Discussion
C6,Communicating,Communicating
Overall,Overall,Overall",
  show_col_types = FALSE
)

extract_text <- function(path) {
  doc <- officer::read_docx(path)
  officer::docx_summary(doc)$text
}


extract_comments <- function(path) {
  text <- extract_text(path)

  comment_pattern <- "^(?<criterion>C[1-6]):\\s*.*Evidence your answers, using as much space as you require:(?<comment>.*)1st2.12.23rdMarginal failFail"

  purrr::map(
    Filter(length, stringr::str_match_all(text, comment_pattern)),
    as.data.frame
  ) |>
    dplyr::bind_rows() |>
    dplyr::select(-1) |>
    tibble::as_tibble() |>
    dplyr::left_join(criteria_df, by = "criterion") |>
    dplyr::select(criterion, comment)
}

extract_marker <- function(path) {
  text <- extract_text(path)
  marker_pattern <- "^Marker:\\s*(?<marker>[A-Za-z]+)$"
  Filter(length, stringr::str_match_all(text, marker_pattern))[[1]][[1, "marker"]]
}
