## ============================================================
## BRSR Section C scraper (clean, readable output)
##
## What was wrong with the earlier output, and what this fixes:
##  1. Every row was tagged "P2": the "PRINCIPLE n:" headers were missed
##     because the PDF text carries a margin tab ("P1", "P2"...) at the
##     start of lines. Margin tabs are now blanked out and the principle
##     is tracked properly.
##  2. Question numbers ("1.", "2.") and page numbers were read as
##     data values. Questions are now detected as questions, page
##     footers are dropped, and ONLY real table cells become values.
##  3. Whole raw lines were dumped as the "Indicator". Now each row has
##     the question text (Indicator) + the table row label (Item) +
##     the column / year it belongs to (Period, Column) + a clean number.
##  4. Narrative answers (Yes/No + explanation) were lost or mixed with
##     numbers. They now go to their own "Qualitative" sheet.
##
## OUTPUT (Excel, 5 sheets):
##   Read Me          - what each column means
##   Section C Data   - one row per number: Principle | Q No | Indicator |
##                      Essential/Leadership | Item | Period | Column |
##                      Value | Unit | As Reported
##   Wide Tables      - big multi-column tables (e.g. benefits coverage)
##                      kept as one row per table row, values in order
##   Qualitative      - the written answer for every indicator
##   Raw Text         - cleaned Section C text, for checking
##
## USAGE: put the PDF next to this script, set pdf_path, then
##        source("brsr_sectionC_clean_scraper.R")
## ============================================================

pdf_path       <- "C:/Users/HomePC/Documents/tiss/sem-5/ESG reporting/BRSR reports/Bajaj Finserv.pdf"
out_path       <- "Bajaj_Finserv_BRSR_SectionC_clean.xlsx"
company_footer <- "BAJAJ FINSERV LIMITED"   # running page header to ignore

pkgs <- c("pdftools", "openxlsx")
need <- pkgs[!pkgs %in% installed.packages()[, "Package"]]
if (length(need) > 0) install.packages(need)
library(pdftools)
library(openxlsx)

## ------------------------------------------------------------
## 1. Read the PDF keeping the horizontal layout (column positions
##    matter for assigning numbers to the right column)
## ------------------------------------------------------------
get_lines <- function(path) {
  stopifnot(file.exists(path))
  if (nchar(Sys.which("pdftotext")) > 0) {
    out <- system2("pdftotext", c("-layout", shQuote(path), "-"),
                   stdout = TRUE, stderr = FALSE)
  } else {
    out <- unlist(strsplit(pdftools::pdf_text(path), "\n", fixed = TRUE))
  }
  iconv(out, from = "UTF-8", to = "UTF-8", sub = " ")
}

# blank out the "P1".."P9" margin tab but keep character positions
clean_line <- function(l) {
  l <- gsub("\a", " ", l, fixed = TRUE)
  l <- gsub("\t", " ", l, fixed = TRUE)
  l <- gsub("\f", "", l, fixed = TRUE)
  m <- regexpr("^P[1-9](?=\\s|$)", l, perl = TRUE)
  if (m > 0) {
    k <- attr(m, "match.length")
    l <- paste0(strrep(" ", k), substring(l, k + 1))
  }
  sub("\\s+$", "", l)
}

FOOT_RE <- paste0(
  "(^\\s*[0-9]{1,3}\\s+[0-9]{1,2}(st|nd|rd|th)\\s+Annual Report)|",
  "(Business Responsibility and Sustainability Report\\s+[0-9]+)|",
  "(", company_footer, ")|",
  "(^\\s*Section A\\s+Section B\\s+Section C)|",
  "(^\\s*Corporate\\s+Statutory\\s+Financial)|",
  "(^\\s*Overview\\s+Reports\\s+Statements)|",
  "(^\\s*[0-9]{1,3}\\s*$)|(^\\s*Section [ABC]\\s*$)|",
  "(^\\s*(Financial\\s+)?Statements\\s*$)|(^\\s*(Statutory\\s+)?Reports\\s*$)|",
  "(^\\s*Corporate\\s*$)|(^\\s*Overview\\s*$)"
)

# a table cell counts as a VALUE if it is a number, a %, a number with a
# unit word (e.g. "8.50 lakh"), or NIL / N.A. / dash
VAL_RE <- paste0(
  "^(-?[0-9,]+(\\.[0-9]+)?%?|",
  "-?[0-9,]+(\\.[0-9]+)?\\s*(lakh|crore|million|MW|KW|KWh)|",
  "NIL|Nil|N\\.A\\.?|NA|-)$"
)

# split a line into cells separated by 2+ spaces, remembering positions
get_cells <- function(l) {
  m <- gregexpr("\\S+(?: \\S+)*", l, perl = TRUE)[[1]]
  if (m[1] == -1) {
    return(data.frame(start = integer(0), end = integer(0),
                      text = character(0), stringsAsFactors = FALSE))
  }
  len <- attr(m, "match.length")
  data.frame(start = as.integer(m),
             end   = as.integer(m) + len - 1L,
             text  = substring(l, m, m + len - 1L),
             stringsAsFactors = FALSE)
}

parse_val <- function(txt) {
  mm <- regmatches(txt, regexec(
    "^(-?[0-9,]+(?:\\.[0-9]+)?)\\s*(lakh|crore|million|MW|KW|KWh)$",
    txt, perl = TRUE))[[1]]
  if (length(mm) > 0) {
    return(list(num = as.numeric(gsub(",", "", mm[2])), unit_word = mm[3]))
  }
  if (grepl("^-?[0-9,]+(\\.[0-9]+)?%?$", txt, perl = TRUE)) {
    return(list(num = suppressWarnings(as.numeric(gsub("[,%]", "", txt))),
                unit_word = ""))
  }
  list(num = NA_real_, unit_word = "")
}

infer_unit <- function(txt, unit_word, ctx) {
  if (grepl("%$", txt)) return("%")
  if (nzchar(unit_word)) {
    return(switch(tolower(unit_word),
                  lakh = "INR lakh", crore = "INR crore",
                  million = "Million", mw = "MW", kw = "KW", kwh = "KWh",
                  unit_word))
  }
  if (grepl("^(NIL|Nil|N\\.A\\.?|NA|-)$", txt, perl = TRUE)) return("")
  rules <- c(
    "Gigajoules"         = "gigajoule",
    "Kilolitres"         = "kilolitre|\\bKL\\b",
    "Metric tonnes CO2e" = "CO2",
    "Metric tonnes"      = "metric tonne|\\bMT\\b",
    "INR crore"          = "crore|\\bCr\\b",
    "INR lakh"           = "lakh",
    "Days"               = "days of"
  )
  for (nm in names(rules)) {
    if (grepl(rules[[nm]], ctx, ignore.case = TRUE, perl = TRUE)) return(nm)
  }
  "Number"
}

## ------------------------------------------------------------
## 2. Isolate Section C
## ------------------------------------------------------------
lines <- get_lines(pdf_path)
s <- grep("SECTION C\\s*:\\s*PRINCIPLE WISE", lines,
          ignore.case = TRUE, perl = TRUE)[1]
if (is.na(s)) stop("Could not find the 'SECTION C' header in the PDF text.")
e_all <- grep("Independent Assurance Statement", lines, ignore.case = TRUE)
e_all <- e_all[e_all > s]
e <- if (length(e_all) > 0) e_all[1] - 1 else length(lines)
sec <- vapply(lines[s:e], clean_line, character(1), USE.NAMES = FALSE)

## ------------------------------------------------------------
## 3. Walk through Section C: track Principle -> Essential/Leadership
##    -> Question, and classify every line as header / data row / text
## ------------------------------------------------------------
long_list <- list()
wide_list <- list()
qmeta <- list()
qtext <- list()
narr  <- list()
qorder <- character(0)

princ <- NA_character_
kind  <- NA_character_
qkey  <- NA_character_
hdr   <- list()
seen  <- FALSE

for (l in sec) {
  if (!nzchar(trimws(l))) next
  if (grepl(FOOT_RE, l, perl = TRUE, ignore.case = TRUE)) next
  
  # --- Principle header -------------------------------------------------
  mp <- regmatches(l, regexec("PRINCIPLE\\s*([0-9])\\s*:", l,
                              ignore.case = TRUE, perl = TRUE))[[1]]
  if (length(mp) > 0) {
    princ <- paste0("P", mp[2]); kind <- NA_character_; qkey <- NA_character_
    next
  }
  # --- Essential / Leadership header ----------------------------------------
  if (grepl("^\\s*Essential Indicators", l, ignore.case = TRUE)) {
    kind <- "Essential"; qkey <- NA_character_; next
  }
  if (grepl("^\\s*Leadership Indicators", l, ignore.case = TRUE)) {
    kind <- "Leadership"; qkey <- NA_character_; next
  }
  if (is.na(princ) || is.na(kind)) next
  
  # --- New numbered question -------------------------------------------------
  mq <- regmatches(l, regexec("^\\s{0,8}([0-9]{1,2})\\.\\s+([^0-9].*)$", l,
                              perl = TRUE))[[1]]
  if (length(mq) > 0) {
    qn   <- as.integer(mq[2])
    qkey <- paste(princ, kind, qn, sep = "|")
    qmeta[[qkey]]  <- list(princ = princ, kind = kind, qn = qn)
    qtext[[qkey]]  <- trimws(mq[3])
    narr[[qkey]]   <- character(0)
    qorder <- c(qorder, qkey)
    hdr  <- list()
    seen <- FALSE
    next
  }
  if (is.na(qkey)) next
  
  cs <- get_cells(l)
  if (nrow(cs) == 0) next
  isv <- grepl(VAL_RE, cs$text, perl = TRUE)
  
  # --- No numbers (or a lone cell): header row or plain text ---------------------
  if (!any(isv) || nrow(cs) < 2) {
    if (!any(isv) && nrow(cs) >= 2) {
      if (seen) { hdr <- list(); seen <- FALSE }   # a new table starts
      hdr[[length(hdr) + 1]] <- cs
      next
    }
    t <- trimws(l)
    if (length(narr[[qkey]]) == 0 && nchar(qtext[[qkey]]) < 400 &&
        !grepl("[?:.)]$", qtext[[qkey]])) {
      qtext[[qkey]] <- paste(qtext[[qkey]], t)      # wrapped question text
    } else {
      narr[[qkey]] <- c(narr[[qkey]], t)            # answer text
    }
    next
  }
  
  # --- Data row: label + values ----------------------------------------------------
  seen  <- TRUE
  vi    <- which(isv)
  first <- cs$start[vi[1]]
  lab_i <- which(!isv & cs$end <= first & nchar(cs$text) <= 45)
  label <- paste(cs$text[lab_i], collapse = " ")
  n     <- length(vi)
  
  entries <- list()
  for (k in seq_len(n)) {
    ctr <- (cs$start[vi[k]] + cs$end[vi[k]]) / 2
    per <- ""
    parts <- character(0)
    for (hl in hdr) {
      d <- abs((hl$start + hl$end) / 2 - ctr)
      j <- which.min(d)
      if (d[j] <= 13 || (ctr >= hl$start[j] - 6 && ctr <= hl$end[j] + 6)) {
        ht <- hl$text[j]
        if (grepl("^FY\\s?[0-9]{4}", ht, ignore.case = TRUE, perl = TRUE)) {
          per <- ht
        } else if (nchar(ht) <= 30) {
          parts <- c(parts, ht)
        }
      }
    }
    txt <- cs$text[vi[k]]
    pv  <- parse_val(txt)
    entries[[k]] <- list(
      Principle = princ, QN = qmeta[[qkey]]$qn, EL = kind,
      Item = label, Period = per, Column = paste(parts, collapse = " "),
      ColNo = k, Value = pv$num,
      Unit = infer_unit(txt, pv$unit_word, paste(l, qtext[[qkey]])),
      Reported = txt, QKey = qkey
    )
  }
  any_period <- any(vapply(entries, function(x) nzchar(x$Period), logical(1)))
  if (n <= 5 || any_period) {
    long_list <- c(long_list, entries)
  } else {
    hd <- unique(unlist(lapply(hdr, function(h) h$text[order(h$start)])))
    wide_list[[length(wide_list) + 1]] <- list(
      Principle = princ, QN = qmeta[[qkey]]$qn, EL = kind,
      Item = label,
      Headers = paste(hd[nchar(hd) <= 30], collapse = " | "),
      Values = paste(cs$text[vi], collapse = " | "),
      QKey = qkey
    )
  }
}

## ------------------------------------------------------------
## 4. Numbers that appear inside written answers
##    (e.g. "102.41 Metric Tonnes", "2.72 Metric Tonnes of compost")
## ------------------------------------------------------------
num_unit_re <- paste0(
  "[0-9][0-9,]*(?:\\.[0-9]+)?\\s*(?:Metric Tonnes|metric tonnes|million units|",
  "million hours|million|MW|KL|crore|lakh|per cent|%)")

for (qk in qorder) {
  txt <- paste(narr[[qk]], collapse = " ")
  if (!nzchar(txt)) next
  mm <- gregexpr(num_unit_re, txt, perl = TRUE)[[1]]
  if (mm[1] == -1) next
  ml <- attr(mm, "match.length")
  for (i in seq_along(mm)) {
    hit  <- substring(txt, mm[i], mm[i] + ml[i] - 1)
    num  <- as.numeric(gsub(",", "", sub("\\s*[A-Za-z% ].*$", "", hit)))
    u    <- tolower(trimws(sub("^[0-9][0-9,]*(?:\\.[0-9]+)?\\s*", "", hit,
                               perl = TRUE)))
    u    <- switch(u, mw = "MW", kl = "KL", `%` = "%",
                   paste0(toupper(substring(u, 1, 1)), substring(u, 2)))
    snip <- trimws(substring(txt, max(1, mm[i] - 70), mm[i] - 1))
    long_list[[length(long_list) + 1]] <- list(
      Principle = qmeta[[qk]]$princ, QN = qmeta[[qk]]$qn, EL = qmeta[[qk]]$kind,
      Item = paste0("(from written answer) ...", snip),
      Period = "", Column = "", ColNo = 1L, Value = num, Unit = u,
      Reported = hit, QKey = qk
    )
  }
}

## ------------------------------------------------------------
## 5. Build the output tables
## ------------------------------------------------------------
ind_label <- function(qk, width = 220) {
  t <- qtext[[qk]]
  if (nchar(t) > width) t <- paste0(substr(t, 1, width), "...")
  paste0("Q", qmeta[[qk]]$qn, ". ", t)
}

long_df <- do.call(rbind, lapply(long_list, function(x) {
  data.frame(Principle = x$Principle, `Q No` = x$QN,
             Indicator = ind_label(x$QKey), `Essential/Leadership` = x$EL,
             Item = x$Item, Period = x$Period, Column = x$Column,
             `Col No` = x$ColNo, Value = x$Value, Unit = x$Unit,
             `As Reported` = x$Reported,
             stringsAsFactors = FALSE, check.names = FALSE)
}))
# keep document order, but make sure P1..P9 and Essential->Leadership are sorted
long_df <- long_df[order(as.integer(sub("P", "", long_df$Principle)),
                         long_df$`Essential/Leadership` != "Essential",
                         long_df$`Q No`), ]
rownames(long_df) <- NULL

wide_df <- if (length(wide_list) > 0) {
  do.call(rbind, lapply(wide_list, function(x) {
    data.frame(Principle = x$Principle, `Q No` = x$QN,
               Indicator = ind_label(x$QKey), `Essential/Leadership` = x$EL,
               `Table Row` = x$Item, `Column Headings (approx.)` = x$Headers,
               `Values (in table order)` = x$Values,
               stringsAsFactors = FALSE, check.names = FALSE)
  }))
} else {
  data.frame(Note = "No wide tables found", stringsAsFactors = FALSE)
}

qual_df <- do.call(rbind, lapply(qorder, function(qk) {
  ans <- paste(narr[[qk]], collapse = " ")
  if (!nzchar(ans)) ans <- "(table / numbers only - see 'Section C Data' and 'Wide Tables')"
  if (nchar(ans) > 2500) ans <- paste0(substr(ans, 1, 2500), "...")
  data.frame(Principle = qmeta[[qk]]$princ, `Q No` = qmeta[[qk]]$qn,
             `Essential/Leadership` = qmeta[[qk]]$kind,
             Indicator = paste0("Q", qmeta[[qk]]$qn, ". ", qtext[[qk]]),
             Response = ans,
             stringsAsFactors = FALSE, check.names = FALSE)
}))

readme <- data.frame(
  Column = c("Principle", "Q No", "Indicator", "Essential/Leadership", "Item",
             "Period", "Column", "Col No", "Value", "Unit", "As Reported"),
  Meaning = c(
    "NGRBC principle P1-P9 the indicator belongs to",
    "Question number within that principle's Essential or Leadership list",
    "The BRSR question being answered (shortened)",
    "Essential = mandatory indicator; Leadership = voluntary indicator",
    "Row label of the table the number came from (e.g. Male, Total, a metric name)",
    "Financial year column the number sits under (FY2024, FY2023, ...), if the table has one",
    "Other column heading(s) detected above the number (best effort, from page layout)",
    "Position of the number within its table row (1 = first number)",
    "The number as a numeric value (blank for NIL / N.A.)",
    "Unit inferred from the number (%, INR crore, Gigajoules, ...) or 'Number' for counts",
    "Exactly what the report printed (keeps NIL, N.A., 8.50 lakh etc.)"
  ),
  stringsAsFactors = FALSE
)

## ------------------------------------------------------------
## 6. Write the Excel file
## ------------------------------------------------------------
wb <- createWorkbook()
hs <- createStyle(textDecoration = "bold", fgFill = "#1F4E78",
                  fontColour = "white", halign = "center",
                  border = "TopBottomLeftRight", wrapText = TRUE)
bs <- createStyle(wrapText = TRUE, valign = "top",
                  border = "TopBottomLeftRight", borderColour = "#D9D9D9")

add_sheet <- function(name, df, widths) {
  addWorksheet(wb, name)
  writeData(wb, name, df, headerStyle = hs)
  setColWidths(wb, name, cols = seq_len(ncol(df)),
               widths = rep_len(widths, ncol(df)))
  if (nrow(df) > 0) {
    addStyle(wb, name, bs, rows = 2:(nrow(df) + 1),
             cols = seq_len(ncol(df)), gridExpand = TRUE)
  }
  freezePane(wb, name, firstRow = TRUE)
  addFilter(wb, name, rows = 1, cols = seq_len(ncol(df)))
}

add_sheet("Read Me", readme, c(24, 100))
add_sheet("Section C Data", long_df, c(10, 7, 55, 14, 34, 10, 26, 7, 12, 18, 14))
add_sheet("Wide Tables", wide_df, c(10, 7, 50, 14, 30, 40, 60))
add_sheet("Qualitative", qual_df, c(10, 7, 14, 60, 110))
add_sheet("Raw Text", data.frame(`Cleaned Section C text` = sec,
                                 check.names = FALSE), 150)
saveWorkbook(wb, out_path, overwrite = TRUE)

cat("Done:", out_path, "\n")
cat("  Section C Data rows:", nrow(long_df), "\n")
cat("  Wide table rows    :", ifelse("Note" %in% names(wide_df), 0, nrow(wide_df)), "\n")
cat("  Indicators found   :", length(qorder), "\n")