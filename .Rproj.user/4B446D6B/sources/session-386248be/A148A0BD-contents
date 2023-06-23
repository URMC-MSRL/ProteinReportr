#Make a helper function to change column names ----
## Make vectors to change column names
change_col_names <- function(df1,
                             df2) {
   name_vector <- dplyr::slice(df1,
                               1) %>%
      unlist(use.names = FALSE) %>%
      unname()

   blank_vector <- dplyr::slice(df1,
                                2) %>%
      unlist(use.names = FALSE) %>%
      unname()

   merged_vector <- stats::setNames(blank_vector,
                                    name_vector)

   renamed_df2 <- dplyr::rename_all(df2,
                                    ~stringr::str_replace_all(.x,
                                                              merged_vector))

   return(renamed_df2)
}

#' Convert Perseus report to Excel report
#'
#' @param perseus_file The file name of the Perseus output. Needs to be in quotes and end in '.txt'.
#' @param user_name The last name of the user. Needs to be in quotes and all lowercase.
#' @param work_order The work order number of the current project. Needs to be in quotes and year and work number separated by '_'.
#' @param excel_file The file name you wish for you excel report to be. Needs to be in quotes and end with '.xlsx'.
#' @param is_median_abundance Declare whether or not your data has median abundance columns. If there is, enter 1. If not, enter 0.
#' @return An excel file written to your active directory.
#' @export
#' @examples
#' make_excel_protein_report('Swovick_ProteinReport_23-001.txt',
#'                           'swovick',
#'                           '23_001',
#'                           'Swovick_Excel_ProteinReport.xlsx')

#Make the actual function -----

make_excel_protein_report <- function(perseus_file,
                                      user_name,
                                      work_order,
                                      excel_file,
                                      is_median_abundance = 1) {

   #Establish values needed later to change column names----
   is_median_abundance2 = is_median_abundance
   user2 <- glue::glue('{user_name}_')
   wo2 <- glue::glue('_{work_order}')
   number_peptides <- 'number_peptides_'
   ttest_pvalue <- 'Student.s.T.test.p.value.'
   comp <- '_'
   ttest_log2fc <- 'Student.s.T.test.Difference.'

   #Load data from Perseus----
   protein_report <- PerseusR::read.perseus(perseus_file)

   #Separate the protein report into constiuent data frames----
   ##Protein annotation----
   protein_annotation <- PerseusR::annotCols(protein_report) %>%
      dplyr::select(c('Protein.Accession',
                      'Genes',
                      'Protein.Name')) %>%
      purrr::set_names('Protein Acession',
                       'Genes',
                       'Protein Name')

   ##Normalized Abundance----
   abundances <- PerseusR::main(protein_report) %>%
      janitor::clean_names()
   ###Use the custom name change function to remove user name and work order
   ####Make a df with the names 'user2' and 'wo2' that contains the user name
   ####and work order in one row and then add another empty row
   abundance_name_change <- tibble::tibble(user2,
                                           wo2) %>%
      dplyr::add_row(user2 = '',
                     wo2 = '')
   ####Using the new df just made, we repalce all unwanted naming
   abundances <- change_col_names(df1 = abundance_name_change,
                                  df2 = abundances)

   ##Number of Peptides----
   n_peptides <- PerseusR::annotCols(protein_report) %>%
      dplyr::select(dplyr::contains('Peptide')) %>%
      janitor::clean_names()
   ###Change the column names using the same method as normalized abundances.
   ###There is just one additional paramter added to remove 'number_peptides_'
   ###from each column.
   peptide_name_change <- tibble::tibble(number_peptides,
                                         user2,
                                         wo2) %>%
      dplyr::add_row(number_peptides = '',
                     user2 = '',
                     wo2 = '')

   n_peptides <- change_col_names(df1 = peptide_name_change,
                                  df2 = n_peptides)

   ## t-test Data----
   ### p-value----
   pvalue <- PerseusR::annotCols(protein_report) %>%
      dplyr::select(dplyr::contains('p.value'))
   ####Use the name change function as before the change column names. Also
   #### removes 'ttest_pvalue' and the quotient symbol.
   pvalue_name_change <- tibble::tibble(ttest_pvalue,
                                        comp) %>%
      dplyr::add_row(ttest_pvalue = '',
                     comp = '/')

   pvalue <- change_col_names(df1 = pvalue_name_change,
                              df2 = pvalue)

   ### log2 Fold Change----
   log2fc <- PerseusR::annotCols(protein_report) %>%
      dplyr::select(dplyr::contains('Difference'))
   ####Use the same name change function again but now we also remove
   ####Student.s.T.test.Difference.
   log2fc_name_change <- tibble::tibble(ttest_log2fc,
                                        comp) %>%
      dplyr::add_row(ttest_log2fc = '',
                     comp = '/')

   log2fc <- change_col_names(df1 = log2fc_name_change,
                              df2 = log2fc)

   ### Median Protein Abundance----
   median_abundance <- PerseusR::annotCols(protein_report) %>%
      dplyr::select(-c(dplyr::contains('Peptide'),
                       dplyr::contains('p.value'),
                       dplyr::contains('Difference'),
                       'Protein.Accession',
                       'Genes',
                       'Protein.Name'))

   # Build the excel workbook ----
   ## Create an empty workbook ----
   wb <- openxlsx::createWorkbook()
   openxlsx::addWorksheet(wb,
                          'Data')
   ## Add column category names ----
   ### Establish starting position for each category ----
   ###This helps automate the formatting regardless of number of samples,
   ###groups, and samples per group.
   #### Number of peptides
   pep_col <- 4
   pep_end <- 4 + ncol(n_peptides) -1
   #### Establish a conditional based on if there is a median abundance column.
   #### Otherwise, if there are no medians, the column number for everything
   #### after 'number of peptides' gets shifted to the left by 1.
   if(is_median_abundance2 == 1) {
      #### Median abundance
      med_col <- pep_end + 1
      med_end <- med_col + ncol(median_abundance) - 1
      #### Log2 fold change
      log2_col <- med_end + 1
      log2_end <- log2_col + ncol(log2fc) - 1
   } else {
      #### Log2 fold change
      log2_col <- pep_end +1
      log2_end <- log2_col + ncol(log2fc) - 1
   }
   #### p-value
   pval_col <- log2_col + ncol(log2fc)
   pval_end <- pval_col + ncol(pvalue) - 1
   #### Normalized abundance
   norm_col <- pval_col + ncol(pvalue)
   norm_end <- norm_col + ncol(abundances) - 1

   ### Create the column names ----
   #### Create conditional for if there are median abundances. If there are no
   #### median abundances, the column names will not have the correct spacing.
   if(is_median_abundance2 == 1) {
      openxlsx::writeData(wb,
                          'Data',
                          'Number of Peptides',
                          startCol = pep_col,
                          startRow = 1)
      openxlsx::writeData(wb,
                          'Data',
                          'Median Abundances',
                          startCol = med_col,
                          startRow = 1)
   } else {
      openxlsx::writeData(wb,
                          'Data',
                          'Number of Peptides',
                          startCol = pep_col,
                          startRow = 1)
   }
   openxlsx::writeData(wb,
                       'Data',
                       'Log2 Fold Change',
                       startCol = log2_col,
                       startRow = 1)
   openxlsx::writeData(wb,
                       'Data',
                       'P-Value',
                       startCol = pval_col)
   openxlsx::writeData(wb,
                       'Data',
                       'Normalized Abundances',
                       startCol = norm_col)
   ### Merge column names together ----
   #### Create conditional for if there are median abundances. If there are no
   #### median abundances, the cell merging will not function correctly since
   #### the columns will be shifted to the left by 1.
   if(is_median_abundance2 == 1) {
      openxlsx::mergeCells(wb,
                           'Data',
                           cols = pep_col:pep_end,
                           rows = 1)
      openxlsx::mergeCells(wb,
                           'Data',
                           cols = med_col:med_end,
                           rows = 1)
   } else {
      openxlsx::mergeCells(wb,
                           'Data',
                           cols = pep_col:pep_end,
                           rows = 1)
   }
   openxlsx::mergeCells(wb,
                        'Data',
                        cols = log2_col:log2_end,
                        rows = 1)
   openxlsx::mergeCells(wb,
                        'Data',
                        cols = pval_col:pval_end,
                        rows = 1)
   openxlsx::mergeCells(wb,
                        'Data',
                        cols = norm_col:norm_end,
                        rows = 1)
   openxlsx::mergeCells(wb,
                        'Data',
                        cols = 1,
                        rows = 1:2)
   openxlsx::mergeCells(wb,
                        'Data',
                        cols = 2,
                        rows = 1:2)
   openxlsx::mergeCells(wb,
                        'Data',
                        cols = 3,
                        rows = 1:2)

   ### Add Perseus data into workbook ----
   #### Protein annotation ----
   openxlsx::writeData(wb,
                       'Data',
                       protein_annotation,
                       startCol = 1,
                       startRow = 1)
   #### Number of peptides ----
   openxlsx::writeData(wb,
                       'Data',
                       n_peptides,
                       startCol = pep_col,
                       startRow = 2)
   #### Median abundance ----
   #### Only performs this if 'median_abundance' is == 1.
   if(is_median_abundance2 == 1) {
      openxlsx::writeData(wb,
                          'Data',
                          median_abundance,
                          startCol = med_col,
                          startRow = 2)
   }
   #### Log2 Fold Change ----
   openxlsx::writeData(wb,
                       'Data',
                       log2fc,
                       startCol = log2_col,
                       startRow = 2)
   #### p-value ----
   openxlsx::writeData(wb,
                       'Data',
                       pvalue,
                       startCol = pval_col,
                       startRow = 2)
   #### Normalized abundance ----
   openxlsx::writeData(wb,
                       'Data',
                       abundances,
                       startCol = norm_col,
                       startRow = 2)
   ## Formatting ----
   ### Add vertical borders ----
   #### Thick borders between data types ----
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(border = 'Right',
                                                    borderStyle = 'medium'),
                      cols = 3,
                      rows = 3:length(protein_annotation[[1]]))
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(border = 'Right',
                                                    borderStyle = 'medium'),
                      cols = pep_end,
                      rows = 3:length(protein_annotation[[1]]))
   #### Only performs if 'is_median_abundance' = 1.
   if(is_median_abundance2 == 1) {
      openxlsx::addStyle(wb = wb,
                         sheet = 'Data',
                         style = openxlsx::createStyle(border = 'Right',
                                                       borderStyle = 'medium'),
                         cols = med_end,
                         rows = 3:length(protein_annotation[[1]]))
   }
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(border = 'Right',
                                                    borderStyle = 'medium'),
                      cols = log2_end,
                      rows = 3:length(protein_annotation[[1]]))
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(border = 'Right',
                                                    borderStyle = 'medium'),
                      cols = pval_end,
                      rows = 3:length(protein_annotation[[1]]))
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(border = 'Right',
                                                    borderStyle = 'medium'),
                      cols = norm_end,
                      rows = 3:length(protein_annotation[[1]]))

   ### Column Headers ----
   #### Bold and bottom borders ----
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(fontColour = 'black',
                                                    halign = 'center',
                                                    valign = 'center',
                                                    textDecoration = 'Bold'),
                      rows = 1,
                      cols = 1:norm_end)
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(fontColour = 'black',
                                                    halign = 'center',
                                                    valign = 'center',
                                                    textDecoration = 'Bold',
                                                    border = 'Bottom',
                                                    borderStyle = 'medium'),
                      rows = 2,
                      cols = 1:norm_end)

   #### Add borders to most right category column ----
   #### Protein annotation ----
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(fontColour = 'black',
                                                    halign = 'center',
                                                    valign = 'center',
                                                    textDecoration = 'Bold',
                                                    border = c('right',
                                                               'bottom'),
                                                    borderStyle = 'medium'),
                      rows = 1:2,
                      cols = 3)

   #### Number of Peptides ----
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(fontColour = 'black',
                                                    halign = 'center',
                                                    valign = 'center',
                                                    textDecoration = 'Bold',
                                                    border = 'Right',
                                                    borderStyle = 'medium'),
                      rows = 1,
                      cols = pep_end)
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(fontColour = 'black',
                                                    halign = 'center',
                                                    valign = 'center',
                                                    textDecoration = 'Bold',
                                                    border = c('right',
                                                               'bottom'),
                                                    borderStyle = 'medium'),
                      rows = 2,
                      cols = pep_end)

   #### Median abundance ----
   #### Only performs if 'is_median_abundance' == 1.
   if(is_median_abundance2 == 1) {
      openxlsx::addStyle(wb = wb,
                         sheet = 'Data',
                         style = openxlsx::createStyle(fontColour = 'black',
                                                       halign = 'center',
                                                       valign = 'center',
                                                       textDecoration = 'Bold',
                                                       border = 'Right',
                                                       borderStyle = 'medium'),
                         rows = 1,
                         cols = med_end)
      openxlsx::addStyle(wb = wb,
                         sheet = 'Data',
                         style = openxlsx::createStyle(fontColour = 'black',
                                                       halign = 'center',
                                                       valign = 'center',
                                                       textDecoration = 'Bold',
                                                       border = c('right',
                                                                  'bottom'),
                                                       borderStyle = 'medium'),
                         rows = 2,
                         cols = med_end)
   }

   #### Log2 fold change ----
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(fontColour = 'black',
                                                    halign = 'center',
                                                    valign = 'center',
                                                    textDecoration = 'Bold',
                                                    border = 'Right',
                                                    borderStyle = 'medium'),
                      rows = 1,
                      cols = log2_end)
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(fontColour = 'black',
                                                    halign = 'center',
                                                    valign = 'center',
                                                    textDecoration = 'Bold',
                                                    border = c('right',
                                                               'bottom'),
                                                    borderStyle = 'medium'),
                      rows = 2,
                      cols = log2_end)

   #### P-value ----
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(fontColour = 'black',
                                                    halign = 'center',
                                                    valign = 'center',
                                                    textDecoration = 'Bold',
                                                    border = 'Right',
                                                    borderStyle = 'medium'),
                      rows = 1,
                      cols = pval_end)
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(fontColour = 'black',
                                                    halign = 'center',
                                                    valign = 'center',
                                                    textDecoration = 'Bold',
                                                    border = c('right',
                                                               'bottom'),
                                                    borderStyle = 'medium'),
                      rows = 2,
                      cols = pval_end)

   #### Normalized Abundances ----
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(fontColour = 'black',
                                                    halign = 'center',
                                                    valign = 'center',
                                                    textDecoration = 'Bold',
                                                    border = 'Right',
                                                    borderStyle = 'medium'),
                      rows = 1,
                      cols = norm_end)
   openxlsx::addStyle(wb = wb,
                      sheet = 'Data',
                      style = openxlsx::createStyle(fontColour = 'black',
                                                    halign = 'center',
                                                    valign = 'center',
                                                    textDecoration = 'Bold',
                                                    border = c('right',
                                                               'bottom'),
                                                    borderStyle = 'medium'),
                      rows = 2,
                      cols = norm_end)

   ### Formatting Cell Contents ----
   #### Protein Annotation ----
   openxlsx::setColWidths(wb = wb,
                          sheet = 'Data',
                          cols = 1:2,
                          widths = 20)
   openxlsx::setColWidths(wb = wb,
                          sheet = 'Data',
                          cols = 3,
                          widths = 45)
   #### Number of Peptides ----
   openxlsx::setColWidths(wb = wb,
                          sheet = 'Data',
                          cols = 4:pep_end,
                          widths = 5)
   # Save the excel workbook ----
   openxlsx::saveWorkbook(wb,
                          excel_file,
                          overwrite = TRUE)
}
