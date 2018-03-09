# stream-analysis-programs

Spreadsheets and scripts for the analysis of river morphology, hydraulics and hydrology.
###
excelerateTemplatePebble.xlsm - for creating histogram/cumulative percentage charts of stream pebble counts.

excelerateTemplateProfile_FxlAssessment.xlsm	- for analyzing single-year stream profile data. Created with TN state functional assessments in mind.
excelerateTemplateProfile_Monitoring.xlsm - for analyzing and displaying multi-year stream profile evolution. Created with TN state monitoring protocols in mind.

excelerateTemplateXS_NC.xlsm	- for analyzing and displaying multi-year stream cross section data. Created with NC state monitoring protocols and functional assessments in mind.

excelerateTemplateXS_TN.xlsm - for analyzing and displaying multi-year stream cross section data. Created with TN state monitoring protocols and functional assessments in mind.

offcoordinateElevationCorrection.xlsx	- for finding and correcting survey deviations. Matches survey shots based on description and applies an appropriate correct to all shots in the same setup or linked setup.

regCurvePredictionWorksheet.xlsx	- for finding confidence bands for regional curve data.

stretchWorksheet.xlsx	- worksheet associated with surveyStretch.R

surveyCleaner.R	- separates stream survey data using the description field of a text file and returns multi-tabbed Excel sheets formatted to be pasted into Excelerate spreadsheets.

surveyStretch.R	- uses user-designated points to stretch or compress stream profiles to reduce year-to-year profile length deviation due to changes in real or assumed thalweg path. Points designated are assumed to be immovable, e.g., rock and log structures

xsReduction.xlsx - Uses as stochastic algorithm to reduce the number of shots in a cross section for the purposes of fitting the shots onto a chart
