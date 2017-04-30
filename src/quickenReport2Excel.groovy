#!/usr/bin/env groovy
//used as a shell script
//https://www.eriwen.com/groovy/groovy-shell-scripts/

def cli = new CliBuilder()

println ("args: $args")
cli.with {
	f args: 1, argName: 'file', 'Quicken Report input file', required: true
}

def arguments = cli.parse(args)
if (!arguments) {
	println "\nExample usage:\nquickenReport2Excel.groovy -f itemized_categories.TXT"
	System.exit(1)
}

println ("Converting quicken report=$arguments.f to excel document=$arguments.f" + ".xls")
QuickenReport2Excel.parseReport(arguments.f)

@Grab( 'org.apache.poi:poi:3.15' )

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFDateUtil
import org.apache.poi.hssf.usermodel.HSSFCellStyle
import org.apache.poi.hssf.usermodel.HSSFFont

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CreationHelper

import org.apache.poi.ss.util.CellReference

import groovy.time.*
import java.io.File
import java.time.*
import java.time.format.DateTimeFormatter

class QuickenReport2Excel {

	static final DateTimeFormatter QUICKEN_REPORT_FORMATTER = DateTimeFormatter.ofPattern("M/d/yyyy")

	static final int QUICKEN_TOKEN_CATEGORY = 1 //shared column
	static final int QUICKEN_TOKEN_DATE = 1 //shared column
	static final int QUIKCEN_TOKEN_TAG = 2
	static final int QUICKEN_TOKEN_COMMENTS = 5
	static final int QUICKEN_TOKEN_AMOUNT = 8

	static final DateTimeFormatter EXCEL_MONTH_YEAR_FORMATTER = DateTimeFormatter.ofPattern("yy-MMM")
	static final int EXCEL_MONTHS_START_BEING_DISPLAYED_COLUMN = 2
	static HSSFWorkbook workbook
	static HSSFSheet sheet
	static HSSFCellStyle excelDateStyle
	static HSSFCellStyle excelBoldStyle
	static int currentRowNum = 0

	static {
		//init apache poi objects
		workbook = new HSSFWorkbook()
		sheet = workbook.createSheet("quicken export")

		//excel bold style
		excelBoldStyle = workbook.createCellStyle()
		HSSFFont font = workbook.createFont()
		font.setBold(true)
		excelBoldStyle.setFont(font)

		//excel date style
		excelDateStyle = workbook.createCellStyle()
		CreationHelper createHelper = workbook.getCreationHelper()
		excelDateStyle.setDataFormat(createHelper.createDataFormat().getFormat("mmm-yy"))
	}

	static void parseReport(def quickenReportFileName) {

		File quickenTabExportFile = new File( quickenReportFileName )
		println "Opening quicken report=" + quickenTabExportFile
		assert quickenTabExportFile.exists() : "quicken report doesn't exist"
		assert workbook != null
		assert sheet != null

		// convenient to get min and max dates first
		println "Getting min and max month dates from quicken report=" + quickenTabExportFile
		def minDate
		def maxDate
		( minDate, maxDate) = getBoundsOfReport(quickenTabExportFile)
		println "Found min date=" + minDate + " and max date=" + maxDate

		println "Parsing quicken report by month=" + quickenTabExportFile
		parseReport(quickenTabExportFile, minDate, maxDate)

		def excelOutputFile = "" + quickenTabExportFile.getAbsoluteFile() + ".xls"
		println "Writing excel file=$excelOutputFile"
		FileOutputStream outputStream = new FileOutputStream(excelOutputFile)
		workbook.write(outputStream)
		workbook.close()
	}

	static def getBoundsOfReport(File quickenTabExportFile) {

		def minDate = null
		def maxDate = null

		quickenTabExportFile.eachLine { line ->
			def quickentokens = line.split( '\t' ).collect { it.trim() }

			try {
				LocalDate date = LocalDate.parse(quickentokens[QUICKEN_TOKEN_DATE], QUICKEN_REPORT_FORMATTER)
				if (minDate == null) {
					minDate = date
				}
				if (maxDate == null) {
					maxDate = date
				}

				if (date.isAfter(maxDate)) {
					maxDate = date
				} else if (date.isBefore(minDate)) {
					minDate = date
				}
			} catch (Exception e) {
				//ignore this, these date column are shared with other types of data and might be empty
			}
		}

		return [minDate.withDayOfMonth(minDate.lengthOfMonth()),maxDate]
	}

	static def parseReport(File quickenTabExportFile, LocalDate minDate, LocalDate maxDate) {

		def startColumn = getExcelColumnLocation(minDate)
		def maxColumns = getExcelColumnLocation(maxDate, startColumn)

		String lastCategory = ""
		int numberSubTotalRowsForCurrentCategory = 0

		quickenTabExportFile.eachLine { line ->
			def quickenTokens = line.split( '\t' ).collect { it.trim() }

			try {
				if (isNewQuickenCateogry(quickenTokens)) {

					if (numberSubTotalRowsForCurrentCategory > 0) {
						createSubTotalRows(sheet, numberSubTotalRowsForCurrentCategory, maxColumns)
						numberSubTotalRowsForCurrentCategory = 0
					} 

					createCategorySectionRows(quickenTokens, maxColumns, minDate)
				}
				else if (isTransactionRow(quickenTokens)) {

					createTransactionRow(startColumn, quickenTokens)
					numberSubTotalRowsForCurrentCategory++
				} else {
					//dont do anything
				}

			} catch(Exception ex) {
				//do nothing, quicken reports are inconsistent
				//println ex
			}			
		}
		createSubTotalRows(sheet, numberSubTotalRowsForCurrentCategory, maxColumns)
	}

	static boolean isTransactionRow(def quickenTokens) {
		return quickenTokens[QUICKEN_TOKEN_AMOUNT] != null
	}

	static boolean isNewQuickenCateogry(def quickenTokens) {
		return quickenTokens[QUICKEN_TOKEN_CATEGORY] != '' && quickenTokens[QUICKEN_TOKEN_AMOUNT] != '' && quickenTokens[QUIKCEN_TOKEN_TAG] == ''
	}

	static int getExcelColumnLocation(LocalDate minDate) {
		return 	(minDate.getYear().intValue() * 12) + minDate.getMonth().getValue()
	}

	static int getExcelColumnLocation(LocalDate minDate, int startColumn) {
		return 	getExcelColumnLocation(minDate) - startColumn
	}

	static def createSubTotalRows(HSSFSheet sheet, int subTotal, int maxColumns) {
		HSSFRow subtotalRow = sheet.createRow(currentRowNum++)

		HSSFCell subTotalCell = subtotalRow.createCell(0)
		subTotalCell.setCellValue("Subtotal")
		subTotalCell.setCellStyle(excelBoldStyle)
		
		for (int column = EXCEL_MONTHS_START_BEING_DISPLAYED_COLUMN; column <= maxColumns + EXCEL_MONTHS_START_BEING_DISPLAYED_COLUMN; column++ ) {
			HSSFCell cell = subtotalRow.createCell(column)
			String sumAllTransactionsInMonthFormula = "SUM("+ CellReference.convertNumToColString(column) + (subtotalRow.getRowNum() - subTotal + 1 ) + ":"+ CellReference.convertNumToColString(column) + (subtotalRow.getRowNum() ) + ")"
			cell.setCellType(HSSFCell.CELL_TYPE_FORMULA)
			cell.setCellFormula(sumAllTransactionsInMonthFormula )
			cell.setCellStyle(excelBoldStyle)
		}

		HSSFCell totalOfAllMonthsCell = subtotalRow.createCell(maxColumns + EXCEL_MONTHS_START_BEING_DISPLAYED_COLUMN + 1)
		String sumOfAllMonthsFormula= "SUM("+ CellReference.convertNumToColString(0 + EXCEL_MONTHS_START_BEING_DISPLAYED_COLUMN) + (subtotalRow.getRowNum() + 1) + ":" + CellReference.convertNumToColString(maxColumns + EXCEL_MONTHS_START_BEING_DISPLAYED_COLUMN) + (subtotalRow.getRowNum() + 1) + ")"
		totalOfAllMonthsCell.setCellType(HSSFCell.CELL_TYPE_FORMULA)
		totalOfAllMonthsCell.setCellFormula(sumOfAllMonthsFormula)
		totalOfAllMonthsCell.setCellStyle(excelBoldStyle)
	}

	static def createCategorySectionRows(def quickeTokens, int maxColumns, LocalDate monthHeader) {

		//space between categories
		sheet.createRow(currentRowNum++)

		//category
		HSSFRow categoryRow = sheet.createRow(currentRowNum++)
		HSSFCell cell = categoryRow.createCell(0)
		cell.setCellStyle(excelBoldStyle)
		cell.setCellValue(quickeTokens[QUICKEN_TOKEN_CATEGORY])

		//month headings
		HSSFRow monthRow = sheet.createRow(currentRowNum++)
		for (int columnNum = EXCEL_MONTHS_START_BEING_DISPLAYED_COLUMN; columnNum <= maxColumns + EXCEL_MONTHS_START_BEING_DISPLAYED_COLUMN; columnNum++ ) {

			HSSFCell monthCell = monthRow.createCell(columnNum)
			monthCell.setCellValue(Date.from(monthHeader.atStartOfDay(ZoneId.systemDefault()).toInstant()))
			monthCell.setCellStyle(excelDateStyle)

			monthHeader = monthHeader.plusMonths(1)
		}
		HSSFCell totalCell = monthRow.createCell(maxColumns + EXCEL_MONTHS_START_BEING_DISPLAYED_COLUMN + 1)
		totalCell.setCellValue("Total")
	}

	static def createTransactionRow(int startColumn, def quickenTokens) {
				
		HSSFRow row = sheet.createRow(currentRowNum++)
		
		HSSFCell commentCell = row.createCell(0)
		def transactionComments = quickenTokens[QUICKEN_TOKEN_COMMENTS]
		commentCell.setCellValue(transactionComments)

		def transactionDate = LocalDate.parse(quickenTokens[QUICKEN_TOKEN_DATE], QUICKEN_REPORT_FORMATTER)
		HSSFCell amountCell = row.createCell(getExcelColumnLocation(transactionDate, startColumn) + EXCEL_MONTHS_START_BEING_DISPLAYED_COLUMN)
		def amount = new BigDecimal(quickenTokens[QUICKEN_TOKEN_AMOUNT].replaceAll(",", ""))
		amountCell.setCellValue(amount)
	}

}