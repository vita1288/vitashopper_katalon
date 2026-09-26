package com.vitashopper

import com.kms.katalon.core.configuration.RunConfiguration
import com.kms.katalon.core.util.KeywordUtil

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory

import internal.GlobalVariable

/**
 * Excel read/write for test data (.xlsx / .xls).
 * File: <ProjectDir>/<GlobalVariable.testDataFile>. Row 1 = header.
 * Close the file in Excel before running, otherwise write-back is skipped.
 */
class ExcelHelper {

	private static final DataFormatter FMT = new DataFormatter()

	static File file() {
		File f = new File(RunConfiguration.getProjectDir(), GlobalVariable.testDataFile.toString())
		if (!f.exists()) KeywordUtil.markFailedAndStop("Excel file not found: ${f.absolutePath}")
		return f
	}

	private static Workbook open() {
		return file().withInputStream { WorkbookFactory.create(it) }
	}

	private static Sheet sheet(Workbook wb, String sheetName) {
		Sheet sh = wb.getSheet(sheetName)
		if (sh == null) KeywordUtil.markFailedAndStop("Sheet '${sheetName}' not found in ${file().name}")
		if (sh.getRow(0) == null) KeywordUtil.markFailedAndStop("Sheet '${sheetName}' has no header row")
		return sh
	}

	/** Cell as text: formulas evaluated, numbers without scientific notation or ".0". */
	private static String cellText(Cell c, FormulaEvaluator ev) {
		if (c == null) return ''
		CellType type = c.getCellType() == CellType.FORMULA ? ev.evaluateFormulaCell(c) : c.getCellType()
		if (type == CellType.NUMERIC && !DateUtil.isCellDateFormatted(c)) {
			return BigDecimal.valueOf(c.getNumericCellValue()).stripTrailingZeros().toPlainString()
		}
		return FMT.formatCellValue(c, ev).trim()
	}

	private static List<String> headers(Sheet sh) {
		Row h = sh.getRow(0)
		int last = Math.max(h.getLastCellNum() as int, 0)
		return (0..<last).collect { FMT.formatCellValue(h.getCell(it)).trim() }
	}

	/** Returns rows as [header: value] maps; '_rowIndex' holds the physical Excel row (0-based). */
	static List<Map<String, String>> readSheet(String sheetName) {
		Workbook wb = open()
		try {
			Sheet sh = sheet(wb, sheetName)
			FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator()
			List<String> cols = headers(sh)
			List<Map<String, String>> rows = []

			for (int r = 1; r <= sh.getLastRowNum(); r++) {
				Row row = sh.getRow(r)
				if (row == null) continue
				Map<String, String> m = [:]
				cols.eachWithIndex { h, i -> if (h) m[h] = cellText(row.getCell(i), ev) }
				if (m.values().any { it }) {
					m['_rowIndex'] = r.toString()
					rows << m
				}
			}
			KeywordUtil.logInfo("Loaded ${rows.size()} row(s) from '${sheetName}'")
			return rows
		} finally {
			wb.close()
		}
	}

	static List<Map<String, String>> readExecutableRows(String sheetName) {
		return readSheet(sheetName).findAll { get(it, false, 'Execute')?.equalsIgnoreCase('Y') }
	}

	static Map<String, String> loginCredential() {
		List<Map<String, String>> rows = readSheet('Login')
		if (rows.isEmpty()) KeywordUtil.markFailedAndStop("Sheet 'Login' is empty")
		return rows[0]
	}

	/** Normalizes a header: lowercase, strips spaces / underscores / dashes / dots. */
	private static String norm(String s) {
		return (s ?: '').toLowerCase().replaceAll(/[\s_\-\.]+/, '')
	}

	/**
	 * Gets a value by header name, ignoring case, spaces, underscores and dashes, with aliases.
	 * e.g. get(row, 'Nama Customer', 'Nama') matches "Nama_Customer", "NAMA CUSTOMER", "Nama".
	 * Stops the test with the available headers if nothing matches or the cell is empty.
	 */
	static String get(Map<String, String> row, String... keys) {
		return get(row, true, keys)
	}

	static String get(Map<String, String> row, boolean required, String... keys) {
		for (String k : keys) {
			def entry = row.find { h, v -> norm(h) == norm(k) }
			if (entry != null && entry.value) return entry.value
		}
		if (required) {
			def available = row.keySet().findAll { it != '_rowIndex' }
			KeywordUtil.markFailedAndStop("Excel value not found for ${keys as List} " +
					"(Excel row ${(row['_rowIndex'] as int) + 1}). Available headers: ${available}")
		}
		return null
	}

	/** Writes values into the given row by header name; creates the column if missing. */
	static synchronized void writeValues(String sheetName, int rowIndex, Map<String, String> values) {
		try {
			Workbook wb = open()
			try {
				Sheet sh = sheet(wb, sheetName)
				Row header = sh.getRow(0)
				Row row = sh.getRow(rowIndex) ?: sh.createRow(rowIndex)

				values.each { String key, String val ->
					int col = headers(sh).findIndexOf { norm(it) == norm(key) }
					if (col < 0) {
						col = Math.max(header.getLastCellNum() as int, 0)
						header.createCell(col).setCellValue(key)
					}
					(row.getCell(col) ?: row.createCell(col)).setCellValue(val ?: '')
				}
				file().withOutputStream { wb.write(it) }
				KeywordUtil.logInfo("Excel updated '${sheetName}' row ${rowIndex + 1}: ${values}")
			} finally {
				wb.close()
			}
		} catch (IOException e) {
			KeywordUtil.markWarning("Excel write-back skipped - close the file in Excel first: ${e.message}")
		} catch (Exception e) {
			KeywordUtil.markWarning("Excel write-back failed: ${e.class.simpleName}: ${e.message}")
		}
	}
}