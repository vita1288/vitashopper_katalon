/*
 * Test Case : TC02_Insert_Customer_Logout (data-driven)
 * Data      : Data Files/vitashopper_testdata.xlsx -> sheets "Login" and "Customer"
 * Excel     : Customer sheet columns -> Execute | TC_ID | Nama_Customer | Alamat | HP
 *             (auto-filled: Generated_Kode | Status | Remarks | Executed_At)
 * Rule      : Execute is the ONLY switch -> Y = insert, N = skip
 * Flow      : Login (always)
 *             -> check Customer list (last Kode)
 *             -> for each Customer row with Execute = Y:
 *                  next Kode -> Tambah Customer -> fill -> Simpan
 *                  -> Execute auto-set to N immediately -> verify row in list
 *             -> Logout (always)
 *
 * Test Suite settings: Retry failed executions = 0, NO Data Binding on this test case.
 */
import com.kms.katalon.core.util.KeywordUtil
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.vitashopper.ExcelHelper

final String SHEET = 'Customer'
final int MAX_INSERT = Integer.MAX_VALUE   // 1 = one customer per run; MAX_VALUE = all pending rows

// ---------- Read Excel + diagnostic log ----------
Map<String, String> cred = ExcelHelper.loginCredential()
List<Map<String, String>> allRows = ExcelHelper.readSheet(SHEET)
List<Map<String, String>> customers = allRows.findAll {
	(ExcelHelper.get(it, false, 'Execute') ?: '').trim().equalsIgnoreCase('Y')
}

KeywordUtil.logInfo("Excel file: ${ExcelHelper.file().absolutePath}")
allRows.each { r ->
	KeywordUtil.logInfo("  Excel row ${(r['_rowIndex'] as int) + 1}: Execute='${ExcelHelper.get(r, false, 'Execute') ?: ''}' " +
			"Nama='${ExcelHelper.get(r, false, 'Nama_Customer', 'Nama', 'Name') ?: ''}'")
}
KeywordUtil.logInfo("Pending rows (Execute = Y): ${customers.collect { (it['_rowIndex'] as int) + 1 }}")

// ---------- STEP 1: Login (always) ----------
WebUI.comment("STEP 1: Login")
CustomKeywords.'com.vitashopper.CustomerManager.login'(
		ExcelHelper.get(cred, 'username', 'user'),
		ExcelHelper.get(cred, 'password', 'pass'))

int inserted = 0

try {
	// ---------- Check Customer list ----------
	WebUI.comment("CHECK: Customer list - next Kode Customer")
	String nextKode = CustomKeywords.'com.vitashopper.CustomerManager.getNextKodeCustomer'()
	KeywordUtil.logInfo("Customer list OK. Next Kode would be: ${nextKode}")
	WebUI.takeScreenshot()

	if (customers.isEmpty()) {
		KeywordUtil.markWarning("No '${SHEET}' rows with Execute = Y. Login + check done, nothing to insert. " +
				"Set Execute = Y on the rows you want to insert.")
		return
	}

	for (Map<String, String> row : customers) {
		int    rowIndex = row['_rowIndex'] as int
		String tcId     = ExcelHelper.get(row, false, 'TC_ID', 'TC ID') ?: '-'
		String kode       = ''
		String status     = 'FAILED'
		String remarks    = ''
		boolean submitted = false

		try {
			String nama   = ExcelHelper.get(row, 'Nama_Customer', 'Nama', 'Name')
			String alamat = ExcelHelper.get(row, 'Alamat', 'Address')
			String hp     = ExcelHelper.get(row, 'HP', 'No_HP', 'Phone')

			WebUI.comment("STEP 2 [${tcId}]: Generate next Kode Customer")
			kode = CustomKeywords.'com.vitashopper.CustomerManager.getNextKodeCustomer'()

			WebUI.comment("STEP 3 [${tcId}]: Fill customer ${kode} - ${nama} (Excel row ${rowIndex + 1})")
			CustomKeywords.'com.vitashopper.CustomerManager.fillCustomer'(kode, nama, alamat, hp)

			// Click Simpan -> IMMEDIATELY auto-set Execute = N
			CustomKeywords.'com.vitashopper.CustomerManager.clickSimpan'()
			submitted = true
			inserted++
			boolean saved = ExcelHelper.writeValuesWithRetry(SHEET, rowIndex, ['Execute': 'N', 'Generated_Kode': kode])
			if (!saved) {
				KeywordUtil.markFailedAndStop("[${tcId}] Simpan clicked for ${kode}, but Excel could not be updated " +
						"after 3 attempts. File is locked: ${ExcelHelper.file().absolutePath}")
			}
			WebUI.comment("[${tcId}] Simpan clicked -> Execute auto-set to N (Excel row ${rowIndex + 1})")

			CustomKeywords.'com.vitashopper.CustomerManager.waitAfterSimpan'()

			WebUI.comment("STEP 4 [${tcId}]: Verify customer in list")
			boolean ok = CustomKeywords.'com.vitashopper.CustomerManager.verifyCustomerExists'(kode, nama)
			status  = ok ? 'PASSED' : 'FAILED'
			remarks = ok ? 'Customer created' : 'Submitted, but row not found in Customer list'

		} catch (Exception e) {
			WebUI.takeScreenshot()
			remarks = e.message
			KeywordUtil.markFailed("[${tcId}] ${e.message}")
		} finally {
			Map<String, String> result = [
				'Status'     : status,
				'Remarks'    : remarks,
				'Executed_At': new Date().format('yyyy-MM-dd HH:mm:ss')
			]
			if (submitted) {
				result['Execute']        = 'N'
				result['Generated_Kode'] = kode
			}
			ExcelHelper.writeValuesWithRetry(SHEET, rowIndex, result)
		}

		if (inserted >= MAX_INSERT) {
			KeywordUtil.logInfo("Inserted ${inserted} customer(s). MAX_INSERT reached, stopping loop.")
			break
		}
		if (!submitted) {
			KeywordUtil.markWarning("[${tcId}] failed before Simpan. Stopping loop.")
			break
		}
	}
	KeywordUtil.logInfo("Run finished. Customers inserted: ${inserted}")

} finally {
	// ---------- Logout (always) ----------
	WebUI.comment("STEP 5: Logout")
	CustomKeywords.'com.vitashopper.CustomerManager.logout'()
}
