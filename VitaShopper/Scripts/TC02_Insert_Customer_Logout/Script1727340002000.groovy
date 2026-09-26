/*
 * Test Case : TC02_Insert_Customer_Logout (data-driven)
 * Data      : Data Files/vitashopper_testdata.xlsx -> sheets "Login" and "Customer"
 * Excel     : Customer sheet columns -> Execute | TC_ID | Nama_Customer | Alamat | HP
 * Flow      : Login
 *             -> take the first Customer row with Execute = Y
 *                  read last Kode Customer (last page) -> increment (C0265 -> C0266)
 *                  -> Tambah Customer -> fill Nama / Alamat / HP -> Simpan
 *                  -> Execute set to N immediately after Simpan (prevents duplicate insert)
 *                  -> verify row in list -> write result back to Excel
 *             -> STOP after 1 insert
 *             -> Logout
 */
import com.kms.katalon.core.util.KeywordUtil
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.vitashopper.ExcelHelper

final String SHEET = 'Customer'
final int MAX_INSERT = 1          // stop after this many inserts

Map<String, String> cred = ExcelHelper.loginCredential()
List<Map<String, String>> customers = ExcelHelper.readExecutableRows(SHEET)   // Execute = Y only

if (customers.isEmpty()) {
	KeywordUtil.markWarning("No '${SHEET}' rows with Execute = Y. Nothing to insert.")
	return
}

WebUI.comment("STEP 1: Login")
CustomKeywords.'com.vitashopper.CustomerManager.login'(
		ExcelHelper.get(cred, 'username', 'user'),
		ExcelHelper.get(cred, 'password', 'pass'))

int inserted = 0

try {
	for (Map<String, String> row : customers) {
		String tcId      = ExcelHelper.get(row, false, 'TC_ID', 'TC ID') ?: '-'
		int    rowIndex  = row['_rowIndex'] as int
		String kode      = ''
		String status    = 'FAILED'
		String remarks   = ''
		boolean submitted = false

		try {
			String nama   = ExcelHelper.get(row, 'Nama_Customer', 'Nama', 'Name')
			String alamat = ExcelHelper.get(row, 'Alamat', 'Address')
			String hp     = ExcelHelper.get(row, 'HP', 'No_HP', 'Phone')

			WebUI.comment("STEP 2 [${tcId}]: Generate next Kode Customer")
			kode = CustomKeywords.'com.vitashopper.CustomerManager.getNextKodeCustomer'()

			WebUI.comment("STEP 3 [${tcId}]: Insert customer ${kode} - ${nama}")
			CustomKeywords.'com.vitashopper.CustomerManager.addCustomer'(kode, nama, alamat, hp)

			// Simpan clicked -> mark row as done right away, so it is never inserted again
			submitted = true
			inserted++
			ExcelHelper.writeValues(SHEET, rowIndex, ['Execute': 'N', 'Generated_Kode': kode])
			WebUI.comment("[${tcId}] Execute changed to N")

			WebUI.comment("STEP 4 [${tcId}]: Verify customer in list")
			boolean ok = CustomKeywords.'com.vitashopper.CustomerManager.verifyCustomerExists'(kode, nama)
			status  = ok ? 'PASSED' : 'FAILED'
			remarks = ok ? 'Customer created' : 'Submitted, but row not found in Customer list - check manually'

		} catch (Exception e) {
			WebUI.takeScreenshot()
			remarks = e.message
			KeywordUtil.markFailed("[${tcId}] ${e.message}")
		} finally {
			Map<String, String> result = [
				'Generated_Kode': kode,
				'Status'        : status,
				'Remarks'       : remarks,
				'Executed_At'   : new Date().format('yyyy-MM-dd HH:mm:ss')
			]
			if (submitted) result['Execute'] = 'N'
			ExcelHelper.writeValues(SHEET, rowIndex, result)
		}

		// Stop after insert, or on failure before insert
		if (inserted >= MAX_INSERT) {
			KeywordUtil.logInfo("Inserted ${inserted} customer(s). Stopping loop.")
			break
		}
		if (!submitted) {
			KeywordUtil.markWarning("[${tcId}] failed before Simpan. Stopping loop.")
			break
		}
	}
} finally {
	WebUI.comment("STEP 5: Logout")
	CustomKeywords.'com.vitashopper.CustomerManager.logout'()
}