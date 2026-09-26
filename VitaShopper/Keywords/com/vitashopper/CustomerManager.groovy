package com.vitashopper

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testobject.ConditionType
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.util.KeywordUtil
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI

import org.openqa.selenium.WebElement

import internal.GlobalVariable

/**
* Backoffice VitaShopper x Kirisuna - Customer module
* Kode Customer is auto-incremented from the last record in the Customer list (e.g. C0265 -> C0266).
*/
class CustomerManager {

   private static final int TIMEOUT = 10

   // ---------- Locators (verified on live site, 26 Sep 2026) ----------
   // Login (Default.aspx) - title "Vita Shopper | Log in"
   private static final String TXT_USERNAME = "//input[@id='t1']"
   private static final String TXT_PASSWORD = "//input[@id='t2']"
   private static final String BTN_SIGN_IN  = "//input[@type='submit' and @value='Sign In']"

   // Add Customer (AddCustomer.aspx) - all 4 fields are required
   private static final String TXT_KODE   = "//input[@id='ctl00_ContentPlaceHolder1_t1']"   // Kode Customer
   private static final String TXT_NAMA   = "//input[@id='ctl00_ContentPlaceHolder1_t2']"   // Nama Customer
   private static final String TXT_ALAMAT = "//input[@id='ctl00_ContentPlaceHolder1_t3']"   // Alamat
   private static final String TXT_HP     = "//input[@id='ctl00_ContentPlaceHolder1_t4']"   // HP
   private static final String BTN_SIMPAN = "//input[@id='ctl00_ContentPlaceHolder1_btnSubmit']" // value="Simpan"

   // Sidebar
   private static final String MENU_CUSTOMER = "//a[@href='Customer.aspx']"
   private static final String MENU_LOGOUT   = "//a[@href='Logout.aspx']"

   // Customer list (Customer.aspx)
   private static final String BTN_TAMBAH_CUSTOMER = "//a[@href='AddCustomer.aspx']"
   private static final String CUSTOMER_TABLE = "//table[@id='ctl00_ContentPlaceHolder1_GridView1']"
   private static final String TABLE_CELLS    = CUSTOMER_TABLE + "/tbody/tr/td[2]"            // Kode Customer column
   private static final String PAGER_ROW      = CUSTOMER_TABLE + "//tr[contains(@class,'pagination')]"
   private static final String PAGER_LINKS    = PAGER_ROW + '//a[contains(@href,"Page$")]'   // incl. "Last Page" (Page$Last)
   private static final String PAGER_CURRENT  = PAGER_ROW + "//span"                            // current page number

   // Kode format: letters + digits, e.g. C0265
   private static final String KODE_REGEX = /^([A-Za-z]+)(\d+)$/

   // ---------- Helpers ----------
   private static TestObject obj(String name, String xpath) {
	   TestObject to = new TestObject(name)
	   to.addProperty("xpath", ConditionType.EQUALS, xpath)
	   return to
   }

   private static String baseUrl() {
	   return GlobalVariable.baseUrl.toString().replaceAll('/+$', '')
   }

   private static List<WebElement> findAll(String name, String xpath, int timeout = 3) {
	   try {
		   return WebUI.findWebElements(obj(name, xpath), timeout) ?: []
	   } catch (Exception ignored) {
		   return []
	   }
   }

   private static void waitPostBack() {
	   WebUI.waitForPageLoad(TIMEOUT)
	   WebUI.delay(1)
   }

   private static int pageFromHref(String href) {
	   def m = (href =~ /Page\$(\d+)/)
	   return m.find() ? (m.group(1) as int) : -1
   }

   private static int currentPage() {
	   def spans = findAll('pagerCurrent', PAGER_CURRENT, 1)
	   def num = spans.collect { it.getText().trim() }.find { it ==~ /\d+/ }
	   return num ? (num as int) : 1
   }

   // ---------- Keywords ----------
   @Keyword
   def login(String username, String password) {
	   WebUI.openBrowser("")
	   WebUI.maximizeWindow()
	   WebUI.navigateToUrl(baseUrl() + "/Default.aspx")

	   WebUI.waitForElementVisible(obj("txtUsername", TXT_USERNAME), TIMEOUT, FailureHandling.STOP_ON_FAILURE)
	   WebUI.setText(obj("txtUsername", TXT_USERNAME), username)
	   WebUI.setText(obj("txtPassword", TXT_PASSWORD), password)
	   WebUI.click(obj("btnSignIn", BTN_SIGN_IN))
	   WebUI.waitForPageLoad(TIMEOUT)

	   if (!WebUI.waitForElementNotPresent(obj("txtUsername", TXT_USERNAME), 5, FailureHandling.OPTIONAL)) {
		   WebUI.takeScreenshot()
		   KeywordUtil.markFailedAndStop("Login failed for user '${username}'")
	   }
	   KeywordUtil.logInfo("Login successful: ${WebUI.getUrl()}")
   }

   @Keyword
   def openCustomerList() {
	   WebUI.click(obj("menuCustomer", MENU_CUSTOMER))
	   WebUI.waitForElementVisible(obj("customerTable", CUSTOMER_TABLE), TIMEOUT, FailureHandling.STOP_ON_FAILURE)
   }

   /** Navigates the GridView pager to the last page ("Last Page" link or highest page number, incl. "..."). */
   @Keyword
   def goToLastPage() {
	   openCustomerList()

	   for (int i = 0; i < 50; i++) {
		   List<WebElement> links = findAll('pagerLinks', PAGER_LINKS)
		   if (links.isEmpty()) break   // single page, no pager

		   int before = currentPage()
		   WebElement target = links.find { it.getAttribute('href')?.contains('Page$Last') }

		   if (target == null) {
			   target = links.max { pageFromHref(it.getAttribute('href')) }
			   if (pageFromHref(target.getAttribute('href')) <= before) break   // already on last page
		   }

		   target.click()
		   waitPostBack()
		   if (currentPage() == before) break
	   }
	   KeywordUtil.logInfo("Customer list on page ${currentPage()}")
   }

   /** Reads the highest Kode Customer on the last page and returns the next one, keeping prefix and zero padding. */
   @Keyword
   String getNextKodeCustomer() {
	   goToLastPage()

	   List<String> kodes = findAll('tableCells', TABLE_CELLS, TIMEOUT)
			   .collect { it.getText().trim() }
			   .findAll { it ==~ KODE_REGEX }

	   if (kodes.isEmpty()) {
		   String first = GlobalVariable.defaultKode.toString()
		   KeywordUtil.markWarning("No existing Kode Customer found. Using default: ${first}")
		   return first
	   }

	   def last = kodes.collect { (it =~ KODE_REGEX)[0] }.max { it[2] as long }
	   String prefix = last[1]
	   String digits = last[2]
	   String next   = prefix + ((digits as long) + 1).toString().padLeft(digits.length(), '0')

	   KeywordUtil.logInfo("Last Kode Customer: ${last[0]} -> Next: ${next}")
	   return next
   }

   @Keyword
   def addCustomer(String kode, String nama, String alamat, String hp) {
	   openCustomerList()
	   WebUI.click(obj("btnTambahCustomer", BTN_TAMBAH_CUSTOMER))
	   WebUI.waitForElementVisible(obj("txtKode", TXT_KODE), TIMEOUT, FailureHandling.STOP_ON_FAILURE)

	   WebUI.setText(obj("txtKode", TXT_KODE), kode)
	   WebUI.setText(obj("txtNama", TXT_NAMA), nama)
	   WebUI.setText(obj("txtAlamat", TXT_ALAMAT), alamat)
	   WebUI.setText(obj("txtHp", TXT_HP), hp)
	   WebUI.takeScreenshot()

	   WebUI.click(obj("btnSimpan", BTN_SIMPAN))
	   if (WebUI.waitForAlert(2, FailureHandling.OPTIONAL)) {
		   KeywordUtil.logInfo("Alert: ${WebUI.getAlertText()}")
		   WebUI.acceptAlert()
	   }
	   waitPostBack()
   }

   /** Verifies the new row (Kode + Nama) exists on the last page of the Customer list. */
   @Keyword
   boolean verifyCustomerExists(String kode, String nama) {
	   goToLastPage()
	   String rowXpath = "//tr[td[normalize-space(.)='${kode}'] and td[normalize-space(.)='${nama}']]"
	   boolean found = WebUI.waitForElementPresent(obj("customerRow", rowXpath), 5, FailureHandling.OPTIONAL)
	   WebUI.takeScreenshot()

	   if (found) KeywordUtil.markPassed("Customer ${kode} - ${nama} found in list")
	   else KeywordUtil.markFailed("Customer ${kode} - ${nama} NOT found in list")
	   return found
   }

   @Keyword
   def logout() {
	   TestObject menuLogout = obj("menuLogout", MENU_LOGOUT)
	   if (WebUI.waitForElementClickable(menuLogout, 5, FailureHandling.OPTIONAL)) {
		   WebUI.click(menuLogout)
	   } else {
		   WebUI.navigateToUrl(baseUrl() + "/Logout.aspx")
	   }
	   WebUI.waitForPageLoad(TIMEOUT)

	   // Session must be invalidated: protected page redirects to login
	   WebUI.navigateToUrl(baseUrl() + "/AddCustomer.aspx")
	   if (WebUI.waitForElementVisible(obj("txtUsername", TXT_USERNAME), TIMEOUT, FailureHandling.OPTIONAL)) {
		   KeywordUtil.markPassed("Logout successful, session ended")
	   } else {
		   WebUI.takeScreenshot()
		   KeywordUtil.markFailed("Logout failed: protected page still accessible")
	   }
	   WebUI.closeBrowser()
   }
}