/*
 * Test Case : TC01_Login_Logout
 * Data      : Data Files/vitashopper_testdata.xlsx -> sheet "Login"
 * Flow      : Login -> verify -> Logout -> verify session ended
 */
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.vitashopper.ExcelHelper

Map<String, String> cred = ExcelHelper.loginCredential()

WebUI.comment("STEP 1: Login")
CustomKeywords.'com.vitashopper.CustomerManager.login'(cred['username'], cred['password'])
WebUI.takeScreenshot()

WebUI.comment("STEP 2: Logout")
CustomKeywords.'com.vitashopper.CustomerManager.logout'()
