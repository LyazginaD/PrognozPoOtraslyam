import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys

WebUI.openBrowser('')

WebUI.navigateToUrl(findTestData('PlanFact').getValue(10, 4), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.setText(findTestObject('Страница авторизации/input__username'), findTestData('PlanFact').getValue(8, 1))

WebUI.setText(findTestObject('Страница авторизации/input__password'), findTestData('PlanFact').getValue(9, 1))

WebUI.click(findTestObject('Страница авторизации/button_'))

'выбрать фильтр ДЗО'
WebUI.click(findTestObject('null'))

'выбрать ПАО Россети'
WebUI.doubleClick(findTestObject('null'))

'выбрать "РаспредКомплекс"'
WebUI.click(findTestObject('null'))

'выбрать Россети Ленэнерго (ГК)'
WebUI.click(findTestObject('null'))

'скролл г.Санкт-Петербург'
WebUI.scrollToElement(findTestObject('null'), 30)

'выбрать Ленинградская область'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText = WebUI.getText(findTestObject('null'))

println(liText)

String noSpaceText = liText.replaceAll('\\s+', '')

println(noSpaceText)

WebUI.verifyEqual(noSpaceText, 'Ленинградскаяобласть')

if (WebUI.verifyEqual(noSpaceText, 'Ленинградскаяобласть') == false) {
	println('Текст в списке ' + liText //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
		//System.exit(1)
		)
} else {
}

'выбрать г. Санкт-Петербург'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText1 = WebUI.getText(findTestObject('null'))

println(liText1)

String noSpaceText1 = liText1.replaceAll('\\s+', '')

println(noSpaceText1)

WebUI.verifyEqual(noSpaceText1, 'г.Санкт-Петербург')

if (WebUI.verifyEqual(noSpaceText1, 'г.Санкт-Петербург') == false) {
    println('Текст в списке ' + liText1 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

WebUI.closeBrowser()


/*
 WebUI.getText(findTestObject(''))

String liText = WebUI.getText(findTestObject(''))

println(liText)

String noSpaceText = liText.replaceAll('\\s+', '')

println(noSpaceText)

WebUI.verifyEqual(noSpaceText, '')

if (WebUI.verifyEqual(noSpaceText, '') == false) {
    println('Текст в списке ' + liText //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}
 */
