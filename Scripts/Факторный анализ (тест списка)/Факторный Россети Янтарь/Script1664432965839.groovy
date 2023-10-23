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

'скролл Тюмень'
WebUI.scrollToElement(findTestObject('null'), 30)

'выбрать Россети Янтарь(ul)'
WebUI.click(findTestObject('null'))

'скролл Россети Янтарь'
WebUI.scrollToElement(findTestObject('null'), 30)

'выбрать Россети Янтарь(li)'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText = WebUI.getText(findTestObject('null'))

println(liText)

String noSpaceText = liText.replaceAll('\\s+', '')

println(noSpaceText)

WebUI.verifyEqual(noSpaceText, 'РоссетиЯнтарь')

if (WebUI.verifyEqual(noSpaceText, 'РоссетиЯнтарь') == false) {
    println('Текст в списке ' + liText //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

WebUI.closeBrowser()

