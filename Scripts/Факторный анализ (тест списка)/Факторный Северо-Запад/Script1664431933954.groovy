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

'скролл Россети Томск'
WebUI.scrollToElement(findTestObject('null'), 30)

'выбрать Россети Северо-Запад'
WebUI.click(findTestObject('null'))

'выбрать Архангельский филиал'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText = WebUI.getText(findTestObject('null'))

println(liText)

String noSpaceText = liText.replaceAll('\\s+', '')

println(noSpaceText)

WebUI.verifyEqual(noSpaceText, 'Архангельскийфилиал')

if (WebUI.verifyEqual(noSpaceText, 'Архангельскийфилиал') == false) {
    println('Текст в списке ' + liText //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Вологодский филиал'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText1 = WebUI.getText(findTestObject('null'))

println(liText1)

String noSpaceText1 = liText1.replaceAll('\\s+', '')

println(noSpaceText1)

WebUI.verifyEqual(noSpaceText1, 'Вологодскийфилиал')

if (WebUI.verifyEqual(noSpaceText1, 'Вологодскийфилиал') == false) {
    println('Текст в списке ' + liText1 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Карельский филиал'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText2 = WebUI.getText(findTestObject('null'))

println(liText2)

String noSpaceText2 = liText2.replaceAll('\\s+', '')

println(noSpaceText2)

WebUI.verifyEqual(noSpaceText2, 'Карельскийфилиал')

if (WebUI.verifyEqual(noSpaceText2, 'Карельскийфилиал') == false) {
    println('Текст в списке ' + liText2 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Колэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText3 = WebUI.getText(findTestObject('null'))

println(liText3)

String noSpaceText3 = liText3.replaceAll('\\s+', '')

println(noSpaceText3)

WebUI.verifyEqual(noSpaceText3, 'Колэнерго')

if (WebUI.verifyEqual(noSpaceText3, 'Колэнерго') == false) {
    println('Текст в списке ' + liText3 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'скролл Колэнерго'
WebUI.scrollToElement(findTestObject('null'), 30)

'выбрать Новгородский филиал'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText4 = WebUI.getText(findTestObject('null'))

println(liText4)

String noSpaceText4 = liText4.replaceAll('\\s+', '')

println(noSpaceText)

WebUI.verifyEqual(noSpaceText4, 'Новгородскийфилиал')

if (WebUI.verifyEqual(noSpaceText4, 'Новгородскийфилиал') == false) {
    println('Текст в списке ' + liText4 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Псковский филиал'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText5 = WebUI.getText(findTestObject('null'))

println(liText5)

String noSpaceText5 = liText5.replaceAll('\\s+', '')

println(noSpaceText5)

WebUI.verifyEqual(noSpaceText5, 'Псковскийфилиал')

if (WebUI.verifyEqual(noSpaceText5, 'Псковскийфилиал') == false) {
    println('Текст в списке ' + liText5 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать филиал в Республике Коми'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText6 = WebUI.getText(findTestObject('null'))

println(liText6)

String noSpaceText6 = liText6.replaceAll('\\s+', '')

println(noSpaceText6)

WebUI.verifyEqual(noSpaceText6, 'филиалвРеспубликеКоми')

if (WebUI.verifyEqual(noSpaceText6, 'филиалвРеспубликеКоми') == false) {
    println('Текст в списке ' + liText6 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

WebUI.closeBrowser()

