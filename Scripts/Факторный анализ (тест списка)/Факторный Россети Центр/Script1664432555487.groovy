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

'выбрать Россети Центр'
WebUI.click(findTestObject('null'))

'выбрать Белгородэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText10 = WebUI.getText(findTestObject('null'))

println(liText10)

String noSpaceText10 = liText10.replaceAll('\\s+', '')

println(noSpaceText10)

WebUI.verifyEqual(noSpaceText10, 'Белгородэнерго')

if (WebUI.verifyEqual(noSpaceText10, 'Белгородэнерго') == false) {
    println('Текст в списке ' + liText10 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Брянскэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText = WebUI.getText(findTestObject('null'))

println(liText)

String noSpaceText = liText.replaceAll('\\s+', '')

println(noSpaceText)

WebUI.verifyEqual(noSpaceText, 'Брянскэнерго')

if (WebUI.verifyEqual(noSpaceText, 'Брянскэнерго') == false) {
    println('Текст в списке ' + liText //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Воронежэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText1 = WebUI.getText(findTestObject('null'))

println(liText1)

String noSpaceText1 = liText1.replaceAll('\\s+', '')

println(noSpaceText1)

WebUI.verifyEqual(noSpaceText1, 'Воронежэнерго')

if (WebUI.verifyEqual(noSpaceText1, 'Воронежэнерго') == false) {
    println('Текст в списке ' + liText1 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'скролл Костромаэнерго'
WebUI.scrollToElement(findTestObject('null'), 30)

'выбрать Костромаэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText2 = WebUI.getText(findTestObject('null'))

println(liText2)

String noSpaceText2 = liText2.replaceAll('\\s+', '')

println(noSpaceText2)

WebUI.verifyEqual(noSpaceText2, 'Костромаэнерго')

if (WebUI.verifyEqual(noSpaceText2, 'Костромаэнерго') == false) {
    println('Текст в списке ' + liText2 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Курскэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText3 = WebUI.getText(findTestObject('null'))

println(liText3)

String noSpaceText3 = liText3.replaceAll('\\s+', '')

println(noSpaceText3)

WebUI.verifyEqual(noSpaceText3, 'Курскэнерго')

if (WebUI.verifyEqual(noSpaceText3, 'Курскэнерго') == false) {
    println('Текст в списке ' + liText3 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Липецкэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText4 = WebUI.getText(findTestObject('null'))

println(liText4)

String noSpaceText4 = liText4.replaceAll('\\s+', '')

println(noSpaceText4)

WebUI.verifyEqual(noSpaceText4, 'Липецкэнерго')

if (WebUI.verifyEqual(noSpaceText4, 'Липецкэнерго') == false) {
    println('Текст в списке ' + liText4 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Орелэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText5 = WebUI.getText(findTestObject('null'))

println(liText5)

String noSpaceText5 = liText5.replaceAll('\\s+', '')

println(noSpaceText5)

WebUI.verifyEqual(noSpaceText5, 'Орелэнерго')

if (WebUI.verifyEqual(noSpaceText5, 'Орелэнерго') == false) {
    println('Текст в списке ' + liText5 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Смоленскэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText6 = WebUI.getText(findTestObject('null'))

println(liText6)

String noSpaceText6 = liText6.replaceAll('\\s+', '')

println(noSpaceText6)

WebUI.verifyEqual(noSpaceText6, 'Смоленскэнерго')

if (WebUI.verifyEqual(noSpaceText6, 'Смоленскэнерго') == false) {
    println('Текст в списке ' + liText6 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Тамбовэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText7 = WebUI.getText(findTestObject('null'))

println(liText7)

String noSpaceText7 = liText7.replaceAll('\\s+', '')

println(noSpaceText7)

WebUI.verifyEqual(noSpaceText7, 'Тамбовэнерго')

if (WebUI.verifyEqual(noSpaceText7, 'Тамбовэнерго') == false) {
    println('Текст в списке ' + liText7 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'скролл Тверьэнерго'
WebUI.scrollToElement(findTestObject('null'), 30)

'выбрать Тверьэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText8 = WebUI.getText(findTestObject('null'))

println(liText8)

String noSpaceText8 = liText8.replaceAll('\\s+', '')

println(noSpaceText8)

WebUI.verifyEqual(noSpaceText8, 'Тверьэнерго')

if (WebUI.verifyEqual(noSpaceText8, 'Тверьэнерго') == false) {
    println('Текст в списке ' + liText8 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Ярэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText9 = WebUI.getText(findTestObject('null'))

println(liText9)

String noSpaceText9 = liText9.replaceAll('\\s+', '')

println(noSpaceText9)

WebUI.verifyEqual(noSpaceText9, 'Ярэнерго')

if (WebUI.verifyEqual(noSpaceText9, 'Ярэнерго') == false) {
    println('Текст в списке ' + liText9 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

WebUI.closeBrowser()

