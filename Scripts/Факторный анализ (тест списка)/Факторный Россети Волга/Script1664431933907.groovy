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

'выбрать Россети Волга'
WebUI.click(findTestObject('null'))

'выбрать Мордовэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText = WebUI.getText(findTestObject('null'))

println(liText)

String noSpaceText = liText.replaceAll('\\s+', '')

println(noSpaceText)

WebUI.verifyEqual(noSpaceText, 'Мордовэнерго')

if (WebUI.verifyEqual(noSpaceText, 'Мордовэнерго') == false) {
    println('Текст в списке ' + liText //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Оренбургэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText1 = WebUI.getText(findTestObject('null'))

println(liText1)

String noSpaceText1 = liText1.replaceAll('\\s+', '')

println(noSpaceText1)

WebUI.verifyEqual(noSpaceText1, 'Оренбургэнерго')

if (WebUI.verifyEqual(noSpaceText1, 'Оренбургэнерго') == false) {
    println('Текст в списке ' + liText1 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'скролл Чувашэнерго'
WebUI.scrollToElement(findTestObject('null'), 30)

'выбрать Пензаэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText2 = WebUI.getText(findTestObject('null'))

println(liText2)

String noSpaceText2 = liText2.replaceAll('\\s+', '')

println(noSpaceText2)

WebUI.verifyEqual(noSpaceText2, 'Пензаэнерго')

if (WebUI.verifyEqual(noSpaceText2, 'Пензаэнерго') == false) {
    println('Текст в списке ' + liText2 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Самарские РС'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText3 = WebUI.getText(findTestObject('null'))

println(liText3)

String noSpaceText3 = liText3.replaceAll('\\s+', '')

println(noSpaceText3)

WebUI.verifyEqual(noSpaceText3, 'СамарскиеРС')

if (WebUI.verifyEqual(noSpaceText3, 'СамарскиеРС') == false) {
    println('Текст в списке ' + liText3 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Саратовские РС'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText4 = WebUI.getText(findTestObject('null'))

println(liText4)

String noSpaceText4 = liText4.replaceAll('\\s+', '')

println(noSpaceText4)

WebUI.verifyEqual(noSpaceText4, 'СаратовскиеРС')

if (WebUI.verifyEqual(noSpaceText4, 'СаратовскиеРС') == false) {
    println('Текст в списке ' + liText4 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Ульяновские РС'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText5 = WebUI.getText(findTestObject('null'))

println(liText5)

String noSpaceText5 = liText5.replaceAll('\\s+', '')

println(noSpaceText5)

WebUI.verifyEqual(noSpaceText5, 'УльяновскиеРС')

if (WebUI.verifyEqual(noSpaceText5, 'УльяновскиеРС') == false) {
    println('Текст в списке ' + liText5 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

'выбрать Чувашэнерго'
WebUI.click(findTestObject('null'))

WebUI.getText(findTestObject('null'))

String liText6 = WebUI.getText(findTestObject('null'))

println(liText6)

String noSpaceText6 = liText6.replaceAll('\\s+', '')

println(noSpaceText6)

WebUI.verifyEqual(noSpaceText6, 'Чувашэнерго')

if (WebUI.verifyEqual(noSpaceText6, 'Чувашэнерго') == false) {
    println('Текст в списке ' + liText6 //System.err.println('Не совпадение списка, должен быть Читаэнерго ' + liText)
        //System.exit(1)
        )
} else {
}

WebUI.closeBrowser()

