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
import com.kms.katalon.keyword.excel.ExcelKeywords as ExcelKeywords

WebUI.openBrowser('')

WebUI.navigateToUrl(findTestData('PlanFact').getValue(10, 6))

WebUI.setText(findTestObject('Страница авторизации/input__username'), findTestData('PlanFact').getValue(8, 1))

WebUI.setText(findTestObject('Страница авторизации/input__password'), findTestData('PlanFact').getValue(9, 1))

WebUI.click(findTestObject('Страница авторизации/button_'))

'Открыть фильтр "Организация"'
WebUI.click(findTestObject('Прогноз по отраслям/Фильтр Организация'))

'Нажать "Снять выделение"'
WebUI.click(findTestObject('Прогноз по отраслям/Снять выделение в фильтре Организация'))

WebUI.click(findTestObject('Прогноз по отраслям/Применить в фильтре Организация'))

int year = 2020

String prognoz = 'Прогноз по отраслям/'

YearChange(year, prognoz)

String month

String typeOfData

String kpoOtpusk

String absPerc = 'абс.'

WebUI.click(findTestObject('Прогноз по отраслям/Переключить отображение с процентов на абс'))

def dzotest = DZOTest(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

WebUI.closeBrowser()

static def DZOTest(def absPerc, def year, def typeOfData, def month, def kpoOtpusk, def prognoz) {
    'Открыть фильтр "ДЗО"'
    WebUI.click(findTestObject('Прогноз по отраслям/Фильтр Организация'))

    'Нажать "Снять выделение"'
    WebUI.click(findTestObject('Прогноз по отраслям/Снять выделение в фильтре Организация'))

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список ПАО Россети'))

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список ПАО Россети'))

    WebUI.click(findTestObject('Общие объекты/Раскрыть список РаспредКомплекс'))

    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Янтарь'), 30)

    WebUI.click(findTestObject('Прогноз по отраслям/Россети Янтарь'))

    Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Янтарь'), 30)

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети Янтарь'))

    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Россети Янтарь нижний уровень'), 30)

    WebUI.click(findTestObject('Прогноз по отраслям/Россети Янтарь нижний уровень'))

    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Заголовок фильтра Организация'), 30)

    'Нажать "Применить"'
    WebUI.click(findTestObject('Прогноз по отраслям/Применить в фильтре Организация'))

    Params(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
}


static def YearChange(def year, def prognoz) {
   
}

static def MonthChange(def month, def prognoz) {
 
}

static def TypeOfDataChange(def typeOfData) {

}

static def KpoOtpusk(def kpoOtpusk, def prognoz) {

}

static def Params(def absPerc, def year, def typeOfData, def month, def kpoOtpusk, def prognoz) {
	month = 'Январь'

	MonthChange(month, prognoz)

}

static def ParamsChangePFKO(def absPerc, def year, def typeOfData, def month, def kpoOtpusk, def prognoz) {
	typeOfData = 'Факт'

	TypeOfDataChange(typeOfData)

	kpoOtpusk = 'КПО'

	KpoOtpusk(kpoOtpusk, prognoz)

	Test(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	kpoOtpusk = 'Отпуск'

	KpoOtpusk(kpoOtpusk, prognoz)

	Test(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	typeOfData = 'План'

	TypeOfDataChange(typeOfData)

	Test(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	kpoOtpusk = 'КПО'

	KpoOtpusk(kpoOtpusk, prognoz)

	Test(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	typeOfData = 'План СУ'

	TypeOfDataChange(typeOfData)

	Test(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	kpoOtpusk = 'Отпуск'

	KpoOtpusk(kpoOtpusk, prognoz)

	Test(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
}

static def Change(def absPerc, def year, def typeOfData, def month, def kpoOtpusk, def prognoz) {
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Заголовок дашборда'), 30)

	'Нажать "Применить"'
	WebUI.click(findTestObject('Прогноз по отраслям/Применить в фильтре Организация'))

	Params(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	'Открыть фильтр "ДЗО"'
	WebUI.click(findTestObject('Прогноз по отраслям/Фильтр Организация'))

	'Нажать "Снять выделение"'
	WebUI.click(findTestObject('Прогноз по отраслям/Снять выделение в фильтре Организация'))
}

static def Test(def absPerc, def year, def typeOfData, def month, def kpoOtpusk, def prognoz) {

}

static def WriteToExcel(def absPerc, def year, def typeOfData, def month, def kpoOtpusk, def prognoz) {
	String sheetName = 'Sheet1'

	def workbook01 = ExcelKeywords.getWorkbook(GlobalVariable.excelFilePath)

	def data = findTestData('PlanFact')

	int n = data.getRowNumbers() + 1

	String dZO = WebUI.getText(findTestObject('Прогноз по отраслям/Фильтр Организация'))

	String dashboardName = 'Прогноз по отраслям'

	def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)

	ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

	ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

	ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

	ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, month)

	ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, typeOfData)

	ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, kpoOtpusk)

	ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, absPerc)

	n = (n + 1)

	ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)
}

