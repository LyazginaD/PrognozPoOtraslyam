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

WebUI.delay(30)

if (WebUI.verifyElementText(findTestObject('Страница авторизации/button_'), 'Вход') == true) {
    WebUI.setText(findTestObject('Страница авторизации/input__username'), findTestData('PlanFact').getValue(8, 1))

    WebUI.setText(findTestObject('Страница авторизации/input__password'), findTestData('PlanFact').getValue(9, 1))

    WebUI.click(findTestObject('Страница авторизации/button_'))

    WebUI.delay(30)
} else {
    WebUI.refresh()

    WebUI.delay(30)

    if (WebUI.verifyElementText(findTestObject('Страница авторизации/button_'), 'Вход') == true) {
        WebUI.setText(findTestObject('Страница авторизации/input__username'), findTestData('PlanFact').getValue(8, 1))

        WebUI.setText(findTestObject('Страница авторизации/input__password'), findTestData('PlanFact').getValue(9, 1))

        WebUI.click(findTestObject('Страница авторизации/button_'))

        WebUI.delay(30)
    } else {
        WebUI.refresh()

        WebUI.delay(30)

        WebUI.setText(findTestObject('Страница авторизации/input__username'), findTestData('PlanFact').getValue(8, 1))

        WebUI.setText(findTestObject('Страница авторизации/input__password'), findTestData('PlanFact').getValue(9, 1))

        WebUI.click(findTestObject('Страница авторизации/button_'))
    }
}

WebUI.delay(30)

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

String absPerc = 'проц.'

DZOTest(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

WebUI.closeBrowser()


static def DZOTest(def absPerc, def year, def typeOfData, def month, def kpoOtpusk, def prognoz) {
	
	'Открыть фильтр "Организация"'
	WebUI.click(findTestObject('Прогноз по отраслям/Фильтр Организация'))

	'Нажать "Снять выделение"'
	WebUI.click(findTestObject('Прогноз по отраслям/Снять выделение в фильтре Организация'))

	WebUI.click(findTestObject('Прогноз по отраслям/Применить в фильтре Организация'))

	'Открыть фильтр "Организация"'
	WebUI.click(findTestObject('Прогноз по отраслям/Фильтр Организация'))

	'Нажать "Снять выделение"'
	WebUI.click(findTestObject('Прогноз по отраслям/Снять выделение в фильтре Организация'))

	WebUI.click(findTestObject('Прогноз по отраслям/ПАО Россети'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	WebUI.click(findTestObject('Прогноз по отраслям/Магистральные сети'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Магистральные сети'))

	WebUI.click(findTestObject('Прогноз по отраслям/Россети ФСК ЕЭС'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Магистральные сети'))

	WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети ФСК ЕЭС'))

	WebUI.click(findTestObject('Прогноз по отраслям/ИА ФСК ЕЭС'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Магистральные сети'))

	WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети ФСК ЕЭС'))

	WebUI.click(findTestObject('Прогноз по отраслям/Итого по ФСК ЕЭС'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	WebUI.click(findTestObject('Прогноз по отраслям/РаспредКомплекс'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	WebUI.click(findTestObject('Прогноз по отраслям/Тываэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Тываэнерго'))

	WebUI.click(findTestObject('Прогноз по отраслям/Тываэнерго нижний уровень'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	WebUI.click(findTestObject('Прогноз по отраслям/Чеченэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Чеченэнерго'))

	WebUI.click(findTestObject('Прогноз по отраслям/Чеченэнерго нижний уровень'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Volga()

	WebUI.click(findTestObject('Прогноз по отраслям/Россети Волга'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Volga()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Мордовэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Мордовэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Volga()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Оренбургэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Оренбургэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Volga()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Пензаэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Пензаэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Volga()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Самарские РС'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Самарские РС'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Volga()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Саратовские РС'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Саратовские РС'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Volga()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Ульяновские РС'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Ульяновские РС'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Volga()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Чувашэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Чувашэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Кубань'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Россети Кубань'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Кубань'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети Кубань'))

	WebUI.click(findTestObject('Прогноз по отраслям/Кубаньэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Len()

	WebUI.click(findTestObject('Прогноз по отраслям/Россети Ленэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Len()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Ленинградская область'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Ленинградская область'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Len()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Санкт-Петербург'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Санкт-Петербург'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	MosReg()

	WebUI.click(findTestObject('Прогноз по отраслям/Россети Московский регион'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	MosReg()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Московская область'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Московская область'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	MosReg()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Москва'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Москва'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevKav()

	WebUI.click(findTestObject('Прогноз по отраслям/Россети Северный Кавказ'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevKav()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Дагестанская сетевая компания'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Дагестанская сетевая компания'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevKav()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Дагестанская энергосбытовая компания'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Дагестанская энергосбытовая компания'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevKav()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Дагэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Дагэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevKav()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Ингушэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Ингушэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevKav()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Каббалкэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Каббалкэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevKav()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Карачаево-Черкесскэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Карачаево-Черкесскэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevKav()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Севкавказэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Севкавказэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevKav()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Ставропольэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Ставропольэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevZap()

	WebUI.click(findTestObject('Прогноз по отраслям/Россети Северо-Запад'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevZap()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Архангельский филиал'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Архангельский филиал'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevZap()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Вологодский филиал'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Вологодский филиал'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevZap()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Карельский филиал'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Карельский филиал'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevZap()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Мурманский филиал'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Мурманский филиал'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevZap()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Новгородский филиал'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Новгородский филиал'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevZap()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Псковский филиал'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Псковский филиал'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	SevZap()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/филиал в Республике Коми'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/филиал в Республике Коми'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Sibir()

	WebUI.click(findTestObject('Прогноз по отраслям/Россети Сибирь'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Sibir()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Алтайэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Алтайэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Sibir()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Бурятэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Бурятэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Sibir()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/ГАЭС'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/ГАЭС'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Sibir()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Красноярскэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Красноярскэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Sibir()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Кузбассэнерго-РЭС'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Кузбассэнерго-РЭС'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Sibir()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Омскэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Омскэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Sibir()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Хакасэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Хакасэнерго'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)

	Sibir()

	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Читаэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Читаэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Томск'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Россети Томск'))

	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Томск'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Россети Томск'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Томск'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети Томск'))
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Томская распределительная компания'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Томская распределительная компания'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Тюмень'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Россети Тюмень'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Тюмень'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети Тюмень'))
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Россети Тюмень нижний уровень'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Россети Тюмень нижний уровень'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Ural()
	
	WebUI.click(findTestObject('Прогноз по отраслям/Россети Урал'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Ural()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Пермэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Пермэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Ural()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Свердловэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Свердловэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Ural()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Челябэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Челябэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Centr()
	
	WebUI.click(findTestObject('Прогноз по отраслям/Россети Центр'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Centr()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Белгородэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Белгородэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Centr()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Брянскэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Брянскэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Centr()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Воронежэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Воронежэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Centr()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Костромаэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Костромаэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Centr()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Курскэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Курскэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Centr()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Липецкэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Липецкэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Centr()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Орелэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Орелэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Centr()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Смоленскэнерго'), 30)

	WebUI.click(findTestObject('Прогноз по отраслям/Смоленскэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Centr()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Тамбовэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Тамбовэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Centr()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Тверьэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Тверьэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Centr()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Ярэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Ярэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	CentrIP()
	
	WebUI.click(findTestObject('Прогноз по отраслям/Россети Центр и Приволжье'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	CentrIP()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Владимирэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Владимирэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	CentrIP()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Ивэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Ивэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	CentrIP()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Калугаэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Калугаэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	CentrIP()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Кировэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Кировэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	CentrIP()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Мариэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Мариэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	CentrIP()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Нижновэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Нижновэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	CentrIP()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Рязаньэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Рязаньэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	CentrIP()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Тулэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Тулэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	CentrIP()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Удмуртэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Удмуртэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Ug()
	
	WebUI.click(findTestObject('Прогноз по отраслям/Россети Юг'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Ug()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Астраханьэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Астраханьэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Ug()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Волгоградэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Волгоградэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Ug()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Калмэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Калмэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
	Ug()
	
	WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Ростовэнерго'), 30)
	
	WebUI.click(findTestObject('Прогноз по отраслям/Ростовэнерго'))
	
	Change(absPerc, year, typeOfData, month, kpoOtpusk, prognoz)
	
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

}


static def YearChange(def year, def prognoz) {
    'Раскрыть фильтр "Год"'
    WebUI.click(findTestObject('Прогноз по отраслям/Фильтр Год'))

    WebUI.click(findTestObject(prognoz + year), FailureHandling.CONTINUE_ON_FAILURE)
}

static def MonthChange(def month, def prognoz) {
    WebUI.deleteAllCookies()

    WebUI.click(findTestObject('Прогноз по отраслям/Фильтр Месяц'))

    WebUI.scrollToElement(findTestObject(prognoz + month), 30)

    WebUI.click(findTestObject(prognoz + month), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Заголовок дашборда'), 30)
}

static def TypeOfDataChange(def typeOfData) {
    WebUI.click(findTestObject('Прогноз по отраслям/Фильтр Тип показателя'))

    WebUI.click(findTestObject('Прогноз по отраслям/Тип показателя ' + typeOfData), FailureHandling.CONTINUE_ON_FAILURE)
}

static def KpoOtpusk(def kpoOtpusk, def prognoz) {
    'Раскрыть фильтр "КПО-Отпуск"'
    WebUI.click(findTestObject('Прогноз по отраслям/Фильтр КПО-Отпуск'))

    WebUI.click(findTestObject(prognoz + kpoOtpusk), FailureHandling.CONTINUE_ON_FAILURE)
}

static def Params(def absPerc, def year, def typeOfData, def month, def kpoOtpusk, def prognoz) {
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

    WebUI.refresh()

    WebUI.deleteAllCookies()

    WebUI.delay(15)

    if (WebUI.verifyElementPresent(findTestObject('Прогноз по отраслям/Заголовок дашборда'), 30) == false) {
        WebUI.refresh()

        WebUI.delay(15)

        if (WebUI.verifyElementPresent(findTestObject('Прогноз по отраслям/Заголовок дашборда'), 30) == false) {
            WebUI.refresh()

            WebUI.delay(15)
        }
    }
    
    YearChange(year, prognoz)

    'Открыть фильтр "ДЗО"'
    WebUI.click(findTestObject('Прогноз по отраслям/Фильтр Организация'))

    'Нажать "Снять выделение"'
    WebUI.click(findTestObject('Прогноз по отраслям/Снять выделение в фильтре Организация'))

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список ПАО Россети'))

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список ПАО Россети'))

    WebUI.click(findTestObject('Общие объекты/Раскрыть список РаспредКомплекс'))
}

static void Volga() {
    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Волга'), 30)

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети Волга'))
}

static void Len() {
    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Ленэнерго'), 30)

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети Ленэнерго'))
}

static void MosReg() {
    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Московский регион'), 30)

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети Московский регион'))
}

static void SevKav() {
    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Северный Кавказ'), 30)

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети Северный Кавказ'))
}

static void SevZap() {
    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Северо-Запад'), 30)

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети Северо-Запад'))
}

static void Sibir() {
    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Сибирь'), 30)

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети Сибирь'))
}

static void Ural() {
    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Урал'), 30)

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети Урал'))
}

static void Centr() {
    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Центр'), 30)

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети Центр'))
}

static void CentrIP() {
    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Центр и Приволжье'), 30)

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети Центр и Приволжье'))
}

static void Ug() {
    WebUI.scrollToElement(findTestObject('Прогноз по отраслям/Раскрыть список Россети Юг'), 30)

    WebUI.click(findTestObject('Прогноз по отраслям/Раскрыть список Россети Юг'))
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

