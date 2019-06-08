VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ОтчетПоКассе"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'/-----------------------------------------
'/ vmm, 2017
'/ Документы.ОтчетОРозничныхПродажах
'/
'/ Из-за постоянных апгрейдов в 1С
'/ возможны изменения полей
'/-----------------------------------------
Private Sub Add_Click()
  Dim rst As DAO.Recordset
  Dim refs As Object
  Dim selDoc As Object
  Dim newDoc As Object
  Dim selSkl As Object
  Dim selVal As Object
  Dim newGoods As Object
  Dim selEI As Object
  Dim valNDS As Object
  Dim selNom As Object
  Dim SelTypPriceR As Object
  Dim newNom As Object
  Dim N, СуммаНДС
  Dim sh As Variant
  
  If IsNull(Me!NomerDoc) Then
     MsgBox "Нет номера накладной !", vbInformation, "ВНИМАНИЕ"
     Exit Sub
  End If
  If rez.METADATA.Name = "БухгалтерияПредприятия" Then N = 0
  If Err = 0 And rez.TransactionActive = False Then
     On Error GoTo ex_rd
     
     Set SelTypPriceR = rez.Справочники.ТипыЦенНоменклатуры.НайтиПоНаименованию("розничная")
     If Me.НДС = 18 Then
        Set valNDS = rez.Перечисления.СтавкиНДС.НДС18
     Else
        Set valNDS = rez.Перечисления.СтавкиНДС.БезНДС
     End If
     Set selEI = rez.Справочники.КлассификаторЕдиницИзмерения.НайтиПоНаименованию("шт")
     Set selVal = rez.Справочники.Валюты.НайтиПоНаименованию("руб")
     Set selSkl = rez.Справочники.Склады.НайтиПоНаименованию("Основной склад")
     sh = Me!Shops
     Set refs = rez.Документы.ОтчетОРозничныхПродажах
     Set selDoc = refs.НайтиПоНомеру(CStr(Me!NomerDoc), CDate(Me!DataDoc))
     If selDoc.Пустая() Then
        If MsgBox("Добавить накладную в 1С ?", vbYesNo, "") = vbNo Then Exit Sub
        On Error GoTo ex_wr
        rez.BeginTransaction
        Set newDoc = refs.CreateDocument()
        newDoc.Номер = CStr(Me!NomerDoc)
        newDoc.Дата = CDate(Me!DataDoc)
        'If Me!НУ = True Then newDoc.ОтражатьВНалоговомУчете = True Else newDoc.ОтражатьВНалоговомУчете = False
        newDoc.ВалютаДокумента = selVal
        newDoc.Комментарий = sh
        newDoc.Организация = rez.Справочники.Организации.НайтиПоНаименованию(CStr(Me!Организация))
        newDoc.Склад = selSkl               '"Основной склад"
        newDoc.СчетКасса = rez.ПланыСчетов.Хозрасчетный.НайтиПоКоду("50.01")
        newDoc.Выручка = Me!Сумма
        newDoc.Ответственный = rez.Справочники.Пользователи.НайтиПоНаименованию("Не авторизован")
     Else
        If MsgBox("Номер " & CStr(Me!NomerDoc) & " входного документа от подразделения " & CStr(Me!Shops) & " уже существует в базе 1С, продолжить?", vbYesNo, "ВНИМАНИЕ") = vbNo Then Exit Sub
        On Error GoTo ex_wr
        rez.BeginTransaction
        Set newDoc = selDoc.ПолучитьОбъект()
        For Each newGoods In newDoc.Товары
            If newGoods.Номенклатура.Наименование <> "" Then
               MsgBox "В этом документе уже есть запчасти, добавить больше невозможно!", vbInformation, "ВНИМАНИЕ"
               'rez.RollbackTransaction
               GoTo ex_ex_wr:
            End If
        Next
     End If
     '/-----------------------------------------
     '/ Добавить
     '/-----------------------------------------
     Set rst = CurrentDb.OpenRecordset("SELECT * FROM tbList")
     If rst.RecordCount = 0 Then
        MsgBox "Список пуст, добавить нечего!", vbInformation, "ВНИМАНИЕ"
        'rez.RollbackTransaction
        GoTo ex_ex_wr:
     Else
        newDoc.ТипЦен = SelTypPriceR
        newDoc.СуммаВключаетНДС = True
        newDoc.УдалитьУчитыватьНДС = False
        'newDoc.СуммаДокумента = Me!Сумма
        newDoc.КратностьДокумента = 1
        newDoc.КурсДокумента = 1
        Do While Not rst.EOF
           If rst!Есть_1С = True Then
              Set selNom = rez.Справочники.Номенклатура.НайтиПоНаименованию(CStr(rst!Номенклатура_1C))
              If selNom.Пустая() Then
                 MsgBox "В справочнике Номенклатура не найден код " & rst!Номенклатура_1C, vbInformation, "ВНИМАНИЕ"
                 GoTo ex_ex_wr:
              Else
                 '/ Новая создается товарная накладная
                 Set newNom = selNom.ПолучитьОбъект()
                 Set newGoods = newDoc.Товары.Добавить()
                 
                 newGoods.Номенклатура = newNom.ссылка  '.Номенклатура
                 'newGoods.ЕдиницаИзмерения = selEI     'newNom.ЕдиницаХраненияОстатков
                 'newGoods.Коэффициент = 1
                 newGoods.Количество = CLng(rst!Кол)
                 newGoods.Цена = Format(rst!Цена, "0.00")
                 newGoods.Сумма = Format(rst!Сумма, "0.00")
                 newGoods.СчетУчета = rez.ПланыСчетов.Хозрасчетный.НайтиПоКоду("41.01")
                 newGoods.СчетУчетаНДСПоРеализации = rez.ПланыСчетов.Хозрасчетный.НайтиПоКоду("90.03")
                 newGoods.Субконто = rez.Справочники.НоменклатурныеГруппы.НайтиПоНаименованию("Розничная торговля автомобильными деталями")
                 
                 If Me.НДС = 18 Then
                    newGoods.СтавкаНДС = valNDS
                    newGoods.СуммаНДС = Format(rst!Сумма * 18 / 118, "0.00")
                    newGoods.СчетДоходов = rez.ПланыСчетов.Хозрасчетный.НайтиПоКоду("90.01.1")
                    newGoods.СчетРасходов = rez.ПланыСчетов.Хозрасчетный.НайтиПоКоду("90.02.1")
                 Else
                    newGoods.СтавкаНДС = valNDS
                    newGoods.СчетДоходов = rez.ПланыСчетов.Хозрасчетный.НайтиПоКоду("90.01.2")
                    newGoods.СчетРасходов = rez.ПланыСчетов.Хозрасчетный.НайтиПоКоду("90.02.2")
                    'newGoods.СуммаНДС = ""  'Format(rst!Сумма, "0.00")
                 End If
                 
              End If
           End If
           rst.MoveNext
        Loop
        '/-----------------------------------------
        '/ Записать
        '/-----------------------------------------
        'If newDoc.ПроверитьЗаполнение Then
           On Error Resume Next
           If rez.ПараметрыСеанса.ВыполняетсяОбновлениеИБ = False Then Debug.Print Err
           If Err <> 0 Then rez.ПараметрыСеанса.ВыполняетсяОбновлениеИБ = False: Err.Clear
           If rez.ПараметрыСеанса.ПропуститьПроверкуЗапретаИзменения = False Then Debug.Print Err
           If Err <> 0 Then rez.ПараметрыСеанса.ПропуститьПроверкуЗапретаИзменения = False: Err.Clear
           On Error GoTo ex_wr
           newDoc.Записать (rez.РежимЗаписиДокумента.Запись)
           newDoc.Записать (rez.РежимЗаписиДокумента.Проведение)
           rez.CommitTransaction
           MsgBox "Накладная № " & Me!NomerDoc & " добавлена в 1С !", vbInformation, "ВНИМАНИЕ"
        'Else
        '   GoTo ex_wr
        'End If
     End If
  Else
     MsgBox "Подключитесь к 1С !", vbInformation, "ВНИМАНИЕ"
     Exit Sub
  End If
ex_ex_wr:
  Set refs = Nothing
  Set selDoc = Nothing
  Set newDoc = Nothing
  Set selSkl = Nothing
  Set selVal = Nothing
  Set newGoods = Nothing
  Set selEI = Nothing
  Set valNDS = Nothing
  Set selNom = Nothing
  Set SelTypPriceR = Nothing
  Set newNom = Nothing
  Set rst = Nothing
  Exit Sub
ex_wr:
  Set refs = Nothing
  Set selDoc = Nothing
  Set newDoc = Nothing
  Set selSkl = Nothing
  Set selVal = Nothing
  Set newGoods = Nothing
  Set selEI = Nothing
  Set valNDS = Nothing
  Set selNom = Nothing
  Set SelTypPriceR = Nothing
  Set newNom = Nothing
  Set rst = Nothing
  rez.RollbackTransaction
ex_rd:
  MsgBox "Произошла ошибка, повторите снова!" & vbCrLf & Err.Description, vbInformation, "ВНИМАНИЕ"
End Sub
'/---------------------------------- Fin

