VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_nakladnayaTo_1C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'/ -----------------------------------------------------------------
'/ vmm, 2018
'/ Форма nakladnayaTo_1C
'/
'/ Вставка прих. накладной из корпоративной системы предприятия в 1С 
'/ в Документы.ПоступлениеТоваровУслуг
'/
'/ Из-за постоянных апгрейдов в 1С
'/ возможны изменения полей
'/ -----------------------------------------------------------------
Private Sub Add_Click()
  Dim rst As DAO.Recordset
  Dim refs As Object
  Dim newSel As Object
  Dim newDoc As Object
  Dim newNom As Object
  Dim newGoods As Object
  Dim valVidOper As Object
  Dim selDoc As Object
  Dim selDog As Object
  Dim selSkl As Object
  Dim selNom As Object
  Dim selVusn As Object
  Dim valAvans As Object
  Dim selEI As Object
  Dim PLAccBU As Object
  Dim PLAccNDS As Object
  Dim selClient As Object
  Dim valNDS As Object
  Dim regNom As Object
  Dim Admin As Object
  Dim DocPriceZ As Object
  Dim DocPriceR As Object
  Dim LinesPriceZ As Object
  Dim LinesPriceR As Object
  Dim SelTypPriceZ As Object
  Dim SelTypPriceR As Object
  Dim selVal As Object
  Dim selAcc As Object
  Dim regPrice As Object
  Dim selDocPTR As Object
  Dim DataNewDoc As Date
  Dim DataNewGds As Date
  Dim N, sel, Discount, sek
  Dim pr As Double
  
  On Error Resume Next
  'If rez.METADATA.Name = "БухгалтерияПредприятия" Then N = 0
  If Err = 0 And Not rez Is Nothing Then
     If rez.TransactionActive = True Then
        MsgBox "Блокировка базы данных, принять накладную в данный момент невозможно.", vbInformation, "ВНИМАНИЕ"
        rez.RollbackTransaction
        Exit Sub
     End If
     On Error GoTo ex_rd
     Me.Add.ForeColor = 255
     If IsNull(Me.Nomer) Or Me.Nomer = "" Or IsNull(Me.DataNakl) Or Me.DataNakl = "" Then MsgBox "Нет номера накладной!", vbInformation, "ВНИМАНИЕ": Exit Sub
     If Me.Child2.Form.НомНакл <> Me.Child0.Form.НомНакл Then MsgBox "Спецификация не от этой накладной!", vbInformation, "ВНИМАНИЕ": Exit Sub
     If DCount("[Код]", "Номенклатура_Z", "[Yes]=true") <> 0 Then MsgBox "Не все позиции Номенклатуры отредактированы!", vbInformation, "ВНИМАНИЕ": Exit Sub
     If DCount("[Номенклатура]", "ПодчинПТУ_Z", "[Номенклатура_1C]=''") <> 0 Then MsgBox "Заполните Номенклатуру!", vbInformation, "ВНИМАНИЕ": Exit Sub
     If Me.NomDoc = "" And Me.DataDoc = "" Then MsgBox "Номер или дату входного документа поставщика!", vbInformation, "ВНИМАНИЕ": Exit Sub
     DataNewDoc = Me!DataNakl
     DataNewGds = Me!DataDoc
     If IsNull(Me!Наценка) Then Discount = 1.59 Else Discount = CDbl(Me!Наценка) 'SelTypPriceR.ПроцентСкидкиНаценки + 100) / 100
     If rez.Документы.ПоступлениеТоваровУслуг.НайтиПоРеквизиту("НомерВходящегоДокумента", CStr(Me!NomDoc)).Контрагент.Наименование = CStr(Me!Delivery) Then
        MsgBox "Номер " & CStr(Me!NomDoc) & " входного документа от поставщика " & CStr(Me!Delivery) & " уже существует в базе 1С !", vbInformation, "ВНИМАНИЕ"
        Exit Sub
     End If
     If MsgBox("Добавить накладную в 1С:Зудилин К.Н.?", vbYesNo, "") = vbYes Then
        On Error GoTo ex_wr
        rez.BeginTransaction
        '/ -----------------------------------------------------------                    
        '/ Справочники
        '/ -----------------------------------------------------------                    
        'Set Admin = rez.Справочники.Пользователи.НайтиПоНаименованию("Администратор")
        Set selVal = rez.Справочники.Валюты.НайтиПоНаименованию("руб")
        Set selSkl = rez.Справочники.Склады.НайтиПоНаименованию("Основной склад")   '"000000001"
        Set selClient = rez.Справочники.Контрагенты.НайтиПоНаименованию(CStr(Me!Delivery))
        Set selEI = rez.Справочники.КлассификаторЕдиницИзмерения.НайтиПоНаименованию("шт")
        If selClient.Пустая Then MsgBox "Нет  такого поставщика " & CStr(Me!Delivery), vbInformation, "ВНИМАНИЕ": Exit Sub
        'If Me.Account = "" Or IsNull(Me.Account) Then
        '   If MsgBox("Нет номера счета оплаты поставщику, продолжить?", vbYesNo, "ВНИМАНИЕ") = vbNo Then GoTo ex_ex_rd
        'Else
        '   Set sel = rez.Документы.СчетНаОплатуПоставщика.Выбрать(, selClient)
        '   Do While selAcc.Next
        '      If selAcc.Code = Me.Account Then Exit Do
        '   Loop
        'End If
        'If Me.contract = "" Or IsNull(Me.contract) Then
        '   If MsgBox("Нет номера договора у этого поставщика, продолжить?", vbYesNo, "ВНИМАНИЕ") = vbNo Then GoTo ex_ex_rd
        'Else
        Set selDog = rez.Справочники.ДоговорыКонтрагентов.НайтиПоКоду(CStr(Me.Contract), , , selClient)
        If selDog.Пустая() Then MsgBox "Нет номера договора у этого поставщика!", vbInformation, "ВНИМАНИЕ": Exit Sub
        'End If
        Set PLAccBU = rez.ПланыСчетов.Хозрасчетный.НайтиПоКоду("41.01")
        Set PLAccNDS = rez.ПланыСчетов.Хозрасчетный.НайтиПоКоду("19.03")
        Set SelTypPriceZ = rez.Справочники.ТипыЦенНоменклатуры.НайтиПоНаименованию("закупочная")
        Set SelTypPriceR = rez.Справочники.ТипыЦенНоменклатуры.НайтиПоНаименованию("розничная")
        '/ Перечисления
        Set valVidOper = rez.Перечисления.ВидыОперацийПоступлениеТоваровУслуг.ПокупкаКомиссия
        Set valNDS = rez.Перечисления.СтавкиНДС.НДС18
        Set valAvans = rez.Перечисления.СпособыЗачетаАвансов.НеЗачитывать
        'Set selVusn = rez.Перечисления.ОтражениеВУСН.Принимаются
        '/ СОЗДАНИЕ ДОКУМЕНТА
        Set refs = rez.Документы.ПоступлениеТоваровУслуг
        Set selDoc = refs.НайтиПоНомеру(CStr(Me!Nomer), Date)
        'Set selDoc = refs.НайтиПоНомеру(CStr("АВМ00000004"))
        If Not selDoc.Пустая Then
           'Set sel = selDoc.ПолучитьОбъект()
           'If selDoc.Товары.Количество() = 0 Then
           'Else
           'End If
           rez.RollbackTransaction
           MsgBox "В 1С:Зудилин К.Н. накладная поставщика с номером " & CStr(Me!Nomer) & " уже существует!"
           GoTo ex_ex_wr
        Else
           Set newDoc = refs.CreateDocument()
           newDoc.Номер = CStr(Me!Nomer)
           newDoc.Дата = DataNewDoc            'Me!DataNakl
           newDoc.Организация = rez.Справочники.Организации.НайтиПоНаименованию(CStr(Me.Organization))
           newDoc.Контрагент = selClient
           newDoc.СпособЗачетаАвансов = valAvans
           'newDoc.Сделка = selAcc             ' Счет на оплату поставщика
           'If Me!НУ = True Then newDoc.ОтражатьВНалоговомУчете = True Else newDoc.ОтражатьВНалоговомУчете = False
           newDoc.Склад = selSkl               '"Основной склад"
           newDoc.ВидОперации = valVidOper     '"Покупка, комиссия"
           If (Not selDog.Пустая()) And (Not selDog.ПометкаУдаления) And (Not selDog.ВалютаВзаиморасчетов.Пустая()) Then
              newDoc.ДоговорКонтрагента = selDog
              newDoc.ВалютаДокумента = selDog.ВалютаВзаиморасчетов
              newDoc.КратностьВзаиморасчетов = 1
              newDoc.КурсВзаиморасчетов = 1
           Else
              rez.RollbackTransaction
              MsgBox "Либо нет номера договора, либо стоит пометка удаления на этой накладной!", vbInformation, "ВНИМАНИЕ"
              GoTo ex_ex_wr
           End If
           'newDoc.Комментарий = "Выгрузка накладной из основной базы"
           '/ Вкладка Учета расчетов
           newDoc.СчетУчетаРасчетовСКонтрагентом = rez.ПланыСчетов.Хозрасчетный.НайтиПоКоду("60.01")
           newDoc.СчетУчетаРасчетовПоАвансам = rez.ПланыСчетов.Хозрасчетный.НайтиПоКоду("60.02")
           '/ Вкладка Дополнительно
           newDoc.Грузоотправитель = selClient
           newDoc.Ответственный = Admin   '"Администратор"
           newDoc.НомерВходящегоДокумента = CStr(Me!NomDoc)
           newDoc.ДатаВходящегоДокумента = DataNewGds
           '------------------------------------------------------------
           '/ Ссылки на документы
           '/ Установка цен номенклатуры
           '/ -----------------------------------------------------------                                                         
           newDoc.СуммаВключаетНДС = False
           newDoc.УдалитьУчитыватьНДС = False
           newDoc.НДСВключенВСтоимость = False
           'newDoc.РегистрироватьЦеныПоставщика = True
           newDoc.ТипЦен = SelTypPriceZ
           '/ -----------------------------------------------------------                                                         
           '/  Создание закупочной цены
           '/ -----------------------------------------------------------                                                         
           If Me.flPrice = True Then
              Set DocPriceZ = rez.Документы.УстановкаЦенНоменклатуры.СоздатьДокумент()
              DocPriceZ.Дата = DataNewDoc
              'DocPriceZ.Ответственный = Admin
              DocPriceZ.НеПроводитьНулевыеЗначения = True
              DocPriceZ.Информация = "Закупочная"
              DocPriceZ.ТипЦен = SelTypPriceZ
              '/ -----------------------------------------------------------                                                          
              '/  Создание розничной цены
              '/ -----------------------------------------------------------                                                          
              Set DocPriceR = rez.Документы.УстановкаЦенНоменклатуры.СоздатьДокумент()
              DocPriceR.Дата = DataNewDoc
              'DocPriceR.Ответственный = Admin
              DocPriceR.НеПроводитьНулевыеЗначения = False
              DocPriceR.Информация = "Розничная"
              DocPriceR.ТипЦен = SelTypPriceR
           End If
           'DocPrice.Комментарий = ""
           '/ -----------------------------------------------------------                                                             
           '/ Вкладка Товар
           '/ -----------------------------------------------------------                                                             
           Set rst = CurrentDb.OpenRecordset("SELECT * FROM ПодчинПТУ_Z ORDER BY N")
           If rst.RecordCount = 0 Then
              rez.RollbackTransaction
              MsgBox "В справочнике Номенклатура не найден код " & CStr(rst!Код), vbInformation, "ВНИМАНИЕ"
              GoTo ex_ex_wr
           Else
              '/ -----------------------------------------------------------                                                              
              '/ Добавить Товары
             '/ -----------------------------------------------------------                                                               
              Do While Not rst.EOF
                 Set selNom = rez.Справочники.Номенклатура.НайтиПоКоду(CStr(rst!Код))
                 If selNom.Пустая() Then
                    '/ Новая создается во время редактирования накладной
                    rez.RollbackTransaction
                    MsgBox "В справочнике Номенклатура не найден код " & CStr(rst!Код), vbInformation, "ВНИМАНИЕ"
                    GoTo ex_ex_wr
                 Else
                    Set newNom = selNom.ПолучитьОбъект()
                    'If newNom.ПометкаУдаления Then
                    '   newNom.ПометкаУдаления = False
                    '   newNom.Записать
                    'End If
                    '/ -----------------------------------------------------------
                    '/  Добавить                                                               
                    '/ -----------------------------------------------------------                                                                
                    Set newGoods = newDoc.Товары.Добавить()
                    newGoods.Номенклатура = newNom.ссылка  '.Номенклатура
                    newGoods.Количество = CLng(rst!Количество)
                    newGoods.СтавкаНДС = valNDS
                    newGoods.ЕдиницаИзмерения = selEI   'newNom.ЕдиницаХраненияОстатков
                    newGoods.Коэффициент = 1            'newNom.ЕдиницаХраненияОстатков.Коэффициент
                    If Me.flCorrect = False Then
                       newGoods.Цена = Format(rst!Сумма / rst!Количество, "0.00")
                       newGoods.Сумма = Format(rst!Сумма, "0.00")
                    Else
                       newGoods.Цена = Format(rst!Цена, "0.00")
                       newGoods.Сумма = Format(rst!Цена * rst!Количество, "0.00")
                    End If
                    newGoods.СуммаНДС = Format(rst!СуммаНДС, "0.00")
                    newGoods.СчетУчета = PLAccBU
                    newGoods.СчетУчетаНДС = PLAccNDS
                    
                    'newGoods.ОтражениеВУСН = selVusn
                    If rst!ГТД = "" Or IsNull(rst!ГТД) Then
                    Else
                       Set sek = rez.Справочники.НомераГТД.НайтиПоКоду(CStr(rst!ГТД))
                       If sek.Пустая() Then
                          Set sek = rez.Справочники.НомераГТД.СоздатьЭлемент()
                          sek.Код = CStr(rst!ГТД)
                          sek.Записать
                       Else
                          Set sek = sek.ПолучитьОбъект()
                       End If
                       newGoods.НомерГТД = sek.ссылка
                    End If
                    If rst!Страна = "" Or IsNull(rst!Страна) Then
                    Else
                      newGoods.СтранаПроисхождения = rez.Справочники.СтраныМира.НайтиПоНаименованию(CStr(rst!Страна)).ссылка 'Nz(CStr(rst!Страна))
                    End If
                    If Me.flPrice = True Then
                       '/ -----------------------------------------------------------                                                                       
                       '/ Цена закупочная
                       '/ -----------------------------------------------------------                                                                       
                       Set LinesPriceZ = DocPriceZ.Товары.Добавить()
                       LinesPriceZ.Номенклатура = newNom.ссылка
                       LinesPriceZ.Валюта = selVal
                       If Me.flCorrect = False Then
                          LinesPriceZ.Цена = Format((rst!Всего / rst!Количество) - rst!СуммаНДС, "0.00")
                       Else
                          LinesPriceZ.Цена = Format(rst!Цена, "0.00")
                       End If
                       '/ -----------------------------------------------------------                                                                         
                       '/ Цена розничная
                       '/ -----------------------------------------------------------                                                                         
                       Set LinesPriceR = DocPriceR.Товары.Добавить()
                       LinesPriceR.Номенклатура = newNom.ссылка
                       LinesPriceR.Валюта = selVal
                       If Me.flCorrect = False Then
                          pr = Format(rst!Всего / rst!Количество * Discount, "0.00")
                       Else
                          pr = Format((rst!Цена * Discount), "0.00")
                       End If
                       LinesPriceR.Цена = Round((Format(pr, "0") + 4) / 10) * 10
                    End If
                 End If
                 rst.MoveNext
              Loop
           End If
           '/
           'DocPrice.Записать (rez.РежимПроведенияДокумента.Неоперативный)
           '/ -----------------------------------------------------------
           '/  Записать
           '/ -----------------------------------------------------------                                                                               
           If Me.flPrice = True Then
             'If DocPriceZ.ПроверитьЗаполнение Then
                 On Error Resume Next
                 If rez.ПараметрыСеанса.ВыполняетсяОбновлениеИБ = False Then Debug.Print Err
                 If Err <> 0 Then rez.ПараметрыСеанса.ВыполняетсяОбновлениеИБ = False: Err.Clear
                 If rez.ПараметрыСеанса.ПропуститьПроверкуЗапретаИзменения = False Then Debug.Print Err
                 If Err <> 0 Then rez.ПараметрыСеанса.ПропуститьПроверкуЗапретаИзменения = False: Err.Clear
                 On Error GoTo ex_wr
                 DocPriceZ.Записать (rez.РежимЗаписиДокумента.Запись)
                 DocPriceZ.Записать (rez.РежимЗаписиДокумента.Проведение)
                 DocPriceR.Записать (rez.РежимЗаписиДокумента.Запись)
                 DocPriceR.Записать (rez.РежимЗаписиДокумента.Проведение)
                 If Err <> 0 Then
                    MsgBox "Проведите накладную другим числом !", vbInformation, "ВНИМАНИЕ"
                    GoTo ex_wr
                 End If
              'Else
              '   GoTo ex_wr
              'End If
           End If
           On Error GoTo ex_wr
           newDoc.Записать (rez.РежимЗаписиДокумента.Запись)
           'newDoc.Записать (rez.РежимЗаписиДокумента.Проведение)
           rez.CommitTransaction
           MsgBox "Накладная № " & Me!Nomer & " добавлена в 1С !", vbInformation, "ВНИМАНИЕ"
        End If
     End If
  Else
     MsgBox "Подключитесь к 1С !", vbInformation, "ВНИМАНИЕ"
     'If Err <> 0 Then
     ' MsgBox "Подключитесь к 1С !", vbInformation, "ВНИМАНИЕ"
     'ElseIf rez.TransactionActive = True Then
     '   MsgBox "Вы не одни в базе, принять накладную в данный момент невозможно,", vbInformation, "ВНИМАНИЕ"
     'End If
     Exit Sub
  End If
ex_ex_wr:
  Set refs = Nothing
  Set newSel = Nothing
  Set newDoc = Nothing
  Set newNom = Nothing
  Set newGoods = Nothing
  Set valVidOper = Nothing
  Set selDoc = Nothing
  Set selDog = Nothing
  Set selSkl = Nothing
  Set selNom = Nothing
  Set selVusn = Nothing
  Set PLAccBU = Nothing
  Set PLAccNDS = Nothing
  Set selClient = Nothing
  Set valNDS = Nothing
  Set regNom = Nothing
  Set Admin = Nothing
  Set DocPriceZ = Nothing
  Set DocPriceR = Nothing
  Set LinesPriceZ = Nothing
  Set LinesPriceR = Nothing
  Set SelTypPriceZ = Nothing
  Set SelTypPriceR = Nothing
  Set regPrice = Nothing
  Set selVal = Nothing
  Set selAcc = Nothing
  Set selDocPTR = Nothing
  Set selEI = Nothing
  Me.btGet.ForeColor = 0
  Me.AddFolders.ForeColor = 0
  Me.Add.ForeColor = 0
  Me.AddNomen.ForeColor = 0
ex_ex_rd:
  Exit Sub
'/ -----------------------------------------------------------
'/ Ошибка
'/ -----------------------------------------------------------                                                                                                                                                      
ex_wr:
  Set refs = Nothing
  Set newSel = Nothing
  Set newDoc = Nothing
  Set newNom = Nothing
  Set newGoods = Nothing
  Set valVidOper = Nothing
  Set selDoc = Nothing
  Set selDog = Nothing
  Set selSkl = Nothing
  Set selNom = Nothing
  Set selVusn = Nothing
  Set PLAccBU = Nothing
  Set PLAccNDS = Nothing
  Set selClient = Nothing
  Set valNDS = Nothing
  Set regNom = Nothing
  Set Admin = Nothing
  Set DocPriceZ = Nothing
  Set DocPriceR = Nothing
  Set LinesPriceZ = Nothing
  Set LinesPriceR = Nothing
  Set SelTypPriceZ = Nothing
  Set SelTypPriceR = Nothing
  Set regPrice = Nothing
  Set selVal = Nothing
  Set selAcc = Nothing
  Set selDocPTR = Nothing
  Set selEI = Nothing
  rez.RollbackTransaction
ex_rd:
  MsgBox "Произошла ошибка, повторите операцию !" & vbCrLf & Err.Description, vbInformation, "ВНИМАНИЕ"
End Sub
'/ -------------------------------- Fin
