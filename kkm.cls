VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_KKM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'/ --------------------------------
'/ Форма ККМ (ФР Штрих-М)
'/ --------------------------------
Option Compare Database
Option Explicit

Private ECR As DrvFR
Private pass As String
Private chk As Long
Dim frm As Form
'Private DrvPay As ShtrihPayMan

'/ --------------------------------
'/ Функция при обращении к форме
'/ --------------------------------
Public Function setSource(chk)
  Dim sql As String
  Dim rst As DAO.Recordset
  Dim ret
  
  '/ Инициализация
  'ret = Status()
  'ret = Internet()
  sql = "SELECT Clients.f0, Clients.Computer, Контрагенты.Type, Контрагенты.Name, Контрагенты.FullName, Контрагенты.Mail, Контрагенты.Tel " _
      & "FROM Контрагенты RIGHT JOIN Clients ON Контрагенты.id = Clients.BPayerId WHERE (((Clients.f0)=" & chk & ") AND ((Контрагенты.BShopId) In (1,20,IsShop())));"
  Set rst = CurrentDb.OpenRecordset(sql)
  If rst.RecordCount = 0 Then
  Else
     '/ Накладная
     Me!Check = chk
     Me.Child0.Form.RecordSource = "SELECT f0 AS Чек, f1 AS Наименование, f13 AS Цена, f16 AS Количество, [f13]*[f16] AS Всего, N, f3 " _
                                 & "FROM ListBill WHERE (((ListBill.f0)=" & chk & ")) ORDER BY N WITH OWNERACCESS OPTION;"
     '/ Оплата
     Select Case Forms.НаСкладах.grPay
      Case 1
       Me!Pay = 1 '"Наличными"
      Case 2
       Me!Pay = 3  '"Карта Visa"
      Case 3
       Me!Pay = "Безналичный"
     End Select
     '/ Клиент
     If rst!Type = "Частное лицо" Then
        Me!Client = rst!FullName   'DLookup("FullName", "Контрагенты", "BShopId In (1,20,IsShop()) AND id=" & Forms.НаСкладах.lstPayer)
     ElseIf rst!Type = "Клиент" Then
        Me!Client = rst!FullName   'DLookup("FullName", "Контрагенты", "BShopId In (1,20,IsShop()) AND id=" & Forms.НаСкладах.lstClients)
     Else
        Me!Client = ""
     End If
     '/ Tel
     Dim sTel As String
     sTel = Nz(rst!Tel, "")
     If sTel = "" Then
     Else
        sTel = Replace(rst!Tel, " ", "")
        sTel = Split(sTel, ",")(0)
        sTel = Replace(sTel, "-", ""): sTel = Replace(sTel, "(", ""): sTel = Replace(sTel, ")", "")
        If Mid(sTel, 1, 1) = 8 Then Mid(sTel, 1, 1) = 7
        If Mid(sTel, 1, 1) <> 7 Then sTel = 7 & sTel
        If Len(sTel) <> 11 Then sTel = ""
        Me!Tel = sTel
     End If
     Me!Mail = rst!Mail
  End If
End Function

'/ --------------------------------
'/ Button оплата ККМ
'/ --------------------------------
Private Sub btnPayMan_Click()
  Dim ret
  'ret = DrvPay.ShowProperties
End Sub

'/ --------------------------------
'/ Button Свойства ККМ
'/ --------------------------------
Private Sub btnShowProperties_Click()
  Dim ret
  ret = ECR.ShowProperties
End Sub

'/ --------------------------------
'/ Выход из программы.
'/ --------------------------------
Private Sub Exit_Click()
  Set ECR = Nothing
  DoCmd.Close
End Sub

'/ --------------------------------
'/ Вставка почты по просьбе клиента
'/ --------------------------------
Private Sub flMail_Click()
  If Me.flTel Then
     Me.flMail = False
  Else
      If Me!flMail Then Me.Tel_Mail = Me!Mail Else Me.Tel_Mail = ""
  End If
End Sub
'/ --------------------------------
'/ Вставка тел. по просьбе клиента
'/ --------------------------------
Private Sub flTel_Click()
  If Me.flMail Then
     Me.flTel = False
  Else
     If Me.flTel Then
        If Mid(Me!Tel, 1, 1) <> "+" And Len(Me!Tel) <> 0 Then Me!Tel = "+" & Me!Tel
        Me.Tel_Mail = Me!Tel
     Else
        Me.Tel_Mail = ""
     End If
  End If
End Sub

'/ --------------------------------
'/ Точка входа в форму ККМ
'/ --------------------------------
Private Sub Form_Load()
  Call Info_Click
  If Not frm Is Nothing Then
     If frm.Name = "НаСкладах" Then
        Me.flSkaner.Locked = True
        Call setSource(Forms.НаСкладах.Check)
     Else
        Me.flSkaner.Locked = False
        Me.Child0.Form.RecordSource = "SELECT f0 AS Чек, f1 AS Наименование, f13 AS Цена, f16 AS Количество, [f13]*[f16] AS Всего, N, f3 " _
                                 & "FROM ListBill WHERE (((ListBill.f0)=0)) WITH OWNERACCESS OPTION;"
     End If
  Else
     Me.flSkaner.Locked = False
     Me.Child0.Form.RecordSource = "SELECT f0 AS Чек, f1 AS Наименование, f13 AS Цена, f16 AS Количество, [f13]*[f16] AS Всего, N, f3 " _
                                 & "FROM ListBill WHERE (((ListBill.f0)=0)) WITH OWNERACCESS OPTION;"
  End If
  Me!Нал = Format(Nz(dSUM("IncomeRU-LostRU", "GrossBush", "Примечание='" & Check & "'"), 0), "0.00")
  Me!БН = Format(Nz(dSUM("IncomeRU-LostRU", "GrossBush_Bank", "Примечание='" & Check & "'"), 0), "0.00")
  Me!getFromCln.SetFocus
End Sub

'/ --------------------------------
'/ Точка открытия формы ККМ
'/ --------------------------------
Private Sub Form_Open(Cancel As Integer)
  Dim ret
  
  ' создаем объект общего драйвера ККМ
  ' если объект создать не удается генерируется исключение, по которому завершается работа приложения
  If IsLoaded("НаСкладах") Then Set frm = Screen.ActiveForm
  ret = AdminMode()
  If ret <> 0 Then
     MsgBox "У Вас нет прав, обратитесь к Администратору!", vbInformation, "ВНИМАНИЕ"
     DoCmd.Close acForm, "KKM"
     Exit Sub
  End If
  On Error GoTo No_KKM
  Set ECR = New DrvFR   'Set ECR = CreateObject("AddIn.DrvFR")
  '/ Продавец и его пароль
  pass = GetSetting("STOCK", "Preferences", "WhoIs", 1)
  Me.Caption = "ККМ :: " & DLookup("ФИО", "Salers", "Номер=" & pass)
  If pass = 1 Then pass = 30
  ECR.ClearResult
  ECR.Password = pass
  ECR.Connect
  If ECR.ResultCode = -2 Then
     If ECR.Connected Then ECR.Disconnect
     MsgBox "Включите ФР или подключите кабель!", vbInformation, "ВНИМАНИЕ"
     DoCmd.Close acForm, "KKM"
     Exit Sub
  End If
  If ECR.ResultCode = -3 Then
     If ECR.Connected Then ECR.Disconnect
     MsgBox "ФР занят, подключитесь позже!", vbInformation, "ВНИМАНИЕ"
     DoCmd.Close acForm, "KKM"
     Exit Sub
  End If
  If ECR.OperatorNumber = 0 Then
     If ECR.Connected Then ECR.Disconnect
     MsgBox "Вы не являетесь кассиром, обратитесь к Администратору!", vbInformation, "ВНИМАНИЕ"
     DoCmd.Close acForm, "KKM"
     Exit Sub
  End If
  ret = Status()
  ret = getOpenDay
  ret = Internet()
  Call setStatErr
  Call getReadyFR
  ECR.Disconnect
  Exit Sub
No_KKM: 'Обработка ошибки при создании OLE-объекта
  If ECR.ResultCode <> 0 Or Err <> 0 Then Call setStatErr
  Cancel = 1
End Sub

'/ --------------------------------
'/ 
'/ --------------------------------
Private Sub getFromCln_AfterUpdate()
  Me!change = Me!getFromCln - Me!Total
End Sub

'/ ------------------------------------------------------
' проверка связи с ОФД (оператор фискализированных данных
'/ ------------------------------------------------------
Private Function Internet() As String
  ' ECR.CheckConnection
  ' ECR.CheckFMConnection
  'If pingComputer("www.taxcom.ru") = 0 Then
  If ECR.ServerConnected Then
     Me.I0.Visible = False
     Me.I1.Visible = True
     Internet = 0
  Else
     Me.I0.Visible = True
     Me.I1.Visible = False
     Internet = 1
  End If
End Function

'/ --------------------------------
'/ Состояние ФР
'/ --------------------------------
Private Function Status() As String
  Dim ret
  
  ' получаем состояние ККМ
  If ECR.GetShortECRStatus = 0 Then
     ECR.Beep
     DoEvents
     Me.K0.Visible = False
     Me.K1.Visible = True
     Status = 0
     ret = FMode()
  Else
     If ECR.ResultCode <> 0 Then Call setStatErr
     '/ ККМ подключен
     Me.K0.Visible = True
     Me.K1.Visible = False
     '/ Фискализация
     Me.f0.Visible = True
     Me.f1.Visible = False
          
     Status = 1
  End If
End Function

'/ ---------------------------------------------------------------------------
'/ Проверка, включён ли фискальный режим или вставлен ли фискальный накопитель
'/ ---------------------------------------------------------------------------
Private Function FMode() As String
  
  ' получаем режим работы ККМ (фискальный или demo)
  If ECR.CheckFM = 0 Then
     Me.f0.Visible = False
     Me.f1.Visible = True
     FMode = 0
  Else
     Me.f0.Visible = True
     Me.f1.Visible = False
     FMode = 1   'ECR.ResultCode
     ECR.ClearResult
  End If
  Me.lbFiscal.Caption = "Фискализация:" & ECR.OpenDocumentNumber
  Call setStatErr
End Function

'/ --------------------------------------------------------
'/ Проверка, есть ли доступ к ФР в качестве администратора.
'/ --------------------------------------------------------
Private Function AdminMode() As String
  Dim RA As Variant
  
  RA = RightAccess("B4")
  If RA = "Администратор" Or RightAccess("B4") <> False Then
     Me.A0.Visible = False
     Me.A1.Visible = True
     Me.Info.Visible = True
     AdminMode = 0
  Else
     Me.A0.Visible = True
     Me.A1.Visible = False
     Me.Info.Visible = False
     AdminMode = 1
  End If
End Function

'/ --------------------------------
'/ Печать ФР
'/ --------------------------------
Private Sub Operations()
  Dim rsk As DAO.Recordset
  Dim rst As DAO.Recordset
  Dim Str As String
  Dim PrintWidth As Integer
  Dim Summ As Double
  Dim Discount As Double
  Dim str1 As String
  Dim str2 As String
  Dim str3 As String
  Dim str4 As String
  
  chk = Nz(Me!Check, 0)
  If IsNull(chk) Or chk = 0 Then Exit Sub
  If ECR Is Nothing Then Set ECR = CreateObject("Addin.DRvFR")
  PrintWidth = 48                       '/максимальная ширина строки печати в символах на ШТРИХ-М-01Ф
  Summ = 0
  Set rsk = CurrentDb.OpenRecordset("SELECT Tel, Address, Mail FROM Контрагенты WHERE id=" & ShopIdx())
  Set rst = CurrentDb.OpenRecordset("SELECT f1,f3,f10,f13,f16,f20,Discount FROM ListBill WHERE f0=" & chk & " ORDER BY N;")
  On Error GoTo exit_sub_KKM
  ECR.Connect
  ECR.ClearResult
  ' получаем состояние ККМ
  If ECR.ResultCode <> 0 Then
     Call setStatErr
     If ECR.Connected Then ECR.Disconnect
  Else
     '/ далее идут проверки режимов ФР, чтобы можно было регистрировать продажу.
     '/ здесь проверяется на 3 режима, если кончилась смена, незавершённый чек и если кончилась чековая лента.
     If ECR.ECRMode = 3 Then
        MsgBox ("Закончилась смена, выведите Z-отчёт!")
        If ECR.Connected Then ECR.Disconnect
        Exit Sub
     End If
     If ECR.ECRMode <> 0 Then
        If ECR.ECRMode = 8 Then
           '/ При открытом документе выполнить закрытие
           'If MsgBox("Документ открыт, чек печатать?", vbYesNo + vbInformation, "ВНИМАНИЕ") = vbNo Then
           ECR.Password = "30"
           ECR.SysAdminCancelCheck
        ElseIf ECR.ECRMode = 2 Then
        Else
           If ECR.Connected Then ECR.Disconnect
           MsgBox "Принтер не готов: " & ECR.ECRModeDescription, vbInformation, "ВНИМАНИЕ"
           Exit Sub
        End If
     End If
     If ECR.ECRAdvancedMode = 1 Then
        'Do Until oDriver.ECRAdvancedMode = 0
           MsgBox ("Закончилась чековая лента, заправьте новую и нажмите Ок!")
           If ECR.Connected Then ECR.Disconnect
           Exit Sub
           'ECR.GetShortECRStatus
        'Loop
     End If
     If MsgBox("Чек печатать?", vbYesNo + vbInformation, "ВНИМАНИЕ") = vbNo Then
        If ECR.Connected Then ECR.Disconnect
        Exit Sub
     End If
     '/ Проверка, есть ли фиск. накопитель.
     '/ Если есть то устанавливаем номер и делаем его активным.
     '/ IsFMSessionOpen
     If ECR.CheckFM = 0 Then
        ECR.LDNumber = 1     '/ номер логического устройства = 1
        ECR.SetActiveLD
     End If
     '/ 1 Блок, Заголовок, брать из B_Shop или Контрагенты
     ECR.StringForPrinting = "................................................"
     ECR.PrintString
     
     '/ 2 Блок, Печать фискальных данных
     ECR.UseReceiptRibbon = True   '/ печать на кассовой ленте
     ECR.UseJournalRibbon = False  '/ печать в журнале
     '/ открываем чек
     '  ECR.CheckMode:=0;  '// 0 - электронный 1 - бумажный
     ECR.CheckType = 0 '// 0 - продажа, 1 - покупка, 2 - возврат продажи, 3 - возврат покупки.
     '  if ECR.OpenCheck <> 0 then Exit;
                
     '/ Отправка данных о покупателе
     'ECR.CustomerEmail = "+7хххххххххх"  '/ В качестве параметра можно передавать или телефон (+7хххххххххх) или e-mail: test@test.com
     'ECR.FNSendCustomerEmail
     If Me.flTel Then
        If Mid(Me.Tel_Mail, 1, 1) <> "+" And Len(Me!Tel) <> 12 Then
           MsgBox "Неверный формат телефона!", vbInformation, "ВНИМАНИЕ"
           If ECR.Connected Then ECR.Disconnect
           Exit Sub
        End If
        ECR.CustomerEmail = Me.Tel_Mail
        ECR.FNSendCustomerEmail
     End If
     If Me.flMail Then
        If InStr(1, Me.Tel_Mail, "@") = 0 Then
           MsgBox "Неверный формат e-mail!", vbInformation, "ВНИМАНИЕ"
           If ECR.Connected Then ECR.Disconnect
           Exit Sub
        End If
        ECR.CustomerEmail = Me.Tel_Mail
        ECR.FNSendCustomerEmail
     End If
     
     ECR.Password = pass                '/ по умолчанию пароль администратора 30 или продавца
     ECR.openCheck                      '/ открытие чека
     If ECR.ResultCode <> 0 Then
        Call setStatErr
        If ECR.Connected Then ECR.Disconnect
        Exit Sub
     End If
     '/ 3 Блок, Заполнение чека
     Dim I As Integer
     If rst.RecordCount <> 0 Then
        'str = "Кассир" & Space(42 - Len(Mid(Me.Caption, 7))) & Mid(Me.Caption, 7)
        'ECR.StringForPrinting = str
        'ECR.PrintString
        ECR.StringForPrinting = "------------------------------------------------"
        ECR.PrintString
        Do While Not rst.EOF
           I = I + 1
           'If I = 5 Then Exit Do
           '/ 1 строка Наименование
           str1 = I & ". " & Mid(rst!f1, 1, 45)
           ECR.StringForPrinting = Mid(str1, 1, 48)
           ECR.PrintString
           '/ 2 строка Количество-Цена-Сумма
           str2 = Format(rst!f16, "0.000") & " Х " & Format(rst!f10, "0.00") & " = " & Format(rst!f16 * rst!f10, "0.00")  '/ Chr(&H1F)
           ECR.StringForPrinting = Space(46 - Len(str2)) & Replace(str2, ",", ".") & "_A"
           ECR.PrintString
           '/ 3 строка Скидка
           str3 = Format(rst!f16 * (rst!f10 - rst!f13), "0.00")
           If str3 = 0 Then
              ECR.StringForPrinting = "Скидка" & Space(42 - Len(str3)) & Replace(str3, ",", ".")
           Else
              ECR.StringForPrinting = "Скидка" & Space(41 - Len(str3)) & "-" & Replace(str3, ",", ".")
           End If
           ECR.PrintString
           '/ 4 строка НДС
           str4 = Format(rst!f13 - (rst!f13 / 1.18), "0.00")
           ECR.StringForPrinting = "НДС 18%" & Space(41 - Len(str4)) & str4
           ECR.PrintString
           
           ECR.Quantity = rst!f16                                      ' Quantity, количество
           ECR.Price = rst!f13                                         ' Price, цена
           ECR.Department = 1                                          ' Department, номер отдела
           ECR.discountvalue = rst!f16 * (rst!f10 - rst!f13)           ' Скидка
           ECR.Tax1 = 1
           'ECR.ChargeValue = 0
           'ECR.BarCode = "123456789"
           '/ Новая реализация
           ECR.StringForPrinting = "//" & Mid(rst!f1, 1, 48) 'Наименование"
           ECR.DocumentName = ""
           ECR.FNDiscountOperation
           'ECR.CashIncome '/ CahInCome - внесение отличается от Sale тем, что не фиксирует продажу в фискальном памяти,
           '/ используется в начале смены, для того чтобы у кассира были наличные для сдачи.
           If ECR.ResultCode <> 0 Then
              rst.Close
              Set rst = Nothing
              GoTo exit_sub_KKM
           End If
           Summ = Summ + rst!f16 * rst!f13
           Discount = Discount + rst!f16 * (rst!f10 - rst!f13)
           rst.MoveNext
        Loop
     End If
     rst.Close
     Set rst = Nothing
     '/Производим продажу на всю сумму чека
     'ECR.Price = Round(Summ, 0)
     'ECR.Quantity = 1
     'ECR.Department = 0
     'ECR.Sale

     ECR.StringForPrinting = "------------------------------------------------"
     ECR.PrintString
     If Discount = 0 Then
        ECR.StringForPrinting = "Ваша скидка составила" & Space(23) & "0.00"
     Else
        ECR.StringForPrinting = "Ваша скидка составила" & Space(29 - Len(Discount)) & Replace(Format(Discount, "0.00"), ",", ".")
     End If
     ECR.PrintString
     ECR.StringForPrinting = Space(48)
     ECR.PrintString
     If Me.getFromCln = 0 Then
        ECR.StringForPrinting = "Получено" & Space(41 - Len(Summ)) & Replace(Format(Summ, "0.00"), ",", ".")
        ECR.PrintString
        ECR.StringForPrinting = "Оплата" & Space(43 - Len(Summ)) & Replace(Format(Summ, "0.00"), ",", ".")
        ECR.PrintString
        ECR.StringForPrinting = "Сдача" & Space(39) & "0.00"
        ECR.PrintString
     Else
        ECR.StringForPrinting = "Получено" & Space(41 - Len(Discount)) & Replace(Format(Me!getFromCln, "0.00"), ",", ".")
        ECR.PrintString
        ECR.StringForPrinting = "Оплата" & Space(43 - Len(Summ)) & Replace(Format(Summ, "0.00"), ",", ".")
        ECR.PrintString
        ECR.StringForPrinting = "Сдача" & Space(45 - Len(Discount)) & Replace(Format(Me!getFromCln - Me!Total, "0.00"), ",", ".")
        ECR.PrintString
     End If
     ECR.StringForPrinting = "------------------------------------------------"
     '/ Закрытие чека
     'ECR.Password = pass
     Select Case Forms.НаСкладах.grPay
       Case 1
        ECR.Summ1 = Round(Summ, 0)
        ECR.Summ2 = 0
        ECR.Summ3 = 0
        ECR.Summ4 = 0
       Case 2
        ECR.Summ1 = 0
        ECR.Summ2 = 0
        ECR.Summ3 = Round(Summ, 0)
        ECR.Summ4 = 0
     End Select
     'ECR.discountvalue = Discount
     'ECR.DiscountOnCheck = 5   '/ Скидка 5%
     ECR.Discount   '/ со скидкой
     ECR.Tax1 = 1
     ECR.Tax2 = 0
     ECR.Tax3 = 0
     ECR.Tax4 = 0
          
     'ECR.DiscountOnCheck = 17 '/ Скидка 5%
     'ECR.Discount   '/ со скидкой
     
     '/ 4 Блок, Итоговая сумма
     ECR.Password = pass
     ECR.CheckSubTotal
     '/ 5 Блок, № фискального документа и QR документа
     ECR.CloseCheck
     If ECR.ResultCode <> 0 Then Call setStatErr
     ECR.Disconnect
  End If
  Exit Sub
exit_sub_KKM:
  If ECR.ResultCode <> 0 Or Err <> 0 Then Call setStatErr
End Sub

'/ --------------------------------
'/ Вывод Инфы
'/ --------------------------------
Private Sub Info_Click()
  Dim Left As Variant
  Dim Top As Variant
  Dim Width As Variant
  Dim Height As Variant
  
  If Me.Form.WindowWidth = 10500 Then Width = 905 Else Width = 700
  Left = 300
  Top = 300
  Height = 530
  Call MoveWindow(Me.hwnd, Left, Top, Width, Height, 1)
End Sub

'/ --------------------------------
'/ Основные функции ККМ
'/ Bottons
'/ --------------------------------

'/ --------------------------------
'/ Сессия открыта ли
'/ --------------------------------
Private Function getOpenDay() As String
  If ECR.ECRMode = 2 Then
     Me.S0.Visible = False
     Me.S1.Visible = True
  Else
     Me.S0.Visible = True
     Me.S1.Visible = False
  End If
End Function

'/ --------------------------------
'/ Начать день или открыть смену
'/ --------------------------------
Private Sub OpenDay_Click()
  Dim ret
  
  If ECR.ECRMode = 4 Then
     '/ закрытая смена
     If MsgBox("Открыть смену?", vbYesNo + vbInformation, "ВНИМАНИЕ") = vbYes Then
        ECR.Password = "30"
        ECR.Connect
        ECR.FNOpenSession
        ret = getOpenDay
        Me.Refresh
        setStatErr
        ECR.Disconnect
     End If
  ElseIf ECR.ECRMode = 2 Then
     '/ открытая смена
     MsgBox "Смена открыта!", vbInformation, "ВНИМАНИЕ"
     setStatErr
  Else
     setStatErr
  End If
End Sub

'/ --------------------------------
'/ Закончить день или закрыть смену
'/ --------------------------------
Private Sub CloseDay_Click()
  If MsgBox("Закрыть смену?", vbYesNo + vbInformation, "ВНИМАНИЕ") = vbYes Then
     ECR.Password = "30"
     ECR.Connect
     If ECR.IsFMSessionOpen Then
        ECR.FNCloseSession
        If ECR.ResultCode <> 0 Then Call setStatErr
     End If
     ECR.Disconnect
  End If
End Sub

'/ --------------------------------
'/ Повтор последнего документа
'/ --------------------------------
Private Sub Repeat_Click()
  ECR.Password = "30"
  ECR.Connect
  ECR.RepeatDocument
  Call setStatErr
  ECR.Disconnect
End Sub

'/ --------------------------------
'/ Печатать Х-Отчёт
'/ --------------------------------
Private Sub XReport_Click()
  If MsgBox("Снять X-Отчёт?", vbYesNo + vbInformation, "ВНИМАНИЕ") = vbYes Then
     If ECR Is Nothing Then Set ECR = CreateObject("Addin.DRvFR")
     ECR.Password = "30"
     ECR.Connect
     If ECR.ResultCode <> 0 Then
        Me.StatusBar.Panels(1).text = "Ошибка : " & ECR.ResultCode & " - " & ECR.ResultCodeDescription
        If ECR.Connected Then ECR.Disconnect
        Exit Sub
     End If
     ECR.PrintReportWithoutCleaning
     If ECR.ResultCode <> 0 Then
        Me.StatusBar.Panels(1).text = "Ошибка : " & ECR.ResultCode & " - " & ECR.ResultCodeDescription
        If ECR.Connected Then ECR.Disconnect
        Exit Sub
     End If
     ECR.Disconnect
  End If
End Sub

'/ --------------------------------
'/ Печатать Z-Отчёт
'/ --------------------------------
Private Sub ZReport_Click()
  If MsgBox("Снять Z-Отчёт?", vbYesNo + vbInformation, "ВНИМАНИЕ") = vbYes Then
     If ECR Is Nothing Then Set ECR = CreateObject("Addin.DRvFR")
     ECR.Password = "30"
     ECR.Connect
     If ECR.ResultCode <> 0 Then
        Me.StatusBar.Panels(1).text = "Ошибка : " & ECR.ResultCode & " - " & ECR.ResultCodeDescription
        If ECR.Connected Then ECR.Disconnect
        Exit Sub
     End If
     ECR.PrintReportWithCleaning
     If ECR.ResultCode <> 0 Then
        Me.StatusBar.Panels(1).text = "Ошибка : " & ECR.ResultCode & " - " & ECR.ResultCodeDescription
        If ECR.Connected Then ECR.Disconnect
        Exit Sub
     End If
     ECR.Disconnect
  End If
End Sub

'/ --------------------------------
'/ Открыть денежный ящик
'/ --------------------------------
Private Sub Case_Click()
  ECR.Password = pass
  ECR.Connect
  ECR.DrawerNumber = 0
  ECR.OpenDrawer
  Call setStatErr
  ECR.Disconnect
End Sub

'/ --------------------------------
'/ Вывести Чек на принтер ФР
'/ --------------------------------
Private Sub PrintCheck_Click()
  Call Operations
End Sub

'/ --------------------------------
'/ Готовность ФР,режим ККМ в данный момент
'/------------------------------------
' Режим ККМ Описание режима ККМ
'
' 0  Принтер в рабочем режиме
' 1  Выдача данных
' 2  Открытая смена, 24 часа не кончились
' 3  Открытая смена, 24 часа кончились
' 4  Закрытая смена
' 5  Блокировка по неправильному паролю налогового инспектора
' 6  Ожидание подтверждения ввода даты
' 7  Разрешение изменения положения десятичной точки
' 8  Открытый Документ
' 9  Режим разрешения технологического обнуления
' 10 Тестовый прогон
' 11 Печать полного фискального отчета
' 12 Печать длинного отчета ЭКЛЗ
' 13 Работа с фискальным подкладным документом
' 14 Печать подкладного документа
' 15 Фискальный подкладной документ сформирован
'/-----------------------------------
Private Sub getReadyFR()
  If ECR.ECRMode <> 0 Then
     If ECR.ECRMode = 2 Then
     Else
        MsgBox "Принтер не готов: " & ECR.ECRModeDescription, vbInformation, "ВНИМАНИЕ"
     End If
  End If
End Sub

'/ --------------------------------
'/ Вывести состояние ККМ в StatusBar
'/ --------------------------------
Private Sub setStatErr()
  Select Case ECR.ResultCode
    Case 0
      If Err <> 0 Then Me.StatusBar.Panels(1).text = Err.Description Else Me.StatusBar.Panels(1).text = "Ок"
    Case -2
    Case Else
      Me.StatusBar.Panels(1).text = "Ошибка : " & ECR.ResultCode & " - " & ECR.ResultCodeDescription
  End Select
End Sub
'/----------------------------------- Fin
