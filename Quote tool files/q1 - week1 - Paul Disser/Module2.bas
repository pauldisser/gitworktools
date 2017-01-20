Attribute VB_Name = "Module2"
Dim i As Long
Dim emptyRow As Long
Dim custNum, orderNum, accName, email, special, check, publisher, quote, txt, selectedCell, htmlFormat, TSR, repFirstName, oneYear, threeYear, greeting, expWarning, datetxt As String
Dim currentRepName() As String
Dim expdate As Date
Dim count As Integer
Dim HtmlBefore, HtmlAfter As String


Public Sub sendQuote()
Attribute sendQuote.VB_ProcData.VB_Invoke_Func = "W\n14"
    getSelectedCell
    looper
    custFacingPublisherSwitch (publisher)
    setGreeting
    setExpirationWarning (expdate)
    custFacingSendMail (txt)
    
End Sub


Private Sub looper()
count = 0
Sheet3.Activate
  For i = 1 To Rows.count
        check = Cells(i, 1).Value
        If selectedCell = check Then
            custNum = Cells(i, 23).Text
            orderNum = Cells(i, 20).Text
            accName = Cells(i, 4).Text
            email = Cells(i, 25).Text
            expdate = Cells(i, 14).Value
            datetxt = Cells(i, 14).Text
            special = Cells(i, 22).Text
            publisher = Cells(i, 12).Text
            TSR = Cells(i, 27).Text
            count = count + 1
        End If
    Next i
Sheet2.Activate
For i = 1 To Rows.count
        check = Cells(i, 1).Value
        If selectedCell = check Then
            quote = Cells(i, 20).Text
            currentRepName() = Split(Cells(i, 7).Text, ", ")
            currentRepName(1) = Replace(currentRepName(1), " ", "")
        End If
    Next i
If count > 1 Then MsgBox "bad primary key, make sure info matches"

End Sub


Private Sub getSelectedCell()
Dim myCells As Range
Set myCells = Selection
selectedCell = myCells.Text
End Sub


Private Sub custFacingSendMail(msg As String)
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Working in Office 2000-2016

    Dim OutApp As Object
    Dim OutMail As Object
    Dim str, signature As String
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    actualMessage = msg
    htmlFormat = "<p style='font-family:'Calibri'' style='font-size:11px' style='color:#1F497D'>"
    If publisher = "VMware" Then
      OutMail.Attachments.Add "C:\Users\" + currentRepName(1) + "_" + currentRepName(0) + "\Documents\Quotes\" + oneYear + ".pdf"
      OutMail.Attachments.Add "C:\Users\" + currentRepName(1) + "_" + currentRepName(0) + "\Documents\Quotes\" + threeYear + ".pdf"
    Else
      OutMail.Attachments.Add "C:\Users\" + currentRepName(1) + "_" + currentRepName(0) + "\Documents\Quotes\" + quote + ".pdf"
    End If
    
    signature = GetBoiler("C:\Users\" + currentRepName(1) + "_" + currentRepName(0) + "\AppData\Roaming\Microsoft\Signatures\Main.htm")
    With OutMail
        .Recipients.Add email
        .CC = TSR
        .Recipients.Resolve
        .Subject = expWarning + publisher + " Renewal for " + accName
        .Display   'or use .Display
        .HTMLBody = htmlFormat + actualMessage + "</p>" + signature
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Private Sub custFacingPublisherSwitch(pub As String)
Select Case pub
    Case "VMware"
            oneYear = Left(quote, Len(quote) / 2)
            threeYear = Right(quote, Len(quote) / 2)
     
        txt = greeting + _
            "I知 " + currentRepName(1) + ", a member of your Software Renewals Team here at Dell. I am reaching out to you to let you know that your VMware Support & Subscription is due to expire. Renewal quotes from VMware are attached.<br><br>" + _
            "Your VMware Support & Subscription expire on " + datetxt + ".<br><br>" + _
            oneYear + " will reflect the price for the one year.<br><br>" + _
            threeYear + " will reflect the price for the three year renewal. Three year quotes come with a 12% discount built in.<br><br>" + _
            "If you want to maintain your licenses please respond, confirming which quote you would like to have processed into an order and how you would like to place the order.<br><br>" + _
            "**Please note: If you let this expire you will no longer receive technical support from VMware, or the ability to upgrade to the newest versions. If your subscription does lapse and you want to upgrade, you will have to either - pay for the months you missed plus reinstatement fees or repurchase the license**<br><br>" + _
            "Are you spinning up more than 25 Virtual Machines? Check out VMware's <a href=https://www.vmware.com/assessment/voa>Virtualization Optimization Assessment</a> to make sure everything is running efficiently<br><br>" + _
            "Thank you,"
    Case "Symantec"
        txt = greeting + _
            "I知 " + currentRepName(1) + ", a member of your Software Renewals Team here at Dell. We are reaching out to you to let you know that your " + publisher + " Protection is expiring. <br><br>" + _
            "This contract expires on " + datetxt + ".<br><br>" + _
            "Quote # " + quote + " is aligned with your previous order.<br><br>" + _
            "If you want to maintain your security please respond confirming that you would like to have this quote processed into an order and how you would like to place the order.<br><br>" & _
            "**Please note: If you let your license expire, you will have to pay reinstatement fees and will no longer receive up to date protection.**<br>" + _
            "More info: <a href=https://www.symantec.com/support-center/renewals/renewals-faq>https://www.symantec.com/support-center/renewals/renewals-faq</a><br><br>" + _
            "Thank you,"
    Case "Trend Micro"
        txt = greeting + _
            "I知 " + currentRepName(1) + ", a member of your Software Renewals Team here at Dell. We are reaching out to you to let you know that your " + publisher + " Maintenance is expiring. <br><br>" + _
            "This contract expires on " + datetxt + ".<br><br>" + _
            "Quote # " + quote + " is aligned with your previous order.<br><br>" + _
            "If you want to maintain your security please respond confirming that you would like to have this quote processed into an order and how you would like to place the order.<br><br>" & _
            "**Please note: When your Maintenance Agreement expires you will no longer receive up to date protection.**<br>" + _
            "More info: <a href= http://docs.trendmicro.com/all/ent/pp/v2.1/en-us/pp_2.1_olh/maintenance_agreement.htm>http://docs.trendmicro.com/all/ent/pp/v2.1/en-us/pp_2.1_olh/maintenance_agreement.htm</a><br><br>" + _
            "Thank you,"
    Case "VERITAS"
        txt = greeting + _
            "I知 " + currentRepName(1) + ", a member of your Software Renewals Team here at Dell. We are reaching out to you to let you know that your " + publisher + " Maintenance is expiring. <br><br>" + _
            "This contract expires on " + datetxt + ".<br><br>" + _
            "Quote # " + quote + " is aligned with your previous order.<br><br>" + _
            "If you want to maintain your licenses please respond confirming that you would like to have this quote processed into an order and how you would like to place the order.<br><br>" & _
            "**Please note: If you let your license expire, you will have to pay reinstatement fees.**<br>" + _
            "More Info: <a href=http://info.veritas.com/Global-Enterprise-Renewals-Policy?cname=17Q2-EMEA-CHNL-EN/EN-CAMP-VSPEAK_201608&eid=2110&cid=>http://info.veritas.com/Global-Enterprise-Renewals-Policy?cname=17Q2-EMEA-CHNL-EN/EN-CAMP-VSPEAK_201608&eid=2110&cid=</a><br><br>" + _
            "Thank you,"
    Case "Autodesk"
         txt = greeting + _
            "I知 " + currentRepName(1) + ", a member of your Software Renewals Team here at Dell. We are reaching out to you to let you know that your " + publisher + " Maintenance is expiring. <br><br>" + _
            "This contract expires on " + datetxt + ".<br><br>" + _
            "Quote # " + quote + " is aligned with your previous order.<br><br>" + _
            "If you want to maintain your licenses please respond confirming that you would like to have this quote processed into an order and how you would like to place the order.<br><br>" & _
            "**Please note: If you let your maintenance expire, you will no longer receive updates, access to new/old releases, and technical support (subscriptions will lose all access).**<br>" + _
            "More Info: <a href=https://knowledge.autodesk.com/customer-service/account-management/subscription-management/manage-contracts/renew-cancel/renew-maintenance-subscription>https://knowledge.autodesk.com/customer-service/account-management/subscription-management/manage-contracts/renew-cancel/renew-maintenance-subscription</a><br><br>" + _
            "Thank you,"
    Case "Microsoft Open Business"
        txt = greeting + _
            "I知 " + currentRepName(1) + ", a member of your Software Renewals Team here at Dell. We are reaching out to you to let you know that your Microsoft Software Assurance is expiring. <br><br>" + _
            "This contract expires on " + datetxt + ".<br><br>" + _
            "Quote # " + quote + " is aligned with your previous order.<br><br>" + _
            "If you want to maintain your Software Assurance please respond confirming that you would like to have this quote processed into an order and how you would like to place the order.<br><br>" & _
            "**Please note: If you let your SA expire, you will no longer receive updates, access to new/old releases, and technical support **<br><br>" + _
            "Thank you,"
    Case Else
        txt = greeting + _
            "I知 " + currentRepName(1) + ", a member of your Software Renewals Team here at Dell. We are reaching out to you to let you know that your " + publisher + " Maintenance is expiring. <br><br>" + _
            "This contract expires on " + datetxt + ".<br><br>" + _
            "Quote # " + quote + " is aligned with your previous order.<br><br>" + _
            "If you want to maintain your licenses please respond confirming that you would like to have this quote processed into an order and how you would like to place the order.<br><br>" & _
            "**Please note: If you let your maintenance expire, you will no longer receive updates, access to new/old releases, and technical support **<br><br>" + _
            "Thank you,"
    
End Select
End Sub

Function setGreeting()
    If Time > TimeValue("12:01") Then
        greeting = "Good Afternoon,<br><br>"
    Else
        greeting = "Good Morning,<br><br>"
    End If
End Function

Function setExpirationWarning(ByVal checkDate As Date)

If DateDiff("d", checkDate, Date) = 0 Then
    expWarning = "[expires today] "
ElseIf DateDiff("d", Date, checkDate) = 1 Then
    expWarning = "[expires tomorrow] "
ElseIf DateDiff("d", Date, checkDate) < 8 Then
    expWarning = "[expiring] "
Else
    expWarning = ""
End If
        
End Function

Function GetBoiler(ByVal sFile As String) As String
'Dick Kusleika
Dim FSO As Object
Dim ts As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Set ts = FSO.GetFile(sFile).OpenAsTextStream(1, -2)
GetBoiler = ts.ReadAll
ts.Close
End Function






