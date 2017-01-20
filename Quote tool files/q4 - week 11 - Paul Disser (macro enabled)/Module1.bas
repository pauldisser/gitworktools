Attribute VB_Name = "Module1"
Dim i As Long
Dim emptyRow As Long
Dim custNum, orderNum, accName, email, expdate, special, check, publisher, quote, txt, recip, selectedCell, signature, htmlFormat As String
Dim count As Integer
Dim currentRepName() As String

Public Sub requestQuote()
Attribute requestQuote.VB_ProcData.VB_Invoke_Func = "Q\n14"
getSelectedCell
Sheet3.Activate
    looper
    Sheet2.Activate
    publisherSwitch (publisher)
    sendMail (txt)
End Sub

Private Sub getSelectedCell()
Attribute getSelectedCell.VB_ProcData.VB_Invoke_Func = " \n14"
Dim myCells As Range
Set myCells = Selection
selectedCell = myCells.text
End Sub

Private Sub sendMail(msg As String)
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Working in Office 2000-2016

    Dim OutApp As Object
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    signature = GetBoiler("C:\Users\" + currentRepName(1) + "_" + currentRepName(0) + "\AppData\Roaming\Microsoft\Signatures\Main.htm")
    htmlFormat = "<p style='font-family:'Calibri'' style='font-size:11px' style='color:#1F497D'>"
    actualMessage = msg
    With OutMail
        .Recipients.Add recip
        .Subject = publisher + " Renewal for " + accName + ", DCN: " + custNum
        .Display
        .HTMLBody = htmlFormat + actualMessage + "</p>" + signature
        
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Private Sub looper()
count = 0
Sheet3.Activate
  For i = 1 To Rows.count
        check = Cells(i, 1).Value
        If selectedCell = check Then
            custNum = Cells(i, 23).text
            orderNum = Cells(i, 20).text
            accName = Cells(i, 4).text
            email = Cells(i, 25).text
            expdate = Cells(i, 14).text
            special = Cells(i, 22).text
            publisher = Cells(i, 12).text
            count = count + 1
        End If
    Next i
Sheet2.Activate
For i = 1 To Rows.count
        check = Cells(i, 1).Value
        If selectedCell = check Then
            currentRepName() = Split(Cells(i, 7).text, ", ")
            currentRepName(1) = Replace(currentRepName(1), " ", "")
        End If
    Next i
If count > 1 Then MsgBox "bad primary key, make sure info matches"

End Sub


Private Sub publisherSwitch(pub As String)
Select Case pub
    Case "VMware"
        txt = "Hello,<br><br>" + _
            "Could I get a quote for this?<br><br>" + _
            "Account name: " + accName & _
            "<br>Cust num: " + custNum & _
            "<br>Order num: " + orderNum & _
            "<br>Email: " + email & _
            "<br>Expiration Date: " + expdate + _
            "<br>Contract: " + special + _
            "<br><br>Thank you,"
        recip = "Joshua_Pence@Dell.com"
    Case "Symantec"
          txt = "Hello,<br><br>" + _
            "Could I get a quote for this?<br><br>" + _
            "Account name: " + accName & _
            "<br>Renewal pin: " + special + _
            "<br>Cust num: " + custNum & _
            "<br>Order num: " + orderNum & _
            "<br>Email: " + email & _
            "<br>Expiration Date: " + expdate + _
            "<br><br>Thank you,"
        recip = "symantec-dell@ingrammicro.com"
    Case "Autodesk"
        txt = "Hello,<br><br>" + _
            "Could I get a quote for this?<br><br>" + _
            "Account name: " + accName & _
            "<br>Serial Number: " + special + _
            "<br>Cust num: " + custNum & _
            "<br>Order num: " + orderNum & _
            "<br>Email: " + email & _
            "<br>Expiration Date: " + expdate + _
            "<br><br>Thank you,"
        recip = "James_Lout@Dell.com"
    Case "Intel Security"
        txt = "Hello,<br><br>" + _
            "Could I get a quote for this?<br><br>" + _
            "Account name: " + accName & _
            "<br>Grant Number: " + special + _
            "<br>Cust num: " + custNum & _
            "<br>Order num: " + orderNum & _
            "<br>Email: " + email & _
            "<br>Expiration Date: " + expdate + _
            "<br><br>Thank you,"
        recip = "IntelSecurity@techdata.com"
    Case "Trend Micro"
        txt = "Hello,<br><br>" + _
            "Could I get a quote for this?<br><br>" + _
            "Account name: " + accName & _
            "<br>License Authorization Number: " + special + _
            "<br>Cust num: " + custNum & _
            "<br>Order num: " + orderNum & _
            "<br>Email: " + email & _
            "<br>Expiration Date: " + expdate + _
            "<br><br>Thank you,"
        recip = "Trend-Licensing@ingrammicro.com"
    Case "VERITAS"
        txt = "Hello,<br><br>" + _
            "Could I get a quote for this?<br><br>" + _
            "Account name: " + accName & _
            "<br>Special: " + special + _
            "<br>Cust num: " + custNum & _
            "<br>Order num: " + orderNum & _
            "<br>Email: " + email & _
            "<br>Expiration Date: " + expdate + _
            "<br><br>Thank you,"
        recip = "Veritas-Dell@ingrammicro.com"
    Case Else
        txt = "Hello,<br><br>" + _
            "Could I get a quote for this?<br><br>" + _
            "Account name: " + accName & _
            "<br>Special: " + special + _
            "<br>Cust num: " + custNum & _
            "<br>Order num: " + orderNum & _
            "<br>Email: " + email & _
            "<br>Expiration Date: " + expdate + _
            "<br><br>Thank you,"
        recip = Null
End Select
End Sub


Function GetBoiler(ByVal sFile As String) As String
'Dick Kusleika
Dim FSO As Object
Dim ts As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Set ts = FSO.GetFile(sFile).OpenAsTextStream(1, -2)
GetBoiler = ts.ReadAll
ts.Close
End Function
