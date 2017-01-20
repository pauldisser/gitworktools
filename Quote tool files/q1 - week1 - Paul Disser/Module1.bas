Attribute VB_Name = "Module1"
Dim i As Long
Dim emptyRow As Long
Dim custNum, orderNum, accName, email, expdate, special, check, publisher, quote, txt, recip, selectedCell, signature, htmlFormat As String
Dim count As Integer
Dim currentRepName() As String

Public Sub requestQuote()
Attribute requestQuote.VB_ProcData.VB_Invoke_Func = "Q\n14"
    getSelectedCell
    looper
    publisherSwitch (publisher)
    sendMail (txt)
End Sub

Private Sub getSelectedCell()
Dim myCells As Range
Set myCells = Selection
selectedCell = myCells.Text
End Sub

Private Sub sendMail(msg As String)
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Working in Office 2000-2016

    Dim OutApp As Object
    Dim OutMail As Object
    
    On Error Resume Next
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
  For i = 1 To Rows.count
        check = Cells(i, 1).Value
        If selectedCell = check Then
            custNum = Cells(i, 16).Text
            orderNum = Cells(i, 15).Text
            accName = Cells(i, 3).Text
            email = Cells(i, 24).Text
            expdate = Cells(i, 4).Text
            special = Cells(i, 25).Text
            publisher = Cells(i, 14).Text
            currentRepName() = Split(Cells(i, 7).Text, ", ")
            currentRepName(1) = Replace(currentRepName(1), " ", "")
            count = count + 1
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
            "<br>Expiration Date: " + expdate + _
            "<br>Contract: " + special

        recip = "Joshua_Pence@Dell.com"
    Case "Symantec"
          txt = "Hello,<br><br>" + _
            "Could I get a quote for this?<br><br>" + _
            "Account name: " + accName & _
            "<br>Renewal pin: " + special + _
            "<br>Cust num: " + custNum & _
            "<br>Expiration Date: " + expdate
        recip = "symantec-dell@ingrammicro.com"
    Case "Autodesk"
        txt = "Hello,<br><br>" + _
            "Could I get a quote for this?<br><br>" + _
            "Account name: " + accName & _
            "<br>Serial Number: " + special + _
            "<br>Cust num: " + custNum & _
            "<br>Expiration Date: " + expdate

        recip = "James_Lout@Dell.com"
    Case "Intel Security"
        txt = "Hello,<br><br>" + _
            "Could I get a quote for this?<br><br>" + _
            "Account name: " + accName & _
            "<br>Grant Number: " + special + _
            "<br>Cust num: " + custNum & _
            "<br>Expiration Date: " + expdate

        recip = "IntelSecurity@techdata.com"
    Case "Trend Micro"
        txt = "Hello,<br><br>" + _
            "Could I get a quote for this?<br><br>" + _
            "Account name: " + accName & _
            "<br>License Authorization Number: " + special + _
            "<br>Cust num: " + custNum & _
            "<br>Expiration Date: " + expdate
             
        recip = "Trend-Licensing@ingrammicro.com"
    Case "VERITAS"
        txt = "Hello,<br><br>" + _
            "Could I get a quote for this?<br><br>" + _
            "Account name: " + accName & _
            "<br>Special: " + special + _
            "<br>Cust num: " + custNum & _
            "<br>Expiration Date: " + expdate

        recip = "Veritas-Dell@ingrammicro.com"
    Case Else
        txt = "Hello,<br><br>" + _
            "Could I get a quote for this?<br><br>" + _
            "Account name: " + accName & _
            "<br>Special: " + special + _
            "<br>Cust num: " + custNum & _
            "<br>Expiration Date: " + expdate

        recip = ""
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


