Attribute VB_Name = "Module4"
Dim selectedCell As String
Dim i As Long
Dim IE As Object
Dim objElement As Object
Dim objCollection As Object

Public Sub customerSearch()
Attribute customerSearch.VB_ProcData.VB_Invoke_Func = "R\n14"
getSelectedCell
searchDSA
End Sub

Private Sub searchDSA()
' Create InternetExplorer Object
 
    Set IE = New InternetExplorerMedium
 
    ' You can uncoment Next line To see form results
    IE.Visible = True
 
    ' Send the form data To URL As POST binary request
    IE.Navigate "http://sales.dell.com/#/customer2/organization/search"

    newPage
    Application.Wait (Now + TimeValue("0:00:01"))
    
    IE.Document.getElementById("customerSearch_customerNumber").Value = selectedCell
    IE.Document.getElementById("customerSearch_customerNumber").Focus
    SendKeys " ", True
    'Application.Wait (Now + TimeValue("0:00:01"))
    IE.Document.getElementById("customerSearch_searchAction").Click
     
End Sub

Private Sub getSelectedCell()
Dim myCells As Range
Set myCells = Selection
selectedCell = myCells.Text
End Sub

Private Sub newPage()
    Do While IE.Busy
       DoEvents
    Loop
End Sub
