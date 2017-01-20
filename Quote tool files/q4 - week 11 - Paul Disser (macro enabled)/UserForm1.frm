VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5304
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8610
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long
Dim emptyRow As Long
Dim custNum, orderNum, accName, email, expdate, special, check, publisher, quote, txt, recip, text As String
Dim obj As New MSForms.DataObject
Dim quotesArr() As String
Dim count As Integer
Dim HtmlBefore, HtmlAfter As String


Public Sub Form_Initalize()
apidbox.text = ""
End Sub

Public Sub looper()
count = 0
Sheet3.Activate
  For i = 1 To Rows.count
        check = Cells(i, 1).Value
        If apidbox.text = check Then
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
        If apidbox.text = check Then
            quote = Cells(i, 20).text
        End If
    Next i
If count > 1 Then MsgBox "bad primary key"

End Sub

Public Sub CommandButton1_Click()
Sheet3.Activate
    If apidbox.text = vbNullString Then Call MsgBox("apidbox is null") Else looper
    Sheet2.Activate
    publisherSwitch (publisher)
    sendMail (txt)
    atest
    Unload Me
End Sub

Sub atest()
Dim myCells As Range
Set myCells = Selection
MsgBox myCells.Value
End Sub

Public Sub CommandButton2_Click()
Sheet3.Activate
    If apidbox.text = vbNullString Then Call MsgBox("apidbox is null") Else looper
    Sheet2.Activate
    custFacingPublisherSwitch (publisher)
    custFacingSendMail text
    
    Unload Me
End Sub


Public Sub custFacingSendMail(msg As String)
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Working in Office 2000-2016

    Dim OutApp As Object
    Dim OutMail As Object
    Dim str As String
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    actualMessage = msg
    If publisher = "VMware" Then
      OutMail.Attachments.Add "C:\Users\paul_disser\Documents\Quotes\" + quotesArr(0) + ".pdf"
      OutMail.Attachments.Add "C:\Users\paul_disser\Documents\Quotes\" + quotesArr(1) + ".pdf"
    Else
      OutMail.Attachments.Add "C:\Users\paul_disser\Documents\Quotes\" + quote + ".pdf"
    End If
    setHTML
    With OutMail
        .Recipients.Add email
        .Subject = publisher + " Renewal for " + accName
        .Display   'or use .Display
        .HTMLBody = HtmlBefore + actualMessage + HtmlAfter
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub



Public Sub sendMail(msg As String)
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Working in Office 2000-2016

    Dim OutApp As Object
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    actualMessage = msg
    setHTML
    With OutMail
        .Recipients.Add recip
        .Subject = publisher + " Renewal for " + accName
        .HTMLBody = HtmlBefore + actualMessage + HtmlAfter
        .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub






Public Sub custFacingPublisherSwitch(pub As String)
Select Case pub
    Case "VMware"
        quotesArr = Split(quote)
        
        text = "Good Morning,<br><br>" + vbNewLine + vbNewLine + _
            "I知 Paul, a member of your Software Renewals Team here at Dell. I am reaching out to you to let you know that your VMware subscription is due to expire. I have created quotes for you. <br><br>" + vbNewLine + vbNewLine + _
            "Your mandatory VMware Support subscription expires on " + expdate + ".<br><br>" + vbNewLine + vbNewLine + _
            "Quote # " + quotesArr(0) + " will reflect the price for the one year.<br><br>" + vbNewLine + vbNewLine + _
            "Quote # " + quotesArr(1) + " will reflect the price for the three year renewal. (3 year quotes come with a 12% discount built in)<br><br>" + vbNewLine + vbNewLine + _
            "If you want to maintain your licenses please respond, confirming the quote that you would like to have processed into an order and how you would like to place the order.<br><br>" + vbNewLine + vbNewLine + _
            "**Please note: to avoid late/reinstatement fees please respond before the expiration date listed.**<br><br>" + vbNewLine + vbNewLine + _
            "Are you spinning up more than 25 Virtual Machines? Check out VMware's Virtualization Optimization Assessment to make sure everything is running efficiently<br>" + vbNewLine + _
            "More info: https://www.vmware.com/assessment/voa<br><br>" + vbNewLine + vbNewLine + _
            "Thank you,"
    Case "Symantec"
        text = "Good Morning,<br><br>" + vbNewLine + vbNewLine + _
            "I知 Paul, a member of your Software Renewals Team here at Dell. We are reaching out to you to let you know that your " + publisher + " Protection is expiring. <br><br>" + vbNewLine + vbNewLine + _
            "This contract expires on " + expdate + ".<br><br>" + vbNewLine + vbNewLine + _
            "Quote # " + quote + " is aligned with your previous order.<br><br>" + vbNewLine & vbNewLine & _
            "If you want to maintain your security please respond confirming that you would like to have this quote processed into an order and how you would like to place the order.<br><br>" & vbNewLine & vbNewLine & _
            "**Please note: If you let your license expire, you will have to pay reinstatement fees and will no longer receive up to date protection.**<br>" + vbNewLine + _
            "More info: https://www.symantec.com/support-center/renewals/renewals-faq<br><br>" + vbNewLine + vbNewLine + _
            "Thank you,"
    Case "Trend Micro"
        text = "Good Morning,<br><br>" + vbNewLine + vbNewLine + _
            "I知 Paul, a member of your Software Renewals Team here at Dell. We are reaching out to you to let you know that your " + publisher + " Maintenance is expiring. <br><br>" + vbNewLine + vbNewLine + _
            "This contract expires on " + expdate + ".<br><br>" + vbNewLine + vbNewLine + _
            "Quote # " + quote + " is aligned with your previous order.<br><br>" + vbNewLine & vbNewLine & _
            "If you want to maintain your security please respond confirming that you would like to have this quote processed into an order and how you would like to place the order.<br><br>" & vbNewLine & vbNewLine & _
            "**Please note: When your Maintenance Agreement expires you will no longer receive up to date protection.**<br>" + vbNewLine + _
            "More info: http://docs.trendmicro.com/all/ent/pp/v2.1/en-us/pp_2.1_olh/maintenance_agreement.htm<br><br>" + vbNewLine & vbNewLine & _
            "Thank you,"
    Case "VERITAS"
        text = "Good Morning,<br><br>" + vbNewLine + vbNewLine + _
            "I知 Paul, a member of your Software Renewals Team here at Dell. We are reaching out to you to let you know that your " + publisher + " Maintenance is expiring. <br><br>" + vbNewLine + vbNewLine + _
            "This contract expires on " + expdate + ".<br><br>" + vbNewLine + vbNewLine + _
            "Quote # " + quote + " is aligned with your previous order.<br><br>" + vbNewLine & vbNewLine & _
            "If you want to maintain your licenses please respond confirming that you would like to have this quote processed into an order and how you would like to place the order.<br><br>" & vbNewLine & vbNewLine & _
            "**Please note: If you let your license expire, you will have to pay reinstatement fees.**<br>" + vbNewLine + _
            "More Info: http://info.veritas.com/Global-Enterprise-Renewals-Policy?cname=17Q2-EMEA-CHNL-EN/EN-CAMP-VSPEAK_201608&eid=2110&cid=<br><br>" + vbNewLine & vbNewLine & _
            "Thank you,"
    Case "Autodesk"
         text = "Good Morning,<br><br>" + vbNewLine + vbNewLine + _
            "I知 Paul, a member of your Software Renewals Team here at Dell. We are reaching out to you to let you know that your " + publisher + " Maintenance is expiring. <br><br>" + vbNewLine + vbNewLine + _
            "This contract expires on " + expdate + ".<br><br>" + vbNewLine + vbNewLine + _
            "Quote # " + quote + " is aligned with your previous order.<br><br>" + vbNewLine & vbNewLine & _
            "If you want to maintain your licenses please respond confirming that you would like to have this quote processed into an order and how you would like to place the order.<br><br>" & vbNewLine & vbNewLine & _
            "**Please note: If you let your maintenance expire, you will no longer receive updates, access to new/old releases, and technical support (subscriptions will lose all access).**<br>" + vbNewLine + _
            "More Info: https://knowledge.autodesk.com/customer-service/account-management/subscription-management/manage-contracts/renew-cancel/renew-maintenance-subscription<br><br>" + vbNewLine & vbNewLine & _
            "Thank you,"
    Case Else
        text = "Good Morning,<br><br>" + vbNewLine + vbNewLine + _
            "I知 Paul, a member of your Software Renewals Team here at Dell. We are reaching out to you to let you know that your " + publisher + " Maintenance is expiring. <br><br>" + vbNewLine + vbNewLine + _
            "This contract expires on " + expdate + ".<br><br>" + vbNewLine + vbNewLine + _
            "Quote # " + quote + " is aligned with your previous order.<br><br>" + vbNewLine & vbNewLine & _
            "If you want to maintain your licenses please respond confirming that you would like to have this quote processed into an order and how you would like to place the order.<br><br>" & vbNewLine & vbNewLine & _
            "**Please note: If you let your maintenance expire, you will no longer receive updates, access to new/old releases, and technical support **<br><br>" + vbNewLine + vbNewLine + _
            "Thank you,"
    
End Select
End Sub

Public Sub publisherSwitch(pub As String)
Select Case pub
    Case "VMware"
        txt = "Hello,<br><br>" + vbNewLine + vbNewLine + _
            "Could I get a quote for this?<br><br>" + vbNewLine + vbNewLine + _
            "Account name: " + accName & vbNewLine & _
            "<br>Cust num: " + custNum & vbNewLine & _
            "<br>Order num: " + orderNum & vbNewLine & _
            "<br>Email: " + email & vbNewLine & _
            "<br>Expiration Date: " + expdate + vbNewLine + _
            "<br>Contract: " + special + vbNewLine + vbNewLine + _
            "<br><br>Thank you,"
        recip = "Joshua_Pence@Dell.com"
    Case "Symantec"
          txt = "Hello,<br><br>" + vbNewLine + vbNewLine + _
            "Could I get a quote for this?<br><br>" + vbNewLine + vbNewLine + _
            "Account name: " + accName & vbNewLine & _
            "<br>Renewal pin: " + special + vbNewLine + _
            "<br>Cust num: " + custNum & vbNewLine & _
            "<br>Order num: " + orderNum & vbNewLine & _
            "<br>Email: " + email & vbNewLine & _
            "<br>Expiration Date: " + expdate + vbNewLine + vbNewLine + _
            "<br><br>Thank you,"
        recip = "symantec-dell@ingrammicro.com"
    Case "Autodesk"
        txt = "Hello,<br><br>" + vbNewLine + vbNewLine + _
            "Could I get a quote for this?<br><br>" + vbNewLine + vbNewLine + _
            "Account name: " + accName & vbNewLine & _
            "<br>Serial Number: " + special + vbNewLine + _
            "<br>Cust num: " + custNum & vbNewLine & _
            "<br>Order num: " + orderNum & vbNewLine & _
            "<br>Email: " + email & vbNewLine & _
            "<br>Expiration Date: " + expdate + vbNewLine + vbNewLine + _
            "<br><br>Thank you,"
        recip = "James_Lout@Dell.com"
    Case "Intel Security"
        txt = "Hello,<br><br>" + vbNewLine + vbNewLine + _
            "Could I get a quote for this?<br><br>" + vbNewLine + vbNewLine + _
            "Account name: " + accName & vbNewLine & _
            "<br>Grant Number: " + special + vbNewLine + _
            "<br>Cust num: " + custNum & vbNewLine & _
            "<br>Order num: " + orderNum & vbNewLine & _
            "<br>Email: " + email & vbNewLine & _
            "<br>Expiration Date: " + expdate + vbNewLine + vbNewLine + _
            "<br><br>Thank you,"
        recip = "IntelSecurity@techdata.com"
    Case "Trend Micro"
        txt = "Hello,<br><br>" + vbNewLine + vbNewLine + _
            "Could I get a quote for this?<br><br>" + vbNewLine + vbNewLine + _
            "Account name: " + accName & vbNewLine & _
            "<br>License Authorization Number: " + special + vbNewLine + _
            "<br>Cust num: " + custNum & vbNewLine & _
            "<br>Order num: " + orderNum & vbNewLine & _
            "<br>Email: " + email & vbNewLine & _
            "<br>Expiration Date: " + expdate + vbNewLine + vbNewLine + _
            "<br><br>Thank you,"
        recip = "Trend-Licensing@ingrammicro.com"
    Case "VERITAS"
        txt = "Hello,<br><br>" + vbNewLine + vbNewLine + _
            "Could I get a quote for this?<br><br>" + vbNewLine + vbNewLine + _
            "Account name: " + accName & vbNewLine & _
            "<br>Special: " + special + vbNewLine + _
            "<br>Cust num: " + custNum & vbNewLine & _
            "<br>Order num: " + orderNum & vbNewLine & _
            "<br>Email: " + email & vbNewLine & _
            "<br>Expiration Date: " + expdate + vbNewLine + vbNewLine + _
            "<br><br>Thank you,"
        recip = "Veritas-Dell@ingrammicro.com"
    Case Else
        txt = "Hello,<br><br>" + vbNewLine + vbNewLine + _
            "Could I get a quote for this?<br><br>" + vbNewLine + vbNewLine + _
            "Account name: " + accName & vbNewLine & _
            "<br>Special: " + special + vbNewLine + _
            "<br>Cust num: " + custNum & vbNewLine & _
            "<br>Order num: " + orderNum & vbNewLine & _
            "<br>Email: " + email & vbNewLine & _
            "<br>Expiration Date: " + expdate + vbNewLine + vbNewLine + _
            "<br><br>Thank you,"
        recip = Null
End Select
End Sub

Private Sub UserForm_Click()

End Sub



Private Sub setHTML()


HtmlBefore = "<html xmlns:v='urn:schemas-microsoft-com:vml' xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns:m='http://schemas.microsoft.com/office/2004/12/omml' xmlns='http://www.w3.org/TR/REC-html40'><head><meta http-equiv=Content-Type content='text/html; charset=us-ascii'><meta name=Generator content='Microsoft Word 15 (filtered medium)'><!--[if !mso]><style>v\:* {behavior:url(#default#VML);}" + _
"o\:* {behavior:url(#default#VML);}w\:* {behavior:url(#default#VML);}.shape {behavior:url(#default#VML);}</style><![endif]--><style><!--" + _
"/* Font Definitions */@font-face'{font-family:'Cambria Math';panose-1:2 4 5 3 5 4 6 3 2 4;}" + _
"@font-face{font-family:Calibri;panose-1:2 15 5 2 2 2 4 3 2 4;}@font-face{font-family:'Trebuchet MS';panose-1:2 11 6 3 2 2 2 2 2 4;}" + _
"/* Style Definitions */p.MsoNormal, li.MsoNormal, div.MsoNormal{margin:0in;margin-bottom:.0001pt;font-size:10.0pt;font-family:'Calibri',sans-serif;}" + _
"a:link, span.MsoHyperlink{mso-style-priority:99;color:#0563C1;text-decoration:underline;}" + _
"a:visited, span.MsoHyperlinkFollowed{mso-style-priority:99;color:#954F72;text-decoration:underline;}" + _
"span.EmailStyle17{mso-style-type:personal-compose;font-family:'Calibri',sans-serif;color:windowtext;}" + _
".MsoChpDefault{mso-style-type:export-only;font-size:10.0pt;font-family:'Calibri',sans-serif;}" + _
"@page WordSection1{size:8.5in 11.0in;margin:1.0in 1.0in 1.0in 1.0in;}div.WordSection1  {page:WordSection1;}--></style><!--[if gte mso 9]><xml>" + _
"<o:shapedefaults v:ext='edit' spidmax='1027' /></xml><![endif]--><!--[if gte mso 9]><xml><o:shapelayout v:ext='edit'><o:idmap v:ext='edit' data='1' />" + _
"</o:shapelayout></xml><![endif]--></head><body lang=EN-US link='#0563C1' vlink='#954F72'><div class=WordSection1><p class=MsoNormal><span style='font-size:11.0pt'>"

HtmlAfter = "</span><span style='font-size:11.0pt;font-family:'Times New Roman',serif'><o:p></o:p></span></p><p class=MsoNormal><span style='font-size:11.0pt'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><b><span style='font-size:9.0pt;font-family:'Trebuchet MS',sans-serif;color:#1F4E79'>Paul Disser,</span></b><b><span style='font-size:9.0pt;font-family:'Trebuchet MS',sans-serif;color:#1F497D'> </span></b><span style='font-size:9.0pt;font-family:'Trebuchet MS',sans-serif;color:#1F497D'>Partner<b> </b>Software Support Team<o:p></o:p></span></p><p class=MsoNormal><!--[if gte vml 1]><v:shapetype id='_x0000_t75' coordsize='21600,21600' o:spt='75' o:preferrelative='t' path='m@4@5l@4@11@9@11@9@5xe' filled='f' stroked='f'>" + _
"<v:stroke joinstyle='miter' /><v:formulas><v:f eqn='if lineDrawn pixelLineWidth 0' /><v:f eqn='sum @0 1 0' />" + _
"<v:f eqn='sum 0 0 @1' /><v:f eqn='prod @2 1 2' /><v:f eqn='prod @3 21600 pixelWidth' /><v:f eqn='prod @3 21600 pixelHeight' /><v:f eqn='sum @0 0 1' /><v:f eqn='prod @6 1 2' /><v:f eqn='prod @7 21600 pixelWidth' /><v:f eqn='sum @8 21600 0' /><v:f eqn='prod @7 21600 pixelHeight' /><v:f eqn='sum @10 21600 0' /></v:formulas>" + _
"<v:path o:extrusionok='f' gradientshapeok='t' o:connecttype='rect' /><o:lock v:ext='edit' aspectratio='t' /></v:shapetype><v:shape id='Picture_x0020_2' o:spid='_x0000_s1026' type='#_x0000_t75' alt='Description: Description: Description: cid:image001.gif@01CACC3B.A16F4080' style='position:absolute;margin-left:0;margin-top:0;width:30pt;height:30pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;mso-width-relative:page;mso-height-relative:page'><v:imagedata src='cid:image001.png@01D2456D.41941020' o:title='image001.gif@01CACC3B' /><w:wrap type='through'/>" + _
"</v:shape><![endif]--><![if !vml]><img width=40 height=40 src='cid:image003.jpg@01D2456D.BFBA86F0' align=left hspace=12 alt='Description: Description: Description: cid:image001.gif@01CACC3B.A16F4080' v:shapes='Picture_x0020_2'><![endif]><span style='font-size:8.0pt;font-family:'Trebuchet MS',sans-serif;color:#1F497D'>North America Partner Software Solution Sales<o:p></o:p></span></p><p class=MsoNormal><span lang=FR style='font-size:8.0pt;font-family:'Trebuchet MS',sans-serif;color:#1F497D'>512.513.8444 | </span><span style='font-size:11.0pt'><a href='mailto:Paul_Disser@Dell.com'><span lang=FR style='font-size:8.0pt;font-family:'Trebuchet MS',sans-serif;color:blue'>Paul_Disser@Dell.com</span></a></span><span style='font-size:8.0pt;font-family:'Trebuchet MS',sans-serif;color:#1F497D'> <span lang=FR><o:p></o:p></span></span></p><p class=MsoNormal><span style='font-size:11.0pt'><o:p>&nbsp;</o:p></span></p></div></body></html>"

End Sub
