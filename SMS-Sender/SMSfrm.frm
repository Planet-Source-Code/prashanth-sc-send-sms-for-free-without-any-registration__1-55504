VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   Caption         =   "SMS-Sender V-1.0"
   ClientHeight    =   2850
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   10215
   FillStyle       =   3  'Vertical Line
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2850
   ScaleMode       =   0  'User
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPut 
      Caption         =   "<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1095
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&ADD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6285
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2100
      Width           =   1185
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&REMOVE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8670
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2100
      Width           =   1185
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&EDIT"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7485
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2100
      Width           =   1185
   End
   Begin VB.ComboBox cboNumber 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2895
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   255
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&CLEAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2145
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2130
      Width           =   1155
   End
   Begin VB.TextBox txtMessage 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   2895
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   660
      Width           =   2955
   End
   Begin VB.TextBox txtMobileNumber 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3900
      MaxLength       =   6
      TabIndex        =   4
      Top             =   255
      Width           =   1950
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3285
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2130
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&SEND"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   930
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2130
      Width           =   1215
   End
   Begin VB.Timer timeTimer 
      Interval        =   1000
      Left            =   2925
      Top             =   5175
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   6465
      Top             =   5460
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   5895
      Top             =   5460
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComctlLib.ListView lstAddress 
      Height          =   1680
      Left            =   6645
      TabIndex        =   14
      Top             =   270
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   2963
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MobileNumber"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Line Line3 
      X1              =   5925
      X2              =   5925
      Y1              =   2565
      Y2              =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "        Mobile Number:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   420
      TabIndex        =   9
      Top             =   270
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Message:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   330
      TabIndex        =   8
      Top             =   675
      Width           =   2310
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10350
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7455
      TabIndex        =   1
      Top             =   2595
      Width           =   60
   End
   Begin VB.Label lblOperationNow 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   2610
      Width           =   720
   End
   Begin VB.Line Line2 
      X1              =   -150
      X2              =   10200
      Y1              =   2565
      Y2              =   2565
   End
   Begin VB.Menu IdHelp 
      Caption         =   "&Help"
      Begin VB.Menu IdContents 
         Caption         =   "Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu IdAbout 
         Caption         =   "About"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu idSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu IDExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/*Code by Prashanth SC*/
'This Form takes care of finding the id and sending it to the particular address specified
Option Explicit
Dim objAddress As New Scripting.FileSystemObject
Dim objRead As TextStream
Dim objWrite As TextStream
Dim blnDataChanged As Boolean
Dim g_NoTimes As Integer
Dim blnTrue As Boolean

Private Sub cmdAdd_Click()
    frmAdd.txtIdentify.Text = "0"
    frmAdd.Show vbModal
End Sub

Private Sub cmdClear_Click()
    txtMessage.Text = ""
End Sub

Private Sub cmdEdit_Click()
    Dim strMobileNumber As String
    Dim strName As String
    Dim intCount As Integer
    Dim blnFlag As Boolean
    For intCount = 1 To lstAddress.ListItems.Count
        If lstAddress.ListItems.Item(intCount).Selected = True Then
            strName = lstAddress.ListItems.Item(intCount).Text
            strMobileNumber = lstAddress.ListItems.Item(intCount).SubItems(1)
            Exit For
        End If
    Next
    frmAdd.txtName.Text = strName
    frmAdd.txtNumber.Text = strMobileNumber
    frmAdd.txtIdentify.Text = CStr(intCount)
    frmAdd.Show vbModal
End Sub

Private Sub cmdExit_Click()
    Dim strResult As String
    If blnDataChanged Then
        strResult = MsgBox("Want to Save Chages made to Contacts list ?", vbYesNoCancel, "SAVE")
        blnTrue = True
        If strResult = vbYes Then
            saveContactList
            Unload Me
        ElseIf strResult = vbNo Then
            Unload Me
        Else
            Exit Sub
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdPut_Click()
    Dim strMobileNumber As String
    Dim intCount As Integer
    Dim blnFlag As Boolean
    For intCount = 1 To lstAddress.ListItems.Count
        If lstAddress.ListItems.Item(intCount).Selected = True Then
            strMobileNumber = lstAddress.ListItems.Item(intCount).SubItems(1)
            Exit For
        End If
    Next
    blnFlag = False
    For intCount = 0 To cboNumber.ListCount
        If cboNumber.List(intCount) = Left$(strMobileNumber, 4) Then
            cboNumber.Text = Left$(strMobileNumber, 4)
            blnFlag = True
            Exit For
        End If
    Next
    If blnFlag Then
        txtMobileNumber.Text = Right$(strMobileNumber, 6)
        txtMessage.SetFocus
    Else
        MsgBox "SMS Sender does not support this Service Provider!", vbExclamation, "Failed"
        Exit Sub
    End If
End Sub

Private Sub cmdRemove_Click()
    Dim intCount As Integer
    For intCount = 1 To lstAddress.ListItems.Count
        If lstAddress.ListItems.Item(intCount).Selected = True Then
            If MsgBox("Are you Sure to Remove " + lstAddress.ListItems.Item(intCount).Text, vbYesNo, "REMOVE") = vbYes Then
                lblOperationNow.Caption = "Removed " + lstAddress.ListItems.Item(intCount).Text + "from the Contacts"
                lstAddress.ListItems.Remove (intCount)
                blnDataChanged = True
                Exit For
            End If
        End If
    Next
End Sub

Private Sub cmdSend_Click()
    On Error GoTo errHandler
    Dim strMailID As String
    If cboNumber.Text <> "" And txtMobileNumber <> "" Then
        If Len(txtMobileNumber.Text) < 6 Then
            MsgBox "Check the Number Once Again", vbInformation, "Check"
            Exit Sub
        End If
        If txtMessage.Text <> "" Then
            SendMail
            g_NoTimes = g_NoTimes + 1
        Else
            MsgBox "Please Enter the Message !"
        End If
    Else
        MsgBox "Please Enter the Number !", vbInformation, "Number"
        txtMobileNumber.SetFocus
    End If
    Exit Sub
errHandler:
End Sub

Private Sub Form_Load()
    Dim strText As String
    Dim ItemX As ListItem
    timeTimer.Enabled = True
    timeTimer.Interval = 1000
    lblOperationNow.Caption = "Welcome to SMS Sender"
    lstAddress.ColumnHeaders(1).Width = lstAddress.Width / 2
    lstAddress.ColumnHeaders(2).Width = lstAddress.Width / 2
    Set objRead = objAddress.OpenTextFile(App.Path + "\Address.txt", ForReading, False)
    While Not objRead.AtEndOfStream
        strText = objRead.ReadLine
        Set ItemX = lstAddress.ListItems.Add(, , Split(strText, "::")(0))
        ItemX.SubItems(1) = Split(strText, "::")(1)
    Wend
    objRead.Close
    loadCboBox
End Sub
Private Sub loadCboBox()
    With cboNumber
        .AddItem "9810", 0
        .AddItem "9812", 1
        .AddItem "9815", 2
        .AddItem "9816", 3
        .AddItem "9818", 4
        .AddItem "9820", 5
        .AddItem "9821", 6
        .AddItem "9822", 7
        .AddItem "9823", 8
        .AddItem "9824", 9
        .AddItem "9825", 10
        .AddItem "9831", 11
        .AddItem "9837", 12
        .AddItem "9840", 13
        .AddItem "9841", 14
        .AddItem "9842", 15
        .AddItem "9843", 16
        .AddItem "9844", 17
        .AddItem "9845", 18
        .AddItem "9846", 19
        .AddItem "9847", 20
        .AddItem "9848", 21
        .AddItem "9849", 22
        .AddItem "9880", 23
        .AddItem "9885", 24
        .AddItem "9886", 25
        .AddItem "9890", 26
        .AddItem "9892", 27
        .AddItem "9893", 28
        .AddItem "9894", 29
        .AddItem "9895", 30
        .AddItem "9896", 31
        .AddItem "9898", 32
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blnTrue = False Then
        cmdExit_Click
    End If
End Sub

Private Sub IdAbout_Click()
    frmAbout.lblVersion = "        SMS-Sender Version-1.0" & vbCrLf & _
                         "                         Prashanth SC"
    frmAbout.Show
End Sub

Private Sub IdContents_Click()
   frmHelp.Show
End Sub

Private Sub IDExit_Click()
    Unload Me
End Sub

Private Sub lstAddress_Click()
    cmdPut.Enabled = True
    cmdEdit.Enabled = True
    cmdRemove.Enabled = True
    lblOperationNow.Caption = "No of Messages Sent : " + CStr(g_NoTimes)
End Sub

Private Sub timeTimer_Timer()
    lblTime.Caption = CStr(Now)
End Sub

Private Sub txtMobileNumber_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If (Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
    Exit Sub
errHandler:
    MsgBox "txtMobileNumber->KeyPress " & Err.Number & "::" & Err.Description
End Sub


Private Function getMailID() As String
    On Error GoTo errHandler
    Dim strTemp As String
    Dim strMailID As String
    If cboNumber.Text <> "" Then
        strTemp = cboNumber.Text
        Select Case strTemp
        Case 9810, 9895, 9896, 9898, 9890, 9894, 9892, 9893, 9815, 9816, 9818, 9840:
                strMailID = "91" + strTemp + txtMobileNumber.Text + "@airtelmail.com"
                lblOperationNow.Caption = "Service by Airtel"
        Case 9894:
                strMailID = "91" + strTemp + txtMobileNumber.Text + "@airteltn.com"
                lblOperationNow.Caption = "Service by Airtel"
        Case 9831:
                strMailID = "91" + strTemp + txtMobileNumber.Text + "@airtelkol.com"
                lblOperationNow.Caption = "Service by Airtel"
        Case 9840:
                strMailID = "91" + strTemp + txtMobileNumber.Text + "@airtelchennai.com"
                lblOperationNow.Caption = "Service by Airtel"
        Case 9849:
                strMailID = "91" + strTemp + txtMobileNumber.Text + "@airtelap.com"
                lblOperationNow.Caption = "Service by Airtel"
        Case 9845, 9880:
                strMailID = "91" + strTemp + txtMobileNumber.Text + "@airtelkk.com"
                lblOperationNow.Caption = "Service by Airtel"
                '****************AirTel*************
        Case 9821, 9823, 9843:
                strMailID = "91" + strTemp + txtMobileNumber.Text + "@bplmobile.com"
                lblOperationNow.Caption = "Service by BPL Mobile"
                '****************bplmobile**********
        Case 9837, 9812, 9847:
                strMailID = strTemp + txtMobileNumber.Text + "@escotelmobile.com"
                lblOperationNow.Caption = "Service by Escotel Mobile"
                '***************escotelmobile*******
        Case 9820:
                strMailID = strTemp + txtMobileNumber.Text + "@orangemail.co.in"
                lblOperationNow.Caption = "Service by Orange"
                '****************orangemail*********
        Case 9822, 9848:
                strMailID = strTemp + txtMobileNumber.Text + "@ideacellular.net"
                lblOperationNow.Caption = "Service by IdeaCellular"
                '***************ideacellular********
        Case 9825:
                strMailID = strTemp + txtMobileNumber.Text + "@celforce.com"
                lblOperationNow.Caption = "Service by CelForce"
                '***************celforce************
        Case 9841:
                strMailID = strTemp + txtMobileNumber.Text + "@rpgmail.net"
                lblOperationNow.Caption = "Service by RPG"
                '***************rpgmail*************
        Case 9842:
                strMailID = strTemp + txtMobileNumber.Text + "@airsms.com"
                lblOperationNow.Caption = "Service by AirSMS"
                '*****************airsms************
        Case 9885, 9886, 9884:
                strMailID = strTemp + txtMobileNumber.Text + "@south.hutch.co.in"
                lblOperationNow.Caption = "Service by HUTCH"
        Case 9811:
                strMailID = strTemp + txtMobileNumber.Text + "@delhi.hutch.co.in"
                lblOperationNow.Caption = "Service by HUTCH"
                '****************hutch**************
        Case 9844:
                strMailID = strTemp + txtMobileNumber.Text + "@mobile.spicetele.com"
                lblOperationNow.Caption = "Service by SPICE"
        Case Else:
                strMailID = strTemp + txtMobileNumber.Text + "@airtelmail.com  "
                lblOperationNow.Caption = "Service by Airtel"
        End Select
    End If
    getMailID = strMailID
    Exit Function
errHandler:
    MsgBox "Error" & Err.Description
End Function
Private Sub SendMail()
    On Error GoTo errHandler
    Dim strMailID As String
    strMailID = getMailID
    MAPISession1.SignOn
    With MAPIMessages1
        .SessionID = MAPISession1.SessionID
        .Compose
        .RecipDisplayName = strMailID
        .RecipAddress = strMailID
        .MsgSubject = "Hai"
        .MsgNoteText = txtMessage.Text
        .Send
    End With
    MAPISession1.SignOff
    lblOperationNow.Caption = "Message sent to  " + cboNumber.Text + txtMobileNumber.Text
    Exit Sub
errHandler:
    MsgBox "Error" + Err.Description
End Sub
Public Sub fillList(a_Name As String, a_Number As String, a_ID As String)
    Dim ItemX As ListItem
    If a_ID = "0" Then
        Set ItemX = lstAddress.ListItems.Add(, , a_Name)
        ItemX.SubItems(1) = a_Number
        blnDataChanged = True
        lblOperationNow.Caption = "New Contact Added " + a_Name
    Else
        lstAddress.ListItems.Remove CInt(a_ID)
        Set ItemX = lstAddress.ListItems.Add(a_ID, , a_Name)
        ItemX.SubItems(1) = a_Number
        blnDataChanged = True
        lblOperationNow.Caption = "Contact Edited Successfully :" + a_Name
    End If
End Sub
Public Sub saveContactList()
    Dim intListCount As Integer
    Dim strSelecteditem As String
    Set objWrite = objAddress.OpenTextFile(App.Path + "\address.txt", ForWriting, True)
    For intListCount = 1 To lstAddress.ListItems.Count
        strSelecteditem = lstAddress.ListItems.Item(intListCount).Text
        strSelecteditem = strSelecteditem + "::" + lstAddress.ListItems.Item(intListCount).SubItems(1)
        objWrite.WriteLine strSelecteditem
    Next
    objWrite.Close
End Sub
