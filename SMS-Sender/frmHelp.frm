VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "HELP"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHelp 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7380
      Left            =   90
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   105
      Width           =   4470
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
txtHelp.Text = "         Welcome to SMS Sender V-1.0" & vbCrLf & vbCrLf & _
"        SMS Sender V-1.0 is a utility which can be used to send SMS" & _
" to any mobile user.SMS Sender uses Microsoft Outlook to send the" & _
" SMS to respective Numbers." & vbCrLf & vbCrLf & _
"        SMS Sender V-1.0 enables You to compose and send the messages in more organised and formated way." & vbCrLf & _
"It reduces lot of time typing on Cell Phone and stress on fingers." & vbCrLf & vbCrLf & _
"Features Provided in this Version are " & vbCrLf & _
"1.Can send SMS to any mobile Users. " & vbCrLf & _
"2.You can Add,Edit,Remove the friends Name and Contact number from the " & vbCrLf & _
" List. " & vbCrLf & vbCrLf & _
"Features to be added in Coming version are" & vbCrLf & _
"1.To send SMS to Multiple recipients." & vbCrLf & _
"3.More attractive user interface and ease of Use with enhanced features." & vbCrLf & vbCrLf & _
"For any suggestions or bugs" & vbCrLf & vbCrLf & _
"mailto :prashanth.sc@itreya.com OR scprashi@rediffmail.com " & vbCrLf & vbCrLf & _
"Its my request don't make use of this tool for any illegal purpose." & vbCrLf & _
"Thank U" & vbCrLf & _
"Your Friend"
End Sub
