VERSION 5.00
Begin VB.Form frmAdd 
   Caption         =   "ADD CONTACT"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtIdentify 
      Enabled         =   0   'False
      Height          =   420
      Left            =   165
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1395
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CANCEL"
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
      Left            =   3210
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1350
      Width           =   1140
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
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
      Left            =   1995
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1350
      Width           =   1215
   End
   Begin VB.TextBox txtNumber 
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
      Left            =   2070
      MaxLength       =   10
      TabIndex        =   1
      Top             =   765
      Width           =   3060
   End
   Begin VB.TextBox txtName 
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
      Left            =   2055
      MaxLength       =   50
      TabIndex        =   0
      Top             =   255
      Width           =   3105
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE NUMBER :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   4
      Top             =   780
      Width           =   1935
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAME :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1185
      TabIndex        =   3
      Top             =   285
      Width           =   750
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()

End Sub

Private Sub Text2_Change()

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If txtName.Text <> "" And txtNumber.Text <> "" Then
        If Len(txtNumber.Text) < 10 Then
            MsgBox "Check the Number", vbExclamation, "Check"
            Exit Sub
        End If
        frmMain.fillList txtName.Text, txtNumber.Text, txtIdentify.Text
        Unload Me
    Else
        MsgBox "Enter Both Fields", vbExclamation, "ADD"
    End If
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)
On Error GoTo errHandler
    If (Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
    Exit Sub
errHandler:
    MsgBox "txtNumber_KeyPress " & Err.Number & "::" & Err.Description
End Sub

