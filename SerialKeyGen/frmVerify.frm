VERSION 5.00
Begin VB.Form frmVerify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "frmVerify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtChallengeCode 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Challenge Code"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtVerify 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Image imgOK 
      Height          =   480
      Left            =   3120
      Picture         =   "frmVerify.frx":030A
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblHidden 
      Caption         =   "This program is protected by encryption. Don't ever try to break it."
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Please appreciate and respect my afford. Thank you."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   $"frmVerify.frx":0FD4
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    txtChallengeCode.Text = GenKey
End Sub

Private Sub txtVerify_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtVerify.Text = GenSerial(txtChallengeCode.Text) Then
            imgOK.Visible = True ' Show icon
            txtVerify.Enabled = False 'Disable user input textbox
            Label2.Caption = "Registration successful!" & vbCrLf & "Thank you for using this program."
            'MsgBox "Thank you for using this program.", vbInformation, "Registration successful"
            'Me.Hide ' Hide this form
            'mdiMain.Show ' Show Program form
        Else
            MsgBox "Incorrect code.", vbCritical, App.ProductName
            ' You can use a counter to quit this program
            ' when wrong code is entered 3 times repeatedly
        End If
    End If
End Sub
