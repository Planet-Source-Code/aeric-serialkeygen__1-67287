VERSION 5.00
Begin VB.Form frmKeyGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Key Generator"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   Icon            =   "frmKeyGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3945
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtResult 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate!"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmKeyGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is a Key Generator to 'crack' Serial Number Verification
Option Explicit

Private Sub cmdCopy_Click()
    'Copy generated Serial Number to Windows Clipboard
    Clipboard.SetText txtResult.Text
    cmdCopy.Enabled = False
End Sub

Private Sub cmdGenerate_Click()
    If Len(txtCode) = 6 Then
        txtResult = GenSerial(txtCode.Text)
        txtResult.Visible = True
        cmdCopy.Enabled = True
    Else
        MsgBox "Invalid length.", vbExclamation, App.CompanyName
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'MsgBox "Thank you for supporting.", vbInformation, App.Title
End Sub

Private Sub txtCode_Change()
    txtResult.Visible = False
    cmdGenerate.Enabled = True
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then
        MsgBox "Code must be Integer only.", vbExclamation, App.CompanyName
        SendKeys "{Home}+{End}"
    End If
End Sub
