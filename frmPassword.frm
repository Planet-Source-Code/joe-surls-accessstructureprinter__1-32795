VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password?"
   ClientHeight    =   1065
   ClientLeft      =   3525
   ClientTop       =   2070
   ClientWidth     =   3360
   ControlBox      =   0   'False
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPassword.frx":030A
   ScaleHeight     =   1065
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   225
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   525
      Width           =   2790
   End
   Begin VB.Label Label1 
      Caption         =   "This Database requires a password. Please enter below..."
      Height          =   390
      Left            =   600
      TabIndex        =   1
      Top             =   75
      Width           =   2565
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pstrPassword As String
Public pblnCancel As Boolean

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Not Trim(txtPassword) = "" Then
        pstrPassword = Trim(txtPassword)
        txtPassword = ""
        pblnCancel = False
        Me.Hide
    End If
    If KeyAscii = vbKeyEscape Then
        pstrPassword = ""
        txtPassword = ""
        pblnCancel = True
        Unload Me
    End If
End Sub
