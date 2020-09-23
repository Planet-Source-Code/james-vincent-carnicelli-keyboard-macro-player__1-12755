VERSION 5.00
Begin VB.Form AboutForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Keyboard Macro Player"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Contents 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "AboutForm.frx":0000
      Top             =   180
      Width           =   3570
   End
   Begin VB.CommandButton Ok 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   420
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Width           =   1140
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Ok_Click()
    Unload Me
End Sub
