VERSION 5.00
Begin VB.Form InsertSymbol 
   Caption         =   "Insert Symbol"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ExtraStuff 
      BackColor       =   &H0080FFFF&
      Caption         =   "<blank line>"
      Height          =   285
      Index           =   6
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2970
      Width           =   1095
   End
   Begin VB.CommandButton ExtraStuff 
      BackColor       =   &H0080FFFF&
      Caption         =   "# Comment"
      Height          =   285
      Index           =   5
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2655
      Width           =   1095
   End
   Begin VB.CommandButton ExtraStuff 
      BackColor       =   &H0080FFFF&
      Caption         =   "End Ignore"
      Height          =   285
      Index           =   4
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton ExtraStuff 
      BackColor       =   &H0080FFFF&
      Caption         =   "Start Ignore"
      Height          =   285
      Index           =   3
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2025
      Width           =   1095
   End
   Begin VB.CommandButton ExtraStuff 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sleep 100ms"
      Height          =   285
      Index           =   2
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1710
      Width           =   1095
   End
   Begin VB.CommandButton ExtraStuff 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sleep 1s"
      Height          =   285
      Index           =   1
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1395
      Width           =   1095
   End
   Begin VB.CommandButton ExtraStuff 
      BackColor       =   &H0080FFFF&
      Caption         =   "Pause"
      Height          =   285
      Index           =   0
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton CtrlAltShift 
      BackColor       =   &H0080FF80&
      Caption         =   "Shift"
      Height          =   285
      Index           =   2
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "+"
      Top             =   765
      Width           =   1095
   End
   Begin VB.CommandButton CtrlAltShift 
      BackColor       =   &H0080FF80&
      Caption         =   "ALT"
      Height          =   285
      Index           =   1
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "%"
      Top             =   450
      Width           =   1095
   End
   Begin VB.CommandButton CtrlAltShift 
      BackColor       =   &H0080FF80&
      Caption         =   "CTRL"
      Height          =   285
      Index           =   0
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "^"
      Top             =   135
      Width           =   1095
   End
   Begin VB.CheckBox KeepOpen 
      Caption         =   "Keep Open"
      Height          =   240
      Left            =   225
      TabIndex        =   47
      Top             =   3645
      Width           =   1275
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5490
      TabIndex        =   0
      Top             =   3555
      Width           =   1545
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F16}"
      Height          =   285
      Index           =   35
      Left            =   5085
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1710
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F15}"
      Height          =   285
      Index           =   34
      Left            =   5085
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1395
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F14}"
      Height          =   285
      Index           =   33
      Left            =   5085
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F13}"
      Height          =   285
      Index           =   32
      Left            =   5085
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   765
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F12}"
      Height          =   285
      Index           =   31
      Left            =   5085
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   450
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F11}"
      Height          =   285
      Index           =   30
      Left            =   5085
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   135
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F10}"
      Height          =   285
      Index           =   29
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2970
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F9}"
      Height          =   285
      Index           =   28
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2655
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F8}"
      Height          =   285
      Index           =   27
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2340
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F7}"
      Height          =   285
      Index           =   26
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2025
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F6}"
      Height          =   285
      Index           =   25
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1710
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F5}"
      Height          =   285
      Index           =   24
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1395
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F4}"
      Height          =   285
      Index           =   23
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F3}"
      Height          =   285
      Index           =   22
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   765
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F2}"
      Height          =   285
      Index           =   21
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   450
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{F1}"
      Height          =   285
      Index           =   20
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   135
      Width           =   735
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{UP}"
      Height          =   285
      Index           =   19
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2970
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{TAB}"
      Height          =   285
      Index           =   18
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2655
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{SCROLLLOCK}"
      Height          =   285
      Index           =   17
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2340
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{RIGHT}"
      Height          =   285
      Index           =   16
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2025
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{PRTSC}"
      Height          =   285
      Index           =   15
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1710
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{PGUP}"
      Height          =   285
      Index           =   14
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1395
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{PGDN}"
      Height          =   285
      Index           =   13
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1080
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{NUMLOCK}"
      Height          =   285
      Index           =   12
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   765
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{LEFT}"
      Height          =   285
      Index           =   11
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   450
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{INSERT}"
      Height          =   285
      Index           =   10
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   135
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{HOME}"
      Height          =   285
      Index           =   9
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2970
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{HELP}"
      Height          =   285
      Index           =   8
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2655
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{ESC}"
      Height          =   285
      Index           =   7
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2340
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{ENTER}"
      Height          =   285
      Index           =   6
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2025
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{END}"
      Height          =   285
      Index           =   5
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1710
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{DOWN}"
      Height          =   285
      Index           =   4
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1395
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{DELETE}"
      Height          =   285
      Index           =   3
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{CAPSLOCK}"
      Height          =   285
      Index           =   2
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   765
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{BREAK}"
      Height          =   285
      Index           =   1
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   450
      Width           =   1950
   End
   Begin VB.CommandButton Symbol 
      BackColor       =   &H0080C0FF&
      Caption         =   "{BACKSPACE}"
      Height          =   285
      Index           =   0
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   135
      Width           =   1950
   End
End
Attribute VB_Name = "InsertSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
    Me.Hide
End Sub

Private Sub CtrlAltShift_Click(Index As Integer)
    Cancel.SetFocus
    MainForm.Script.SelText = CtrlAltShift(Index).Tag
    If KeepOpen.Value = 0 Then Me.Hide
End Sub

Private Sub ExtraStuff_Click(Index As Integer)
    Cancel.SetFocus
    If Left(ExtraStuff(Index).Caption, 1) = "<" Then
        MainForm.Script.SelText = vbCrLf
    Else
        MainForm.Script.SelText = ExtraStuff(Index).Caption & vbCrLf
    End If
    If KeepOpen.Value = 0 Then Me.Hide
End Sub

Private Sub Symbol_Click(Index As Integer)
    Cancel.SetFocus
    MainForm.Script.SelText = Symbol(Index).Caption
    If KeepOpen.Value = 0 Then Me.Hide
End Sub
