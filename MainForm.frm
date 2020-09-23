VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainForm 
   Caption         =   "<untitled> - KMP"
   ClientHeight    =   4275
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2850
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   2850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Continue 
      BackColor       =   &H0080C0FF&
      Caption         =   "Continue"
      Height          =   285
      Left            =   945
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Timer SleepTimer 
      Enabled         =   0   'False
      Left            =   90
      Top             =   3825
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   585
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Keyboard Macro Files (*.kmp)|*.kmp|All Files (*.*)|*.*"
   End
   Begin VB.TextBox Script 
      Height          =   3615
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   315
      Width           =   2805
   End
   Begin VB.CommandButton Execute 
      Caption         =   "E&xecute"
      Height          =   285
      Left            =   1890
      TabIndex        =   1
      Top             =   3960
      Width           =   915
   End
   Begin VB.Label Status 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Editing"
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2805
   End
   Begin VB.Menu Menu_File 
      Caption         =   "&File"
      Begin VB.Menu Menu_File_New 
         Caption         =   "&New"
      End
      Begin VB.Menu Menu_File_Open 
         Caption         =   "&Open..."
      End
      Begin VB.Menu Menu_File_Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu Menu_File_SaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu Menu_File_Separator_1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_File_Reexecute 
         Caption         =   "Stop and &Reexecute"
      End
      Begin VB.Menu Menu_File_Separator_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_File_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Menu_Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Menu_Edit_InsertSymbol 
         Caption         =   "&Insert Symbol..."
      End
   End
   Begin VB.Menu Menu_Help 
      Caption         =   "&Help"
      Begin VB.Menu Menu_Help_Contents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu Menu_Help_About 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Executing As Boolean
Private FileName As String
Private FileChanged As Boolean
Private Paused As Boolean

Private Sub Continue_Click()
    Paused = False
    Continue.Visible = False
End Sub

Private Sub Execute_Click()
    Dim Lines, Line, StartTime As Date, Milliseconds As Long
    Dim Count, Ignore As Boolean
    If Executing Then
        Executing = False
        SleepTimer.Enabled = False
        DoEvents
        Execute.Caption = "E&xecute"
        Status.Caption = "Editing"
    Else
        InsertSymbol.Hide
        Execute.Caption = "&Stop"
        Executing = True
        Script.SetFocus
        DoEvents
        SendKeys "%{TAB}"
        Sleep 100
        
        Lines = Split(Script.Text, vbCrLf)
        Do
            For Each Line In Lines
                Count = Count + 1
                Status.Caption = "Executing line " & Count & " of " & (UBound(Lines) + 1)
                
                DoEvents
                If Not Executing Then Exit Do
                
                If Line = "" Then
                    'Blank line
                
                ElseIf Left(Line, 1) = "#" And Len(Line) > 1 Then
                    'Comment
                
                ElseIf UCase(Line) = "PAUSE" Then
                    Paused = True
                    Continue.Visible = True
                    Do Until Not Paused Or Not Executing
                        DoEvents
                    Loop
                    Sleep 100
                    If Executing Then
                        SendKeys "%{TAB}"
                        DoEvents
                    End If
                
                ElseIf UCase(Line) = "START IGNORE" Then
                    Ignore = True
                
                ElseIf UCase(Line) = "END IGNORE" Then
                    Ignore = False
                
                ElseIf UCase(Left(Line, 6)) = "SLEEP " Then
                    Line = Mid(Line, 7)
                    If UCase(Right(Line, 2)) = "MS" Then
                        'Milliseconds
                        Line = Trim(Left(Line, Len(Line) - 2))
                        Milliseconds = Line
                        Status.Caption = "Executing line " & Count & " of " & (UBound(Lines) + 1) & " (sleep " & Line & " s)"
                    ElseIf UCase(Right(Line, 1)) = "S" Then
                        'Seconds
                        Line = Trim(Left(Line, Len(Line) - 1))
                        Milliseconds = Line * 1000
                        Status.Caption = "Executing line " & Count & " of " & (UBound(Lines) + 1) & " (sleep " & Line & " s)"
                    Else
                        'Seconds
                        Line = Trim(Line)
                        Milliseconds = Line * 1000
                        Status.Caption = "Executing line " & Count & " of " & (UBound(Lines) + 1) & " (sleep " & Line & " s)"
                    End If
                    Sleep Milliseconds
                
                Else
                    SendKeys Line
                
                End If
            Next
            Exit Do
        Loop
        
        If Executing Then
            Execute_Click
        End If
        Continue.Visible = False
    End If
End Sub

Private Sub Sleep(Milliseconds As Long)
    If Milliseconds < 10 Then Milliseconds = 10
    SleepTimer.Interval = Milliseconds
    SleepTimer.Enabled = True
    Do Until Not SleepTimer.Enabled Or Not Executing
        DoEvents
    Loop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload InsertSymbol
    If Execute.Caption <> "E&xecute" Then
        Execute.Value = True
        Do Until Execute.Caption = "E&xecute"
            DoEvents
        Loop
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    On Error Resume Next
    Execute.Left = Me.ScaleWidth - Execute.Width - Script.Left
    Execute.Top = Me.ScaleHeight - Execute.Height - Script.Left
    Script.Width = Me.ScaleWidth - 2 * Script.Left
    Script.Height = Execute.Top - Script.Top - 2 * Script.Left
End Sub

Private Sub Menu_Edit_InsertSymbol_Click()
    InsertSymbol.Show
End Sub

Private Sub Menu_File_Exit_Click()
    SendKeys "%{F4}"
End Sub

Private Sub Menu_File_New_Click()
    FileName = ""
    Me.Caption = "<untitled> - KMP"
    Script.Text = ""
End Sub

Private Sub Menu_File_Open_Click()
    Dim Ret As VbMsgBoxResult
    
    'Save if needed
    If FileChanged Then
        Ret = MsgBox("File has changed.  Want to save?", vbYesNoCancel)
        If Ret = vbCancel Then
            Exit Sub
        ElseIf Ret = vbYes Then
            Menu_File_Save_Click
            If FileChanged Then  'User cancelled save
                Exit Sub
            End If
        End If
    End If
    
    On Error Resume Next
    FileDialog.ShowOpen
    If Err.Description = "Cancel was selected." Then
        Exit Sub
    Else
        On Error GoTo 0
        FileName = FileDialog.FileName
        Me.Caption = FileDialog.FileTitle & " - KMP"
        Script.Text = ReadFile(FileName)
    End If
End Sub

Private Sub Menu_File_Reexecute_Click()
    If Executing Then
        Execute.Value = True
        DoEvents
    End If
    Execute.Value = True
End Sub

Private Sub Menu_File_Save_Click()
    If FileName = "" Then
        Menu_File_SaveAs_Click
    Else
        WriteFile FileName, Script.Text
        FileChanged = False
    End If
End Sub

Private Sub Menu_File_SaveAs_Click()
    On Error Resume Next
    FileDialog.ShowSave
    If Err.Description = "Cancel was selected." Then
        Exit Sub
    Else
        On Error GoTo 0
        FileName = FileDialog.FileName
        Me.Caption = FileDialog.FileTitle & " - KMP"
        Menu_File_Save_Click
    End If
End Sub

Private Sub Menu_Help_About_Click()
    AboutForm.Show vbModal
End Sub

Private Sub Menu_Help_Contents_Click()
    Shell "Notepad.exe " & App.Path & "\help.txt", vbNormalFocus
End Sub

Private Sub SleepTimer_Timer()
    SleepTimer.Enabled = False
End Sub

'Read an entire text file into and return a string
Public Function ReadFile(ByVal FileName As String) As String
    Dim fh As Integer
    fh = FreeFile
    Open FileName For Binary As #fh
    ReadFile = Input$(LOF(fh), fh)
    Close #fh
End Function

'Write the contents of the string to a text file
Public Sub WriteFile(ByVal FileName As String, ByVal Contents As String)
    Dim fh As Integer
    fh = FreeFile
    Open FileName For Output As #fh
    Print #fh, Contents;
    Close #fh
End Sub

'Determine if the given file already exists
Public Function FileExists(ByVal FileName As String) As Boolean
    FileExists = (Dir(FileName) <> "")
End Function
