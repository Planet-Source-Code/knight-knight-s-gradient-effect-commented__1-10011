VERSION 5.00
Begin VB.Form K_FrmDither 
   Caption         =   "Knight's Form Dithering - [Main Window]"
   ClientHeight    =   3600
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5160
   DrawStyle       =   6  'Inside Solid
   Icon            =   "K_FrmDither.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraSelect 
      Caption         =   "Select Colour:"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Select The Colour You Want To Dither To Black"
      Top             =   120
      Width           =   1695
      Begin VB.CheckBox ChkAutoReDraw 
         Caption         =   "Auto Re-Draw"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Auto Re-Draw Form Slower But Smoother"
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton OptColour 
         Caption         =   "Red"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   5
         ToolTipText     =   "Dither Form Red"
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton OptColour 
         Caption         =   "Green"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Dither Form Green"
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton OptColour 
         Caption         =   "Blue"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   4
         ToolTipText     =   "Dither Form Blue"
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton OptColour 
         Caption         =   "White"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Dither Form White"
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton OptColour 
         Caption         =   "Random"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Dither Form Random"
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "K_FrmDither"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' Detect Non-Dimmed Variables

Private Function Dither(FrmDither As Form, Colour As String)
On Error Resume Next
    ' Declare IntLoop Variable As Integer
    Dim IntLoop As Integer
    ' Change Critical FrmDither Property to Draw Dithered Form
    FrmDither.DrawStyle = vbInsideSolid ' 6
    FrmDither.DrawMode = vbCopyPen ' 13
    FrmDither.ScaleMode = vbPixels ' 3
    FrmDither.DrawWidth = 2 ' Must Be > 1 to Work Properly
    FrmDither.ScaleHeight = 256 ' Maximum Value Number of Colours
    ' Draw Line The Lines Using IntLoop
    For IntLoop = 0 To 255 ' Draw A Colour Line For Each Value to 255
        DoEvents ' Let Windows Process Other Thingz
        ' Check Color & Draw Line
        Select Case Colour
            Case "Red"
                ' Red
                FrmDither.Line (0, IntLoop)-(Screen.Width, IntLoop - 1), RGB(255 - IntLoop, 0, 0), B
            Case "Green"
                ' Green
                FrmDither.Line (0, IntLoop)-(Screen.Width, IntLoop - 1), RGB(0, 255 - IntLoop, 0), B
            Case "Blue"
                ' Blue
                FrmDither.Line (0, IntLoop)-(Screen.Width, IntLoop - 1), RGB(0, 0, 255 - IntLoop), B
            Case "White"
                ' White
                FrmDither.Line (0, IntLoop)-(Screen.Width, IntLoop - 1), RGB(255 - IntLoop, 255 - IntLoop, 255 - IntLoop), B
            Case "Random"
                ' Random
                FrmDither.Line (0, IntLoop)-(Screen.Width, IntLoop - 1), RGB(Int(Rnd * IntLoop), Int(Rnd * IntLoop), Int(Rnd * IntLoop)), B
        End Select
        DoEvents ' Let Windows Process Other Thingz
    Next IntLoop ' Done Drawing Colour Lines for Each Colour
End Function

Private Sub ChkAutoReDraw_Click()
On Error Resume Next
    ' Enabled/Disabled Form Auto Redraw
    Me.AutoRedraw = ChkAutoReDraw.Value
End Sub

Private Sub Form_Click()
On Error Resume Next
    ' Hide/Show FraSelect
    FraSelect.Visible = Not FraSelect.Visible
    Call ChkSelectFra
End Sub

Private Sub Form_Load()
On Error Resume Next
    ' Check Selected Option & Call DitherForm
    Me.Caption = "Knight's Form Dithering - [Dithering Form...]"
    FraSelect.Enabled = False
    If OptColour(0).Value = True Then
        ' Red
        Call Dither(Me, "Red")
    ElseIf OptColour(1).Value = True Then
        ' Green
        Call Dither(Me, "Green")
    ElseIf OptColour(2).Value = True Then
        ' Blue
        Call Dither(Me, "Blue")
    ElseIf OptColour(3).Value = True Then
        ' White
        Call Dither(Me, "Write")
    ElseIf OptColour(4).Value = True Then
        ' Random
        Call Dither(Me, "Random")
    End If
    FraSelect.Enabled = True
    Call ChkSelectFra
End Sub

Private Sub Form_Paint()
On Error Resume Next
    ' Re-Draw Dithering
    Call Form_Load
End Sub

Private Sub Form_Resize()
On Error Resume Next
    ' Re-Center Form Window
    Call Me.Move((Screen.Width / 2) - (Me.Width / 2), (Screen.Height / 2) - (Me.Height / 2))
    ' Re-Draw Dithering
    Call Form_Load
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    ' Display About Program & Vote MsgBox
    Call MsgBox(App.Title & vbNewLine & "Version: " & App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "0000") & vbNewLine & vbNewLine & "Please Remember To Vote For Me! Thank You." & vbNewLine & vbNewLine & "By: " & App.CompanyName & vbNewLine & App.LegalCopyright, vbInformation, App.Title & " - About Message Box")
End Sub

Private Sub OptColour_Click(Index As Integer)
On Error Resume Next
    Me.Caption = "Knight's Form Dithering - [Dithering Form...]"
    FraSelect.Enabled = False
    ' Find Colour Chosen
    Select Case Index
        Case "0" ' Red
            Call Dither(Me, "Red")
        Case "1" ' Green
            Call Dither(Me, "Green")
        Case "2" ' Blue
            Call Dither(Me, "Blue")
        Case "3" ' White
            Call Dither(Me, "White")
        Case "4" ' Random
            Call Dither(Me, "Random")
    End Select
    FraSelect.Enabled = True
    Call ChkSelectFra
End Sub

Private Function ChkSelectFra()
On Error Resume Next
    ' Select Correct Caption
    If FraSelect.Visible = True Then
        Me.Caption = "Knight's Form Dithering - [Click to Hide Frame]"
    Else
        Me.Caption = "Knight's Form Dithering - [Click to Show Frame]"
    End If
End Function
