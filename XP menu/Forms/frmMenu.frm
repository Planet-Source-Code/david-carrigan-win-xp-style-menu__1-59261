VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Splitter 
      Index           =   0
      Visible         =   0   'False
      X1              =   480
      X2              =   2640
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblMenuItem 
      BackStyle       =   0  'Transparent
      Caption         =   "MENU ITEM"
      Height          =   255
      Index           =   0
      Left            =   430
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image imgHighlight 
      Height          =   300
      Left            =   30
      Picture         =   "frmMenu.frx":0000
      Top             =   30
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image imgMenu 
      Height          =   195
      Index           =   2
      Left            =   0
      Picture         =   "frmMenu.frx":2A72
      Top             =   170
      Width           =   2760
   End
   Begin VB.Image imgMenu 
      Height          =   195
      Index           =   0
      Left            =   0
      Picture         =   "frmMenu.frx":46BC
      Top             =   0
      Width           =   2760
   End
   Begin VB.Image imgMenu 
      Height          =   5925
      Index           =   1
      Left            =   0
      Picture         =   "frmMenu.frx":6306
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intMenuItems As Integer
Dim intSplitterItems As Integer

Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const CS_DROPSHADOW = &H20000
Private Const GCL_STYLE = (-26)
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000


Private Sub Form_Load()
        ' Add your Menu things here. or how ever you feel
        ' by adding "//Splitter//" will add a line instead
        AddMenu "Test Menu"
        AddMenu "Test Menu"
        AddMenu "Test Menu"
        AddMenu "//Splitter//"
        AddMenu "Test Menu"
        AddMenu "Test Menu"
        AddMenu "Test Menu"
        
    FadeIn Me, Me.hwnd, 255
    DropShadow Me.hwnd
End Sub

Function AddMenu(strMenuCaption As String)
    If strMenuCaption = "//Splitter//" Then
        intSplitterItems = intSplitterItems + 1
        lblMenuItem(intMenuItems).Tag = "SplitterAfter"
        Load Splitter(intSplitterItems)
        With Splitter(intSplitterItems)
            .Y1 = lblMenuItem(intMenuItems).Top + lblMenuItem(intMenuItems).Height + 10
            .Y2 = lblMenuItem(intMenuItems).Top + lblMenuItem(intMenuItems).Height + 10
            .BorderColor = RGB(172, 168, 153)
            .Visible = True
            .ZOrder vbBringToFront
        End With
        Exit Function
    End If
    intMenuItems = intMenuItems + 1
    Load lblMenuItem(intMenuItems)
    With lblMenuItem(intMenuItems)
        .Caption = strMenuCaption
        If intMenuItems = 1 Then
            .Top = 75
        ElseIf lblMenuItem(intMenuItems - 1).Tag = "SplitterAfter" Then
            .Top = lblMenuItem(intMenuItems - 1).Top + lblMenuItem(intMenuItems - 1).Height + 90
        ElseIf lblMenuItem(intMenuItems - 1).Tag = "" Then
            .Top = lblMenuItem(intMenuItems - 1).Top + lblMenuItem(intMenuItems - 1).Height + 80
        End If
        .ZOrder vbBringToFront
        .Visible = True
    End With
    ResizeMenu
End Function

Function ResizeMenu()
    If intSplitterItems <= 0 Then
        imgMenu(2).Top = lblMenuItem(intMenuItems).Top + 80
    Else
        imgMenu(2).Top = lblMenuItem(intMenuItems).Top + 90
    End If
    frmMenu.Height = imgMenu(2).Top + imgMenu(2).Height
End Function

Private Sub lblMenuItem_Click(Index As Integer)
    MsgBox lblMenuItem(Index).Tag
End Sub

Private Sub lblMenuItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgHighlight
        .Top = lblMenuItem(Index).Top - 50
        .Visible = True
    End With
End Sub

Sub DropShadow(hwnd As Long)
    SetClassLong hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) Or CS_DROPSHADOW
End Sub

Public Sub FadeIn(Form As Form, ByVal hwnd As Long, Final As Integer)
Call MakeTransparent(hwnd, 0)
Dim X As Integer

    Form.Show
    X = 0
    Do Until X = Final
    DoEvents
        X = X + 1
        Call MakeTransparent(hwnd, X)
    Loop
End Sub

Public Sub FadeOut(Form As Form, ByVal hwnd As Long, UnldForm As Boolean)
Call MakeTransparent(hwnd, 220)
Dim Y As Integer

    Y = 300
    Do Until Y = 0
    DoEvents
        Y = Y - 1
        Call MakeTransparent(hwnd, Y)
    Loop
    If UnldForm = True Then
        Unload Form
    End If
End Sub

Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
Dim Msg As Long
On Error Resume Next

    If Perc < 0 Or Perc > 255 Then
        MakeTransparent = 1
    Else
        Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong hwnd, GWL_EXSTYLE, Msg
        SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
        MakeTransparent = 0
    End If
    If Err Then
        MakeTransparent = 2
    End If
End Function
