VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TimeEdit100 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   LockControls    =   -1  'True
   ScaleHeight     =   2310
   ScaleWidth      =   2475
   Begin VB.Timer ClickTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1275
      Top             =   1050
   End
   Begin MSComctlLib.Toolbar TBar 
      Height          =   195
      Left            =   1035
      TabIndex        =   5
      Tag             =   "TBarSeqEdit"
      Top             =   30
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   344
      ButtonWidth     =   291
      ButtonHeight    =   344
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   4
      ImageHeight     =   7
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TimeEdit100.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TimeEdit100.ctx":00B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   360
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "txtSeqEdit"
      Text            =   "00"
      Top             =   0
      Width           =   300
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   720
      MaxLength       =   3
      TabIndex        =   1
      Tag             =   "txtSeqEdit"
      Text            =   "00"
      Top             =   0
      Width           =   300
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   0
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "txtSeqEdit"
      Text            =   "00"
      Top             =   0
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   4
      Top             =   -15
      Width           =   60
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   660
      TabIndex        =   3
      Top             =   -15
      Width           =   60
   End
End
Attribute VB_Name = "TimeEdit100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim MinPacks As Long, MaxPacks As Long
Dim BtnPressed As Integer, FieldFocus As Integer

Public Event Change(newPacks As Long)

Private Sub UserControl_Initialize()
    FieldFocus = 2  'Default is to increase 1/100th column
End Sub

Private Sub ChangeEvent(newval As Long)
    RaiseEvent Change(newval)
End Sub


Private Sub ClickTimer_Timer()
    ClickBtn
End Sub

Private Sub TBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x < TBar.Buttons(2).Left Then    'Left tool button
        BtnPressed = IIf(Button = vbLeftButton, vbKeyDown, vbKeyPageDown)
    Else
        BtnPressed = IIf(Button = vbLeftButton, vbKeyUp, vbKeyPageUp)
    End If
    
    txt(FieldFocus).SetFocus
    ClickTimer.Interval = 150
    ClickTimer.Enabled = True
    ClickBtn
End Sub

Private Sub TBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ClickTimer.Enabled = False
End Sub

Private Sub ClickBtn()
    'Must do BtnPressed +0 otherwise txt_KeyDown sets it to zero
    txt_KeyDown FieldFocus, BtnPressed + 0, 0
End Sub

Private Sub txt_GotFocus(Index As Integer)
    txt(Index).SelStart = 0
    txt(Index).SelLength = Len(txt(Index))
    
    Select Case Index   'All tabs go round in a circle
        Case 0: txt(0).TabIndex = 0: txt(1).TabIndex = 1: txt(2).TabIndex = 2
        Case 1: txt(0).TabIndex = 2: txt(1).TabIndex = 0: txt(2).TabIndex = 1
        Case 2: txt(0).TabIndex = 1: txt(1).TabIndex = 2: txt(2).TabIndex = 0
    End Select
    FieldFocus = Index
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim tval(2), t1%, tadd As Long, tPacks As Long
    
    Select Case KeyCode
        Case vbKeyDown:     tadd = -1
        Case vbKeyUp:       tadd = 1
        Case vbKeyPageDown: tadd = -10
        Case vbKeyPageUp:   tadd = 10
    End Select
    
    If tadd <> 0 Then
        tPacks = GetVal
        Select Case Index
            Case 0: tPacks = tPacks + tadd * 6000
            Case 1: tPacks = tPacks + tadd * 100
            Case 2: tPacks = tPacks + tadd
        End Select
        If tPacks < MinPacks Then tPacks = MinPacks
        If tPacks > MaxPacks Then tPacks = MaxPacks
        
        SetVal tPacks, MinPacks, MaxPacks
        ChangeEvent tPacks
        
        KeyCode = 0
        txt(Index).SelStart = 0
        txt(Index).SelLength = Len(txt(Index))
        
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And _
        KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txt_LostFocus(Index As Integer)
    txt_Validate Index, False
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Dim tval
    With txt(Index)
        tval = Int(Abs(Val(.Text)))
        Select Case Index
            Case 0
                .Text = format$(tval Mod 100, "00")
            Case 1
                .Text = format$(tval Mod 60, "00")
            Case 2
                .Text = format$(tval Mod 100, "00")
        End Select
    End With
    tval = GetVal
    If tval < MinPacks Then
        SetVal MinPacks, MinPacks, MaxPacks
    ElseIf MaxPacks <> 0 And tval > MaxPacks Then
        SetVal MaxPacks, MinPacks, MaxPacks
    End If
    ChangeEvent GetVal
End Sub

'Returns the number of packs
Public Function GetVal() As Long
    GetVal = (Val(txt(0)) * 60 + Val(txt(1))) * 100 + Val(txt(2))
End Function

Public Function MinVal() As Long
    MinVal = MinPacks
End Function

Public Function MaxVal() As Long
    MaxVal = MinPacks
End Function

Public Function PackString(NumPacks)
    Dim tsecs As Long
    
    tsecs = Int(NumPacks / 100)
    PackString = format$(Int(tsecs / 60), "00") & ":" & _
             format$(tsecs Mod 60, "00") & ":" & _
             format$(NumPacks Mod 100, "00")
End Function

'Sets the number of packs
Public Sub SetVal(NumPacks As Long, Optional MinP As Long = 0, Optional MaxP As Long = 1799999)
    Dim tsecs As Long
    
    tsecs = Int(NumPacks / 100)
    txt(0).Text = format$(Int(tsecs / 60), "00")
    txt(1).Text = format$(tsecs Mod 60, "00")
    txt(2).Text = format$(NumPacks Mod 100, "00")
    MinPacks = MinP: MaxPacks = MaxP
End Sub

Private Sub UserControl_LostFocus()
    ChangeEvent GetVal
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 1110 + TBar.Width + 30
    UserControl.Height = 315
End Sub

