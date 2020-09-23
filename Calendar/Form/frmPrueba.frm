VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrueba 
   BackColor       =   &H00D2C8BE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar 1.1 - HACKPRO TM 2004"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   Icon            =   "frmPrueba.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picColor 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   9
      Left            =   3450
      MouseIcon       =   "frmPrueba.frx":058A
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   825
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1260
      Width           =   885
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   8
      Left            =   3450
      MouseIcon       =   "frmPrueba.frx":0894
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   825
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   975
      Width           =   885
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   7
      Left            =   3450
      MouseIcon       =   "frmPrueba.frx":0B9E
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   825
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   690
      Width           =   885
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   6
      Left            =   3450
      MouseIcon       =   "frmPrueba.frx":0EA8
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   825
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   405
      Width           =   885
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   5
      Left            =   3450
      MouseIcon       =   "frmPrueba.frx":11B2
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   825
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   885
   End
   Begin VB.CheckBox chk1 
      BackColor       =   &H00D2C8BE&
      Caption         =   "UseBackGroundPicture"
      ForeColor       =   &H007C4C2B&
      Height          =   240
      Index           =   4
      Left            =   105
      TabIndex        =   14
      Top             =   2370
      Width           =   2040
   End
   Begin VB.CheckBox chk1 
      BackColor       =   &H00D2C8BE&
      Caption         =   "ShowToolTipText"
      ForeColor       =   &H007C4C2B&
      Height          =   240
      Index           =   3
      Left            =   2745
      TabIndex        =   12
      Top             =   1680
      Width           =   1605
   End
   Begin VB.CheckBox chk1 
      BackColor       =   &H00D2C8BE&
      Caption         =   "Enabled"
      ForeColor       =   &H007C4C2B&
      Height          =   240
      Index           =   2
      Left            =   2265
      TabIndex        =   15
      Top             =   2370
      Width           =   900
   End
   Begin VB.CheckBox chk1 
      BackColor       =   &H00D2C8BE&
      Caption         =   "AutoSelect"
      ForeColor       =   &H007C4C2B&
      Height          =   240
      Index           =   1
      Left            =   105
      TabIndex        =   13
      Top             =   2040
      Width           =   1110
   End
   Begin VB.CheckBox chk1 
      BackColor       =   &H00D2C8BE&
      Caption         =   "ImageSelect"
      ForeColor       =   &H007C4C2B&
      Height          =   240
      Index           =   0
      Left            =   105
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cmbCustomImage 
      ForeColor       =   &H00C56A31&
      Height          =   315
      Left            =   2190
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtPrompChar 
      ForeColor       =   &H00C56A31&
      Height          =   330
      Left            =   2235
      MaxLength       =   1
      TabIndex        =   11
      Top             =   1815
      Width           =   350
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   4
      Left            =   1200
      MouseIcon       =   "frmPrueba.frx":14BC
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   825
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1260
      Width           =   885
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   3
      Left            =   1200
      MouseIcon       =   "frmPrueba.frx":17C6
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   825
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   975
      Width           =   885
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   1200
      MouseIcon       =   "frmPrueba.frx":1AD0
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   825
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   690
      Width           =   885
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   1
      Left            =   1200
      MouseIcon       =   "frmPrueba.frx":1DDA
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   825
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   405
      Width           =   885
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   1200
      MouseIcon       =   "frmPrueba.frx":20E4
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   825
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   885
   End
   Begin MSComDlg.CommonDialog cdDialog 
      Left            =   -285
      Top             =   -480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   405
      Left            =   3585
      TabIndex        =   17
      Top             =   3165
      Width           =   810
   End
   Begin Calendario.Calendar Calendar1 
      Height          =   300
      Left            =   2175
      TabIndex        =   18
      Top             =   3930
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   529
      CustomImage     =   2
      DayColor        =   7021576
      MousePointer    =   99
      MouseIcon       =   "frmPrueba.frx":23EE
      UseBackGroundPicture=   -1  'True
   End
   Begin VB.Image imgUserImage 
      Height          =   255
      Left            =   4095
      MouseIcon       =   "frmPrueba.frx":2708
      MousePointer    =   99  'Custom
      Picture         =   "frmPrueba.frx":2A12
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UserImageSelect:"
      ForeColor       =   &H007C4C2B&
      Height          =   195
      Index           =   12
      Left            =   2730
      TabIndex        =   33
      Top             =   2070
      Width           =   1260
   End
   Begin VB.Line linS 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   60
      X2              =   4385
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Line linS 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   75
      X2              =   4400
      Y1              =   2715
      Y2              =   2715
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomImage"
      ForeColor       =   &H007C4C2B&
      Height          =   195
      Index           =   11
      Left            =   2190
      TabIndex        =   31
      Top             =   3105
      Width           =   960
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PrompChar:"
      ForeColor       =   &H007C4C2B&
      Height          =   195
      Index           =   10
      Left            =   1380
      TabIndex        =   30
      Top             =   1875
      Width           =   825
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TodayColor:"
      ForeColor       =   &H007C4C2B&
      Height          =   195
      Index           =   9
      Left            =   2520
      TabIndex        =   29
      Top             =   1290
      Width           =   855
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SeparatorColor:"
      ForeColor       =   &H007C4C2B&
      Height          =   195
      Index           =   8
      Left            =   2280
      TabIndex        =   28
      Top             =   1005
      Width           =   1095
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SelectColor:"
      ForeColor       =   &H007C4C2B&
      Height          =   195
      Index           =   7
      Left            =   2505
      TabIndex        =   27
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NormalColor:"
      ForeColor       =   &H007C4C2B&
      Height          =   195
      Index           =   6
      Left            =   2475
      TabIndex        =   26
      Top             =   420
      Width           =   900
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MonthColor:"
      ForeColor       =   &H007C4C2B&
      Height          =   195
      Index           =   5
      Left            =   2520
      TabIndex        =   25
      Top             =   135
      Width           =   855
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HighLightColor:"
      ForeColor       =   &H007C4C2B&
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   24
      Top             =   1290
      Width           =   1080
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DisabledColor:"
      ForeColor       =   &H007C4C2B&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   1005
      Width           =   1020
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DayColor:"
      ForeColor       =   &H007C4C2B&
      Height          =   195
      Index           =   2
      Left            =   450
      TabIndex        =   22
      Top             =   720
      Width           =   690
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BorderColor:"
      ForeColor       =   &H007C4C2B&
      Height          =   195
      Index           =   1
      Left            =   270
      TabIndex        =   21
      Top             =   420
      Width           =   870
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BackColor:"
      ForeColor       =   &H007C4C2B&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   20
      Top             =   135
      Width           =   780
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HACKPRO TM © 2004 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2505
      TabIndex        =   19
      Top             =   2820
      Width           =   1980
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HACKPRO TM © 2004 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007C4C2B&
      Height          =   195
      Index           =   1
      Left            =   2490
      TabIndex        =   32
      Top             =   2805
      Width           =   1980
   End
   Begin VB.Image imgBackGroundPicture 
      Height          =   1575
      Left            =   120
      MouseIcon       =   "frmPrueba.frx":2B07
      MousePointer    =   99  'Custom
      Picture         =   "frmPrueba.frx":2E11
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1950
   End
End
Attribute VB_Name = "frmPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
          '***********************************'
          '* Copyright (C) 2004 - HACKPRO TM *'
          '*  Heriberto Mantilla Santamaría  *'
          '*        Barrancabermeja          *'
          '***********************************'
Option Explicit
 
 Private i As Integer

Private Sub cmdShow_Click()
 SetValue False
End Sub

Private Sub Form_Load()
 For i = 1 To 6
  cmbCustomImage.AddItem "Custom_" & i
 Next
 cmbCustomImage.ListIndex = Calendar1.CustomImage - 1
 Set Calendar1.BackGroundPicture = imgBackGroundPicture.Picture
 Set Calendar1.UserImageSelect = imgUserImage.Picture
 SetValue
End Sub

Private Function ShowDialog(Optional ByVal WhatIts As Boolean = False) As Variant
 If (WhatIts = True) Then
  With cdDialog
   .Filter = "Todas las Imagenes|*.bmp;*.gif;*.jpg;*.jpeg;*.wmf"
   .Flags = &H4
   .ShowOpen
   If (.CancelError Or .FileName = "") Then Exit Function
   ShowDialog = .FileName
  End With
 Else
  cdDialog.ShowColor
  ShowDialog = cdDialog.Color
 End If
End Function

Private Sub SetValue(Optional ByVal DefaultValue As Boolean = True)
 For i = 0 To 4
  If (Calendar1.CustomImage = Mid$(cmbCustomImage.Text, 8, 1)) Then
   Exit For
  Else
   Calendar1.CustomImage = cmbCustomImage.ListIndex + 1
   Exit For
  End If
 Next
 If (DefaultValue = True) Then
  picColor(0).BackColor = Calendar1.BackColor
  picColor(1).BackColor = Calendar1.BorderColor
  picColor(2).BackColor = Calendar1.DayColor
  picColor(3).BackColor = Calendar1.DisabledColor
  picColor(4).BackColor = Calendar1.HighLightColor
  picColor(5).BackColor = Calendar1.MonthColor
  picColor(6).BackColor = Calendar1.NormalColor
  picColor(7).BackColor = Calendar1.SelectColor
  picColor(8).BackColor = Calendar1.SeparatorColor
  picColor(9).BackColor = Calendar1.TodayColor
  chk1(0).Value = Int(Calendar1.ImageSelect) * -1
  chk1(1).Value = Int(Calendar1.AutoSelect) * -1
  chk1(2).Value = Int(Calendar1.Enabled) * -1
  chk1(3).Value = Int(Calendar1.ShowToolTipText) * -1
  chk1(4).Value = Int(Calendar1.UseBackGroundPicture) * -1
  txtPrompChar.Text = Calendar1.PrompChar
  Set imgUserImage.Picture = Calendar1.UserImageSelect
  Set imgBackGroundPicture.Picture = Calendar1.BackGroundPicture
 Else
  Calendar1.BackColor = picColor(0).BackColor
  Calendar1.BorderColor = picColor(1).BackColor
  Calendar1.DayColor = picColor(2).BackColor
  Calendar1.DisabledColor = picColor(3).BackColor
  Calendar1.HighLightColor = picColor(4).BackColor
  Calendar1.MonthColor = picColor(5).BackColor
  Calendar1.NormalColor = picColor(6).BackColor
  Calendar1.SelectColor = picColor(7).BackColor
  Calendar1.SeparatorColor = picColor(8).BackColor
  Calendar1.TodayColor = picColor(9).BackColor
  Calendar1.ImageSelect = Int(chk1(0).Value) * -1
  Calendar1.AutoSelect = Int(chk1(1).Value) * -1
  Calendar1.Enabled = Int(chk1(2).Value) * -1
  Calendar1.ShowToolTipText = Int(chk1(3).Value) * -1
  Calendar1.UseBackGroundPicture = Int(chk1(4).Value) * -1
  Calendar1.PrompChar = txtPrompChar.Text
  Set Calendar1.UserImageSelect = imgUserImage.Picture
  Set Calendar1.BackGroundPicture = imgBackGroundPicture.Picture
 End If
End Sub

Private Sub imgBackGroundPicture_Click()
 Set imgBackGroundPicture.Picture = LoadPicture(ShowDialog(True), vbResBitmap)
End Sub

Private Sub imgUserImage_Click()
 Set imgUserImage.Picture = LoadPicture(ShowDialog(True), vbResBitmap)
End Sub

Private Sub picColor_Click(Index As Integer)
 picColor(Index).BackColor = ShowDialog()
End Sub
