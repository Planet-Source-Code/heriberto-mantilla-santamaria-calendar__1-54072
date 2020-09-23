VERSION 5.00
Begin VB.UserControl Calendar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EAFFFF&
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2220
   LockControls    =   -1  'True
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   148
   ToolboxBitmap   =   "Calendar.ctx":0000
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   -6180
      Top             =   3060
   End
   Begin VB.TextBox txtDate 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   75
      MaxLength       =   9
      TabIndex        =   1
      Top             =   45
      Width           =   1890
   End
   Begin VB.PictureBox picCalendar 
      BackColor       =   &H00EED5C4&
      BorderStyle     =   0  'None
      Height          =   1905
      Left            =   0
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   3
      Top             =   300
      Visible         =   0   'False
      Width           =   2220
      Begin VB.Label lblMonthNext 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C4C2B&
         Height          =   195
         Left            =   1980
         TabIndex        =   11
         Top             =   60
         Width           =   195
      End
      Begin VB.Label lblMonthPrev 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007C4C2B&
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   195
      End
      Begin VB.Label lblMonthYear 
         Alignment       =   2  'Center
         BackColor       =   &H00C56A31&
         BackStyle       =   0  'Transparent
         Caption         =   "Month 0000"
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
         Height          =   255
         Left            =   0
         TabIndex        =   55
         Top             =   60
         Width           =   2235
      End
      Begin VB.Label lblMove 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mo"
         Height          =   195
         Left            =   405
         TabIndex        =   56
         Top             =   2310
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgSelect 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2340
         Picture         =   "Calendar.ctx":0312
         Stretch         =   -1  'True
         Top             =   2190
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgCustom 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   0
         Left            =   2340
         Picture         =   "Calendar.ctx":081C
         Stretch         =   -1  'True
         Top             =   1635
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgCustom 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   255
         Picture         =   "Calendar.ctx":0D26
         Stretch         =   -1  'True
         Top             =   2310
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgCustom 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   2
         Left            =   2355
         Picture         =   "Calendar.ctx":1230
         Stretch         =   -1  'True
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgCustom 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   3
         Left            =   2640
         Picture         =   "Calendar.ctx":173A
         Stretch         =   -1  'True
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgCustom 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   4
         Left            =   2640
         Picture         =   "Calendar.ctx":1C44
         Stretch         =   -1  'True
         Top             =   2190
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   34
         Left            =   1860
         TabIndex        =   13
         Top             =   1620
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   33
         Left            =   1560
         TabIndex        =   14
         Top             =   1620
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   32
         Left            =   1260
         TabIndex        =   15
         Top             =   1620
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   31
         Left            =   960
         TabIndex        =   16
         Top             =   1620
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   30
         Left            =   660
         TabIndex        =   17
         Top             =   1620
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   29
         Left            =   360
         TabIndex        =   18
         Top             =   1620
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   28
         Left            =   60
         TabIndex        =   19
         Top             =   1620
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   27
         Left            =   1860
         TabIndex        =   20
         Top             =   1380
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   26
         Left            =   1560
         TabIndex        =   21
         Top             =   1380
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   25
         Left            =   1260
         TabIndex        =   22
         Top             =   1380
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   24
         Left            =   960
         TabIndex        =   23
         Top             =   1380
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   23
         Left            =   660
         TabIndex        =   24
         Top             =   1380
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   22
         Left            =   360
         TabIndex        =   25
         Top             =   1380
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   21
         Left            =   60
         TabIndex        =   26
         Top             =   1380
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   20
         Left            =   1860
         TabIndex        =   27
         Top             =   1140
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   19
         Left            =   1560
         TabIndex        =   28
         Top             =   1140
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   18
         Left            =   1260
         TabIndex        =   29
         Top             =   1140
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   17
         Left            =   960
         TabIndex        =   30
         Top             =   1140
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   16
         Left            =   660
         TabIndex        =   31
         Top             =   1140
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   15
         Left            =   360
         TabIndex        =   32
         Top             =   1140
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   14
         Left            =   60
         TabIndex        =   33
         Top             =   1140
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   13
         Left            =   1860
         TabIndex        =   34
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   12
         Left            =   1560
         TabIndex        =   35
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   11
         Left            =   1260
         TabIndex        =   36
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   10
         Left            =   960
         TabIndex        =   37
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   9
         Left            =   660
         TabIndex        =   38
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   39
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   7
         Left            =   60
         TabIndex        =   40
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   42
         Top             =   660
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   4
         Left            =   1260
         TabIndex        =   43
         Top             =   660
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   44
         Top             =   660
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   2
         Left            =   660
         TabIndex        =   45
         Top             =   660
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   46
         Top             =   660
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   47
         Top             =   660
         Width           =   255
      End
      Begin VB.Line linSep 
         BorderColor     =   &H00808080&
         X1              =   5
         X2              =   141
         Y1              =   40
         Y2              =   40
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H00C56A31&
         Height          =   255
         Index           =   6
         Left            =   1860
         TabIndex        =   41
         Top             =   660
         Width           =   255
      End
      Begin VB.Line line 
         BorderColor     =   &H00C56A31&
         Index           =   1
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   24
      End
      Begin VB.Line line 
         BorderColor     =   &H00C56A31&
         Index           =   0
         X1              =   147
         X2              =   147
         Y1              =   0
         Y2              =   24
      End
      Begin VB.Label lblDoW 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Th"
         ForeColor       =   &H006A240A&
         Height          =   255
         Index           =   4
         Left            =   1260
         TabIndex        =   50
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblDoW 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Su"
         ForeColor       =   &H006A240A&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   54
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblDoW 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mo"
         ForeColor       =   &H006A240A&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   53
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblDoW 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tu"
         ForeColor       =   &H006A240A&
         Height          =   255
         Index           =   2
         Left            =   660
         TabIndex        =   52
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblDoW 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "We"
         ForeColor       =   &H006A240A&
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   51
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblDoW 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fr"
         ForeColor       =   &H006A240A&
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   49
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblDoW 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sa"
         ForeColor       =   &H006A240A&
         Height          =   255
         Index           =   6
         Left            =   1860
         TabIndex        =   48
         Top             =   360
         Width           =   255
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H00C56A31&
         Height          =   1605
         Left            =   0
         Top             =   300
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   41
         Left            =   1860
         TabIndex        =   10
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   40
         Left            =   1560
         TabIndex        =   9
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   39
         Left            =   1260
         TabIndex        =   8
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   38
         Left            =   960
         TabIndex        =   7
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   37
         Left            =   660
         TabIndex        =   6
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   36
         Left            =   360
         TabIndex        =   5
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   35
         Left            =   60
         TabIndex        =   4
         Top             =   1860
         Width           =   255
      End
      Begin VB.Shape shpMonthYearBack 
         BackColor       =   &H00D0C7B0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   2220
      End
   End
   Begin VB.Image img4 
      Height          =   300
      Left            =   -6180
      Picture         =   "Calendar.ctx":214E
      Top             =   2490
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgDropDown 
      Height          =   300
      Left            =   1995
      MouseIcon       =   "Calendar.ctx":25A8
      MousePointer    =   99  'Custom
      Picture         =   "Calendar.ctx":28B2
      Top             =   0
      Width           =   225
   End
   Begin VB.Shape shpHighLight 
      Height          =   225
      Left            =   -6180
      Top             =   3330
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label lblSelectColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      Height          =   195
      Left            =   -6180
      TabIndex        =   2
      Top             =   2940
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image img1 
      Height          =   300
      Left            =   -6180
      MouseIcon       =   "Calendar.ctx":2D0C
      MousePointer    =   99  'Custom
      Picture         =   "Calendar.ctx":3016
      Top             =   2505
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image img2 
      Height          =   300
      Left            =   -6180
      Picture         =   "Calendar.ctx":3470
      Top             =   2505
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image img3 
      Height          =   300
      Left            =   -6180
      Picture         =   "Calendar.ctx":38CA
      Top             =   2505
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblBlock 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Shape shpList 
      BorderColor     =   &H00DEEDEF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   2220
   End
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
          '***********************************'
          '* Copyright (C) 2004 - HACKPRO TM *'
          '*  Heriberto Mantilla Santamaría  *'
          '*        Barrancabermeja          *'
          '***********************************'
Option Explicit
 
 Public Event MouseMove()
 
 Private CalendarVisible As Boolean
 Private ListExpanded    As Boolean
 Private ImgSel          As Boolean
 Private IsMove          As Boolean
 Private myAutoSelect    As Boolean
 Private ToolTipS        As Boolean
 Private myUseBack       As Boolean
 
 Private TodayIndex      As Integer
 Private DayIndex        As Integer
 
 Private m_MonthColor    As Long
 Private m_DayColor      As Long
 Private m_TodayColor    As Long
 Private m_BorderColor   As Long
 Private m_BackColor     As Long
 Private sIndex          As Long
 Private sngScaleX       As Long
 Private sngScaleY       As Long
  
 Private CalDate         As Date
 Private CustomDate      As Date
 
 Private TodayDay        As String
 
 Private myPicture       As StdPicture
 Private myBackGround    As StdPicture
 
 Private Const SiColor = &HC56A31
 Private Const NoColor = &H9900FF
 Private Const HighColor = &HFE0099
 Private Const kBackColor = &HEED5C4
 Private Const kBorderColor = &HD0C7B0
 Private Const kMonthColor = &H8A4500
 Private Const kDayColor = &H6A240A
 Private Const kTodayColor = &H97080E
 Private Const kSepaColor = &H73B0BB
 Private Const m_def_DisabledColor = &H808080
  
 Private iX        As Long
 Private sCaption  As String, m_PrompChar As String
 Private Der       As Long, iY            As Long
 Private Izq       As Long
 Private m_Pointer As MousePointerConstants
 Private m_Icon    As Picture
  
 Public Enum Select_Image
  [Custom_1] = 1
  [Custom_2] = 2
  [Custom_3] = 3
  [Custom_4] = 4
  [Custom_5] = 5
  [Custom_6] = 6
 End Enum
 
 Private Type RECT
  Left   As Long
  Top    As Long
  Right  As Long
  Bottom As Long
 End Type
 
 Private Type POINTAPI
  X As Long
  Y As Long
 End Type
 
 Private sImge           As Select_Image
 Private m_DisabledColor As OLE_COLOR
 Private m_SepaColor     As OLE_COLOR
 
 Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
 Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
 Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
 Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
 Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
 Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
 Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
 Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Sub imgDropDown_Click()
 '* Español: Muestra o oculta la lista del Calendario.
 '* English: Shows or hides the list of the Calendar.
 CalendarVisible = Not (CalendarVisible)
 lblMove.Visible = False
 If (CalendarVisible = True) Then
  shpBorder.Visible = True
  shpMonthYearBack.Visible = True
  '* Español: Si (CustomDate = 0) devuelve la fecha del sistema.
  '* English: If (CustomDate = 0) returns the date of the system.
  If (CustomDate = 0) Then
   CalDate = Format$(Now, "dd/mm/yy")
  Else
   CalDate = Format$(CustomDate, "dd/mm/yy")
  End If
  '* Español: Oculta la lista.
  '* English: Hidden the list.
  ListExpanded = False
  shpList.BorderColor = &HC56A31
  imgDropDown.Picture = img3.Picture
  '* Español: Procedimientos Externos.
  '* English: External Procedures.
  DrawCalendar
  ShowPopUp
  picCalendar.Visible = True
 Else
  shpBorder.Visible = False
  shpMonthYearBack.Visible = False
  picCalendar.Visible = False
  ListExpanded = False
  '* Español: Procedimiento Externo.
  '* English: External Procedure.
  ChangeImg
 End If
 UserControl_Resize
 Refresh
End Sub

Private Sub lblDay_Click(Index As Integer)
 lblMove_Click
End Sub

Private Sub lblDay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Char As String
 
 '* Español: Simula un efecto de selección para cuando el mouse se detiene en cualquier día.
 '* English: Simulates a selection effect for when the mouse stops in any day.
 sIndex = Index
 lblMove.Caption = lblDay(Index).Caption
 lblMonthNext.ForeColor = MonthColor
 lblMonthPrev.ForeColor = MonthColor
 '* Español: Procedimiento Externo.
 '* English: External Procedure.
 SetTodayColor
 If (Len(lblMove.Caption) > 1) Then
  If (lblDay(Index).Tag <> "Today") Then
   lblMove.Move lblDay(Index).Left + 2, lblDay(Index).Top
  Else
   lblMove.Move -10, -10
  End If
 Else
  If (lblDay(Index).Tag <> "Today") Then
   lblMove.Move lblDay(Index).Left + 5, lblDay(Index).Top
  Else
   lblMove.Move -10, -10
  End If
 End If
 lblMove.Visible = True
 '* Español: Devuelve la fecha en formato largo para mostrarla en la propiedad ToolTipText, cuando ShowToolTipText = True.
 '* English: Returns the date in long format to show it in the property ToolTipText, when ShowToolTipText = True.
 Char = Format$(Format$(lblMove.Caption & " " & lblMonthYear.Caption, "dd/mm/yy"), "Long Date")
 If (ShowToolTipText = True) Then lblMove.ToolTipText = Char Else lblMove.ToolTipText = ""
 '* Español: Indica si se esta movimiento por los días.
 '* English: Indicates if you this movement for the days.
 IsMove = True
End Sub

Private Sub SetTodayColor()
 '* Español: Establece el color del día actual.
 '* English: Establishes the color of the current day.
 If (lblDay(sIndex).Caption = TodayDay) And (lblDay(TodayIndex).Caption = TodayDay) Then lblDay(TodayIndex).ForeColor = TodayColor
End Sub

Private Sub SetMonthColor()
 '* Español: Establece el color normal de Month Previous y Month Next.
 '* English: Establishes Month Previous and Month Next to their normal color.
 lblMonthNext.ForeColor = MonthColor
 lblMonthPrev.ForeColor = MonthColor
End Sub

Private Sub lblDoW_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 SetMonthColor
End Sub

Private Sub lblMonthNext_Click()
 '* Español: Avanza al siguiente Mes.
 '* English: Advances to the following Month.
 IsMove = True
 CalDate = DateSerial(Year(CalDate), Month(CalDate) + 1, Day(CalDate))
 DrawCalendar
End Sub

Private Sub lblMonthNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Char As Integer
 
 lblMonthNext.ForeColor = HighLightColor
 Char = Format$(lblMonthYear.Caption, "MM")
 Char = Format$(Char + 1, "00")
 If (ShowToolTipText = True) Then lblMonthNext.ToolTipText = FormatMonth(Char) Else lblMonthNext.ToolTipText = ""
End Sub

Private Function FormatMonth(ByVal Month As Integer) As String
 '* Español: Devuelve el mes.
 '* English: Return the month.
 Select Case Month
  Case 1:  FormatMonth = "Enero"
  Case 2:  FormatMonth = "Febrero"
  Case 3:  FormatMonth = "Marzo"
  Case 4:  FormatMonth = "Abril"
  Case 5:  FormatMonth = "Mayo"
  Case 6:  FormatMonth = "Junio"
  Case 7:  FormatMonth = "Julio"
  Case 8:  FormatMonth = "Agosto"
  Case 9:  FormatMonth = "Septiembre"
  Case 10: FormatMonth = "Octubre"
  Case 11: FormatMonth = "Noviembre"
  Case 12: FormatMonth = "Diciembre"
 End Select
End Function

Private Sub lblMonthPrev_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Char As Integer
 
 lblMonthPrev.ForeColor = HighLightColor
 Char = Format$(lblMonthYear.Caption, "MM")
 Char = Format$(Char - 1, "00")
 If (ShowToolTipText = True) Then lblMonthPrev.ToolTipText = FormatMonth(Char) Else lblMonthPrev.ToolTipText = ""
End Sub

Private Sub lblMonthPrev_Click()
 '* Español: Regresa un Mes.
 '* English: Back one Month.
 IsMove = True
 CalDate = DateSerial(Year(CalDate), Month(CalDate) - 1, Day(CalDate))
 DrawCalendar
End Sub

Private Sub lblMonthYear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 SetMonthColor
End Sub

Private Sub lblMove_Click()
 '* Español: Devuelve o establece la fecha seleccionada.
 '* English: Gets/Sets the selected date.
 Text = Format$(lblDay(sIndex).Caption & " " & lblMonthYear, "dd/mm/yy")
 CalendarVisible = False
 ListExpanded = False
 ChangeImg
 UserControl_Resize
End Sub

Private Sub picCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 lblMove.Visible = False
 SetMonthColor
 SetTodayColor
End Sub

Private Sub tmrFocus_Timer()
 Dim p     As POINTAPI
 Dim Rec   As RECT
 Dim sView As Boolean
 
 '* Español: Si el mouse no se encuentra sobre el objeto, oculta la lista.
 '* English: If the mouse is not on the object, hidden the list.
 sView = False
 GetWindowRect UserControl.hWnd, Rec
 GetCursorPos p
 If (p.X >= Rec.Left) And (p.X <= Rec.Right) And (p.Y >= Rec.Top) And (p.Y <= Rec.Bottom) Then
  IsMove = False
 Else
  IsMove = True
 End If
 If (GetAsyncKeyState(vbLeftButton)) Or (GetAsyncKeyState(vbRightButton)) Then
  If (IsMove = True) Then sView = True
 End If
 If (Height <> "300") And (sView = False) Then Exit Sub
 If (iX = p.X) Or (iY = p.Y) Then Exit Sub
 iX = p.X: iY = p.Y
 If (hWnd <> WindowFromPoint(p.X, p.Y)) Then
  If (Height = "300") Then
   If (imgDropDown.Tag <> "Ya") Then
    SetTodayColor
    ChangeImg
   End If
  ElseIf (IsMove = True) Then
   imgDropDown_Click
  End If
 Else
  HighLightList
 End If
 Refresh
End Sub

Private Sub txtDate_GotFocus()
 '* Español: Selecciona el texto, si AutoSelect = True.
 '* English: Selects the text, if AutoSelect = True.
 If (myAutoSelect = True) Then
  With txtDate
   .SelStart = 0
   .SelLength = Len(.Text)
  End With
 End If
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
 Dim iText As String
 
 '* Español: Pregunta si es válida la fecha.
 '* English: Asks if it's valid the date.
 iText = Replace(txtDate.Text, Mid$(txtDate.Text, 3, 1), "/")
 If (IsDate(iText) = True) Then
  Text = iText
 Else
  Cancel = True
 End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
 '* Español: Para permite únicamente Números.
 '* English: For it only allows Numbers.
 If (KeyAscii = 8) Or (KeyAscii = vbKeyShift) Then Exit Sub
 If (KeyAscii < 47) Or (KeyAscii > 58) Then
  KeyAscii = 0
  Beep
 End If
On Error Resume Next
 Text = txtDate.Text
End Sub

Private Sub txtDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '* Español: Si esta oculta la lista muestra el enfoque.
 '* English: If this hidden list shows the focus.
 If (UserControl.Height = "300") Then HighLightList
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
 picCalendar.Visible = False
 CalendarVisible = False
 ListExpanded = False
 ChangeImg
 Refresh
End Sub

Private Sub UserControl_ExitFocus()
 picCalendar.Visible = False
 CalendarVisible = False
 ChangeImg
 ColorDays
 Refresh
End Sub

Private Sub ChangeImg()
 '* Español: Cambia a la apariencia normal.
 '* English: Changes to the normal appearance.
 ListExpanded = False
 UserControl.Height = 300
 imgDropDown.Picture = img1.Picture
 shpList.BorderColor = &HDEEDEF
 picCalendar.Visible = False
 imgDropDown.Tag = "Ya"
 tmrFocus.Enabled = False
 SetMonthColor
 lblMove.Visible = False
 Refresh
End Sub

Private Sub UserControl_Initialize()
 Dim intIndex As Integer

 CalendarVisible = False
 Width = 2220
 For intIndex = 0 To 6
  lblDoW(intIndex).Caption = Left$(Format$(intIndex + 1, "dddd", vbSunday), 2)
 Next
 TodayDay = Day(Now)
End Sub

Private Sub UserControl_InitProperties()
 Dim i As Integer
 
 For i = 0 To 41
  lblDay(i).Caption = ""
 Next
 CustomDate = Format$(Now, "dd/mm/yy")
 myAutoSelect = True
 BackColor = kBackColor
 NormalColor = SiColor
 m_MonthColor = kMonthColor
 m_BorderColor = kBorderColor
 m_SepaColor = kSepaColor
 m_DayColor = kDayColor
 m_TodayColor = kTodayColor
 myUseBack = False
 SelectColor = NoColor
 Set myPicture = Nothing
 Set myBackGround = Nothing
 m_PrompChar = "/"
 ToolTipS = False
 sImge = Custom_1
 linSep.BorderColor = MonthColor
 HighLightColor = HighColor
 SetMonthColor
 txtDate.ForeColor = NormalColor
 UserControl.Width = 2220
 m_DisabledColor = m_def_DisabledColor
 Set m_Icon = Nothing
 m_Pointer = vbDefault
 ImgSel = True
End Sub

Private Sub UserControl_LostFocus()
 UserControl_ExitFocus
End Sub

Public Property Get NormalColor() As OLE_COLOR
Attribute NormalColor.VB_Description = "Determina el Color Normal del Texto."
 NormalColor = UserControl.ForeColor
End Property

Public Property Let NormalColor(ByVal New_FontColor As OLE_COLOR)
 '* Español: Color normal del Texto.
 '* English: Normal color of the Text.
 UserControl.ForeColor = New_FontColor
 If (Enabled = True) Then txtDate.ForeColor = New_FontColor
 PropertyChanged "NormalColor"
 Refresh
End Property

Public Property Get ShowToolTipText() As Boolean
Attribute ShowToolTipText.VB_Description = "Devuelve o establece si se muestra información adicional en el Calendario."
 ShowToolTipText = ToolTipS
End Property

Public Property Let ShowToolTipText(ByVal ShowTip As Boolean)
 '* Español: Devuelve o establece si se Muestra información adicional.
 '* English: Gets/Sets shows additional information.
 ToolTipS = ShowTip
 PropertyChanged "ShowToolTipText"
End Property

Public Property Get DayColor() As OLE_COLOR
Attribute DayColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar el texto de los días de cada mes."
 DayColor = m_DayColor
End Property

Public Property Let DayColor(ByVal New_FontColor As OLE_COLOR)
 '* Español: Color normal de los días.
 '* English: Normal color of the days.
 m_DayColor = New_FontColor
 PropertyChanged "DayColor"
 Refresh
End Property

Public Property Get TodayColor() As OLE_COLOR
Attribute TodayColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar el día actual"
 TodayColor = m_TodayColor
End Property

Public Property Let TodayColor(ByVal New_FontColor As OLE_COLOR)
 '* Español: Color del día actual del sistema.
 '* English: Color of the current day of the system.
 m_TodayColor = New_FontColor
 PropertyChanged "TodayColor"
 Refresh
End Property

Public Property Get CustomImage() As Select_Image
Attribute CustomImage.VB_Description = "Devuelve el tipo de Imagen a mostrar cuando se selecciona un elemento."
 CustomImage = sImge
End Property

Public Property Let CustomImage(ByVal New_Image As Select_Image)
 '* Español: Devuelve o establece la imagen de selección para la fecha seleccionada.
 '* English: Gets/Sets selection image for the selected date.
 sImge = New_Image
 Select Case sImge
  Case "1"
   Set imgSelect.Picture = imgCustom(4).Picture
  Case "2"
   Set imgSelect.Picture = imgCustom(3).Picture
  Case "3"
   Set imgSelect.Picture = imgCustom(2).Picture
  Case "4"
   Set imgSelect.Picture = imgCustom(1).Picture
  Case "5"
   Set imgSelect.Picture = imgCustom(0).Picture
  Case "6"
   Set imgSelect.Picture = myPicture
 End Select
 PropertyChanged "CustomImage"
 Refresh
End Property

Public Property Get MonthColor() As OLE_COLOR
Attribute MonthColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar el texto de cada mes."
 MonthColor = m_MonthColor
End Property

Public Property Let MonthColor(ByVal New_FontColor As OLE_COLOR)
 '* Español: Color para el mes.
 '* English: Color for the month.
 m_MonthColor = New_FontColor
 PropertyChanged "MonthColor"
 Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un Objeto."
 BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_FontColor As OLE_COLOR)
 '* Español: Color de fondo del Usercontrol.
 '* English: BackColor of the Usercontrol.
 m_BackColor = New_FontColor
 shpBorder.BackColor = m_BackColor
 picCalendar.BackColor = m_BackColor
 PropertyChanged "BackColor"
 Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar el fondo de los meses."
 BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_FontColor As OLE_COLOR)
 '* Español: Color de fondo del Mes.
 '* English: BackColor of the Month.
 m_BorderColor = New_FontColor
 shpMonthYearBack.BackColor = m_BorderColor
 PropertyChanged "BorderColor"
 Refresh
End Property

Public Property Get SeparatorColor() As OLE_COLOR
Attribute SeparatorColor.VB_Description = "Devuelve o establece el Color de separación."
 SeparatorColor = m_SepaColor
End Property

Public Property Let SeparatorColor(ByVal New_SepaColor As OLE_COLOR)
 '* Español: Color del borde de separación entre Días y Día.
 '* English: Color of the separation border between Days and Day.
 m_SepaColor = New_SepaColor
 linSep.BorderColor = m_SepaColor
 PropertyChanged "SeparatorColor"
 Refresh
End Property

Public Property Get DisabledColor() As OLE_COLOR
Attribute DisabledColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un Objeto cuando el Objeto este Deshabilitado."
 DisabledColor = m_DisabledColor
End Property

Public Property Let DisabledColor(ByVal New_DisabledColor As OLE_COLOR)
 '* Español: Color del texto deshabilitado.
 '* English: Color of the disabled text.
 m_DisabledColor = New_DisabledColor
 lblBlock.ForeColor = m_DisabledColor
 txtDate.ForeColor = m_DisabledColor
 PropertyChanged "DisabledColor"
 Refresh
End Property

Public Property Get SelectColor() As OLE_COLOR
Attribute SelectColor.VB_Description = "Determina el Color de Selección cuando el mouse pasa por un elemento de la Lista."
 SelectColor = lblSelectColor.ForeColor
End Property

Public Property Let SelectColor(ByVal New_FontColor As OLE_COLOR)
 '* Español: Color de selección de un día.
 '* English: Color of selection of one day.
 lblSelectColor.ForeColor = New_FontColor
 PropertyChanged "SelectColor"
 Refresh
End Property

Public Property Get HighLightColor() As OLE_COLOR
Attribute HighLightColor.VB_Description = "Determina el Color de segundo plano para cuando se cambia de mes."
 HighLightColor = shpHighLight.BorderColor
End Property

Public Property Let HighLightColor(ByVal New_FontColor As OLE_COLOR)
 '* Español: Color de selección de cualquier día.
 '* English: Color of selection of any day.
 shpHighLight.BorderColor = New_FontColor
 PropertyChanged "HighLightColor"
 Refresh
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Establece un icono personalizado para el mouse."
 Set MouseIcon = m_Icon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
 Dim i As Long
 
 Set m_Icon = New_MouseIcon
 imgDropDown.MouseIcon = m_Icon
 lblMonthNext.MouseIcon = m_Icon
 lblMonthPrev.MouseIcon = m_Icon
 lblMove.MouseIcon = m_Icon
 imgSelect.MouseIcon = m_Icon
 For i = 0 To 34
  lblDay(i).MouseIcon = m_Icon
 Next
 PropertyChanged "MouseIcon"
 Refresh
End Property

Public Property Get MousePointer() As MousePointerConstants
 MousePointer = m_Pointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
 Dim i As Long
 
 m_Pointer = New_MousePointer
 imgDropDown.MousePointer = m_Pointer
 lblMonthNext.MousePointer = m_Pointer
 lblMonthPrev.MousePointer = m_Pointer
 lblMove.MousePointer = m_Pointer
 imgSelect.MousePointer = m_Pointer
 For i = 0 To 34
  lblDay(i).MousePointer = m_Pointer
 Next
 PropertyChanged "MousePointer"
 Refresh
End Property

Public Property Get hWnd() As Long
 hWnd = UserControl.hWnd
 Refresh
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un Objeto puede responder a eventos generados por el usuario."
 Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
 ChangeImg
 UserControl.Enabled() = New_Enabled
 If (Enabled = True) Then
  imgDropDown.Picture = img1.Picture
  lblBlock.ForeColor = UserControl.ForeColor
  txtDate.ForeColor = UserControl.ForeColor
  tmrFocus.Enabled = UserControl.Ambient.UserMode
 Else
  imgDropDown.Picture = img4.Picture
  lblBlock.ForeColor = DisabledColor
  txtDate.ForeColor = DisabledColor
  tmrFocus.Enabled = False
 End If
 PropertyChanged "Enabled"
 Refresh
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 tmrFocus.Enabled = True
 SetMonthColor
 SetTodayColor
 lblMove.Visible = False
 RaiseEvent MouseMove
End Sub

Private Sub imgDropDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 tmrFocus.Enabled = True
End Sub

Private Sub HighLightList()
 If (ListExpanded = False) Then
  imgDropDown.Picture = img2.Picture
  shpList.BorderColor = &HC56A31
  shpBorder.Visible = True
  shpMonthYearBack.Visible = True
  imgDropDown.Tag = ""
 End If
 Refresh
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ChangeImg
End Sub

Private Sub UserControl_Resize()
 With UserControl
  .Width = 2220
  shpList.Width = .Width
  If (CalendarVisible = True) Then
   .Height = 2198
  Else
   .Height = 300
  End If
 End With
 Refresh
End Sub

Private Sub DrawCalendar()
 Dim intIndex As Integer, DateF As String
 Dim intDay   As Integer, DateG As String
 Dim i        As Integer

 '* Español: Crea el calendario.
 '* English: Create the calendar.
On Error GoTo myErr
 lblMonthYear.ForeColor = MonthColor
 lblMove.ForeColor = HighLightColor
 lblMove.Visible = False
 lblMonthYear = Format$(CalDate, "MMMM YYYY")
 imgSelect.Visible = False
 If (UseBackGroundPicture = True) Then
  Set picCalendar.Picture = BackGroundPicture
 Else
  Set picCalendar.Picture = Nothing
 End If
 For i = 0 To 6
  lblDoW(i).ForeColor = DayColor
 Next
 ColorDays
 intIndex = Format("1 " & Month(CalDate) & " " & Year(CalDate), "w") - 1
 intDay = 1
 For intIndex = intIndex To intIndex + Day(DateSerial(Year(CalDate), Month(CalDate) + 1, 0)) - 1
  With lblDay(intIndex)
   .Caption = CInt(intDay)
   DateF = Format$(intDay & " " & lblMonthYear, "dd/mm/yy")
   DateG = Format$(Now, "dd/mm/yy")
   If (DateF = DateG) Then
    .ForeColor = TodayColor
    .FontBold = True
    .Tag = "Today"
    TodayIndex = intIndex
   End If
   DateF = Format$(intDay & " " & lblMonthYear, "dd/mm/yy")
   DateG = Format$(CustomDate, "dd/mm/yy")
   If (DateF = DateG) Then sCaption = CInt(.Caption)
   .Visible = True
  End With
  intDay = intDay + 1
 Next
 For i = 0 To 34
  If (lblDay(i).Caption = sCaption) Then Exit For
 Next
 If (ImgSel = True) Then
  If (lblDay(i).Visible = True) And (lblDay(i).Caption <> "") And (i <= 34) Then
   imgSelect.Move lblDay(i).Left - 2, lblDay(i).Top - 3
   imgSelect.Visible = True
  End If
 ElseIf (lblDay(i).Tag <> "Today") Then
  lblDay(i).ForeColor = SelectColor
 End If
 txtDate.Text = Replace(CalDate, Mid$(CalDate, 3, 1), PrompChar)
 Refresh
 Exit Sub
myErr:
End Sub

Private Sub ColorDays()
 Dim intIndex As Integer
 
 '* Español: Color normal de los días.
 '* English: Normal color of the days.
 For intIndex = 0 To 41
  With lblDay(intIndex)
   .ForeColor = NormalColor
   .Visible = False
   .FontBold = False
   .Tag = ""
  End With
 Next
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
  PrompChar = PropBag.ReadProperty("PrompChar", "/")
  Text = .ReadProperty("Text", Format$(Now, "dd/mm/yy"))
  myAutoSelect = .ReadProperty("AutoSelect", True)
  ImgSel = .ReadProperty("ImageSelect", True)
  NormalColor = PropBag.ReadProperty("NormalColor", SiColor)
  MonthColor = PropBag.ReadProperty("MonthColor", kMonthColor)
  BorderColor = PropBag.ReadProperty("BorderColor", kBorderColor)
  SeparatorColor = PropBag.ReadProperty("SeparatorColor", kSepaColor)
  BackColor = PropBag.ReadProperty("BackColor", kBackColor)
  DayColor = PropBag.ReadProperty("DayColor", kDayColor)
  TodayColor = PropBag.ReadProperty("TodayColor", kTodayColor)
  SelectColor = PropBag.ReadProperty("SelectColor", NoColor)
  DisabledColor = PropBag.ReadProperty("DisabledColor", m_def_DisabledColor)
  HighLightColor = PropBag.ReadProperty("HighLightColor", HighColor)
  MousePointer = PropBag.ReadProperty("MousePointer", 0)
  Enabled = PropBag.ReadProperty("Enabled", True)
  Locked = PropBag.ReadProperty("Locked", False)
  DisabledColor = PropBag.ReadProperty("DisabledColor", m_def_DisabledColor)
  CustomImage = PropBag.ReadProperty("CustomImage", Select_Image.Custom_1)
  Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
  ShowToolTipText = PropBag.ReadProperty("ShowToolTipText", False)
  Set UserImageSelect = PropBag.ReadProperty("UserImageSelect", Nothing)
  Set BackGroundPicture = PropBag.ReadProperty("BackGroundPicture", Nothing)
  UseBackGroundPicture = PropBag.ReadProperty("UseBackGroundPicture", False)
 End With
End Sub

Private Sub UserControl_Show()
 Dim L As Long
 
 '* Español: La función GetWindowLong recupera información sobre la ventana especificada, SetWindowLong hace cambios a un atributo de la ventana especificada.
 '* English: The GetWindowLong function retrieves information about the specified window, The SetWindowLong function changes an attribute of the specified window.
On Error GoTo myErr
 SetParent picCalendar.hWnd, 0
 L& = GetWindowLong(picCalendar.hWnd, -20)
 Call SetWindowLong(picCalendar.hWnd, -20, L& Or &H80)
 SetWindowPos picCalendar.hWnd, picCalendar.hWnd, 0, 0, 0, 0, 39
 L& = SetWindowLong(picCalendar.hWnd, -8, Parent.hWnd)
 tmrFocus.Enabled = False
 UserControl.BackColor = Parent.BackColor
 If (Enabled = True) Then txtDate.ForeColor = UserControl.ForeColor
 Refresh
 Exit Sub
myErr:
End Sub

Private Sub UserControl_Terminate()
 tmrFocus.Enabled = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
  .WriteProperty "PrompChar", m_PrompChar, "/"
  .WriteProperty "Text", CustomDate, Format$(Now, "dd" & PrompChar & "mm" & PrompChar & "yy")
  .WriteProperty "AutoSelect", myAutoSelect, True
  .WriteProperty "ImageSelect", ImgSel, True
  .WriteProperty "CustomImage", sImge, Select_Image.Custom_1
  .WriteProperty "NormalColor", UserControl.ForeColor, SiColor
  .WriteProperty "MonthColor", m_MonthColor, kMonthColor
  .WriteProperty "DayColor", m_DayColor, kDayColor
  .WriteProperty "TodayColor", m_TodayColor, kTodayColor
  .WriteProperty "BorderColor", m_BorderColor, kBorderColor
  .WriteProperty "SeparatorColor", m_SepaColor, kSepaColor
  .WriteProperty "BackColor", m_BackColor, kBackColor
  .WriteProperty "SelectColor", lblSelectColor.ForeColor, NoColor
  .WriteProperty "HighLightColor", shpHighLight.BorderColor, HighColor
  .WriteProperty "DisabledColor", m_DisabledColor, m_def_DisabledColor
  .WriteProperty "MousePointer", m_Pointer, 0
  .WriteProperty "MouseIcon", m_Icon, Nothing
  .WriteProperty "Enabled", Enabled, True
  .WriteProperty "Locked", Locked, False
  .WriteProperty "ShowToolTipText", ToolTipS, False
  .WriteProperty "UserImageSelect", myPicture, Nothing
  .WriteProperty "BackGroundPicture", myBackGround, Nothing
  .WriteProperty "UseBackGroundPicture", myUseBack, False
 End With
End Sub

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determina si es posible modificar un Control."
 Locked = txtDate.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
 txtDate.Locked = New_Locked
 PropertyChanged "Locked"
 Refresh
End Property

Public Property Get Text() As Date
Attribute Text.VB_Description = "Devuelve o establece el texto contenido en el Control."
 If (CustomDate = 0) Then
  Text = Format$(Now, "dd/mm/yy")
 Else
  Text = Format$(CustomDate, "dd/mm/yy")
 End If
End Property

Public Property Let Text(ByVal NewValue As Date)
 '* Español: Devuelve o establece la fecha.
 '* English: Gets/Sets the date.
 CustomDate = NewValue
 CalDate = NewValue
 If Not (CustomDate = 0) Then txtDate.Text = Replace(CustomDate, Mid$(CustomDate, 3, 1), PrompChar)
 If (CalendarVisible = True) Then DrawCalendar
 PropertyChanged "Text"
End Property

Public Property Get PrompChar() As String
Attribute PrompChar.VB_Description = "Devuelve o establece el tipo de Promp para separar los días de los meses y los años."
 PrompChar = m_PrompChar
End Property

Public Property Let PrompChar(ByVal NewValue As String)
 m_PrompChar = NewValue
 txtDate.Text = Replace(txtDate.Text, Mid$(txtDate.Text, 3, 1), NewValue)
 PropertyChanged "PrompChar"
 Refresh
End Property

Public Property Get AutoSelect() As Boolean
Attribute AutoSelect.VB_Description = "Determina o establece si se selecciona el texto cuando el Objeto toma el enfoque."
 AutoSelect = myAutoSelect
End Property

Public Property Let AutoSelect(ByVal NewValue As Boolean)
 '* Español: Devuelve o establece si se seleccione el texto cuando se tome el enfoque.
 '* English: Gets/Sets if the text is selected when takes the focus.
 myAutoSelect = NewValue
 PropertyChanged "AutoSelect"
End Property

Public Property Get UseBackGroundPicture() As Boolean
Attribute UseBackGroundPicture.VB_Description = "Devuelve o establece si se muestra la imagen de fondo del Calendario."
 UseBackGroundPicture = myUseBack
End Property

Public Property Let UseBackGroundPicture(ByVal NewValue As Boolean)
 '* Español: Devuelve o establece si se usa la imagen escogida por el usuario.
 '* English: Uses the image selected by the user like selection.
 myUseBack = NewValue
 PropertyChanged "UseBackGroundPicture"
End Property

Public Property Get ImageSelect() As Boolean
Attribute ImageSelect.VB_Description = "Devuelve o establece si se muestra una Imagen cuando se selecciona un día del mes."
 ImageSelect = ImgSel
End Property

Public Property Let ImageSelect(ByVal ImageSel As Boolean)
 '* Español: Establece la imagen de selección de un día.
 '* English: Establishes the image of selection of one day.
 ImgSel = ImageSel
 PropertyChanged "ImageSelect"
End Property

Public Property Get UserImageSelect() As StdPicture
Attribute UserImageSelect.VB_Description = "Devuelve o establece una imagen de selección."
 Set UserImageSelect = myPicture
End Property

Public Property Set UserImageSelect(ByVal Imagen As StdPicture)
 '* Español: Imagen establecida por el usuario como selección.
 '* English: Image of chosen selection for the user.
 Set myPicture = Imagen
 PropertyChanged "UserImageSelect"
End Property

Public Property Get BackGroundPicture() As StdPicture
Attribute BackGroundPicture.VB_Description = "Devuelve o establece la imagen de fondo del Calendario."
 Set BackGroundPicture = myBackGround
End Property

Public Property Set BackGroundPicture(ByVal Imagen As StdPicture)
 '* Español: Establece un fondo para el Calendario.
 '* English: Establishes an image for the Calendar.
 Set myBackGround = Imagen
 PropertyChanged "BackGroundPicture"
End Property

Private Sub ShowPopUp()
 MovePos picCalendar
 picCalendar.Top = Der + 5
 picCalendar.Left = Izq
End Sub

Private Sub MovePos(ByRef Control As Object)
 Dim Rec As RECT
 
 '* Español: Devuelve o establece la posición donde se muestra la lista.
 '* English: Gets/Sets the position where the list is shown.
 GetWindowRect hWnd, Rec
 SetMousePos
 If (Rec.Bottom + (Control.Height / sngScaleX) > Screen.Height / sngScaleY) Then
  Der = (Rec.Top - (Control.Height / sngScaleY)) * sngScaleY
 Else
  Der = Rec.Bottom * sngScaleY
 End If
 If ((Rec.Right - Rec.Left) > Control.Width / sngScaleX) Then
  If (Rec.Right > Screen.Width / sngScaleX) Then
   Izq = Screen.Width - Control.Width
  Else
   Izq = Rec.Right * sngScaleX - Control.Width
  End If
  If (Izq < 0) Then Izq = 0
 Else
  If (Rec.Left < 0) Then
   Izq = 0
  Else
   Izq = Rec.Left * sngScaleX
  End If
  If (Izq + Control.Width > Screen.Width) Then Izq = Screen.Width - Control.Width
 End If
End Sub

Private Sub SetMousePos()
 sngScaleY = Screen.TwipsPerPixelY
 sngScaleX = Screen.TwipsPerPixelX
End Sub
