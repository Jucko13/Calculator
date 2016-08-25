VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F1E4D9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator by Ricardo"
   ClientHeight    =   4545
   ClientLeft      =   4065
   ClientTop       =   3345
   ClientWidth     =   8520
   BeginProperty Font 
      Name            =   "Jucko13"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8520
   WhatsThisHelp   =   -1  'True
   Begin Project1.uTextBox Text1 
      Height          =   315
      Left            =   90
      TabIndex        =   41
      Top             =   150
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   556
      BorderColor     =   16745771
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin VB.PictureBox picNormal4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   420
      Picture         =   "Form1.frx":0CCA
      ScaleHeight     =   330
      ScaleWidth      =   5040
      TabIndex        =   39
      Top             =   7770
      Width           =   5040
   End
   Begin VB.PictureBox PicHigh4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   420
      Picture         =   "Form1.frx":63AE
      ScaleHeight     =   330
      ScaleWidth      =   5040
      TabIndex        =   38
      Top             =   8190
      Width           =   5040
   End
   Begin Project1.LynxGrid List1 
      Height          =   2505
      Left            =   3375
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1470
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   4419
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorEdit   =   16777215
      BackColorSel    =   6934527
      BackColorSel2   =   13034495
      ForeColor       =   16711680
      ForeColorEdit   =   0
      ForeColorHdr    =   16711680
      ForeColorSel    =   7078143
      BackColorEvenRows=   15589589
      CustomColorFrom =   16745771
      CustomColorTo   =   15852761
      GridColor       =   16745771
      BorderStyle     =   0
      DisplayEllipsis =   0   'False
      FocusRectColor  =   9895934
      FocusRectStyle  =   0
      ThemeColor      =   5
      ThemeStyle      =   6
      CenterRowImage  =   0   'False
      ColumnHeaderLines=   2
      Caption         =   "Berekening Geshiedenis"
      ColumnHeaderSmall=   -1  'True
      ShowRowNumbers  =   -1  'True
      ShowRowNumbersVary=   -1  'True
      AllowColumnResizing=   -1  'True
      AllowWordWrap   =   -1  'True
      ColumnSort      =   -1  'True
      AllowDelete     =   -1  'True
      EditTrigger     =   2
      FocusRowHighlightStyle=   1
      HotHeaderTracking=   0   'False
      AutoToolTips    =   0   'False
   End
   Begin VB.PictureBox picNormal3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   1680
      Picture         =   "Form1.frx":BA92
      ScaleHeight     =   750
      ScaleWidth      =   540
      TabIndex        =   23
      Top             =   6510
      Width           =   540
   End
   Begin VB.PictureBox picHigh3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2310
      Picture         =   "Form1.frx":D6F6
      ScaleHeight     =   750
      ScaleWidth      =   540
      TabIndex        =   22
      Top             =   6510
      Width           =   540
   End
   Begin VB.PictureBox picHigh2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   420
      Picture         =   "Form1.frx":F35A
      ScaleHeight     =   330
      ScaleWidth      =   1170
      TabIndex        =   20
      Top             =   7350
      Width           =   1170
   End
   Begin VB.PictureBox picNormal2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   420
      Picture         =   "Form1.frx":10E6E
      ScaleHeight     =   330
      ScaleWidth      =   1170
      TabIndex        =   19
      Top             =   6930
      Width           =   1170
   End
   Begin VB.PictureBox picHigh 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1050
      Picture         =   "Form1.frx":12982
      ScaleHeight     =   330
      ScaleWidth      =   540
      TabIndex        =   11
      Top             =   6510
      Width           =   540
   End
   Begin VB.PictureBox picNormal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   420
      Picture         =   "Form1.frx":13626
      ScaleHeight     =   330
      ScaleWidth      =   540
      TabIndex        =   10
      Top             =   6510
      Width           =   540
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   1
      Left            =   105
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2715
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "1"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   2
      Left            =   735
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2715
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "2"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   3
      Left            =   1365
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2715
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "3"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   4
      Left            =   105
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2295
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "4"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   5
      Left            =   735
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2295
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "5"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   6
      Left            =   1365
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2295
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "6"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   7
      Left            =   105
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1875
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "7"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   8
      Left            =   735
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1875
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "8"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   0
      Left            =   105
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3135
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "0"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   9
      Left            =   1365
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1875
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "9"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   20
      Left            =   105
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1455
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "<--"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   15
      Left            =   1365
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1455
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Clear"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   11
      Left            =   1995
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1875
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "§("
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   12
      Left            =   1995
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2295
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "/"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   13
      Left            =   1995
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2715
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "*"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   14
      Left            =   1995
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3135
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "-"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   16
      Left            =   1365
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3135
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ","
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   750
      Index           =   10
      Left            =   2625
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2715
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   1323
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "="
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   17
      Left            =   2625
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2295
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "+"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdExtras 
      Height          =   330
      Index           =   0
      Left            =   105
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3660
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tan("
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdExtras 
      Height          =   330
      Index           =   1
      Left            =   105
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4080
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "aTan("
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdExtras 
      Height          =   330
      Index           =   2
      Left            =   735
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3660
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sin("
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdExtras 
      Height          =   330
      Index           =   3
      Left            =   735
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4080
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "aSin("
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdExtras 
      Height          =   330
      Index           =   4
      Left            =   1365
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3660
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cos("
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdExtras 
      Height          =   330
      Index           =   5
      Left            =   1365
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4080
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "aCos("
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdExtras 
      Height          =   330
      Index           =   6
      Left            =   1995
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3660
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "^"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdExtras 
      Height          =   330
      Index           =   7
      Left            =   2625
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3660
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Pi"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   18
      Left            =   1995
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4080
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "("
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   19
      Left            =   2625
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4080
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ")"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdExtras 
      Height          =   330
      Index           =   8
      Left            =   2625
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1455
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Rel"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdClearList 
      Height          =   330
      Left            =   3360
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4080
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Clear List"
      ForeColor       =   16711680
   End
   Begin Project1.TransPicBox cmdNumbers 
      Height          =   330
      Index           =   21
      Left            =   2625
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1875
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackStyle       =   1
      BackColor       =   16745771
      MaskColor       =   255
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "b/c"
      ForeColor       =   16711680
   End
   Begin Project1.uTextBox Text2 
      Height          =   570
      Left            =   90
      TabIndex        =   42
      Top             =   450
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   1005
      BorderColor     =   16745771
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      RowNumberOnEveryLine=   -1  'True
      WordWrap        =   -1  'True
      MultiLine       =   -1  'True
   End
   Begin Project1.uTextBox Text3 
      Height          =   315
      Left            =   90
      TabIndex        =   43
      Top             =   1005
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   556
      BorderColor     =   16745771
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FF852B&
      X1              =   8385
      X2              =   8385
      Y1              =   1470
      Y2              =   3990
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF852B&
      X1              =   3360
      X2              =   8385
      Y1              =   3975
      Y2              =   3975
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF852B&
      X1              =   3360
      X2              =   8400
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF852B&
      X1              =   3360
      X2              =   3360
      Y1              =   1470
      Y2              =   3990
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpslaan 
         Caption         =   "Geschiedenis Opslaan"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSerp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditOmtrek 
         Caption         =   "Omtrek van:"
         Begin VB.Menu mnuEditOmtrekCirkel 
            Caption         =   "Cirkel"
         End
         Begin VB.Menu mnuEditOmtrekDriehoek 
            Caption         =   "Driehoek"
         End
         Begin VB.Menu mnuEditOmtrekVierkant 
            Caption         =   "Vierkant"
         End
      End
      Begin VB.Menu mnuEditArea 
         Caption         =   "Oppervlakte van:"
         Begin VB.Menu mnuEditAreaCircle 
            Caption         =   "Cirkel"
         End
         Begin VB.Menu mnuEditAreaDriehoek 
            Caption         =   "Driehoek"
         End
         Begin VB.Menu mnuEditAreaSquare 
            Caption         =   "Vierhoek"
         End
         Begin VB.Menu mnuEditAreaVijfhoek 
            Caption         =   "Vijfhoek"
         End
         Begin VB.Menu mnuEditAreaZeshoek 
            Caption         =   "Zeshoek"
         End
      End
      Begin VB.Menu mnuEditInhoud 
         Caption         =   "Inhoud van:"
         Begin VB.Menu mnuEditInhoudCirkel 
            Caption         =   "Cilinder"
         End
         Begin VB.Menu mnuEditInhoudPrisma 
            Caption         =   "Prisma"
         End
         Begin VB.Menu mnuEditInhoudDriehoek 
            Caption         =   "Piramide"
         End
         Begin VB.Menu mnuEditInhoudVierkant 
            Caption         =   "Kubus"
         End
      End
      Begin VB.Menu mnuEditFormules 
         Caption         =   "Formules"
         Begin VB.Menu mnuEditFormulesABC 
            Caption         =   "ABC"
         End
      End
      Begin VB.Menu mnuEditSerp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCopyAnsware 
         Caption         =   "Kopiëer Antwoord          "
      End
      Begin VB.Menu mnuEditCopyCalc 
         Caption         =   "Kopiëer Berekening"
      End
      Begin VB.Menu mnuEditCopyBoth 
         Caption         =   "Kopiëer Bijde"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "Help"
   End
   Begin VB.Menu mnuVerberg 
      Caption         =   "Verberg"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuCalculations 
      Caption         =   "Berekeningen: -"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WM_HOTKEY As Integer = &H312

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim LastOver As Long
Dim ClickBusy As Boolean

Dim OldIndex As Integer
Dim OldIndex2 As Integer

Dim charLastAdded As Integer

Const MAX_PATH = 260&

Private Const Wortel As Long = &H221A

Private Const MOD_SINGLE_KEY As Long = &H0
Private Const MOD_SHIFT As Long = &H4

Private TotalCalculations As Long
Private objWinApi As winapi


Private Sub cmdClearList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (X < 0) Or (Y < 0) Or (X > cmdClearList.Width) Or (Y > cmdClearList.Height) Then
        ReleaseCapture
        Set cmdClearList.Picture = picNormal4.Picture

        
    ElseIf GetCapture() <> cmdClearList.hwnd Then
        SetCapture cmdClearList.hwnd
        Set cmdClearList.Picture = PicHigh4.Picture
    End If
    List1.Redraw = True
End Sub

Private Sub cmdClearList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim result As Integer

result = MsgBox("Weet je zeker dat je de lijst wilt leeg maken?", vbYesNo, "Lijst Wissen.")
If result = vbYes Then List1.Clear
End Sub

Private Sub cmdExtras_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim IIndex1 As Integer
Dim IIndex2 As Integer
Dim IIndex3 As Integer
Dim IIndex4 As Integer

IIndex1 = cmdExtras(OldIndex2).Width
IIndex2 = cmdExtras(Index).Width
IIndex3 = cmdExtras(OldIndex2).Height
IIndex4 = cmdExtras(Index).Height

    If (X < 0) Or (Y < 0) Or (X > cmdExtras(Index).Width) Or (Y > cmdExtras(Index).Height) Then
        ReleaseCapture
        If IIndex1 > 540 Then
                Set cmdExtras(OldIndex2).Picture = picNormal2.Picture
                cmdExtras(OldIndex2).ForeColor = vbBlue
        Else
            If IIndex3 > 330 Then
                Set cmdExtras(OldIndex2).Picture = picNormal3.Picture
                cmdExtras(OldIndex2).ForeColor = vbBlue
            Else
                Set cmdExtras(OldIndex2).Picture = picNormal.Picture
                cmdExtras(OldIndex2).ForeColor = vbBlue
            End If
        End If
        
    ElseIf GetCapture() <> cmdExtras(OldIndex2).hwnd Then
        SetCapture cmdExtras(Index).hwnd
        If IIndex2 > 540 Then
            Set cmdExtras(Index).Picture = picHigh2.Picture
            cmdExtras(Index).ForeColor = &H6C00FF
        Else
            If IIndex4 > 330 Then
                Set cmdExtras(Index).Picture = picHigh3.Picture
                cmdExtras(Index).ForeColor = &H6C00FF
            Else
                Set cmdExtras(Index).Picture = picHigh.Picture
                cmdExtras(Index).ForeColor = &H6C00FF
            End If
        End If
        OldIndex2 = Index
    End If
End Sub

Private Sub SubAddText(NewStr As String)
Text1.AddCharAtCursor NewStr
Text1.updateCaretPos
Text1.Redraw

'Dim selST As Long
'Dim selLn As Integer
'Dim strText As String
'
'Dim part1 As String
'Dim part2 As String
'
'Dim dimStr1 As String
'Dim dimStr2 As String
'
'selST = Text1.SelStart
'selLn = Text1.SelLength
'strText = Text1.Text
'
'If selST > 0 Then
'    dimStr1 = Mid(strText, selST, 1)
'    dimStr2 = Mid(strText, selST + 1, 1)
'End If
'
'part1 = Mid(strText, 1, selST)
'part2 = Mid(strText, selST + selLn + 1, (Len(strText) - selST - selLn))
'
'If NewStr = "/" Or NewStr = "*" Or NewStr = "-" Or NewStr = "+" Then
'    If dimStr1 = "/" Or dimStr1 = "*" Or dimStr1 = "-" Or dimStr1 = "+" Then
'        Text1.Text = Mid(strText, 1, selST - 1) & NewStr & Mid(strText, selST + 1, Len(strText) - selST - selLn)
'        Text1.SelStart = (selST)
'    ElseIf dimStr2 = "/" Or dimStr2 = "*" Or dimStr2 = "-" Or dimStr2 = "+" Then
'        Text1.Text = Mid(strText, 1, selST) & NewStr & Mid(strText, selST - 1, Len(strText) - selST - selLn)
'        Text1.SelStart = (selST)
'    Else
'        Text1.Text = part1 & NewStr & part2
'        Text1.SelStart = (selST + Len(NewStr))
'    End If
'Else
'        Text1.Text = part1 & NewStr & part2
'        Text1.SelStart = (selST + Len(NewStr))
'End If
End Sub

Private Sub cmdExtras_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0 To 7
        If Text1.Text = "0" Then
            Text1.Text = cmdExtras(Index).Caption
        Else
            Text1.Text = Text1.Text & cmdExtras(Index).Caption
        End If
        Text1.SetFocus
        'Text1.SelStart = Len(Text1)
    Case 8
        initializeScript
    Case 9
    
End Select



End Sub

Private Sub cmdNumbers_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim IIndex1 As Integer
Dim IIndex2 As Integer
Dim IIndex3 As Integer
Dim IIndex4 As Integer

IIndex1 = cmdNumbers(OldIndex).Width
IIndex2 = cmdNumbers(Index).Width
IIndex3 = cmdNumbers(OldIndex).Height
IIndex4 = cmdNumbers(Index).Height

    If (X < 0) Or (Y < 0) Or (X > cmdNumbers(Index).Width) Or (Y > cmdNumbers(Index).Height) Then
        ReleaseCapture
        If IIndex1 > 540 Then
                Set cmdNumbers(OldIndex).Picture = picNormal2.Picture
                cmdNumbers(OldIndex).ForeColor = vbBlue
        Else
            If IIndex3 > 330 Then
                Set cmdNumbers(OldIndex).Picture = picNormal3.Picture
                cmdNumbers(OldIndex).ForeColor = vbBlue
            Else
                Set cmdNumbers(OldIndex).Picture = picNormal.Picture
                cmdNumbers(OldIndex).ForeColor = vbBlue
            End If
        End If
        
    ElseIf GetCapture() <> cmdNumbers(OldIndex).hwnd Then
        SetCapture cmdNumbers(Index).hwnd
        If IIndex2 > 540 Then
            Set cmdNumbers(Index).Picture = picHigh2.Picture
            cmdNumbers(Index).ForeColor = &H40C0&
        Else
            If IIndex4 > 330 Then
                Set cmdNumbers(Index).Picture = picHigh3.Picture
                cmdNumbers(Index).ForeColor = &H40C0&
            Else
                Set cmdNumbers(Index).Picture = picHigh.Picture
                cmdNumbers(Index).ForeColor = &H40C0&
            End If
        End If
        OldIndex = Index
    End If
End Sub

Private Sub cmdNumbers_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpTx As String
Static tmpVal As String

Select Case Index
    
    Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
        PressedCalc = ""
        If Text1.Text = "0" Then
            'SubAddText Index & ""
            Text1.Text = Index
            Text1.SelStart = 1
        Else
            SubAddText Index & ""
            'Text1.Text = Text1.Text & Index
        End If
        TempStr = Text1.Text
    Case 10
    
        If Text1.Text <> "" And (Text1.Text Like "*+" Or _
            Text1.Text Like "*-" Or _
            Text1.Text Like "*[*]" Or _
            Text1.Text Like "*/") Then
                SubAddText Text2.Text
            'Text1.Text = Text1.Text & TempStr
        End If
        'Call DoIt
        Text2.Text = CheckCalculation(Text1.Text)
        tmpVal = Text2.Text
        
        If List1.CellText(0, 0) <> Text1.Text Or List1.CellText(0, 1) <> Text2.Text Then
            ReleaseCapture
            SetCapture List1.hwnd
            DoEvents
            List1.AddItem Text1.Text & Chr(9) & Text2.Text, 0
            List1.Refresh
            ReleaseCapture
            List1.Redraw = True
        End If
        
    Case 16
        SubAddText ","
        'Text1.Text = Text1.Text & ","
    Case 12, 13, 14, 17
        If Len(Text1.Text) = 0 Then
            SubAddText cmdNumbers(Index).Caption
            'Text1.Text = Text1.Text & cmdNumbers(Index).Caption
        Else
            SubAddText cmdNumbers(Index).Caption
        End If
    Case 18, 11
        If Text1.Text = "0" Then
            Text1.Text = cmdNumbers(Index).Caption
            Text1.SelStart = Len(cmdNumbers(Index).Caption)
        Else
            SubAddText cmdNumbers(Index).Caption
        End If
            

        
    Case 19
        If Text2.Text = "0" Then
            SubAddText cmdNumbers(Index).Caption
        Else
            SubAddText cmdNumbers(Index).Caption
        End If
        
    Case 20
        If Len(Text1.Text) > 0 Then
            Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 1)
        End If
        
    Case 21
        If InStr(1, Text2.Text, "/") > 0 Then
            If tmpVal <> "" Then
                Text2.Text = tmpVal
            End If
        Else
            If Val(Text2.Text) Then
                
                tmpTx = GetFraction(Round(Val(Text2.Text), 14))
                If tmpTx = Text2.Text Then
                    tmpTx = Dec2Frac(Text2.Text)
                End If
                Text2.Text = tmpTx
            End If
        End If
        
        
    Case 15
        Text1.Text = ""
        Text2.Text = "0"
        MayLog = False
        TypedText = ""
End Select

End Sub

Sub initializeScript()
    On Error GoTo Err:
    
    Set objScript = CreateObject("MSScriptControl.ScriptControl")
    objScript.Language = "VBScript"
    objScript.AddCode "setlocale ""en-us"""
    Set objWinApi = New winapi
    
    objScript.AddObject "winapi", objWinApi
    objScript.AddCode GetFileContent("functionlist.txt")
    Exit Sub
Err:
    MsgBox Err.Description, vbOKOnly Or vbCritical, "ERROR: " & Err.Number
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim TotalRows As Integer
    Dim sKeyName As String
    Dim sKeyValue As String
    Dim sKeyIcon As String
    Dim ret&
    Dim lphKey&
    Dim Reg As Object
    Dim strLast As String

    
    initializeScript
    'MsgBox objScript.Eval("Maths.asin(1)")
    
    LoadDataIntoFile 101, ("C:\windows\fonts\" & "Jucko13.ttf")
    DoEvents
    
    SetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", "Jucko13 (lol Type)", "Jucko13.ttf", REG_SZ
    
    'Me.Width = 3360
    For i = 1 To 9
        Set cmdNumbers(i).Picture = picNormal.Picture
    Next i
    
    Set cmdNumbers(11).Picture = picNormal.Picture
    Set cmdNumbers(12).Picture = picNormal.Picture
    Set cmdNumbers(13).Picture = picNormal.Picture
    Set cmdNumbers(14).Picture = picNormal.Picture
    Set cmdNumbers(16).Picture = picNormal.Picture
    Set cmdNumbers(18).Picture = picNormal.Picture
    Set cmdNumbers(19).Picture = picNormal.Picture
    Set cmdNumbers(21).Picture = picNormal.Picture
    
    Set cmdNumbers(0).Picture = picNormal2.Picture
    Set cmdNumbers(15).Picture = picNormal2.Picture
    Set cmdNumbers(20).Picture = picNormal2.Picture
    
    Set cmdNumbers(10).Picture = picNormal3.Picture
    Set cmdNumbers(17).Picture = picNormal.Picture
    
    Set cmdClearList.Picture = picNormal4.Picture
    cmdClearList.BackStyle = Transparent
    
    On Error Resume Next
    
    For i = 0 To 21
        cmdNumbers(i).BackStyle = Transparent
    Next i
    
    For i = 0 To 8
        cmdExtras(i).BackStyle = Transparent
        Set cmdExtras(i).Picture = picNormal.Picture
    Next i
    'Set cmdNumbers(i).Picture = picNormal.Picture
    
    List1.AddColumn "Berekening", List1.Width / 2 - 270, lgAlignCenterCenter, lgString, , , , True, , True, False
    List1.AddColumn "Antwoord", List1.Width / 2 - 270, lgAlignCenterCenter, lgString, , , , True, , True, True
    
    List1.BackColorBkg = &HF1E4D9
    List1.ForeColorHdr = &HFF0000
    List1.Refresh
    List1.Redraw = True
    Text1.SelStart = 1
    
    TotalRows = GetSetting("Calculator", "Berekeningen", "Rows", -1)
    Text1.Text = GetSetting("Calculator", "Berekeningen", "Text1.text", 0)
    Text2.Text = GetSetting("Calculator", "Berekeningen", "Text2.text", 0)
    Text3.Text = GetSetting("Calculator", "Berekeningen", "Text3.text", 0)
    Text1.SelStart = GetSetting("Calculator", "Berekeningen", "Text1.SelStart", 0)
    Text2.SelStart = GetSetting("Calculator", "Berekeningen", "Text2.SelStart", 0)
    
    'cmdExtras_MouseUp 8, 0, 0, 0, 0
    If TotalRows > -1 Then
        For i = 0 To TotalRows
            List1.AddItem GetSetting("Calculator", "Berekeningen", "Row" & i)
        Next i
    End If
End Sub


Function CheckCalculation(CalculateString As String, Optional ParentCall As Boolean = True) As String
    Dim result As Variant
    
    
    On Error GoTo EndIt
    objScript.AddCode Text3.Text
    
    result = objScript.Eval(Replace(CalculateString, "§", "sqr"))
    If TypeName(result) = "Double" Then
        CheckCalculation = Replace(result, ",", ".")
    Else
        CheckCalculation = result
    End If

    If ParentCall = True Then
        mnuCalculations.Caption = "Berekeningen: " & TotalCalculations
    End If

    Exit Function
EndIt:
    
    If Err.Number = 6 Or Err.Number = 1031 Then
        CheckCalculation = "Error: Overflow"
    Else
        CheckCalculation = "Error: " & Err.Description & " [" & Err.Number & "]"
    End If

End Function


Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer

RemoveKeyboardHook

SaveSetting "Calculator", "Berekeningen", "Rows", List1.Rows

SaveSetting "Calculator", "Berekeningen", "Text1.Text", Text1.Text
SaveSetting "Calculator", "Berekeningen", "Text2.Text", Text2.Text
SaveSetting "Calculator", "Berekeningen", "Text3.Text", Text3.Text

SaveSetting "Calculator", "Berekeningen", "Text1.SelStart", Text1.SelStart
SaveSetting "Calculator", "Berekeningen", "Text2.SelStart", Text2.SelStart

For i = 0 To List1.Rows
    SaveSetting "Calculator", "Berekeningen", "Row" & i, List1.CellText(i, 0) & Chr(9) & List1.CellText(i, 1)
Next i
End Sub



Private Sub List1_DblClick()
Dim RowMouse As Long
Dim RowStr() As String

RowMouse = List1.MouseRow
Text1.Text = List1.CellText(RowMouse, 0)
Text2.Text = List1.CellText(RowMouse, 1)
List1.Redraw = True
End Sub

Private Sub List1_ItemAdded(ByVal Row As Long)
'List1.SBVisible(efsVertical) = True
End Sub

Private Sub List1_MouseEnter()
'List1.GridColor = &HDBFF&
End Sub

Private Sub List1_MouseLeave()
'List1.GridColor = &HFF852B
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim RowMouse As Long
Dim i As Integer


RowMouse = List1.MouseRow
If LastOver = RowMouse Then Exit Sub

If RowMouse > -1 Then
    For i = 0 To List1.Rows
        If List1.RowSelected(i) = True Then List1.RowSelected(i) = False
    Next i
    List1.RowSelected(List1.MouseRow) = True
Else
    For i = 0 To List1.Rows
        If List1.RowSelected(i) = True Then List1.RowSelected(i) = False
    Next i
End If
LastOver = RowMouse
End Sub

Private Sub mnuAbout_Click()
MsgBox "Dit is een simpele RekenMachine die geschreven is in Visual Basics 6.0" & vbCrLf & _
       "Gecodeert door: Ricardo de Roode." & vbCrLf & _
       vbCrLf & _
       "How To Use:" & vbCrLf & _
       "        - Press ""Shift+9"" / ""("" to start loggin your calculation." & vbCrLf & _
       "        - Press ""Shift+Enter"" to Finnish calculation, remove the " & vbCrLf & _
       "          calculation, and replace it with the answer." & vbCrLf & _
       "        - Press ""Escape"" to Reset the logged calculation." & vbCrLf & _
       vbCrLf & _
       "if you typed something wrong you can just press ""BackSpace"".", vbInformation, "About"
        
End Sub

Private Sub mnuEditAreaCircle_Click()
Dim mm As String

mm = InputBox("Hier de Diameter", "Oppervlakte van een: " & mnuEditAreaCircle.Caption)

If mm <> "" Then
    Text1.Text = "(0.25*Pi)*" & mm & "*" & mm
    Text2.Text = CheckCalculation(Text1.Text)
End If
End Sub

Private Sub mnuEditAreaDriehoek_Click()
Dim mm1 As String
Dim mm2 As String

mm1 = InputBox("Hier de Basis", "Oppervlakte van een: " & mnuEditAreaSquare.Caption)
mm2 = InputBox("Hier de hoogte", "Oppervlakte van een: " & mnuEditAreaSquare.Caption)

If mm1 <> "" And mm2 <> "" Then
    Text1.Text = "(" & mm1 & "*" & mm2 & ")/2"
    Text2.Text = CheckCalculation(Text1.Text)
End If
End Sub

Private Sub mnuEditAreaSquare_Click()
Dim mm1 As String
Dim mm2 As String

mm1 = InputBox("Hier de breedte", "Oppervlakte van een: " & mnuEditAreaSquare.Caption)
mm2 = InputBox("Hier de hoogte", "Oppervlakte van een: " & mnuEditAreaSquare.Caption)

If mm1 <> "" And mm2 <> "" Then
    Text1.Text = "(" & mm1 & "*" & mm2 & ")"
    Text2.Text = CheckCalculation(Text1.Text)
End If
End Sub

Private Sub mnuEditAreaVijfhoek_Click()
Dim mm1 As String

mm1 = InputBox("Hier de lengte van 1 zijde", "Oppervlakte van een: " & mnuEditAreaZeshoek.Caption)

If mm1 <> "" Then
    Text1.Text = "(2,5*" & mm1 & "*§(§(3)-3)"
    Text2.Text = CheckCalculation(Text1.Text)
End If
End Sub

Private Sub mnuEditAreaZeshoek_Click()
Dim mm1 As String

mm1 = InputBox("Hier de lengte van 1 zijde", "Oppervlakte van een: " & mnuEditAreaZeshoek.Caption)

If mm1 <> "" Then
    Text1.Text = "((3/2)*" & mm1 & "^2)*§(3)"
    Text2.Text = CheckCalculation(Text1.Text)
End If
End Sub

Private Sub mnuEditCopyAnsware_Click()
Clipboard.Clear
Clipboard.SetText Text2.Text
End Sub

Private Sub mnuEditCopyBoth_Click()

Clipboard.Clear
Clipboard.SetText Text1.Text & " = " & Text2.Text
End Sub

Private Sub mnuEditCopyCalc_Click()
Clipboard.Clear
Clipboard.SetText Text1.Text
End Sub

Private Sub mnuEditFormulesABC_Click()
Dim a As String
Dim b As String
Dim c As String

a = InputBox("Geef waarde voor A:")
b = InputBox("Geef waarde voor B:")
c = InputBox("Geef waarde voor C:")

Text1.Text = "abc(" & a & "," & b & "," & c & ")"
cmdNumbers_MouseUp 10, 0, 0, 0, 0
End Sub

Private Sub mnuEditInhoudPrisma_Click()
Dim mm1 As String
Dim mm2 As String
Dim mm3 As String

mm1 = InputBox("Hier de Breedte", "Oppervlakte van een: " & mnuEditInhoudVierkant.Caption)
mm2 = InputBox("Hier de Hoogte", "Oppervlakte van een: " & mnuEditOmtrekDriehoek.Caption)
mm3 = InputBox("Hier de Diepte", "Oppervlakte van een: " & mnuEditOmtrekDriehoek.Caption)

If mm1 <> "" And mm2 <> "" Then
    Text1.Text = "(" & mm1 & "*" & mm2 & "*" & mm3 & ")"
    Text2.Text = CheckCalculation(Text1.Text)
End If
End Sub

Private Sub mnuEditInhoudVierkant_Click()
Dim mm1 As String
Dim mm2 As String
Dim mm3 As String

mm1 = InputBox("Hier de Breedte", "Oppervlakte van een: " & mnuEditInhoudVierkant.Caption)
mm2 = InputBox("Hier de Hoogte", "Oppervlakte van een: " & mnuEditOmtrekDriehoek.Caption)
mm3 = InputBox("Hier de Diepte", "Oppervlakte van een: " & mnuEditOmtrekDriehoek.Caption)

If mm1 <> "" And mm2 <> "" Then
    Text1.Text = "(" & mm1 & "*" & mm2 & "*" & mm3 & ")"
    Text2.Text = CheckCalculation(Text1.Text)
End If
End Sub

Private Sub mnuEditOmtrekCirkel_Click()
Dim mm1 As String

mm1 = InputBox("Diameter van de Circel:", "Omtrek van een: " & mnuEditOmtrekCirkel.Caption)

If mm1 <> "" Then
    Text1.Text = "(2" & "*" & "Pi*" & (mm1 / 2) & ")"
    Text2.Text = CheckCalculation(Text1.Text)
End If
End Sub

Private Sub mnuEditOmtrekDriehoek_Click()
Dim mm1 As String
Dim mm2 As String
Dim mm3 As String

mm1 = InputBox("Hier Zijde 1", "Oppervlakte van een: " & mnuEditOmtrekDriehoek.Caption)
mm2 = InputBox("Hier Zijde 2", "Oppervlakte van een: " & mnuEditOmtrekDriehoek.Caption)
mm3 = InputBox("Hier Zijde 3", "Oppervlakte van een: " & mnuEditOmtrekDriehoek.Caption)

If mm1 <> "" And mm2 <> "" Then
    Text1.Text = "(" & mm1 & "+" & mm2 & "+" & mm3 & ")"
    Text2.Text = CheckCalculation(Text1.Text)
End If
End Sub

Private Sub mnuEditOmtrekVierkant_Click()
Dim mm1 As String
Dim mm2 As String

mm1 = InputBox("Hier de breedte", "Oppervlakte van een: " & mnuEditOmtrekVierkant.Caption)
mm2 = InputBox("Hier de hoogte", "Oppervlakte van een: " & mnuEditOmtrekVierkant.Caption)

If mm1 <> "" And mm2 <> "" Then
    Text1.Text = "((" & mm1 & "+" & mm2 & ")*2)"
    Text2.Text = CheckCalculation(Text1.Text)
End If
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub Text1_Change()
'If Len(Text1.Text) = 0 Then
'    Text1.Text = 0
'    Text1.SelStart = 1
'End If
'If Len(Text1.Text) > 1 Then
'    If Mid(Text1.Text, 1, 1) = "0" Then
'        Text1.Text = Mid(Text1.Text, 2, Len(Text1.Text) - 1)
'        Text1.SelStart = 1
'    End If
'End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdNumbers_MouseUp 10, 0, 0, 0, 0
    KeyAscii = 0
End If
End Sub

Private Sub Text2_Changed()
formatTextBox Text2
End Sub

Private Sub Text3_Changed()
    formatTextBox Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdNumbers_MouseUp 10, 0, 0, 0, 0
    KeyAscii = 0
End If
End Sub

Sub formatTextBox(txt As uTextBox)
    Dim fColors(0 To 4) As Long
    'Dim bColors(0 To 6) As Long
    
    fColors(0) = RGB(60, 140, 255)
    fColors(1) = RGB(255, 126, 0)
    fColors(2) = RGB(67, 200, 0)
    fColors(3) = RGB(191, 112, 0)
    fColors(4) = RGB(170, 98, 255)
    
'    bColors(0) = RGB()
'    bColors(1) = -1
'    bColors(2) = -1
'    bColors(3) = -1
'    bColors(4) = -1
'    bColors(5) = -1
'    bColors(6) = -1
    
    Dim i As Long
    Dim OT As Long 'opened tags
    Dim TT As Long 'total tags
    Dim s As String
    Dim t As String
    Dim lstart As Long
    Dim lend As Long
    Dim lStep As Long
    
    Dim BT As Long 'bold tag
    Dim CFT As Boolean 'colored first tag
    Dim CS As Boolean 'color string
    
    
    s = txt.Text
    
    
    For i = 1 To Len(s)
        t = Mid$(s, i, 1)
        
        txt.setCharBold i - 1, False
        txt.setCharBackColor i - 1, -1
        txt.setCharForeColor i - 1, IIf(CS, fColors(3), -1)
        
        If t = Chr(34) Then
            CS = Not CS
            txt.setCharForeColor i - 1, fColors(3)
        End If
        
        If Not CS Then
            Select Case t
                Case "(", ")"
                    If t = ")" Then
                        OT = OT - 1
                    End If
                    
                    txt.setCharBold i - 1, True
                    
                    If OT < 0 Then
                        txt.setCharForeColor i - 1, vbRed
                    Else
                        txt.setCharForeColor i - 1, fColors(OT Mod (UBound(fColors) + 1))
                    End If
    
                    If t = "(" Then
                        OT = OT + 1
                    End If
    
                Case "0" To "9", "."
                    txt.setCharForeColor i - 1, vbMagenta
                
                Case "/", "-", "+", "*"
                    txt.setCharForeColor i - 1, RGB(100, 100, 100)
                    
                Case Else
            End Select
        End If
        
    Next i
    
    txt.Redraw
End Sub


Private Sub Text1_Changed()

    formatTextBox Text1
    
End Sub

Private Sub Text1_SelectionChanged()
Text1_Changed
End Sub

