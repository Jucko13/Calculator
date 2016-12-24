VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F1E4D9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator by Ricardo"
   ClientHeight    =   4545
   ClientLeft      =   4065
   ClientTop       =   3345
   ClientWidth     =   8505
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   303
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   567
   WhatsThisHelp   =   -1  'True
   Begin Project1.uListBox lstComplete 
      Height          =   2460
      Left            =   3420
      TabIndex        =   45
      Top             =   400
      Visible         =   0   'False
      Width           =   2350
      _ExtentX        =   4154
      _ExtentY        =   4339
      BorderColor     =   16745771
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Text            =   ""
      SelectionBackgroundColor=   15852761
      SelectionBorderColor=   16761024
      SelectionForeColor=   16711680
      ItemHeight      =   32
   End
   Begin VB.Timer tmrFly 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3060
      Top             =   2235
   End
   Begin Project1.uTextBox txtFly 
      Height          =   220
      Left            =   3200
      TabIndex        =   44
      Top             =   3710
      Visible         =   0   'False
      Width           =   60
      _ExtentX        =   529
      _ExtentY        =   397
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
      HideCursor      =   -1  'True
      AutoResize      =   -1  'True
   End
   Begin MSComDlg.CommonDialog comm1 
      Left            =   3165
      Top             =   1335
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.uButton cmdClearList 
      Height          =   330
      Left            =   3360
      TabIndex        =   12
      Top             =   4080
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      CaptionBorderColor=   14737632
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "Clear"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":0CCA
      PictureMouseOver=   "Form1.frx":1074
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uListBox List1 
      Height          =   2535
      Left            =   3360
      TabIndex        =   11
      Top             =   1440
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   4471
      BackgroundColor =   15852761
      BorderColor     =   16745771
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      Text            =   "uFrame"
      SelectionBackgroundColor=   16745771
      SelectionBorderColor=   16745771
      SelectionForeColor=   16711680
      ItemHeight      =   33
   End
   Begin Project1.uTextBox Text1 
      Height          =   315
      Left            =   90
      TabIndex        =   0
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
      MousePointer    =   3
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
      Picture         =   "Form1.frx":17E2
      ScaleHeight     =   330
      ScaleWidth      =   5040
      TabIndex        =   8
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
      Picture         =   "Form1.frx":6EC6
      ScaleHeight     =   330
      ScaleWidth      =   5040
      TabIndex        =   7
      Top             =   8190
      Width           =   5040
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
      Picture         =   "Form1.frx":C5AA
      ScaleHeight     =   750
      ScaleWidth      =   540
      TabIndex        =   6
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
      Picture         =   "Form1.frx":E20E
      ScaleHeight     =   750
      ScaleWidth      =   540
      TabIndex        =   5
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
      Picture         =   "Form1.frx":FE72
      ScaleHeight     =   330
      ScaleWidth      =   1170
      TabIndex        =   4
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
      Picture         =   "Form1.frx":11986
      ScaleHeight     =   330
      ScaleWidth      =   1170
      TabIndex        =   3
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
      Picture         =   "Form1.frx":1349A
      ScaleHeight     =   330
      ScaleWidth      =   540
      TabIndex        =   2
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
      Picture         =   "Form1.frx":1413E
      ScaleHeight     =   330
      ScaleWidth      =   540
      TabIndex        =   1
      Top             =   6510
      Width           =   540
   End
   Begin Project1.uTextBox Text2 
      Height          =   570
      Left            =   90
      TabIndex        =   9
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
      MousePointer    =   3
      RowNumberOnEveryLine=   -1  'True
      WordWrap        =   -1  'True
      MultiLine       =   -1  'True
   End
   Begin Project1.uTextBox Text3 
      Height          =   315
      Left            =   90
      TabIndex        =   10
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
      MousePointer    =   3
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   20
      Left            =   105
      TabIndex        =   13
      Top             =   1440
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "<---"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":14DE2
      PictureMouseOver=   "Form1.frx":1501F
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   15
      Left            =   1365
      TabIndex        =   14
      Top             =   1440
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "Clear"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":155C9
      PictureMouseOver=   "Form1.frx":15806
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdExtras 
      Height          =   330
      Index           =   8
      Left            =   2625
      TabIndex        =   15
      Top             =   1440
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "Rel"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":15DB0
      PictureMouseOver=   "Form1.frx":15F0F
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   21
      Left            =   2625
      TabIndex        =   16
      Top             =   1860
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "b/c"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1643E
      PictureMouseOver=   "Form1.frx":1659D
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   17
      Left            =   2625
      TabIndex        =   17
      Top             =   2280
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "+"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":16ACC
      PictureMouseOver=   "Form1.frx":16C2B
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   7
      Left            =   105
      TabIndex        =   18
      Top             =   1860
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "7"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1715A
      PictureMouseOver=   "Form1.frx":172B9
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   1
      Left            =   105
      TabIndex        =   19
      Top             =   2700
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "1"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":177E8
      PictureMouseOver=   "Form1.frx":17947
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   2
      Left            =   735
      TabIndex        =   20
      Top             =   2700
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "2"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":17E76
      PictureMouseOver=   "Form1.frx":17FD5
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   3
      Left            =   1365
      TabIndex        =   21
      Top             =   2700
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "3"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":18504
      PictureMouseOver=   "Form1.frx":18663
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   4
      Left            =   105
      TabIndex        =   22
      Top             =   2280
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "4"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":18B92
      PictureMouseOver=   "Form1.frx":18CF1
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   5
      Left            =   735
      TabIndex        =   23
      Top             =   2280
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "5"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":19220
      PictureMouseOver=   "Form1.frx":1937F
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   6
      Left            =   1365
      TabIndex        =   24
      Top             =   2280
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "6"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":198AE
      PictureMouseOver=   "Form1.frx":19A0D
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   8
      Left            =   735
      TabIndex        =   25
      Top             =   1860
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "8"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":19F3C
      PictureMouseOver=   "Form1.frx":1A09B
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   9
      Left            =   1365
      TabIndex        =   26
      Top             =   1860
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "9"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1A5CA
      PictureMouseOver=   "Form1.frx":1A729
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   11
      Left            =   1995
      TabIndex        =   27
      Top             =   1860
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "§("
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1AC58
      PictureMouseOver=   "Form1.frx":1ADB7
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   12
      Left            =   1995
      TabIndex        =   28
      Top             =   2280
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "/"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1B2E6
      PictureMouseOver=   "Form1.frx":1B445
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   13
      Left            =   1995
      TabIndex        =   29
      Top             =   2700
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "*"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1B974
      PictureMouseOver=   "Form1.frx":1BAD3
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   14
      Left            =   1995
      TabIndex        =   30
      Top             =   3120
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "-"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1C002
      PictureMouseOver=   "Form1.frx":1C161
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   16
      Left            =   1365
      TabIndex        =   31
      Top             =   3120
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "."
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1C690
      PictureMouseOver=   "Form1.frx":1C7EF
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   0
      Left            =   105
      TabIndex        =   42
      Top             =   3120
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "0"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1CD1E
      PictureMouseOver=   "Form1.frx":1CF5B
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   750
      Index           =   10
      Left            =   2630
      TabIndex        =   43
      Top             =   2700
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   1323
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "="
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1D505
      PictureMouseOver=   "Form1.frx":1D6DE
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdExtras 
      Height          =   330
      Index           =   0
      Left            =   105
      TabIndex        =   32
      Top             =   3660
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "Tan"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1DD6E
      PictureMouseOver=   "Form1.frx":1DECD
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdExtras 
      Height          =   330
      Index           =   1
      Left            =   105
      TabIndex        =   33
      Top             =   4080
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "aTn"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1E3FC
      PictureMouseOver=   "Form1.frx":1E55B
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdExtras 
      Height          =   330
      Index           =   2
      Left            =   735
      TabIndex        =   34
      Top             =   3660
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "Sin"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1EA8A
      PictureMouseOver=   "Form1.frx":1EBE9
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdExtras 
      Height          =   330
      Index           =   3
      Left            =   735
      TabIndex        =   35
      Top             =   4080
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "aSn"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1F118
      PictureMouseOver=   "Form1.frx":1F277
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdExtras 
      Height          =   330
      Index           =   4
      Left            =   1365
      TabIndex        =   36
      Top             =   3660
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "Cos"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1F7A6
      PictureMouseOver=   "Form1.frx":1F905
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdExtras 
      Height          =   330
      Index           =   5
      Left            =   1365
      TabIndex        =   37
      Top             =   4080
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "aCs"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":1FE34
      PictureMouseOver=   "Form1.frx":1FF93
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdExtras 
      Height          =   330
      Index           =   6
      Left            =   1995
      TabIndex        =   38
      Top             =   3660
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "^"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":204C2
      PictureMouseOver=   "Form1.frx":20621
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdExtras 
      Height          =   330
      Index           =   7
      Left            =   2625
      TabIndex        =   39
      Top             =   3660
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "PI"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":20B50
      PictureMouseOver=   "Form1.frx":20CAF
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   18
      Left            =   1995
      TabIndex        =   40
      Top             =   4080
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   "("
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":211DE
      PictureMouseOver=   "Form1.frx":2133D
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   19
      Left            =   2625
      TabIndex        =   41
      Top             =   4080
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      BackgroundColor =   15852761
      BorderColor     =   16745771
      ForeColor       =   16711680
      MouseOverBackgroundColor=   15852761
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      Caption         =   ")"
      Border          =   0   'False
      BorderAnimation =   0
      Picture         =   "Form1.frx":2186C
      PictureMouseOver=   "Form1.frx":219CB
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileHighDPI 
         Caption         =   "High DPI"
      End
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
         Caption         =   "Kopiëer Beide"
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
   Begin VB.Menu mnuExecTime 
      Caption         =   "ExecTime: -"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type GUITHREADINFO
    cbSize As Long
    flags As Long
    hwndActive As Long
    hwndFocus As Long
    hwndCapture As Long
    hwndMenuOwner As Long
    hwndMoveSize As Long
    hwndCaret As Long
    rcCaret As RECT
End Type
 
Private Declare Function GetGUIThreadInfo Lib "user32" (ByVal hThreadId As Long, pGuiThreadInfo As GUITHREADINFO) As Long
 
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long


Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As uGlobals.POINTAPI) As Long
  

Private Const WM_HOTKEY As Integer = &H312

Private objWinApi As winapi


Private Sub cmdClearList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If (x < 0) Or (y < 0) Or (x > cmdClearList.Width) Or (y > cmdClearList.Height) Then
    '    ReleaseCapture
    '    Set cmdClearList.Picture = picNormal4.Picture

        
    'ElseIf GetCapture() <> cmdClearList.hWnd Then
    '    SetCapture cmdClearList.hWnd
    '    Set cmdClearList.Picture = PicHigh4.Picture
    'End If
    'List1.Redraw = True
End Sub

Private Sub cmdClearList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim result As Integer
    Dim i As Long
    On Error Resume Next
    
    result = MsgBox("Weet je zeker dat je de lijst wilt leeg maken?", vbYesNo, "Lijst Wissen.")
    If result = vbYes Then
        For i = 0 To List1.ListCount - 1
            DeleteSetting "Calculator", "Berekeningen", "Row" & i
        Next i

        SaveSetting "Calculator", "Berekeningen", "Rows", 0
        
        List1.Clear
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

Private Sub cmdExtras_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim selection As Long

Select Case Index
    Case 0 To 6
        selection = Text1.SelStart
        Text1.AddCharAtCursor cmdExtras(Index).Caption & "()"
        Text1.SelStart = selection + Len(cmdExtras(Index).Caption) + 1

    Case 7
        Text1.AddCharAtCursor cmdExtras(Index).Caption
        
    Case 8
        initializeScript
        Text1_Changed
        Text2_Changed
        Text3_Changed
    Case 9
    
End Select

Text1.SetFocus


End Sub


Private Sub cmdNumbers_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
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
    
        Text2.Text = CheckCalculation(Text1.Text)
        tmpVal = Text2.Text
        
        If List1.ListCount > 0 Then
            If List1.Cell(0, 0) <> Text1.Text Or List1.Cell(0, 1) <> vbCrLf & Text2.Text Then
                List1.AddItem Text1.Text & Chr(9) & vbCrLf & Text2.Text, , 0
            End If
        Else
            List1.AddItem Text1.Text & Chr(9) & vbCrLf & Text2.Text
        End If
        Text1.SetFocus
        
        
    Case 16
        SubAddText ","
        'Text1.Text = Text1.Text & ","
    Case 11, 12, 13, 14, 17, 18
        SubAddText cmdNumbers(Index).Caption
        
        
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
        Text2.Text = ""
        MayLog = False
        TypedText = ""
        
End Select

End Sub

Sub initializeScript()
    On Error GoTo Err:
    Dim f As String
    Dim i As Long
    Dim tmpLines() As String
    Dim t As String
    Dim c As Long
    
    ReDim ExternalFunctions(0) As String
    ReDim ExternalCustomFunctions(0) As String
    ReDim ExternalConstants(0) As String
    ReDim ExternalOperators(0) As String
    
    f = GetFileContent(App.Path & "\functionlist.txt")
    
    Set objScript = CreateObject("MSScriptControl.ScriptControl")
    objScript.Language = "VBScript"
    objScript.AddCode "setlocale ""en-us"""
    Set objWinApi = New winapi
    objWinApi.initialize comm1
    
    objScript.AddObject "winapi", objWinApi
    objScript.AddCode f
    objScript.AddCode "function help(): winapi.help(): end function "
    objScript.AllowUI = True
    
    '
    
    ExternalFunctions = Split("ans abs array asc atn cbool cbyte ccur cdate cdbl chr cint clng conversions cos createobject csng cstr date dateadd datediff datepart dateserial datevalue day escape eval exp filter formatcurrency formatdatetime formatnumber formatpercent getlocale getobject getref hex hour inputbox instr instrrev int fix isarray isdate isempty isnull isnumeric isobject join lbound lcase left len loadpicture log ltrim rtrim trim maths mid minute month monthname msgbox now oct replace rgb right rnd round scriptengine scriptenginebuildversion scriptenginemajorversion scriptengineminorversion second setlocale sgn sin space split sqr strcomp string strreverse tan time timer timeserial timevalue typename ubound ucase unescape vartype weekday weekdayname year", " ")
    
    c = 0
    
    ReDim ExternalCustomFunctions(0 To 7) As String

    ExternalCustomFunctions(0) = "winapi"
    ExternalCustomFunctions(1) = "colorpicker"
    ExternalCustomFunctions(2) = "help"
    ExternalCustomFunctions(3) = "longtorgb"
    ExternalCustomFunctions(4) = "gettickcount"
    ExternalCustomFunctions(5) = "findwindow"
    ExternalCustomFunctions(6) = "commondialog"
    ExternalCustomFunctions(7) = "showcommands"

    c = UBound(ExternalCustomFunctions) + 1
    
    tmpLines = Split(LCase(f), vbCrLf)
    For i = 0 To UBound(tmpLines)
        t = Text1.GetMidText(tmpLines(i), "function ", "(")
        
        If t <> "" Then
            ReDim Preserve ExternalCustomFunctions(0 To c) As String
            ExternalCustomFunctions(c) = t
            c = c + 1
        End If
    Next i
    
    MergeSort ExternalCustomFunctions
    
    
    ExternalConstants = Split("vbabortretryignore vbapplicationmodal vbarray vbblack vbblue vbboolean vbbyte vbcr vbcritical vbcrlf vbcurrency vbcyan vbdataobject vbdate vbdecimal vbdefaultbutton1 vbdefaultbutton2 vbdefaultbutton3 vbdefaultbutton4 vbdouble vbempty vberror vbexclamation vbfalse vbformfeed vbgreen vbinformation vbinteger vblf vblong vbmagenta vbnewline vbnull vbnullchar vbnullstring vbobject vbokcancel vbokonly vbquestion vbred vbretrycancel vbsingle vbstring vbsystemmodal vbtab vbtrue vbusedefault vbvariant vbverticaltab vbwhite vbyellow vbyesno vbyesnocancel vbbinarycompare vbtextcompare", " ")
    
    
    ExternalOperators = Split("xor and or not is * - + / :", " ")
    
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
    
    SetKeyboardHook
    
    'txtFly.Visible = False
    
    SetParent txtFly.hWnd, GetParent(Me.hWnd)
    SetTopMostWindow txtFly.hWnd, True
    
    initializeScript
    'MsgBox objScript.Eval("Maths.asin(1)")
    
    'LoadDataIntoFile 101, ("C:\windows\fonts\" & "Jucko13.ttf")
    'DoEvents
    
    'SetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", "Jucko13 (lol Type)", "Jucko13.ttf", REG_SZ
    
    On Error Resume Next
    
    List1.Redraw
    
    TotalRows = GetSetting("Calculator", "Berekeningen", "Rows", -1)
    Text1.Text = GetSetting("Calculator", "Berekeningen", "Text1.text", 0)
    Text2.Text = GetSetting("Calculator", "Berekeningen", "Text2.text", 0)
    Text3.Text = GetSetting("Calculator", "Berekeningen", "Text3.text", "")
    Text1.SelStart = GetSetting("Calculator", "Berekeningen", "Text1.SelStart", 0)
    Text2.SelStart = GetSetting("Calculator", "Berekeningen", "Text2.SelStart", 0)
    
    mnuFileHighDPI.Checked = GetSetting("Calculator", "Settings", "high dpi", False)
    If mnuFileHighDPI.Checked Then ApplyDPI

    If TotalRows > -1 Then
        For i = 0 To TotalRows - 1
            List1.AddItem GetSetting("Calculator", "Berekeningen", "Row" & i)
        Next i
    End If
End Sub


Function CheckCalculation(CalculateString As String, Optional ParentCall As Boolean = True) As String
    Dim result As Variant
    Dim t As clsTimer
    Dim tend As Double
    Dim allfunctions As String
    
    Set t = New clsTimer
    
    On Error GoTo EndIt
    
    t.tStart
    
    initializeScript
    
    If InStr(1, LCase$(CalculateString), "ans") > 0 Then
        objScript.AddCode "dim ans: ans = " & Text2.Text
    End If
    
    objScript.AddCode Text3.Text
    
    'allfunctions = CharExecution(objScript.CodeObject, False)
    'allfunctions = allfunctions & CharExecution(objScript.CodeObject.winapi, True)
    
    'MsgBox allfunctions
    ' MsgBox
    result = objScript.Eval(Replace(CalculateString, "§", "sqr"))
    If TypeName(result) = "Double" Then
        CheckCalculation = Replace(result, ",", ".")
    Else
        CheckCalculation = result
    End If
    
    tend = t.tStop
    mnuExecTime.Caption = "ExecTime: " & mircoToTime(tend)

    Exit Function
EndIt:
    
    If Err.Number = 6 Or Err.Number = 1031 Then
        CheckCalculation = "Error: Overflow"
    Else
        CheckCalculation = "Error: " & Err.Description & " [" & Err.Number & "]"
    End If
    
    t.tStop
    
    mnuExecTime.Caption = "ExecTime: -"
End Function

Function mircoToTime(ByVal lTime As Double) As String
    Dim ltimes As Long
    Dim newTime As Double
    
    If lTime = 0 Then mircoToTime = "Instant": Exit Function
    
    Do While lTime < 1
        ltimes = ltimes + 1
        lTime = lTime * 1000
    Loop
    
    newTime = Round(lTime, 4)
    
    Select Case ltimes
        Case 0
            mircoToTime = newTime & " s"
            
        Case 1
            mircoToTime = newTime & " ms"
        
        Case 2
            mircoToTime = newTime & " us"
        
        Case Else
            mircoToTime = newTime & " ?"
    End Select
    
    
End Function


Sub ApplyDPI()
    Dim c As Control
    
    For Each c In Me.Controls
        If TypeName(c) = "uButton" Then
            c.Width = c.Width - IIf(mnuFileHighDPI.Checked, 1, -1)
            c.Height = c.Height - IIf(mnuFileHighDPI.Checked, 1, -1)
        End If
    Next c
    
    
    List1.ScrollBarWidth = IIf(mnuFileHighDPI.Checked, 30, 20)
    List1.setTabStop 1, List1.Width - List1.ScrollBarWidth - 4, vbRightJustify
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Long

RemoveKeyboardHook

SaveSetting "Calculator", "Berekeningen", "Rows", List1.ListCount

SaveSetting "Calculator", "Berekeningen", "Text1.Text", Text1.Text
SaveSetting "Calculator", "Berekeningen", "Text2.Text", Text2.Text
SaveSetting "Calculator", "Berekeningen", "Text3.Text", Text3.Text

SaveSetting "Calculator", "Berekeningen", "Text1.SelStart", Text1.SelStart
SaveSetting "Calculator", "Berekeningen", "Text2.SelStart", Text2.SelStart

For i = 0 To List1.ListCount - 1
    SaveSetting "Calculator", "Berekeningen", "Row" & i, List1.Cell(i, 0) & Chr(9) & List1.Cell(i, 1)
Next i
End Sub



Private Sub List1_DblClick()
Dim RowMouse As Long
Dim RowStr() As String

RowMouse = List1.ListIndex
Text1.Text = List1.Cell(RowMouse, 0)
Text2.Text = Replace(List1.Cell(RowMouse, 1), vbCrLf, "")
End Sub

Private Sub List1_ItemAdded(ItemIndex As Long)
    List1.RedrawPause
    Dim i As Long
    
    For i = 0 To List1.ListCount - 1
        If i Mod 2 = 0 Then
            List1.ItemColor(i) = RGB(241, 244, 250)
        Else
            List1.ItemColor(i) = &HF1E4D9
        End If
    Next i
    
    List1.RedrawResume
End Sub

Private Sub lstComplete_DblClick()
    If lstComplete.ListIndex <> -1 And lstComplete.ListCount > 0 Then
        Text1.ReplaceWord lstComplete.List(lstComplete.ListIndex)
    End If
End Sub

'Private Sub List1_MouseEnter()
''List1.GridColor = &HDBFF&
'End Sub
'
'Private Sub List1_MouseLeave()
''List1.GridColor = &HFF852B
'End Sub

'Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Dim RowMouse As Long
'Dim i As Integer
'
'
'RowMouse = List1.MouseRow
'If LastOver = RowMouse Then Exit Sub
'
'If RowMouse > -1 Then
'    For i = 0 To List1.Rows
'        If List1.RowSelected(i) = True Then List1.RowSelected(i) = False
'    Next i
'    List1.RowSelected(List1.MouseRow) = True
'Else
'    For i = 0 To List1.Rows
'        If List1.RowSelected(i) = True Then List1.RowSelected(i) = False
'    Next i
'End If
'LastOver = RowMouse
'End Sub

Private Sub mnuAbout_Click()
MsgBox "This Calculator is made in Visual Basics 6.0 and uses the MSScripting library to parse the calculations." & vbCrLf & _
       "Programmed by: Ricardo de Roode." & vbCrLf & _
       vbCrLf & _
       "How to use the external-program-mode:" & vbCrLf & _
       "        - Press ""Ctrl+Shift+9"" to start logging your calculation." & vbCrLf & _
       "        - Press ""Ctrl+Shift+0"" to calculate and paste." & vbCrLf & _
       "        - Press ""Escape"" to cancel the calculation." & vbCrLf & _
       vbCrLf & _
       "When you made a typo during the calculation you can press ""BackSpace"".", vbInformation, "About"
        
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

Text1.Text = """x: "" & iif(d < 0 , ""-"", (-b + §(iif(d<0,0,d))) / (2 * a)) & "" OF x:"" & iif(d < 0 , ""-"", (-b - §(iif(d<0,0,d))) / (2 * a))"
Text3.Text = "const a = " & a & ":const b = " & b & ": const c = " & c & ": dim d: d = b^2 - 4*a*c"
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


Private Sub mnuFileHighDPI_Click()
    mnuFileHighDPI.Checked = Not mnuFileHighDPI.Checked
    
    SaveSetting "Calculator", "Settings", "high dpi", mnuFileHighDPI.Checked
    
    ApplyDPI
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim wordNr As Long
    
    If KeyCode = vbKeySpace And Shift = 2 Then
        KeyCode = 0
        Shift = 0
        
        wordNr = RefillAutocomplete
        
        If lstComplete.ListCount = 1 Then
            Text1.ReplaceWord lstComplete.List(0), wordNr
        End If
        
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        If lstComplete.Visible Then
            If KeyCode = vbKeyUp Then
                If lstComplete.ListIndex > 0 Then lstComplete.ListIndex = lstComplete.ListIndex - 1
            Else
                If lstComplete.ListIndex < lstComplete.ListCount - 1 Then lstComplete.ListIndex = lstComplete.ListIndex + 1
            End If
            
            KeyCode = 0
            Shift = 0
            
        End If
    ElseIf KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        If lstComplete.Visible Then
            KeyCode = 0
            Shift = 0
            
            wordNr = Text1.getWordFromChar(Text1.m_CursorPos)
            
            Text1.ReplaceWord lstComplete.List(lstComplete.ListIndex), wordNr
        End If
        
    ElseIf KeyCode = vbKeyEscape And lstComplete.Visible Then
        KeyCode = 0
        Shift = 0
        
        lstComplete.Visible = False

    End If
End Sub

Function RefillAutocomplete() As Long
    Dim wordNr As Long
    Dim wordS As String
    Dim totalText As String
    Dim i As Long
    Dim foundCount As Long
    Dim foundLast As Long
    
    totalText = Text1.Text
    lstComplete.Clear
    
    wordNr = Text1.getWordFromChar(Text1.m_CursorPos)
    If wordNr = -1 Then
        lstComplete.Visible = False
        Exit Function
    End If
    
    wordS = Mid$(totalText, Text1.getWordStart(wordNr) + 1, Text1.getWordLength(wordNr))
    
    
    For i = 0 To UBound(ExternalCustomFunctions)
        If ExternalCustomFunctions(i) <> "" Then
        
            If InStr(1, ExternalCustomFunctions(i), wordS) = 1 And ExternalCustomFunctions(i) <> wordS Then
                lstComplete.AddItem (ExternalCustomFunctions(i))
            End If
            
        End If
        
    Next i
    
    
    If lstComplete.ListCount = 0 Then
        lstComplete.Visible = False
    Else

        lstComplete.ListIndex = 0
        lstComplete.Visible = True
        Dim cPos As RECT
        GetGlobalCaretPos cPos, False
        lstComplete.Left = cPos.Left + Text1.Left
        lstComplete.Top = Text1.Top + Text1.Height - 1
        lstComplete.Height = IIf(lstComplete.ListCount > 6, 6 * 30, lstComplete.ListCount * 30)
        lstComplete.ItemsVisible = IIf(lstComplete.ListCount > 6, 6, lstComplete.ListCount)

    End If
    
    RefillAutocomplete = wordNr
End Function

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
    Dim fColors(0 To 6) As Long

    fColors(0) = RGB(60, 140, 255)
    fColors(1) = RGB(255, 126, 0)
    fColors(2) = RGB(55, 170, 0)
    fColors(3) = RGB(191, 112, 0)
    fColors(4) = RGB(170, 98, 255)
    fColors(5) = RGB(0, 200, 242)
    fColors(6) = RGB(100, 100, 100)
    
    Dim i As Long
    Dim j As Long
    Dim K As Long
    
    Dim OT As Long 'opened tags
    Dim TT As Long 'total tags
    Dim s As String
    Dim t As String
    Dim lStart As Long
    Dim lend As Long
    Dim lStep As Long
    
    Dim BT As Long 'bold tag
    Dim CFT As Boolean 'colored first tag
    Dim CS As Boolean 'color string
    Dim CD As Boolean 'color dim
    
    s = LCase(txt.Text)
    txt.RedrawPause
    txt.ReCalculateWords
    
    For i = 1 To Len(s)
        t = Mid$(s, i, 1)
        
        txt.setCharBold i - 1, False
        txt.setCharBackColor i - 1, -1
        txt.setCharForeColor i - 1, IIf(CS, fColors(3), IIf(CD, fColors(4), -1))
        
        If t = Chr(34) Then
            CS = Not CS
            txt.setCharForeColor i - 1, fColors(3)
        End If
        
        
        If (t = "[" Or t = "]") And Not CS Then
            CD = Not CD
            txt.setCharForeColor i - 1, fColors(4)
        End If
        
        If Not CS And Not CD Then
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
    
    Dim instrstart As Long
    Dim External() As String
    Dim ExternalColor As Long
    Dim maypaint As Boolean
    
    
    
    For K = 0 To 3
        If K = 0 Then
            External = ExternalFunctions
            ExternalColor = 2
        ElseIf K = 1 Then
            External = ExternalConstants
            ExternalColor = 5
        ElseIf K = 2 Then
            External = ExternalOperators
            ExternalColor = 6
        ElseIf K = 3 Then
            External = ExternalCustomFunctions
            ExternalColor = 2
        End If
        
    
        For i = 0 To UBound(External)
            If External(i) <> "" Then
                instrstart = InStr(instrstart + 1, s, External(i))
                
                While (instrstart > 0)
                    maypaint = True
                    
                    If instrstart - 2 >= 0 Then
                        If txt.getWordFromChar(instrstart - 2) = txt.getWordFromChar(instrstart - 1) Then
                            maypaint = False
                        End If
                    End If
                    
                    If instrstart + Len(External(i)) - 1 < Len(s) Then
                        If txt.getWordFromChar(instrstart - 1) = txt.getWordFromChar(instrstart + Len(External(i)) - 1) Then
                            maypaint = False
                        End If
                    End If
                    
                    
                    
                    If maypaint And txt.getCharForeColor(instrstart - 1) = -1 Then
                        For j = 0 To Len(External(i)) - 1
                            txt.setCharForeColor instrstart - 1 + j, fColors(ExternalColor)
                        Next j
                    End If
                            
                    instrstart = InStr(instrstart + 1, s, External(i))
                Wend
                instrstart = 0
            End If
        Next i
    Next K
    
    txt.RedrawResume
End Sub


Private Sub Text1_Changed()

    formatTextBox Text1
    If lstComplete.Visible = True Then
        RefillAutocomplete
    End If
    
End Sub

Private Sub Text1_SelectionChanged()
Text1_Changed
lstComplete.Visible = False
End Sub

Private Sub tmrFly_Timer()
    Dim cur As RECT

    If MayLog Then
        GetGlobalCaretPos cur
        txtFly.Left = cur.Left
        txtFly.Top = cur.Top
        
        
        If deactivateLogAndSend Then
            If ControlDown = False And ShiftDown = False Then
                deactivateLogAndSend = False
                MayLog = False
                If Len(TypedText) > 0 Then
                    With Form1
                        .Text1.Text = TypedText
                        .Text2.Text = .CheckCalculation(TypedText)
                        
                        If InStr(1, LCase(.Text2.Text), "error") > 0 Then
                            Sendkeys "ERROR"
                        Else
        
                            'Sendkeys ("{backspace " & Len(TypedText) & "}")
                            Sendkeys .Text2.Text
                            
                            TypedText = ""
                        End If
                        
                    End With
                End If
                
            End If
        End If
    Else
        Form1.tmrFly.Enabled = False
        Form1.txtFly.Visible = False
    End If
    
End Sub

Sub txtFly_Changed()
    formatTextBox txtFly
    tmrFly_Timer
End Sub



Private Sub GetGlobalCaretPos(ByRef lPos As RECT, Optional RealPosition As Boolean = True)
    ' get the caret's position
    Dim GUIInfo As GUITHREADINFO
    Dim threadidhwnd As Long
    Dim lres As Long
    Dim crect As RECT
    Dim wind As Long
    
    GUIInfo.cbSize = Len(GUIInfo)
    wind = GetForegroundWindow
    
    lres = GetWindowThreadProcessId(wind, threadidhwnd)
    
    GetGUIThreadInfo lres, GUIInfo
    If RealPosition Then
        GetWindowRect GUIInfo.hwndCaret, crect
    End If
    
    crect.Top = crect.Top + GUIInfo.rcCaret.Top
    crect.Left = crect.Left + GUIInfo.rcCaret.Left
    
    lPos = crect
End Sub





