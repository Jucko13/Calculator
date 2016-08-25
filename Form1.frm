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
   MaxButton       =   0   'False
   ScaleHeight     =   303
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   567
   WhatsThisHelp   =   -1  'True
   Begin MSComDlg.CommonDialog comm1 
      Left            =   3165
      Top             =   1350
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
      ForeColor       =   16711680
      Text            =   "uFrame"
      SelectionBackgroundColor=   16745771
      SelectionBorderColor=   16745771
      SelectionForeColor=   16777215
      ItemHeight      =   33
   End
   Begin Project1.uTextBox Text1 
      Height          =   315
      Left            =   90
      TabIndex        =   8
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
      Picture         =   "Form1.frx":17E2
      ScaleHeight     =   330
      ScaleWidth      =   5040
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   0
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
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   20
      Left            =   105
      TabIndex        =   13
      Top             =   1455
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
      Top             =   1455
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
      Top             =   1455
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
      Top             =   1875
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
      Top             =   2295
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
      Top             =   1875
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
      Top             =   2715
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
      Top             =   2715
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
      Top             =   2715
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
      Top             =   2295
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
      Top             =   2295
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
      Top             =   2295
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
      Top             =   1875
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
      Top             =   1875
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
      Top             =   1875
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
      Top             =   2295
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
      Top             =   2715
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
      Top             =   3135
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
      Top             =   3135
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
      Caption         =   ","
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
      Top             =   3135
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
      Left            =   2625
      TabIndex        =   43
      Top             =   2715
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
    x As Long
    y As Long
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
        
        If List1.ListCount > 0 Then
            If List1.Cell(0, 0) <> Text1.Text Or List1.Cell(0, 1) <> Text2.Text Then
                List1.AddItem Text1.Text & Chr(9) & vbCrLf & Text2.Text, 0
            End If
        Else
            List1.AddItem Text1.Text & Chr(9) & vbCrLf & Text2.Text, 0
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
    objWinApi.initialize comm1
    
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
    'For i = 1 To 9
    '    Set cmdNumbers(i).Picture = picNormal.Picture
    'Next i
    
'
'    Set cmdNumbers(11).Picture = picNormal.Picture
'    Set cmdNumbers(12).Picture = picNormal.Picture
'    Set cmdNumbers(13).Picture = picNormal.Picture
'    Set cmdNumbers(14).Picture = picNormal.Picture
'    Set cmdNumbers(16).Picture = picNormal.Picture
'    Set cmdNumbers(18).Picture = picNormal.Picture
'    Set cmdNumbers(19).Picture = picNormal.Picture
'    Set cmdNumbers(21).Picture = picNormal.Picture
'
'    Set cmdNumbers(0).Picture = picNormal2.Picture
'    Set cmdNumbers(15).Picture = picNormal2.Picture
'    Set cmdNumbers(20).Picture = picNormal2.Picture
'
'    Set cmdNumbers(10).Picture = picNormal3.Picture
'    Set cmdNumbers(17).Picture = picNormal.Picture
'
'    Set cmdClearList.Picture = picNormal4.Picture
    'cmdClearList.BackStyle = Transparent
    
    On Error Resume Next
    
'    For i = 0 To 21
'        cmdNumbers(i).BackStyle = Transparent
'    Next i
'
'    For i = 0 To 8
'        cmdExtras(i).BackStyle = Transparent
'        Set cmdExtras(i).Picture = picNormal.Picture
'    Next i
    'Set cmdNumbers(i).Picture = picNormal.Picture
    
    'List1.AddColumn "Berekening", List1.Width / 2 - 270, lgAlignCenterCenter, lgString, , , , True, , True, False
    'List1.AddColumn "Antwoord", List1.Width / 2 - 270, lgAlignCenterCenter, lgString, , , , True, , True, True
    
    List1.BackgroundColor = &HF1E4D9
    List1.setTabStop 0, 0, vbLeftJustify
    List1.setTabStop 1, List1.Width - 4 - 20, vbRightJustify
    
    'List1.ForeColorHdr = &HFF0000
    List1.Redraw
    'List1.Redraw = True
    Text1.SelStart = 1
    
    TotalRows = GetSetting("Calculator", "Berekeningen", "Rows", -1)
    Text1.Text = GetSetting("Calculator", "Berekeningen", "Text1.text", 0)
    Text2.Text = GetSetting("Calculator", "Berekeningen", "Text2.text", 0)
    Text3.Text = GetSetting("Calculator", "Berekeningen", "Text3.text", "")
    Text1.SelStart = GetSetting("Calculator", "Berekeningen", "Text1.SelStart", 0)
    Text2.SelStart = GetSetting("Calculator", "Berekeningen", "Text2.SelStart", 0)
    
    'cmdExtras_MouseUp 8, 0, 0, 0, 0
    If TotalRows > -1 Then
        For i = 0 To TotalRows - 1
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
    If ItemIndex Mod 2 = 0 Then
        List1.ItemColor(ItemIndex) = RGB(228, 235, 244)
    Else
        List1.ItemColor(ItemIndex) = -1
    End If
    
End Sub

Private Sub List1_MouseEnter()
'List1.GridColor = &HDBFF&
End Sub

Private Sub List1_MouseLeave()
'List1.GridColor = &HFF852B
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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

