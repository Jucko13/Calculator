VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0024211E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator by Ricardo"
   ClientHeight    =   4965
   ClientLeft      =   4065
   ClientTop       =   3345
   ClientWidth     =   9735
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
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   649
   WhatsThisHelp   =   -1  'True
   Begin Project1.uListBox lstComplete 
      Height          =   2385
      Left            =   4485
      TabIndex        =   35
      Top             =   2055
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   4207
      BackgroundColor =   3551534
      BorderColor     =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      Text            =   ""
      SelectionBackgroundColor=   3551534
      SelectionBorderColor=   16777215
      SelectionForeColor=   8500547
      ItemHeight      =   31
   End
   Begin VB.Timer tmrFly 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3420
      Top             =   2850
   End
   Begin Project1.uTextBox txtFly 
      Height          =   330
      Left            =   3315
      TabIndex        =   34
      Top             =   4065
      Visible         =   0   'False
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   582
      BackgroundColor =   3551534
      BorderColor     =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ConsoleColors   =   0   'False
      HideCursor      =   -1  'True
      AutoResize      =   -1  'True
   End
   Begin MSComDlg.CommonDialog comm1 
      Left            =   3375
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.uButton cmdClearList 
      Height          =   330
      Left            =   3240
      TabIndex        =   4
      Top             =   4545
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   14322034
      ForeColor       =   14322034
      MouseOverBackgroundColor=   5913650
      CaptionBorderColor=   14737632
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "Clear"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uListBox List1 
      Height          =   2445
      Left            =   3240
      TabIndex        =   3
      Top             =   2025
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   4313
      BackgroundColor =   3551534
      BorderColor     =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      Text            =   "uFrame"
      SelectionBackgroundColor=   3551534
      SelectionBorderColor=   16777215
      SelectionForeColor=   8500547
      ItemHeight      =   40
      VisibleItems    =   4
   End
   Begin Project1.uTextBox Text1 
      Height          =   465
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   820
      BackgroundColor =   3551534
      BorderColor     =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      MousePointer    =   3
      ConsoleColors   =   0   'False
   End
   Begin Project1.uTextBox Text2 
      Height          =   735
      Left            =   90
      TabIndex        =   1
      Top             =   630
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   1296
      BackgroundColor =   3551534
      BorderColor     =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      MousePointer    =   3
      ConsoleColors   =   0   'False
      RowNumberOnEveryLine=   -1  'True
      WordWrap        =   -1  'True
      MultiLine       =   -1  'True
   End
   Begin Project1.uTextBox Text3 
      Height          =   510
      Left            =   90
      TabIndex        =   2
      Top             =   1440
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   900
      BackgroundColor =   3551534
      BorderColor     =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      MousePointer    =   3
      ConsoleColors   =   0   'False
   End
   Begin Project1.uButton cmdNumbers 
      Height          =   330
      Index           =   20
      Left            =   90
      TabIndex        =   5
      Top             =   2025
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   14322034
      ForeColor       =   14322034
      MouseOverBackgroundColor=   5913650
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "<-"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   720
      TabIndex        =   6
      Top             =   2025
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   14322034
      ForeColor       =   14322034
      MouseOverBackgroundColor=   5913650
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "Clear"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   2610
      TabIndex        =   7
      Top             =   2430
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   1746682
      ForeColor       =   1746682
      MouseOverBackgroundColor=   1584197
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "b/c"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   1980
      TabIndex        =   8
      Top             =   2430
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   11944815
      ForeColor       =   11944815
      MouseOverBackgroundColor=   6894151
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "+"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   90
      TabIndex        =   9
      Top             =   2430
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   4671472
      ForeColor       =   4671472
      MouseOverBackgroundColor=   2434394
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "7"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   90
      TabIndex        =   10
      Top             =   3240
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   4671472
      ForeColor       =   4671472
      MouseOverBackgroundColor=   2434394
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "1"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   720
      TabIndex        =   11
      Top             =   3240
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   4671472
      ForeColor       =   4671472
      MouseOverBackgroundColor=   2434394
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "2"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   1350
      TabIndex        =   12
      Top             =   3240
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   4671472
      ForeColor       =   4671472
      MouseOverBackgroundColor=   2434394
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "3"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   90
      TabIndex        =   13
      Top             =   2835
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   4671472
      ForeColor       =   4671472
      MouseOverBackgroundColor=   2434394
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "4"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   720
      TabIndex        =   14
      Top             =   2835
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   4671472
      ForeColor       =   4671472
      MouseOverBackgroundColor=   2434394
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "5"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   1350
      TabIndex        =   15
      Top             =   2835
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   4671472
      ForeColor       =   4671472
      MouseOverBackgroundColor=   2434394
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "6"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   1350
      TabIndex        =   16
      Top             =   2430
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   4671472
      ForeColor       =   4671472
      MouseOverBackgroundColor=   2434394
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "9"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Index           =   11
      Left            =   1980
      TabIndex        =   17
      Top             =   4140
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   1746682
      ForeColor       =   1746682
      MouseOverBackgroundColor=   1584197
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "sqr"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   1980
      TabIndex        =   18
      Top             =   2835
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   11944815
      ForeColor       =   11944815
      MouseOverBackgroundColor=   6894151
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "/"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   1980
      TabIndex        =   19
      Top             =   3240
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   11944815
      ForeColor       =   11944815
      MouseOverBackgroundColor=   6894151
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "*"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   1980
      TabIndex        =   20
      Top             =   3645
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   11944815
      ForeColor       =   11944815
      MouseOverBackgroundColor=   6894151
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "-"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   1350
      TabIndex        =   21
      Top             =   3645
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   4671472
      ForeColor       =   4671472
      MouseOverBackgroundColor=   2434394
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "."
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   90
      TabIndex        =   32
      Top             =   3645
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   4671472
      ForeColor       =   4671472
      MouseOverBackgroundColor=   2434394
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "0"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   2610
      TabIndex        =   33
      Top             =   3240
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   1323
      BackgroundColor =   2367774
      BorderColor     =   8500547
      ForeColor       =   8500547
      MouseOverBackgroundColor=   3425832
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "="
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   90
      TabIndex        =   22
      Top             =   4140
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   1746682
      ForeColor       =   1746682
      MouseOverBackgroundColor=   1584197
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "Tan"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   90
      TabIndex        =   23
      Top             =   4545
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   1746682
      ForeColor       =   1746682
      MouseOverBackgroundColor=   1584197
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "aTn"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   720
      TabIndex        =   24
      Top             =   4140
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   1746682
      ForeColor       =   1746682
      MouseOverBackgroundColor=   1584197
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "Sin"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   720
      TabIndex        =   25
      Top             =   4545
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   1746682
      ForeColor       =   1746682
      MouseOverBackgroundColor=   1584197
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "aSn"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   1350
      TabIndex        =   26
      Top             =   4140
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   1746682
      ForeColor       =   1746682
      MouseOverBackgroundColor=   1584197
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "Cos"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   1350
      TabIndex        =   27
      Top             =   4545
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   1746682
      ForeColor       =   1746682
      MouseOverBackgroundColor=   1584197
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "aCs"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   2610
      TabIndex        =   28
      Top             =   2835
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   1746682
      ForeColor       =   1746682
      MouseOverBackgroundColor=   1584197
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "^"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionOffsetTop=   4
   End
   Begin Project1.uButton cmdExtras 
      Height          =   330
      Index           =   7
      Left            =   2610
      TabIndex        =   29
      Top             =   4140
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   1746682
      ForeColor       =   1746682
      MouseOverBackgroundColor=   1584197
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "PI"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   1980
      TabIndex        =   30
      Top             =   4545
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   1746682
      ForeColor       =   1746682
      MouseOverBackgroundColor=   1584197
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "("
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   2610
      TabIndex        =   31
      Top             =   4545
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   1746682
      ForeColor       =   1746682
      MouseOverBackgroundColor=   1584197
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   ")"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   720
      TabIndex        =   36
      Top             =   2430
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   4671472
      ForeColor       =   4671472
      MouseOverBackgroundColor=   2434394
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "8"
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Left            =   1980
      TabIndex        =   37
      Top             =   2025
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      BackgroundColor =   2367774
      BorderColor     =   1746682
      ForeColor       =   1746682
      MouseOverBackgroundColor=   1584197
      FocusColor      =   0
      BackgroundColorDisabled=   0
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   " & "" "" & "
      BorderAnimation =   0
      AlignPictureInCorner=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
      Begin VB.Menu mnuFileOpenFuncties 
         Caption         =   "Open Functies"
      End
      Begin VB.Menu mnuFileReloadFunctions 
         Caption         =   "Herlaad functies"
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


Private Sub cmdClearList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If (x < 0) Or (y < 0) Or (x > cmdClearList.Width) Or (y > cmdClearList.Height) Then
    '    ReleaseCapture
    '    Set cmdClearList.Picture = picNormal4.Picture

        
    'ElseIf GetCapture() <> cmdClearList.hWnd Then
    '    SetCapture cmdClearList.hWnd
    '    Set cmdClearList.Picture = PicHigh4.Picture
    'End If
    'List1.Redraw = True
End Sub

Private Sub cmdClearList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub cmdExtras_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim selection As Long

Select Case Index
    Case 0 To 6, 11
        If Text1.SelLength > 0 Then
            Dim tmpStr As String
            selection = Text1.SelStart
            tmpStr = Text1.GetSelectionText
            Text1.AddCharAtCursor cmdExtras(Index).Caption & "(" & tmpStr & ")"
            Text1.SelStart = selection + Len(cmdExtras(Index).Caption) + Len(tmpStr) + 1
        Else
            selection = Text1.SelStart
            Text1.AddCharAtCursor cmdExtras(Index).Caption & "()"
            Text1.SelStart = selection + Len(cmdExtras(Index).Caption) + 1
        End If
        
    Case 7 To 8
        Text1.AddCharAtCursor cmdExtras(Index).Caption
        
End Select

Text1.SetFocus


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
        If Text1.TextLength > 0 Then
            If Text1.SelLength > 0 Then
                Text1.AddCharAtCursor ""
                Text1.SetFocus
            
            ElseIf Text1.SelStart > 0 Then
                Text1.SelStart = Text1.SelStart - 1
                Text1.SelLength = 1
                Text1.AddCharAtCursor ""
                Text1.SetFocus
            Else
                Text1.SelStart = 0
                Text1.SelLength = 1
                Text1.AddCharAtCursor ""
                Text1.SetFocus
                'Text1.SelStart = 0
            End If
            'Text1.Text = Mid(Text1.Text, 1, Text1.TextLength - 1)
        Else
            'Text1.SetFocus
        End If
        
        
    Case 21
        'If InStr(1, Text2.Text, "/") > 0 Then
        '    If tmpVal <> "" Then
        '        Text2.Text = tmpVal
        '    End If
        'Else
        If Val(Text2.Text) Then
            tmpTx = GetFraction(Text1.Text)
            'If tmpTx = Text2.Text Then
                'tmpTx = Dec2Frac(Text2.Text)
            'End If
            Text2.Text = tmpTx
        End If
        'End If
        
        
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
        Else
            t = Text1.GetMidText(tmpLines(i), "sub ", "(")
            If t <> "" Then
                ReDim Preserve ExternalCustomFunctions(0 To c) As String
                ExternalCustomFunctions(c) = t
                c = c + 1
            End If
        End If
    Next i
    
    MergeSort ExternalCustomFunctions
    
    
    ExternalConstants = Split("pi e integer string double float long byte vbabortretryignore vbapplicationmodal vbarray vbblack vbblue vbboolean vbbyte vbcr vbcritical vbcrlf vbcurrency vbcyan vbdataobject vbdate vbdecimal vbdefaultbutton1 vbdefaultbutton2 vbdefaultbutton3 vbdefaultbutton4 vbdouble vbempty vberror vbexclamation vbfalse vbformfeed vbgreen vbinformation vbinteger vblf vblong vbmagenta vbnewline vbnull vbnullchar vbnullstring vbobject vbokcancel vbokonly vbquestion vbred vbretrycancel vbsingle vbstring vbsystemmodal vbtab vbtrue vbusedefault vbvariant vbverticaltab vbwhite vbyellow vbyesno vbyesnocancel vbbinarycompare vbtextcompare", " ")
    
    
    ExternalOperators = Split("xor and or not is * - + / ^ : false true", " ")
    
    Exit Sub
Err:
    MsgBox Err.Description, vbOKOnly Or vbCritical, "ERROR: " & Err.Number
    
End Sub

Sub drawMenu()
    Dim mi As MENUINFO
    Dim lbBrushInfo As LOGBRUSH
    Dim ret As Long
    
    Dim menuColor As Long
    Dim lRGBColor As Long
    
    menuColor = &H36312E
    OleTranslateColor menuColor, 0, lRGBColor
    
    mi.cbSize = Len(mi)
    
    ret = GetMenuInfo(GetMenu(Me.hWnd), mi) ' 0 means failure
    
    
    With mi
        .dwMenuData = 900
        .fMask = MIM_BACKGROUND Or MIM_STYLE Or MIM_APPLYTOSUBMENUS
        
        .dwStyle = MNS_NOCHECK Or MNS_NOTIFYBYPOS
        
        lbBrushInfo.lbStyle = 0
        lbBrushInfo.lbColor = RGB(155, 100, 200)
        lbBrushInfo.lbHatch = 0
        
        
        .hbrBack = CreateBrushIndirect(lbBrushInfo)
        
        
        SetMenuInfo GetMenu(Me.hWnd), mi  'main menu bar
        
    End With
    
    'DrawMenuBar Me.hWnd
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


Function GetFraction(Calculation As String) As String
    Dim dOne As Double
    
     Dim tend As Double
    Dim tim As clsTimer
    Set tim = New clsTimer
    
    tim.tStart
    
    dOne = Val(CheckCalculation("1 / (" & Calculation & ")"))
    
    Dim bigNumber As String
    
    bigNumber = dOne
    
    If InStr(1, LCase(bigNumber), "e") > 0 Then
        GetFraction = "??"
        Exit Function
    End If
    
    Dim lPlace As Long
    
    lPlace = InStr(1, bigNumber, ",")
    Dim toPowerOf As Long
    
    toPowerOf = Len(bigNumber) - lPlace
    
    If toPowerOf <= 0 Then
        GetFraction = "??"
        Exit Function
    End If
    
    Dim upperBound As Double
    Dim lowerBound As Double
    
    upperBound = 10 ^ toPowerOf
    lowerBound = dOne * upperBound
    
    If upperBound >= 100000000 Then
        GetFraction = GetFractionSlow(dOne, "1 / (" & Calculation & ")")
    Else
        Dim t As Double
        Dim a As Double
        Dim b As Double
        
        a = upperBound
        b = lowerBound
        While b <> 0
            t = b
            b = FMod(a, b)
            a = t
        Wend
        
        GetFraction = upperBound / a & " / " & lowerBound / a
        
    End If
    
    tend = tim.tStop
    mnuExecTime.Caption = "ExecTime: " & mircoToTime(tend)
End Function

Function GetFractionSlow(startNum As Double, Calculation As String) As String
    Dim i As Long
    
    'If startNum < 2 Then
    '    GetFractionSlow = "??"
    '    Exit Function
    'End If
    
    
    For i = 2 To 2000
        Dim tmp As String
        
        'Debug.Assert i <> 86
        
        tmp = CheckCalculation("(" & Calculation & ")" & " * " & i)
        If tmp = Fix(Val(tmp)) Then
            Dim upperBound As Double
            Dim lowerBound As Double
            
            upperBound = i
            lowerBound = CheckCalculation("(" & Calculation & ")" & " * " & upperBound)
    
            GetFractionSlow = upperBound & " / " & lowerBound
            Exit Function
        End If
        
    Next i
End Function


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
    objScript.AddObject "Me", Me, True
    objScript.AllowUI = True
    
    'allfunctions = CharExecution(objScript.CodeObject, False)
    'allfunctions = allfunctions & CharExecution(objScript.CodeObject.winapi, True)
    
    'MsgBox allfunctions
    ' MsgBox
    result = objScript.Eval(CalculateString)
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
If RowMouse = -1 Then Exit Sub

Text1.Text = List1.Cell(RowMouse, 0)
Text2.Text = Replace(List1.Cell(RowMouse, 1), vbCrLf, "")
End Sub

Private Sub List1_ItemAdded(ItemIndex As Long)
    List1.RedrawPause
    Dim i As Long
    
    For i = 0 To List1.ListCount - 1
        If i Mod 2 = 0 Then
            List1.ItemBackColor(i) = &H3F3936
        Else
            List1.ItemBackColor(i) = -1
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

Private Sub mnuFileOpenFuncties_Click()
    ShellExecute Me.hWnd, "open", App.Path & "/functionlist.txt", "", "", vbNormalFocus
End Sub

Private Sub mnuFileReloadFunctions_Click()
    initializeScript
    Text1_Changed
    Text2_Changed
    Text3_Changed
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

Private Sub Text2_SelectionChanged()
    Dim tmpStr As String
    
    tmpStr = Replace(Text2.GetSelectionText, " ", "")
    
    If Text2.BackgroundColor <> 3551534 Then Text2.BackgroundColor = 3551534
    
    If tmpStr Like "rgb(*,*,*)" Then
        Dim tmpSplit() As String
        tmpSplit = Split(Mid$(tmpStr, 5, Len(tmpStr) - 5), ",")
        If IsNumeric(tmpSplit(0)) And IsNumeric(tmpSplit(1)) And IsNumeric(tmpSplit(2)) Then
            Text2.BackgroundColor = RGB(tmpSplit(0), tmpSplit(1), tmpSplit(2))
        End If
    End If
    
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

    fColors(0) = RGB(60, 140, 255) 'light blue
    fColors(1) = RGB(255, 126, 0) 'bright orange
    fColors(2) = &H81B543 'very soft green/cyan    'RGB(55, 170, 0)
    fColors(3) = RGB(191, 112, 0) 'dark orange
    fColors(4) = RGB(170, 98, 255) 'soft light purple
    fColors(5) = RGB(0, 200, 242) 'bright light blue/cyan
    fColors(6) = &HDA8972 'light/soft blue/purple
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim OT As Long 'opened tags
    Dim TT As Long 'total tags
    Dim s As String 'string
    Dim sl As Long 'string length
    Dim t As String 'text
    Dim tn As String 't next
    Dim lStart As Long
    Dim lend As Long
    Dim lStep As Long
    
    Dim BT As Long 'bold tag
    Dim CFT As Boolean 'colored first tag
    Dim CS As Boolean 'color string
    Dim CD As Boolean 'color dim
    Dim CH As Boolean 'color hex
    Dim cc As Boolean 'color comment
    
    s = LCase(txt.Text)
    sl = Len(s)
    
    txt.RedrawPause
    txt.ReCalculateWords
    
    For i = 1 To sl
        t = Mid$(s, i, 1)
        If i + 1 < sl Then tn = Mid$(s, i + 1, 1)
        
        txt.setCharBold i - 1, True
        txt.setCharBackColor i - 1, -1
        txt.setCharForeColor i - 1, IIf(CS, fColors(1), IIf(CD, fColors(4), IIf(CH, fColors(4), IIf(cc, fColors(2), -1))))
        
        
        If t = "'" And Not cc And Not CS Then
            cc = Not cc
            txt.setCharForeColor i - 1, fColors(2)
        End If
        
        
        If t = Chr(34) And Not CH And Not CD Then
            CS = Not CS
            txt.setCharForeColor i - 1, fColors(1)
        End If
        
        
        If (t = "[" Or t = "]") And Not CS And Not CH Then
            CD = Not CD
            txt.setCharForeColor i - 1, fColors(4)
        End If
        
        If ((t = "&" And tn = "h" And Not CH) Or (t = "&" And CH)) And Not CS And Not CD Then
            CH = Not CH
            txt.setCharForeColor i - 1, fColors(4)
        End If
        
        
        If Not CS And Not CD And Not CH Then
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
                    txt.setCharForeColor i - 1, &H4747F0
                
                Case "="
                    txt.setCharForeColor i - 1, &H81B543
                    
                Case "/", "-", "+", "*", "^"
                    txt.setCharForeColor i - 1, &HFF62AA
                    
                Case Else
            End Select
        End If
        
    Next i
    
    Dim instrstart As Long
    Dim External() As String
    Dim ExternalColor As Long
    Dim maypaint As Boolean
    
    
    
    For k = 0 To 3
        If k = 0 Then
            External = ExternalFunctions
            ExternalColor = 2
        ElseIf k = 1 Then
            External = ExternalConstants
            ExternalColor = 5
        ElseIf k = 2 Then
            External = ExternalOperators
            ExternalColor = 6
        ElseIf k = 3 Then
            External = ExternalCustomFunctions
            ExternalColor = 2
        End If
        
    
        For i = 0 To UBound(External)
            If External(i) <> "" Then
                instrstart = InStr(instrstart + 1, s, External(i))
                
                While (instrstart > 0)
                    maypaint = True
                    
                    If instrstart - 3 >= 0 Then
                        'Debug.Print txt.getWordFromChar(instrstart - 3)
                        If txt.getWordFromChar(instrstart - 2) = txt.getWordFromChar(instrstart - 3) Then
                            maypaint = False
                        End If
                    ElseIf instrstart > 2 Then
                        maypaint = False
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
    Next k
    
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





