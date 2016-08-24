VERSION 5.00
Begin VB.UserControl TransPicBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2040
   ControlContainer=   -1  'True
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   4575
   ScaleWidth      =   2040
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   195
      ScaleHeight     =   855
      ScaleWidth      =   1020
      TabIndex        =   0
      Top             =   2190
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hallo"
      BeginProperty Font 
         Name            =   "Jucko13"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   615
   End
End
Attribute VB_Name = "TransPicBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''global HowMany
'' if Replaced = true then Howmany = Howmany + 1

Option Explicit

Const m_def_AutoSize = False

Dim WasDown As Boolean

Public Event Click()
Public Event DblClick()
Public Event Resize()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Enum pbBackStyle
    [Opaque] = 1
    [Transparent] = 0
End Enum

Enum pbBorderstyle
    [No Border] = 0
    [Fixed Single] = 1
End Enum

Enum pbAppearance
    [3D] = 1
    [Flat] = 0
End Enum

Dim m_AutoSize As Boolean
Dim mCaptionColor As OLE_COLOR

Public Property Get Appearance() As pbAppearance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As pbAppearance)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

Public Property Get BackStyle() As pbBackStyle
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As pbBackStyle)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As pbBorderstyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As pbBorderstyle)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    Set UserControl.MaskPicture = New_Picture
    Set Picture1.Picture = New_Picture
    If m_AutoSize Then
        UserControl.Width = Picture1.Width
        UserControl.Height = Picture1.Height
    End If
      
    PropertyChanged "Picture"
End Property

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
WasDown = True
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

RaiseEvent MouseUp(Button, Shift, X, Y)
If WasDown = True Then
    WasDown = False
    RaiseEvent Click
End If
End Sub

Sub UserControl_Resize()
    If m_AutoSize Then
        UserControl.Width = Picture1.Width
        UserControl.Height = Picture1.Height
    End If
    Label1.Left = 0
    Label1.Width = UserControl.Width
    Label1.Top = (UserControl.Height / 2) - (TextHeight(Label1.Caption) / 2)
    RaiseEvent Resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Font.Size = 10
Font.Bold = False
Font.Name = "Jucko13"
Font.Italic = False

    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", 16777215)
    Set FontName = PropBag.ReadProperty("FontName", Font)
    Label1.Caption = PropBag.ReadProperty("Caption", "")
    ForeColor = PropBag.ReadProperty("ForeColor", mCaptionColor)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, 16777215)
    Call PropBag.WriteProperty("FontName", FontName, "Jucko13")
    Call PropBag.WriteProperty("Caption", Label1.Caption, "")
    Call PropBag.WriteProperty("ForeColor", ForeColor, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    UserControl_Resize
End Sub


Private Sub UserControl_InitProperties()
    m_AutoSize = m_def_AutoSize
    mCaptionColor = 0
End Sub

Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    If m_AutoSize Then
        UserControl.Width = Picture1.Width
        UserControl.Height = Picture1.Height
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
    PropertyChanged "AutoSize"
    UserControl.BackColor = New_Color
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "122c"
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Str As String)
    PropertyChanged "Caption"
    Label1.Caption = New_Str
    UserControl_Resize
End Property

Public Property Get FontName() As Font
    Set FontName = UserControl.Font
End Property

Public Property Set FontName(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Set Label1.Font = New_Font
    PropertyChanged "Font"
    UserControl_Resize
End Property

Property Let ForeColor(NewValue As OLE_COLOR)
mCaptionColor = NewValue
Label1.ForeColor = NewValue
PropertyChanged "ForeColor"
End Property

Property Get ForeColor() As OLE_COLOR
ForeColor = mCaptionColor
End Property
