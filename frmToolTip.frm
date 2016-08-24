VERSION 5.00
Begin VB.Form frmToolTip 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   1950
   ClientLeft      =   -8520
   ClientTop       =   -900
   ClientWidth     =   5115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmToolTip.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   341
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   360
      Left            =   2865
      TabIndex        =   2
      Top             =   885
      Width           =   525
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   1200
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   0
      Top             =   765
      Visible         =   0   'False
      Width           =   1350
      Begin VB.Line Line1 
         BorderColor     =   &H0000FF00&
         X1              =   44
         X2              =   52
         Y1              =   1
         Y2              =   1
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000FF00&
         X1              =   43
         X2              =   43
         Y1              =   2
         Y2              =   10
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000FF00&
         X1              =   44
         X2              =   52
         Y1              =   10
         Y2              =   10
      End
      Begin VB.Line Line4 
         BorderColor     =   &H0000FF00&
         X1              =   52
         X2              =   52
         Y1              =   2
         Y2              =   10
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H0000FF00&
         Height          =   210
         Left            =   45
         TabIndex        =   1
         Top             =   0
         Width           =   630
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3180
      Top             =   15
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1050
      Top             =   45
   End
End
Attribute VB_Name = "frmToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X_Pos As Long
    Y_Pos As Long
End Type

Function GetXas() As Long
   Dim pt As POINTAPI
   GetCursorPos pt
   GetXas = pt.X_Pos
End Function

Function GetYas() As Long
   Dim pt As POINTAPI
   GetCursorPos pt
   GetYas = pt.Y_Pos
End Function

Private Sub Command1_Click()
Picture1_Resize
End Sub

Private Sub Form_Load()
'SetWindowLong Picture1.hwnd, GWL_STYLE, GetWindowLong(Picture1.hwnd, GWL_STYLE) + WS_BORDER

    Call SetWindowLong(Picture1.hWnd, GWL_EXSTYLE, &H80)
    Call SetParent(Picture1.hWnd, GetParent(Form1.hWnd))
    
    SetTopmostWindow Picture1.hWnd, True
    frmToolTip.Visible = False
    frmToolTip.Hide


'SetTopmostWindow Me.hwnd, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If ShowCORDSTool = True Then stopMouseHook2
End Sub

Sub SetFormShit(mode As Boolean)
If mode = True Then
    startMouseHook2
    
ElseIf mode = False Then
    stopMouseHook2
    
End If
End Sub

Sub Picture1_Resize()
    Dim R1 As Rect
    Dim R2 As Rect
    Dim i As Integer
    
    'Picture1.Cls
    
    R1.Top = 0
    R1.Left = 0
    R1.Right = Picture1.Width
    R1.Bottom = ((Picture1.Height) / 2) - 3
    
    R2.Top = ((Picture1.Height) / 2) + 3
    R2.Left = 0
    R2.Right = Picture1.Width
    R2.Bottom = Picture1.Height

    'FillGradient Picture1.hDC, R1, Form1.Jucko13.BackColor, RGB(255, 255, 255), True, True
    'FillGradient Picture1.hDC, R2, RGB(255, 255, 255), Form1.Jucko13.BackColor, True, True
    
    'Picture1.AutoRedraw = True
    'Picture1.Refresh
    
Line1.X1 = 0
Line1.X2 = Picture1.ScaleWidth
Line1.Y1 = 0
Line1.Y2 = 0

Line2.X1 = 0
Line2.X2 = 0
Line2.Y1 = 0
Line2.Y2 = Picture1.ScaleHeight

Line3.X1 = 0
Line3.X2 = Picture1.ScaleWidth
Line3.Y1 = Picture1.ScaleHeight - 1
Line3.Y2 = Picture1.ScaleHeight - 1

Line4.X1 = Picture1.ScaleWidth - 1
Line4.X2 = Picture1.ScaleWidth - 1
Line4.Y1 = 0
Line4.Y2 = Picture1.ScaleHeight
End Sub

Private Sub Timer1_Timer()
Dim Xi As Integer
Dim Yi As Integer

Xi = (GetXas * 15)
Yi = (GetYas * 15)

If Picture1.Visible = True Then
    If (Xi + Picture1.Width + 215) > Screen.Width Then
        Picture1.Left = Screen.Width - Picture1.Width
        If (Yi + Picture1.Height + 215) > Screen.Height Then
            Picture1.Top = Screen.Height - Picture1.Height
        Else
            Picture1.Top = ((Yi / 15) + 15) * 15
        End If
    ElseIf (Yi + Picture1.Height + 215) > Screen.Height Then
            Picture1.Top = Screen.Height - Picture1.Height
            If (Xi + Picture1.Width + 215) > Screen.Width Then
                Picture1.Left = Screen.Width - Picture1.Width
            Else
                Picture1.Left = ((Xi / 15) + 15) * 15
            End If
    Else
        Picture1.Top = ((Yi / 15) + 15) * 15
        Picture1.Left = ((Xi / 15) + 15) * 15
    End If
    If ShowCORDSTool = True Then
        lblText.Caption = "Click to confirm this position: " & (Xi / 15) & ", " & (Yi / 15)
        Picture1.Width = TextWidth(lblText.Caption) * 15 + 160
    End If
End If
End Sub


Private Sub Timer2_Timer()
Dim i As Integer
Timer2.Enabled = False
ShowIngToolTip = False

SetTrans Me.Picture1.hWnd, 220
Me.Picture1.Visible = True
For i = 240 To 0 Step -20
    SetTrans Me.Picture1.hWnd, i
    Wait 0
Next i
Picture1.Visible = False
End Sub
