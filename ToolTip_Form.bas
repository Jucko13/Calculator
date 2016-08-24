Attribute VB_Name = "ToolTip_Form"
Private Declare Function GetTickCount Lib "kernel32" () As Long

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Global Const AlphaBet As String = " abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ[.-=>]{,_+<}?/\|:;'`~!@#$%^&*()"
Global Const AlphaiS  As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
Global Const Numbers As String = "0123456789, "


Global SYSBackColor
Global SYSForeColor
Global ShowCORDSTool As Boolean
Global DifCords As String

Global Const InBox As Integer = 11
Global Const vbFontMode As Integer = 31
Global Const vbRunMode As Integer = 62

Global ButtonPressed

Enum ShowMode
    [Unload / Load] = 0
    [Load / Edit] = 1
End Enum

Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Global GlobalFormClosing As Boolean

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Dim pos As Integer
Dim LargestLabel As Integer
Global ShowIngToolTip As Boolean

'------------------------------------

'------------------------------------

Sub SetTopmostWindow(ByVal hWnd As Long, Optional topmost As Boolean = True)
    Const HWND_NOTOPMOST = -2
    Const HWND_TOPMOST = -1
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    SetWindowPos hWnd, IIf(topmost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, _
        SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE
End Sub

Sub Wait(ByVal dblMilliseconds As Double)
    Dim dblStart As Double
    Dim dblEnd As Double
    Dim dblTickCount As Double
    
    dblTickCount = GetTickCount()
    dblStart = GetTickCount()
    dblEnd = GetTickCount + dblMilliseconds
    
    Do
    DoEvents
    dblTickCount = GetTickCount()
    Loop Until dblTickCount > dblEnd Or dblTickCount < dblStart
       
End Sub

Sub ShowToolTip(Str As String, Tijd As Integer, Optional BgClr = vbBlack, _
                Optional TxtClr = vbBlue, Optional SkinNumber As String = 0, _
                Optional IsControl As Boolean = False, Optional ShowMD As ShowMode = [Unload / Load], _
                Optional ShowCords As Boolean = False)
                
Dim Multilined As String
Dim entry() As String
Dim i As Integer

        If pos > 0 Then
            For i = 1 To pos - 1
                frmToolTip.Controls.Remove ("boe" & i)
            Next i
        End If

entry = Split(Str, vbCrLf, , vbTextCompare)

If ShowCords = True Then ShowCORDSTool = True Else ShowCORDSTool = False

pos = 0
LargestLabel = 0

        If SkinNumber = 0 Then
            With frmToolTip
                .lblText.Forecolor = Form1.jucko13.Forecolor
                .Line1.BorderColor = Form1.jucko13.Forecolor
                .Line2.BorderColor = Form1.jucko13.Forecolor
                .Line3.BorderColor = Form1.jucko13.Forecolor
                .Line4.BorderColor = Form1.jucko13.Forecolor
                .Picture1.BackColor = Form1.jucko13.BackColor
            End With
        ElseIf SkinNumber = 1 Then
            With frmToolTip
                .lblText.Forecolor = vbGreen
                .Line1.BorderColor = vbGreen
                .Line2.BorderColor = vbGreen
                .Line3.BorderColor = vbGreen
                .Line4.BorderColor = vbGreen
                .Picture1.BackColor = vbBlack
            End With
        ElseIf SkinNumber = 2 Then
            With frmToolTip
                .lblText.Forecolor = vbRed
                .Line1.BorderColor = vbRed
                .Line2.BorderColor = vbRed
                .Line3.BorderColor = vbRed
                .Line4.BorderColor = vbRed
                .Picture1.BackColor = vbBlack
            End With
        ElseIf SkinNumber = 3 Then
            With frmToolTip
                .lblText.Forecolor = &HFF572D
                .Line1.BorderColor = &HFF572D
                .Line2.BorderColor = &HFF572D
                .Line3.BorderColor = &HFF572D
                .Line4.BorderColor = &HFF572D
                .Picture1.BackColor = vbBlack
            End With
        ElseIf SkinNumber = 4 Then
            With frmToolTip
                .lblText.Forecolor = vbGreen
                .Line1.BorderColor = vbGreen
                .Line2.BorderColor = vbGreen
                .Line3.BorderColor = vbGreen
                .Line4.BorderColor = vbGreen
                .Picture1.BackColor = vbBlack
            End With
        ElseIf SkinNumber = 5 Then
            With frmToolTip
                .lblText.Forecolor = vbRed
                .Line1.BorderColor = vbRed
                .Line2.BorderColor = vbRed
                .Line3.BorderColor = vbRed
                .Line4.BorderColor = vbRed
                .Picture1.BackColor = vbBlack
            End With
        ElseIf SkinNumber = 6 Then
            With frmToolTip
                .lblText.Forecolor = &HFF572D
                .Line1.BorderColor = &HFF572D
                .Line2.BorderColor = &HFF572D
                .Line3.BorderColor = &HFF572D
                .Line4.BorderColor = &HFF572D
                .Picture1.BackColor = vbBlack
            End With
        ElseIf SkinNumber = 7 Then
            With frmToolTip
                .lblText.Forecolor = vbGreen
                .Line1.BorderColor = vbGreen
                .Line2.BorderColor = vbGreen
                .Line3.BorderColor = vbGreen
                .Line4.BorderColor = vbGreen
                .Picture1.BackColor = vbBlack
            End With
        ElseIf SkinNumber = 8 Then
            With frmToolTip
                .lblText.Forecolor = vbRed
                .Line1.BorderColor = vbRed
                .Line2.BorderColor = vbRed
                .Line3.BorderColor = vbRed
                .Line4.BorderColor = vbRed
                .Picture1.BackColor = vbBlack
            End With
        ElseIf SkinNumber = 9 Then
            With frmToolTip
                .lblText.Forecolor = &HFF572D
                .Line1.BorderColor = &HFF572D
                .Line2.BorderColor = &HFF572D
                .Line3.BorderColor = &HFF572D
                .Line4.BorderColor = &HFF572D
                .Picture1.BackColor = vbBlack
            End With
        ElseIf SkinNumber = 10 Then
            With frmToolTip
                .lblText.Forecolor = TxtClr
                .Line1.BorderColor = TxtClr
                .Line2.BorderColor = TxtClr
                .Line3.BorderColor = TxtClr
                .Line4.BorderColor = TxtClr
                .Picture1.BackColor = BgClr
            End With
        End If
        
If UBound(entry) > 0 Then
    Do While pos < UBound(entry) + 1
        If Trim$(entry(pos)) <> "" Then
            If pos = 0 Then
                If frmToolTip.TextWidth(entry(pos)) * 15 > LargestLabel Then
                    LargestLabel = frmToolTip.TextWidth(entry(pos)) * 15
                End If
                
                frmToolTip.lblText.Caption = entry(pos)
            Else
                If frmToolTip.TextWidth(entry(pos)) * 15 > LargestLabel Then
                    LargestLabel = frmToolTip.TextWidth(entry(pos)) * 15
                End If
                
                frmToolTip.Picture1.Height = ((pos + 1) * 16) * 15
                PasteLabel "boe" & pos, entry(pos), 2, pos * (16), frmToolTip.Line1.BorderColor
            End If
        End If
        pos = pos + 1
    Loop
Else
    If frmToolTip.TextWidth(Str) * 15 > LargestLabel Then
        LargestLabel = frmToolTip.TextWidth(entry(pos)) * 15
    End If
    frmToolTip.Picture1.Height = 16 * 15
    frmToolTip.lblText.Caption = Str
End If

DoEvents
    With frmToolTip
        .Timer2.Enabled = False
        .Picture1.Width = LargestLabel + 100
        .Timer2.Interval = Tijd
        .SetFormShit ShowCords
        ShowIngToolTip = True
        .Timer2.Enabled = True
        .Timer1.Enabled = True
        
        SetTrans .Picture1.hWnd, 0
        .Picture1.Visible = True
        For i = 0 To 240 Step 20
            SetTrans .Picture1.hWnd, i
            Wait 0
        Next i
        
        .Picture1_Resize
    End With

End Sub


Sub HideToolTip()
    frmToolTip.SetFormShit False
    frmToolTip.Picture1.Visible = False
    frmToolTip.Timer2.Enabled = False
    frmToolTip.Timer1.Enabled = False
    ShowIngToolTip = False
End Sub


Function PasteLabel(lblName As String, lblCaption As String, xValue As Integer, yValue As Integer, TxtClr As String)
    'Create label within a picturebox at specific coordinates
    On Error Resume Next
    frmToolTip.Controls.Add "VB.label", lblName
    With frmToolTip.Controls(lblName)
        .Height = 200
        .Caption = lblCaption
        .Visible = True
        .Top = yValue
        Set .Container = frmToolTip.Picture1
        .Left = xValue
        .Width = 2000
        .Forecolor = TxtClr
        .BackStyle = 0
    End With
End Function

Function ShowMsgBox(Optional Str As String, Optional Buttons As Integer = vbOKOnly, Optional Title As String = Empty, _
                    Optional BackColor1 As Long = -1, Optional ForeColor1 As Long = -1, _
                    Optional VisibleSeconds As Integer = -1, Optional DefauldButton As String = "ok")
    Dim entry() As String
    Dim pos As Integer
    Dim LabelL As Integer
    
    MessageBeep 0
    Form1.Enabled = False
    frmToolTip.Enabled = False
    frmStopPlay.Enabled = False
    
    If Str = "" Then Str = " "
    Str = Replace(Str, "{ENTER}", vbCrLf, , , vbTextCompare)
    
    entry = Split(Str, vbCrLf, , vbTextCompare)
    
    If UBound(entry) > 0 Then
        Do While pos < UBound(entry) + 1
            If Trim$(entry(pos)) <> "" Then
                    If MessageBox.TextWidth(entry(pos)) > LabelL Then
                        LabelL = MessageBox.TextWidth(entry(pos))
                    End If
                End If
            pos = pos + 1
        Loop
    Else
        LabelL = MessageBox.TextWidth(entry(0))
    End If

    MessageBox.label1.Caption = Str
    MessageBox.SetModeOfMSGBOX (Buttons)
    
    If Title = Empty Then MessageBox.Caption = App.Title Else MessageBox.Caption = Title
    If BackColor1 < 0 Then MessageBox.SetBackColors Form1.jucko13.BackColor, Form1.jucko13.Forecolor Else MessageBox.SetBackColors BackColor1, ForeColor1
    
    MessageBox.Show
    SetTopmostWindow MessageBox.hWnd, True
    
    If VisibleSeconds > 0 Then MessageBox.SShow VisibleSeconds, DefauldButton, LabelL Else MessageBox.SShow 0, "", LabelL
    
    Do While MessageBox.Visible
    DoEvents
    Wait 50
    
    Loop
    
    Form1.Enabled = True
    frmToolTip.Enabled = True
    frmStopPlay.Enabled = True
    
    ShowMsgBox = ButtonPressed
    'If Form1.WindowState = 0 Or Form1.WindowState = 2 Then Form1.SetFocus
    
End Function

'ShowInputBox "text", Title, "wat is het normale", xPos, yPos

Function ShowInputBox(Optional Str As String, Optional Title As String = Empty, _
                    Optional ByVal Default As String, Optional BackColor1 As Long = -1, _
                    Optional ForeColor1 As Long = -1, Optional ExtraMode As Integer = 0, _
                    Optional ExtraText As String = "", Optional ExtraTextSize As Integer = 15)
    Dim entry() As String
    Dim pos As Integer
    Dim LabelL As Integer
    
    DefauldButton = "ok"
    MessageBeep 0
    Form1.Enabled = False
    
    If Str = "" Then Str = " "
    'Str = Replace(Str, "{ENTER}", vbCrLf, , , vbTextCompare)
    
    entry = Split(Str, vbCrLf, , vbTextCompare)
    
    If UBound(entry) > 0 Then
        Do While pos < UBound(entry) + 1
            If Trim$(entry(pos)) <> "" Then
                    If MessageBox.TextWidth(entry(pos)) > LabelL Then
                        LabelL = MessageBox.TextWidth(entry(pos))
                    End If
                End If
            pos = pos + 1
        Loop
    Else
        LabelL = MessageBox.TextWidth(entry(0))
    End If
    
    If ExtraMode = vbFontMode Then
        MessageBox.SetModeOfMSGBOX (vbFontMode)
        MessageBox.label1.Caption = Str
        MessageBox.Label3.Caption = ExtraText
        MessageBox.Label3.FontSize = ExtraTextSize
    Else
        MessageBox.label1.Caption = Str
        MessageBox.SetModeOfMSGBOX (InBox)
    End If
    
    If Title = Empty Then MessageBox.Caption = App.Title Else MessageBox.Caption = Title
    If BackColor1 < 0 Then MessageBox.SetBackColors Form1.jucko13.BackColor, Form1.jucko13.Forecolor Else MessageBox.SetBackColors BackColor1, ForeColor1

    MessageBox.Show
    SetTopmostWindow MessageBox.hWnd, True
    
    MessageBox.SShow 0, Default, LabelL
    MessageBox.Refresh
    Do While MessageBox.Visible: Wait 50:  Loop
    
    Form1.Enabled = True
    
    If Len(ButtonPressed) = 0 Then
        ShowInputBox = vbEmptyErr
    Else
        ShowInputBox = ButtonPressed
    End If
    'If Form1.WindowState = 0 Or Form1.WindowState = 2 Then Form1.SetFocus
    
End Function

