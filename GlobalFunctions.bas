Attribute VB_Name = "GlobalFunctions"
Option Explicit


Global Const MIM_BACKGROUND As Long = &H2
Global Const MIM_APPLYTOSUBMENUS As Long = &H80000000
Global Const MIM_MENUDATA As Long = &H8

Global Const MIM_STYLE As Long = &H10
Global Const MNS_MODELESS As Long = &H40000000
Global Const MNS_NOCHECK As Long = &H80000000
Global Const MNS_NOTIFYBYPOS As Long = &H8000000

Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type


Type MENUINFO
    cbSize As Long
    fMask As Long
    dwStyle As Long
    cyMax As Long
    hbrBack As Long
    dwContextHelpID As Long
    dwMenuData As Long
End Type

Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, mi As MENUINFO) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function GetMenuInfo Lib "user32" (ByVal hMenu As Long, lpcmi As MENUINFO) As Long

Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long

Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long


Sub main()
    
    
    uEnableMouseHooks = True
    
    
    Form1.Show
    


End Sub
