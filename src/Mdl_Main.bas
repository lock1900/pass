Attribute VB_Name = "Mdl_Main"
Option Explicit
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetWindowPlacement _
    Lib "user32" _
    (ByVal hwnd As Long _
    , lpwndpl As WINDOWPLACEMENT) As Long
Private Declare PtrSafe Function SetWindowPlacement _
    Lib "user32" _
    (ByVal hwnd As Long _
    , lpwndpl As WINDOWPLACEMENT) As Long
Private Declare PtrSafe Function GetForegroundWindow _
    Lib "user32" () As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type
Sub Sub_Main()
    Dim lngHeight As Long
    Dim lngWidth As Long
    
    lngWidth = Application.Width / 1.25
    lngHeight = Application.Height / 1.25
    
    Dim rng As Range
    Dim idx As Long
    Dim str名前定義 As String
    Dim strパス As String
    For idx = 1 To 4
        Set rng = Sh_Main.Cells(idx, 2)
        str名前定義 = rng.Name.Name

        strパス = rng.Value
        Call Sub_ウィンドウ配置メイン(str名前定義, strパス, lngWidth, lngHeight)
    Next idx
'    For idx = 1 To 4
'        Set rng = Sh_Main.Cells(idx, 2)
''        str名前定義 = rng.Name.Name
'
'        strパス = rng.Value
'        CreateObject("Wscript.Shell").Run strパス
'    Next idx
'    For idx = 1 To 4
'        Set rng = Sh_Main.Cells(idx, 2)
'        str名前定義 = rng.Name.Name
'
'        strパス = rng.Value
'        Call Sub_ウィンドウ配置メイン(str名前定義, strパス, lngWidth, lngHeight)
'    Next idx
End Sub
Sub Sub_ウィンドウ配置メイン(ByVal str名前定義 As String, ByVal strパス As String, ByVal lngWidth As Long, ByVal lngHeight As Long)
    Dim lngTop As Long
    Dim lngLeft As Long
  
    Select Case str名前定義
        Case "定義_左上"
            lngTop = 0
            lngLeft = 0
        Case "定義_左下"
            lngTop = lngHeight
            lngLeft = 0
            lngHeight = lngHeight * 2
        Case "定義_右上"
            lngTop = 0
            lngLeft = lngWidth
            lngWidth = lngWidth * 2
        Case "定義_右下"
            lngTop = lngHeight
            lngLeft = lngWidth
            lngHeight = lngHeight * 2
            lngWidth = lngWidth * 2
    End Select
    
    Call Sub_ウィンドウ配置(strパス, lngTop, lngLeft, lngWidth, lngHeight)
End Sub
Sub Sub_ウィンドウ配置(ByVal FPath As String, ByVal lngTop As Long, ByVal lngLeft As Long, ByVal lngWidth As Long, ByVal lngHeight As Long)
    CreateObject("Wscript.Shell").Run FPath
    Application.Wait Now() + TimeValue("0:00:05")

    'ウィンドウハンドルの取得
    Dim myHwnd As Long
    myHwnd = GetForegroundWindow()
    
    'ウィンドウ情報の取得
    Dim myWindowPlacement As WINDOWPLACEMENT
    GetWindowPlacement myHwnd, myWindowPlacement
    
    'ウィンドウ情報を変更して設定
    With myWindowPlacement.rcNormalPosition
        .Left = lngLeft
        .Top = lngTop
        .Right = lngWidth
        .Bottom = lngHeight
    End With
    SetWindowPlacement myHwnd, myWindowPlacement
End Sub

