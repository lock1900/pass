Attribute VB_Name = "Mdl_Main"
Option Explicit
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetWindowPlacement _
    Lib "user32" _
    (ByVal hWnd As Long _
    , lpwndpl As WINDOWPLACEMENT) As Long
Private Declare PtrSafe Function SetWindowPlacement _
    Lib "user32" _
    (ByVal hWnd As Long _
    , lpwndpl As WINDOWPLACEMENT) As Long
Private Declare PtrSafe Function GetForegroundWindow _
    Lib "user32" () As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Sub SetForegroundWindow Lib "user32" (ByVal hWnd As Long)

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
Public uiAuto As CUIAutomation
Sub Sub_Main()
    Dim lngHeight As Long
    Dim lngWidth As Long
    
    lngWidth = Application.Width / 1.25
    lngHeight = Application.Height / 1.25
    
    Dim rng As Range
    Dim idx As Long
    Dim str���O��` As String
    Dim str�p�X As String
    For idx = 1 To 4
        Set rng = Sh_Main.Cells(idx, 2)
        str���O��` = rng.Name.Name

        str�p�X = rng.Value
        Call Sub_�E�B���h�E�z�u���C��(str���O��`, str�p�X, lngWidth, lngHeight)
        
    Next idx

End Sub
Sub Sub_�E�B���h�E�z�u���C��(ByVal str���O��` As String, ByVal str�p�X As String, ByVal lngWidth As Long, ByVal lngHeight As Long)
    Dim lngTop As Long
    Dim lngLeft As Long
  
    Select Case str���O��`
        Case "��`_����"
            lngTop = 0
            lngLeft = 0
        Case "��`_����"
            lngTop = lngHeight
            lngLeft = 0
            lngHeight = lngHeight * 2
        Case "��`_�E��"
            lngTop = 0
            lngLeft = lngWidth
            lngWidth = lngWidth * 2
        Case "��`_�E��"
            lngTop = lngHeight
            lngLeft = lngWidth
            lngHeight = lngHeight * 2
            lngWidth = lngWidth * 2
    End Select
    
    Call Sub_�E�B���h�E�z�u(str�p�X, lngTop, lngLeft, lngWidth, lngHeight)
End Sub
Sub Sub_�E�B���h�E�z�u(ByVal FPath As String, ByVal lngTop As Long, ByVal lngLeft As Long, ByVal lngWidth As Long, ByVal lngHeight As Long)
    CreateObject("Wscript.Shell").Run FPath
    DoEvents
    Application.Wait Now() + TimeValue("0:00:03")
    
    Dim str�Ώۃp�X As String
    str�Ώۃp�X = Mid(FPath, InStrRev(FPath, "\") + 1)
    Call searchWin(str�Ώۃp�X)
    
    '�E�B���h�E�n���h���̎擾
    Dim myHwnd As Long
    myHwnd = GetForegroundWindow()
    
    '�E�B���h�E���̎擾
    Dim myWindowPlacement As WINDOWPLACEMENT
    GetWindowPlacement myHwnd, myWindowPlacement
    
    '�E�B���h�E����ύX���Đݒ�
    With myWindowPlacement.rcNormalPosition
        .Left = lngLeft
        .Top = lngTop
        .Right = lngWidth
        .Bottom = lngHeight
    End With
    SetWindowPlacement myHwnd, myWindowPlacement
End Sub
Sub searchWin(ByVal title As String)
    If uiAuto Is Nothing Then
        Set uiAuto = New CUIAutomation
    End If
    
    Dim elmDeskTop As IUIAutomationElement
    Set elmDeskTop = uiAuto.GetRootElement
    
    Dim condIE As IUIAutomationCondition
    Dim aryIEControls As IUIAutomationElementArray
    
    Set condIE = uiAuto.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_WindowControlTypeId)
    Set aryIEControls = elmDeskTop.FindAll(TreeScope_Children, condIE)
    
    Dim i As Long
    Dim hdle As Variant
    Dim elm1 As IUIAutomationElement
    Dim cnd1 As IUIAutomationCondition
    
    For i = 0 To aryIEControls.Length - 1
        DoEvents
        
        If InStr(aryIEControls.GetElement(i).CurrentName, title) > 0 Then
            aryIEControls.GetElement(i).SetFocus
            Exit For
        End If
    Next i
End Sub
