Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------------
' Module    : modMain
' DateTime  : 10/1/02 13:42
' Author    : Mike S. Firoved
'---------------------------------------------------------------------------------------
Public leftPass As Long
Public rightPass As Long
Public fireNow As Long
Public upPass As Long
Public downPass As Long

Public curLevel As Long

Public gameOver As Long

Public moveRate As Long  '= 64
Public shotRate As Long  '= 128
Public starRate As Long  '= 64
Public managedShots As New Collection
Public managedStars As New Collection

Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'---------------------------------------------------------------------------------------
' Procedure : findAvailStar
' DateTime  : 10/1/02 13:41
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Public Function findAvailStar() As Integer

    'SEARCH FROM BOTTOM TO TOP
    On Error GoTo ERRfindAvailStar
    For xd = (frmGameBoard.pctStar.Count - 1) To 0 Step -1
        sts = frmGameBoard.pctStar(xd).Visible
        If sts = False Then
            findAvailStar = xd
            Exit Function
        End If
    Next
    findAvailStar = -1
    Exit Function
    
ERRfindAvailStar:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.modMain.findAvailStar." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    Resume Next
End Function

'---------------------------------------------------------------------------------------
' Procedure : findAvailShot
' DateTime  : 10/1/02 13:41
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Public Function findAvailShot() As Integer

    'SEARCH FROM BOTTOM TO TOP
    On Error GoTo ERRfindAvailShot
    For xd = (frmGameBoard.pctShot.Count - 1) To 0 Step -1
        sts = frmGameBoard.pctShot(xd).Visible
        If sts = False Then
            findAvailShot = xd
            Exit Function
        End If
    Next
    findAvailShot = -1
    Exit Function
    
ERRfindAvailShot:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.modMain.findAvailShot." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    Resume Next
End Function


'---------------------------------------------------------------------------------------
' Procedure : endNow
' DateTime  : 10/1/02 13:40
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Public Sub endNow()
    On Error GoTo ERRendNow
    ShowCursor 1
    End
    Exit Sub
    
ERRendNow:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.modMain.endNow." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    Resume Next
End Sub
