VERSION 5.00
Begin VB.Form frmGameBoard 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "HelloGame"
   ClientHeight    =   7785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11190
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrManageStars 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   4260
      Top             =   900
   End
   Begin VB.Timer tmrMutateStarBehavior 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3840
      Top             =   900
   End
   Begin VB.Timer tmrAimStar 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   3420
      Top             =   900
   End
   Begin VB.Timer tmrGameBoundry 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   1320
      Top             =   900
   End
   Begin VB.Timer tmrStarLauncher 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3000
      Top             =   900
   End
   Begin VB.Timer tmrExploStar 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   2580
      Top             =   900
   End
   Begin VB.Timer tmrCollisionTest 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   2160
      Top             =   900
   End
   Begin VB.Timer tmrManageShots 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   1740
      Top             =   900
   End
   Begin VB.Timer tmrFireGun 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   900
      Top             =   900
   End
   Begin VB.PictureBox pctGameBoard 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   7995
      Left            =   2700
      ScaleHeight     =   7995
      ScaleWidth      =   11055
      TabIndex        =   0
      Top             =   1800
      Width           =   11055
      Begin VB.PictureBox pctExplo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   600
         Picture         =   "frmGameBoard.frx":0000
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   4
         Top             =   2040
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox pctShot 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   0
         Left            =   1260
         Picture         =   "frmGameBoard.frx":055B
         ScaleHeight     =   435
         ScaleWidth      =   375
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox pctStar 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   0
         Left            =   120
         Picture         =   "frmGameBoard.frx":08E1
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox pctShip 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   5820
         Picture         =   "frmGameBoard.frx":0D54
         ScaleHeight     =   435
         ScaleWidth      =   375
         TabIndex        =   1
         Top             =   2940
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblGameOver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G A M E  O V E R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   2640
         TabIndex        =   5
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmGameBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' DateTime  : 10/1/02 13:34
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ERRForm_Load
    With pctGameBoard
        .Top = 0
        .Left = 0
        .Width = Screen.Width
        .Height = Screen.Height
    End With
    Randomize Int(Timer / 256) * 256 + 3.1415
    DoEvents
    
    'center gameover label (hidden now)
    lblGameOver.Top = (pctGameBoard.ScaleHeight - lblGameOver.Height) / 2
    lblGameOver.Left = (pctGameBoard.ScaleWidth - lblGameOver.Width) / 2
    
    'set the speed of the objects
    moveRate = 64
    shotRate = 128
    starRate = 32
    
    pctShip.Top = (pctGameBoard.ScaleHeight - pctShip.Height - 50) / 2
    pctShip.Visible = True
    
    'load all the objects
    For cv = 1 To 64
        Load pctShot(cv)
        Load pctStar(cv)
        
        'randomize behavior patterns. Using fillcolor see notes
        pctStar(cv).FillColor = RGB(Int(Rnd(2) * 255), Int(Rnd(2) * 255), Int(Rnd(2) * 255))
        pctStar(cv).ForeColor = 0
    Next cv
    
    curLevel = 1
    
    'hide cursor
    ShowCursor 0
    Exit Sub
    
ERRForm_Load:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.frmGameBoard.Form_Load." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    
    Resume Next
End Sub



'---------------------------------------------------------------------------------------
' Procedure : pctGameBoard_KeyDown
' DateTime  : 10/1/02 13:34
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Private Sub pctGameBoard_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo ERRpctGameBoard_KeyDown
    Select Case KeyCode

        Case vbKeySpace
            fireNow = 1

        Case vbKeyEnd
            endNow

        Case vbKeyLeft
            If leftPass = 1 Then pctShip.Left = pctShip.Left - moveRate
            
        Case vbKeyUp
            If upPass = 1 Then pctShip.Top = pctShip.Top - moveRate

        Case vbKeyRight
            If rightPass = 1 Then pctShip.Left = pctShip.Left + moveRate

        Case vbKeyDown
            If downPass = 1 Then pctShip.Top = pctShip.Top + moveRate

        Case vbKeyNumpad1
            If downPass = 1 Then pctShip.Top = pctShip.Top + moveRate
            If leftPass = 1 Then pctShip.Left = pctShip.Left - moveRate

        Case vbKeyNumpad2
            If downPass = 1 Then pctShip.Top = pctShip.Top + moveRate

        Case vbKeyNumpad3
            If downPass = 1 Then pctShip.Top = pctShip.Top + moveRate
            If rightPass = 1 Then pctShip.Left = pctShip.Left + moveRate

        Case vbKeyNumpad4
            If leftPass = 1 Then pctShip.Left = pctShip.Left - moveRate

        Case vbKeyNumpad5
            fireNow = 1

        Case vbKeyNumpad6
            If rightPass = 1 Then pctShip.Left = pctShip.Left + moveRate

        Case vbKeyNumpad7
            If upPass = 1 Then pctShip.Top = pctShip.Top - moveRate
            If leftPass = 1 Then pctShip.Left = pctShip.Left - moveRate

        Case vbKeyNumpad8
            If upPass = 1 Then pctShip.Top = pctShip.Top - moveRate

        Case vbKeyNumpad9
            If upPass = 1 Then pctShip.Top = pctShip.Top - moveRate
            If rightPass = 1 Then pctShip.Left = pctShip.Left + moveRate

    End Select
    Exit Sub
    
ERRpctGameBoard_KeyDown:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.frmGameBoard.pctGameBoard_KeyDown." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    
    Resume Next
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : pctGameBoard_KeyUp
' DateTime  : 10/1/02 13:43
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Private Sub pctGameBoard_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo ERRpctGameBoard_KeyUp
    If KeyCode = vbKeyEscape Then endNow
    Exit Sub
    
ERRpctGameBoard_KeyUp:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.frmGameBoard.pctGameBoard_KeyUp." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrAimStar_Timer
' DateTime  : 10/1/02 13:33
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Private Sub tmrAimStar_Timer()
    'Using semi-ai
    On Error GoTo ERRtmrAimStar_Timer
    For nm = 1 To managedStars.Count
        xc = managedStars.Item(nm).FillColor
        rd = Val("&h" & Mid(Hex(xc), 1, 2))
        gr = Val("&h" & Mid(Hex(xc), 3, 2))
        bl = Val("&h" & Mid(Hex(xc), 5, 2))
        
        cr = rd - bl
        If Abs(cr) <> cr Then cr = 0
        
        cg = gr - bl
        If Abs(cg) <> cg Then cg = 0
        
        msl = managedStars.Item(nm).Left
        
        If cr = 0 And cg = 0 Then Exit Sub
        
        starAimRate = starRate / 2
        
        If cr > cg Then
            If pctShip.Left > msl Then
                managedStars.Item(nm).Left = managedStars.Item(nm).Left + starAimRate
            Else
                managedStars.Item(nm).Left = managedStars.Item(nm).Left - starAimRate
            End If
        Else
            If pctShip.Left > msl Then
                managedStars.Item(nm).Left = managedStars.Item(nm).Left - starAimRate
            Else
                managedStars.Item(nm).Left = managedStars.Item(nm).Left + starAimRate
            End If
        End If
    Next nm
    If gameOver = 1 Then
        lblGameOver.FontSize = lblGameOver.FontSize + 2
        If lblGameOver.FontSize > 144 Then
            gameOver = 0
            lblGameOver.Visible = False
            endNow
        End If
        lblGameOver.Top = (pctGameBoard.ScaleHeight - lblGameOver.Height) / 2
        lblGameOver.Left = (pctGameBoard.ScaleWidth - lblGameOver.Width) / 2
        pctGameBoard.SetFocus
    End If
    Exit Sub
    
ERRtmrAimStar_Timer:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.frmGameBoard.tmrAimStar_Timer." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrGameBoundry_Timer
' DateTime  : 10/1/02 13:38
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Private Sub tmrGameBoundry_Timer()
    On Error GoTo ERRtmrGameBoundry_Timer
    leftPass = 1
    If pctShip.Left <= pctGameBoard.Left Then leftPass = 0
    rightPass = 1
    If pctShip.Left >= pctGameBoard.ScaleWidth - pctShip.Width Then rightPass = 0
    upPass = 1
    If pctShip.Top <= (pctStar(0).Height * 3) Then upPass = 0
    downPass = 1
    If pctShip.Top >= pctGameBoard.ScaleHeight - pctShip.Height - 50 Then downPass = 0
    Exit Sub
    
ERRtmrGameBoundry_Timer:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.frmGameBoard.tmrGameBoundry_Timer." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrFireGun_Timer
' DateTime  : 10/1/02 13:38
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Private Sub tmrFireGun_Timer()
    'Fire shots
    On Error GoTo ERRtmrFireGun_Timer
    If fireNow = 1 Then
        fireNow = 0
        xa = findAvailShot
        
        If xa = -1 Then Exit Sub 'no more shots left, only 64 at one time
        
        pctShot(xa).Left = pctShip.Left
        pctShot(xa).Top = pctShip.Top - pctShip.Height
        pctShot(xa).Visible = True
        managedShots.Add pctShot(xa)
    End If
    DoEvents
    Exit Sub
    
ERRtmrFireGun_Timer:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.frmGameBoard.tmrFireGun_Timer." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrManageShots_Timer
' DateTime  : 10/1/02 13:39
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Private Sub tmrManageShots_Timer()
    'Manage shots
    On Error GoTo ERRtmrManageShots_Timer
    For cv = 1 To managedShots.Count
        With managedShots
            .Item(cv).Top = .Item(cv).Top - shotRate
        End With
    Next
    For cv = 1 To managedShots.Count
        If managedShots.Item(cv).Top <= pctGameBoard.Top Then
            managedShots.Item(cv).Visible = False
            'Unload managedShots.Item(cv)
            managedShots.Remove cv
            Exit For
        End If
    Next
    Exit Sub
    
ERRtmrManageShots_Timer:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.frmGameBoard.tmrManageShots_Timer." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrCollisionTest_Timer
' DateTime  : 9/23/02 18:27
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Private Sub tmrCollisionTest_Timer()
    'collision test
    On Error GoTo ERRtmrCollisionTest_Timer
    For df = 1 To managedStars.Count
        tstar = managedStars.Item(df).Top
        lstar = managedStars.Item(df).Left
        hstar = managedStars.Item(df).Height
        wstar = managedStars.Item(df).Width
        For cv = 1 To managedShots.Count
            If managedShots.Item(cv).Top <= (tstar + hstar) Then
                msl = managedShots.Item(cv).Left
                msr = managedShots.Item(cv).Left + managedShots.Item(cv).Width
                If (msl < (lstar + wstar) And msl > lstar) Or (msr < (lstar + wstar) And msr > lstar) Then
                    managedStars.Item(df).Picture = pctExplo.Picture
                    managedStars.Item(df).Tag = "MFD"
                    managedStars.Item(df).ForeColor = 0
                    managedShots.Item(cv).Visible = False
                    managedShots.Remove cv
                End If
            End If
        Next cv
        psl = pctShip.Left
        psr = pctShip.Left + pctShip.Width
        If pctShip.Top < (tstar + hstar) And pctShip.Top + pctShip.Height > tstar Then
            If (psl < (lstar + wstar) And psl > lstar) Or (psr < (lstar + wstar) And psr > lstar) Then
                managedStars.Item(df).Picture = pctExplo.Picture
                managedStars.Item(df).Tag = "MFD"
                pctShip.Visible = False
                gameOver = 1
                For mn = 1 To 64
                    If managedStars.Item(mn).Tag = "" Then managedStars.Item(mn).Visible = False
                    managedShots.Item(mn).Visible = False
                Next mn
                tmrFireGun.Enabled = False
                tmrGameBoundry.Enabled = False
                tmrManageShots.Enabled = False
                tmrCollisionTest.Enabled = False
                tmrStarLauncher.Enabled = False
                lblGameOver.Visible = True
                Exit Sub
            End If
        End If
    Next df
    Exit Sub
    
ERRtmrCollisionTest_Timer:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.frmGameBoard.tmrCollisionTest_Timer." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrExploStar_Timer
' DateTime  : 9/24/02 17:56
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Private Sub tmrExploStar_Timer()
    On Error GoTo ERRtmrExploStar_Timer
    For rt = 1 To managedStars.Count
        If managedStars.Item(rt).Tag = "SFD" Then
            managedStars.Item(rt).Tag = ""
            managedStars.Item(rt).Visible = False
            managedStars.Item(rt).Picture = pctStar(0).Picture
            managedStars.Remove rt
            If gameOver = 1 Then tmrExploStar.Enabled = False
            Exit For
        End If
        If managedStars.Item(rt).Tag = "MFD" Then
            managedStars.Item(rt).Tag = "SFD"
        End If
    Next rt
    Exit Sub
    
ERRtmrExploStar_Timer:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.frmGameBoard.tmrExploStar_Timer." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrManageStars_Timer
' DateTime  : 9/26/02 18:06
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Private Sub tmrManageStars_Timer()
    On Error GoTo ERRtmrManageStars_Timer
    For cv = 1 To managedStars.Count
        With managedStars
            .Item(cv).Top = .Item(cv).Top + starRate
        End With
    Next
    For cv = 1 To managedStars.Count
        If managedStars.Item(cv).Top >= pctGameBoard.ScaleHeight Then
            managedStars.Item(cv).Visible = False
            managedStars.Item(cv).ForeColor = 1
            managedStars.Remove cv
            Exit For
        End If
    Next
    Exit Sub
    
ERRtmrManageStars_Timer:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.frmGameBoard.tmrManageStars_Timer." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrMutateStarBehavior_Timer
' DateTime  : 9/24/02 17:56
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Private Sub tmrMutateStarBehavior_Timer()
    On Error GoTo ERRtmrMutateStarBehavior_Timer
    'The star's behavior is stored in it's fillcolor.
    'this was convient because each object carried it's own data
    'So you might be asking why I didn't use the tag property,
    'because it is being used by the explo animation
    'If a star makes it past the ship then the we do not "mutate" it's behavior
    'Behavior is stored using RGB, red=aggresiveness. green=timidness and blue=inhibitor
    pctGameBoard.Cls
    For sd = 1 To 64
        
        If pctStar(sd).ForeColor = 0 Then
            xc = pctStar(sd).FillColor
            
            lr = CStr(Val("&h" & Mid(Hex(xc), 1, 2)))
            lg = CStr(Val("&h" & Mid(Hex(xc), 3, 2)))
            lb = CStr(Val("&h" & Mid(Hex(xc), 5, 2)))
            
            rd = Val("&h" & Mid(Hex(xc), 1, 2))
            gr = Val("&h" & Mid(Hex(xc), 3, 2))
            bl = Val("&h" & Mid(Hex(xc), 5, 2))
            bl = Int(bl * (Rnd(2) + 0.5))
            gr = Int(gr * (Rnd(2) + 0.5))
            rd = Int(rd * (Rnd(2) + 0.5))
            If Rnd(2) > 0.99 Then rd = rd + 255
            If Rnd(2) > 0.99 Then gr = gr + 255
            If Rnd(2) > 0.999 Then rd = rd = 0
            If Rnd(2) > 0.999 Then gr = gr = 0
            If Rnd(2) > 0.6 Then bl = bl + 255
            pctStar(sd).FillColor = RGB(rd, gr, bl)
        End If
        DoEvents
    Next sd
    
    
    Exit Sub
    
ERRtmrMutateStarBehavior_Timer:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.frmGameBoard.tmrMutateStarBehavior_Timer." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrStarLauncher_Timer
' DateTime  : 10/1/02 13:40
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Private Sub tmrStarLauncher_Timer()
    On Error GoTo ERRtmrStarLauncher_Timer
    xv = Int(Rnd(2) * (6 - curLevel))
    Debug.Print xv
    If xv = 1 Then 'drop star
        xa = findAvailStar
        rdl = Int(Rnd(2) * (pctGameBoard.ScaleWidth - pctStar(xa).Width))
        pctStar(xa).Left = rdl
        pctStar(xa).Top = 0
        pctStar(xa).Visible = True
        managedStars.Add pctStar(xa)
    End If
    Exit Sub
    
ERRtmrStarLauncher_Timer:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.frmGameBoard.tmrStarLauncher_Timer." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    Resume Next
End Sub


