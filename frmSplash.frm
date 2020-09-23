VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   375
         Left            =   5580
         TabIndex        =   5
         Top             =   3480
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Click here to start"
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Press ESC to end Game"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   4
         Top             =   3660
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   $"frmSplash.frx":000C
         Height          =   2175
         Left            =   2520
         TabIndex        =   3
         Top             =   1260
         Width           =   4275
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":02AD
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Hello Game"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2520
         TabIndex        =   1
         Top             =   420
         Width           =   3540
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'---------------------------------------------------------------------------------------
' Procedure : Command1_Click
' DateTime  : 10/1/02 14:06
' Author    : Mike S. Firoved
' Owner     : hellogme
'---------------------------------------------------------------------------------------
Private Sub Command1_Click()
    On Error GoTo ERRCommand1_Click
    Me.Hide
    With frmGameBoard
        .Show
        .tmrFireGun.Enabled = True
        .tmrAimStar.Enabled = True
        .tmrCollisionTest.Enabled = True
        .tmrExploStar.Enabled = True
        .tmrGameBoundry.Enabled = True
        .tmrManageShots.Enabled = True
        .tmrManageStars.Enabled = True
        .tmrMutateStarBehavior.Enabled = True
        .tmrStarLauncher.Enabled = True
        .pctGameBoard.SetFocus
    End With
    Exit Sub
    
ERRCommand1_Click:
    SaveSetting "hellogme", "error", "2nd", GetSetting("hellogme", "error", "Last", "")
    SaveSetting "hellogme", "error", "3rd", GetSetting("hellogme", "error", "2nd", "")
    SaveSetting "hellogme", "error", "Last", "hellogme.frmSplash.Command1_Click." & App.ThreadID & " Produced the following error:  Error #" & Err.Number & "," & Err.Description
    Resume Next
End Sub

Private Sub Command2_Click()
    End
End Sub
