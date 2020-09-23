VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "RegSvr32 for Windows"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5775
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   5535
   End
   Begin VB.Frame frmFilename 
      Caption         =   "Options :"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdUnreg 
         Caption         =   "UNREGISTER FILE"
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton cmdReg 
         Caption         =   "REGISTER FILE"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label lblCopyright 
         Alignment       =   2  'Zentriert
         Caption         =   "RegSvr32 for Windows (c) 2001 Thomas Sturm"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   3735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   30
         X2              =   5520
         Y1              =   1340
         Y2              =   1340
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   30
         X2              =   5520
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblFilename 
         Caption         =   "Filename :"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdReg_Click()
If txtFilename.Text = "" Then Exit Sub
If FileExist(txtFilename.Text) Then
    If RegServer(txtFilename.Text) = True Then
        MsgBox txtFilename.Text & " was correctly registered.", vbInformation, "Success"
    Else
        MsgBox "Failure registering " & txtFilename.Text & ".", vbCritical, "Failure"
    End If
Else
    MsgBox "Specified File not found !", vbCritical, "Error"
End If
End Sub

Private Sub cmdUnreg_Click()
If txtFilename.Text = "" Then Exit Sub
If FileExist(txtFilename.Text) Then
    If UnRegServer(txtFilename.Text) = True Then
        MsgBox txtFilename.Text & " was correctly unregistered.", vbInformation, "Success"
    Else
        MsgBox "Failure unregistering " & txtFilename.Text & ".", vbCritical, "Failure"
    End If
Else
    MsgBox "Specified File not found !", vbCritical, "Error"
End If
End Sub

Function FileExist(sFileName As String) As Boolean
FileExist = IIf(Dir(sFileName) <> "", True, False)
End Function

Private Sub Command1_Click()
Load frmBrowse
frmBrowse.Show vbModal, Me
End Sub
