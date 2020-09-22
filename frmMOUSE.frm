VERSION 5.00
Begin VB.Form frmMOUSE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   2055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton mnuTRAY 
      Caption         =   "Tray"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdSELF 
      Caption         =   "Self"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdMOUSE 
      Caption         =   "Mouse"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frmMOUSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ImplodeFormToMouse Me.hWnd, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ImplodeFormToMouse Me.hWnd
End Sub
Private Sub cmdMOUSE_Click()
    Dim frm As frmMOUSE
    Set frm = New frmMOUSE
    frm.Show
End Sub
Private Sub cmdSELF_Click()
    Dim frm As frmSELF
    Set frm = New frmSELF
    frm.Show
End Sub
Private Sub mnuTRAY_Click()
    Dim frm As frmTRAY
    Set frm = New frmTRAY
    frm.Show
End Sub
