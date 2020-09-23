VERSION 5.00
Begin VB.Form FrmHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Artificial Intelligence Unit Help"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FrmHelp.frx":0000
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'Close the Help Form
Unload Me
End Sub
