VERSION 5.00
Object = "{53DD3497-214A-11D5-9934-444553540001}#1.0#0"; "AI.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AI Control Test Application"
   ClientHeight    =   2355
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin AI.AiUnit AiUnit1 
      Left            =   2280
      Top             =   1680
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.CommandButton EditAi 
      Caption         =   "&Edit Properties"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton RunAi 
      Caption         =   "&Process Input"
      Height          =   375
      Left            =   15
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox AiOutput 
      Height          =   285
      Left            =   15
      TabIndex        =   1
      Top             =   840
      Width           =   4575
   End
   Begin VB.TextBox AiInput 
      Height          =   285
      Left            =   15
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Output from AI Control:"
      Height          =   255
      Left            =   15
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Input to AI Control:"
      Height          =   195
      Left            =   15
      TabIndex        =   4
      Top             =   0
      Width           =   1320
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit Properties"
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu MnuHlp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EditAi_Click()
'Show the edit form, so that users can modify the AIUnit's Settings
FrmEdit.Show
End Sub

Private Sub Form_Load()
'Sets the size of the form. I made the form a little bigger in edit mode so that you can see the actual AI unit.
Me.Height = 2250
Me.Width = 4695
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Close the program. I did this to prevent problems with the closing of this main form, and other forms such as the edit and help dialogs staying open
End
End Sub

Private Sub mnuAbout_Click()
'Just display some basic information about the application
MsgBox "Test Application for the AI Control by Michael Dzicek", vbOKOnly, "About the AI Application"
End Sub

Private Sub mnuEdit_Click()
'Show the edit form, so that users can modify the AIUnit's Settings
FrmEdit.Show
End Sub

Private Sub mnuHelp_Click()
'Show the help form, which displays thorough information about the AIunit and its uses.
FrmHelp.Show
End Sub

Private Sub mnuQuit_Click()
'Close the program
End
End Sub

Private Sub RunAi_Click()
'Use the AI unit to process the data in AIinput, and display the response in AIOutput
AiOutput.Text = AiUnit1.GoAI(AiInput.Text)
End Sub
