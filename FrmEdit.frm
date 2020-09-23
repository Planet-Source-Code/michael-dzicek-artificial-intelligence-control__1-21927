VERSION 5.00
Begin VB.Form FrmEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit AI Properties"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Update 
         Caption         =   "Add / Update"
         Height          =   615
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox OutputTxt 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox InputTxt 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Output:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Input:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "AI Inputs and Outputs (Double Click an Entry to delete it)"
      Height          =   2175
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
      Begin VB.ListBox AiInputs 
         Height          =   1815
         ItemData        =   "FrmEdit.frx":0000
         Left            =   120
         List            =   "FrmEdit.frx":0002
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.ListBox AiOutputs 
         Height          =   1815
         ItemData        =   "FrmEdit.frx":0004
         Left            =   2400
         List            =   "FrmEdit.frx":0006
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton OK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "FrmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AiInputs_Click()
'Match up the listfocuses so that we can visually see which statement goes with which response.
AiOutputs.ListIndex = AiInputs.ListIndex
'Fill the statement/resposne textboxes with the selected entry, just in case someone wants to easily modify it.
InputTxt.Text = AiInputs.List(AiInputs.ListIndex)
OutputTxt.Text = AiOutputs.List(AiOutputs.ListIndex)
End Sub

Private Sub AiInputs_DblClick()
'Delete the entries in the lists, because after all, thats why the list was double clicked.
AiOutputs.RemoveItem AiInputs.ListIndex
AiInputs.RemoveItem AiInputs.ListIndex
'Clean out the textboxes just in case the double clicking caused them to be filled with the now deleted data.
InputTxt.Text = ""
OutputTxt.Text = ""
End Sub

Private Sub AiOutputs_Click()
'Match up the listfocuses so that we can visually see which statement goes with which response.
AiInputs.ListIndex = AiOutputs.ListIndex
'Fill the statement/resposne textboxes with the selected entry, just in case someone wants to easily modify it.
InputTxt.Text = AiInputs.List(AiInputs.ListIndex)
OutputTxt.Text = AiOutputs.List(AiOutputs.ListIndex)
End Sub

Private Sub AiOutputs_DblClick()
'Delete the entries in the lists, because after all, thats why the list was double clicked.
AiInputs.RemoveItem AiOutputs.ListIndex
AiOutputs.RemoveItem AiOutputs.ListIndex
'Clean out the textboxes just in case the double clicking caused them to be filled with the now deleted data.
InputTxt.Text = ""
OutputTxt.Text = ""
End Sub

Private Sub OK_Click()
Dim Counter As Long
'Clear the unit's lists so that there is no repeat data in the unit, which allows for potential errors
FrmMain.AiUnit1.DumpLists
'Refill the AIUnit's statement/response lists with the lists that are contained within the edit form.
For Counter = 0 To (AiInputs.ListCount - 1)
 FrmMain.AiUnit1.AddStatement AiInputs.List(Counter), AiOutputs.List(Counter)
Next Counter
'Unload the form. We must use the unload insted of hide, otherwise when the form is reloaded, the listboxes will not be updated
Unload Me
End Sub

Private Sub Cancel_Click()
'Unload the form. We must use the unload insted of hide, otherwise when the form is reloaded, the listboxes will not be updated
Unload Me
End Sub

Private Sub Form_Load()
Dim Statements As AIStatement, Counter As Long
'clear out the listboxes so that there will not be repeat data in them, which could cause an error.
AiInputs.Clear
AiOutputs.Clear
'Fill in The two list boxes with the Statement/Response data that is retireved from the unit's internal lists.
For Counter = 0 To (FrmMain.AiUnit1.StatementAmount - 1)
 Statements = FrmMain.AiUnit1.GetStatement(Counter)
 AiInputs.AddItem Statements.Input
 AiOutputs.AddItem Statements.Response
Next
End Sub

Private Sub Update_Click()
Dim Index As Long
'First, Check to make sure that Both the Statement and Response fields are filled in
If InputTxt.Text > "" And OutputTxt.Text > "" Then
 'Run through the Statement List to make sure that the Newly Input Statement isnt already there
 For Index = 0 To (AiInputs.ListCount - 1)
  'We use the CleanUpThisMess Statement here to ensure that input statement isnt just a rework of a previous one
  If FrmMain.AiUnit1.CleanUpThisMess(InputTxt.Text) = AiInputs.List(Index) Then Exit For
 Next
 'At this point the the previous entry in the lists is overwritten if one was alreadyt found
 'If the previous loop did not find the statement, then this will just add the new statement to the end of the lists
 AiInputs.List(Index) = FrmMain.AiUnit1.CleanUpThisMess(InputTxt.Text)
 AiOutputs.List(Index) = OutputTxt.Text
 'Empty the textboxes and set foceus to the Statement box, just for cosmetic purposes
 InputTxt.Text = ""
 OutputTxt.Text = ""
 InputTxt.SetFocus
Else
 'in case one of the earlier text boxes was not filled in, display and error, and set focus to the appropriate box.
 MsgBox "Enter both a valid Input and Output statement."
 If InputTxt.Text = "" Then InputTxt.SetFocus
 If OutputTxt.Text = "" Then OutputTxt.SetFocus
End If
End Sub
