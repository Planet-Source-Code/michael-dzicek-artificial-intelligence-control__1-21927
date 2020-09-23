VERSION 5.00
Begin VB.UserControl AiUnit 
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4260
   InvisibleAtRuntime=   -1  'True
   Picture         =   "AiOCX.ctx":0000
   PropertyPages   =   "AiOCX.ctx":030A
   ScaleHeight     =   2085
   ScaleWidth      =   4260
   ToolboxBitmap   =   "AiOCX.ctx":0321
   Begin VB.ListBox List2 
      Height          =   1425
      ItemData        =   "AiOCX.ctx":0633
      Left            =   0
      List            =   "AiOCX.ctx":0685
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox List3 
      Height          =   1425
      ItemData        =   "AiOCX.ctx":07F3
      Left            =   2160
      List            =   "AiOCX.ctx":0845
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "AiUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Compare Text
'Create the public type that output by the GetStatement Function
Public Type AIStatement
 Input As String
 Response As String
End Type
'All of the subs and functions in this OCX are more thoroughly explained in the property page, and the about file.
Public Sub DumpLists()
'empty out the statement/resposne lists
List2.Clear
List3.Clear
End Sub

Public Sub AddStatement(StatementToGoIn As String, StatementToGoOut As String)
'Add a statement/response combination to the internal lists.
'Here we use the CleanUpThisMess function to ensure maximum compatibility in the GoAI function.
List2.AddItem CleanUpThisMess(StatementToGoIn)
'Add the statement response
List3.AddItem StatementToGoOut
End Sub

Public Function StatementAmount() As Long
'return the number of entries in the statements list. Should give an exact output of the statement/resposne combinations
StatementAmount = List2.ListCount
End Function

Public Function GetStatement(Index As Long) As AIStatement
'Make sure that the given index is in the bounds of the internal lists.
If Index <= List2.ListCount - 1 And Index <= List3.ListCount - 1 And Index >= 0 Then
 'if it is, then return the entry corresponding to the given index.
 GetStatement.Input = List2.List(Index)
 GetStatement.Response = List3.List(Index)
Else
 'If the number was out of bounds, or the entry did not exist, then return null responses.
 GetStatement.Input = vbNullString
 GetStatement.Response = vbNullString
End If
End Function

Public Function RemoveStatement(Index As Long)
'Make sure that the given index is in the bounds of the internal lists.
If Index <= List2.ListCount - 1 And Index <= List3.ListCount - 1 And Index >= 0 Then
 'If it is, then remove the statement/response combinations.
 List1.RemoveItem Index
 List2.RemoveItem Index
End If
End Function

Public Function CleanUpThisMess(Mess As String) As String
Dim MessInput() As String, Word As Variant, TempClean As String
'make the statement lowercase, so that capitalization does not become an issue later on in the search process(es)
Mess = LCase(Mess)
'trim any spaces that may be on the end, so that they are not an issue later as fas as case sensitivity goes.
While Right(Mess, 1) = " "
 Mess = Left(Mess, Len(Mess) - 1)
Wend
'Split the given phrase apart, so that each word can be scanned individually, and converted to a standardized form.
MessInput = Split(Mess, " ")
'Now, we check each individual word, and convert it to a standard form if we have to.
For Each Word In MessInput
 'Here, we trim off any non alphabetic characters that are in the word. This opens up a slight error possibility by the exclusion of numbers, but i have had very few problems with it, and find
 'That it helps the scanning by sometimes removing typos.
 While IsNumeric(Right(Word, 1)) = False And Right(Word, 1) Like "[A-Z]" = False And Right(Word, 1) Like "[a-z]" = False
  Word = Left(Word, Len(Word) - 1)
 Wend
 'Scan it again, but this time we do it from the left, ratehr than the right.
 While IsNumeric(Left(Word, 1)) = False And Left(Word, 1) Like "[A-Z]" = False And Left(Word, 1) Like "[a-z]" = False
  Word = Right(Word, Len(Word) - 1)
 Wend
  'This is where the bulk of the processing happens. here many forms of slang are processed into more english-friendly equivalents. This helps standardize the input. Feel free to add to or modify this list.
  'Just make sure that you write the replacement in lowercase, otherwise you could have problems with case sensitivity later.
  If Word = "c" Then Word = "see"
  If Word = "ic" Then Word = "i see"
  If Word = "oic" Then Word = "oh i see"
  If Word = "u" Then Word = "you"
  If Word = "ya" Then Word = "you"
  If Word = "ur" Then Word = "you are"
  If Word = "cu" Then Word = "see you"
  If Word = "im" Then Word = "i am"
  If Word = "lol" Then Word = "laughs out loud"
  If Word = "rotf" Then Word = "rolling on the floor"
  If Word = "rotflmao" Then Word = "rolling on the floor laughing my ass off"
  If Word = "brb" Then Word = "be right back"
  If Word = "dunno" Then Word = "don't know"
  If Word = "dont" Then Word = "don't"
  If Word = "cant" Then Word = "can't"
  If Word = "wtf" Then Word = "what the fuck"
  If Word = "r" Then Word = "are"
  If Word = "b" Then Word = "be"
  If Word = "whos" Then Word = "who is"
  If Word = "goin" And Left(Word, 5) <> "going" Then Word = "going"
  If Word = "who's" Then Word = "who is"
  If Word = "whats" Then Word = "what is"
  If Word = "what's" Then Word = "what is"
  If Word = "whens" Then Word = "when is"
  If Word = "when's" Then Word = "when is"
  If Word = "wheres" Then Word = "where is"
  If Word = "where's" Then Word = "where is"
  If Word = "whys" Then Word = "why is"
  If Word = "why's" Then Word = "why is"
  If Word = "hows" Then Word = "how is"
  If Word = "how's" Then Word = "how is"
  If Word = "whatcha" Then Word = "what are you"
  If Word = "whatya" Then Word = "what are you"
  If Word = "whacha" Then Word = "what are you"
  If Word = "wazzap" Then Word = "what is up"
  If Word = "wussup" Then Word = "what is up"
  If Word = "sup" Then Word = "what is up"
  If Word = "wuzzup" Then Word = "what is up"
  If Word = "doin" Then Word = "doing"
  'Here the new sentance is formed from the cleaned words by joining the sentance formed up to this point with the new word and a space.
  TempClean = TempClean & " " & Word
Next
'We trim off the last space of the sentance for both cosmetic, and possible case matching reasons.
CleanUpThisMess = Right(TempClean, Len(TempClean) - 1)
End Function
Function GoAI(AInput As String) As String
Dim Count As Long, TempAI As String
'Clean up the input first, in order to standardize the language in the lists, and in the input
AInput = CleanUpThisMess(AInput)
'Check first to make sure that both the statement and response lists are the same legnth, in order to prevent finding a statement with a null answer.
If List2.ListCount <> List3.ListCount Then GoTo ErrorEnd
'run through the entire list of possible statements, in order to look for our input.
For Count = 0 To (List2.ListCount - 1)
 'if we find our input statement in the list, then transfer the response into a temporary variable
 If InStr(AInput, List2.List(Count)) > 0 Then
  TempAI = List3.List(Count)
  Exit For
 End If
Next
'now assign the value of the function to that of the temporary variable so that the value can be retrieved.
GoAI = TempAI
'exit the function here, so that the error handler is not called by accident.
Exit Function
'Error handler
ErrorEnd:
MsgBox "The # of inputs does not match the outputs."
End Function
Private Sub UserControl_Resize()
'This just standardizes the size of the aiunit so that it is not expanded to a large size, much in the way that controls such as the timer are.
UserControl.Height = 33 * Screen.TwipsPerPixelY
UserControl.Width = 33 * Screen.TwipsPerPixelX
End Sub
