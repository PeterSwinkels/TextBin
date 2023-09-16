VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextBinClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class contains the Text Bin interface.
Option Explicit

'The events provided by this class.
Public Event FoundText(Text As String, ByRef Continue As Boolean)   'The event raised when a character is found that is not considered to be text.
Public Event HandleError(ErrorO As Object)                          'The event raised when an error occurs.

'This enumeration lists the character categories used by this class.
Public Enum CharacterCategoriesE
   CCUnreadable = 0           'The character is not "human readable."
   CCRangeCharacter           'The character falls inside the specified range.
   CCAdditionalCharacter      'The character is a "human readable" character outside the specified range.
   CCExcludedCharacter        'The character is excluded from the "human readable" characters.
   CCUnicodeNullCharacter     'The character is null character between two "human readable" characters.
End Enum

'This enumeration lists the unicode handling options for this class.
Public Enum UnicodeOptionsE
   UOExcludeUnicode = 0       'Unicode is not checked for or included in a search.
   UOIncludeUnicode           'Unicode is checked for and included in a search.
   UOExclusiveUnicode         'Unicode is checked for and a search result is only included when unicode is detected.
End Enum

'This structure  defines what is considered to be text.
Private Type TextDefinitionStr
   RangeStart As Long                 'Defines the first character in the human readable character range.
   RangeEnd As Long                   'Defines the last character in the human readable character range.
   Additional As String               'Defines any characters outside the defined range, but should be included.
   Excluded As String                 'Defines any characters inside the defined range, but should be excluded.
   UnicodeOption As UnicodeOptionsE   'Defines the unicode handling option.
End Type

'The constants used by this class:
Private Const NullCharacter As Long = 0   'Defines the null character.

'The variables used by this class:
Private TextDefinition As TextDefinitionStr  'Contains the definition of what is considered to be text.



'This procedure returns the specified character's category.
Public Function CharacterCategory(Character As Long, Optional PreviousCharacter As Long = Empty, Optional NextCharacter As Long = Empty, Optional NextNextCharacter As Long = Empty) As CharacterCategoriesE
On Error GoTo ErrorTrap
Dim Category As CharacterCategoriesE

   With TextDefinition
      If InStrB(.Excluded, Chr$(Character)) > 0 Then
         Category = CCExcludedCharacter
      ElseIf Character = NullCharacter And NextNextCharacter = NullCharacter Then
         If Not .UnicodeOption = UOExcludeUnicode Then
            If Not (PreviousCharacter = Empty Or NextCharacter = Empty) Then
               If InStrB(.Excluded, Chr$(PreviousCharacter)) = 0 And InStrB(.Excluded, Chr$(NextCharacter)) = 0 Then
                  If (PreviousCharacter >= .RangeStart And PreviousCharacter <= .RangeEnd) Or (InStrB(.Additional, Chr$(PreviousCharacter)) > 0) Then
                     If (NextCharacter >= .RangeStart And NextCharacter <= .RangeEnd) Or (InStrB(.Additional, Chr$(NextCharacter)) > 0) Then
                        Category = CCUnicodeNullCharacter
                     End If
                  End If
               End If
            End If
         End If
      ElseIf Character >= .RangeStart And Character <= .RangeEnd Then
         Category = CCRangeCharacter
      ElseIf InStrB(.Additional, Chr$(Character)) > 0 Then
         Category = CCAdditionalCharacter
      Else
         Category = CCUnreadable
      End If
   End With
   
EndRoutine:
   CharacterCategory = Category
   Exit Function
   
ErrorTrap:
   RaiseEvent HandleError(Err)
   Resume EndRoutine
End Function




'This procedure changes the text definition used by this class.
Public Sub DefineText(RangeStart As Long, RangeEnd As Long, Optional Additional As String = vbNullString, Optional Excluded As String = vbNullString, Optional UnicodeOption As UnicodeOptionsE = UOExcludeUnicode)
On Error GoTo ErrorTrap
   With TextDefinition
      .RangeStart = RangeStart
      .RangeEnd = RangeEnd
      .Additional = Additional
      .Excluded = Excluded
      .UnicodeOption = UnicodeOption
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   RaiseEvent HandleError(Err)
   Resume EndRoutine
End Sub


'This procedure returns the text definition used by this class.
Public Sub GetTextDefinition(Optional RangeStart As Long = Empty, Optional RangeEnd As Long = Empty, Optional Additional As String = Empty, Optional Excluded As String = vbNullString, Optional UnicodeOption As UnicodeOptionsE = UOExcludeUnicode)
On Error GoTo ErrorTrap
   With TextDefinition
      RangeStart = .RangeStart
      RangeEnd = .RangeEnd
      Additional = .Additional
      Excluded = .Excluded
      UnicodeOption = .UnicodeOption
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   RaiseEvent HandleError(Err)
   Resume EndRoutine
End Sub


'This procedure searches the binary data for strings of human readable characters.
Public Sub FindText(BinaryData() As Byte)
On Error GoTo ErrorTrap
Dim Category As CharacterCategoriesE
Dim Character As Long
Dim Continue As Boolean
Dim Index As Long
Dim IsUnicode As Boolean
Dim NextCharacter As Long
Dim NextNextCharacter As Long
Dim PreviousCategory As CharacterCategoriesE
Dim PreviousCharacter As Long
Dim Text As String

   Character = Empty
   Continue = False
   IsUnicode = False
   Text = vbNullString
   
   ReDim Preserve BinaryData(LBound(BinaryData()) To UBound(BinaryData()) + 3) As Byte
   
   For Index = LBound(BinaryData()) To UBound(BinaryData()) - 2
      PreviousCharacter = Character
      Character = BinaryData(Index)
      NextCharacter = BinaryData(Index + 1)
      NextNextCharacter = BinaryData(Index + 2)
   
      PreviousCategory = Category
      Category = CharacterCategory(Character, PreviousCharacter, NextCharacter, NextNextCharacter)
      
      If Category = CCUnicodeNullCharacter Then
         If PreviousCategory = CCAdditionalCharacter Or PreviousCategory = CCRangeCharacter Then
            IsUnicode = True
         Else
            Category = CCUnreadable
         End If
      ElseIf Not Category = CCUnicodeNullCharacter Then
         If Category = CCAdditionalCharacter Or Category = CCRangeCharacter Then
            Text = Text & ChrW$(Character)
         ElseIf Category = CCExcludedCharacter Or Category = CCUnreadable Then
            If Not Text = vbNullString Then
               
               If TextDefinition.UnicodeOption = UOExclusiveUnicode Then
                  If IsUnicode Then
                     RaiseEvent FoundText(Text, Continue)
                     If Not Continue Then Exit Sub
                  End If
               Else
                  RaiseEvent FoundText(Text, Continue)
                  If Not Continue Then Exit Sub
               End If
               
               IsUnicode = False
               Text = vbNullString
            End If
         End If
      End If
   Next Index
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   RaiseEvent HandleError(Err)
   Resume EndRoutine
End Sub

'This function returns the binary data from the specified file.
Public Function GetBinaryData(BinaryFile As String) As Byte()
On Error GoTo ErrorTrap
   Dim BinaryData() As Byte
   Dim FileH As Integer
   
   ReDim BinaryData(0 To 0) As Byte
   
   If Not BinaryFile = vbNullString Then
      FileH = FreeFile()
      Open BinaryFile For Input Lock Read Write As FileH: Close FileH
      Open BinaryFile For Binary Lock Read Write As FileH
         BinaryData() = InputB$(LOF(FileH), FileH)
      Close FileH
   End If
   
EndRoutine:
   GetBinaryData = BinaryData()
   Exit Function
   
ErrorTrap:
   RaiseEvent HandleError(Err)
   Close FileH
   Resume EndRoutine
End Function




'This procedure initializes this class.
Private Sub Class_Initialize()
On Error GoTo ErrorTrap
   With TextDefinition
      .Additional = vbCr & vbTab
      .Excluded = vbNullString
      .RangeEnd = Asc("~")
      .RangeStart = Asc(" ")
      .UnicodeOption = UOExcludeUnicode
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   RaiseEvent HandleError(Err)
   Resume
End Sub


