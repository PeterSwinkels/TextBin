VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TextBinDemoBox 
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4995
   ClipControls    =   0   'False
   Icon            =   "TextBinDemo.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   13.188
   ScaleMode       =   4  'Character
   ScaleWidth      =   41.625
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TextBox 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
   Begin VB.Menu QuitMenu 
      Caption         =   "&Quit"
   End
   Begin VB.Menu RestartMenu 
      Caption         =   "&Restart"
   End
   Begin VB.Menu InformationMenu 
      Caption         =   "&Information"
   End
End
Attribute VB_Name = "TextBinDemoBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main interface.
Option Explicit

'The Microsoft Windows API function used by this program.
Private Declare Function SafeArrayGetDim Lib "Oleaut32.dll" (ByRef saArray() As Any) As Long

'This enumeration lists the data types supported by this demo.
Private Enum DataTypesE
   DTGuidIds = 0         'Anything that could be a valid GUID id.
   DTDLLReferences       'Anything that could be a valid DLL filename.
   DTEMailAddresses      'Anything that could be a valid e-mail address.
   DTNames               'Anything that could be a name (last, initials (first).)
   DTText                'Any string of "human readable" characters (character codes 31-127.)
   DTTextBlocks          'Any string of "human readable" characters (character codes 31-127) including line breaks.
   DTURLs                'Anything that could be a valid url (with the protocol specified.)
End Enum

'This enumeration lists the line break types supported by this demo.
Private Enum LineBreakTypesE
   LBMSDOS = 0           'The carriage return character line break type.
   LBwindows             'The carriage return character followed by a line feed character line break type.
   LBUnix                'The line feed character line break type.
End Enum

'This enumeration lists the relative text fragment positions checked for by this demo.
Private Enum RelativePositionsE
   RPNone = 0          'No position.
   RPStart = 1         'The start position.
   RPMiddle = 2        'The middle position.
   RPEnd = 4           'The end position.
End Enum

'This structure defines a search action.
Private Type SearchStr
   Aborted As Boolean                        'Indicates whether the search has been aborted by the user.
   AlreadyFound() As String                  'Defines the text fragments already found.
   CurrentDataType As DataTypesE             'Defines the data type selected by the user.
   CurrentLineBreakType As LineBreakTypesE   'Defines the line break type selected by the user.
   CurrentPath As String                     'Defines the path of the file to be searched.
   CurrentUnicodeOption As UnicodeOptionsE   'Defines the unicode option selected by the user.
End Type

Private Const GUID_MASK As String = "########-####-####-####-############"   'Defines a GUID's mask. The hash ("#") character stands for "hexadecimal digit".

Private Search As SearchStr                  'The parameters and status information of a search action.
Private WithEvents TextBin As TextBinClass   'Indicates that Text Bin class contains events to be used by this window.
Attribute TextBin.VB_VarHelpID = -1

'This procedure adds the text found to the list, if it is a unique text fragment.
Private Sub AddToList(Text As String)
On Error GoTo ErrorTrap
Dim Index As Long

   With Search
      If SafeArrayGetDim(.AlreadyFound()) = 0 Then
         ReDim .AlreadyFound(0 To 0) As String
      Else
         For Index = LBound(.AlreadyFound()) To UBound(.AlreadyFound())
            If Text = .AlreadyFound(Index) Then Exit Sub
         Next Index
      
         ReDim Preserve .AlreadyFound(LBound(.AlreadyFound()) To UBound(.AlreadyFound()) + 1) As String
      End If
      
      .AlreadyFound(UBound(.AlreadyFound())) = Text
   End With
   
   TextBox.Text = TextBox.Text & Text & vbCrLf
   
   DisplayStatus
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns whether the specified fragment occurs the specified number of times in the specified text. Negative values indicate any value greater than one.
Private Function CheckCount(Text As String, Fragment As String, ExpectedOccurrences As Long) As Boolean
On Error GoTo ErrorTrap
Dim Fragments() As String
Dim Occurrences As Long

   Fragments = Split(Text, Fragment)
   Occurrences = UBound(Fragments()) - LBound(Fragments())

EndRoutine:
   CheckCount = (ExpectedOccurrences < 0 And Occurrences > 0) Or (ExpectedOccurrences >= 0 And Occurrences = ExpectedOccurrences)
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function



'This procedure returns whether the specified fragment's positions match the expected positions specified. Negative position values check whether a fragment is not present.
Private Function CheckPositions(Text As String, Fragment As String, ExpectedPositions As RelativePositionsE) As Boolean
On Error GoTo ErrorTrap
Dim Positions As Long
   
   Positions = RPNone
   
   If Text = Fragment Then
      Positions = (RPStart Or RPMiddle Or RPEnd)
   Else
      If Left$(Text, Len(Fragment)) = Fragment Then Positions = Positions Or RPStart
      If Right$(Text, Len(Fragment)) = Fragment Then Positions = Positions Or RPEnd
      If InStrB(2, Left$(Text, Len(Text) - 1), Fragment) > 0 Then Positions = Positions Or RPMiddle
   End If
   
EndRoutine:
   If Sgn(ExpectedPositions) < 0 Then
      CheckPositions = Not (Positions = Abs(ExpectedPositions))
   Else
      CheckPositions = (Positions = ExpectedPositions)
   End If
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure displays the status information.
Private Sub DisplayStatus()
On Error GoTo ErrorTrap
   With Search
      Me.Caption = App.Title & " - "
      If SafeArrayGetDim(.AlreadyFound()) = 0 Then
         Me.Caption = Me.Caption & "0 results."
      Else
         Me.Caption = Me.Caption & CStr(UBound(.AlreadyFound()) - LBound(.AlreadyFound()))
         If UBound(.AlreadyFound()) - LBound(.AlreadyFound()) = 1 Then Me.Caption = Me.Caption & " result" Else Me.Caption = Me.Caption & " results"
      End If
   
      Me.Caption = Me.Caption & " found in " & .CurrentPath
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub



'This procedure handles any errors that occur.
Private Sub HandleError()
Dim ErrorCode As Long
Dim Message As String
   
   ErrorCode = Err.Number
   Message = Err.Description
   On Error Resume Next
   MsgBox Message, vbOKOnly Or vbExclamation
End Sub


'This procedure requests the user to specify a data type and file to be searched.
Private Sub RequestParameters()
On Error GoTo ErrorTrap
Dim DataTypePrompt As String
Dim LineBreakPrompt As String
Dim SelectedDataType As String
Dim SelectedLineBreakType As String
Dim SelectedUnicodeOption As String
Dim UnicodeOptionPrompt As String

   With Search
      If .CurrentPath = vbNullString Then .CurrentPath = ShowFileDialog()
   
      If Not .CurrentPath = vbNullString Then
         DataTypePrompt = "0. GUID Ids" & vbCr
         DataTypePrompt = DataTypePrompt & "1. DLL references" & vbCr
         DataTypePrompt = DataTypePrompt & "2. E-Mail addresses" & vbCr
         DataTypePrompt = DataTypePrompt & "3. Names (last, initials (first))" & vbCr
         DataTypePrompt = DataTypePrompt & "4. Text (strings of character codes 31-127)" & vbCr
         DataTypePrompt = DataTypePrompt & "5. Text blocks (strings of character codes 31-127)" & vbCr
         DataTypePrompt = DataTypePrompt & "6. URLs"
      
         SelectedDataType = InputBox$(DataTypePrompt, "Data Type:", CStr(.CurrentDataType))
      
         If Not SelectedDataType = vbNullString Then
            UnicodeOptionPrompt = "0 = Exclude unicode" & vbCr
            UnicodeOptionPrompt = UnicodeOptionPrompt & "1 = Include unicode" & vbCr
            UnicodeOptionPrompt = UnicodeOptionPrompt & "2 = Exclusively unicode"
            
            SelectedUnicodeOption = InputBox$(UnicodeOptionPrompt, "Data Type:", CStr(.CurrentUnicodeOption))
            If Not SelectedUnicodeOption = vbNullString Then .CurrentUnicodeOption = CLng(Val(SelectedUnicodeOption))
           
            If SelectedDataType = DTTextBlocks Then
               LineBreakPrompt = "0. MS-DOS (Carriage Return)" & vbCr
               LineBreakPrompt = LineBreakPrompt & "1. Windows (Carriage Return + Line Feed)" & vbCr
               LineBreakPrompt = LineBreakPrompt & "2. Linux (Line Feed)" & vbCr
               
               SelectedLineBreakType = InputBox$(LineBreakPrompt, "Line Break Type:", CStr(.CurrentLineBreakType))
               If Not SelectedLineBreakType = vbNullString Then .CurrentLineBreakType = CLng(Val(SelectedLineBreakType))
            End If
            
            SetDataType CLng(Val(SelectedDataType))
            StartSearch
         End If
      End If
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure sets the text definition used by the Text Bin class to fit the specified data type.
Private Sub SetDataType(NewDataType As DataTypesE)
On Error GoTo ErrorTrap

   With Search
      .CurrentDataType = NewDataType
   
      If Not (.CurrentUnicodeOption = UOExcludeUnicode Or .CurrentUnicodeOption = UOIncludeUnicode Or .CurrentUnicodeOption = UOExclusiveUnicode) Then
         .CurrentUnicodeOption = UOExcludeUnicode
         MsgBox "Invalid unicode option. Using default unicode option.", vbExclamation
      End If
   
      Select Case .CurrentDataType
         Case DTGuidIds
            TextBin.DefineText Asc("0"), Asc("9"), "ABCDEFabcdef-", , .CurrentUnicodeOption
         Case DTDLLReferences
            TextBin.DefineText Asc(" "), Asc("~"), , "\/:*?""<>|", .CurrentUnicodeOption
         Case DTEMailAddresses
            TextBin.DefineText Asc("!"), Asc("~"), , "()[]\;:,<>""", .CurrentUnicodeOption
         Case DTNames
            TextBin.DefineText Asc("A"), Asc("Z"), "abcdefghijklmnopqrstuvwxyz(,.) ", , .CurrentUnicodeOption
         Case DTText
            TextBin.DefineText Asc(" "), Asc("~"), vbTab, , .CurrentUnicodeOption
         Case DTTextBlocks
            TextBin.DefineText Asc(" "), Asc("~"), vbCrLf & vbTab, , .CurrentUnicodeOption
                  
            If Not (.CurrentLineBreakType = LBMSDOS Or .CurrentLineBreakType = LBMSDOS Or .CurrentLineBreakType = LBwindows) Then
               .CurrentLineBreakType = LBwindows
               MsgBox "Invalid line break type. Using default line break type.", vbExclamation
            End If
         Case DTURLs
            TextBin.DefineText Asc("!"), Asc("~"), , "<>""'", .CurrentUnicodeOption
         Case Else
            .CurrentDataType = DTText
            TextBin.DefineText Asc(" "), Asc("~"), vbCr & vbTab, vbNullString, .CurrentUnicodeOption
            MsgBox "Invalid data type. Using default data type.", vbExclamation
      End Select
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure displays the file dialog and returns the path of the file selected by the user.
Private Function ShowFileDialog() As String
On Error GoTo ErrorTrap

   With FileDialog
      .Flags = cdlOFNFileMustExist
      .Flags = .Flags Or cdlOFNLongNames
      .Flags = .Flags Or cdlOFNNoChangeDir
      .Flags = .Flags Or cdlOFNNoDereferenceLinks
      .Flags = .Flags Or cdlOFNNoReadOnlyReturn
      .Flags = .Flags Or cdlOFNPathMustExist
      .Flags = .Flags Or cdlOFNShareAware
   End With
   
   FileDialog.ShowOpen
EndRoutine:
   ShowFileDialog = FileDialog.FileName
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure starts a search action.
Private Sub StartSearch()
On Error GoTo ErrorTrap

   With Search
      If Not .CurrentPath = vbNullString Then
         Screen.MousePointer = vbHourglass
         .Aborted = False
         Erase .AlreadyFound()
   
         DisplayStatus
         TextBox.Text = vbNullString
   
         If Left$(.CurrentPath, 1) = """" Then .CurrentPath = Mid$(.CurrentPath, 2)
         If Right$(.CurrentPath, 1) = """" Then .CurrentPath = Left$(.CurrentPath, Len(.CurrentPath) - 1)
   
         TextBin.FindText TextBin.GetBinaryData(.CurrentPath)
      End If
   End With
   
   Screen.MousePointer = vbDefault
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to request the user to specify the search parameters.
Private Sub Form_Activate()
On Error GoTo ErrorTrap
   RequestParameters
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure processes any key strokes made by the user.
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap
   If KeyCode = vbKeyEscape Then Search.Aborted = True
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   Me.Caption = App.Title
   Me.Width = Screen.Width / 1.5
   Me.Height = Screen.Height / 1.5
   
   With Search
      .Aborted = False
      .CurrentDataType = DTText
      .CurrentLineBreakType = LBwindows
      .CurrentPath = Command$()
      .CurrentUnicodeOption = UOExcludeUnicode
   End With
   
   Set TextBin = New TextBinClass
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure adjusts this window's controls to its new size.
Private Sub Form_Resize()
On Error Resume Next
   TextBox.Width = Me.ScaleWidth
   TextBox.Height = Me.ScaleHeight
End Sub


'This procedure closes this program when this window is closed.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
Dim Choice As VbMsgBoxResult

   Choice = MsgBox("Quit?", vbQuestion Or vbYesNo Or vbDefaultButton2)
   
   Select Case Choice
      Case vbNo
         Cancel = CInt(True)
      Case vbYes
         Cancel = CInt(False)
         End
   End Select
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure displays information about this program.
Private Sub InformationMenu_Click()
On Error GoTo ErrorTrap
   With App
      MsgBox .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName, vbInformation
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to close this window.
Private Sub QuitMenu_Click()
On Error GoTo ErrorTrap
   Unload Me
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to request the user to specify the search parameters.
Private Sub RestartMenu_Click()
On Error GoTo ErrorTrap
   Search.CurrentPath = vbNullString
   RequestParameters
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure is called each time a string that fits the specified text definition is found.
Private Sub TextBin_FoundText(Text As String, ContinueSearch As Boolean)
On Error GoTo ErrorTrap
Dim Match As Boolean
Dim Position As Long

   Text = Trim$(Text)
   
   If Not Text = vbNullString Then
      Select Case Search.CurrentDataType
         Case DTGuidIds
            If Len(Text) >= Len(GUID_MASK) Then
               Match = True
               If InStr(Text, "-") > InStr(GUID_MASK, "-") Then Text = Mid$(Text, InStr(Text, "-") - InStr(GUID_MASK, "-"))
               Text = Left$(Text, Len(GUID_MASK))
               For Position = 1 To Len(GUID_MASK)
                  Select Case Mid$(GUID_MASK, Position, 1)
                     Case "#"
                        If Mid$(Text, Position, 1) = "-" Then
                           Match = False
                           Exit For
                        End If
                     Case "-"
                        If Not Mid$(Text, Position, 1) = "-" Then
                           Match = False
                           Exit For
                        End If
                  End Select
               Next Position
               If Match Then AddToList "{" & UCase$(Text) & "}"
            End If
         Case DTDLLReferences
            If CheckPositions(Text, ".", -RPStart) Then
               If CheckPositions(LCase$(Text), ".dll", RPEnd) Then AddToList LCase$(Text)
            End If
         Case DTEMailAddresses
            If CheckPositions(Text, "@", RPMiddle) Then
               If CheckPositions(Text, ".", RPMiddle) Then
                  If CheckPositions(Text, ".@", RPNone) Then
                     If CheckPositions(Text, "@.", RPNone) Then
                        If CheckPositions(Text, "..", RPNone) Then
                           If CheckCount(Text, "@", 1) Then AddToList LCase$(Text)
                        End If
                     End If
                  End If
               End If
            End If
         Case DTNames
            If CheckCount(Text, ",", 1) Then
               If CheckCount(Text, "(", 1) Then
                  If CheckCount(Text, ")", 1) Then
                     If CheckCount(Text, " ", -1) Then AddToList LCase$(Text$)
                  End If
               End If
            End If
         Case DTText
            AddToList Text
         Case DTTextBlocks
            Select Case Search.CurrentLineBreakType
               Case LBMSDOS
                  If CheckCount(Text, vbCr, -1) And CheckCount(Text, vbCrLf, 0) Then AddToList Replace(Text, vbCr, vbCrLf) & vbCrLf
               Case LBwindows
                  If CheckCount(Text, vbCrLf, -1) Then AddToList Text & vbCrLf
               Case LBUnix
                  If CheckCount(Text, vbLf, -1) And CheckCount(Text, vbCrLf, 0) Then AddToList Replace(Text, vbLf, vbCrLf) & vbCrLf
            End Select
         Case DTURLs
            If CheckPositions(Text, "://", RPMiddle) Then AddToList LCase$(Text)
      End Select
   End If
   
   DoEvents
   
   ContinueSearch = Not Search.Aborted
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure is called when an error occurs in the Text Bin class.
Private Sub TextBin_HandleError(ErrorO As Object)
   MsgBox ErrorO.Description & vbCr & "Error code: " & ErrorO.Number, vbExclamation
End Sub

'This procedure loads a file dropped into
Private Sub TextBox_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap
   If Data.Files.Count > 0 Then
      Search.CurrentPath = Data.Files.Item(1)
      RequestParameters
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


