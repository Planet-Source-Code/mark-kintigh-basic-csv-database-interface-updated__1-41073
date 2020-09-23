VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSVDBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 1

Private Type OPENFILENAME
   lStructSize As Long
   hWndOwner As Long
   hInstance As Long
   lpstrFilter As String
   lpstrCustomFilter As String
   nMaxCustFilter As Long
   nFilterIndex As Long
   lpstrFile As String
   nMaxFile As Long
   lpstrFileTitle As String
   nMaxFileTitle As Long
   lpstrInitialDir As String
   lpstrTitle As String
   flags As Long
   nFileOffset As Integer
   nFileExtension As Integer
   lpstrDefExt As String
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type

Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
'
' force no long names for 4.x modules
'
Private Const OFN_NOLONGNAMES = &H40000
'
' new look commdlg
'
Private Const OFN_EXPLORER = &H80000
Private Const OFN_NODEREFERENCELINKS = &H100000
'
' force long names for 3.x modules
'
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
   Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
   Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'
' Private TYPE defined
'
Private Type ValueInfo
   info() As Variant
End Type
'
' Local variables that are only used within the class
'
Private Fields() As String
Private Values() As ValueInfo
Private FieldCount As Long
'
' Local variables to hold property values
'
Private mvarRecordCount As Long
Private mvarCurrentRecord As Long
Private mvarAutoGenerateKey As Boolean
Private mvarChangesMade As Boolean
Private mvarKeyField As String
Private mvarFileName As String
Private mvarHeaderLine As Boolean
Private mvarSeperator As String
'
' Other Private Const values
'
Private Const UniqueKeyLength = 12
'
'
'
Enum SortOrder
   AtoZ = 1
   ZtoA = 2
End Enum

'------------------------------------------------------------------------
' Read-only properties
'------------------------------------------------------------------------
Public Property Get ChangesMade() As Boolean
    ChangesMade = mvarChangesMade
End Property

Public Property Get CurrentRecord() As Long
    CurrentRecord = mvarCurrentRecord
End Property

Public Property Get RecordCount() As Long
    RecordCount = mvarRecordCount
End Property

'------------------------------------------------------------------------
' Read/Write properties
'------------------------------------------------------------------------
Public Property Let Seperator(ByVal vData As String)
   mvarSeperator = vData
End Property

Public Property Get Seperator() As String
    HeaderLine = mvarSeperator
End Property

Public Property Let HeaderLine(ByVal vData As Boolean)
    mvarHeaderLine = vData
End Property

Public Property Get HeaderLine() As Boolean
    HeaderLine = mvarHeaderLine
End Property

Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Property Let KeyField(ByVal vData As String)
    mvarKeyField = vData
End Property

Public Property Get KeyField() As String
    KeyField = UCase(Trim(mvarKeyField))
End Property

Public Property Let AutoGenerateKey(ByVal vData As Boolean)
    mvarAutoGenerateKey = vData
End Property

Public Property Get AutoGenerateKey() As Boolean
    AutoGenerateKey = mvarAutoGenerateKey
End Property

'------------------------------------------------------------------------
' Class initializer
'------------------------------------------------------------------------
Private Sub Class_Initialize()
   mvarRecordCount = -1
   mvarCurrentRecord = -1
   mvarAutoGenerateKey = False
   mvarChangesMade = False
   mvarKeyField = ""
   mvarFileName = ""
   mvarHeaderLine = True
   mvarSeperator = ","
End Sub

'------------------------------------------------------------------------
' Public methods
'------------------------------------------------------------------------
Public Function AskForFile() As Boolean
   Dim s As String
   
   AskForFile = False                        ' Assume no file was choosen
   s = Trim(OpenDialog("CSV Files|*.CSV|All Files|*.*", _
      "Select the source file", App.Path))   ' As for a file name
   If Len(s) > 1 Then                        ' If a file was fiven...
      mvarFileName = s                       '> Store the file name
      s = ""
      Do                                     '> DO...LOOP WHILE
         s = Trim(InputBox("What is used " & _
            "to seperate the values?", _
            "Field seperator", _
            mvarSeperator))                  '>> Ask for seperator
         If Len(s) = 0 Then                  '>> If nothing entered...
            If MsgBox("Do you want to cancel?", _
               vbYesNo + vbQuestion, _
               "No seperator value given") = _
               vbYes Then                    '>>> ...ask to abort
               mvarFileName = ""             '>>>> Remove file name
               Exit Function                 '>>>> Exit function
            End If
         Else
            mvarSeperator = s                '>> Store seperator
         End If
      Loop While Len(s) < 1                  '> Loop while no seperator
      If MsgBox("Does the first line of the " & _
         "file contain header information/" & _
         "field names?", vbYesNo + vbQuestion, _
         "Header line?") = _
         vbYes Then                          '> Ask if there is a header
         mvarHeaderLine = True               '>> Set flag to TRUE
      Else                                   '> ...otherwise...
         mvarHeaderLine = False              '>> Set flag to FALSE
      End If
      AskForFile = True                      '> Set function flag to TRUE
   End If
End Function

Public Function FileExists(Optional ByVal FileName As String = "") _
   As Boolean
   Dim fs, s As String
   
   If Len(Trim(FileName)) > 0 Then           ' Check to see if given name
      s = FileName                           '> Use given name
   Else                                      ' ...otherwise...
      s = mvarFileName                       '> use internal name
   End If
   If Len(s) < 1 Then                        ' If it is still empty...
      FileExists = False                     '> file does not exist
      Exit Function                          '> Exit function
   End If
   '
   ' Use external object -- Scripting.FileSystemObject -- to check
   ' for the file's existance.
   '
   Set fs = CreateObject("Scripting.FileSystemObject")
   FileExists = fs.FileExists(s)
   Set fs = Nothing                          ' Release external object
End Function

Public Function Find(ByVal FieldValue As Variant, _
   Optional ByVal FieldName As String = "") As Long
   Dim x As Long, fi As Integer, found As Boolean
   
   found = False                             ' Set found to FALSE
   x = 1                                     ' Start counter at 1
   If Len(FieldName) > 1 Then                ' If field name given...
      fi = FieldIndex(FieldName)             '> look for field's index
   ElseIf Len(mvarKeyField) > 0 Then         ' If a key was defined...
      fi = FieldIndex(mvarKeyField)          '> look for field's index
   Else                                      ' otherwise...
      Find = -1                              '> No field was defined, so
      Exit Function                          '    exit the function
   End If
   If fi = -1 Then                           ' Could not find the field
      Find = -1                              '> Set Find to failure and
      Exit Function                          '    exit the function
   End If
   While Not found And x <= mvarRecordCount  ' While !found & x < array
      If UCase(Trim(FieldValue)) = _
         UCase(Trim(Values(x).info(fi))) Then '> Case insensitive compare
         found = True                        '>> Set found flag
         Find = x                            '>> Store index value
         mvarCurrentRecord = x               '>> Set internal index value
      Else                                   '> Not a match
         x = x + 1                           '>> Increment counter
      End If
      DoEvents                               '> Time slicing
   Wend
   If Not found Then Find = -1               ' If not found, set failure
End Function

Public Function NewRecord() As Boolean
   Dim fi As Integer

   NewRecord = True                          ' Assume success
   mvarChangesMade = True                    ' Set changes made flag
   If mvarCurrentRecord = -1 Then            ' Check for first record
      mvarCurrentRecord = 1                  '> Set index to 1
   Else                                      ' otherwise...
      mvarCurrentRecord = _
         UBound(Values, 1) + 1               '> Set index to +1
   End If
   ReDim Preserve Values(mvarCurrentRecord) _
      As ValueInfo                           ' Resize main array
   ReDim Values(mvarCurrentRecord).info(FieldCount) As Variant
   If mvarAutoGenerateKey Then               ' If class generates a key..
      If Len(mvarKeyField) < 1 Then          '> Error check: key defined
         fi = -1                             '>> set field index (FI) = -1
      Else
         fi = FieldIndex(mvarKeyField)       '>> Find the field's index
      End If
      If fi = -1 Then                        '> Check for failure
         NewRecord = False                   '>> Set to failure
         mvarCurrentRecord = _
            mvarCurrentRecord - 1            '>> Reset index (-1)
         If mvarCurrentRecord > 0 Then _
            ReDim Preserve Values( _
            mvarCurrentRecord, FieldCount) _
            As ValueInfo                     '>> Reverse array resizing
      Else                                   '> otherwise...
         Values(mvarCurrentRecord).info(fi) = _
            UniqueKey                        '>> Set the unique key
      End If
   End If
   mvarRecordCount = UBound(Values)          ' Reset record count
End Function

Public Function GotoRecord(ByVal RecordNumber As Long) As Boolean
Attribute GotoRecord.VB_Description = "Moves the index pointer to the given record.  If successful it returns TRUE, otherwise it returns FALSE"
   '
   ' Preform a range check then, if valid, change the internal index
   ' and return TRUE; otherwise, return FALSE
   '
   If RecordNumber >= 1 And RecordNumber <= mvarRecordCount Then
      mvarCurrentRecord = RecordNumber
      GotoRecord = True
   Else
      GotoRecord = False
   End If
End Function

Public Function WriteFile() As Boolean
Attribute WriteFile.VB_Description = "Will attempt to write the values back to the CSV file, overwriting the current values.  If successful it returns TRUE, otherwise it returns FALSE."
   Dim x As Long, fh As Integer, s As String

   WriteFile = True                          ' Assume success
   fh = FreeFile                             ' Grab a free file handle
   On Local Error GoTo WriteFile_Error       ' Set local error routine
   Open mvarFileName For Output Access Write _
      As fh                                  ' Open file in W/O mode
   If mvarHeaderLine Then                    ' If a header line is used
      s = Join(Fields, mvarSeperator)        '> Create the output line
      Print #fh, s                           '> Write the output line
   End If
   For x = 1 To mvarRecordCount              ' Process all of the records
      s = Join(Values(x).info, mvarSeperator) '> Create the output line
      Print #fh, s                           '> Write the output line
      DoEvents                               '> Time slicing
   Next
   Close fh                                  ' Close the file
   On Local Error GoTo 0                     ' Unset local error routine
   If WriteFile Then _
      mvarChangesMade = False                ' Unset changes made flag
   Exit Function                             ' Exit the function
WriteFile_Error:
   WriteFile = False                         ' Set return to FALSE
   Close fh                                  ' Close the file
   Err.Clear                                 ' Clear the error
   On Local Error GoTo 0                     ' Unset local error routine
   Exit Function                             ' Exit the function
End Function

Public Function LoadFile() As Boolean
Attribute LoadFile.VB_Description = "Will attempt to load the CSV file's information into memory.  If successful it returns TRUE, otherwise it returns FALSE."
   Dim x As Long, y As Integer, fh As Integer, s As String, a
   
   fh = FreeFile                             ' Grab a free file handle
   LoadFile = True                           ' Assume success
   On Local Error GoTo LoadFile_Error        ' Set local error routine
   Open mvarFileName For Input Access Read _
      As fh                                  ' Open file in R/O mode
   FieldCount = 0
   mvarCurrentRecord = _
      IIf(mvarHeaderLine, -1, 0)             ' Start the index counter
   While Not EOF(fh)                         ' While not at EOF
      Line Input #fh, s                      '> Read one line from file
      a = Split(s, mvarSeperator)            '> Split the field entries
      If mvarCurrentRecord = -1 Then         '> If it has a header...
         FieldCount = UBound(a) + 1          '>> Store # of read fields
         ReDim Fields(FieldCount) As String  '>> Resize array
         For y = 1 To FieldCount             '>> Save in the Fields()
            Fields(y) = Trim(a(y - 1))       '>>> Trim value then store
         Next
         mvarCurrentRecord = 0               '>> Set index to values
      ElseIf mvarCurrentRecord = 0 Then      '> else if on 1st record...
         If FieldCount < 1 Then              '>> Check if fields defined
            FieldCount = UBound(a) + 1       '>>> Store # of read fields
            ReDim Fields(FieldCount) As String '>>> Resize array
            For y = 1 To FieldCount          '>>> Create field names
               Fields(y) = "Field " & y      '>>>> Make "Field #"
            Next
         End If
         mvarCurrentRecord = 1               '>> Set counter to 1
         ReDim Values(1) As ValueInfo        '>> Resize array
         ReDim Values(1).info(FieldCount) _
            As Variant                       '>> Resize information array
         For y = 1 To FieldCount             '>> Save in Values().info()
            Values(1).info(y) = Trim(a(y - 1)) '>>> Trim value then store
         Next
      Else                                   '> otherwise, just add cell
         mvarCurrentRecord = _
            mvarCurrentRecord + 1            '>> Increase counter by 1
         ReDim Preserve Values( _
            mvarCurrentRecord) As ValueInfo  '>> Resize array
         ReDim Values( _
            mvarCurrentRecord).info( _
            FieldCount) As Variant           '>> Resize information array
         For y = 1 To FieldCount             '>> Save in Values().info()
            Values(mvarCurrentRecord).info(y) = _
               Trim(a(y - 1))               '>>> Trim value then store
         Next
      End If
      DoEvents                               '> Time slicing
   Wend
   mvarRecordCount = UBound(Values, 1)       ' Set the record counter
   If mvarRecordCount > 0 Then               ' If there are records...
      mvarCurrentRecord = 1                  '> Set counter = first entry
   Else                                      ' otherwise...
      mvarCurrentRecord = -1                 '> Set counter = -1
   End If
   Close fh                                  ' Close the file
   On Local Error GoTo 0                     ' Unset local error routine
   mvarChangesMade = False                   ' Set changes made flag
   Exit Function                             ' Exit the function
LoadFile_Error:                             ' -- LOCAL ERROR HANDLER --
   LoadFile = False                          ' Set return to FALSE
   Close fh                                  ' Close the file
   Err.Clear                                 ' Clear the error
   On Local Error GoTo 0                     ' Unset local error routine
   Exit Function                             ' Exit the function
End Function

Public Function DeleteRecord(Optional ByVal RecordNumber As Long = -1) _
   As Boolean
   Dim dIndex As Long, x As Long, s As String, y As Long
   
   DeleteRecord = False                      ' Assume failure
   If mvarRecordCount < 1 Then               ' Check for any records first
      Exit Function                          '> Exit the function
   End If
   dIndex = IIf(RecordNumber = -1, _
      mvarCurrentRecord, RecordNumber)       ' Determine entry to delete
   If dIndex < 1 Then                        ' Make sure a record selected
      Exit Function                          '> Exit the function
   End If
   While dIndex < mvarRecordCount            ' While not at the end
      DeleteRecord = True                    '> Set return to TRUE
      For y = 1 To FieldCount                '> Transfer all of the values
         Values(dIndex).info(y) = _
            Values(dIndex + 1).info(y)       '  from the next cell down 1
      Next
      dIndex = dIndex + 1                    '> Increment counter
      DoEvents                               '> Time slice
   Wend
   If DeleteRecord Then                      ' If something was deleted
      mvarRecordCount = mvarRecordCount - 1  '> -1 from total count
      ReDim Preserve Values(mvarRecordCount) _
         As ValueInfo                        '> Resize array
      mvarChangesMade = True                 '> Set changes made flag
   End If
End Function

Public Function BottomRecord() As Boolean
Attribute BottomRecord.VB_Description = "Moves the index pointer to the last record in memory.  If successful it returns TRUE, otherwise it returns FALSE."
   If mvarRecordCount < 1 Then               ' Error check the range
      BottomRecord = False
   Else
      BottomRecord = True
   End If
   mvarCurrentRecord = mvarRecordCount       ' Change index pointer
End Function

Public Function FirstRecord() As Boolean
Attribute FirstRecord.VB_Description = "Moves the index pointer to the first record in memory.  If successful it returns TRUE, otherwise it returns FALSE."
   If mvarRecordCount < 1 Then               ' Error check the range
      FirstRecord = False
   Else
      FirstRecord = True
   End If
   mvarCurrentRecord = 1                     ' Change index pointer
End Function

Public Function LastRecord() As Boolean
Attribute LastRecord.VB_Description = "Subtracts 1 from the pointer.  If successful it returns TRUE, otherwise FALSE"
   If mvarRecordCount < 1 Then               ' Error check the range
      LastRecord = False
   Else
      LastRecord = True
   End If
   If mvarCurrentRecord > 1 Then _
      mvarCurrentRecord = mvarCurrentRecord - 1 ' -1 from index pointer
End Function

Public Function NextRecord() As Boolean
Attribute NextRecord.VB_Description = "Will add 1 to the index counter.  If successful it returns TRUE, otherwise it returns FALSE."
   If mvarRecordCount < 1 Then               ' Error check the range
      NextRecord = False
   Else
      NextRecord = True
   End If
   If mvarCurrentRecord < mvarRecordCount Then _
      mvarCurrentRecord = mvarCurrentRecord + 1 ' +1 to index pointer
End Function

Public Function SetField(Optional ByVal FieldName As String = "", _
   Optional ByVal FieldValue As Variant = "", _
   Optional ByVal RecordNumber As Long = -1) As Boolean
   Dim x As Long, s As String, rn As Long
   
   s = IIf(Len(Trim(FieldName)) > 0, _
      Trim(FieldName), mvarKeyField)        ' Determine which field set
   rn = IIf(RecordNumber = -1, _
      mvarCurrentRecord, RecordNumber)       ' Determine record to change
   x = FieldIndex(s)                         ' Look up the field index
   If x = -1 Or rn < 1 Then                  ' Error check
      SetField = False                       '> ERROR: set return to FALSE
      Exit Function                          '> and exit the function
   End If
   SetField = True                           ' Set return to TRUE
   If Len(FieldValue) > 0 Then               ' Storing a normal value?
      Values(rn).info(x) = FieldValue        '> Yes, store the new value
   Else                                      ' No, asking to set key
      If mvarAutoGenerateKey Then            '> Auto generate the key?
         Values(rn).info(x) = UniqueKey      '>> Generate and set the key
      Else                                   '> No, use the given value
         If Len(Trim(FieldValue)) < 1 Then   '>> Error check: <>"" key
            SetField = False                 '>>> Change return to FALSE
         Else
            Values(rn).info(x) = FieldValue  '>>> Set the value
         End If
      End If
   End If
   If SetField Then _
      mvarChangesMade = True                 ' Set changes made flag
End Function

Public Function GetField(ByVal FieldName As String, _
   Optional ByVal RecordNumber As Long = -1) As Variant
   Dim x As Long, s As String, rn As Long
   
   s = IIf(Len(Trim(FieldName)) > 0, _
      Trim(FieldName), mvarKeyField)        ' Determine which field set
   rn = IIf(RecordNumber = -1, _
      mvarCurrentRecord, RecordNumber)       ' Determine record to change
   x = FieldIndex(s)                         ' Look up the field index
   If x = -1 Or rn < 1 Then                  ' Error check
      GetField = False                       '> ERROR: set return to FALSE
      Exit Function                          '> and exit the function
   End If
   GetField = Values(rn).info(x)             ' Set return value
End Function

Public Function DefineFields(Elements() As String) As Boolean
Attribute DefineFields.VB_Description = "Defines the list of fields used within the records, thereby allowing the user to call the field by name rather than array position.  If successful it returns TRUE, otherwise it returns FALSE."
   Dim x As Long, l As Long

   If UBound(Elements) < 1 Then              ' Make sure array > 0
      DefineFields = False                   '> Set return to FALSE
      Exit Function                          '> Exit the function
   End If
   l = UBound(Elements) - LBound(Elements) + 1 ' Determine cell count
   FieldCount = l                            ' Store count of fields
   ReDim Fields(FieldCount) As String        ' Resize interal FIELDS array
   x = 1
   For l = LBound(Elements) To _
      UBound(Elements)                       ' Process all of the cells
      Fields(x) = UCase(Trim(Elements(l)))   '> Store trimmed, upper case
      x = x + 1                              '> Increase interal counter
   Next
   For l = 1 To mvarRecordCount              ' Resize value arrays
      With Values(l)
         ReDim Preserve .info(FieldCount) _
            As Variant
      End With
   Next
   DefineFields = True                       ' Set return to TRUE
End Function

Public Function SortData( _
   Optional ByVal Direction As SortOrder = AtoZ, _
   Optional ByVal FieldName As String = "") As Boolean
   Dim s As String, fi As Long, x As Long, y As Long
   
   s = IIf(Len(Trim(FieldName)) > 0, _
      Trim(FieldName), mvarKeyField)         ' Determine which field set
   fi = FieldIndex(s)                        ' Get field's index
   If fi < 1 Then                            ' Error check
      SortData = False                       '> Set return to FALSE
      Exit Function                          '> Exit the function
   End If
   '
   ' Execute a bubble sort of all of the values
   '
   For x = 1 To mvarRecordCount - 1          ' For each record - 1
      For y = x + 1 To mvarRecordCount       '> For current+1 to last cell
         If Direction = AtoZ Then            '>> Sort A-to-Z?
            '
            ' If the current cell's value is
            ' greater than the target cell's
            ' value then swap the values.
            '
            If Values(x).info(fi) > _
               Values(y).info(fi) Then _
               SwapValues Values(x), Values(y)
         ElseIf Direction = ZtoA Then        '>> no, sort Z-to-A?
            '
            ' If the current cell's value is
            ' less than the target cell's
            ' value then swap the values.
            '
            If Values(x).info(fi) < _
               Values(y).info(fi) Then _
               SwapValues Values(x), Values(y)
         Else                                '>> no, ABORT (bad argument)
            SortData = False                 '>>> Set return to FALSE
            Exit Function                    '>>> Exit the function
         End If
         DoEvents                            '>> Time slice
      Next y
   Next x
End Function

Public Function GetFieldList()
   GetFieldList = Fields                     ' Return the field list
End Function

Public Function ChangeFieldName(ByVal OldKey As String, _
   ByVal NewKey As String) As Boolean
   Dim fi As Long, x As Long, y As Long
   
   fi = FieldIndex(OldKey)                   ' Get field's index
   If fi < 1 Then                            ' Error check
      ChangeFieldName = False                '> Set return to FALSE
      Exit Function                          '> Exit the function
   End If
   Fields(fi) = Trim(NewKey)                 ' Rename the field reference
   ChangeFieldName = True                    ' Set return to TRUE
End Function

Public Function AddField(ByVal NewKey As String, _
   Optional DefaultValue As Variant = "") As Boolean
   Dim i As Long
   
   If FieldIndex(NewKey) <> -1 Then          ' If the field exists...
      AddField = False                       '> Set return to FALSE
      Exit Function                          '> Exit the function
   End If
   FieldCount = FieldCount + 1               ' Increase field count
   ReDim Preserve Fields(FieldCount) _
      As String                              ' Resize the field array
   Fields(FieldCount) = UCase(Trim(NewKey))  ' Store the field name
   For i = 1 To mvarRecordCount              ' For every record...
      With Values(i)                         '> With each record
         ReDim Preserve .info(FieldCount) _
            As Variant                       '>> Resize information array
         .info(FieldCount) = DefaultValue    '>> Set default value
      End With
      DoEvents                               '> Time slicing
   Next
   AddField = True                           ' Set return to TRUE
   mvarChangesMade = True                    ' Set changes flag to TRUE
End Function

Public Function DeleteField(ByVal FieldName As String) As Boolean
   Dim i As Long, y As Long, fi As Integer

   fi = FieldIndex(FieldName)
   If fi = -1 Then                           ' If field doesn't exists...
      DeleteField = False                    '> Set return to FALSE
      Exit Function                          '> Exit the function
   End If
   For i = fi To FieldCount - 1              ' Remove the field reference
      Fields(i) = Fields(i + 1)              '> Copy down the reference
   Next
   FieldCount = FieldCount - 1               ' Decrease field count by 1
   ReDim Preserve Fields(FieldCount) _
      As String                              ' Resize Fields() array
   For i = 1 To mvarRecordCount              ' For every record...
      With Values(i)                         '> With each record
         For y = fi To FieldCount            '>> For each remaining field
            .info(y) = .info(y + 1)          '>>> Copy the value down 1
         Next
         ReDim Preserve .info(FieldCount) _
            As Variant                       '>> Resize the array
      End With
      DoEvents                               '> Time slicing
   Next
   DeleteField = True                        ' Set return to TRUE
   mvarChangesMade = True                    ' Set changes flag to TRUE
End Function

'------------------------------------------------------------------------
' Private methods
'------------------------------------------------------------------------
Private Function OpenDialog(Filter As String, _
   Title As String, InitDir As String) As String
   Dim ofn As OPENFILENAME
   Dim a As Long

   ofn.lStructSize = Len(ofn)
   ofn.hWndOwner = 0
   ofn.hInstance = App.hInstance
   If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
   For a = 1 To Len(Filter)
      If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
   Next
   ofn.lpstrFilter = Filter
   ofn.lpstrFile = Space$(254)
   ofn.nMaxFile = 255
   ofn.lpstrFileTitle = Space$(254)
   ofn.nMaxFileTitle = 255
   ofn.lpstrInitialDir = InitDir
   ofn.lpstrTitle = Title
   ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
   a = GetOpenFileName(ofn)
   If (a) Then
      OpenDialog = Trim$(ofn.lpstrFile)
   Else
      OpenDialog = ""
   End If
End Function

Private Function FieldIndex(FieldName As String) As Integer
   Dim x As Integer, found As Boolean
   
   found = False                             ' Set found to FALSE
   x = 1                                     ' Set counter to 1
   While Not found And x <= UBound(Fields)   ' While !found & x<array
      If UCase(Trim(FieldName)) = _
         UCase(Trim(Fields(x))) Then         '> Case insensitive compare
         found = True                        '>> Set found flag
         FieldIndex = x                      '>> Set return value
      Else                                   '> Not a match
         x = x + 1                           '>> Increase counter
      End If
      DoEvents                               '> Time slicing
   Wend
   If Not found Then FieldIndex = -1         ' If not found, set failure
End Function

Private Function UniqueKey() As String
   Dim x As Integer, s As String, y As Long, found As Boolean, kv As Long
   
   found = True                              ' Set flag to TRUE
   Randomize                                 ' Seed the RND() function
   kv = FieldIndex(mvarKeyField)             ' Get the key's index
   If kv < 1 Then                            ' Error check
      UniqueKey = ""                         '> Set return to blank
      Exit Function                          '> Exit the function
   End If
   While found                               ' While the flag is true...
      s = ""                                 '> Blank desired key
      For x = 1 To UniqueKeyLength           '> Build the a key
         Select Case Int(Rnd() * 2)          '> Randomly select char or #
            Case 0                           '>> Use a number
               s = s & Chr(Asc("0") + _
                  Int(Rnd() * (Asc("9") - _
                  Asc("0"))))
            Case 1                           '>> Use a character
               s = s & Chr(Asc("A") + _
                  Int(Rnd() * (Asc("Z") - _
                  Asc("A"))))
         End Select
      Next
      found = False                          '> Assume it won't be found
      For y = 1 To mvarRecordCount           '> Check records
         If Values(y).info(kv) = s Then      '>> Values the same?
            found = True                     '>>> Yes, set flag to TRUE
            Exit For                         '>>> Exit loop
         End If
         DoEvents                            '>> Time slicing
      Next
   Wend
   If Not found Then UniqueKey = s _
   Else UniqueKey = ""                       ' Set the return value
End Function

Private Sub SwapValues(ByRef v1 As ValueInfo, v2 As ValueInfo)
   Dim s As Variant, x As Long
   
   For x = 1 To FieldCount                   ' For each field...
      s = v2.info(x)                         '> Copy v2   into temp
      v2.info(x) = v1.info(x)                '> Copy v1   into v2
      v1.info(x) = s                         '> Copy temp into v1
   Next
End Sub
