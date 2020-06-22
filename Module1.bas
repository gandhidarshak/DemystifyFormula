Attribute VB_Name = "Module1"
'*****************************************************************************
' Summary: Macros to better handle In cell excel fromulae using name range definition
' ****************************************************************************

' Externally visible sub-routine for Variable -> Cell Ref
' Don't change this interface if possible as people may have made shortcuts around this
Sub ConvertVariableToCellRef()
   InternalConvertVariableToCellRef (True)
End Sub

'*****************************************************************************

' Internal  sub-routine for Variable -> Cell Ref
Function InternalConvertVariableToCellRef(Optional ByVal showMsgs As Boolean = True)
   Dim Review As VbMsgBoxResult

   ' Does user want to review changes?
   If showMsgs = True Then
      Review = MsgBox("Do you want to review each change before applying?", vbYesNo)
   End If

   ' Get user selected range
   Dim selRange As Range
   Set selRange = Application.Selection

   On Error Resume Next

   ' Iterate over all cells and do the conversion
   Dim selCel As Range
   For Each selCel In selRange.Cells
      Dim strFormula As String
      Dim OrgFormula As String
      OrgFormula = selCel.Formula
      strFormula = OneCellVariableToCellRef(selCel)

      Dim Msg1, Msg2, Msg3 As String
      Msg1 = "Do you accept changes for cell: " & selCel.Address & " ?" & Chr(13) & Chr(13)
      Msg2 = "Original Formula " & Chr(13) & OrgFormula & Chr(13) & Chr(13)
      Msg3 = "Modified Formula " & Chr(13) & strFormula

      If showMsgs = True And Review = vbYes Then
         Accept = MsgBox(Msg1 & Msg2 & Msg3, vbYesNo)
      Else
         Accept = vbYes
      End If

      If Accept = vbYes Then
         selCel.Formula = strFormula
      End If
   Next selCel

   If showMsgs = True Then
      MsgBox ("Name range to Cell reference conversion completed")
   End If
End Function

'*****************************************************************************

' Atom internal API for Var->Cell Ref conversion. Operates on one cell
Function OneCellVariableToCellRef(ByVal selCel As Range) As String
   On Error Resume Next
   Dim strFormula As String
   Dim OrgFormula As String
   strFormula = selCel.Formula
   OrgFormula = strFormula

   ' Add Space at the end so it can be treated like any other part of formula
   strFormula = strFormula & " "

   ' Clean up Formula name to remove current sheet's name.
   Dim shtName As String
   shtName = ActiveSheet.Name & "!"
   strFormula = Replace(strFormula, shtName, "")

   ' Separate all cell references by putting a space around all operators
   Dim opCharArray As Variant
   opCharArray = Array("+", "-", "*", "/", "(", ")", ",", ">", "<", "=")
   Dim opChar As Variant
   For Each opChar In opCharArray
      strFormula = Replace(strFormula, CStr(opChar), " " & CStr(opChar) & " ")
   Next opChar

   ' Replace all " 1 * "
   strFormula = Replace(strFormula, " 1 * ", " ")

   Dim delimArray() As String
   Dim i As Integer
   delimArray = Split(strFormula, " ")
   For i = LBound(delimArray) To UBound(delimArray)
      If delimArray(i) Like "*_Col_*" Then
         ' MsgBox (delimArray(i))
         Dim nameCol As Range
         Set nameRng = ActiveSheet.Names(delimArray(i)).RefersToRange
         If nameRng.Count > 100 Then
            ' Def a column
            strFormula = Replace(strFormula, delimArray(i), Cells(selCel.Row, nameRng.Column).Address(RowAbsolute:=False))
         End If
      End If
   Next i

   strFormula = Replace(strFormula, shtName, "")
   strFormula = Replace(strFormula, " ", "")

   ' To return string
   OneCellVariableToCellRef = strFormula

End Function

'*****************************************************************************

' Externally visible API for Cell Ref -> Variable conversion
Sub ConvertCellRefToVariable()
   InternalConvertCellRefToVariable (True)
End Sub

'*****************************************************************************

' Internal API for Cell Ref -> Variable conversion
Function InternalConvertCellRefToVariable(Optional ByVal showMsgs As Boolean = True)
   Dim Review As VbMsgBoxResult

   If showMsgs = True Then
      Review = MsgBox("Do you want to review each change before applying?", vbYesNo)
   End If


   ' Counter for exiting if it goes too long
   Dim counter As Integer
   counter = 0

   ' Selected range
   Dim selCel As Range
   Dim selRange As Range
   Set selRange = Application.Selection

   On Error Resume Next

   For Each selCel In selRange.Cells
      Dim strFormula As String
      Dim OrgFormula As String
      OrgFormula = selCel.Formula
      strFormula = OneCellRefToVariable(selCel)

      Dim Msg1, Msg2, Msg3 As String
      Msg1 = "Do you accept changes for cell: " & selCel.Address & " ?" & Chr(13) & Chr(13)
      Msg2 = "Original Formula " & Chr(13) & OrgFormula & Chr(13) & Chr(13)
      Msg3 = "Modified Formula " & Chr(13) & strFormula

      If showMsgs = True And Review = vbYes Then
         Accept = MsgBox(Msg1 & Msg2 & Msg3, vbYesNo)
      Else
         Accept = vbYes
      End If

      If Accept = vbYes Then
         selCel.Formula = strFormula
      End If
   Next selCel
   If showMsgs = True Then
      MsgBox ("Cell reference to Name Range conversion completed")
   End If
End Function

'*****************************************************************************

Function OneCellRefToVariable(ByVal selCel As Range) As String
   On Error Resume Next
   Dim strFormula As String
   Dim OrgFormula As String
   OrgFormula = selCel.Formula

   ' If the selction already has some column name refernce then we have to convert it to
   ' Cell reference first before we can convert all cell ref to formula
   strFormula = OneCellVariableToCellRef(selCel)

   strFormula = strFormula & " "

   ' Clean up Formula name to remove current sheet's name. And add spaces before, after all operators
   Dim shtName As String
   shtName = ActiveSheet.Name & "!"

   strFormula = Replace(strFormula, shtName, "")

   Dim matchOpArray As Variant
   mathOpArray = Array("SUM", "MIN", "MAX")
   Dim mathOpChar As Variant

   Dim opCharArray As Variant
   opCharArray = Array("+", "-", "*", "/", "(", ")", ",", ">", "<", "=")
   Dim opChar As Variant
   For Each opChar In opCharArray
      strFormula = Replace(strFormula, CStr(opChar), " " & CStr(opChar) & " ")
   Next opChar

   ' This cell will be wiped at the end of implementation. Change the cell if
   ' ZZ1 has some important data
   Dim tempCell As Range
   Set tempCell = Range("ZZ1")

   ' Covert range with : to ,
   Set objRegEx = CreateObject("VBScript.RegExp")
   objRegEx.Global = True
   objRegEx.Pattern = "([\S]+):([\S]+)"
   Set f_Matches = objRegEx.Execute(strFormula)
   For Each f_match In f_Matches
      ' Check if this : is only of this row?
      Dim isOfThisLine As Boolean
      isOfThisLine = True
      For Each f_SubMatch In f_match.Submatches
         If Not selCel.Row = Range(f_SubMatch).Row Then
            isOfThisLine = False
         End If
      Next f_SubMatch

      If isOfThisLine = True Then
         Dim f_MatchReplacement As String
         f_MatchReplacement = " "
         ' MsgBox f_Match
         tempCell.Formula = "=" & f_match
         For Each tempPrecCell In tempCell.DirectPrecedents.Cells
            f_MatchReplacement = f_MatchReplacement & tempPrecCell.Address(RowAbsolute:=False) & " , "
         Next tempPrecCell
         tempCell.Formula = ""
         ' Remove last ", " by reversing and removing first! Stupid VBA!!
         f_MatchReplacement = StrReverse(Replace(StrReverse(f_MatchReplacement), " ,", "", 1, 1))
         ' MsgBox (f_MatchReplacemnet)
         strFormula = Replace(strFormula, f_match, f_MatchReplacement)
         ' MsgBox (strFormula)
      End If
   Next f_match


   Dim precCel As Range
   Dim precRange As Range
   Set precRange = selCel.DirectPrecedents
   For Each precCel In precRange.Cells
      counter = counter + 1
      If counter > 200 Then
         MsgBox ("Number of iterations exceeded 200! Exiting function.")
         Exit Function
      End If

      ' Detect for math formula which can confuse Column name with entire column instead of one cell
      Dim isMathFormula As Boolean
      isMathFormula = False

      For Each mathOpChar In mathOpArray
         If InStr(1, strFormula, CStr(mathOpChar)) <> 0 Then
            isMathFormula = True
         End If
      Next mathOpChar

      Dim strPrefix As String
      strPrefix = " "
      If isMathFormula = True Then
         strPrefix = " 1* "
      End If


      ' MsgBox (precCel.Address())
      strFormula = Replace(strFormula, " " & precCel.Address(RowAbsolute:=False) & " ", strPrefix & precCel.Name.Name & " ")
      strFormula = Replace(strFormula, " " & precCel.Address(ColumnAbsolute:=False) & " ", strPrefix & precCel.Name.Name & " ")
      strFormula = Replace(strFormula, " " & precCel.Address(RowAbsolute:=False, ColumnAbsolute:=False) & " ", strPrefix & precCel.Name.Name & " ")
      strFormula = Replace(strFormula, " " & precCel.Address() & " ", strPrefix & precCel.Name.Name & " ")
      If selCel.Row = precCel.Row Then
         ' Same row, let's see if column name is there for precedent
         Dim colOfPrecRange As Range
         Set colOfPrecRange = Columns(precCel.Column)


         If isMathFormula = True Then
            preCelColName = "1*" & preCelColName
         End If

         ' MsgBox (colOfPrecRange.Name.Name)
         strFormula = Replace(strFormula, " " & precCel.Address(RowAbsolute:=False) & " ", strPrefix & colOfPrecRange.Name.Name & " ")
         strFormula = Replace(strFormula, " " & precCel.Address(ColumnAbsolute:=False) & " ", strPrefix & colOfPrecRange.Name.Name & " ")
         strFormula = Replace(strFormula, " " & precCel.Address(RowAbsolute:=False, ColumnAbsolute:=False) & " ", strPrefix & colOfPrecRange.Name.Name & " ")
         strFormula = Replace(strFormula, " " & precCel.Address() & " ", strPrefix & colOfPrecRange.Name.Name & " ")

         'MsgBox (strFormula)
      End If
   Next precCel

   ' handling off-sheet nameranges
   ' directprecedents only gives precedents from current sheet
   objRegEx.Pattern = "([\S]+)!([\S]+)"
   Set f_Matches = objRegEx.Execute(strFormula)
   Dim OtherSheetName As String
   Dim OtherSheetCell As String
   
   Dim toWrkSheet, toAddr, toName As String
   For Each f_match In f_Matches
       toWrkSheet = f_match.Submatches.Item(0)
       toAddr = f_match.Submatches.Item(1)
       Set preCel = Worksheets(toWrkSheet).Range(toAddr)
       strFormula = Replace(strFormula, " " & f_match & " ", " " & preCel.Name.Name & " ")
   Next f_match
   
   ' Replace sheet name
   strFormula = Replace(strFormula, shtName, "")

   'Remove extra " 1* " which are not part of math functions
   Dim delimArray() As String
   Dim i As Integer
   Dim isInAMathFunc As Boolean
   isInAMathFunc = False

   Dim bracketCount As Integer
   bracketCount = 0

   delimArray = Split(strFormula, " ")
   For i = LBound(delimArray) To UBound(delimArray)
      ' If this is a math func term, start the watch
      For Each mathOpChar In mathOpArray
         If delimArray(i) = CStr(mathOpChar) Then  ' Math function started
            ' MsgBox (delimArray(i))
            isInAMathFunc = True
         End If
      Next mathOpChar

      ' If we are inside a math func, count brackets
      If delimArray(i) = "(" And isInAMathFunc Then
         bracketCount = bracketCount + 1
      End If
      If delimArray(i) = ")" And isInAMathFunc Then
         bracketCount = bracketCount - 1
         If bracketCount = 0 Then
            isInAMathFunc = False
         End If
      End If

      ' If Not isInAMathFunc and we find 1* then remove it.
      If delimArray(i) = "1*" And Not isInAMathFunc Then
         delimArray(i) = " "
      End If
   Next i

   strFormula = ""

   ' Reconstruct Formula
   For i = LBound(delimArray) To UBound(delimArray)
      strFormula = strFormula & " " & delimArray(i)
   Next i

   strFormula = Replace(strFormula, " ", "")

   OneCellRefToVariable = strFormula
End Function


'*****************************************************************************

Sub NameThisColumn()
   Dim counter As Integer
   counter = 0
   Dim selCel As Range
   Dim selRange As Range
   Set selRange = Application.Selection
   Dim shtName As String
   shtName = ActiveSheet.Name & "!"
   On Error Resume Next

   For Each selCel In selRange.Cells
      Dim cellVal As String
      cellVal = selCel.Value
      If cellVal <> "" Then
         Dim colRange As Range
         Set colRange = Columns(selCel.Column)
         cellVal = Replace(cellVal, " ", "_")
         cellVal = Replace(cellVal, ")", "_")
         cellVal = Replace(cellVal, "(", "_")
         cellVal = Replace(cellVal, "__", "_")
         cellVal = Replace(cellVal, "+", "_")
         cellVal = Replace(cellVal, "?", "_")

         cellVal = cellVal & "_Col_" & ConvertToLetter(selCel.Column)
         ' MsgBox (cellVal)
         ActiveSheet.Names.Add Name:=cellVal, RefersTo:=colRange
         counter = counter + 1
      End If
   Next selCel
   MsgBox ("Created Column Names for " & counter & " columns")
End Sub

'*****************************************************************************
'
Function ConvertToLetter(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int((iCol - 1) / 26)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function



