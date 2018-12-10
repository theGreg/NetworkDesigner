Attribute VB_Name = "stringStuff"
Function ScanFence(lv94x() As Long, lv96x() As Long, LV_SP_P() As String, LV_DP_P() As String) As String
    Dim counter As Integer
    Dim oElement As Element
    Dim oEnumerator As ElementEnumerator
    Dim ofence As Fence
    Dim lv940s, lv960s, pole944, LVpoles As Integer
    Dim lv964 As Integer
    'Dim LVSPcount(1 To 10) As Integer
    LVpoles = 0
    lv940s = 0
    lv960s = 0
    pole944 = 0
    lv964 = 0
    'Dim ii As Integer
    'For ii = 1 To 10
    '    MsgBox LV_SP_P(ii)
    'Next ii
    'For ii = 1 To 10
    '    MsgBox LV_DP_P(ii)
    'Next ii
    Set ofence = ActiveDesignFile.Fence
    If ofence.IsDefined = True Then
        Set oEnumerator = ofence.GetContents
        MsgBox "defined"
        Dim Elements As String
            Elements = ""
        Do While oEnumerator.MoveNext
            Set oElement = oEnumerator.Current
            'If oElement.Type = 3 Then
            
                If oElement.type = msdElementTypeText Then
                    counter = counter + 1
                    Elements = Elements & counter & " : " & oElement.AsTextElement.Text & vbCrLf
                    'results = results + oElement.AsTextElement.Text & Chr(11) '& vbCrLf
                    Select Case (oElement.AsTextElement.Text)
                    
                        Case "091600960", "091600961", "091600962", "091600963", "091600964" & _
                                 "091600965", "091600966", "091600967", "091600968", "091600969"
                            lv960s = lv960s + 1
                        Case "091600964"
                            
                            lv964 = lv964 + 1
                            'lv96x = updateCount(Low:=1, High:=10, Prefix:=oElement.AsTextElement.Text, polePrefixArray:=LV_SP_P)
                        
                        Case LV_DP_P(1), LV_DP_P(2), LV_DP_P(3), LV_DP_P(4), LV_DP_P(5) & _
                                LV_DP_P(6), LV_DP_P(7), LV_DP_P(8), LV_DP_P(9), LV_DP_P(10)
                        
                            lv940s = lv940s + 1
                            '   lv94x = updateCount(Low:=1, High:=10, Prefix:=oElement.AsTextElement.Text, polePrefixArray:=LV_DP_P)
                        
                        'Case "091600944"
                        '  pole944 = pole944 + 1
                    
                    End Select
                    If oElement.AsTextElement.Text = "091600964" Then
                        lv960s = lv960s + 1
                    End If
                    
                    
                End If
           'elements = elements +
           'show results
           
           'results = "Pole 940 = " & CStr(pole940) & _
                    ' "Pole 940 = " & CStr(pole942) & _
                    ' "Pole 940 = " & CStr(lv960s)
        Loop
        MsgBox "The total number of elements scanned were: " & CStr(counter) & vbCrLf & _
                Elements

        
           MsgBox " Lv poles is 940s: " & lv940s & " , 960s: " & lv960s
    Else
        MsgBox "A Fence Has Not Been Defined!"
        
   End If
   
    ScanFence = CStr(counter)
End Function

Private Function displayResults() As String
    displayResults = ScanDGN()

End Function
Function ScanDGN() As String
    Dim counter As Integer
    Dim oEnumerator As ElementEnumerator
    Set oEnumerator = ActiveModelReference.Scan
    Do While oEnumerator.MoveNext
        counter = counter + 1
    Loop
    MsgBox "The total number of elements scanned were: " & CStr(counter)
    ScanDGN = CStr(counter)
End Function

Sub TestStrings()
    'Dim Arr() As String
    Dim LV_SP_poles() As String ' array of strings are the strings that will be looked for in the drawing
    Dim LV_DP_poles() As String ' these are the "0916009XX" and "1118017XX" series
    Dim MV_174X_poles() As String ' mv poles
    Dim MV_177X_poles() As String 'mv poles
    
    Dim LV_SP_Count() As Long
    
    Dim LV_SP_prefix As String
    Dim LV_DP_prefix As String
    Dim MV_174X_prefix As String
    Dim MV_177X_prefix As String
    
    LV_SP_prefix = "115"
    LV_DP_prefix = "114"
    MV_174X_prefix = "11180174"
    MV_177X_prefix = "11180177"
         
    'ReDim LV_SP_poles(1 To 5)
    'MsgBox "BEFORE TestStrings: Number Of Elements in LVS: " & CStr(UBound(LV_SP_poles) - LBound(LV_SP_poles) + 1)
    'Arr = LoadStrings(Low:=1, High:=10)
    'Dim arr() As String
    LV_SP_poles = InitializePrefixStrings(low:=1, high:=10, prefix:=LV_SP_prefix)
    LV_DP_poles = InitializePrefixStrings(low:=1, high:=10, prefix:=LV_DP_prefix)
    'MV_174X_poles = InitializePrefixStrings(low:=1, high:=10, prefix:=MV_174X_prefix)
    'MV_177X_poles = InitializePrefixStrings(low:=1, high:=10, prefix:=MV_177X_prefix)
    'MsgBox "AFTER Teststrings: Number Of Elements in LVS: " & CStr(UBound(LV_SP_poles) - LBound(LV_SP_poles) + 1)
    'Dim i As Integer
    'For i = 1 To UBound(LV_SP_poles)
        'MsgBox LV_SP_poles(i)
        'Debug.Print LV_SP_poles(i)
   ' Next i
    
End Sub
Function InitializePrefixStrings(low As Long, high As Long, prefix As String) As String()
    ''''''''''''''''''''''''''''''''''''''''''''''''', poletypes() As String
    ' LoadNumbers
    ' Returns an array of Longs containing the numbers
    ' between Low and High.
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Dim resultArray() As String
    'ReDim poletypes(1 To (High - Low + 1))
    'poletypes(1) = 8
    Dim Ndx As Long
    'Dim Val As Long
    If low > high Then
        MsgBox "Size of array cannot be negative"
        Exit Function
    End If
    ReDim resultArray(1 To (high - low + 1))
    'Val = Low
    For Ndx = LBound(resultArray) To UBound(resultArray)
        resultArray(Ndx) = prefix & CStr(Ndx - 1)
        'ResultArray(Ndx) = CStr(poletypes(1))
        'Val = Val + 1
    Next Ndx
    InitializePrefixStrings = resultArray()
End Function

Function InitializeNumbers(low As Long, high As Long) As Long()
    ''''''''''''''''''''''''''''''''''''''''''''''''', poletypes() As String
    ' LoadNumbers
    ' Returns an array of Longs containing the numbers
    ' between Low and High.
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Dim resultArray() As Long
    'ReDim poletypes(1 To (High - Low + 1))
    Dim Ndx As Long
    'Dim Val As Long
    If low > high Then
        MsgBox "Size of array cannot be negative"
        Exit Function
    End If
    ReDim resultArray(1 To (high - low + 1))
    'Val = Low
    For Ndx = LBound(resultArray) To UBound(resultArray)
        resultArray(Ndx) = 0
    Next Ndx
    InitializeNumbers = resultArray()
End Function


Function updateCount(low As Long, high As Long, prefix As String, polePrefixArray() As String) As Long()
    '''''''''''''''''''''''''''''''''''''''''''''''''
    ' LoadNumbers
    ' Returns an array of Longs containing the numbers
    ' between Low and High.
    '''''''''''''''''''''''''''''''''''''''''''''''''

    Dim resultArray() As Long
    Dim Ndx As Long
    'Dim Val As Long
    If low > high Then
        MsgBox "Size of array cannot be negative"
        Exit Function
    End If
    ReDim resultArray(1 To (high - low + 1))
    
    'Val = Low
    For Ndx = LBound(resultArray) To UBound(resultArray)
        If prefix = polePrefixArray(Ndx) Then
         resultArray(Ndx) = resultArray(Ndx) + 1
         'Exit For
        End If
    Next Ndx
        
    updateCount = resultArray()
End Function
