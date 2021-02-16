Attribute VB_Name = "Module1"
Sub selekcja()
    Dim input_fullpath As String
    Dim output_name As String
    Dim output_fullpath As String
    Dim input_fileNumber As Integer
    Dim output_fileNumber As Integer
    Dim linefromfile As String
    Dim Xmin As String
    Dim Ymin As String
    Dim Xmax As String
    Dim Ymax As String
    Dim Zmin As String
    Dim Zmax As String

    
    Xmin = Range("D13").Value
    Xmax = Range("D3").Value
    Ymin = Range("B8").Value
    Ymax = Range("E8").Value
    
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .AllowMultiSelect = False
        .Title = "Wybór pliku"
        .InitialFileName = "C:\Users\nk041169\Desktop\" ' zmieniæ
        .Filters.Add "Text files", "*.txt;*.csv; *.xyz", 1
        .Filters.Add "All files", "*.*"
        If .Show = -1 Then
            input_fullpath = .SelectedItems.Item(1)
            Else: GoTo error1
        End If
    End With
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Save a folder"
        .InitialFileName = "C:\Users\nk041169\Desktop\" ' zmieniæ
        If .Show = -1 Then
            output_fullpath = .SelectedItems(1)
        Else: GoTo error1
        End If
    End With
    
    output_name = InputBox("Podja nazwê pliku wyjœciowego")
    If output_name = "" Then
        output_name = "output_file"
        MsgBox "Nie wybrano nazwy pliku. Plik zostanie zapisany pod nazw¹ output_file"
    End If
    
    input_fileNumber = FreeFile
    Open input_fullpath For Input As #input_fileNumber
    
    output_fileNumber = FreeFile
    output_fullpath = output_fullpath & "\" & output_name & ".txt"
    Open output_fullpath For Output As #output_fileNumber
    
    If Range("D13") = Empty Then
        Xmin = min_max_from_file("X1", input_fileNumber)
    End If
    If Range("D3") = Empty Then
        Xmax = min_max_from_file("X2", input_fileNumber)
    End If
    If Range("B8") = Empty Then
        Ymin = min_max_from_file("Y1", input_fileNumber)
    End If
    If Range("E8") = Empty Then
        Ymax = min_max_from_file("Y2", input_fileNumber)
    End If
    If Range("C22") = Empty Then
        Zmin = min_max_from_file("Z1", input_fileNumber)
    End If
    If Range("C16") = Empty Then
        Zmax = min_max_from_file("Z2", input_fileNumber)
    End If
        
    Do Until EOF(input_fileNumber)
        Line Input #input_fileNumber, linefromfile
        lineitems = Split(linefromfile, " ")
        If CDbl(lineitems(1)) >= CDbl(Xmin) Then
            If CDbl(lineitems(1)) <= CDbl(Xmax) Then
                If CDbl(lineitems(0)) >= CDbl(Ymin) Then
                    If CDbl(lineitems(0)) <= CDbl(Ymax) Then
                        ' zapisz liniê do pliku
                        Write #output_fileNumber, linefromfile
                    End If
                End If
            End If
        End If
    Loop

    MsgBox "Plik zosta³ zapisany pod nazw¹: " & output_fullpath, , "raport"
GoTo koniec

'errors
error1:
'    If input_fullpath = "" Then
      MsgBox "Nie wybrano pliku. Program zostanie zakoñczony", , "OK"
      GoTo koniec
'    End If
    
koniec:
    Close #input_fileNumber
    Close #output_fileNumber

Call repace_quotes(output_fullpath)
    
End Sub

Private Function min_max_from_file(X_Y_min_max As String, input_fileNumber As Integer) As Single
    Dim lineitems

    Do Until EOF(input_fileNumber)
        Line Input #input_fileNumber, linefromfile
        lineitems = Split(linefromfile, " ")
        Select Case X_Y_min_max
            Case "X1"
                If lineitems(0) < min_max_from_file Then
                    min_max_from_file = lineitems(0)
            End If
            Case "X2"
                If lineitems(0) > min_max_from_file Then
                    min_max_from_file = lineitems(0)
            End If
            Case "Y1"
                If lineitems(1) < min_max_from_file Then
                    min_max_from_file = lineitems(1)
            End If
            Case "Y2"
                If lineitems(1) > min_max_from_file Then
                    min_max_from_file = lineitems(1)
            End If
            Case "Z1"
                If lineitems(2) < min_max_from_file Then
                    min_max_from_file = lineitems(2)
            End If
            Case "Z2"
                If lineitems(2) > min_max_from_file Then
                    min_max_from_file = lineitems(2)
            End If
            
        End Select
    Loop
    Seek input_fileNumber, 1
End Function

Sub repace_quotes(file_path As String)
    Dim filesize As Integer
    Dim entireline As String
    Dim strData() As String
    Dim MyData As String
    
    filesize = FreeFile()
    Open file_path For Binary As filesize
    MyData = Space$(LOF(1))
    Get filesize, , MyData
    Close filesize
    strData() = Split(MyData, vbCrLf)
    
    
    '~~> Open your file
    filesize = FreeFile()
    Open file_path For Output As #filesize
    
    For i = LBound(strData) To UBound(strData) - 1
        entireline = Replace(strData(i), """", "")
        '~~> Export Text
        Print #filesize, entireline
    Next i
    Close #filesize
End Sub

