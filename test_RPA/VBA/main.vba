


Option Explicit

Function FindIndex(arr, val)
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If Not IsError(Application.Match(val, Application.Index(arr, r, 0), 0)) Then
            FindIndex = Application.Match(val, Application.Index(arr, r, 0), 0)
            Exit Function
        End If
    Next r
    End Function
Function SortArrayBySecondColumn(arr() As Variant) As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    
    ' Сортировка массива по второму столбцу
    For i = LBound(arr, 1) To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            If arr(i, 2) < arr(j, 2) Then
                ' Перестановка элементов
                temp = arr(i, 1)
                arr(i, 1) = arr(j, 1)
                arr(j, 1) = temp
                
                temp = arr(i, 2)
                arr(i, 2) = arr(j, 2)
                arr(j, 2) = temp
            End If
        Next j
    Next i
    
    SortArrayBySecondColumn = arr
End Function

Sub main()


    Dim tabel_number() As String
    Dim last_name() As String
    Dim first_name() As String
    Dim second_name() As String
    Dim department() As Long
    Dim ID_department() As Long
    Dim department_name() As String
    Dim task_tabel_number() As String
    Dim final_array(1 To 14, 1 To 2) As Variant
    Dim tmp_string As String
    Dim tmp_count As Long
    Dim sort_array(1 To 4, 1 To 2) As Variant
    Dim final_sort_array() As Variant
    Dim tmp_array(1 To 7, 1 To 2) As Variant
    Dim final_tmp_array() As Variant

    Dim wsTabel As Worksheet
    Dim wsDepartment As Worksheet
    Dim wsTask As Worksheet
    Dim task_counts
    Dim department_dict
    

    Set wsTabel = ThisWorkbook.Worksheets("Сотрудники")
    Set wsDepartment = ThisWorkbook.Worksheets("Отделы")
    Set wsTask = ThisWorkbook.Worksheets("Задачи")

    Dim i As Long
    Dim j As Long
    Dim tmp As Long
    Dim tmp1 As Long

    ReDim tabel_number(1 To 7)
    ReDim last_name(1 To 7)
    ReDim second_name(1 To 7)
    ReDim first_name(1 To 7)
    ReDim department(1 To 7)

    For i = 1 To 7
        tabel_number(i) = wsTabel.Range("A1").Offset(i)
        last_name(i) = wsTabel.Range("B1").Offset(i)
        first_name(i) = wsTabel.Range("C1").Offset(i)
        second_name(i) = wsTabel.Range("D1").Offset(i)
        department(i) = wsTabel.Range("F1").Offset(i)
    Next i

    ReDim ID_department(1 To 4)
    ReDim department_name(1 To 4)

    For i = 1 To 4
        ID_department(i) = wsDepartment.Range("A1").Offset(i)
        department_name(i) = wsDepartment.Range("B1").Offset(i)
    Next i

    ReDim task_tabel_number(1 To 21)

    For i = 1 To 21
        task_tabel_number(i) = wsTask.Range("B1").Offset(i)
    Next i
    
     Set task_counts = CreateObject("Scripting.Dictionary")
    
    ' Составляем словарь и подсчитываем выполненные задачи каждым сотрудником
    For i = 1 To UBound(task_tabel_number)
        If Not task_counts.Exists(task_tabel_number(i)) Then
            task_counts.Add task_tabel_number(i), 1
        Else
            task_counts.Item(task_tabel_number(i)) = task_counts.Item(task_tabel_number(i)) + 1
        End If
    Next i
    
    ' Составляем словарь Название департамента : номер департамента
    Set department_dict = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(department_name)
        department_dict.Add department_name(i), ID_department(i)
    Next i
    
    tmp = 1
    tmp1 = 1
    
    For j = 1 To UBound(department_name)
        sort_array(j, 1) = department_name(j)
        tmp_count = 0
        
        For i = 1 To UBound(tabel_number)
            If department(i) = ID_department(j) Then
                tmp_count = tmp_count + task_counts.Item(task_tabel_number(i))
            End If
        Next i
        sort_array(j, 2) = tmp_count
        tmp_count = 0
    Next j
        final_sort_array = SortArrayBySecondColumn(sort_array)
        
    
    For j = 1 To 4
    ' Вносим названия отделов и сумму выполненных задач
        final_array(tmp1, 1) = final_sort_array(j, 1)
        final_array(tmp1, 2) = final_sort_array(j, 2)
        tmp_count = 0
        i = 1
        tmp = 1
        
        ' Очищаем временный массив перед новым использованием
        Erase tmp_array
        
        For i = 1 To UBound(tabel_number)
        
            If department(i) = department_dict.Item(final_sort_array(j, 1)) Then
    
                tmp_array(tmp, 1) = last_name(i) & " " & Mid(first_name(i), 1, 1) & ". " & Mid(second_name(i), 1, 1) & ". "
                tmp_array(tmp, 2) = task_counts.Item(task_tabel_number(i))
                tmp = tmp + 1
                tmp_count = tmp_count + task_counts.Item(task_tabel_number(i))
            
            End If
        
        Next i
        
        ' Сортируем по убыванию
        final_tmp_array = SortArrayBySecondColumn(tmp_array)
        
        tmp = 1
        For i = 1 To UBound(final_tmp_array, 1)
        If Len(final_tmp_array(i, 1)) <> 0 Then

            final_array(tmp1 + i, 1) = final_tmp_array(i, 1)
            final_array(tmp1 + i, 2) = final_tmp_array(i, 2)
            
        tmp = tmp + 1
       End If
        Next i
        tmp1 = tmp1 + tmp
    Next j
    
    For i = 1 To 14
    
    ' Временная таблица
    Worksheets("Сотрудники").Cells(i + 10, 1) = final_array(i, 1)
    Worksheets("Сотрудники").Cells(i + 10, 2) = final_array(i, 2)
    Next i
    
   ' Создаём Word файл и заносим таблицу туда
    
    Dim word_app As Word.Application
    Dim my_doc As Word.Document
    
    Set word_app = New Word.Application
    
    word_app.Visible = True
    
    Set my_doc = word_app.Documents.Add()
    
    ThisWorkbook.Worksheets("Сотрудники").Range("A11:B21").Copy
    
    my_doc.Paragraphs(1).Range.PasteExcelTable _
 LinkedToExcel:=False, _
 WordFormatting:=False, _
 RTF:=False
 
  my_doc.SaveAs2 "Отчёт о загрузке"
  
  
  ' Удаляем временную таблицу
  ThisWorkbook.Worksheets("Сотрудники").Range("A11:B21").Delete
    

End Sub
