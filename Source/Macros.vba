Attribute VB_Name = "Macros"
Private Sub protect_workbook()
    Main.Protect AllowFormattingCells:=True, _
                        AllowDeletingRows:=False, _
                        AllowInsertingRows:=False, _
                        UserInterfaceOnly:=True
                        
    ActiveSheet.EnableSelection = xlUnlockedCells
End Sub


Private Sub get_area_from_mag()

    If Not IsEmpty(vars.magnitude) Then
        vars.mag_area.Value = 10 ^ (-3.49 + 0.91 * vars.magnitude.Value)
    Else
        vars.mag_area.Value = ""
    End If
    
End Sub

Private Sub get_mechanism()
    
    If ((vars.rake.Value > -180 And vars.rake.Value < -150) Or _
            (vars.rake.Value > -30 And vars.rake.Value < 30) Or _
            (vars.rake.Value > 150 And vars.rake.Value < 180)) Then
        vars.mechanism.Value = "Strike-Slip"
    ElseIf vars.rake.Value > -120 And vars.rake.Value < -60 Then
        vars.mechanism.Value = "Normal"
    ElseIf vars.rake.Value > 60 And vars.rake.Value < 120 Then
        vars.mechanism.Value = "Reverse"
    ElseIf ((vars.rake.Value > 30 And vars.rake.Value < 60) Or _
                (vars.rake.Value > 120 And vars.rake.Value < 150) Or _
                (vars.rake.Value > -150 And vars.rake.Value < -120) Or _
                (vars.rake.Value > -60 And vars.rake.Value < -30)) Then
        vars.mechanism.Value = "Unspecified"
    End If
    
End Sub

Private Sub get_area_from_segment()

End Sub

Private Sub get_strike()

End Sub

Private Sub get_dip()

End Sub

Private Sub manage_segments()
    ' find out how many segments we currently have
    seg_count = Split(Main.Cells(Rows.Count, "B").End(xlUp).Value, " ")(1) * 1
    last_row = Main.Cells(Rows.Count, "C").End(xlUp).Row + 2
    
    Dim last_seg As Integer
    
    If seg_count > vars.segment_count Then
        ' find the last row we don't want to delete
        last_seg = vars.segment_start + (vars.segment_count * vars.segment_height)
        ' delete all the rows after this one
        Main.Rows(last_seg & ":" & last_row).EntireRow.Delete
    
    ElseIf seg_count < vars.segment_count Then
        Dim vert_num As Range
        new_seg_count = vars.segment_count - seg_count
        
        last_seg = vars.segment_start + (seg_count * vars.segment_height)
        
        new_seg_start = last_seg + 1
        'last_new_seg = last_row + 1 + new_seg_count * vars.segment_height
        seg_row = new_seg_start
        For current_seg = seg_count + 1 To seg_count + new_seg_count
            Main.Range("B" & seg_row).Value = "Segment " & current_seg
            vars.blank_seg.Copy
            Main.Range("C" & seg_row, "G" & seg_row + vars.segment_height - 1).PasteSpecial
            
            Set vert_num = Main.Range("C" & seg_row)
            
            With vert_num.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="='Lookup Values'!$A$1:$A$100"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .ShowInput = True
                .ShowError = True
            End With
            
            seg_row = seg_row + vars.segment_height
        Next
        
    End If
    
End Sub

Private Sub manage_vertices(vertices As Range)
    ' number of vertices selected
    new_vert_count = vertices.Value
    color_index = vertices.Interior.ColorIndex
    row_num = vertices.Row
    col_num = vertices.Column
    
    ' current number of vertices that already exist
    vert_count = 0
    For Each cell In Main.Rows(row_num + 1).Cells
        If cell.Interior.ColorIndex = color_index Then
            vert_count = vert_count + 1
        End If
    Next
    
    If new_vert_count < vert_count Then
        Main.Range(Cells(row_num, col_num + 1 + new_vert_count), _
                           Cells(row_num + 3, col_num + vert_count)).Delete shift:=xlToLeft
    
    ElseIf new_vert_count > vert_count Then
        vert_diff = new_vert_count - vert_count
        
        For new_vert = vert_count + 1 To new_vert_count
            Main.Cells(row_num, col_num + new_vert).Value = new_vert
            vars.blank_seg_col.Copy
            With Main.Range(Cells(row_num + 1, col_num + new_vert), Cells(row_num + 3, col_num + new_vert))
                .PasteSpecial
                .Locked = False
            End With
            
            
        Next
        
    End If

End Sub

Private Sub check_fault_ref()
    If IsEmpty(vars.fault_ref.Value) Then
        vars.fault_ref.Value = "None"
    End If

End Sub

Private Sub get_time()

End Sub

Private Sub get_date()

End Sub

Function num_check(ByVal num_in As String)
    Dim num_str As String
    num_str = "0123456789"
    
    For num = 1 To Len(num_in)
        check_num = Mid(num_in, num, 1)
        If InStr(num_str, check_num) = 0 Then
            num_check = False
            Exit Function
        End If
    Next
    
    num_check = True
End Function

Private Sub finite_fault_model()

    last_row = Main.Range("C1").SpecialCells(xlCellTypeLastCell).Row
    
    If vars.finite_fault_model.Value = "Yes" Then
        ' Just unhide all rows, because finding the right ones is too hard
        Main.Range("C:C").EntireRow.Hidden = False
    ElseIf vars.finite_fault_model.Value = "No" Then
        Main.Range("A" & vars.segment_count.Row, "A" & last_row).EntireRow.Hidden = True
    End If

End Sub
