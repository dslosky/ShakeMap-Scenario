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
        area_val = 10 ^ (-3.49 + 0.91 * vars.magnitude.Value)
        
        If area_val < 2 Then
            vars.mag_area.Value = Round(area_val, 2)
        Else
            vars.mag_area.Value = Round(area_val, 0)
        End If
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
    ' instead of copying and pasting, let's just hide and unhide rows
    total_rows = vars.segment_height * vars.segment_max + vars.segment_start
    show_rows = vars.segment_height * vars.segment_count + vars.segment_start

    
    Main.Range("A" & show_rows + 1, "A" & total_rows).Rows.Hidden = True
    Main.Range("A" & vars.segment_start, "A" & show_rows).Rows.Hidden = False
    
End Sub

Private Sub manage_segments_old()
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
                xlBetween, Formula1:="='Lookup Values'!$A$4:$A$20"
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
    
    Run "copy_vertices"

End Sub

Private Sub copy_vertices()
    
    ' copy segments
    vars.seg1_copy.Value = vars.seg1_range.Value
    vars.seg2_copy.Value = vars.seg2_range.Value
    vars.seg3_copy.Value = vars.seg3_range.Value
    vars.seg4_copy.Value = vars.seg4_range.Value
    vars.seg5_copy.Value = vars.seg5_range.Value
    
    If IsEmpty(Lookup.Range("N1")) Then
        Lookup.Range("N1").Value = 0
    End If
    If IsEmpty(Lookup.Range("N2")) Then
        Lookup.Range("N2").Value = 0
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
        Main.Range("A" & vars.segment_count.Row, "A" & vars.segment_start).EntireRow.Hidden = False
        
        Run "manage_segments"
        
        ' make a plot
        Run "make_plot"
    ElseIf vars.finite_fault_model.Value = "No" Then
        Main.Range("A" & vars.segment_count.Row, "A" & last_row).EntireRow.Hidden = True
        
        ' delete the segment plot if it exists
        If Main.ChartObjects.Count > 0 Then
            Main.ChartObjects.Delete
        End If
    End If

End Sub

Private Sub make_plot()
    ' Delete plots that already exist
    If Main.ChartObjects.Count > 0 Then
        Main.ChartObjects.Delete
    End If
    
    Dim chart_obj As ChartObject
    With vars.plot_area
        Set chart_obj = Main.ChartObjects.Add(.Left, .Top, .Width, .Height)
    End With
    
    Set new_chart = chart_obj.Chart
    
    With new_chart
        .ChartType = xlXYScatterLines
        'Set data source range.
        .SetSourceData Source:=vars.seg1_plot, PlotBy:= _
          xlRows
        .HasTitle = True
        .ChartTitle.Text = "Segments"
         'X axis name
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Longitude"
         'y-axis name
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Latitude"
        'The Parent property is used to set properties of
        'the Chart.
        'With .Parent
        '    .Top = Range("F9").Top
        '    .Left = Range("F9").Left
        '    .Name = "ToolsChart2"
        'End With
    End With
    
    If vars.segment_count = 1 Then
        ' segment one is already in there!
    ElseIf vars.segment_count = 2 Then
        vars.seg2_plot.Copy
        new_chart.Paste
    ElseIf vars.segment_count = 3 Then
        vars.seg2_plot.Copy
        new_chart.Paste
        vars.seg3_plot.Copy
        new_chart.Paste
    ElseIf vars.segment_count = 4 Then
        vars.seg2_plot.Copy
        new_chart.Paste
        vars.seg3_plot.Copy
        new_chart.Paste
        vars.seg4_plot.Copy
        new_chart.Paste
    ElseIf vars.segment_count = 5 Then
        vars.seg2_plot.Copy
        new_chart.Paste
        vars.seg3_plot.Copy
        new_chart.Paste
        vars.seg4_plot.Copy
        new_chart.Paste
        vars.seg5_plot.Copy
        new_chart.Paste
    End If
    
End Sub
