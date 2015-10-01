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
    
    vars.hypo_plot.Copy
    new_chart.Paste
    
    With new_chart.SeriesCollection("Hypocenter")
        .MarkerStyle = xlMarkerStyleTriangle
        .MarkerForegroundColor = RGB(255, 0, 0)
        .MarkerBackgroundColor = RGB(255, 0, 0)
        .MarkerSize = 9
        .Format.Line.Visible = False
    End With
End Sub

Private Sub export_docs()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo ExitHandler
    
    Application.Run "export_xml"
    
ExitHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub export_xml()
    Application.Run "make_xml_table"
    If vars.export_ready = False Then GoTo ExitHandler
    
    Dim dir As String
    Dim file_name As String
    Dim xml_path As String
    dir = Application.ActiveWorkbook.Path
    file_name = "shakemap_scenario.xml"

    If InStr(getOS, "Windows") = 0 Then
        xml_path = dir & ":" & file_name
    Else
        xml_path = dir & "\" & file_name
    End If
    
    Dim xml_string As String
    xml_string = get_xml_string()
    
    Open xml_path For Output As #2
        Print #2, xml_string
    Close #2
    
    MsgBox "Your scenario has been successfully exported at: " & vbNewLine & _
                    xml_path
ExitHandler:
End Sub

Public Function get_xml_string()

dtd_string = "<?xml version=""1.0"" encoding=""US-ASCII"" standalone=""yes""?>" & _
                        "<!DOCTYPE earthquake [" & vbNewLine & _
                        "<!ELEMENT  earthquake EMPTY>" & vbNewLine & _
                        "<!ATTLIST earthquake" & vbNewLine & _
                        "  id        ID  #REQUIRED" & vbNewLine & _
                        "  lat       CDATA   #REQUIRED" & vbNewLine & _
                        "  lon       CDATA   #REQUIRED" & vbNewLine & _
                        "  mag       CDATA   #REQUIRED" & vbNewLine & _
                        "  year          CDATA   #REQUIRED" & vbNewLine & _
                        "  month         CDATA   #REQUIRED" & vbNewLine & _
                        "  day           CDATA   #REQUIRED" & vbNewLine & _
                        "  hour          CDATA   #REQUIRED" & vbNewLine & _
                        "  minute        CDATA   #REQUIRED" & vbNewLine & _
                        "  second        CDATA   #REQUIRED" & vbNewLine & _
                        "  timezone      CDATA   #REQUIRED" & vbNewLine & _
                        "  depth     CDATA   #REQUIRED" & vbNewLine & _
                        "  type      CDATA   #REQUIRED" & vbNewLine & _
                        "  locstring CDATA   #REQUIRED" & vbNewLine & _
                        "  pga       CDATA   #REQUIRED" & vbNewLine & _
                        "  pgv       CDATA   #REQUIRED" & vbNewLine & _
                        "  sp03      CDATA   #REQUIRED" & vbNewLine & _
                        "  sp10      CDATA   #REQUIRED" & vbNewLine & _
                        "  sp30      CDATA   #REQUIRED" & vbNewLine & _
                        "  created   CDATA   #REQUIRED"

dtd_string = dtd_string & ">" & vbNewLine & _
                     "]>" & vbNewLine

eq_String = "<earthquake "

Dim att_val As String
For Each att In XML_Table.Range("A1", "Q1")
    att_val = XML_Table.Cells(att.Row + 1, att.Column).Value
    
    eq_String = eq_String & att.Value & "=""" & att_val & """ "
Next att

eq_String = eq_String & "/>"

get_xml_string = dtd_string & eq_String
End Function


Private Sub make_xml_table()

    Application.Run "check_input"
    If vars.export_ready = False Then GoTo ExitHandler

    ' split date
    Dim eq_date() As String
    eq_date = Split(vars.eq_date, "/")
    
    ' split time
    Dim eq_time() As String
    eq_time = Split(vars.eq_time.Value, ":")
    
    ' id
    XML_Table.Range("A2").Value = get_id()
    
    ' lat
    XML_Table.Range("B2").Value = vars.hyp_lat
    
    ' lon
    XML_Table.Range("C2").Value = vars.hyp_long
    
    ' mag
    XML_Table.Range("D2").Value = vars.magnitude
    
    ' year
    XML_Table.Range("E2").Value = eq_date(2)
    
    ' month
    XML_Table.Range("F2").Value = eq_date(0)
    
    ' day
    XML_Table.Range("G2").Value = eq_date(1)
    
    ' hour
    XML_Table.Range("H2").Value = eq_time(0)
    
    ' minute
    XML_Table.Range("I2").Value = eq_time(1)
    
    ' second
    XML_Table.Range("J2").Value = eq_time(2)
    
    ' timezone
    XML_Table.Range("K2").Value = vars.timezone.Value
    
    ' depth
    XML_Table.Range("L2").Value = vars.hyp_depth
    
    ' locstring
    XML_Table.Range("M2").Value = vars.eq_name
    
    ' created
    XML_Table.Range("N2").Value = ""
    
    ' otime
    XML_Table.Range("O2").Value = ""
    
    ' type
    XML_Table.Range("P2").Value = ""
    
    ' network
    XML_Table.Range("Q2").Value = vars.network
    
ExitHandler:
If vars.export_ready = False Then
    MsgBox "Failed to export. Some required fields have been left " & _
                  "blank. These cells have been highlighted for you and must " & _
                  "be completed before exporting."
End If


End Sub

Private Sub check_input()

    Dim req_fields(0 To 12) As Variant

    req_fields(0) = vars.eq_name.Address
    req_fields(1) = vars.eq_date.Address
    req_fields(2) = vars.eq_time.Address
    req_fields(3) = vars.timezone.Address
    req_fields(4) = vars.network.Address
    req_fields(5) = vars.fault_ref.Address
    req_fields(6) = vars.magnitude.Address
    req_fields(7) = vars.rake.Address
    req_fields(8) = vars.hyp_lat.Address
    req_fields(9) = vars.hyp_long.Address
    req_fields(10) = vars.hyp_depth.Address
    req_fields(11) = vars.finite_fault_model.Address
    req_fields(12) = vars.segment_count.Address
    
    Dim check_range As Range
    vars.export_ready = True
    For Each range_address In req_fields
        Set check_range = Main.Range(range_address)
            If IsEmpty(check_range) Then
                check_range.Interior.Color = 6579455
                vars.export_ready = False
            Else
                check_range.Interior.Color = 10213316
            End If
    Next range_address
                         

End Sub

Public Function get_id()
    
    Dim id As String
    id = vars.eq_name
    id = Replace(id, " ", "_")
    id = Replace(id, "`", "")
    id = Replace(id, "~", "")
    id = Replace(id, "!", "")
    id = Replace(id, "@", "")
    id = Replace(id, "#", "")
    id = Replace(id, "$", "")
    id = Replace(id, "%", "")
    id = Replace(id, "^", "")
    id = Replace(id, "&", "")
    id = Replace(id, "*", "")
    id = Replace(id, "(", "")
    id = Replace(id, ")", "")
    id = Replace(id, "=", "")
    id = Replace(id, "+", "")
    id = Replace(id, ":", "")
    id = Replace(id, ";", "")
    id = Replace(id, "'", "")
    id = Replace(id, """", "")
    id = Replace(id, "<", "")
    id = Replace(id, ">", "")
    id = Replace(id, ",", "")
    id = Replace(id, ".", "")
    id = Replace(id, "/", "")
    id = Replace(id, "?", "")
    id = Replace(id, "_", "")
    id = Replace(id, "{", "")
    id = Replace(id, "}", "")
    id = Replace(id, "[", "")
    id = Replace(id, "]", "")
    id = Replace(id, "\", "")
    id = Replace(id, "|", "")
    
    get_id = id & "_eq"
End Function
