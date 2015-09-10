Attribute VB_Name = "vars"
Public eq_name As Range
Public eq_date As Range
Public eq_time As Range
Public fault_ref As Range
Public magnitude As Range
Public mag_area As Range
Public rake As Range
Public mechanism As Range
Public hyp_long As Range
Public hyp_lat As Range
Public hyp_depth As Range
Public segment_range As Range
Public segment_start As Integer
Public segment_height As Integer
Public segment_count As Range
Public blank_seg As Range
Public blank_seg_col As Range
Public finite_fault_model As Range
Public segment_max As Integer

Public seg1_range As Range
Public seg2_range As Range
Public seg3_range As Range
Public seg4_range As Range
Public seg5_range As Range
Public seg1_copy As Range
Public seg2_copy As Range
Public seg3_copy As Range
Public seg4_copy As Range
Public seg5_copy As Range
Public seg1_plot As Range
Public seg2_plot As Range
Public seg3_plot As Range
Public seg4_plot As Range
Public seg5_plot As Range

Public plot_area As Range


Private Sub setup_variables()
    Set vars.eq_name = Main.Range("B7")
    Set vars.eq_date = Main.Range("B8")
    Set vars.eq_time = Main.Range("B9")
    Set vars.fault_ref = Main.Range("B10")
    Set vars.magnitude = Main.Range("B13")
    Set vars.mag_area = Main.Range("B14")
    Set vars.rake = Main.Range("B15")
    Set vars.mechanism = Main.Range("B16")
    Set vars.hyp_long = Main.Range("C17")
    Set vars.hyp_lat = Main.Range("C18")
    Set vars.hyp_depth = Main.Range("C19")
    Set vars.finite_fault_model = Main.Range("B20")
    Set vars.segment_count = Main.Range("B21")
    Set vars.blank_seg = Lookup.Range("E1:I7")
    Set vars.blank_seg_col = Lookup.Range("I2:I4")
    
    vars.segment_start = 23
    vars.segment_height = 7
    vars.segment_max = 5
    
    ' setup ranges to copy segments
    Set seg1_range = Main.Range("D25", "W27")
    Set seg2_range = Main.Range("D32", "W34")
    Set seg3_range = Main.Range("D39", "W41")
    Set seg4_range = Main.Range("D46", "W48")
    Set seg5_range = Main.Range("D53", "W55")
    Set seg1_copy = Lookup.Range("N1", "AG3")
    Set seg2_copy = Lookup.Range("N4", "AG6")
    Set seg3_copy = Lookup.Range("N7", "AG9")
    Set seg4_copy = Lookup.Range("N10", "AG12")
    Set seg5_copy = Lookup.Range("N13", "AG15")
    
    Set seg1_plot = Lookup.Range("M1", "AG2")
    Set seg2_plot = Lookup.Range("M4", "AG5")
    Set seg3_plot = Lookup.Range("M7", "AG8")
    Set seg4_plot = Lookup.Range("M10", "AG11")
    Set seg5_plot = Lookup.Range("M13", "AG14")
    
    Set plot_area = Main.Range("E5", "H20")
End Sub

