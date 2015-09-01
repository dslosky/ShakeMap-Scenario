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
End Sub

