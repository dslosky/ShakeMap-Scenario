VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DateForm 
   ClientHeight    =   4000
   ClientLeft      =   -1600
   ClientTop       =   -9200.001
   ClientWidth     =   7060
   OleObjectBlob   =   "DateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub okay_Click()
    month = Me.month.Text
    day = Me.day.Text
    year = Me.year.Text
    Dim valid As Boolean
    valid = True
    
    If Len(month) < 2 Then
        MsgBox "You have entered an invalid month"
        valid = False
    End If
    If Len(day) < 2 Then
        MsgBox "You have entered and invalid day"
        valid = False
    End If
    If Len(year) < 4 Then
        MsgBox "You have entered an invalid year"
        valid = False
    End If
    
    If valid Then
        vars.eq_date.NumberFormat = "@"
        vars.eq_date.Value = "'" & month & "/" & day & "/" & year
        
        Unload Me
    End If
    
End Sub

Private Sub month_Change()
    input_len = Len(Me.month.Text)
    If input_len > 2 Then
        Me.month.Text = Left(Me.month.Text, 2)
    End If
    If Not num_check(Me.month.Text) Then
        MsgBox "This field must only contain numbers"
        
        If input_len = 1 Then
            Me.month.Text = ""
        Else:
            Me.month.Text = Left(Me.month.Text, 1)
        End If
        
    End If
    
    If Len(Me.month.Text) = 2 Then
        Me.day.SetFocus
    End If
    
End Sub

Private Sub day_Change()
    input_len = Len(Me.day.Text)
    If input_len > 2 Then
        Me.day.Text = Left(Me.day.Text, 2)
    End If
    If Not num_check(Me.day.Text) Then
        MsgBox "This field must only contain numbers"
        
        If input_len = 1 Then
            Me.day.Text = ""
        Else:
            Me.day.Text = Left(Me.day.Text, 1)
        End If
        
    End If
    
    If Len(Me.day.Text) = 2 Then
        Me.year.SetFocus
    End If
End Sub
Private Sub UserForm_Initialize()
    If Not IsEmpty(vars.eq_date) And Not InStr(vars.eq_date.Value, "/") = 0 Then
        Dim date_str() As String
        date_str() = Split(vars.eq_date, "/")
        
        Me.month.Text = date_str(0)
        Me.day.Text = date_str(1)
        Me.year.Text = date_str(2)
    End If
End Sub

Private Sub year_Change()
    input_len = Len(Me.year.Text)
    If input_len > 4 Then
        Me.year.Text = Left(Me.year.Text, 4)
    End If
    If Not num_check(Me.year.Text) Then
        MsgBox "This field must only contain numbers"
        
        If input_len = 1 Then
            Me.year.Text = ""
        ElseIf input_len = 2 Then
            Me.year.Text = Left(Me.year.Text, 1)
        ElseIf input_len = 3 Then
            Me.year.Text = Left(Me.year.Text, 2)
        ElseIf input_len = 4 Then
            Me.year.Text = Left(Me.year.Text, 3)
        End If
        
    End If
    
    If Len(Me.year.Text) = 4 Then
        Me.okay.SetFocus
    End If
End Sub

