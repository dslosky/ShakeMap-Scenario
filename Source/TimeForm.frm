VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TimeForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4000
   ClientLeft      =   -880
   ClientTop       =   -5060
   ClientWidth     =   7060
   OleObjectBlob   =   "TimeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TimeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub okay_Click()
    hour = Me.hour.Text
    minute = Me.minute.Text
    seconds = Me.seconds.Text
    Dim valid As Boolean
    valid = True
    
    If Len(hour) < 2 Then
        MsgBox "You have entered an invalid hour"
        valid = False
    End If
    If Len(minute) < 2 Then
        MsgBox "You have entered and invalid minute"
        valid = False
    End If
    If Len(seconds) < 2 Then
        MsgBox "You have entered an invalid second"
        valid = False
    End If
    
    If valid Then
        ' stop excel from formatting as date...
        vars.eq_time.NumberFormat = "@"
        Dim time_str As String
        time_str = "'" & hour & ":" & minute & ":" & seconds
        vars.eq_time.Value = time_str

        Unload Me
    End If
    
End Sub

Private Sub hour_Change()
    input_len = Len(Me.hour.Text)
    If input_len > 2 Then
        Me.hour.Text = Left(Me.hour.Text, 2)
    End If
    If Not num_check(Me.hour.Text) Then
        MsgBox "This field must only contain numbers"
        
        If input_len = 1 Then
            Me.hour.Text = ""
        Else:
            Me.hour.Text = Left(Me.hour.Text, 1)
        End If
        
    End If
    
    If Len(Me.hour.Text) = 2 Then
        Me.minute.SetFocus
    End If
    
End Sub

Private Sub minute_Change()
    input_len = Len(Me.minute.Text)
    If input_len > 2 Then
        Me.minute.Text = Left(Me.minute.Text, 2)
    End If
    If Not num_check(Me.minute.Text) Then
        MsgBox "This field must only contain numbers"
        
        If input_len = 1 Then
            Me.minute.Text = ""
        Else:
            Me.minute.Text = Left(Me.minute.Text, 1)
        End If
        
    End If
    
    If Len(Me.minute.Text) = 2 Then
        Me.seconds.SetFocus
    End If
End Sub
Private Sub UserForm_Initialize()
    If Not IsEmpty(vars.eq_time) And Not InStr(vars.eq_time.Value, ":") = 0 Then
        Dim time_str() As String
        time_str() = Split(vars.eq_time.Value, ":")
        
        Me.hour.Text = time_str(0)
        Me.minute.Text = time_str(1)
        Me.seconds.Text = time_str(2)
    End If
End Sub

Private Sub seconds_Change()
    input_len = Len(Me.seconds.Text)
    If input_len > 2 Then
        Me.seconds.Text = Left(Me.seconds.Text, 2)
    End If
    If Not num_check(Me.seconds.Text) Then
        MsgBox "This field must only contain numbers"
        
        If input_len = 1 Then
            Me.seconds.Text = ""
        Else
            Me.seconds.Text = Left(Me.seconds.Text, 1)
        End If
        
    End If
    
    If Len(Me.seconds.Text) = 4 Then
        Me.okay.SetFocus
    End If
End Sub


