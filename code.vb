VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim flag1, flag2, flag3, startup, difference As Boolean
Dim frame, x, y As Long
Dim speedtimer, speed1, speed2 As Double
Dim rule1, rule2 As String
Dim ce(4 To 37, 3 To 76) As Variant


Public Sub START_Click()
    If START.Caption = "STOP" Then GoTo label2
    
    Application.ScreenUpdating = False
    START.Caption = "STOP"
    START.BackColor = &HFF00FF
    Cells(28, 82).Value = "Running"
    Cells(29, 82).Value = ""
    
    rule1 = ""
    If CheckBox9.Value = True Then rule1 = rule1 & "1"
    If CheckBox10.Value = True Then rule1 = rule1 & "2"
    If CheckBox11.Value = True Then rule1 = rule1 & "3"
    If CheckBox12.Value = True Then rule1 = rule1 & "4"
    If CheckBox13.Value = True Then rule1 = rule1 & "5"
    If CheckBox14.Value = True Then rule1 = rule1 & "6"
    If CheckBox15.Value = True Then rule1 = rule1 & "7"
    If CheckBox16.Value = True Then rule1 = rule1 & "8"

    rule2 = ""
    If CheckBox1.Value = True Then rule2 = rule2 & "1"
    If CheckBox2.Value = True Then rule2 = rule2 & "2"
    If CheckBox3.Value = True Then rule2 = rule2 & "3"
    If CheckBox4.Value = True Then rule2 = rule2 & "4"
    If CheckBox5.Value = True Then rule2 = rule2 & "5"
    If CheckBox6.Value = True Then rule2 = rule2 & "6"
    If CheckBox7.Value = True Then rule2 = rule2 & "7"
    If CheckBox8.Value = True Then rule2 = rule2 & "8"

    Cells(29, 82).Value = "S" & rule1 & "/B" & rule2
    
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    Application.Cursor = xlNorthwestArrow
    flag1 = False
    flag2 = False
    flag3 = False
    
    Dim i, iii, r As Range
    Dim ii As Integer
    speedtimer = Timer
    
label1:
    For Each i In Range("c4:bx37")
        ii = 0
        
        With i
            If .Offset(-1, -1).Value = "." Then ii = ii + 1
            If .Offset(-1, 0).Value = "." Then ii = ii + 1
            If .Offset(-1, 1).Value = "." Then ii = ii + 1
            If .Offset(0, -1).Value = "." Then ii = ii + 1
            If .Offset(0, 1).Value = "." Then ii = ii + 1
            If .Offset(1, -1).Value = "." Then ii = ii + 1
            If .Offset(1, 0).Value = "." Then ii = ii + 1
            If .Offset(1, 1).Value = "." Then ii = ii + 1
        End With
      
        If i.Value = "" Then
            If CheckBox1.Value = True And ii = 1 _
                Or CheckBox2.Value = True And ii = 2 _
                Or CheckBox3.Value = True And ii = 3 _
                Or CheckBox4.Value = True And ii = 4 _
                Or CheckBox5.Value = True And ii = 5 _
                Or CheckBox6.Value = True And ii = 6 _
                Or CheckBox7.Value = True And ii = 7 _
                Or CheckBox8.Value = True And ii = 8 Then
                ce(i.Row, i.Column) = "."
            Else
                ce(i.Row, i.Column) = ""
            End If
        Else
            If CheckBox9.Value = True And ii = 1 _
                Or CheckBox10.Value = True And ii = 2 _
                Or CheckBox11.Value = True And ii = 3 _
                Or CheckBox12.Value = True And ii = 4 _
                Or CheckBox13.Value = True And ii = 5 _
                Or CheckBox14.Value = True And ii = 6 _
                Or CheckBox15.Value = True And ii = 7 _
                Or CheckBox16.Value = True And ii = 8 Then
                ce(i.Row, i.Column) = "."
            Else
                ce(i.Row, i.Column) = ""
            End If
        End If
    Next i
    
    If CheckBox18.Value Then
        For Each iii In Range("c4:bx37")
            If iii.Value <> ce(iii.Row, iii.Column) Then difference = True
        Next iii
        If difference = False Then flag3 = True
    End If
    
    For Each r In Range("c4:bx37")
        r.Value = ce(r.Row, r.Column)
    Next r
    
    If CheckBox17.Value Then
        frame = frame + 1
        
        If frame Mod 10 = 1 Then
            Cells(31, 82).Value = frame & "st"
        ElseIf frame Mod 10 = 2 Then
            Cells(31, 82).Value = frame & "nd"
        ElseIf frame Mod 10 = 3 Then
            Cells(31, 82).Value = frame & "rd"
        Else
            Cells(31, 82).Value = frame & "th"
        End If
        
        speed1 = Timer - speedtimer
        Cells(33, 82).Value = 1 / speed1
        speed2 = speed2 + speed1
        Cells(35, 82).Value = 1 / (speed2 / frame)
        speedtimer = Timer

        rule1 = ""
        If CheckBox9.Value = True Then rule1 = rule1 & "1"
        If CheckBox10.Value = True Then rule1 = rule1 & "2"
        If CheckBox11.Value = True Then rule1 = rule1 & "3"
        If CheckBox12.Value = True Then rule1 = rule1 & "4"
        If CheckBox13.Value = True Then rule1 = rule1 & "5"
        If CheckBox14.Value = True Then rule1 = rule1 & "6"
        If CheckBox15.Value = True Then rule1 = rule1 & "7"
        If CheckBox16.Value = True Then rule1 = rule1 & "8"

        rule2 = ""
        If CheckBox1.Value = True Then rule2 = rule2 & "1"
        If CheckBox2.Value = True Then rule2 = rule2 & "2"
        If CheckBox3.Value = True Then rule2 = rule2 & "3"
        If CheckBox4.Value = True Then rule2 = rule2 & "4"
        If CheckBox5.Value = True Then rule2 = rule2 & "5"
        If CheckBox6.Value = True Then rule2 = rule2 & "6"
        If CheckBox7.Value = True Then rule2 = rule2 & "7"
        If CheckBox8.Value = True Then rule2 = rule2 & "8"

        Cells(29, 82).Value = "S" & rule1 & "/B" & rule2
    End If
    
    Application.ScreenUpdating = True
    DoEvents
    If flag1 = False And flag2 = False And flag3 = False Then
        Application.ScreenUpdating = False
        difference = False
        GoTo label1
    ElseIf flag3 = True Then
        Application.ScreenUpdating = False
        START.Caption = "START"
        START.BackColor = &HFF797C
        Cells(28, 82).Value = "Paused"
        Application.ScreenUpdating = True
        Exit Sub
    Else
        Application.ScreenUpdating = False
        START.Caption = "START"
        START.BackColor = &HFF797C
        If flag1 = True Then Cells(28, 82).Value = "Paused"
        Application.ScreenUpdating = True
        Exit Sub
    End If

label2:
    flag1 = True
End Sub


Public Sub CLEAR_Click()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    flag2 = True
    frame = "0"
    speed2 = "0"
    Cells(28, 82).Value = "Cleared"
    Cells(29, 82).Value = ""
    Cells(31, 82).Value = ""
    Cells(33, 82).Value = ""
    Cells(35, 82).Value = ""
    Range("c4:bx137").ClearContents
    START.Caption = "START"
    START.BackColor = &HFF797C
    For x = 4 To 36
        For y = 3 To 76
            ce(x, y) = ""
        Next y
    Next x
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Application.EnableEvents = False

    If startup Then
        frame = "0"
        speed2 = "0"
        startup = False
    End If
    
    Dim i As Range
    
    If Selection.Cells(1, 1).Row < 4 Then GoTo label1
    If Selection.Cells(1, 1).Column < 3 Then GoTo label1
    If Selection.Cells(1, 1).Offset(Selection.Rows.Count - 1, Selection.Columns.Count - 1).Row > 37 Then GoTo label1
    If Selection.Cells(1, 1).Offset(Selection.Rows.Count - 1, Selection.Columns.Count - 1).Column > 76 Then GoTo label1
    
    Application.ScreenUpdating = False
    
    If Selection.Cells(1, 1).Value = "" Then
        For Each i In Target
        i.Value = "."
        Next i
    Else
        For Each i In Target
        i.Value = ""
        Next i
    End If
    
    Application.ScreenUpdating = True
    
label1:
    Application.EnableEvents = True
End Sub


Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Application.EnableEvents = False
    
    If Target.Row < 4 Then GoTo label1
    If Target.Column < 3 Then GoTo label1
    If Target.Row > 37 Then GoTo label1
    If Target.Column > 76 Then GoTo label1

    Target.Value = ""
    Cancel = True
    
label1:
    Application.EnableEvents = True
End Sub


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Application.EnableEvents = False
    
    If Target.Row < 4 Then GoTo label1
    If Target.Column < 3 Then GoTo label1
    If Target.Row > 37 Then GoTo label1
    If Target.Column > 76 Then GoTo label1

    Target.Value = "."
    Cancel = True
    
label1:
    Application.EnableEvents = True
End Sub
