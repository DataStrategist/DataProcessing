VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dimChanger 
   Caption         =   "Dimension Changer 2000!!!!!!"
   ClientHeight    =   3285
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   3780
   OleObjectBlob   =   "dimChanger.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dimChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' dimChanger - A way to transform data between list and table format
'     Copyright (C) 2015 Amit Kohli <amit *at* amitkohli.com>
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'   the Free Software Foundation; either version 2 of the License, or
'     (at your option) any later version.
'
'     This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'     GNU General Public License for more details (http://www.gnu.org/licenses/gpl-2.0.txt)
'
'    You should have received a copy of the GNU General Public License along
'   with this program; if not, write to the Free Software Foundation, Inc.,
'   51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
'
'
' For examples, and to obtain new versions of the file, go to:  https://github.com/mexindian/excel.DataProcessingMacros

Option Explicit

Private Sub CommandButton1_Click()

'ok button

'-------------DIMs
Dim rrange1, rrange2, datastarts, x, Y1, targg As Range
Dim i, i_ctr As Integer
Dim r, c As Variant
Dim cmt As Comment
Dim fixxed_cmt As String
Dim arr(99999, 5)
Dim emptyWarning
Dim Xcol As Integer


'-------------ERRORS
If Me.OB_Table_to_List.Value = False And Me.OB_List_to_table.Value = False Then
    MsgBox ("Please select what I should do with your data")
    Exit Sub
End If

'-------------PICK DIMENSIONS
Me.Hide
On Error Resume Next
    Application.DisplayAlerts = False
    If Me.OB_List_to_table Then 'Dim 1, ROW headings
        Set rrange1 = Application.InputBox(Prompt:="Please select the Dimension that will become ROW HEADINGS", Title:="SPECIFY DIM 1", Type:=8)
    Else
        Set rrange1 = Application.InputBox(Prompt:="Please select the ROW HEADINGS", Title:="SPECIFY DIM 1", Type:=8)
    End If
    
If rrange1 Is Nothing Then Exit Sub
    
    If Me.OB_List_to_table Then 'Dim 2, COLUMN headings
        Set rrange2 = Application.InputBox(Prompt:="Please select the Dimension that will become COLUMN HEADINGS", Title:="SPECIFY DIM 2", Type:=8)
    Else
        Set rrange2 = Application.InputBox(Prompt:="Please select the COLUMN HEADINGS", Title:="SPECIFY DIM 2", Type:=8)
    End If
    
On Error GoTo 0
If rrange2 Is Nothing Then Exit Sub

Set datastarts = Application.InputBox(Prompt:="Please select first data-point.", Title:="SPECIFY DIM 2", Type:=8) 'First data point

Application.DisplayAlerts = True

If datastarts Is Nothing Then Exit Sub

If rrange1.Cells(1, 1).Column = datastarts.Column Then
    Set x = rrange1
    Set Y1 = rrange2
Else
    Set x = rrange2
    Set Y1 = rrange1
End If

If Me.CB_formatting Then
        'In comments, replace line breaks with unique character ƒ, and "  with '.  (Just cleaning up for later)
        For Each cmt In ActiveSheet.Comments
            'fixxed_cmt = Replace(cmt.Text, Chr(10), "ƒ")
            'fixxed_cmt = Replace(cmt.Text, Chr(13), "ƒ")
            fixxed_cmt = Replace(cmt.Text, """", "'")
            cmt.Delete
            cmt.Parent.AddComment Text:=fixxed_cmt
        Next
    End If

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=START!    ARR 0=Row counter | 1=Column counter | 2=Value | 3=Cell Color | 4=Font Color | 5=Comment

i = 0

If Me.OB_Table_to_List Then  '================================================================================== TABLE ------> LIST HERE
    datastarts.Activate
    
    For Each r In Y1
        For Each c In x
            Range("A1").Offset(r.Row - 1, c.Column - 1).Activate 'debug
            arr(i, 0) = r
            arr(i, 1) = c
            arr(i, 2) = Range("A1").Offset(r.Row - 1, c.Column - 1).Formula
            If Me.CB_formatting Then
                arr(i, 3) = Range("A1").Offset(r.Row - 1, c.Column - 1).Interior.Color
                arr(i, 4) = Range("A1").Offset(r.Row - 1, c.Column - 1).Font.Color
                On Error Resume Next
                    arr(i, 5) = Range("A1").Offset(r.Row - 1, c.Column - 1).Comment.Text
                On Error GoTo 0
            End If
                
            i = i + 1
        Next
    Next
    
    '====OK, done, now spitting out results
    If Me.OptionButton1 Then
        Workbooks.Add
        Range("B2").Activate
    Else
        Windows(Me.ComboBox1.Value).Activate
        Set targg = Application.InputBox(Prompt:="Select destination", Title:="SPECIFY DIM 1", Type:=8)
        targg.Activate
    End If
    Application.ScreenUpdating = False
    
    For i_ctr = 0 To i - 1
        If Len(arr(i_ctr, 2)) <> 0 Or Me.CB_Blanks Then 'if cell isn't empty or if u want blanks
            ActiveCell.Offset(0, 0).Value = arr(i_ctr, 0)
            ActiveCell.Offset(0, 1).Value = arr(i_ctr, 1)
            ActiveCell.Offset(0, 2).Value = arr(i_ctr, 2)
            If Me.CB_formatting Then
                ActiveCell.Offset(0, 2).Interior.Color = arr(i_ctr, 3)
                ActiveCell.Offset(0, 2).Font.Color = arr(i_ctr, 4)
                If Len(arr(i_ctr, 5)) <> 0 Then
                    ActiveCell.Offset(0, 2).NoteText arr(i_ctr, 5)
                End If
            End If
            ActiveCell.Offset(1, 0).Activate
        End If
        Application.StatusBar = Round(i_ctr / i * 100, 1) & "%"
    Next


Else '=============================================================== LIST ------> TABLE HERE

    For Each c In rrange1
        datastarts.Offset(i, 0).Activate
        
        arr(i, 0) = c.Value
        arr(i, 1) = rrange2.Cells(i + 1, 1).Value
        arr(i, 2) = datastarts.Offset(i, 0).Formula
        If Not WorksheetFunction.IsNumber(datastarts.Offset(i, 0).Value) And datastarts.Offset(i, 0).Value <> "" Then
            MsgBox ("This isn't a number, move to comment or delete or something, but it can't stay here")
            Exit Sub
        End If
        If Me.CB_formatting Then
            arr(i, 3) = datastarts.Offset(i, 0).Interior.Color
            arr(i, 4) = datastarts.Offset(i, 0).Font.Color
            On Error Resume Next
                arr(i, 5) = datastarts.Offset(i, 0).Comment.Text
            On Error GoTo 0
        End If
        
        i = i + 1
    Next
        
    '====OK, done, now spitting out results
    
    If Me.OptionButton1 Then
        Workbooks.Add
        Range("B2").Activate
    Else
        Windows(Me.ComboBox1.Value).Activate
    End If
    Application.ScreenUpdating = False
       
    ' First, create row and column headings for all "slots"
    '' 0   rowheading
    '' 1 is columnheading
    
    Dim d As Object
    Dim ii As Long
    Dim v As Variant
    
    
    Set d = CreateObject("Scripting.Dictionary")
    'Set d = New Scripting.Dictionary
    
    For ii = LBound(arr) To UBound(arr)
        d(arr(ii, 1)) = 1
    Next ii
    
    Range("c1").Activate
    For Each v In d.Keys()
        'd.Keys() is a Variant array of the unique values in myArray.
        ActiveCell.Value = v
        ActiveCell.Offset(0, 1).Activate
    Next v
    
    Set d = New Scripting.Dictionary
    
    For ii = LBound(arr) To UBound(arr)
        d(arr(ii, 0)) = 1
    Next ii
    
    Range("a2").Activate
    For Each v In d.Keys()
        'd.Keys() is a Variant array of the unique values in myArray.
        ActiveCell.Value = v
        ActiveCell.Offset(1, 0).Activate
    Next v
    
    For i_ctr = 0 To i - 1
        Range("1:1").Find(arr(i_ctr, 1)).Activate
        Xcol = ActiveCell.Column
        Range("A:A").Find(arr(i_ctr, 0)).Activate
        ActiveCell.Offset(0, Xcol - 1).Activate
    
        If Len(arr(i_ctr, 2)) <> 0 Or Me.CB_Blanks Then 'if cell isn't empty or if u want blanks
            'Found point! Putting data
            If ActiveCell.Value <> "" Then
                If emptyWarning = "" Then emptyWarning = MsgBox("huh... so there are already results for this. Should I overwrite cells or add em up? " & Chr(13) & "Say YES  for overwrite" & Chr(13) & "Say NO for adding up all instances of that point", vbYesNo)
                If emptyWarning = 6 Then GoTo overwriteAnyway
                    If Left(arr(i_ctr, 2), 1) = "=" Then
                        ActiveCell.Formula = Replace("=" & ActiveCell.Formula & Replace(arr(i_ctr, 2), "=", "+"), "==", "=")
                    Else
                        ActiveCell.Value = ActiveCell.Value + CDbl(arr(i_ctr, 2))
                    End If
            Else
overwriteAnyway:
                ActiveCell.Value = arr(i_ctr, 2)
            End If
            If Me.CB_formatting Then
                ActiveCell.Interior.Color = arr(i_ctr, 3)
                ActiveCell.Font.Color = arr(i_ctr, 4)
                If Len(arr(i_ctr, 5)) <> 0 Then
                    ActiveCell.NoteText arr(i_ctr, 5)
                End If
            End If
            Application.StatusBar = "Step 3  " & Round(i_ctr / i * 100, 1) & "%"
        End If
    Next
    
End If

Application.ScreenUpdating = True
Application.StatusBar = ""

End Sub

Private Sub CommandButton2_Click()
'Cancel
Unload dimChanger
End Sub

Private Sub OB_List_to_table_Click()
If OptionButton2 Then Me.Label1.Visible = True
End Sub

Private Sub OB_Table_to_List_Click()
Me.Label1.Visible = False
End Sub

Private Sub OptionButton1_Click()
Me.ComboBox1.Visible = False
Me.Label1.Visible = False
End Sub

Private Sub OptionButton2_Click()
Dim wkb As Workbook
Me.ComboBox1.Visible = True
With Me.ComboBox1
    For Each wkb In Application.Workbooks
        .AddItem wkb.Name
    Next wkb
End With
If OB_List_to_table Then Me.Label1.Visible = True
End Sub
