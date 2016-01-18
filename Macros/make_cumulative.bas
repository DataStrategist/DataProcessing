Attribute VB_Name = "make_cumulative"
' make_cumulative() - Used to "accumulate" values along a specified axis of a table
'     in order to convert increments into the cumulative amount.
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

Sub make_cumulative()
Direkshun = MsgBox(" ------->   YES" & Chr(13) & "|" & Chr(13) & "|" & Chr(13) & "|" & Chr(13) & "\/" & Chr(13) & "NO", vbYesNoCancel)
If Direkshun = 2 Then Exit Sub
Set Start_point = Application.InputBox(Prompt:="Start point?", Type:=8)

R_n = InputBox("What row number has the Titles?", , Start_point.Row - 1)
C_n = InputBox("What column number has the Titles?  A=1, B=2, etc", , Start_point.Column - 1)


Start_point.Activate

If Direkshun = 6 Then ' GOING RIGHT
    While ActiveCell.Offset(0, C_n - ActiveCell.Column).Value <> ""
        While ActiveCell.Offset(R_n - ActiveCell.Row, 0).Value <> ""
            aa = ActiveCell.Value + aa
            ActiveCell.Formula = aa
            ActiveCell.Offset(0, 1).Activate
        Wend
        aa = 0
        ActiveCell.Offset(1, C_n + 1 - ActiveCell.Column).Activate
    Wend
ElseIf Direkshun = 7 Then ' GOING DOWN
    While ActiveCell.Offset(R_n - ActiveCell.Row, 0).Value <> ""
        While ActiveCell.Offset(0, C_n - ActiveCell.Column).Value <> ""
    
            aa = ActiveCell.Value + aa
            ActiveCell.Formula = aa
            ActiveCell.Offset(1, 0).Activate
        Wend
        aa = 0
        ActiveCell.Offset(R_n + 1 - ActiveCell.Row, 1).Activate
    Wend
End If

End Sub

