VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AddSerialNumbers()

Dim i As Integer

On Error GoTo Last

i = InputBox("Enter Value", "Enter Serial Numbers")

For i = 1 To i

ActiveCell.Value = i

ActiveCell.Offset(1, 0).Activate

Next i

Last: Exit Sub

End Sub
