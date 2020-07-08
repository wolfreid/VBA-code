VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Shape_Create() As shape
   Set Shape_Create = TitleSheet.Shapes.AddShape(msoShapeBalloon, 350, 80, 275, 20)
    With Shape_Create
        .Adjustments(1) = 0.3
        .Adjustments(2) = 1
        .Fill.ForeColor.RGB = RGB(169, 209, 142)
        With .TextFrame
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
                .Characters.Font.name = "Calibri"
                .Characters.Font.size = 12
        End With
    End With
End Function


Public Sub Delete_Shape(ByVal sName As String)
Dim oShape As shape
Dim shps As Shapes: Set shps = TitleSheet.Shapes
For Each oShape In shps
    If oShape.name = sName Then oShape.Delete
Next
End Sub




