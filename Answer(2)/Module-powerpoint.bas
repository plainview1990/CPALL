Attribute VB_Name = "Module1"
Sub AddText1()

Set myDocument = ActivePresentation.Slides(1)
myDocument.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
    left:=200, top:=5, width:=100, height:=20).TextFrame.TextRange.Text = "ข้อมูล 01"
    
End Sub

Sub AddText2()

Set myDocument = ActivePresentation.Slides(1)
myDocument.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
    left:=650, top:=5, width:=100, height:=20).TextFrame.TextRange.Text = "ข้อมูล 02"
    
End Sub

Sub AddText3()

Set myDocument = ActivePresentation.Slides(1)
myDocument.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
    left:=650, top:=235, width:=100, height:=20).TextFrame.TextRange.Text = "ข้อมูล 03"
    
End Sub


