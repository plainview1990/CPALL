Attribute VB_Name = "Module1"
Sub ConslidateWorkbooks_1()
Dim FolderPath As String
Dim Filename As String
Dim Sheet As Worksheet

Application.ScreenUpdating = False
FolderPath = "C:\Data Science\CPALL\"  'input your local folder path
Filename = Dir(FolderPath & "*.xls*")

Do While Filename <> ""
 Workbooks.Open Filename:=FolderPath & Filename, ReadOnly:=True
 
 For Each Sheet In ActiveWorkbook.Sheets
 Sheet.Copy After:=ThisWorkbook.Sheets(1)
 
 Next Sheet
 
 Workbooks(Filename).Close
 Filename = Dir()
 
Loop

Application.ScreenUpdating = True

End Sub

Sub DeleteWorksheet_2()

Application.DisplayAlerts = False
Sheets("Sheet1").delete             'input your worksheet name to delete
Application.DisplayAlerts = True

End Sub

Sub SortWorksheets_3()
 Dim i As Integer
 Dim j As Integer
 Dim iAnswer As VbMsgBoxResult

    iAnswer = MsgBox("Sort Sheets in Ascending Order?" & Chr(10) _
      & "Clicking No will sort in Descending Order", _
      vbYesNoCancel + vbQuestion + vbDefaultButton1, "Sort Worksheets")
    For i = 1 To Sheets.Count
       For j = 1 To Sheets.Count - 1

          If iAnswer = vbYes Then
             If UCase$(Sheets(j).Name) > UCase$(Sheets(j + 1).Name) Then
                Sheets(j).Move After:=Sheets(j + 1)
             End If

          ElseIf iAnswer = vbNo Then
             If UCase$(Sheets(j).Name) < UCase$(Sheets(j + 1).Name) Then
                Sheets(j).Move After:=Sheets(j + 1)

             End If
          End If

       Next j
    Next i

 End Sub

Sub CopyMultiObject_4()

'Before run macro CopyMultiObject_4 _
VBA : Tools > References > Checked : Microsoft PowerPoint 16.0 Object Library > OK

Dim ppt As PowerPoint.Application
Dim pres As PowerPoint.Presentation
Dim sl As PowerPoint.Slide
Dim shp As PowerPoint.Shape

Dim objarray, slarray As Variant
Dim left, top, height, width As Variant
Dim x As Long

Set ppt = New PowerPoint.Application
Set pres = ppt.Presentations.Add
Set sl = pres.Slides.Add(1, ppLayoutBlank)

slarray = Array(1, 2, 3)
objarray = Array(Sheet2.Range("A1").CurrentRegion, Sheet3.Range("A1").CurrentRegion, Sheet4.Range("A1").CurrentRegion)
'check your worksheet index mapping worksheet name

left = Array(50, 480, 480)
top = Array(35, 35, 270)

For x = LBound(slarray) To UBound(slarray)
    objarray(x).Copy
    
    sl.Shapes.PasteSpecial DataType:=ppPasteOLEObject, link:=msoTrue
    Set shp = sl.Shapes(sl.Shapes.Count)
       
    With shp
    .left = left(x)
    .top = top(x)
    End With

Next x

End Sub

Sub CopySingleObject()

Dim ppt As PowerPoint.Application
Dim pres As PowerPoint.Presentation
Dim sl As PowerPoint.Slide
Dim cl As PowerPoint.CustomLayout
Dim shp As PowerPoint.ShapeRange
Dim MyRangeArray As Variant

Set ppt = New PowerPoint.Application
Set pres = ppt.Presentations.Add
Set cl = pres.SlideMaster.CustomLayouts(7)
Set sl = pres.Slides.AddSlide(1, cl)

Range("A1").CurrentRegion.Copy

Set shp = sl.Shapes.PasteSpecial(DataType:=ppPasteOLEObject, link:=msoTrue)
shp(1).top = 50
shp(1).left = 50

End Sub

