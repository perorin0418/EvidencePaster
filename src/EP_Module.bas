Attribute VB_Name = "EP_module"
Option Explicit

Public Const Format_Contents_Cell As String = "C5"
Public Const Keyword1_Contents_Cell As String = "C12"
Public Const Keyword2_Contents_Cell As String = "C15"
Public Const Keyword3_Contents_Cell As String = "C18"
Public Const Keyword4_Contents_Cell As String = "C21"
Public Const Keyword5_Contents_Cell As String = "C24"
Public Const Keyword1_Range_Cell As String = "R12"
Public Const Keyword2_Range_Cell As String = "R15"
Public Const Keyword3_Range_Cell As String = "R18"
Public Const Keyword4_Range_Cell As String = "R21"
Public Const Keyword5_Range_Cell As String = "R24"
Public Const Given_Range_Cell As String = "AB12"
Public Const When_Range_Cell As String = "AQ12"
Public Const Then_Range_Cell As String = "AB19"
Public Const EVIDENCE_A_Range_Cell As String = "J32"
Public Const EVIDENCE_B_Range_Cell As String = "J35"
Public Const INITPOS_Range_Cell As String = "AQ19"
Public Const EVIDENCE_NO_Range_Cell As String = "AB26"

Private Const TEMPLATE_SHEET_NAME As String = "template"
Private Const GIVEN_NAME = "【前提条件】"
Private Const WHEN_NAME = "【イベント】"
Private Const THEN_NAME = "【処理結果】"

Public CellRange As String
Public template As EP_Class_Template
Public TestSuite As EP_class_TestSuite
Public CurrentPasteCell As Range

Private Sub エビ貼りまくろー♪()
    
    EP_Waiting.Show vbModeless
    
    ' 描画停止
    Application.ScreenUpdating = False
    
    Load_Template
    
    Load_TestSuite
    
    Set CurrentPasteCell = Range(Range(INITPOS_Range_Cell).Value)
    
    PasteEvidence
        
    ' 描画再開
    Application.ScreenUpdating = True
    
    Unload EP_Waiting
    
End Sub

Private Sub Load_Template()
    ' テンプレートシートに移動
    ActiveWorkbook.Worksheets(TEMPLATE_SHEET_NAME).Activate
    
    Set template = New EP_Class_Template
    
    ' エビペシートに移動
    ActiveWorkbook.Worksheets("エビ貼りまくろー♪").Activate
End Sub

Private Sub Load_TestSuite()
    Set TestSuite = New EP_class_TestSuite
    
    Dim Keyword1 As String
    Keyword1 = Range(Keyword1_Range_Cell).Value
    Dim Keyword2 As String
    Keyword2 = Range(Keyword2_Range_Cell).Value
    Dim Keyword3 As String
    Keyword3 = Range(Keyword3_Range_Cell).Value
    Dim Keyword4 As String
    Keyword4 = Range(Keyword4_Range_Cell).Value
    Dim Keyword5 As String
    Keyword5 = Range(Keyword5_Range_Cell).Value
    Dim Given As String
    Given = Range(Given_Range_Cell).Value
    Dim When As String
    When = Range(When_Range_Cell).Value
    Dim Zen As String
    Zen = Range(Then_Range_Cell).Value
    Dim eviNo As String
    eviNo = Range(EVIDENCE_NO_Range_Cell).Value
    
    ' テストスウィートシートに移動
    ActiveWorkbook.Worksheets(Range(Range(Given_Range_Cell).Value).Parent.Name).Activate
    
    TestSuite.Init Keyword1, Keyword2, Keyword3, Keyword4, Keyword5, Given, When, Zen, eviNo
    
    ' エビペシートに移動
    ActiveWorkbook.Worksheets("エビ貼りまくろー♪").Activate
End Sub

Private Sub PasteEvidence()
    
    Dim format As String
    
    Dim suite As Collection
    Set suite = TestSuite.GetTestSuite
    Dim Cnt As Long
    For Cnt = 1 To suite.Count
        Dim test As Collection
        Set test = suite(Cnt)
        PasteBDD test, "Given", Cnt
        PasteBDD test, "When", Cnt
        PasteBDD test, "Then", Cnt
    Next Cnt
End Sub

Private Sub PasteBDD(test As Collection, BDD As String, Cnt As Long)
    Dim format As String
    format = Range(Format_Contents_Cell).Value
        
    Dim Keyword1_Contents As String
    Keyword1_Contents = Range(Keyword1_Contents_Cell).Value
    If Keyword1_Contents <> "" Then
        format = Replace(format, Keyword1_Contents, test("Keyword1"))
    End If
        
    Dim Keyword2_Contents As String
    Keyword2_Contents = Range(Keyword2_Contents_Cell).Value
    If Keyword2_Contents <> "" Then
        format = Replace(format, Keyword2_Contents, test("Keyword2"))
    End If
        
    Dim Keyword3_Contents As String
    Keyword3_Contents = Range(Keyword3_Contents_Cell).Value
    If Keyword3_Contents <> "" Then
        format = Replace(format, Keyword3_Contents, test("Keyword3"))
    End If
        
    Dim Keyword4_Contents As String
    Keyword4_Contents = Range(Keyword4_Contents_Cell).Value
    If Keyword4_Contents <> "" Then
        format = Replace(format, Keyword4_Contents, test("Keyword4"))
    End If
        
    Dim Keyword5_Contents As String
    Keyword5_Contents = Range(Keyword5_Contents_Cell).Value
    If Keyword5_Contents <> "" Then
        format = Replace(format, Keyword5_Contents, test("Keyword5"))
    End If
    
    format = Replace(format, "[No]", "*")
    format = Replace(format, "[BDD]", BDD)
        
    Paste format, test, BDD, Cnt
End Sub

Private Sub Paste(image As String, test As Collection, BDD As String, Cnt As Long)
    
    Dim buf_Count1 As Long
    Dim buf_Count2 As Long
    Dim cells() As String
    
    Dim eviApath As String
    eviApath = Range(EVIDENCE_A_Range_Cell).Value
    Dim eviBpath As String
    eviBpath = Range(EVIDENCE_B_Range_Cell).Value
    
    Dim images As Collection
    Set images = New Collection
    Dim buf_image As String
    On Error Resume Next
    buf_image = Dir(eviApath & "\" & image)
    Do While buf_image <> ""
        images.Add buf_image, buf_image
        buf_image = Dir()
    Loop
    buf_image = Dir(eviBpath & "\" & image)
    Do While buf_image <> ""
        images.Add buf_image, buf_image
        buf_image = Dir()
    Loop
    On Error GoTo 0
    
    Dim templateCollection As Collection
    Set templateCollection = template.GetTemplate
    
    Dim imageA_width As Long
    imageA_width = 0
    For buf_Count1 = 1 To templateCollection.Count
        cells = templateCollection(buf_Count1)
        For buf_Count2 = 1 To UBound(templateCollection(buf_Count1)) - LBound(templateCollection(buf_Count1)) + 1
            If cells(buf_Count2) Like "【EVIDENCE_A】_*" Then
                If Val(Split(cells(buf_Count2), "】_")(1)) > imageA_width Then
                    imageA_width = Split(cells(buf_Count2), "】_")(1)
                    GoTo ForBreakA
                End If
            End If
        Next buf_Count2
    Next buf_Count1
ForBreakA:
    Dim imageB_width As Long
    imageB_width = 0
    For buf_Count1 = 1 To templateCollection.Count
        cells = templateCollection(buf_Count1)
        For buf_Count2 = 1 To UBound(templateCollection(buf_Count1)) - LBound(templateCollection(buf_Count1)) + 1
            If cells(buf_Count2) Like "【EVIDENCE_B】_*" Then
                If Val(Split(cells(buf_Count2), "】_")(1)) > imageB_width Then
                    imageB_width = Split(cells(buf_Count2), "】_")(1)
                    GoTo ForBreakB
                End If
            End If
        Next buf_Count2
    Next buf_Count1
ForBreakB:
    
    ' エビデンスシートに移動
    ActiveWorkbook.Worksheets(Range(Range(INITPOS_Range_Cell).Value).Parent.Name).Activate
    
    Dim collCount As Long
    collCount = templateCollection.Count
    
    Dim image_count As Long
    For image_count = 1 To images.Count
        For buf_Count1 = 1 To collCount
            cells = templateCollection(buf_Count1)
            Dim moveCell As Long
            moveCell = 1
            Dim moveCellForImage As Long
            moveCellForImage = 0
            For buf_Count2 = 1 To UBound(templateCollection(buf_Count1))
                Dim cellStr As String
                cellStr = cells(buf_Count2)
                If cellStr <> "" Then
                    If cellStr = "【EVIDENCE_NO】" Then
                        CurrentPasteCell.Offset(buf_Count1 - 2, buf_Count2 - 1).Value = RemoveExstension(images(image_count))
                        
                        ' エビペシートに移動
                        ActiveWorkbook.Worksheets("エビ貼りまくろー♪").Activate
                        
                        ' テストスウィートシートに移動
                        ActiveWorkbook.Worksheets(Range(Range(EVIDENCE_NO_Range_Cell).Value).Parent.Name).Activate
                        
                        Dim eviNo As Range
                        Set eviNo = TestSuite.GetEvidenceNo
                        If eviNo(Cnt) = "" Then
                            eviNo(Cnt) = RemoveExstension(images(image_count))
                        Else
                            eviNo(Cnt) = eviNo(Cnt) & vbLf & RemoveExstension(images(image_count))
                        End If
                        
                        ' エビペシートに移動
                        ActiveWorkbook.Worksheets("エビ貼りまくろー♪").Activate
                        
                        ' エビデンスシートに移動
                        ActiveWorkbook.Worksheets(Range(Range(INITPOS_Range_Cell).Value).Parent.Name).Activate
                        
                    End If
                    
                    If cellStr = "【BDD_SECTION_NAME】" Then
                        If BDD = "Given" Then
                            CurrentPasteCell.Offset(buf_Count1 - 2, buf_Count2 - 1).Value = GIVEN_NAME
                        ElseIf BDD = "When" Then
                            CurrentPasteCell.Offset(buf_Count1 - 2, buf_Count2 - 1).Value = WHEN_NAME
                        ElseIf BDD = "Then" Then
                            CurrentPasteCell.Offset(buf_Count1 - 2, buf_Count2 - 1).Value = THEN_NAME
                        End If
                    End If
                    
                    If cellStr = "【BDD_SECTION_CONTENTS】" Then
                        Dim strs() As String
                        strs = Split(test(BDD), vbLf)
                        Dim str As Variant
                        For Each str In strs
                            CurrentPasteCell.Offset(buf_Count1 - 4 + moveCell, buf_Count2 - 1).Value = str
                            moveCell = moveCell + 1
                        Next str
                        moveCell = moveCell + 1
                    End If
                    
                    If cellStr Like "【EVIDENCE_*】_*" Then
                        Dim myShape As Shape
                        Dim modScale As Double
                        Dim cellHeight As Long
                        CurrentPasteCell.Offset(buf_Count1 - 6, buf_Count2 - 1).Activate
                        If cellStr Like "【EVIDENCE_A】_*" Then
                            If Dir(eviApath & "\" & images(image_count)) <> "" Then
                                Set myShape = ActiveSheet.Shapes.AddPicture( _
                                    fileName:=eviApath & "\" & images(image_count), _
                                    LinkToFile:=False, _
                                    SaveWithDocument:=True, _
                                    Left:=Selection.Left, _
                                    Top:=Selection.Top, _
                                    Width:=0, _
                                    Height:=0)
                                With myShape
                                    .ScaleHeight 1, msoTrue
                                    .ScaleWidth 1, msoTrue
                                End With
                                
                                modScale = imageA_width / myShape.Width
                                
                                myShape.Width = myShape.Width * modScale
                                myShape.Height = myShape.Height * modScale
                                
                                If myShape.Height / CurrentPasteCell.Height > moveCellForImage Then
                                    moveCellForImage = myShape.Height / CurrentPasteCell.Height - 4
                                End If
                            End If
                        ElseIf cellStr Like "【EVIDENCE_B】_*" Then
                            If Dir(eviBpath & "\" & images(image_count)) <> "" Then
                                Set myShape = ActiveSheet.Shapes.AddPicture( _
                                    fileName:=eviBpath & "\" & images(image_count), _
                                    LinkToFile:=False, _
                                    SaveWithDocument:=True, _
                                    Left:=Selection.Left, _
                                    Top:=Selection.Top, _
                                    Width:=0, _
                                    Height:=0)
                                With myShape
                                    .ScaleHeight 1, msoTrue
                                    .ScaleWidth 1, msoTrue
                                End With
                                
                                modScale = imageA_width / myShape.Width
                                
                                myShape.Width = myShape.Width * modScale
                                myShape.Height = myShape.Height * modScale
                                
                                If myShape.Height / CurrentPasteCell.Height > moveCellForImage Then
                                    moveCellForImage = myShape.Height / CurrentPasteCell.Height - 4
                                End If
                            End If
                        End If
                        Set myShape = Nothing
                    End If
                End If
            Next buf_Count2
            Set CurrentPasteCell = CurrentPasteCell.Offset(moveCell + moveCellForImage, 0)
        Next buf_Count1
    Next image_count
    
    ' エビペシートに移動
    ActiveWorkbook.Worksheets("エビ貼りまくろー♪").Activate
    
End Sub

Private Function RemoveExstension(fileName As String) As String
    RemoveExstension = Split(fileName, ".")(0)
End Function

Private Sub キーワード選択1()
    EP_RefEdit.Show
    Dim r As Range
    Set r = Range(Keyword1_Range_Cell)
    r.Value = CellRange
    Set r = Nothing
End Sub

Private Sub キーワード選択2()
    EP_RefEdit.Show
    Dim r As Range
    Set r = Range(Keyword2_Range_Cell)
    r.Value = CellRange
    Set r = Nothing
End Sub

Private Sub キーワード選択3()
    EP_RefEdit.Show
    Dim r As Range
    Set r = Range(Keyword3_Range_Cell)
    r.Value = CellRange
    Set r = Nothing
End Sub

Private Sub キーワード選択4()
    EP_RefEdit.Show
    Dim r As Range
    Set r = Range(Keyword4_Range_Cell)
    r.Value = CellRange
    Set r = Nothing
End Sub

Private Sub キーワード選択5()
    EP_RefEdit.Show
    Dim r As Range
    Set r = Range(Keyword5_Range_Cell)
    r.Value = CellRange
    Set r = Nothing
End Sub

Private Sub 前提条件選択()
    EP_RefEdit.Show
    Dim r As Range
    Set r = Range(Given_Range_Cell)
    r.Value = CellRange
    Set r = Nothing
End Sub

Private Sub イベント選択()
    EP_RefEdit.Show
    Dim r As Range
    Set r = Range(When_Range_Cell)
    r.Value = CellRange
    Set r = Nothing
End Sub

Private Sub 処理結果選択()
    EP_RefEdit.Show
    Dim r As Range
    Set r = Range(Then_Range_Cell)
    r.Value = CellRange
    Set r = Nothing
End Sub

Private Sub エビデンス№選択()
    EP_RefEdit.Show
    Dim r As Range
    Set r = Range(EVIDENCE_NO_Range_Cell)
    r.Value = CellRange
    Set r = Nothing
End Sub

Private Sub エビデンスA選択()
    Dim r As Range
    Set r = Range(EVIDENCE_A_Range_Cell)
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            r.Value = .SelectedItems(1)
        End If
    End With
    Set r = Nothing
End Sub

Private Sub エビデンスB選択()
    Dim r As Range
    Set r = Range(EVIDENCE_B_Range_Cell)
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            r.Value = .SelectedItems(1)
        End If
    End With
    Set r = Nothing
End Sub

Private Sub 初期位置選択()
    EP_RefEdit.Show
    Dim r As Range
    Set r = Range(INITPOS_Range_Cell)
    r.Value = CellRange
    Set r = Nothing
End Sub

