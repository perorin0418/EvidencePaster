VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EP_Class_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TEMPLATE_SHEET_NAME As String = "template"

Private m_Body As Collection

Private Sub Class_Initialize()
    
    Set m_Body = New Collection
    
    ' テンプレート範囲の取得（印刷範囲取得）
    Dim TemplateRange As Range
    Set TemplateRange = Range(ActiveSheet.PageSetup.PrintArea)
    
    ' エビデンスAのシェイプ名を取得
    Dim EvidenceAName As String
    EvidenceAName = GetShapeName("【EVIDENCE_A】")
    
    ' エビデンスAのセルを取得
    Dim EvidenceARange As Range
    Set EvidenceARange = Range(GetShapeRange(EvidenceAName))
    
    ' エビデンスAのセルに幅を書き込む
    EvidenceARange.Value = "【EVIDENCE_A】_" & GetShapeWidth(EvidenceAName)
    
    ' エビデンスBのシェイプ名を取得
    Dim EvidenceBName As String
    EvidenceBName = GetShapeName("【EVIDENCE_B】")
    
    ' エビデンスBのセルを取得
    Dim EvidenceBRange As Range
    Set EvidenceBRange = Range(GetShapeRange(EvidenceBName))
    
    ' エビデンスBのセルに幅を書き込む
    EvidenceBRange.Value = "【EVIDENCE_B】_" & GetShapeWidth(EvidenceBName)
    
    ' タグ範囲を走査してヘッダーとかに分ける
    Dim row As Integer
    For row = 1 To TemplateRange.Rows.Count
        m_Body.Add GetRangeRow2Variant(TemplateRange, row)
    Next
    
    ' 変数破棄
    Dispose TemplateRange
    Dispose EvidenceARange
    Dispose EvidenceBRange
    
End Sub

Private Sub Class_Terminate()
    ' NOP
End Sub

Public Sub Dispose(ByRef obj As Object)
    Set obj = Nothing
End Sub

Public Function GetTemplate() As Collection
    Set GetTemplate = m_Body
End Function

Public Function GetShapeName(text As String) As String
    Dim ShapeName() As String
    Dim shapeCnt As Integer
    Dim Cnt As Integer
    Dim Find As Boolean

    'templateシートのShapes数をカウント
    shapeCnt = ActiveWorkbook.Worksheets(TEMPLATE_SHEET_NAME).Shapes.Count
    '配列を再宣言
    ReDim ShapeName(shapeCnt)

    '配列shapeNameにオブジェクト名を代入
    For Cnt = 1 To shapeCnt
        ShapeName(Cnt) = ActiveWorkbook.Worksheets(TEMPLATE_SHEET_NAME).Shapes(Cnt).TextFrame.Characters.text
    Next Cnt

    Find = False

    '配列shapeNameに代入されたオブジェクト名を表示
    For Cnt = 1 To shapeCnt
        If ShapeName(Cnt) = text Then
            Find = True
            Exit For
        End If
    Next Cnt
    
    If Find = True Then
        GetShapeName = ActiveWorkbook.Worksheets(TEMPLATE_SHEET_NAME).Shapes(Cnt).Name
    Else
        GetShapeName = ""
    End If
    
    ' 変数解放
    Erase ShapeName
End Function

Public Function GetShapeRange(ShapeName As String) As String
    GetShapeRange = ActiveWorkbook.Worksheets(TEMPLATE_SHEET_NAME).Shapes(ShapeName).TopLeftCell.Address(False, False)
End Function

Public Function GetShapeWidth(ShapeName As String) As Integer
    GetShapeWidth = ActiveWorkbook.Worksheets(TEMPLATE_SHEET_NAME).Shapes(ShapeName).Width
End Function

Public Function GetRangeRow2Variant(rg As Range, row As Integer) As Variant
    Dim StrArray() As String
    ReDim StrArray(1 To rg.Columns.Count)
    Dim Cnt As Integer
    For Cnt = 1 To rg.Columns.Count
        StrArray(Cnt) = rg(row, Cnt).MergeArea.cells(1, 1).Value
    Next Cnt
    Dim Ret As Variant
    Ret = StrArray
    GetRangeRow2Variant = Ret
End Function
