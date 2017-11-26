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
    
    ' �e���v���[�g�͈͂̎擾�i����͈͎擾�j
    Dim TemplateRange As Range
    Set TemplateRange = Range(ActiveSheet.PageSetup.PrintArea)
    
    ' �G�r�f���XA�̃V�F�C�v�����擾
    Dim EvidenceAName As String
    EvidenceAName = GetShapeName("�yEVIDENCE_A�z")
    
    ' �G�r�f���XA�̃Z�����擾
    Dim EvidenceARange As Range
    Set EvidenceARange = Range(GetShapeRange(EvidenceAName))
    
    ' �G�r�f���XA�̃Z���ɕ�����������
    EvidenceARange.Value = "�yEVIDENCE_A�z_" & GetShapeWidth(EvidenceAName)
    
    ' �G�r�f���XB�̃V�F�C�v�����擾
    Dim EvidenceBName As String
    EvidenceBName = GetShapeName("�yEVIDENCE_B�z")
    
    ' �G�r�f���XB�̃Z�����擾
    Dim EvidenceBRange As Range
    Set EvidenceBRange = Range(GetShapeRange(EvidenceBName))
    
    ' �G�r�f���XB�̃Z���ɕ�����������
    EvidenceBRange.Value = "�yEVIDENCE_B�z_" & GetShapeWidth(EvidenceBName)
    
    ' �^�O�͈͂𑖍����ăw�b�_�[�Ƃ��ɕ�����
    Dim row As Integer
    For row = 1 To TemplateRange.Rows.Count
        m_Body.Add GetRangeRow2Variant(TemplateRange, row)
    Next
    
    ' �ϐ��j��
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

    'template�V�[�g��Shapes�����J�E���g
    shapeCnt = ActiveWorkbook.Worksheets(TEMPLATE_SHEET_NAME).Shapes.Count
    '�z����Đ錾
    ReDim ShapeName(shapeCnt)

    '�z��shapeName�ɃI�u�W�F�N�g������
    For Cnt = 1 To shapeCnt
        ShapeName(Cnt) = ActiveWorkbook.Worksheets(TEMPLATE_SHEET_NAME).Shapes(Cnt).TextFrame.Characters.text
    Next Cnt

    Find = False

    '�z��shapeName�ɑ�����ꂽ�I�u�W�F�N�g����\��
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
    
    ' �ϐ����
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