VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EP_class_TestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_TestSuite As Collection
Public m_EviNo As Range

Private Sub Class_Initialize()
    ' NOP
End Sub

Private Sub Class_Terminate()
    ' NOP
End Sub

Public Sub Init(ByVal Keyword1 As String, _
                ByVal Keyword2 As String, _
                ByVal Keyword3 As String, _
                ByVal Keyword4 As String, _
                ByVal Keyword5 As String, _
                ByVal Given As String, _
                ByVal When As String, _
                ByVal Zen As String, _
                ByVal eviNo As String)
    
    Dim buf_Keyword1 As Range
    Dim buf_Keyword2 As Range
    Dim buf_Keyword3 As Range
    Dim buf_Keyword4 As Range
    Dim buf_Keyword5 As Range
    Dim buf_Given As Range
    Dim buf_When As Range
    Dim buf_Zen As Range
    
    If Keyword1 <> "" Then
        Set buf_Keyword1 = Range(GetColSuite(Keyword1))
    End If
    
    If Keyword2 <> "" Then
        Set buf_Keyword2 = Range(GetColSuite(Keyword2))
    End If
    
    If Keyword3 <> "" Then
        Set buf_Keyword3 = Range(GetColSuite(Keyword3))
    End If
    
    If Keyword4 <> "" Then
        Set buf_Keyword4 = Range(GetColSuite(Keyword4))
    End If
    
    If Keyword5 <> "" Then
        Set buf_Keyword5 = Range(GetColSuite(Keyword5))
    End If
    
    Set buf_Given = Range(GetColSuite(Given))
    Set buf_When = Range(GetColSuite(When))
    Set buf_Zen = Range(GetColSuite(Zen))
    
    Set m_EviNo = Range(eviNo)
    
    ' 個数チェック
    Dim Cnt As Long
    Cnt = buf_Given.Count
    
    If Keyword1 <> "" Then
        If buf_Keyword1.Count <> Cnt Then
            MsgBox "テストスウィートの個数がおかしいゾ！"
            Exit Sub
        End If
    End If
    
    If Keyword2 <> "" Then
        If buf_Keyword2.Count <> Cnt Then
            MsgBox "テストスウィートの個数がおかしいゾ！"
            Exit Sub
        End If
    End If
    
    If Keyword3 <> "" Then
        If buf_Keyword3.Count <> Cnt Then
            MsgBox "テストスウィートの個数がおかしいゾ！"
            Exit Sub
        End If
    End If
    
    If Keyword4 <> "" Then
        If buf_Keyword4.Count <> Cnt Then
            MsgBox "テストスウィートの個数がおかしいゾ！"
            Exit Sub
        End If
    End If
    
    If Keyword5 <> "" Then
        If buf_Keyword5.Count <> Cnt Then
            MsgBox "テストスウィートの個数がおかしいゾ！"
            Exit Sub
        End If
    End If
    
    If buf_Given.Count <> Cnt Then
        MsgBox "テストスウィートの個数がおかしいゾ！"
        Exit Sub
    End If
    
    If buf_When.Count <> Cnt Then
        MsgBox "テストスウィートの個数がおかしいゾ！"
        Exit Sub
    End If
    
    If buf_Zen.Count <> Cnt Then
        MsgBox "テストスウィートの個数がおかしいゾ！"
        Exit Sub
    End If
    
    Set m_TestSuite = New Collection
    Dim TestRec As Collection
    Set TestRec = New Collection
    
    Dim i As Long
    For i = 1 To Cnt
        Set TestRec = New Collection
        
        If Keyword1 <> "" Then
            TestRec.Add buf_Keyword1(i).MergeArea.cells(1, 1).Value, "Keyword1"
        End If
        
        If Keyword2 <> "" Then
            TestRec.Add buf_Keyword2(i).MergeArea.cells(1, 1).Value, "Keyword2"
        End If
        
        If Keyword3 <> "" Then
            TestRec.Add buf_Keyword3(i).MergeArea.cells(1, 1).Value, "Keyword3"
        End If
        
        If Keyword4 <> "" Then
            TestRec.Add buf_Keyword4(i).MergeArea.cells(1, 1).Value, "Keyword4"
        End If
        
        If Keyword5 <> "" Then
            TestRec.Add buf_Keyword5(i).MergeArea.cells(1, 1).Value, "Keyword5"
        End If
        
        TestRec.Add buf_Given(1).MergeArea.cells(1, 1).Value, "Given"
        TestRec.Add buf_When(1).MergeArea.cells(1, 1).Value, "When"
        TestRec.Add buf_Zen(1).MergeArea.cells(1, 1).Value, "Then"
        
        m_TestSuite.Add TestRec
    Next i
End Sub

Public Function GetTestSuite() As Collection
    Set GetTestSuite = m_TestSuite
End Function

Public Function GetEvidenceNo() As Range
    Set GetEvidenceNo = m_EviNo
End Function

Private Function GetColSuite(Lead As String) As String
    
    Dim LeadRange As Range
    Set LeadRange = Range(Range(Lead).MergeArea.Item(1).Address(0, 0))
    
    Dim i As Long
    For i = 1 To 1000000
        If LeadRange.Offset(i, 0).MergeArea.cells(1, 1).Value = "" Then
            Exit For
        End If
    Next i
    
    Dim TailRange As Range
    Set TailRange = LeadRange.Offset(i - 1, 0)
    
    GetColSuite = LeadRange.Address(0, 0) & ":" & TailRange.Address(0, 0)
End Function
