Attribute VB_Name = "RTFModule"
Option Explicit

Private m_bFileMode As Boolean
Private m_lObj As Long
Private m_sText As String
Private m_lPos As Long
Private m_lLen As Long

Public Property Let RichEdit(ByVal edtThis As LiteRTFViewer)
   m_lObj = ObjPtr(edtThis)
End Property

Public Property Get RichEdit() As LiteRTFViewer
Dim rT As LiteRTFViewer
   If (m_lObj <> 0) Then
      CopyMemory rT, m_lObj, 4
      Set RichEdit = rT
      CopyMemory rT, 0&, 4
   End If
End Property

Public Property Let StreamText(ByRef sText As String)
    m_sText = sText
    m_lPos = 1
    m_lLen = Len(m_sText)
End Property

Public Function LoadCallBack( _
        ByVal dwCookie As Long, _
        ByVal lPtrPbBuff As Long, _
        ByVal cb As Long, _
        ByVal pcb As Long _
    ) As Long

Dim sBuf As String
Dim b() As Byte
Dim lLen As Long
Dim lRead As Long
    
    If (m_bFileMode) Then
        ReadFile dwCookie, ByVal lPtrPbBuff, cb, pcb, ByVal 0&
        CopyMemory lRead, ByVal pcb, 4
        If (lRead < cb) Then
            ' Complete:
            LoadCallBack = 0
        Else
            ' More to read:
            LoadCallBack = 0
        End If
        m_lPos = m_lPos + lRead
    
    Else
        CopyMemory lRead, ByVal pcb, 4
        
        ' Place cb bytes if possible, or place in the whole string:
        If (m_lLen - m_lPos >= 0) Then
            If (m_lLen - m_lPos < cb) Then
                ReDim b(0 To (m_lLen - m_lPos)) As Byte
                b = StrConv(Mid$(m_sText, m_lPos), vbFromUnicode)
                lRead = m_lLen - m_lPos + 1
                CopyMemory ByVal lPtrPbBuff, b(0), lRead
                m_lPos = m_lLen + 1
            Else
                ReDim b(0 To cb - 1) As Byte
                b = StrConv(Mid$(m_sText, m_lPos, cb), vbFromUnicode)
                CopyMemory ByVal lPtrPbBuff, b(0), cb
                m_lPos = m_lPos + cb
                lRead = cb
            End If
                        
            CopyMemory ByVal pcb, lRead, 4
            LoadCallBack = 0
        Else
            lRead = 0
            CopyMemory ByVal pcb, lRead, 4
            LoadCallBack = 0
        End If
        
    End If
End Function

