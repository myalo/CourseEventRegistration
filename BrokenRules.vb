Public Class BrokenRules
   Private m_cBrokenRules As Collection
   Private m_LineSeparator As String
   Public Sub New()
      ' initialize object here
      m_cBrokenRules = New Collection
      m_LineSeparator = "<br />"
   End Sub

   Public Property BrokenRules() As Collection
      Get
         Return m_cBrokenRules
      End Get
      Set(ByVal value As Collection)
         m_cBrokenRules = value
      End Set
   End Property
   Public Property LineSeparator() As String
      Get
         Return m_LineSeparator
      End Get
      Set(ByVal value As String)
         m_LineSeparator = value
      End Set
   End Property

   Public Overrides Function ToString() As String
      Dim i As Integer
      Dim s As String = ""
      For i = 1 To m_cBrokenRules.Count
         s = s & m_cBrokenRules(i) & m_LineSeparator
      Next i
      ToString = s
   End Function
   Public Sub Add(ByVal Description As String)
      m_cBrokenRules.Add(Description)
   End Sub
   Public Sub Add(ByVal Description As String, ByVal Key As Integer)
      m_cBrokenRules.Add(Description, Key)
   End Sub
   Public Function IsValid() As Boolean
      Dim bResult As Boolean
      bResult = False
      If Not m_cBrokenRules Is Nothing Then
         If m_cBrokenRules.Count = 0 Then
            bResult = True
         End If
      End If
      IsValid = bResult

   End Function
End Class
