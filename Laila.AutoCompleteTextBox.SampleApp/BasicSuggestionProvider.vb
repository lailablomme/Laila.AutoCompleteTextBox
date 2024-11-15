Public Class BasicSuggestionProvider
    Implements ISuggestionProviderAsync

    Public Async Function GetSuggestions(filter As String) As Task(Of IEnumerable) Implements ISuggestionProviderAsync.GetSuggestions
        Dim list As List(Of BasicData) = New List(Of BasicData) From
        {
            New BasicData() With {.Id = 1, .Name = "One"},
            New BasicData() With {.Id = 2, .Name = "Two"},
            New BasicData() With {.Id = 3, .Name = "Three"}
        }
        If Not String.IsNullOrWhiteSpace(filter) Then
            Select Case filter
                Case "GetById://1" : Return New List(Of BasicData) From {list(0)}
                Case "GetById://2" : Return New List(Of BasicData) From {list(1)}
                Case "GetById://3" : Return New List(Of BasicData) From {list(2)}
                Case Else : Return list
            End Select
        Else
            Return list
        End If
    End Function
End Class
