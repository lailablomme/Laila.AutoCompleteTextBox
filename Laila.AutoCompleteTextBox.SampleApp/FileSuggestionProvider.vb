Imports System.IO

Public Class FileSuggestionProvider
    Implements ISuggestionProvider

    Public Function GetSuggestions(filter As String) As IEnumerable Implements ISuggestionProvider.GetSuggestions
        If filter = "test" Then
            Throw New Exception("test error")
        ElseIf filter.StartsWith("GetById://") Then
            Return {New FileData() With {.FullPath = filter.Substring("GetById://".Length)}}
        Else
            Try
                Dim path As String = filter
                Dim pathWithoutFileName As String
                If path.Contains(IO.Path.DirectorySeparatorChar) Then
                    pathWithoutFileName = path.Substring(0, path.LastIndexOf(IO.Path.DirectorySeparatorChar))
                Else
                    pathWithoutFileName = path
                End If

                Dim files As List(Of FileData) =
                    Directory.GetFiles(pathWithoutFileName + IO.Path.DirectorySeparatorChar, "*.*") _
                        .Where(Function(fullPath)
                                   Return IO.Path.GetFileName(fullPath).ToLower().Contains(IO.Path.GetFileName(filter).ToLower())
                               End Function) _
                        .Select(
                            Function(fullPath)
                                Return New FileData() With {
                                    .FullPath = fullPath
                                }
                            End Function).ToList()

                files.AddRange(Directory.GetDirectories(pathWithoutFileName + IO.Path.DirectorySeparatorChar, "*.*") _
                        .Where(Function(fullPath)
                                   Return IO.Path.GetFileName(fullPath).ToLower().Contains(IO.Path.GetFileName(filter).ToLower())
                               End Function) _
                        .Select(
                            Function(fullPath)
                                Return New FileData() With {
                                    .FullPath = fullPath
                                }
                            End Function).ToList())

                Return files
            Catch ex As Exception
                Return Nothing
            End Try
        End If
    End Function
End Class
