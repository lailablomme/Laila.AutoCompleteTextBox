Imports System.ComponentModel
Imports System.IO
Imports Microsoft.WindowsAPICodePack.Shell

Public Class FileData
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _fullPath As String

    Public Property FullPath As String
        Get
            Return _fullPath
        End Get
        Set(value As String)
            _fullPath = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("FullPath"))
        End Set
    End Property

    Public ReadOnly Property FileName As String
        Get
            Return Path.GetFileName(Me.FullPath)
        End Get
    End Property

    Public ReadOnly Property Image As ImageSource
        Get
            Try
                Dim attr As FileAttributes = File.GetAttributes(_fullPath)
                If (attr And FileAttributes.Directory) = FileAttributes.Directory Then
                    Return New BitmapImage(New Uri("pack://application:,,,/Laila.AutoCompleteTextBox.SampleApp;component/003-folder.png"))
                Else
                    Return ShellFile.FromFilePath(FullPath).Thumbnail.SmallBitmapSource
                End If
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property
End Class
