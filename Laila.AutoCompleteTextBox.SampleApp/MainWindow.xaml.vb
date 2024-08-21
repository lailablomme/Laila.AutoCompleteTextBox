Imports System.ComponentModel

Class MainWindow
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.DataContext = Me
    End Sub

    Private _file As FileData
    Private _fullPath As String

    Public Property File As FileData
        Get
            Return _file
        End Get
        Set(value As FileData)
            _file = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("File"))
        End Set
    End Property

    Public Property FullPath As String
        Get
            Return _fullPath
        End Get
        Set(value As String)
            _fullPath = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("FullPath"))
        End Set
    End Property
End Class
