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
    Private _intValue As String

    Public Property IntValue As String
        Get
            Return _intValue
        End Get
        Set(value As String)
            _intValue = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("IntValue"))
        End Set
    End Property

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

    Private Sub DoSetToNothing_Click(sender As Object, e As RoutedEventArgs) Handles DoSetToNothing.Click
        Me.FullPath = Nothing
    End Sub

    Private Sub DoSetToWindows_Click(sender As Object, e As RoutedEventArgs) Handles DoSetToWindows.Click
        Me.FullPath = "c:\Windows\"
    End Sub

    Private Sub DoSetFocus_Click(sender As Object, e As RoutedEventArgs) Handles DoSetFocus.Click
        actbFile.Focus()
    End Sub

    Private Sub DoSetToThree_Click(sender As Object, e As RoutedEventArgs) Handles DoSetToThree.Click
        Me.IntValue = "3"
    End Sub
End Class
