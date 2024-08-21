Imports System.Reflection
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports System.Windows.Media

Public Class SelectionAdapter
    Private _selectorControl As Selector

    Public Sub New(ByVal selector As Selector)
        SelectorControl = selector
        AddHandler SelectorControl.PreviewMouseDown, AddressOf OnPreviewMouseDown
    End Sub

    Public Event Cancel()

    Public Event Commit(ByRef handled As Boolean)

    Public Event SelectionChanged()

    Public Property SelectorControl() As Selector
        Get
            Return _selectorControl
        End Get
        Set(ByVal value As Selector)
            _selectorControl = value
        End Set
    End Property

    Public Sub HandleKeyDown(ByVal e As KeyEventArgs)
        Debug.WriteLine(e.Key)
        Select Case e.Key
            Case Key.Down
                e.Handled = True
                IncrementSelection()
            Case Key.Up
                e.Handled = True
                DecrementSelection()
            Case Key.Enter
                RaiseEvent Commit(e.Handled)
            Case Key.Escape
                e.Handled = True
                RaiseEvent Cancel()
            Case Key.Tab
                RaiseEvent Commit(e.Handled)
            Case Key.PageDown
                Dim items As Integer = SelectorControl.ActualHeight / CType(CType(SelectorControl, ListBox).ItemContainerGenerator.ContainerFromIndex(0), ListBoxItem).ActualHeight - 1
                For i = 1 To items
                    If SelectorControl.SelectedIndex < SelectorControl.Items.Count - 1 Then
                        IncrementSelection()
                    End If
                Next
            Case Key.PageUp
                Dim items As Integer = SelectorControl.ActualHeight / CType(CType(SelectorControl, ListBox).ItemContainerGenerator.ContainerFromIndex(0), ListBoxItem).ActualHeight - 1
                For i = 1 To items
                    If SelectorControl.SelectedIndex > 0 Then
                        DecrementSelection()
                    End If
                Next
        End Select
    End Sub

    Public Sub HandleKeyUp(ByVal e As KeyEventArgs)
    End Sub

    Private Sub DecrementSelection()
        If SelectorControl.SelectedIndex = -1 Then
            SelectorControl.SelectedIndex = SelectorControl.Items.Count - 1
        ElseIf SelectorControl.SelectedIndex = 0 Then
            SelectorControl.SelectedIndex = SelectorControl.Items.Count - 1
        Else
            SelectorControl.SelectedIndex -= 1
        End If
        RaiseEvent SelectionChanged()
    End Sub

    Private Sub IncrementSelection()
        If SelectorControl.SelectedIndex = -1 Then
            SelectorControl.SelectedIndex = 0
        ElseIf SelectorControl.SelectedIndex = SelectorControl.Items.Count - 1 Then
            SelectorControl.SelectedIndex = 0
        Else
            SelectorControl.SelectedIndex += 1
        End If
        RaiseEvent SelectionChanged()
    End Sub

    Private Function getListBoxItem(originalSource As Object) As ListBoxItem
        Dim value As Object = originalSource
        Dim field As FieldInfo = value.GetType().GetField("_parent", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)

        While Not field Is Nothing
            value = field.GetValue(value)
            If Not value Is Nothing AndAlso Not TypeOf value Is ListBoxItem Then
                field = value.GetType().GetField("_parent", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
            Else
                field = Nothing
            End If
        End While

        Return value
    End Function

    Private Sub OnPreviewMouseDown(ByVal sender As Object, ByVal e As MouseButtonEventArgs)
        Dim result As HitTestResult = VisualTreeHelper.HitTest(Me.SelectorControl, e.GetPosition(Me.SelectorControl))
        If Not result Is Nothing Then
            Dim listBoxItem As ListBoxItem = getListBoxItem(result.VisualHit)
            If Not listBoxItem Is Nothing Then
                listBoxItem.IsSelected = True
                ' capture the mouse because we're previewing mousedown and we don't want the click to go to the control below
                Mouse.Capture(Me.SelectorControl, CaptureMode.SubTree)
                RaiseEvent SelectionChanged()
                Dim h As Boolean = False
                RaiseEvent Commit(h)
            End If
        Else
            RaiseEvent Cancel()
        End If
    End Sub
End Class
