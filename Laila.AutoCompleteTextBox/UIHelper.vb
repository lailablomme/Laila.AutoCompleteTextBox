Imports System.Reflection
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Media

Friend Class UIHelper
    Public Shared Sub OnUIThread(action As System.Action)
        Dim app As Application = Application.Current
        If Not app Is Nothing Then
            If Not app.Dispatcher.CheckAccess() Then
                app.Dispatcher.Invoke(
                    Sub()
                        action.Invoke()
                    End Sub)
            Else
                action.Invoke()
            End If
        End If
    End Sub

    Public Shared Iterator Function FindVisualChildren(Of T As DependencyObject)(depObj As DependencyObject, Optional isDeep As Boolean = True) As IEnumerable(Of T)
        For i = 0 To VisualTreeHelper.GetChildrenCount(depObj) - 1
            Dim isMatch As Boolean = False
            Dim child As DependencyObject = VisualTreeHelper.GetChild(depObj, i)
            If Not child Is Nothing AndAlso TypeOf child Is T Then
                isMatch = True
                Yield child
            End If

            If Not isMatch Or isDeep Then
                For Each childOfChild In FindVisualChildren(Of T)(child, isDeep)
                    Yield childOfChild
                Next
            End If
        Next
    End Function

    Public Shared Iterator Function FindLogicalChildren(Of T As DependencyObject)(depObj As DependencyObject) As IEnumerable(Of T)
        For Each child In LogicalTreeHelper.GetChildren(depObj)
            If Not child Is Nothing AndAlso TypeOf child Is T Then
                Yield child
            End If

            If TypeOf child Is DependencyObject Then
                For Each childOfChild In FindLogicalChildren(Of T)(child)
                    Yield childOfChild
                Next
            End If
        Next
    End Function

    Public Shared Function IsElementVisible(element As FrameworkElement, container As FrameworkElement) As Boolean
        If Not element.IsVisible Then
            Return False
        End If

        Dim bounds As System.Windows.Rect = element.TransformToAncestor(container).TransformBounds(New System.Windows.Rect(0.0, 0.0, element.ActualWidth, element.ActualHeight))
        Dim rect As System.Windows.Rect = New System.Windows.Rect(0.0, 0.0, container.ActualWidth, container.ActualHeight)
        Return rect.Contains(bounds)
    End Function

    Public Shared Function IsIn(obj As DependencyObject, container As DependencyObject) As Boolean
        Dim parent As DependencyObject = obj
        While Not parent Is Nothing
            If Not container Is Nothing AndAlso container.Equals(parent) Then
                Return True
            End If
            parent = GetParent(parent)
        End While
        Return False
    End Function

    Public Shared Function GetParents(Of T As DependencyObject)(obj As DependencyObject) As List(Of T)
        Dim list As List(Of T) = New List(Of T)
        Dim parent As DependencyObject = GetParent(obj)
        While Not parent Is Nothing
            If TypeOf parent Is T Then
                list.Add(CType(parent, T))
            End If
            parent = GetParent(parent)
        End While
        Return list
    End Function

    Public Shared Function GetParentOfType(Of T As DependencyObject)(child As DependencyObject, root As DependencyObject) As T
        Dim parent = GetParent(child)
        If Not root Is Nothing AndAlso root.Equals(parent) Then parent = Nothing
        If Not parent Is Nothing AndAlso Not TypeOf parent Is T Then
            Return GetParentOfType(Of T)(parent, root)
        End If
        Return parent
    End Function

    Public Shared Function GetParent(obj As DependencyObject) As DependencyObject
        If obj Is Nothing Then
            Return Nothing
        End If

        If TypeOf obj Is ContentElement Then
            Dim parent As DependencyObject = ContentOperations.GetParent(obj)
            If Not parent Is Nothing Then
                Return parent
            End If
            If TypeOf obj Is FrameworkContentElement Then
                Return CType(obj, FrameworkContentElement).Parent
            Else
                Return Nothing
            End If
        End If
        Dim lp As DependencyObject = LogicalTreeHelper.GetParent(obj)
        Dim vp As DependencyObject = VisualTreeHelper.GetParent(obj)
        Return If(lp Is Nothing, vp, lp)
    End Function

    Private Function findFirstInvalidChild(obj As DependencyObject) As DependencyObject
        If TypeOf obj Is Control Then
            If CType(obj, Control).GetValue(Validation.HasErrorProperty) Then
                Return obj
            End If
        End If

        For i = 0 To VisualTreeHelper.GetChildrenCount(obj) - 1
            Dim child As DependencyObject = VisualTreeHelper.GetChild(obj, i)
            Dim match As DependencyObject = findFirstInvalidChild(child)
            If Not match Is Nothing Then
                Return match
            End If
        Next

        Return Nothing
    End Function

    Public Shared Function FindChild(Of T)(obj As DependencyObject, propertyName As String, index As Integer, Optional ByRef count As Integer? = Nothing) As DependencyObject
        If count Is Nothing Then
            count = 0
        End If

        If TypeOf obj Is T Then
            Dim bindings As Dictionary(Of DependencyProperty, Binding) = getBindings(obj)
            For Each item In bindings
                'IoC.Get(Of LoggingDAL).Log(String.Format("{0} : {1} = {2}", obj.ToString(), item.Key.Name, item.Value.Path.Path))
                If item.Value.Path.Path = propertyName Then
                    If count = index Then
                        Return obj
                    Else
                        count += 1
                    End If
                End If
            Next
        End If

        For Each child In LogicalTreeHelper.GetChildren(obj)
            If TypeOf child Is DependencyObject Then
                Dim match As DependencyObject = FindChild(Of T)(child, propertyName, index, count)
                If Not match Is Nothing Then
                    Return match
                End If
            End If
        Next

        Return Nothing
    End Function

    Private Shared Function getBindings(obj As DependencyObject) As Dictionary(Of DependencyProperty, Binding)
        Dim result As Dictionary(Of DependencyProperty, Binding) = New Dictionary(Of DependencyProperty, Binding)()
        Dim properties As List(Of DependencyProperty) = getDependencyProperties(obj)
        For Each p In properties
            Dim b As BindingBase = BindingOperations.GetBindingBase(obj, p)
            If Not b Is Nothing AndAlso TypeOf b Is Binding Then
                result.Add(p, b)
            End If
        Next
        Return result
    End Function

    Private Shared Function getDependencyProperties(obj As DependencyObject) As List(Of DependencyProperty)
        Dim result As List(Of DependencyProperty) = New List(Of DependencyProperty)()
        For Each fi As FieldInfo In obj.GetType().GetFields(BindingFlags.FlattenHierarchy Or BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Static Or BindingFlags.Public)
            If fi.FieldType = GetType(DependencyProperty) Then
                result.Add(fi.GetValue(Nothing))
            End If
        Next
        Return result
    End Function
End Class
