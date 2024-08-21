Imports System.Windows
Imports System.Windows.Data

Public Class BindingEvaluator
    Inherits FrameworkElement

    Public Shared ReadOnly ValueProperty As DependencyProperty = DependencyProperty.Register("Value", GetType(String), GetType(BindingEvaluator), New FrameworkPropertyMetadata(String.Empty))

    Private _valueBinding As Binding

    Public Sub New(ByVal binding As Binding)
        ValueBinding = binding
    End Sub

    Public Property Value() As String
        Get
            Return GetValue(ValueProperty)
        End Get

        Set(ByVal value As String)
            SetValue(ValueProperty, value)
        End Set
    End Property

    Public Property ValueBinding() As Binding
        Get
            Return _valueBinding
        End Get
        Set(ByVal value As Binding)
            _valueBinding = value
        End Set
    End Property

    Public Function Evaluate(ByVal dataItem As Object) As String
        Me.DataContext = dataItem
        SetBinding(ValueProperty, ValueBinding)
        Return Value
    End Function
End Class
