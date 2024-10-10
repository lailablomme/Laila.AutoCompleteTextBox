Imports System.Reflection
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Data
Imports System.Windows.Input
Imports System.Windows.Threading

<TemplatePart(Name:=AutoCompleteTextBox.PartEditor, Type:=GetType(TextBox))>
<TemplatePart(Name:=AutoCompleteTextBox.PartPopup, Type:=GetType(System.Windows.Controls.Primitives.Popup))>
<TemplatePart(Name:=AutoCompleteTextBox.PartSelector, Type:=GetType(Selector))>
<TemplatePart(Name:=AutoCompleteTextBox.PartIcon, Type:=GetType(ContentPresenter))>
<TemplatePart(Name:=AutoCompleteTextBox.PartBorder, Type:=GetType(Border))>
<TemplatePart(Name:=AutoCompleteTextBox.PartDropDownButton, Type:=GetType(Button))>
Public Class AutoCompleteTextBox
    Inherits Control

#Region "Fields"

    Public Const PartIcon As String = "PART_Icon"
    Public Const PartEditor As String = "PART_Editor"
    Public Const PartPopup As String = "PART_Popup"
    Public Const PartSelector As String = "PART_Selector"
    Public Const PartBorder As String = "PART_Border"
    Public Const PartDropDownButton As String = "PART_DropDownButton"

    Public Shared ReadOnly AllowFreeTextProperty As DependencyProperty = DependencyProperty.Register("AllowFreeText", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False))
    Public Shared ReadOnly ShowDropDownButtonProperty As DependencyProperty = DependencyProperty.Register("ShowDropDownButton", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False))
    Public Shared ReadOnly ProviderProperty As DependencyProperty = DependencyProperty.Register("Provider", GetType(ISuggestionProviderSyncOrAsync), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing, AddressOf OnProviderChanged))
    Public Shared ReadOnly ItemPickerTypeProperty As DependencyProperty = DependencyProperty.Register("ItemPickerType", GetType(Type), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing))
    Public Shared ReadOnly WatermarkProperty As DependencyProperty = DependencyProperty.Register("Watermark", GetType(String), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(String.Empty))
    Public Shared ReadOnly DelayProperty As DependencyProperty = DependencyProperty.Register("Delay", GetType(Integer), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(200))
    Public Shared ReadOnly MinCharsProperty As DependencyProperty = DependencyProperty.Register("MinChars", GetType(Integer), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(3))
    Public Shared ReadOnly MaxLengthProperty As DependencyProperty = DependencyProperty.Register("MaxLength", GetType(Integer), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(0))
    Public Shared ReadOnly DisplayMemberProperty As DependencyProperty = DependencyProperty.Register("DisplayMember", GetType(String), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(String.Empty))
    Public Shared ReadOnly IconPlacementProperty As DependencyProperty = DependencyProperty.Register("IconPlacement", GetType(IconPlacement), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(IconPlacement.Left))
    Public Shared ReadOnly IconProperty As DependencyProperty = DependencyProperty.Register("Icon", GetType(Object), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing))
    Public Shared ReadOnly IconVisibilityProperty As DependencyProperty = DependencyProperty.Register("IconVisibility", GetType(Visibility), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Visibility.Visible))
    Public Shared ReadOnly InvalidValueProperty As DependencyProperty = DependencyProperty.Register("InvalidValue", GetType(Object), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Integer.MinValue))
    Public Shared ReadOnly IsDropDownOpenProperty As DependencyProperty = DependencyProperty.Register("IsDropDownOpen", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False))
    Public Shared ReadOnly IsBalloonOpenProperty As DependencyProperty = DependencyProperty.Register("IsBalloonOpen", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False))
    Public Shared ReadOnly IsReadOnlyProperty As DependencyProperty = DependencyProperty.Register("IsReadOnly", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False))
    Public Shared Shadows ReadOnly IsTabStopProperty As DependencyProperty = DependencyProperty.Register("IsTabStop", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(True))
    Public Shared ReadOnly IsToolTipEnabledProperty As DependencyProperty = DependencyProperty.Register("IsToolTipEnabled", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(True))
    Public Shared ReadOnly ItemTemplateProperty As DependencyProperty = DependencyProperty.Register("ItemTemplate", GetType(DataTemplate), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing))
    Public Shared ReadOnly LoadingContentProperty As DependencyProperty = DependencyProperty.Register("LoadingContent", GetType(Object), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing))
    Public Shared ReadOnly SelectedValuePathProperty As DependencyProperty = DependencyProperty.Register("SelectedValuePath", GetType(String), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing))
    Public Shared ReadOnly TextProperty As DependencyProperty = DependencyProperty.Register("Text", GetType(String), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnTextChanged))
    Public Shared ReadOnly ErrorMessageProperty As DependencyProperty = DependencyProperty.Register("ErrorMessage", GetType(String), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
    Public Shared ReadOnly SelectedItemProperty As DependencyProperty = DependencyProperty.Register("SelectedItem", GetType(Object), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedItemChanged))
    Public Shared ReadOnly SelectedValueProperty As DependencyProperty = DependencyProperty.Register("SelectedValue", GetType(Object), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedValueChanged))
    Private Shared ReadOnly IsInvalidPropertyKey As DependencyPropertyKey = DependencyProperty.RegisterReadOnly("IsInvalid", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False))
    Public Shared ReadOnly IsInvalidProperty As DependencyProperty = IsInvalidPropertyKey.DependencyProperty
    Public Shared ReadOnly IsLoadingProperty As DependencyProperty = DependencyProperty.Register("IsLoading", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False))

    Private _bindingEvaluator As BindingEvaluator
    Private _editor As TextBox
    Private _fetchTimer As DispatcherTimer
    Private _itemsSelector As Selector
    Private _popup As System.Windows.Controls.Primitives.Popup
    Private _selectionAdapter As SelectionAdapter
    Private _suggestionsAdapter As SuggestionsAdapter
    Private _iconPart As ContentPresenter
    Private _border As Border
    Private _dropDownButtonPart As Button
    Private _isErrorLoading As Boolean = False
    Private _isInitiallyLoading As Boolean = True
    Private _isWorking As Boolean = False
    Private _isUpdatingText As Boolean
    Private _isAfterInput As Boolean
    Private _isChanged As Boolean

    Protected Friend CancelText As String
    Protected Friend CancelValue As Object
    Protected Friend IsBoundEarly As Boolean

#End Region 'Fields

#Region "Constructors"

    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(GetType(AutoCompleteTextBox)))
    End Sub

    Sub New()
        Me.Focusable = False
        MyBase.IsTabStop = False
        AddHandler Me.Loaded,
            Sub()
                _isInitiallyLoading = False
            End Sub
        _suggestionsAdapter = New SuggestionsAdapter(Me)
    End Sub

#End Region 'Constructors

#Region "Properties"
    Public Property IsInitiallyLoading As Boolean
        Get
            Return _isInitiallyLoading
        End Get
        Set(value As Boolean)
            _isInitiallyLoading = value
        End Set
    End Property

    Public Property IsErrorLoading As Boolean
        Get
            Return _isErrorLoading
        End Get
        Set(value As Boolean)
            _isErrorLoading = value
        End Set
    End Property

    Public Property IsLoading() As Boolean
        Get
            Return GetValue(IsLoadingProperty)
        End Get

        Set(ByVal value As Boolean)
            SetValue(IsLoadingProperty, value)
        End Set
    End Property

    Public Property AllowFreeText As Boolean
        Get
            Return GetValue(AllowFreeTextProperty)
        End Get
        Set(value As Boolean)
            SetValue(AllowFreeTextProperty, value)
        End Set
    End Property

    Public Property IsInvalid As Boolean
        Get
            Return GetValue(IsInvalidProperty)
        End Get
        Friend Set(value As Boolean)
            SetValue(IsInvalidPropertyKey, value)
        End Set
    End Property

    Public Property BindingEvaluator() As BindingEvaluator
        Get
            Return _bindingEvaluator
        End Get
        Set(ByVal value As BindingEvaluator)
            _bindingEvaluator = value
        End Set
    End Property

    Public Property ShowDropDownButton() As Boolean
        Get
            Return GetValue(ShowDropDownButtonProperty)
        End Get

        Set(ByVal value As Boolean)
            SetValue(ShowDropDownButtonProperty, value)
        End Set
    End Property

    Public Property Delay() As Integer
        Get
            Return GetValue(DelayProperty)
        End Get

        Set(ByVal value As Integer)
            SetValue(DelayProperty, value)
        End Set
    End Property

    Public Property MinChars As Integer
        Get
            Return GetValue(MinCharsProperty)
        End Get
        Set(value As Integer)
            SetValue(MinCharsProperty, value)
        End Set
    End Property

    Public Property MaxLength As Integer
        Get
            Return GetValue(MaxLengthProperty)
        End Get
        Set(value As Integer)
            SetValue(MaxLengthProperty, value)
        End Set
    End Property

    Public Property DisplayMember() As String
        Get
            Return GetValue(DisplayMemberProperty)
        End Get

        Set(ByVal value As String)
            SetValue(DisplayMemberProperty, value)
        End Set
    End Property

    Public Property IconPart As ContentPresenter
        Get
            Return _iconPart
        End Get
        Set(value As ContentPresenter)
            _iconPart = value
        End Set
    End Property

    Public Property DropDownButtonPart As Button
        Get
            Return _dropDownButtonPart
        End Get
        Set(value As Button)
            _dropDownButtonPart = value
        End Set
    End Property

    Public Property Border As Border
        Get
            Return _border
        End Get
        Set(value As Border)
            _border = value
        End Set
    End Property

    Public Property Editor() As TextBox
        Get
            Return _editor
        End Get
        Set(ByVal value As TextBox)
            _editor = value
        End Set
    End Property

    Public Property FetchTimer() As DispatcherTimer
        Get
            Return _fetchTimer
        End Get
        Set(ByVal value As DispatcherTimer)
            _fetchTimer = value
        End Set
    End Property

    Public Property Icon() As Object
        Get
            Return GetValue(IconProperty)
        End Get

        Set(ByVal value As Object)
            SetValue(IconProperty, value)
        End Set
    End Property

    Public Property IconPlacement() As IconPlacement
        Get
            Return GetValue(IconPlacementProperty)
        End Get

        Set(ByVal value As IconPlacement)
            SetValue(IconPlacementProperty, value)
        End Set
    End Property

    Public Property IconVisibility() As Visibility
        Get
            Return GetValue(IconVisibilityProperty)
        End Get

        Set(ByVal value As Visibility)
            SetValue(IconVisibilityProperty, value)
        End Set
    End Property

    Public Property InvalidValue As Object
        Get
            Return GetValue(InvalidValueProperty)
        End Get
        Set(value As Object)
            SetValue(InvalidValueProperty, value)
        End Set
    End Property

    Public Property IsDropDownOpen() As Boolean
        Get
            Return GetValue(IsDropDownOpenProperty)
        End Get

        Set(ByVal value As Boolean)
            SetValue(IsDropDownOpenProperty, value)
        End Set
    End Property

    Public Property IsBalloonOpen() As Boolean
        Get
            Return GetValue(IsBalloonOpenProperty)
        End Get

        Set(ByVal value As Boolean)
            SetValue(IsBalloonOpenProperty, value)
        End Set
    End Property

    Public Property IsReadOnly() As Boolean
        Get
            Return GetValue(IsReadOnlyProperty)
        End Get

        Set(ByVal value As Boolean)
            SetValue(IsReadOnlyProperty, value)
        End Set
    End Property

    Public Shadows Property IsTabStop() As Boolean
        Get
            Return GetValue(IsTabStopProperty)
        End Get

        Set(ByVal value As Boolean)
            SetValue(IsTabStopProperty, value)
        End Set
    End Property

    Public Property IsToolTipEnabled As Boolean
        Get
            Return GetValue(IsToolTipEnabledProperty)
        End Get
        Set(value As Boolean)
            SetValue(IsToolTipEnabledProperty, value)
        End Set
    End Property

    Public Property ItemsSelector() As Selector
        Get
            Return _itemsSelector
        End Get
        Set(ByVal value As Selector)
            _itemsSelector = value
        End Set
    End Property

    Public Property ItemTemplate() As DataTemplate
        Get
            Return GetValue(ItemTemplateProperty)
        End Get

        Set(ByVal value As DataTemplate)
            SetValue(ItemTemplateProperty, value)
        End Set
    End Property

    Public Property LoadingContent() As Object
        Get
            Return GetValue(LoadingContentProperty)
        End Get

        Set(ByVal value As Object)
            SetValue(LoadingContentProperty, value)
        End Set
    End Property

    Public Property Popup() As System.Windows.Controls.Primitives.Popup
        Get
            Return _popup
        End Get
        Set(ByVal value As System.Windows.Controls.Primitives.Popup)
            _popup = value
        End Set
    End Property

    Public Property Provider As ISuggestionProviderSyncOrAsync
        Get
            Return GetValue(ProviderProperty)
        End Get

        Set(ByVal value As ISuggestionProviderSyncOrAsync)
            SetValue(ProviderProperty, value)
        End Set
    End Property

    Public Property ItemPickerType As Type
        Get
            Return GetValue(ItemPickerTypeProperty)
        End Get
        Set(value As Type)
            SetValue(ItemPickerTypeProperty, value)
        End Set
    End Property

    Public Property SelectedValue As Object
        Get
            Return GetValue(SelectedValueProperty)
        End Get
        Set(ByVal value As Object)
            SetValue(SelectedValueProperty, value)
        End Set
    End Property

    Public Property SelectedValuePath As String
        Get
            Return GetValue(SelectedValuePathProperty)
        End Get
        Set(ByVal value As String)
            SetValue(SelectedValuePathProperty, value)
        End Set
    End Property

    Public Property SelectedItem() As Object
        Get
            Return GetValue(SelectedItemProperty)
        End Get

        Set(ByVal value As Object)
            SetValue(SelectedItemProperty, value)
        End Set
    End Property

    Public Property SelectionAdapter() As SelectionAdapter
        Get
            Return _selectionAdapter
        End Get
        Set(ByVal value As SelectionAdapter)
            _selectionAdapter = value
        End Set
    End Property

    Public Property Text As String
        Get
            Return GetValue(TextProperty)
        End Get

        Set(ByVal value As String)
            SetValue(TextProperty, value)
        End Set
    End Property

    Public Property ErrorMessage As String
        Get
            Return GetValue(ErrorMessageProperty)
        End Get

        Set(ByVal value As String)
            SetValue(ErrorMessageProperty, value)
        End Set
    End Property

    Public Property Watermark() As String
        Get
            Return GetValue(WatermarkProperty)
        End Get

        Set(ByVal value As String)
            SetValue(WatermarkProperty, value)
        End Set
    End Property

#End Region 'Properties

#Region "Methods"

    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        Editor = Template.FindName(PartEditor, Me)
        Popup = Template.FindName(PartPopup, Me)
        ItemsSelector = Template.FindName(PartSelector, Me)

        IconPart = Template.FindName(PartIcon, Me)
        Border = Template.FindName(PartBorder, Me)
        DropDownButtonPart = Template.FindName(PartDropDownButton, Me)
        AddHandler DropDownButtonPart.Click,
            Sub()
                Me.Focus()
                openDropDown()
            End Sub
        BindingEvaluator = New BindingEvaluator(New Binding(DisplayMember))

        If Not Me.Editor Is Nothing Then
            Me.Editor.MaxLength = Me.MaxLength
            AddHandler Me.Editor.TextChanged, AddressOf OnEditorTextChanged
            AddHandler Me.Editor.PreviewKeyDown, AddressOf OnEditorKeyDownSync
            AddHandler Me.Editor.PreviewKeyDown, AddressOf OnEditorKeyDownAsync
            AddHandler Me.Editor.PreviewKeyUp, AddressOf OnEditorKeyUp
            AddHandler Me.Editor.PreviewLostKeyboardFocus, AddressOf OnEditorLostFocus

            _isUpdatingText = True
            Me.Editor.Text = Me.Text
            _isUpdatingText = False

            If Not Me.SelectedItem Is Nothing Then
                _isUpdatingText = True
                Editor.Text = Me.GetDisplayText(Me.SelectedItem)
                Editor.SelectionStart = 0
                Editor.SelectionLength = Editor.Text.Length
                _isUpdatingText = False
            End If

            AddHandler Me.IsEnabledChanged,
                Sub(s As Object, e As DependencyPropertyChangedEventArgs)
                    If Not Me.IsEnabled AndAlso Me.IsDropDownOpen Then
                        Me.IsDropDownOpen = False
                    End If
                End Sub
        End If

        If Not Me.Popup Is Nothing Then
            Me.Popup.StaysOpen = False
        End If
        If Not Me.ItemsSelector Is Nothing Then
            Me.SelectionAdapter = New SelectionAdapter(Me.ItemsSelector)
            AddHandler Me.SelectionAdapter.Commit, AddressOf OnSelectionAdapterCommit
            AddHandler Me.SelectionAdapter.Cancel, AddressOf OnSelectionAdapterCancel
            AddHandler Me.SelectionAdapter.SelectionChanged, AddressOf OnSelectionAdapterSelectionChanged
        End If
    End Sub

    Private Sub openDropDown()
        If Not FetchTimer Is Nothing Then
            FetchTimer.IsEnabled = False
            FetchTimer.Stop()
        End If
        If Not IsDropDownOpen And IsEnabled Then
            ItemsSelector.ItemsSource = Nothing
            _suggestionsAdapter.GetSuggestions(String.Empty)
        Else
            IsDropDownOpen = False
        End If
    End Sub

    Shared Sub OnProviderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim act As AutoCompleteTextBox = TryCast(d, AutoCompleteTextBox)
        If act.IsBoundEarly AndAlso Not e.NewValue Is Nothing Then
            act.OnSelectedValueChanged(act.SelectedValue)
            act.IsBoundEarly = False
        End If
    End Sub

    Shared Sub OnTextChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim act As AutoCompleteTextBox = TryCast(d, AutoCompleteTextBox)
        If Not act Is Nothing Then
            act.OnTextChanged(e)
        End If
    End Sub

    Friend Sub OnTextChanged(ByVal e As DependencyPropertyChangedEventArgs)
        If Not _isUpdatingText AndAlso Not Me.Editor Is Nothing Then
            Try
                _isUpdatingText = True
                Me.Editor.Text = e.NewValue
            Finally
                _isUpdatingText = False
            End Try
        End If
    End Sub

    Shared Sub OnSelectedItemChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim act As AutoCompleteTextBox = TryCast(d, AutoCompleteTextBox)
        If Not act Is Nothing AndAlso Not EqualityComparer(Of Object).Default.Equals(e.NewValue, e.OldValue) Then
            act.OnSelectedItemChanged(e)
        End If
    End Sub

    Friend Sub OnSelectedItemChanged(ByVal e As DependencyPropertyChangedEventArgs)
        If Not FetchTimer Is Nothing Then
            FetchTimer.Stop()
        End If
        If e.NewValue Is Nothing AndAlso Not Me.ToolTip Is Nothing AndAlso TypeOf (Me.ToolTip) Is ToolTip Then
            CType(Me.ToolTip, ToolTip).IsOpen = False
        End If

        If Not e.NewValue Is Nothing Then
            ' update text
            If Not Me.Editor Is Nothing Then
                Try
                    _isUpdatingText = True
                    Me.Editor.Text = Me.GetDisplayText(e.NewValue)
                    Me.Editor.SelectionStart = 0
                    Me.Editor.SelectionLength = Me.Editor.Text.Length
                Finally
                    _isUpdatingText = False
                End Try
            End If

            Me.CancelText = Me.GetDisplayText(e.NewValue)
        ElseIf Not Me.Editor Is Nothing AndAlso Not _isUpdatingText AndAlso Not Me.AllowFreeText AndAlso Not _isAfterInput Then
            Try
                _isUpdatingText = True
                Me.Editor.Text = Nothing
                Me.Editor.SelectionStart = 0
                Me.Editor.SelectionLength = 0
            Finally
                _isUpdatingText = False
            End Try
        End If

        _isAfterInput = False

        If Not Me._isWorking Then
            Try
                Me._isWorking = True

                If Not e.NewValue Is Nothing Then
                    ' set selected value
                    If Not String.IsNullOrWhiteSpace(Me.SelectedValuePath) Then
                        Dim p As PropertyInfo = e.NewValue.GetType().GetProperty(Me.SelectedValuePath)
                        If Not p Is Nothing Then
                            Dim newValue As Object = p.GetValue(e.NewValue, Nothing)
                            If Not EqualityComparer(Of Object).Default.Equals(newValue, Me.SelectedValue) Then
                                Me.SetCurrentValue(SelectedValueProperty, newValue)
                                Me.CancelValue = Me.SelectedValue
                            End If
                        End If
                    End If
                ElseIf Me.Editor Is Nothing OrElse String.IsNullOrWhiteSpace(Me.Editor.Text) Then
                    If Not Me.SelectedValue Is Nothing Then
                        Me.SetCurrentValue(SelectedValueProperty, Nothing)
                    End If
                ElseIf Not Me.SelectedValue = Me.InvalidValue Then
                    Me.SetCurrentValue(SelectedValueProperty, Me.InvalidValue)
                End If
            Finally
                Me._isWorking = False
            End Try
        End If

        Me.IsInvalid = Not Me.AllowFreeText AndAlso EqualityComparer(Of Object).Default.Equals(Me.InvalidValue, Me.SelectedValue)
    End Sub

    Shared Sub OnSelectedValueChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim act As AutoCompleteTextBox = TryCast(d, AutoCompleteTextBox)
        If Not act Is Nothing AndAlso Not EqualityComparer(Of Object).Default.Equals(e.NewValue, e.OldValue) Then
            act.OnSelectedValueChanged(e)
        End If
    End Sub

    Friend Sub OnSelectedValueChanged(ByVal e As DependencyPropertyChangedEventArgs)
        If Not Me._isWorking Then
            Try
                Me._isWorking = True

                If Not Me.Provider Is Nothing Then
                    Me.OnSelectedValueChanged(e.NewValue)
                Else
                    Me.IsBoundEarly = True
                End If
            Finally
                Me._isWorking = False
            End Try
        End If

        Me.IsInvalid = Not Me.AllowFreeText AndAlso EqualityComparer(Of Object).Default.Equals(Me.InvalidValue, Me.SelectedValue)
    End Sub

    Friend Sub OnSelectedValueChanged(value As Object)
        If Not value Is Nothing AndAlso Not value = Me.InvalidValue Then
            Dim currentValue As Object
            If Not Me.SelectedItem Is Nothing AndAlso Not String.IsNullOrWhiteSpace(Me.SelectedValuePath) Then
                Dim p As PropertyInfo = Me.SelectedItem.GetType().GetProperty(Me.SelectedValuePath)
                If Not p Is Nothing Then
                    currentValue = p.GetValue(Me.SelectedItem, Nothing)
                End If
            End If
            If Me.SelectedItem Is Nothing OrElse Not EqualityComparer(Of Object).Default.Equals(currentValue, Me.SelectedValue) Then
                If TypeOf Me.Provider Is ISuggestionProvider Then
                    tryGetItemSync(String.Format("GetById://{0}", Convert.ToString(value)), Nothing)
                ElseIf TypeOf Me.Provider Is ISuggestionProviderAsync Then
                    tryGetItemAsync(String.Format("GetById://{0}", Convert.ToString(value)), Nothing)
                End If
            End If
        ElseIf (value Is Nothing OrElse value = Me.InvalidValue) AndAlso Not Me.SelectedItem Is Nothing Then
            Me.SetCurrentValue(SelectedItemProperty, Nothing)
        End If
    End Sub

    Private Function GetDisplayText(ByVal dataItem As Object) As String
        If dataItem Is Nothing Then
            Return String.Empty
        ElseIf String.IsNullOrEmpty(DisplayMember) Then
            Return dataItem.ToString()
        Else
            If Me.BindingEvaluator Is Nothing Then
                Me.BindingEvaluator = New BindingEvaluator(New Binding(DisplayMember))
            End If
            Return Me.BindingEvaluator.Evaluate(dataItem)
        End If
    End Function

    Private Sub OnEditorKeyDownSync(ByVal sender As Object, ByVal e As KeyEventArgs)
        If TypeOf Me.Provider Is ISuggestionProvider Then
            Dim startListViewItem As ListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(sender)

            If (e.Key = Key.Enter AndAlso (Me.SelectionAdapter Is Nothing OrElse Me.ItemsSelector.SelectedItem Is Nothing)) _
                    OrElse e.Key = Key.Tab _
                    OrElse (e.Key = Key.Left AndAlso ((Editor.SelectionStart = 0 AndAlso Not startListViewItem Is Nothing) _
                        OrElse (startListViewItem Is Nothing AndAlso Editor.SelectionStart = 0 AndAlso Editor.SelectionLength = 0))) _
                    OrElse (e.Key = Key.Right AndAlso ((Editor.SelectionStart + Editor.SelectionLength = Editor.Text.Length _
                        AndAlso Not startListViewItem Is Nothing) OrElse (startListViewItem Is Nothing _
                        AndAlso Editor.SelectionStart = Editor.Text.Length AndAlso Editor.SelectionLength = 0))) Then
                IsDropDownOpen = False
                If _isChanged Then
                    _isAfterInput = True
                    tryGetItemSync(Editor.Text, Nothing)
                    If e.Key = Key.Right Then
                        Editor.SelectionStart = Editor.Text.Length
                    ElseIf e.Key = Key.Left Then
                        Editor.SelectionLength = 0
                    End If
                End If
                Me.OnAfterSelect()
            ElseIf e.Key = Key.Escape Then
                Me.Cancel()
            ElseIf Me.ShowDropDownButton AndAlso Keyboard.Modifiers.HasFlag(ModifierKeys.Alt) AndAlso Keyboard.IsKeyDown(Key.Down) Then
                openDropDown()
                e.Handled = True
            ElseIf Not Me.SelectionAdapter Is Nothing AndAlso IsDropDownOpen Then
                SelectionAdapter.HandleKeyDown(e)
            End If
        End If
    End Sub

    Private Async Sub OnEditorKeyDownAsync(ByVal sender As Object, ByVal e As KeyEventArgs)
        If TypeOf Me.Provider Is ISuggestionProviderAsync Then
            Dim startListViewItem As ListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(sender)

            If (e.Key = Key.Enter AndAlso (Me.SelectionAdapter Is Nothing OrElse Me.ItemsSelector.SelectedItem Is Nothing)) _
                    OrElse e.Key = Key.Tab _
                    OrElse (e.Key = Key.Left AndAlso ((Editor.SelectionStart = 0 AndAlso Not startListViewItem Is Nothing) _
                        OrElse (startListViewItem Is Nothing AndAlso Editor.SelectionStart = 0 AndAlso Editor.SelectionLength = 0))) _
                    OrElse (e.Key = Key.Right AndAlso ((Editor.SelectionStart + Editor.SelectionLength = Editor.Text.Length _
                        AndAlso Not startListViewItem Is Nothing) OrElse (startListViewItem Is Nothing _
                        AndAlso Editor.SelectionStart = Editor.Text.Length AndAlso Editor.SelectionLength = 0))) Then
                IsDropDownOpen = False
                If _isChanged Then
                    _isAfterInput = True
                    Await tryGetItemAsync(Editor.Text,
                    Async Function() As Task
                        If e.Key = Key.Right Then
                            Editor.SelectionStart = Editor.Text.Length
                        ElseIf e.Key = Key.Left Then
                            Editor.SelectionLength = 0
                        End If
                        Me.OnAfterSelect()
                    End Function)
                Else
                    Me.OnAfterSelect()
                End If
            ElseIf e.Key = Key.Escape Then
                Me.Cancel()
            ElseIf Me.ShowDropDownButton AndAlso Keyboard.Modifiers.HasFlag(ModifierKeys.Alt) AndAlso Keyboard.IsKeyDown(Key.Down) Then
                openDropDown()
                e.Handled = True
            ElseIf Not Me.SelectionAdapter Is Nothing AndAlso IsDropDownOpen Then
                SelectionAdapter.HandleKeyDown(e)
            End If
        End If
    End Sub

    Private Sub OnEditorKeyUp(sender As Object, e As KeyEventArgs)
        If SelectionAdapter IsNot Nothing AndAlso IsDropDownOpen Then
            SelectionAdapter.HandleKeyUp(e)
        End If
    End Sub

    Protected Overridable Sub OnAfterSelect()

    End Sub

    Private Sub tryGetItemSync(text As String, callback As System.Action)
        Dim list As IEnumerable = Nothing
        Dim outerEx As Exception = Nothing
        If (Not IsDropDownOpen OrElse SelectionAdapter.SelectorControl.SelectedItem Is Nothing) AndAlso Not String.IsNullOrEmpty(text) Then 'AndAlso Me.SelectedItem Is Nothing /// (text <> _lastGetItemText Or Me.SelectedItem Is Nothing) AndAlso 
            Using New OverrideCursor()
                Me.IsBalloonOpen = False
                If Not text.StartsWith("GetById://") AndAlso IsDropDownOpen AndAlso SelectionAdapter.SelectorControl.Items.Count > 0 Then
                    list = SelectionAdapter.SelectorControl.Items
                Else
                    Try
                        list = CType(Me.Provider, ISuggestionProvider).GetSuggestions(text)
                    Catch ex As Exception
                        outerEx = ex
                    End Try
                End If
            End Using
        End If
        afterTryGetItem(list, outerEx)
        If Not callback Is Nothing Then
            callback()
        End If
    End Sub

    Private Async Function tryGetItemAsync(text As String, callback As Func(Of Task)) As Task
        Dim list As IEnumerable = Nothing
        Dim outerEx As Exception = Nothing
        If (Not IsDropDownOpen OrElse SelectionAdapter.SelectorControl.SelectedItem Is Nothing) AndAlso Not String.IsNullOrEmpty(text) Then 'AndAlso Me.SelectedItem Is Nothing /// (text <> _lastGetItemText Or Me.SelectedItem Is Nothing) AndAlso 
            Using New OverrideCursor()
                Me.IsBalloonOpen = False
                If Not text.StartsWith("GetById://") AndAlso IsDropDownOpen AndAlso SelectionAdapter.SelectorControl.Items.Count > 0 Then
                    list = SelectionAdapter.SelectorControl.Items
                Else
                    Try
                        list = Await CType(Me.Provider, ISuggestionProviderAsync).GetSuggestions(text)
                    Catch ex As Exception
                        outerEx = ex
                    End Try
                End If
            End Using
        End If
        afterTryGetItem(list, outerEx)
        If Not callback Is Nothing Then
            Await callback()
        End If
    End Function

    Private Sub afterTryGetItem(list As IEnumerable, outerEx As Exception)
        _isAfterInput = False
        _isChanged = False

        If Not outerEx Is Nothing Then
            Me.IsErrorLoading = True
            If Not _isInitiallyLoading Then
                Me.IsDropDownOpen = False
                Me.ErrorMessage = outerEx.Message & If(Not outerEx.InnerException Is Nothing, vbCrLf & outerEx.InnerException.Message, "")
                Me.IsBalloonOpen = True
            End If
            Return
        End If

        If Not list Is Nothing Then
            Dim enumerator As IEnumerator = list.GetEnumerator()
            If enumerator.MoveNext() Then
                Dim item As Object = enumerator.Current
                If Not item Is Nothing Then
                    _isUpdatingText = True
                    If Not Me.Editor Is Nothing Then
                        Editor.Text = GetDisplayText(item)
                        Editor.SelectionStart = 0
                        Editor.SelectionLength = Editor.Text.Length
                    End If
                    _isUpdatingText = False
                        If Not item.Equals(Me.SelectedItem) Then
                            Me.SetCurrentValue(SelectedItemProperty, item)
                        End If
                        IsDropDownOpen = False
                        Return
                    End If
                End If
        End If

        IsDropDownOpen = False
        If Not Me.SelectedItem Is Nothing Then Me.SetCurrentValue(SelectedItemProperty, Nothing)
        If String.IsNullOrWhiteSpace(Me.Editor.Text) Then
            If Not EqualityComparer(Of Object).Default.Equals(Nothing, Me.SelectedValue) Then
                Me.SetCurrentValue(SelectedValueProperty, Nothing)
            End If
        Else
            If Not EqualityComparer(Of Object).Default.Equals(Me.InvalidValue, Me.SelectedValue) Then
                Me.SetCurrentValue(SelectedValueProperty, Me.InvalidValue)
            End If
        End If
        Debug.WriteLine(String.Format("Cannot find anything for {0} ({1})", Text, Me.Provider.ToString()))
    End Sub

    Private Sub OnEditorLostFocus(ByVal sender As Object, ByVal e As RoutedEventArgs)
        If TypeOf Me.Provider Is ISuggestionProvider Then
            If _isChanged Then
                IsDropDownOpen = False
                _isAfterInput = True
                tryGetItemSync(Me.Editor.Text, Nothing)
            End If
        ElseIf TypeOf Me.Provider Is ISuggestionProviderAsync Then
            If _isChanged Then
                IsDropDownOpen = False
                _isAfterInput = True
                Dim isReady As Boolean
                Application.Current.Dispatcher.Invoke(
                        Async Sub()
                            Await tryGetItemAsync(Me.Editor.Text, Nothing)
                            isReady = True
                        End Sub)
                While Not isReady
                    Application.Current.Dispatcher.Invoke(
                            Sub()
                            End Sub, Threading.DispatcherPriority.ContextIdle)
                End While
            End If
        End If
    End Sub

    Private Sub OnEditorTextChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs)
        _isChanged = True
        If Not Me.Text = Me.Editor.Text Then
            Me.SetCurrentValue(TextProperty, Me.Editor.Text)
        End If
        If Not _isUpdatingText Then
            Try
                _isUpdatingText = True
                If Me.Editor.Text Is Nothing OrElse Me.Editor.Text.Length = 0 Then
                    Me.SetCurrentValue(SelectedItemProperty, Nothing)
                    Me.SetCurrentValue(SelectedValueProperty, Nothing)
                End If
                If Me.FetchTimer Is Nothing Then
                    Me.FetchTimer = New DispatcherTimer()
                    Me.FetchTimer.Interval = TimeSpan.FromMilliseconds(Delay)
                    AddHandler Me.FetchTimer.Tick, AddressOf OnFetchTimerTick
                End If
                Me.FetchTimer.IsEnabled = False
                Me.FetchTimer.Stop()
                If Me.Editor.Text.Length >= Me.MinChars Then
                    Me.FetchTimer.IsEnabled = True
                    Me.FetchTimer.Start()
                Else
                    Me.IsDropDownOpen = False
                End If
            Finally
                _isUpdatingText = False
            End Try
        End If
    End Sub

    Private Sub OnFetchTimerTick(ByVal sender As Object, ByVal e As EventArgs)
        FetchTimer.IsEnabled = False
        FetchTimer.Stop()
        _suggestionsAdapter.GetSuggestions(Me.Editor.Text)
    End Sub

    Protected Overridable Sub Cancel()
        IsDropDownOpen = False
        Try
            _isUpdatingText = True
            Me.Editor.Text = Me.CancelText
            Me.Editor.SelectionStart = Me.Editor.Text.Length
        Finally
            _isUpdatingText = False
        End Try
    End Sub

    Protected Sub OnSelectionAdapterCancel()
        Me.Cancel()
    End Sub

    Private Sub OnSelectionAdapterCommit(ByRef handled As Boolean)
        If Not ItemsSelector.SelectedItem Is Nothing Then
            _isChanged = False
            Me.SetCurrentValue(SelectedItemProperty, Me.ItemsSelector.SelectedItem)
            Me.ItemsSelector.SelectedItem = Nothing ' prevent item getting picked next time before the selector even shows
            Me.IsDropDownOpen = False
            _isAfterInput = False
            Me.OnAfterSelect()
        End If
    End Sub

    Private Sub OnSelectionAdapterSelectionChanged()
        If Not Me.ItemsSelector.SelectedItem Is Nothing Then
            Try
                _isUpdatingText = True
                Me.Editor.Text = GetDisplayText(ItemsSelector.SelectedItem)
                Me.Editor.SelectionStart = Editor.Text.Length
                Me.Editor.SelectionLength = 0
            Finally
                _isUpdatingText = False
            End Try

            Dim listBox As ListBox = Me.ItemsSelector
            If Not listBox Is Nothing And Not listBox.SelectedItem Is Nothing Then
                listBox.ScrollIntoView(listBox.SelectedItem)
            End If
        End If
    End Sub
#End Region 'Methods

    Private Class SuggestionsAdapter
        Private _actb As AutoCompleteTextBox

        Public Sub New(ByVal actb As AutoCompleteTextBox)
            _actb = actb
        End Sub

        Public Async Sub GetSuggestions(ByVal filter As String)
            If TypeOf _actb.Provider Is ISuggestionProvider Then
                Dim p As ISuggestionProvider = Nothing
                _actb.IsLoading = True
                _actb.IsDropDownOpen = True
                p = _actb.Provider
                Dim list As IEnumerable = Nothing
                Try
                    Dim f As Func(Of Task) =
                        Async Function() As Task
	                        list = p.GetSuggestions(filter)
                        End Function
                    Await Task.Run(f)
                    _actb.IsBalloonOpen = False
                    Me.DisplaySuggestions(list)
                Catch ex As Exception
                    _actb.IsDropDownOpen = False
                    _actb.ErrorMessage = ex.Message & If(Not ex.InnerException Is Nothing, vbCrLf & ex.InnerException.Message, "")
                    _actb.IsBalloonOpen = True
                End Try
            ElseIf TypeOf _actb.Provider Is ISuggestionProviderAsync Then
                _actb.IsLoading = True
                _actb.IsDropDownOpen = True
                Dim list As IEnumerable = Nothing
                Try
                    list = Await CType(_actb.Provider, ISuggestionProviderAsync).GetSuggestions(filter)
                    _actb.IsBalloonOpen = False
                    Me.DisplaySuggestions(list)
                Catch ex As Exception
                    _actb.IsDropDownOpen = False
                    _actb.ErrorMessage = ex.Message & If(Not ex.InnerException Is Nothing, vbCrLf & ex.InnerException.Message, "")
                    _actb.IsBalloonOpen = True
                End Try
            End If
        End Sub

        Private Sub DisplaySuggestions(ByVal suggestions As IEnumerable)
            Using New OverrideCursor()
                    If _actb.IsKeyboardFocusWithin Then
                        _actb.ItemsSelector.ItemsSource = suggestions
                        _actb.IsDropDownOpen = _actb.ItemsSelector.HasItems AndAlso _actb.IsEnabled
                        _actb.IsLoading = Not _actb.ItemsSelector.HasItems

                        If _actb.IsDropDownOpen Then
                            _actb.CancelText = _actb.Editor.Text
                            _actb.CancelValue = _actb.SelectedValue

                            If Not _actb.SelectedValue Is Nothing AndAlso Not _actb.SelectedValue.Equals(_actb.InvalidValue) Then
                                Dim p As PropertyInfo = suggestions(0).GetType().GetProperty(_actb.SelectedValuePath)
                                Dim selectedItem As Object = Nothing
                                For Each sug In suggestions
                                    If p.GetValue(sug, Nothing) = _actb.SelectedValue Then
                                        selectedItem = sug
                                        Exit For
                                    End If
                                Next
                                If Not selectedItem Is Nothing Then
                                    _actb.ItemsSelector.SelectedItem = selectedItem
                                    CType(_actb.ItemsSelector, ListBox).ScrollIntoView(_actb.ItemsSelector.SelectedItem)
                                Else
                                    CType(_actb.ItemsSelector, ListBox).ScrollIntoView(_actb.ItemsSelector.Items(0))
                                End If
                            Else
                                CType(_actb.ItemsSelector, ListBox).ScrollIntoView(_actb.ItemsSelector.Items(0))
                            End If
                        End If
                    Else
                        _actb.IsDropDownOpen = False
                    End If
            End Using
        End Sub
    End Class
End Class
