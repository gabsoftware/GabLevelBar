Imports System.Threading
'Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.ComponentModel

Namespace GabSoftware.WinControls

    Public Class GabLevelBar

#Region " Variables "

        Private _EndColor As Color
        Private _MiddleColor As Color
        Private _StartColor As Color
        Private _InactiveEndColor As Color
        Private _InactiveMiddleColor As Color
        Private _InactiveStartColor As Color
        Private _AutoInactiveColor As Boolean
        Private _InactiveColorCoefficient As Single
        Private _Orientation As eOrientation
        Private _LedHeight As Integer
        Private _Value As Integer
        Private _ScaledValue As Integer
        Private _Maximum As Integer
        Private _Minimum As Integer
        Private _Smooth As Boolean
        Private _Thread As Thread
        Private _ActiveGradient As LinearGradientBrush
        Private _InactiveGradient As LinearGradientBrush
        Private _IsLoaded As Boolean
        Private _Inversed As Boolean


#End Region

#Region " Constructor "

        Public Sub New()

            ' This call is required by the Windows Form Designer.
            InitializeComponent()

            ' Add any initialization after the InitializeComponent() call.
            'spécifie que l'on dessine soit-même
            SetStyle(ControlStyles.ResizeRedraw, True)
            SetStyle(ControlStyles.UserPaint, True)
            SetStyle(ControlStyles.SupportsTransparentBackColor, True)
            SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
            SetStyle(ControlStyles.Selectable, True)
            SetStyle(ControlStyles.AllPaintingInWmPaint, True)
            SetStyle(ControlStyles.UserMouse, True)

            _IsLoaded = False

            'default values
            _AutoInactiveColor = True
            _InactiveColorCoefficient = 2
            _LedHeight = 5
            _Maximum = 8192
            _Minimum = 0
            Value = 4096
            _Orientation = eOrientation.Vertical
            _Smooth = False
            EndColor = Color.Red
            MiddleColor = Color.MediumPurple
            StartColor = Color.RoyalBlue
            _Inversed = False
            Me.Height = 301
            Me.Width = 25

            _IsLoaded = True

            Refresh()

        End Sub

#End Region

#Region " Enums "
        Public Enum eOrientation
            Horizontal = 1
            Vertical = 2
        End Enum
#End Region

#Region " Properties "
        <Category("LevelBar")> Public Property EndColor() As Color
            Get
                Return _EndColor
            End Get
            Set(ByVal value As Color)
                _EndColor = value
                _InactiveEndColor = Color.FromArgb(_EndColor.A, _
                                         IIf(_EndColor.R >= _InactiveColorCoefficient, _EndColor.R / _InactiveColorCoefficient, 0), _
                                         IIf(_EndColor.G >= _InactiveColorCoefficient, _EndColor.G / _InactiveColorCoefficient, 0), _
                                         IIf(_EndColor.B >= _InactiveColorCoefficient, _EndColor.B / _InactiveColorCoefficient, 0))
                Refresh()
            End Set
        End Property
        <Category("LevelBar")> Public Property MiddleColor() As Color
            Get
                Return _MiddleColor
            End Get
            Set(ByVal value As Color)
                _MiddleColor = value
                _InactiveMiddleColor = Color.FromArgb(_MiddleColor.A, _
                                         IIf(_MiddleColor.R >= _InactiveColorCoefficient, _MiddleColor.R / _InactiveColorCoefficient, 0), _
                                         IIf(_MiddleColor.G >= _InactiveColorCoefficient, _MiddleColor.G / _InactiveColorCoefficient, 0), _
                                         IIf(_MiddleColor.B >= _InactiveColorCoefficient, _MiddleColor.B / _InactiveColorCoefficient, 0))
                Refresh()
            End Set
        End Property
        <Category("LevelBar")> Public Property StartColor() As Color
            Get
                Return _StartColor
            End Get
            Set(ByVal value As Color)
                _StartColor = value
                _InactiveStartColor = Color.FromArgb(_StartColor.A, _
                                         IIf(_StartColor.R >= _InactiveColorCoefficient, _StartColor.R / _InactiveColorCoefficient, 0), _
                                         IIf(_StartColor.G >= _InactiveColorCoefficient, _StartColor.G / _InactiveColorCoefficient, 0), _
                                         IIf(_StartColor.B >= _InactiveColorCoefficient, _StartColor.B / _InactiveColorCoefficient, 0))
                Refresh()
            End Set
        End Property
        <Category("LevelBar")> Public Property InactiveEndColor() As Color
            Get
                Return _InactiveEndColor
            End Get
            Set(ByVal value As Color)
                If _AutoInactiveColor = False Then
                    _InactiveEndColor = value
                    Refresh()
                End If
            End Set
        End Property
        <Category("LevelBar")> Public Property InactiveMiddleColor() As Color
            Get
                Return _InactiveMiddleColor
            End Get
            Set(ByVal value As Color)
                If _AutoInactiveColor = False Then
                    _InactiveMiddleColor = value
                    Refresh()
                End If
            End Set
        End Property
        <Category("LevelBar")> Public Property InactiveStartColor() As Color
            Get
                Return _InactiveStartColor
            End Get
            Set(ByVal value As Color)
                If _AutoInactiveColor = False Then
                    _InactiveStartColor = value
                    Refresh()
                End If
            End Set
        End Property
        <Category("LevelBar")> Public Property AutoInactiveColor() As Boolean
            Get
                Return _AutoInactiveColor
            End Get
            Set(ByVal value As Boolean)
                _AutoInactiveColor = value

                'refresh the inactive colors
                EndColor = _EndColor
                MiddleColor = _MiddleColor
                StartColor = _StartColor
            End Set
        End Property
        <Category("LevelBar")> Public Property InactiveColorCoefficient() As Single
            Get
                Return _InactiveColorCoefficient
            End Get
            Set(ByVal value As Single)
                If value > 0 And value <= 255 Then
                    _InactiveColorCoefficient = value

                    'refresh the inactive colors
                    EndColor = _EndColor
                    MiddleColor = _MiddleColor
                    StartColor = _StartColor
                End If
            End Set
        End Property
        <Category("LevelBar")> Public Property Smooth() As Boolean
            Get
                Return _Smooth
            End Get
            Set(ByVal value As Boolean)
                _Smooth = value
                Refresh()
            End Set
        End Property
        <Category("LevelBar")> Public Property LedHeight() As Integer
            Get
                Return _LedHeight
            End Get
            Set(ByVal value As Integer)
                If value > 0 And value < Me.Height Then
                    _LedHeight = value
                    Me.Value = Me.Value
                    'Refresh()
                End If
            End Set
        End Property
        <Category("LevelBar")> Public Property Orientation() As eOrientation
            Get
                Return _Orientation
            End Get
            Set(ByVal value As eOrientation)
                _Orientation = value
                Dim oldheight As Integer = Me.Height
                Me.Height = Me.Width
                Me.Width = oldheight
                Refresh()
            End Set
        End Property
        <Category("LevelBar")> Public Property Value() As Integer
            Get
                Return _Value
            End Get
            Set(ByVal value As Integer)
                If value >= _Minimum And value <= Maximum Then
                    _Value = value
                    _ScaledValue = IIf(_Orientation = eOrientation.Vertical, _
                                       _Value - (_Value Mod ((_LedHeight * _Maximum) \ Me.Height)), _
                                       _Value - (_Value Mod ((_LedHeight * _Maximum) \ Me.Width)))
                    Refresh()
                End If
            End Set
        End Property
        <Category("LevelBar")> Public ReadOnly Property ScaledValue() As Integer
            Get
                Return _ScaledValue
            End Get
        End Property
        <Category("LevelBar")> Public Property Maximum() As Integer
            Get
                Return _Maximum
            End Get
            Set(ByVal value As Integer)
                If value >= _Value And value > _Minimum Then
                    _Maximum = value
                    Refresh()
                End If
            End Set
        End Property
        <Category("LevelBar")> Public Property Minimum() As Integer
            Get
                Return _Minimum
            End Get
            Set(ByVal value As Integer)
                If value <= _Value And value < _Maximum Then
                    _Minimum = value
                    Refresh()
                End If
            End Set
        End Property
        <Category("LevelBar")> Public Property Inversed() As Boolean
            Get
                Return _Inversed
            End Get
            Set(ByVal value As Boolean)
                _Inversed = value
                Refresh()
            End Set
        End Property

#End Region

#Region " Private methods "
        ''' <summary>
        ''' Launches our painting in a different thread
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub GabLevelBar_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
            If Not _IsLoaded Then Exit Sub

            'redraw
            If (Not _Thread Is Nothing) Then
                If (_Thread.IsAlive) Then
                    _Thread.Abort()
                End If
            End If


            _Thread = New Thread(AddressOf ThreadedPaint)
            _Thread.IsBackground = True
            _Thread.Start(e)
            _Thread.Join()
        End Sub

        ''' <summary>
        ''' We paint everything here
        ''' </summary>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub ThreadedPaint(ByVal e As System.Windows.Forms.PaintEventArgs)

            If _Inversed = True Then
                e.Graphics.TranslateTransform(Me.Width, Me.Height)
                e.Graphics.RotateTransform(180)
            End If

            If _Smooth Then
                If _Orientation = eOrientation.Vertical Then
                    'inactive gradient
                    e.Graphics.FillRectangle(_InactiveGradient, 0, 0, Me.Width, Me.Height - ((_Value * Me.Height) \ _Maximum))
                    'activegradient
                    e.Graphics.FillRectangle(_ActiveGradient, 0, Me.Height - ((_Value * Me.Height) \ _Maximum), Me.Width, (_Value * Me.Height) \ _Maximum)
                Else
                    'inactive gradient
                    e.Graphics.FillRectangle(_InactiveGradient, (_Value * Width) \ _Maximum, 0, Me.Width, Me.Height)
                    'activegradient
                    e.Graphics.FillRectangle(_ActiveGradient, 0, 0, (_Value * Width) \ _Maximum, Me.Height)
                End If
            Else
                If _Orientation = eOrientation.Vertical Then
                    'inactive gradient
                    e.Graphics.FillRectangle(_InactiveGradient, 0, 0, Me.Width, Me.Height - ((_ScaledValue * Me.Height) \ _Maximum))
                    'activegradient
                    e.Graphics.FillRectangle(_ActiveGradient, 0, Me.Height - (((_ScaledValue) * Me.Height) \ _Maximum) - 1, Me.Width, ((_ScaledValue) * Me.Height) \ _Maximum)

                    'leds borders
                    For i As Integer = Me.Height - IIf(_Inversed = False, 1, 0) To 0 Step -_LedHeight
                        e.Graphics.DrawLine(Pens.Black, 0, i, Me.Width, i)
                    Next
                Else
                    'inactive gradient
                    e.Graphics.FillRectangle(_InactiveGradient, (_ScaledValue * Width) \ _Maximum, 0, Me.Width, Me.Height)
                    'activegradient
                    e.Graphics.FillRectangle(_ActiveGradient, 0, 0, (_ScaledValue * Width) \ _Maximum + 1, Me.Height)
                    'leds borders
                    For i As Integer = IIf(_Inversed = False, 0, 1) To Me.Width Step _LedHeight
                        e.Graphics.DrawLine(Pens.Black, i, 0, i, Me.Height)
                    Next
                End If
            End If



        End Sub


#End Region

#Region " Public methods "
        ''' <summary>
        ''' We compute our gradients
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub Refresh()

            If Not _IsLoaded Then Exit Sub

            'StartColor = StartColor
            'MiddleColor = MiddleColor
            'EndColor = EndColor

            Dim cb As New ColorBlend(3)

            cb.Positions = New Single() {0.0F, 0.5F, 1.0F}

            'active gradient
            If _Orientation = eOrientation.Horizontal Then
                cb.Colors = New Color() {_StartColor, _MiddleColor, _EndColor}
                _ActiveGradient = New LinearGradientBrush(Me.Bounds, _StartColor, _EndColor, LinearGradientMode.Horizontal)
            Else
                cb.Colors = New Color() {_EndColor, _MiddleColor, _StartColor}
                _ActiveGradient = New LinearGradientBrush(Me.Bounds, _EndColor, _StartColor, LinearGradientMode.Vertical)
            End If
            _ActiveGradient.InterpolationColors = cb
            _ActiveGradient.WrapMode = WrapMode.TileFlipXY

            'inactive gradient
            If _Orientation = eOrientation.Horizontal Then
                cb.Colors = New Color() {_InactiveStartColor, _InactiveMiddleColor, _InactiveEndColor}
                _InactiveGradient = New LinearGradientBrush(Me.Bounds, _InactiveStartColor, _InactiveEndColor, LinearGradientMode.Horizontal)
            Else
                cb.Colors = New Color() {_InactiveEndColor, _InactiveMiddleColor, _InactiveStartColor}
                _InactiveGradient = New LinearGradientBrush(Me.Bounds, _InactiveEndColor, _InactiveStartColor, LinearGradientMode.Vertical)
            End If
            _InactiveGradient.InterpolationColors = cb
            _InactiveGradient.WrapMode = WrapMode.TileFlipXY

            MyBase.Refresh()

        End Sub
#End Region

    End Class
End Namespace