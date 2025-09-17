<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim CBlendItems1 As ProgBar.cBlendItems = New ProgBar.cBlendItems
        Dim CFocalPoints1 As ProgBar.cFocalPoints = New ProgBar.cFocalPoints
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form2))
        Dim CBlendItems2 As ProgBar.cBlendItems = New ProgBar.cBlendItems
        Dim CFocalPoints2 As ProgBar.cFocalPoints = New ProgBar.cFocalPoints
        Dim CBlendItems3 As ProgBar.cBlendItems = New ProgBar.cBlendItems
        Dim CFocalPoints3 As ProgBar.cFocalPoints = New ProgBar.cFocalPoints
        Dim CBlendItems4 As ProgBar.cBlendItems = New ProgBar.cBlendItems
        Dim CFocalPoints4 As ProgBar.cFocalPoints = New ProgBar.cFocalPoints
        Dim CBlendItems5 As ProgBar.cBlendItems = New ProgBar.cBlendItems
        Dim CFocalPoints5 As ProgBar.cFocalPoints = New ProgBar.cFocalPoints
        Dim CBlendItems6 As ProgBar.cBlendItems = New ProgBar.cBlendItems
        Dim CFocalPoints6 As ProgBar.cFocalPoints = New ProgBar.cFocalPoints
        Me.TrackBar1 = New System.Windows.Forms.TrackBar
        Me.pbarLoadable = New ProgBar.ProgBarPlus
        Me.ProgBarPlus7 = New ProgBar.ProgBarPlus
        Me.ProgBarPlus3 = New ProgBar.ProgBarPlus
        Me.ProgBarPlus2 = New ProgBar.ProgBarPlus
        Me.ProgBarPlus1 = New ProgBar.ProgBarPlus
        Me.ProgBarPlus4 = New ProgBar.ProgBarPlus
        CType(Me.TrackBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TrackBar1
        '
        Me.TrackBar1.Location = New System.Drawing.Point(12, 185)
        Me.TrackBar1.Maximum = 880
        Me.TrackBar1.Name = "TrackBar1"
        Me.TrackBar1.Size = New System.Drawing.Size(407, 45)
        Me.TrackBar1.TabIndex = 29
        Me.TrackBar1.Value = 440
        '
        'pbarLoadable
        '
        CBlendItems1.iColor = New System.Drawing.Color() {System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(100, Byte), Integer), CType(CType(0, Byte), Integer)), System.Drawing.Color.FromArgb(CType(CType(50, Byte), Integer), CType(CType(205, Byte), Integer), CType(CType(50, Byte), Integer))}
        CBlendItems1.iPoint = New Single() {0.0!, 1.0!}
        Me.pbarLoadable.BarColorBlend = CBlendItems1
        Me.pbarLoadable.BarColorSolid = System.Drawing.Color.LimeGreen
        Me.pbarLoadable.BarColorSolidB = System.Drawing.Color.White
        Me.pbarLoadable.BarLength = ProgBar.ProgBarPlus.eBarLength.Full
        Me.pbarLoadable.BarLengthValue = CType(25, Short)
        Me.pbarLoadable.BarPadding = New System.Windows.Forms.Padding(2)
        Me.pbarLoadable.BarStyleFill = ProgBar.ProgBarPlus.eBarStyle.GradientLinear
        Me.pbarLoadable.BarStyleHatch = System.Drawing.Drawing2D.HatchStyle.SmallCheckerBoard
        Me.pbarLoadable.BarStyleLinear = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.pbarLoadable.BarStyleTexture = Nothing
        Me.pbarLoadable.BarStyleWrapMode = System.Drawing.Drawing2D.WrapMode.Clamp
        Me.pbarLoadable.BarType = ProgBar.ProgBarPlus.eBarType.Bar
        Me.pbarLoadable.BorderColor = System.Drawing.Color.DarkGreen
        Me.pbarLoadable.BorderWidth = CType(1, Short)
        Me.pbarLoadable.Corners.All = CType(0, Short)
        Me.pbarLoadable.Corners.LowerLeft = CType(0, Short)
        Me.pbarLoadable.Corners.LowerRight = CType(0, Short)
        Me.pbarLoadable.Corners.UpperLeft = CType(0, Short)
        Me.pbarLoadable.Corners.UpperRight = CType(0, Short)
        Me.pbarLoadable.CornersApply = ProgBar.ProgBarPlus.eCornersApply.Both
        Me.pbarLoadable.CylonInterval = CType(1, Short)
        Me.pbarLoadable.CylonMove = 5.0!
        Me.pbarLoadable.FillDirection = ProgBar.ProgBarPlus.eFillDirection.Up_Right
        CFocalPoints1.CenterPoint = CType(resources.GetObject("CFocalPoints1.CenterPoint"), System.Drawing.PointF)
        CFocalPoints1.FocusScales = CType(resources.GetObject("CFocalPoints1.FocusScales"), System.Drawing.PointF)
        Me.pbarLoadable.FocalPoints = CFocalPoints1
        Me.pbarLoadable.Font = New System.Drawing.Font("Arial", 14.0!, System.Drawing.FontStyle.Bold)
        Me.pbarLoadable.ForeColor = System.Drawing.Color.Gainsboro
        Me.pbarLoadable.Location = New System.Drawing.Point(61, 12)
        Me.pbarLoadable.Max = 880
        Me.pbarLoadable.Name = "pbarLoadable"
        Me.pbarLoadable.Orientation = ProgBar.ProgBarPlus.eOrientation.Horizontal
        Me.pbarLoadable.Shape = ProgBar.ProgBarPlus.eShape.Rectangle
        Me.pbarLoadable.ShapeTextFont = New System.Drawing.Font("Arial Black", 30.0!)
        Me.pbarLoadable.ShapeTextRotate = ProgBar.ProgBarPlus.eRotateText.None
        Me.pbarLoadable.ShowDesignBorder = False
        Me.pbarLoadable.Size = New System.Drawing.Size(502, 37)
        Me.pbarLoadable.TabIndex = 31
        Me.pbarLoadable.TextAlignment = System.Drawing.StringAlignment.Center
        Me.pbarLoadable.TextAlignmentVert = System.Drawing.StringAlignment.Center
        Me.pbarLoadable.TextFormat = "Process {1}% Done"
        Me.pbarLoadable.TextPlacement = ProgBar.ProgBarPlus.eTextPlacement.OverBar
        Me.pbarLoadable.TextRotate = ProgBar.ProgBarPlus.eRotateText.None
        Me.pbarLoadable.TextShadow = True
        Me.pbarLoadable.TextShadowColor = System.Drawing.Color.Green
        Me.pbarLoadable.TextShow = ProgBar.ProgBarPlus.eTextShow.ValueOfMax
        Me.pbarLoadable.Value = 600
        '
        'ProgBarPlus7
        '
        Me.ProgBarPlus7.BarBackColor = System.Drawing.Color.LightSalmon
        CBlendItems2.iColor = New System.Drawing.Color() {System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer)), System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer)), System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))}
        CBlendItems2.iPoint = New Single() {0.0!, 0.4725275!, 1.0!}
        Me.ProgBarPlus7.BarColorBlend = CBlendItems2
        Me.ProgBarPlus7.BarColorSolid = System.Drawing.Color.Teal
        Me.ProgBarPlus7.BarColorSolidB = System.Drawing.Color.Purple
        Me.ProgBarPlus7.BarLength = ProgBar.ProgBarPlus.eBarLength.Full
        Me.ProgBarPlus7.BarLengthValue = CType(25, Short)
        Me.ProgBarPlus7.BarPadding = New System.Windows.Forms.Padding(5)
        Me.ProgBarPlus7.BarStyleFill = ProgBar.ProgBarPlus.eBarStyle.GradientPath
        Me.ProgBarPlus7.BarStyleHatch = System.Drawing.Drawing2D.HatchStyle.BackwardDiagonal
        Me.ProgBarPlus7.BarStyleLinear = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.ProgBarPlus7.BarStyleTexture = Nothing
        Me.ProgBarPlus7.BarStyleWrapMode = System.Drawing.Drawing2D.WrapMode.Clamp
        Me.ProgBarPlus7.BarType = ProgBar.ProgBarPlus.eBarType.Bar
        Me.ProgBarPlus7.BorderColor = System.Drawing.Color.MediumVioletRed
        Me.ProgBarPlus7.BorderWidth = CType(8, Short)
        Me.ProgBarPlus7.Corners.All = CType(-1, Short)
        Me.ProgBarPlus7.Corners.LowerLeft = CType(10, Short)
        Me.ProgBarPlus7.Corners.LowerRight = CType(30, Short)
        Me.ProgBarPlus7.Corners.UpperLeft = CType(50, Short)
        Me.ProgBarPlus7.Corners.UpperRight = CType(30, Short)
        Me.ProgBarPlus7.CornersApply = ProgBar.ProgBarPlus.eCornersApply.Both
        Me.ProgBarPlus7.CylonInterval = CType(1, Short)
        Me.ProgBarPlus7.CylonMove = 5.0!
        Me.ProgBarPlus7.FillDirection = ProgBar.ProgBarPlus.eFillDirection.Up_Right
        CFocalPoints2.CenterPoint = CType(resources.GetObject("CFocalPoints2.CenterPoint"), System.Drawing.PointF)
        CFocalPoints2.FocusScales = CType(resources.GetObject("CFocalPoints2.FocusScales"), System.Drawing.PointF)
        Me.ProgBarPlus7.FocalPoints = CFocalPoints2
        Me.ProgBarPlus7.Font = New System.Drawing.Font("Comic Sans MS", 20.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ProgBarPlus7.ForeColor = System.Drawing.Color.Gold
        Me.ProgBarPlus7.Location = New System.Drawing.Point(61, 311)
        Me.ProgBarPlus7.Max = 65
        Me.ProgBarPlus7.Name = "ProgBarPlus7"
        Me.ProgBarPlus7.Orientation = ProgBar.ProgBarPlus.eOrientation.Horizontal
        Me.ProgBarPlus7.Shape = ProgBar.ProgBarPlus.eShape.Rectangle
        Me.ProgBarPlus7.ShapeTextFont = New System.Drawing.Font("Arial Black", 30.0!)
        Me.ProgBarPlus7.ShapeTextRotate = ProgBar.ProgBarPlus.eRotateText.None
        Me.ProgBarPlus7.ShowDesignBorder = False
        Me.ProgBarPlus7.Size = New System.Drawing.Size(287, 120)
        Me.ProgBarPlus7.TabIndex = 30
        Me.ProgBarPlus7.TextAlignment = System.Drawing.StringAlignment.Center
        Me.ProgBarPlus7.TextAlignmentVert = System.Drawing.StringAlignment.Center
        Me.ProgBarPlus7.TextFormat = "Process {1}% Done"
        Me.ProgBarPlus7.TextPlacement = ProgBar.ProgBarPlus.eTextPlacement.OverBar
        Me.ProgBarPlus7.TextRotate = ProgBar.ProgBarPlus.eRotateText.None
        Me.ProgBarPlus7.TextShadow = True
        Me.ProgBarPlus7.TextShadowColor = System.Drawing.Color.Black
        Me.ProgBarPlus7.TextShow = ProgBar.ProgBarPlus.eTextShow.Percent
        '
        'ProgBarPlus3
        '
        CBlendItems3.iColor = New System.Drawing.Color() {System.Drawing.Color.Navy, System.Drawing.Color.Blue}
        CBlendItems3.iPoint = New Single() {0.0!, 1.0!}
        Me.ProgBarPlus3.BarColorBlend = CBlendItems3
        Me.ProgBarPlus3.BarColorSolid = System.Drawing.Color.Red
        Me.ProgBarPlus3.BarColorSolidB = System.Drawing.Color.White
        Me.ProgBarPlus3.BarLength = ProgBar.ProgBarPlus.eBarLength.Full
        Me.ProgBarPlus3.BarLengthValue = CType(25, Short)
        Me.ProgBarPlus3.BarPadding = New System.Windows.Forms.Padding(2)
        Me.ProgBarPlus3.BarStyleFill = ProgBar.ProgBarPlus.eBarStyle.Solid
        Me.ProgBarPlus3.BarStyleHatch = System.Drawing.Drawing2D.HatchStyle.Percent80
        Me.ProgBarPlus3.BarStyleLinear = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.ProgBarPlus3.BarStyleTexture = Nothing
        Me.ProgBarPlus3.BarStyleWrapMode = System.Drawing.Drawing2D.WrapMode.Clamp
        Me.ProgBarPlus3.BarType = ProgBar.ProgBarPlus.eBarType.Bar
        Me.ProgBarPlus3.BorderColor = System.Drawing.Color.MediumBlue
        Me.ProgBarPlus3.BorderWidth = CType(1, Short)
        Me.ProgBarPlus3.Corners.All = CType(0, Short)
        Me.ProgBarPlus3.Corners.LowerLeft = CType(0, Short)
        Me.ProgBarPlus3.Corners.LowerRight = CType(0, Short)
        Me.ProgBarPlus3.Corners.UpperLeft = CType(0, Short)
        Me.ProgBarPlus3.Corners.UpperRight = CType(0, Short)
        Me.ProgBarPlus3.CornersApply = ProgBar.ProgBarPlus.eCornersApply.Both
        Me.ProgBarPlus3.CylonInterval = CType(1, Short)
        Me.ProgBarPlus3.CylonMove = 5.0!
        Me.ProgBarPlus3.FillDirection = ProgBar.ProgBarPlus.eFillDirection.Down_Left
        CFocalPoints3.CenterPoint = CType(resources.GetObject("CFocalPoints3.CenterPoint"), System.Drawing.PointF)
        CFocalPoints3.FocusScales = CType(resources.GetObject("CFocalPoints3.FocusScales"), System.Drawing.PointF)
        Me.ProgBarPlus3.FocalPoints = CFocalPoints3
        Me.ProgBarPlus3.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ProgBarPlus3.ForeColor = System.Drawing.Color.White
        Me.ProgBarPlus3.Location = New System.Drawing.Point(479, 98)
        Me.ProgBarPlus3.Max = 880
        Me.ProgBarPlus3.Name = "ProgBarPlus3"
        Me.ProgBarPlus3.Orientation = ProgBar.ProgBarPlus.eOrientation.Vertical
        Me.ProgBarPlus3.Shape = ProgBar.ProgBarPlus.eShape.Rectangle
        Me.ProgBarPlus3.ShapeTextFont = New System.Drawing.Font("Arial Black", 30.0!)
        Me.ProgBarPlus3.ShapeTextRotate = ProgBar.ProgBarPlus.eRotateText.None
        Me.ProgBarPlus3.ShowDesignBorder = False
        Me.ProgBarPlus3.Size = New System.Drawing.Size(35, 407)
        Me.ProgBarPlus3.TabIndex = 28
        Me.ProgBarPlus3.TextAlignment = System.Drawing.StringAlignment.Center
        Me.ProgBarPlus3.TextAlignmentVert = System.Drawing.StringAlignment.Center
        Me.ProgBarPlus3.TextFormat = "Process {1}% Done"
        Me.ProgBarPlus3.TextPlacement = ProgBar.ProgBarPlus.eTextPlacement.OverBar
        Me.ProgBarPlus3.TextRotate = ProgBar.ProgBarPlus.eRotateText.None
        Me.ProgBarPlus3.TextShadow = True
        Me.ProgBarPlus3.TextShadowColor = System.Drawing.Color.DarkRed
        Me.ProgBarPlus3.TextShow = ProgBar.ProgBarPlus.eTextShow.None
        Me.ProgBarPlus3.Value = 440
        '
        'ProgBarPlus2
        '
        CBlendItems4.iColor = New System.Drawing.Color() {System.Drawing.Color.Navy, System.Drawing.Color.Blue}
        CBlendItems4.iPoint = New Single() {0.0!, 1.0!}
        Me.ProgBarPlus2.BarColorBlend = CBlendItems4
        Me.ProgBarPlus2.BarColorSolid = System.Drawing.Color.LimeGreen
        Me.ProgBarPlus2.BarColorSolidB = System.Drawing.Color.White
        Me.ProgBarPlus2.BarLength = ProgBar.ProgBarPlus.eBarLength.Full
        Me.ProgBarPlus2.BarLengthValue = CType(25, Short)
        Me.ProgBarPlus2.BarPadding = New System.Windows.Forms.Padding(2)
        Me.ProgBarPlus2.BarStyleFill = ProgBar.ProgBarPlus.eBarStyle.Solid
        Me.ProgBarPlus2.BarStyleHatch = System.Drawing.Drawing2D.HatchStyle.SmallCheckerBoard
        Me.ProgBarPlus2.BarStyleLinear = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.ProgBarPlus2.BarStyleTexture = Nothing
        Me.ProgBarPlus2.BarStyleWrapMode = System.Drawing.Drawing2D.WrapMode.Clamp
        Me.ProgBarPlus2.BarType = ProgBar.ProgBarPlus.eBarType.Bar
        Me.ProgBarPlus2.BorderColor = System.Drawing.Color.MediumBlue
        Me.ProgBarPlus2.BorderWidth = CType(1, Short)
        Me.ProgBarPlus2.Corners.All = CType(0, Short)
        Me.ProgBarPlus2.Corners.LowerLeft = CType(0, Short)
        Me.ProgBarPlus2.Corners.LowerRight = CType(0, Short)
        Me.ProgBarPlus2.Corners.UpperLeft = CType(0, Short)
        Me.ProgBarPlus2.Corners.UpperRight = CType(0, Short)
        Me.ProgBarPlus2.CornersApply = ProgBar.ProgBarPlus.eCornersApply.Both
        Me.ProgBarPlus2.CylonInterval = CType(1, Short)
        Me.ProgBarPlus2.CylonMove = 5.0!
        Me.ProgBarPlus2.FillDirection = ProgBar.ProgBarPlus.eFillDirection.Up_Right
        CFocalPoints4.CenterPoint = CType(resources.GetObject("CFocalPoints4.CenterPoint"), System.Drawing.PointF)
        CFocalPoints4.FocusScales = CType(resources.GetObject("CFocalPoints4.FocusScales"), System.Drawing.PointF)
        Me.ProgBarPlus2.FocalPoints = CFocalPoints4
        Me.ProgBarPlus2.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ProgBarPlus2.ForeColor = System.Drawing.Color.White
        Me.ProgBarPlus2.Location = New System.Drawing.Point(438, 98)
        Me.ProgBarPlus2.Max = 880
        Me.ProgBarPlus2.Name = "ProgBarPlus2"
        Me.ProgBarPlus2.Orientation = ProgBar.ProgBarPlus.eOrientation.Vertical
        Me.ProgBarPlus2.Shape = ProgBar.ProgBarPlus.eShape.Rectangle
        Me.ProgBarPlus2.ShapeTextFont = New System.Drawing.Font("Arial Black", 30.0!)
        Me.ProgBarPlus2.ShapeTextRotate = ProgBar.ProgBarPlus.eRotateText.None
        Me.ProgBarPlus2.ShowDesignBorder = False
        Me.ProgBarPlus2.Size = New System.Drawing.Size(35, 407)
        Me.ProgBarPlus2.TabIndex = 28
        Me.ProgBarPlus2.TextAlignment = System.Drawing.StringAlignment.Center
        Me.ProgBarPlus2.TextAlignmentVert = System.Drawing.StringAlignment.Center
        Me.ProgBarPlus2.TextFormat = "Process {1}% Done"
        Me.ProgBarPlus2.TextPlacement = ProgBar.ProgBarPlus.eTextPlacement.OverBar
        Me.ProgBarPlus2.TextRotate = ProgBar.ProgBarPlus.eRotateText.None
        Me.ProgBarPlus2.TextShadow = True
        Me.ProgBarPlus2.TextShadowColor = System.Drawing.Color.Green
        Me.ProgBarPlus2.TextShow = ProgBar.ProgBarPlus.eTextShow.None
        Me.ProgBarPlus2.Value = 440
        '
        'ProgBarPlus1
        '
        CBlendItems5.iColor = New System.Drawing.Color() {System.Drawing.Color.Navy, System.Drawing.Color.Blue}
        CBlendItems5.iPoint = New Single() {0.0!, 1.0!}
        Me.ProgBarPlus1.BarColorBlend = CBlendItems5
        Me.ProgBarPlus1.BarColorSolid = System.Drawing.Color.Red
        Me.ProgBarPlus1.BarColorSolidB = System.Drawing.Color.White
        Me.ProgBarPlus1.BarLength = ProgBar.ProgBarPlus.eBarLength.Full
        Me.ProgBarPlus1.BarLengthValue = CType(25, Short)
        Me.ProgBarPlus1.BarPadding = New System.Windows.Forms.Padding(2)
        Me.ProgBarPlus1.BarStyleFill = ProgBar.ProgBarPlus.eBarStyle.Solid
        Me.ProgBarPlus1.BarStyleHatch = System.Drawing.Drawing2D.HatchStyle.Percent80
        Me.ProgBarPlus1.BarStyleLinear = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.ProgBarPlus1.BarStyleTexture = Nothing
        Me.ProgBarPlus1.BarStyleWrapMode = System.Drawing.Drawing2D.WrapMode.Clamp
        Me.ProgBarPlus1.BarType = ProgBar.ProgBarPlus.eBarType.Bar
        Me.ProgBarPlus1.BorderColor = System.Drawing.Color.MediumBlue
        Me.ProgBarPlus1.BorderWidth = CType(1, Short)
        Me.ProgBarPlus1.Corners.All = CType(0, Short)
        Me.ProgBarPlus1.Corners.LowerLeft = CType(0, Short)
        Me.ProgBarPlus1.Corners.LowerRight = CType(0, Short)
        Me.ProgBarPlus1.Corners.UpperLeft = CType(0, Short)
        Me.ProgBarPlus1.Corners.UpperRight = CType(0, Short)
        Me.ProgBarPlus1.CornersApply = ProgBar.ProgBarPlus.eCornersApply.Both
        Me.ProgBarPlus1.CylonInterval = CType(1, Short)
        Me.ProgBarPlus1.CylonMove = 5.0!
        Me.ProgBarPlus1.FillDirection = ProgBar.ProgBarPlus.eFillDirection.Down_Left
        CFocalPoints5.CenterPoint = CType(resources.GetObject("CFocalPoints5.CenterPoint"), System.Drawing.PointF)
        CFocalPoints5.FocusScales = CType(resources.GetObject("CFocalPoints5.FocusScales"), System.Drawing.PointF)
        Me.ProgBarPlus1.FocalPoints = CFocalPoints5
        Me.ProgBarPlus1.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ProgBarPlus1.ForeColor = System.Drawing.Color.White
        Me.ProgBarPlus1.Location = New System.Drawing.Point(12, 132)
        Me.ProgBarPlus1.Max = 880
        Me.ProgBarPlus1.Name = "ProgBarPlus1"
        Me.ProgBarPlus1.Orientation = ProgBar.ProgBarPlus.eOrientation.Horizontal
        Me.ProgBarPlus1.Shape = ProgBar.ProgBarPlus.eShape.Rectangle
        Me.ProgBarPlus1.ShapeTextFont = New System.Drawing.Font("Arial Black", 30.0!)
        Me.ProgBarPlus1.ShapeTextRotate = ProgBar.ProgBarPlus.eRotateText.None
        Me.ProgBarPlus1.ShowDesignBorder = False
        Me.ProgBarPlus1.Size = New System.Drawing.Size(407, 28)
        Me.ProgBarPlus1.TabIndex = 28
        Me.ProgBarPlus1.TextAlignment = System.Drawing.StringAlignment.Center
        Me.ProgBarPlus1.TextAlignmentVert = System.Drawing.StringAlignment.Center
        Me.ProgBarPlus1.TextFormat = "Process {1}% Done"
        Me.ProgBarPlus1.TextPlacement = ProgBar.ProgBarPlus.eTextPlacement.OverBar
        Me.ProgBarPlus1.TextRotate = ProgBar.ProgBarPlus.eRotateText.None
        Me.ProgBarPlus1.TextShadow = True
        Me.ProgBarPlus1.TextShadowColor = System.Drawing.Color.DarkRed
        Me.ProgBarPlus1.TextShow = ProgBar.ProgBarPlus.eTextShow.None
        Me.ProgBarPlus1.Value = 440
        '
        'ProgBarPlus4
        '
        CBlendItems6.iColor = New System.Drawing.Color() {System.Drawing.Color.Navy, System.Drawing.Color.Blue}
        CBlendItems6.iPoint = New Single() {0.0!, 1.0!}
        Me.ProgBarPlus4.BarColorBlend = CBlendItems6
        Me.ProgBarPlus4.BarColorSolid = System.Drawing.Color.LimeGreen
        Me.ProgBarPlus4.BarColorSolidB = System.Drawing.Color.White
        Me.ProgBarPlus4.BarLength = ProgBar.ProgBarPlus.eBarLength.Full
        Me.ProgBarPlus4.BarLengthValue = CType(25, Short)
        Me.ProgBarPlus4.BarPadding = New System.Windows.Forms.Padding(2)
        Me.ProgBarPlus4.BarStyleFill = ProgBar.ProgBarPlus.eBarStyle.Solid
        Me.ProgBarPlus4.BarStyleHatch = System.Drawing.Drawing2D.HatchStyle.SmallCheckerBoard
        Me.ProgBarPlus4.BarStyleLinear = System.Drawing.Drawing2D.LinearGradientMode.Horizontal
        Me.ProgBarPlus4.BarStyleTexture = Nothing
        Me.ProgBarPlus4.BarStyleWrapMode = System.Drawing.Drawing2D.WrapMode.Clamp
        Me.ProgBarPlus4.BarType = ProgBar.ProgBarPlus.eBarType.Bar
        Me.ProgBarPlus4.BorderColor = System.Drawing.Color.MediumBlue
        Me.ProgBarPlus4.BorderWidth = CType(1, Short)
        Me.ProgBarPlus4.Corners.All = CType(0, Short)
        Me.ProgBarPlus4.Corners.LowerLeft = CType(0, Short)
        Me.ProgBarPlus4.Corners.LowerRight = CType(0, Short)
        Me.ProgBarPlus4.Corners.UpperLeft = CType(0, Short)
        Me.ProgBarPlus4.Corners.UpperRight = CType(0, Short)
        Me.ProgBarPlus4.CornersApply = ProgBar.ProgBarPlus.eCornersApply.Both
        Me.ProgBarPlus4.CylonInterval = CType(1, Short)
        Me.ProgBarPlus4.CylonMove = 5.0!
        Me.ProgBarPlus4.FillDirection = ProgBar.ProgBarPlus.eFillDirection.Up_Right
        CFocalPoints6.CenterPoint = CType(resources.GetObject("CFocalPoints6.CenterPoint"), System.Drawing.PointF)
        CFocalPoints6.FocusScales = CType(resources.GetObject("CFocalPoints6.FocusScales"), System.Drawing.PointF)
        Me.ProgBarPlus4.FocalPoints = CFocalPoints6
        Me.ProgBarPlus4.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ProgBarPlus4.ForeColor = System.Drawing.Color.White
        Me.ProgBarPlus4.Location = New System.Drawing.Point(12, 98)
        Me.ProgBarPlus4.Max = 880
        Me.ProgBarPlus4.Name = "ProgBarPlus4"
        Me.ProgBarPlus4.Orientation = ProgBar.ProgBarPlus.eOrientation.Horizontal
        Me.ProgBarPlus4.Shape = ProgBar.ProgBarPlus.eShape.Rectangle
        Me.ProgBarPlus4.ShapeTextFont = New System.Drawing.Font("Arial Black", 30.0!)
        Me.ProgBarPlus4.ShapeTextRotate = ProgBar.ProgBarPlus.eRotateText.None
        Me.ProgBarPlus4.ShowDesignBorder = False
        Me.ProgBarPlus4.Size = New System.Drawing.Size(407, 28)
        Me.ProgBarPlus4.TabIndex = 28
        Me.ProgBarPlus4.TextAlignment = System.Drawing.StringAlignment.Center
        Me.ProgBarPlus4.TextAlignmentVert = System.Drawing.StringAlignment.Center
        Me.ProgBarPlus4.TextFormat = "Process {1}% Done"
        Me.ProgBarPlus4.TextPlacement = ProgBar.ProgBarPlus.eTextPlacement.OverBar
        Me.ProgBarPlus4.TextRotate = ProgBar.ProgBarPlus.eRotateText.None
        Me.ProgBarPlus4.TextShadow = True
        Me.ProgBarPlus4.TextShadowColor = System.Drawing.Color.Green
        Me.ProgBarPlus4.TextShow = ProgBar.ProgBarPlus.eTextShow.None
        Me.ProgBarPlus4.Value = 440
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(576, 521)
        Me.Controls.Add(Me.pbarLoadable)
        Me.Controls.Add(Me.ProgBarPlus7)
        Me.Controls.Add(Me.TrackBar1)
        Me.Controls.Add(Me.ProgBarPlus3)
        Me.Controls.Add(Me.ProgBarPlus2)
        Me.Controls.Add(Me.ProgBarPlus1)
        Me.Controls.Add(Me.ProgBarPlus4)
        Me.Name = "Form2"
        Me.Text = "Form2"
        CType(Me.TrackBar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ProgBarPlus4 As ProgBar.ProgBarPlus
    Friend WithEvents ProgBarPlus1 As ProgBar.ProgBarPlus
    Friend WithEvents TrackBar1 As System.Windows.Forms.TrackBar
    Friend WithEvents ProgBarPlus2 As ProgBar.ProgBarPlus
    Friend WithEvents ProgBarPlus3 As ProgBar.ProgBarPlus
    Friend WithEvents ProgBarPlus7 As ProgBar.ProgBarPlus
    Friend WithEvents pbarLoadable As ProgBar.ProgBarPlus
End Class
