Partial Class AtifNaseem
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AtifNaseem))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.GrpCCase1 = Me.Factory.CreateRibbonGroup
        Me.BtnLCase = Me.Factory.CreateRibbonButton
        Me.BtnUCase = Me.Factory.CreateRibbonButton
        Me.BtnPCase = Me.Factory.CreateRibbonButton
        Me.GrpCCase2 = Me.Factory.CreateRibbonGroup
        Me.BtnSCase = Me.Factory.CreateRibbonButton
        Me.BtnTxt2Num = Me.Factory.CreateRibbonButton
        Me.BtnNum2Txt = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.GrpCCase1.SuspendLayout()
        Me.GrpCCase2.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.GrpCCase1)
        Me.Tab1.Groups.Add(Me.GrpCCase2)
        resources.ApplyResources(Me.Tab1, "Tab1")
        Me.Tab1.Name = "Tab1"
        '
        'GrpCCase1
        '
        Me.GrpCCase1.Items.Add(Me.BtnLCase)
        Me.GrpCCase1.Items.Add(Me.BtnUCase)
        Me.GrpCCase1.Items.Add(Me.BtnTxt2Num)
        Me.GrpCCase1.Name = "GrpCCase1"
        '
        'BtnLCase
        '
        resources.ApplyResources(Me.BtnLCase, "BtnLCase")
        Me.BtnLCase.Name = "BtnLCase"
        '
        'BtnUCase
        '
        resources.ApplyResources(Me.BtnUCase, "BtnUCase")
        Me.BtnUCase.Name = "BtnUCase"
        '
        'BtnPCase
        '
        resources.ApplyResources(Me.BtnPCase, "BtnPCase")
        Me.BtnPCase.Name = "BtnPCase"
        '
        'GrpCCase2
        '
        Me.GrpCCase2.Items.Add(Me.BtnPCase)
        Me.GrpCCase2.Items.Add(Me.BtnSCase)
        Me.GrpCCase2.Items.Add(Me.BtnNum2Txt)
        Me.GrpCCase2.Name = "GrpCCase2"
        '
        'BtnSCase
        '
        resources.ApplyResources(Me.BtnSCase, "BtnSCase")
        Me.BtnSCase.Name = "BtnSCase"
        '
        'BtnTxt2Num
        '
        resources.ApplyResources(Me.BtnTxt2Num, "BtnTxt2Num")
        Me.BtnTxt2Num.Name = "BtnTxt2Num"
        '
        'BtnNum2Txt
        '
        resources.ApplyResources(Me.BtnNum2Txt, "BtnNum2Txt")
        Me.BtnNum2Txt.Name = "BtnNum2Txt"
        '
        'AtifNaseem
        '
        Me.Name = "AtifNaseem"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.GrpCCase1.ResumeLayout(False)
        Me.GrpCCase1.PerformLayout()
        Me.GrpCCase2.ResumeLayout(False)
        Me.GrpCCase2.PerformLayout()

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents GrpCCase1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BtnUCase As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnLCase As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnPCase As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents GrpCCase2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BtnSCase As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnTxt2Num As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnNum2Txt As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As AtifNaseem
        Get
            Return Me.GetRibbon(Of AtifNaseem)()
        End Get
    End Property
End Class
