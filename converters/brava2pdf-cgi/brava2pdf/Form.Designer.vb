<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form))
        Me.bravaComponent = New AxBRAVADTXLib.AxBravaDTXView()
        CType(Me.bravaComponent, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'bravaComponent
        '
        Me.bravaComponent.Enabled = True
        Me.bravaComponent.Location = New System.Drawing.Point(12, 12)
        Me.bravaComponent.Name = "bravaComponent"
        Me.bravaComponent.OcxState = CType(resources.GetObject("bravaComponent.OcxState"), System.Windows.Forms.AxHost.State)
        Me.bravaComponent.Size = New System.Drawing.Size(268, 249)
        Me.bravaComponent.TabIndex = 0
        '
        'Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.Add(Me.bravaComponent)
        Me.Name = "Form"
        Me.Text = "brava2pdf"
        CType(Me.bravaComponent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents bravaComponent As AxBRAVADTXLib.AxBravaDTXView

End Class
