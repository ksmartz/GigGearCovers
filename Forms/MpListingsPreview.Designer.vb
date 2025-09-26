<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MpListingsPreview
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
        Me.lblWooSampleTitle = New System.Windows.Forms.Label()
        Me.wbWooLongDescriptionPreview = New System.Windows.Forms.WebBrowser()
        Me.txtReverbPreview = New System.Windows.Forms.TextBox()
        Me.lblWooTitle = New System.Windows.Forms.Label()
        Me.lblWooLongDescriptionSample = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblWooSampleTitle
        '
        Me.lblWooSampleTitle.AutoSize = True
        Me.lblWooSampleTitle.Location = New System.Drawing.Point(287, 57)
        Me.lblWooSampleTitle.Name = "lblWooSampleTitle"
        Me.lblWooSampleTitle.Size = New System.Drawing.Size(133, 20)
        Me.lblWooSampleTitle.TabIndex = 0
        Me.lblWooSampleTitle.Text = "Woo Sample Title"
        '
        'wbWooLongDescriptionPreview
        '
        Me.wbWooLongDescriptionPreview.Location = New System.Drawing.Point(934, 132)
        Me.wbWooLongDescriptionPreview.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbWooLongDescriptionPreview.Name = "wbWooLongDescriptionPreview"
        Me.wbWooLongDescriptionPreview.Size = New System.Drawing.Size(948, 678)
        Me.wbWooLongDescriptionPreview.TabIndex = 1
        '
        'txtReverbPreview
        '
        Me.txtReverbPreview.Location = New System.Drawing.Point(12, 132)
        Me.txtReverbPreview.MaximumSize = New System.Drawing.Size(800, 800)
        Me.txtReverbPreview.Multiline = True
        Me.txtReverbPreview.Name = "txtReverbPreview"
        Me.txtReverbPreview.Size = New System.Drawing.Size(800, 685)
        Me.txtReverbPreview.TabIndex = 2
        '
        'lblWooTitle
        '
        Me.lblWooTitle.AutoSize = True
        Me.lblWooTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWooTitle.Location = New System.Drawing.Point(43, 50)
        Me.lblWooTitle.Name = "lblWooTitle"
        Me.lblWooTitle.Size = New System.Drawing.Size(236, 29)
        Me.lblWooTitle.TabIndex = 3
        Me.lblWooTitle.Text = "Woo Title Sample: "
        '
        'lblWooLongDescriptionSample
        '
        Me.lblWooLongDescriptionSample.AutoSize = True
        Me.lblWooLongDescriptionSample.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWooLongDescriptionSample.Location = New System.Drawing.Point(955, 88)
        Me.lblWooLongDescriptionSample.Name = "lblWooLongDescriptionSample"
        Me.lblWooLongDescriptionSample.Size = New System.Drawing.Size(381, 29)
        Me.lblWooLongDescriptionSample.TabIndex = 4
        Me.lblWooLongDescriptionSample.Text = "Woo Long Description Sample: "
        '
        'MpListingsPreview
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2172, 1146)
        Me.Controls.Add(Me.lblWooLongDescriptionSample)
        Me.Controls.Add(Me.lblWooTitle)
        Me.Controls.Add(Me.txtReverbPreview)
        Me.Controls.Add(Me.wbWooLongDescriptionPreview)
        Me.Controls.Add(Me.lblWooSampleTitle)
        Me.Name = "MpListingsPreview"
        Me.Text = "MP Listings Preview"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblWooSampleTitle As Windows.Forms.Label
    Friend WithEvents wbWooLongDescriptionPreview As Windows.Forms.WebBrowser
    Friend WithEvents txtReverbPreview As Windows.Forms.TextBox
    Friend WithEvents lblWooTitle As Windows.Forms.Label
    Friend WithEvents lblWooLongDescriptionSample As Windows.Forms.Label
End Class
