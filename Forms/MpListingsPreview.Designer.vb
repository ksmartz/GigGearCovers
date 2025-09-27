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
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.wbWooLongDescriptionPreview = New System.Windows.Forms.WebBrowser()
        Me.txtDescriptionPreview = New System.Windows.Forms.TextBox()
        Me.lblWooTitle = New System.Windows.Forms.Label()
        Me.lblWooLongDescriptionSample = New System.Windows.Forms.Label()
        Me.txtFieldNameValues = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Location = New System.Drawing.Point(64, 97)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(133, 20)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "Woo Sample Title"
        '
        'wbWooLongDescriptionPreview
        '
        Me.wbWooLongDescriptionPreview.Location = New System.Drawing.Point(934, 132)
        Me.wbWooLongDescriptionPreview.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbWooLongDescriptionPreview.Name = "wbWooLongDescriptionPreview"
        Me.wbWooLongDescriptionPreview.Size = New System.Drawing.Size(948, 678)
        Me.wbWooLongDescriptionPreview.TabIndex = 1
        '
        'txtDescriptionPreview
        '
        Me.txtDescriptionPreview.Location = New System.Drawing.Point(39, 499)
        Me.txtDescriptionPreview.MaximumSize = New System.Drawing.Size(800, 800)
        Me.txtDescriptionPreview.Multiline = True
        Me.txtDescriptionPreview.Name = "txtDescriptionPreview"
        Me.txtDescriptionPreview.Size = New System.Drawing.Size(800, 685)
        Me.txtDescriptionPreview.TabIndex = 2
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
        'txtFieldNameValues
        '
        Me.txtFieldNameValues.Location = New System.Drawing.Point(39, 147)
        Me.txtFieldNameValues.MaximumSize = New System.Drawing.Size(800, 800)
        Me.txtFieldNameValues.Multiline = True
        Me.txtFieldNameValues.Name = "txtFieldNameValues"
        Me.txtFieldNameValues.Size = New System.Drawing.Size(785, 346)
        Me.txtFieldNameValues.TabIndex = 5
        '
        'MpListingsPreview
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2172, 1404)
        Me.Controls.Add(Me.txtFieldNameValues)
        Me.Controls.Add(Me.lblWooLongDescriptionSample)
        Me.Controls.Add(Me.lblWooTitle)
        Me.Controls.Add(Me.txtDescriptionPreview)
        Me.Controls.Add(Me.wbWooLongDescriptionPreview)
        Me.Controls.Add(Me.lblTitle)
        Me.Name = "MpListingsPreview"
        Me.Text = "MP Listings Preview"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblTitle As Windows.Forms.Label
    Friend WithEvents wbWooLongDescriptionPreview As Windows.Forms.WebBrowser
    Friend WithEvents txtDescriptionPreview As Windows.Forms.TextBox
    Friend WithEvents lblWooTitle As Windows.Forms.Label
    Friend WithEvents lblWooLongDescriptionSample As Windows.Forms.Label
    Friend WithEvents txtFieldNameValues As Windows.Forms.TextBox
End Class
