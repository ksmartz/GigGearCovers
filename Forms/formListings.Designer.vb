<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class formListings
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
        Me.cmbManufacturerName = New System.Windows.Forms.ComboBox()
        Me.cmbSeriesName = New System.Windows.Forms.ComboBox()
        Me.dgvListingInformation = New System.Windows.Forms.DataGridView()
        Me.btnApiTester = New System.Windows.Forms.Button()
        Me.btnReverbListings = New System.Windows.Forms.Button()
        CType(Me.dgvListingInformation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbManufacturerName
        '
        Me.cmbManufacturerName.FormattingEnabled = True
        Me.cmbManufacturerName.Location = New System.Drawing.Point(16, 35)
        Me.cmbManufacturerName.Name = "cmbManufacturerName"
        Me.cmbManufacturerName.Size = New System.Drawing.Size(121, 28)
        Me.cmbManufacturerName.TabIndex = 0
        '
        'cmbSeriesName
        '
        Me.cmbSeriesName.FormattingEnabled = True
        Me.cmbSeriesName.Location = New System.Drawing.Point(249, 45)
        Me.cmbSeriesName.Name = "cmbSeriesName"
        Me.cmbSeriesName.Size = New System.Drawing.Size(121, 28)
        Me.cmbSeriesName.TabIndex = 1
        '
        'dgvListingInformation
        '
        Me.dgvListingInformation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvListingInformation.Location = New System.Drawing.Point(16, 168)
        Me.dgvListingInformation.Name = "dgvListingInformation"
        Me.dgvListingInformation.RowHeadersWidth = 62
        Me.dgvListingInformation.RowTemplate.Height = 28
        Me.dgvListingInformation.Size = New System.Drawing.Size(2587, 248)
        Me.dgvListingInformation.TabIndex = 2
        '
        'btnApiTester
        '
        Me.btnApiTester.Location = New System.Drawing.Point(2105, 48)
        Me.btnApiTester.Name = "btnApiTester"
        Me.btnApiTester.Size = New System.Drawing.Size(160, 39)
        Me.btnApiTester.TabIndex = 3
        Me.btnApiTester.Text = "API Tester"
        Me.btnApiTester.UseVisualStyleBackColor = True
        '
        'btnReverbListings
        '
        Me.btnReverbListings.Location = New System.Drawing.Point(716, 64)
        Me.btnReverbListings.Name = "btnReverbListings"
        Me.btnReverbListings.Size = New System.Drawing.Size(200, 60)
        Me.btnReverbListings.TabIndex = 4
        Me.btnReverbListings.Text = "Reverb"
        Me.btnReverbListings.UseVisualStyleBackColor = True
        '
        'formListings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2615, 522)
        Me.Controls.Add(Me.btnReverbListings)
        Me.Controls.Add(Me.btnApiTester)
        Me.Controls.Add(Me.dgvListingInformation)
        Me.Controls.Add(Me.cmbSeriesName)
        Me.Controls.Add(Me.cmbManufacturerName)
        Me.Name = "formListings"
        Me.Text = "formListings"
        CType(Me.dgvListingInformation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cmbManufacturerName As Windows.Forms.ComboBox
    Friend WithEvents cmbSeriesName As Windows.Forms.ComboBox
    Friend WithEvents dgvListingInformation As Windows.Forms.DataGridView
    Friend WithEvents btnApiTester As Windows.Forms.Button
    Friend WithEvents btnReverbListings As Windows.Forms.Button
End Class
