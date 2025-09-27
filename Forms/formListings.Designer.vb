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
        Me.btnAllListings = New System.Windows.Forms.Button()
        Me.btnClearMarketPlaces = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblEquipmentType = New System.Windows.Forms.Label()
        Me.txtEquipmentType = New System.Windows.Forms.TextBox()
        Me.btnCreateListings = New System.Windows.Forms.Button()
        Me.btnSubmitListings = New System.Windows.Forms.Button()
        CType(Me.dgvListingInformation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbManufacturerName
        '
        Me.cmbManufacturerName.FormattingEnabled = True
        Me.cmbManufacturerName.Location = New System.Drawing.Point(411, 54)
        Me.cmbManufacturerName.Name = "cmbManufacturerName"
        Me.cmbManufacturerName.Size = New System.Drawing.Size(216, 28)
        Me.cmbManufacturerName.TabIndex = 0
        '
        'cmbSeriesName
        '
        Me.cmbSeriesName.FormattingEnabled = True
        Me.cmbSeriesName.Location = New System.Drawing.Point(411, 154)
        Me.cmbSeriesName.Name = "cmbSeriesName"
        Me.cmbSeriesName.Size = New System.Drawing.Size(216, 28)
        Me.cmbSeriesName.TabIndex = 1
        '
        'dgvListingInformation
        '
        Me.dgvListingInformation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvListingInformation.Location = New System.Drawing.Point(2, 536)
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
        Me.btnReverbListings.Location = New System.Drawing.Point(956, 69)
        Me.btnReverbListings.Name = "btnReverbListings"
        Me.btnReverbListings.Size = New System.Drawing.Size(200, 60)
        Me.btnReverbListings.TabIndex = 4
        Me.btnReverbListings.Text = "Reverb"
        Me.btnReverbListings.UseVisualStyleBackColor = True
        '
        'btnAllListings
        '
        Me.btnAllListings.Location = New System.Drawing.Point(1306, 69)
        Me.btnAllListings.Name = "btnAllListings"
        Me.btnAllListings.Size = New System.Drawing.Size(182, 54)
        Me.btnAllListings.TabIndex = 5
        Me.btnAllListings.Text = "All Listings"
        Me.btnAllListings.UseVisualStyleBackColor = True
        '
        'btnClearMarketPlaces
        '
        Me.btnClearMarketPlaces.Location = New System.Drawing.Point(68, 247)
        Me.btnClearMarketPlaces.Name = "btnClearMarketPlaces"
        Me.btnClearMarketPlaces.Size = New System.Drawing.Size(166, 53)
        Me.btnClearMarketPlaces.TabIndex = 6
        Me.btnClearMarketPlaces.Text = "Clear Marketplaces"
        Me.btnClearMarketPlaces.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(407, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(153, 20)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Select Manufacturer"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(407, 120)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(103, 20)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Select Series"
        '
        'lblEquipmentType
        '
        Me.lblEquipmentType.AutoSize = True
        Me.lblEquipmentType.Location = New System.Drawing.Point(407, 218)
        Me.lblEquipmentType.Name = "lblEquipmentType"
        Me.lblEquipmentType.Size = New System.Drawing.Size(124, 20)
        Me.lblEquipmentType.TabIndex = 9
        Me.lblEquipmentType.Text = "Equipment Type"
        '
        'txtEquipmentType
        '
        Me.txtEquipmentType.Location = New System.Drawing.Point(411, 260)
        Me.txtEquipmentType.Name = "txtEquipmentType"
        Me.txtEquipmentType.Size = New System.Drawing.Size(213, 26)
        Me.txtEquipmentType.TabIndex = 10
        '
        'btnCreateListings
        '
        Me.btnCreateListings.Location = New System.Drawing.Point(68, 327)
        Me.btnCreateListings.Name = "btnCreateListings"
        Me.btnCreateListings.Size = New System.Drawing.Size(166, 53)
        Me.btnCreateListings.TabIndex = 11
        Me.btnCreateListings.Text = "Create Listings"
        Me.btnCreateListings.UseVisualStyleBackColor = True
        '
        'btnSubmitListings
        '
        Me.btnSubmitListings.Location = New System.Drawing.Point(68, 412)
        Me.btnSubmitListings.Name = "btnSubmitListings"
        Me.btnSubmitListings.Size = New System.Drawing.Size(166, 53)
        Me.btnSubmitListings.TabIndex = 12
        Me.btnSubmitListings.Text = "Submit Listings"
        Me.btnSubmitListings.UseVisualStyleBackColor = True
        '
        'formListings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2615, 920)
        Me.Controls.Add(Me.btnSubmitListings)
        Me.Controls.Add(Me.btnCreateListings)
        Me.Controls.Add(Me.txtEquipmentType)
        Me.Controls.Add(Me.lblEquipmentType)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnClearMarketPlaces)
        Me.Controls.Add(Me.btnAllListings)
        Me.Controls.Add(Me.btnReverbListings)
        Me.Controls.Add(Me.btnApiTester)
        Me.Controls.Add(Me.dgvListingInformation)
        Me.Controls.Add(Me.cmbSeriesName)
        Me.Controls.Add(Me.cmbManufacturerName)
        Me.Name = "formListings"
        Me.Text = "formListings"
        CType(Me.dgvListingInformation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmbManufacturerName As Windows.Forms.ComboBox
    Friend WithEvents cmbSeriesName As Windows.Forms.ComboBox
    Friend WithEvents dgvListingInformation As Windows.Forms.DataGridView
    Friend WithEvents btnApiTester As Windows.Forms.Button
    Friend WithEvents btnReverbListings As Windows.Forms.Button
    Friend WithEvents btnAllListings As Windows.Forms.Button
    Friend WithEvents btnClearMarketPlaces As Windows.Forms.Button
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents lblEquipmentType As Windows.Forms.Label
    Friend WithEvents txtEquipmentType As Windows.Forms.TextBox
    Friend WithEvents btnCreateListings As Windows.Forms.Button
    Friend WithEvents btnSubmitListings As Windows.Forms.Button
End Class
