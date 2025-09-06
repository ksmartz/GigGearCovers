<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmAddModelInformation
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.lblAddManufacturer = New System.Windows.Forms.Label()
        Me.cmbManufacturerName = New System.Windows.Forms.ComboBox()
        Me.cmbSeriesName = New System.Windows.Forms.ComboBox()
        Me.lblSeriesName = New System.Windows.Forms.Label()
        Me.cmbEquipmentType = New System.Windows.Forms.ComboBox()
        Me.lblEquipmentType = New System.Windows.Forms.Label()
        Me.dgvModelInformation = New System.Windows.Forms.DataGridView()
        Me.btnSaveSeries = New System.Windows.Forms.Button()
        Me.dgvSeriesName = New System.Windows.Forms.DataGridView()
        Me.btnUploadModels = New System.Windows.Forms.Button()
        Me.btn_UpdateCosts = New System.Windows.Forms.Button()
        Me.txtAddManufacturer = New System.Windows.Forms.TextBox()
        Me.lblNewManufacturer = New System.Windows.Forms.Label()
        Me.btnSaveNewManufacturer = New System.Windows.Forms.Button()
        Me.btnUploadWooListings = New System.Windows.Forms.Button()
        Me.btnFixMissingParentSkus = New System.Windows.Forms.Button()
        CType(Me.dgvModelInformation, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSeriesName, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblAddManufacturer
        '
        Me.lblAddManufacturer.AutoSize = True
        Me.lblAddManufacturer.Location = New System.Drawing.Point(23, 21)
        Me.lblAddManufacturer.Name = "lblAddManufacturer"
        Me.lblAddManufacturer.Size = New System.Drawing.Size(150, 20)
        Me.lblAddManufacturer.TabIndex = 1
        Me.lblAddManufacturer.Text = "Manufacturer Name"
        '
        'cmbManufacturerName
        '
        Me.cmbManufacturerName.DropDownHeight = 500
        Me.cmbManufacturerName.FormattingEnabled = True
        Me.cmbManufacturerName.IntegralHeight = False
        Me.cmbManufacturerName.ItemHeight = 20
        Me.cmbManufacturerName.Location = New System.Drawing.Point(27, 58)
        Me.cmbManufacturerName.MaxDropDownItems = 10
        Me.cmbManufacturerName.Name = "cmbManufacturerName"
        Me.cmbManufacturerName.Size = New System.Drawing.Size(173, 28)
        Me.cmbManufacturerName.TabIndex = 2
        '
        'cmbSeriesName
        '
        Me.cmbSeriesName.FormattingEnabled = True
        Me.cmbSeriesName.Location = New System.Drawing.Point(266, 57)
        Me.cmbSeriesName.MaxDropDownItems = 10
        Me.cmbSeriesName.Name = "cmbSeriesName"
        Me.cmbSeriesName.Size = New System.Drawing.Size(223, 28)
        Me.cmbSeriesName.TabIndex = 3
        '
        'lblSeriesName
        '
        Me.lblSeriesName.AutoSize = True
        Me.lblSeriesName.Location = New System.Drawing.Point(283, 10)
        Me.lblSeriesName.Name = "lblSeriesName"
        Me.lblSeriesName.Size = New System.Drawing.Size(100, 20)
        Me.lblSeriesName.TabIndex = 4
        Me.lblSeriesName.Text = "Series Name"
        '
        'cmbEquipmentType
        '
        Me.cmbEquipmentType.FormattingEnabled = True
        Me.cmbEquipmentType.Location = New System.Drawing.Point(575, 60)
        Me.cmbEquipmentType.Name = "cmbEquipmentType"
        Me.cmbEquipmentType.Size = New System.Drawing.Size(158, 28)
        Me.cmbEquipmentType.TabIndex = 5
        '
        'lblEquipmentType
        '
        Me.lblEquipmentType.AutoSize = True
        Me.lblEquipmentType.Location = New System.Drawing.Point(605, 18)
        Me.lblEquipmentType.Name = "lblEquipmentType"
        Me.lblEquipmentType.Size = New System.Drawing.Size(124, 20)
        Me.lblEquipmentType.TabIndex = 6
        Me.lblEquipmentType.Text = "Equipment Type"
        '
        'dgvModelInformation
        '
        Me.dgvModelInformation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvModelInformation.Location = New System.Drawing.Point(12, 525)
        Me.dgvModelInformation.Name = "dgvModelInformation"
        Me.dgvModelInformation.RowHeadersWidth = 62
        Me.dgvModelInformation.RowTemplate.Height = 28
        Me.dgvModelInformation.Size = New System.Drawing.Size(2488, 569)
        Me.dgvModelInformation.TabIndex = 7
        '
        'btnSaveSeries
        '
        Me.btnSaveSeries.Location = New System.Drawing.Point(788, 21)
        Me.btnSaveSeries.Name = "btnSaveSeries"
        Me.btnSaveSeries.Size = New System.Drawing.Size(165, 37)
        Me.btnSaveSeries.TabIndex = 8
        Me.btnSaveSeries.Text = "Save/Update Series"
        Me.btnSaveSeries.UseVisualStyleBackColor = True
        '
        'dgvSeriesName
        '
        Me.dgvSeriesName.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSeriesName.Location = New System.Drawing.Point(12, 294)
        Me.dgvSeriesName.MultiSelect = False
        Me.dgvSeriesName.Name = "dgvSeriesName"
        Me.dgvSeriesName.ReadOnly = True
        Me.dgvSeriesName.RowHeadersWidth = 62
        Me.dgvSeriesName.RowTemplate.Height = 28
        Me.dgvSeriesName.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvSeriesName.Size = New System.Drawing.Size(1177, 225)
        Me.dgvSeriesName.TabIndex = 9
        '
        'btnUploadModels
        '
        Me.btnUploadModels.Location = New System.Drawing.Point(869, 80)
        Me.btnUploadModels.Name = "btnUploadModels"
        Me.btnUploadModels.Size = New System.Drawing.Size(165, 37)
        Me.btnUploadModels.TabIndex = 10
        Me.btnUploadModels.Text = "Upload CSV"
        Me.btnUploadModels.UseVisualStyleBackColor = True
        '
        'btn_UpdateCosts
        '
        Me.btn_UpdateCosts.Location = New System.Drawing.Point(1284, 43)
        Me.btn_UpdateCosts.Name = "btn_UpdateCosts"
        Me.btn_UpdateCosts.Size = New System.Drawing.Size(236, 44)
        Me.btn_UpdateCosts.TabIndex = 11
        Me.btn_UpdateCosts.Text = "Update Costs & Pricing"
        Me.btn_UpdateCosts.UseVisualStyleBackColor = True
        '
        'txtAddManufacturer
        '
        Me.txtAddManufacturer.Location = New System.Drawing.Point(27, 191)
        Me.txtAddManufacturer.Name = "txtAddManufacturer"
        Me.txtAddManufacturer.Size = New System.Drawing.Size(196, 26)
        Me.txtAddManufacturer.TabIndex = 12
        '
        'lblNewManufacturer
        '
        Me.lblNewManufacturer.AutoSize = True
        Me.lblNewManufacturer.Location = New System.Drawing.Point(61, 156)
        Me.lblNewManufacturer.Name = "lblNewManufacturer"
        Me.lblNewManufacturer.Size = New System.Drawing.Size(139, 20)
        Me.lblNewManufacturer.TabIndex = 13
        Me.lblNewManufacturer.Text = "New Manufacturer"
        '
        'btnSaveNewManufacturer
        '
        Me.btnSaveNewManufacturer.Location = New System.Drawing.Point(27, 232)
        Me.btnSaveNewManufacturer.Name = "btnSaveNewManufacturer"
        Me.btnSaveNewManufacturer.Size = New System.Drawing.Size(196, 56)
        Me.btnSaveNewManufacturer.TabIndex = 14
        Me.btnSaveNewManufacturer.Text = "Save new manufacturer"
        Me.btnSaveNewManufacturer.UseVisualStyleBackColor = True
        '
        'btnUploadWooListings
        '
        Me.btnUploadWooListings.Location = New System.Drawing.Point(540, 122)
        Me.btnUploadWooListings.Name = "btnUploadWooListings"
        Me.btnUploadWooListings.Size = New System.Drawing.Size(254, 70)
        Me.btnUploadWooListings.TabIndex = 15
        Me.btnUploadWooListings.Text = "btnUploadWooListings"
        Me.btnUploadWooListings.UseVisualStyleBackColor = True
        '
        'btnFixMissingParentSkus
        '
        Me.btnFixMissingParentSkus.Location = New System.Drawing.Point(367, 227)
        Me.btnFixMissingParentSkus.Name = "btnFixMissingParentSkus"
        Me.btnFixMissingParentSkus.Size = New System.Drawing.Size(173, 49)
        Me.btnFixMissingParentSkus.TabIndex = 16
        Me.btnFixMissingParentSkus.Text = "Fix Missing Parent Skus"
        Me.btnFixMissingParentSkus.UseVisualStyleBackColor = True
        '
        'frmAddModelInformation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2626, 1183)
        Me.Controls.Add(Me.btnFixMissingParentSkus)
        Me.Controls.Add(Me.btnUploadWooListings)
        Me.Controls.Add(Me.btnSaveNewManufacturer)
        Me.Controls.Add(Me.lblNewManufacturer)
        Me.Controls.Add(Me.txtAddManufacturer)
        Me.Controls.Add(Me.btn_UpdateCosts)
        Me.Controls.Add(Me.btnUploadModels)
        Me.Controls.Add(Me.dgvSeriesName)
        Me.Controls.Add(Me.btnSaveSeries)
        Me.Controls.Add(Me.dgvModelInformation)
        Me.Controls.Add(Me.lblEquipmentType)
        Me.Controls.Add(Me.cmbEquipmentType)
        Me.Controls.Add(Me.lblSeriesName)
        Me.Controls.Add(Me.cmbSeriesName)
        Me.Controls.Add(Me.cmbManufacturerName)
        Me.Controls.Add(Me.lblAddManufacturer)
        Me.Name = "frmAddModelInformation"
        Me.Text = "AddModelInformation"
        CType(Me.dgvModelInformation, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSeriesName, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblAddManufacturer As Windows.Forms.Label
    Friend WithEvents cmbManufacturerName As Windows.Forms.ComboBox
    Friend WithEvents cmbSeriesName As Windows.Forms.ComboBox
    Friend WithEvents lblSeriesName As Windows.Forms.Label
    Friend WithEvents cmbEquipmentType As Windows.Forms.ComboBox
    Friend WithEvents lblEquipmentType As Windows.Forms.Label
    Friend WithEvents dgvModelInformation As Windows.Forms.DataGridView
    Friend WithEvents btnSaveSeries As Windows.Forms.Button
    Friend WithEvents dgvSeriesName As Windows.Forms.DataGridView
    Friend WithEvents btnUploadModels As Windows.Forms.Button
    Friend WithEvents btn_UpdateCosts As Windows.Forms.Button
    Friend WithEvents txtAddManufacturer As Windows.Forms.TextBox
    Friend WithEvents lblNewManufacturer As Windows.Forms.Label
    Friend WithEvents btnSaveNewManufacturer As Windows.Forms.Button
    Friend WithEvents btnUploadWooListings As Windows.Forms.Button
    Friend WithEvents btnFixMissingParentSkus As Windows.Forms.Button
End Class
