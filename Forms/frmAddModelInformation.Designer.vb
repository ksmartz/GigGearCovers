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
        Me.btnPublishFabricColor = New System.Windows.Forms.Button()
        Me.pnlVariationBuilder = New System.Windows.Forms.Panel()
        Me.tlVariationRoot = New System.Windows.Forms.TableLayoutPanel()
        Me.dgMatrix = New System.Windows.Forms.DataGridView()
        Me.tlRight = New System.Windows.Forms.TableLayoutPanel()
        Me.grpPreview = New System.Windows.Forms.GroupBox()
        Me.picVariation = New System.Windows.Forms.PictureBox()
        Me.picParent = New System.Windows.Forms.PictureBox()
        Me.btnSet = New System.Windows.Forms.Button()
        Me.TextBox1txtVariationIMageURL = New System.Windows.Forms.TextBox()
        Me.lblComputerPrice = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.nudFabricAdder = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.nudBasePrice = New System.Windows.Forms.NumericUpDown()
        Me.txtTitlePreview = New System.Windows.Forms.TextBox()
        Me.grpAddOns = New System.Windows.Forms.GroupBox()
        Me.lvAddOns = New System.Windows.Forms.ListView()
        CType(Me.dgvModelInformation, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSeriesName, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlVariationBuilder.SuspendLayout()
        Me.tlVariationRoot.SuspendLayout()
        CType(Me.dgMatrix, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tlRight.SuspendLayout()
        Me.grpPreview.SuspendLayout()
        CType(Me.picVariation, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picParent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudFabricAdder, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudBasePrice, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpAddOns.SuspendLayout()
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
        Me.dgvModelInformation.Location = New System.Drawing.Point(12, 588)
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
        'btnPublishFabricColor
        '
        Me.btnPublishFabricColor.Location = New System.Drawing.Point(802, 230)
        Me.btnPublishFabricColor.Name = "btnPublishFabricColor"
        Me.btnPublishFabricColor.Size = New System.Drawing.Size(159, 45)
        Me.btnPublishFabricColor.TabIndex = 17
        Me.btnPublishFabricColor.Text = "Fabric Color"
        Me.btnPublishFabricColor.UseVisualStyleBackColor = True
        '
        'pnlVariationBuilder
        '
        Me.pnlVariationBuilder.Controls.Add(Me.tlVariationRoot)
        Me.pnlVariationBuilder.Location = New System.Drawing.Point(1292, 104)
        Me.pnlVariationBuilder.Name = "pnlVariationBuilder"
        Me.pnlVariationBuilder.Size = New System.Drawing.Size(1155, 467)
        Me.pnlVariationBuilder.TabIndex = 18
        '
        'tlVariationRoot
        '
        Me.tlVariationRoot.ColumnCount = 2
        Me.tlVariationRoot.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 60.0!))
        Me.tlVariationRoot.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 40.0!))
        Me.tlVariationRoot.Controls.Add(Me.dgMatrix, 0, 0)
        Me.tlVariationRoot.Controls.Add(Me.tlRight, 1, 0)
        Me.tlVariationRoot.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tlVariationRoot.Location = New System.Drawing.Point(0, 0)
        Me.tlVariationRoot.Name = "tlVariationRoot"
        Me.tlVariationRoot.RowCount = 1
        Me.tlVariationRoot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlVariationRoot.Size = New System.Drawing.Size(1155, 467)
        Me.tlVariationRoot.TabIndex = 0
        '
        'dgMatrix
        '
        Me.dgMatrix.AllowUserToAddRows = False
        Me.dgMatrix.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgMatrix.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgMatrix.Location = New System.Drawing.Point(3, 3)
        Me.dgMatrix.Name = "dgMatrix"
        Me.dgMatrix.RowHeadersVisible = False
        Me.dgMatrix.RowHeadersWidth = 62
        Me.dgMatrix.RowTemplate.Height = 28
        Me.dgMatrix.Size = New System.Drawing.Size(687, 461)
        Me.dgMatrix.TabIndex = 0
        '
        'tlRight
        '
        Me.tlRight.ColumnCount = 1
        Me.tlRight.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlRight.Controls.Add(Me.grpPreview, 0, 0)
        Me.tlRight.Controls.Add(Me.grpAddOns, 0, 1)
        Me.tlRight.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tlRight.Location = New System.Drawing.Point(696, 3)
        Me.tlRight.Name = "tlRight"
        Me.tlRight.RowCount = 2
        Me.tlRight.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 55.0!))
        Me.tlRight.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 45.0!))
        Me.tlRight.Size = New System.Drawing.Size(456, 461)
        Me.tlRight.TabIndex = 1
        '
        'grpPreview
        '
        Me.grpPreview.Controls.Add(Me.picVariation)
        Me.grpPreview.Controls.Add(Me.picParent)
        Me.grpPreview.Controls.Add(Me.btnSet)
        Me.grpPreview.Controls.Add(Me.TextBox1txtVariationIMageURL)
        Me.grpPreview.Controls.Add(Me.lblComputerPrice)
        Me.grpPreview.Controls.Add(Me.Label2)
        Me.grpPreview.Controls.Add(Me.nudFabricAdder)
        Me.grpPreview.Controls.Add(Me.Label1)
        Me.grpPreview.Controls.Add(Me.nudBasePrice)
        Me.grpPreview.Controls.Add(Me.txtTitlePreview)
        Me.grpPreview.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpPreview.Location = New System.Drawing.Point(3, 3)
        Me.grpPreview.Name = "grpPreview"
        Me.grpPreview.Size = New System.Drawing.Size(450, 247)
        Me.grpPreview.TabIndex = 0
        Me.grpPreview.TabStop = False
        Me.grpPreview.Text = "Preview"
        '
        'picVariation
        '
        Me.picVariation.Location = New System.Drawing.Point(212, 169)
        Me.picVariation.Name = "picVariation"
        Me.picVariation.Size = New System.Drawing.Size(119, 72)
        Me.picVariation.TabIndex = 9
        Me.picVariation.TabStop = False
        '
        'picParent
        '
        Me.picParent.Location = New System.Drawing.Point(23, 175)
        Me.picParent.Name = "picParent"
        Me.picParent.Size = New System.Drawing.Size(119, 72)
        Me.picParent.TabIndex = 8
        Me.picParent.TabStop = False
        '
        'btnSet
        '
        Me.btnSet.Location = New System.Drawing.Point(138, 135)
        Me.btnSet.Name = "btnSet"
        Me.btnSet.Size = New System.Drawing.Size(48, 34)
        Me.btnSet.TabIndex = 7
        Me.btnSet.Text = "Set"
        Me.btnSet.UseVisualStyleBackColor = True
        '
        'TextBox1txtVariationIMageURL
        '
        Me.TextBox1txtVariationIMageURL.Location = New System.Drawing.Point(23, 143)
        Me.TextBox1txtVariationIMageURL.Name = "TextBox1txtVariationIMageURL"
        Me.TextBox1txtVariationIMageURL.Size = New System.Drawing.Size(100, 26)
        Me.TextBox1txtVariationIMageURL.TabIndex = 6
        '
        'lblComputerPrice
        '
        Me.lblComputerPrice.AutoSize = True
        Me.lblComputerPrice.Location = New System.Drawing.Point(308, 81)
        Me.lblComputerPrice.Name = "lblComputerPrice"
        Me.lblComputerPrice.Size = New System.Drawing.Size(57, 20)
        Me.lblComputerPrice.TabIndex = 5
        Me.lblComputerPrice.Text = "Label3"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(177, 66)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 20)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Label2"
        '
        'nudFabricAdder
        '
        Me.nudFabricAdder.Location = New System.Drawing.Point(158, 98)
        Me.nudFabricAdder.Name = "nudFabricAdder"
        Me.nudFabricAdder.Size = New System.Drawing.Size(120, 26)
        Me.nudFabricAdder.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(23, 60)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 20)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Label1"
        '
        'nudBasePrice
        '
        Me.nudBasePrice.Location = New System.Drawing.Point(4, 92)
        Me.nudBasePrice.Name = "nudBasePrice"
        Me.nudBasePrice.Size = New System.Drawing.Size(120, 26)
        Me.nudBasePrice.TabIndex = 1
        '
        'txtTitlePreview
        '
        Me.txtTitlePreview.Dock = System.Windows.Forms.DockStyle.Top
        Me.txtTitlePreview.Location = New System.Drawing.Point(3, 22)
        Me.txtTitlePreview.Name = "txtTitlePreview"
        Me.txtTitlePreview.Size = New System.Drawing.Size(444, 26)
        Me.txtTitlePreview.TabIndex = 0
        '
        'grpAddOns
        '
        Me.grpAddOns.Controls.Add(Me.lvAddOns)
        Me.grpAddOns.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpAddOns.Location = New System.Drawing.Point(3, 256)
        Me.grpAddOns.Name = "grpAddOns"
        Me.grpAddOns.Size = New System.Drawing.Size(450, 202)
        Me.grpAddOns.TabIndex = 1
        Me.grpAddOns.TabStop = False
        Me.grpAddOns.Text = "Add Ons"
        '
        'lvAddOns
        '
        Me.lvAddOns.CheckBoxes = True
        Me.lvAddOns.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvAddOns.HideSelection = False
        Me.lvAddOns.Location = New System.Drawing.Point(3, 22)
        Me.lvAddOns.Name = "lvAddOns"
        Me.lvAddOns.Size = New System.Drawing.Size(444, 177)
        Me.lvAddOns.TabIndex = 0
        Me.lvAddOns.UseCompatibleStateImageBehavior = False
        '
        'frmAddModelInformation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2626, 1183)
        Me.Controls.Add(Me.pnlVariationBuilder)
        Me.Controls.Add(Me.btnPublishFabricColor)
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
        Me.pnlVariationBuilder.ResumeLayout(False)
        Me.tlVariationRoot.ResumeLayout(False)
        CType(Me.dgMatrix, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tlRight.ResumeLayout(False)
        Me.grpPreview.ResumeLayout(False)
        Me.grpPreview.PerformLayout()
        CType(Me.picVariation, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picParent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudFabricAdder, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudBasePrice, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpAddOns.ResumeLayout(False)
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
    Friend WithEvents btnPublishFabricColor As Windows.Forms.Button
    Friend WithEvents pnlVariationBuilder As Windows.Forms.Panel
    Friend WithEvents tlVariationRoot As Windows.Forms.TableLayoutPanel
    Friend WithEvents dgMatrix As Windows.Forms.DataGridView
    Friend WithEvents tlRight As Windows.Forms.TableLayoutPanel
    Friend WithEvents grpPreview As Windows.Forms.GroupBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents nudBasePrice As Windows.Forms.NumericUpDown
    Friend WithEvents txtTitlePreview As Windows.Forms.TextBox
    Friend WithEvents lblComputerPrice As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents nudFabricAdder As Windows.Forms.NumericUpDown
    Friend WithEvents picVariation As Windows.Forms.PictureBox
    Friend WithEvents picParent As Windows.Forms.PictureBox
    Friend WithEvents btnSet As Windows.Forms.Button
    Friend WithEvents TextBox1txtVariationIMageURL As Windows.Forms.TextBox
    Friend WithEvents grpAddOns As Windows.Forms.GroupBox
    Friend WithEvents lvAddOns As Windows.Forms.ListView
End Class
