<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class formAddModels
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
        Me.cmbManufacturer = New System.Windows.Forms.ComboBox()
        Me.lblSelectManufacturer = New System.Windows.Forms.Label()
        Me.cmbSeries = New System.Windows.Forms.ComboBox()
        Me.lblSelectSeries = New System.Windows.Forms.Label()
        Me.dgvModels = New System.Windows.Forms.DataGridView()
        Me.btnReturnToDashboard = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.txtEquipmentType = New System.Windows.Forms.TextBox()
        Me.btnAddManufacturerSeriesInfo = New System.Windows.Forms.Button()
        CType(Me.dgvModels, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbManufacturer
        '
        Me.cmbManufacturer.FormattingEnabled = True
        Me.cmbManufacturer.Location = New System.Drawing.Point(35, 72)
        Me.cmbManufacturer.Name = "cmbManufacturer"
        Me.cmbManufacturer.Size = New System.Drawing.Size(137, 28)
        Me.cmbManufacturer.TabIndex = 0
        '
        'lblSelectManufacturer
        '
        Me.lblSelectManufacturer.AutoSize = True
        Me.lblSelectManufacturer.Location = New System.Drawing.Point(31, 42)
        Me.lblSelectManufacturer.Name = "lblSelectManufacturer"
        Me.lblSelectManufacturer.Size = New System.Drawing.Size(153, 20)
        Me.lblSelectManufacturer.TabIndex = 1
        Me.lblSelectManufacturer.Text = "Select Manufacturer"
        '
        'cmbSeries
        '
        Me.cmbSeries.FormattingEnabled = True
        Me.cmbSeries.Location = New System.Drawing.Point(246, 72)
        Me.cmbSeries.Name = "cmbSeries"
        Me.cmbSeries.Size = New System.Drawing.Size(167, 28)
        Me.cmbSeries.TabIndex = 2
        '
        'lblSelectSeries
        '
        Me.lblSelectSeries.AutoSize = True
        Me.lblSelectSeries.Location = New System.Drawing.Point(242, 42)
        Me.lblSelectSeries.Name = "lblSelectSeries"
        Me.lblSelectSeries.Size = New System.Drawing.Size(103, 20)
        Me.lblSelectSeries.TabIndex = 3
        Me.lblSelectSeries.Text = "Select Series"
        '
        'dgvModels
        '
        Me.dgvModels.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvModels.Location = New System.Drawing.Point(35, 217)
        Me.dgvModels.Name = "dgvModels"
        Me.dgvModels.RowHeadersWidth = 62
        Me.dgvModels.RowTemplate.Height = 28
        Me.dgvModels.Size = New System.Drawing.Size(2587, 248)
        Me.dgvModels.TabIndex = 4
        '
        'btnReturnToDashboard
        '
        Me.btnReturnToDashboard.Location = New System.Drawing.Point(35, 141)
        Me.btnReturnToDashboard.Name = "btnReturnToDashboard"
        Me.btnReturnToDashboard.Size = New System.Drawing.Size(227, 49)
        Me.btnReturnToDashboard.TabIndex = 5
        Me.btnReturnToDashboard.Text = "Dashboard"
        Me.btnReturnToDashboard.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(298, 144)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(115, 46)
        Me.btnSave.TabIndex = 6
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'txtEquipmentType
        '
        Me.txtEquipmentType.Location = New System.Drawing.Point(464, 74)
        Me.txtEquipmentType.Name = "txtEquipmentType"
        Me.txtEquipmentType.Size = New System.Drawing.Size(143, 26)
        Me.txtEquipmentType.TabIndex = 7
        '
        'btnAddManufacturerSeriesInfo
        '
        Me.btnAddManufacturerSeriesInfo.Location = New System.Drawing.Point(872, 67)
        Me.btnAddManufacturerSeriesInfo.Name = "btnAddManufacturerSeriesInfo"
        Me.btnAddManufacturerSeriesInfo.Size = New System.Drawing.Size(179, 65)
        Me.btnAddManufacturerSeriesInfo.TabIndex = 8
        Me.btnAddManufacturerSeriesInfo.Text = "New Manufacturer"
        Me.btnAddManufacturerSeriesInfo.UseVisualStyleBackColor = True
        '
        'formAddModels
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2615, 920)
        Me.Controls.Add(Me.btnAddManufacturerSeriesInfo)
        Me.Controls.Add(Me.txtEquipmentType)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnReturnToDashboard)
        Me.Controls.Add(Me.dgvModels)
        Me.Controls.Add(Me.lblSelectSeries)
        Me.Controls.Add(Me.cmbSeries)
        Me.Controls.Add(Me.lblSelectManufacturer)
        Me.Controls.Add(Me.cmbManufacturer)
        Me.Name = "formAddModels"
        Me.Text = "formAddModels"
        CType(Me.dgvModels, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmbManufacturer As Windows.Forms.ComboBox
    Friend WithEvents lblSelectManufacturer As Windows.Forms.Label
    Friend WithEvents cmbSeries As Windows.Forms.ComboBox
    Friend WithEvents lblSelectSeries As Windows.Forms.Label
    Friend WithEvents dgvModels As Windows.Forms.DataGridView
    Friend WithEvents btnReturnToDashboard As Windows.Forms.Button
    Friend WithEvents btnSave As Windows.Forms.Button
    Friend WithEvents txtEquipmentType As Windows.Forms.TextBox
    Friend WithEvents btnAddManufacturerSeriesInfo As Windows.Forms.Button
End Class
