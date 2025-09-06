<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmDashboard
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
        Me.btnMaterialsDataEntry = New System.Windows.Forms.Button()
        Me.btnAddSupplier = New System.Windows.Forms.Button()
        Me.btnAddModels = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnMaterialsDataEntry
        '
        Me.btnMaterialsDataEntry.Location = New System.Drawing.Point(606, 79)
        Me.btnMaterialsDataEntry.Name = "btnMaterialsDataEntry"
        Me.btnMaterialsDataEntry.Size = New System.Drawing.Size(345, 89)
        Me.btnMaterialsDataEntry.TabIndex = 0
        Me.btnMaterialsDataEntry.Text = "Materials Data Entry "
        Me.btnMaterialsDataEntry.UseVisualStyleBackColor = True
        '
        'btnAddSupplier
        '
        Me.btnAddSupplier.Location = New System.Drawing.Point(97, 45)
        Me.btnAddSupplier.Name = "btnAddSupplier"
        Me.btnAddSupplier.Size = New System.Drawing.Size(193, 56)
        Me.btnAddSupplier.TabIndex = 1
        Me.btnAddSupplier.Text = "Add Supplier"
        Me.btnAddSupplier.UseVisualStyleBackColor = True
        '
        'btnAddModels
        '
        Me.btnAddModels.Location = New System.Drawing.Point(108, 195)
        Me.btnAddModels.Name = "btnAddModels"
        Me.btnAddModels.Size = New System.Drawing.Size(202, 68)
        Me.btnAddModels.TabIndex = 2
        Me.btnAddModels.Text = "Add Model Information"
        Me.btnAddModels.UseVisualStyleBackColor = True
        '
        'frmDashboard
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1429, 803)
        Me.Controls.Add(Me.btnAddModels)
        Me.Controls.Add(Me.btnAddSupplier)
        Me.Controls.Add(Me.btnMaterialsDataEntry)
        Me.Name = "frmDashboard"
        Me.Text = "frmDashboard"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnMaterialsDataEntry As Windows.Forms.Button
    Friend WithEvents btnAddSupplier As Windows.Forms.Button
    Friend WithEvents btnAddModels As Windows.Forms.Button
End Class
