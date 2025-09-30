<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class formDashboard
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
        Me.btnAddModels = New System.Windows.Forms.Button()
        Me.btnListings = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnAddModels
        '
        Me.btnAddModels.Location = New System.Drawing.Point(102, 82)
        Me.btnAddModels.Name = "btnAddModels"
        Me.btnAddModels.Size = New System.Drawing.Size(183, 46)
        Me.btnAddModels.TabIndex = 0
        Me.btnAddModels.Text = "Add Models"
        Me.btnAddModels.UseVisualStyleBackColor = True
        '
        'btnListings
        '
        Me.btnListings.Location = New System.Drawing.Point(536, 96)
        Me.btnListings.Name = "btnListings"
        Me.btnListings.Size = New System.Drawing.Size(173, 57)
        Me.btnListings.TabIndex = 1
        Me.btnListings.Text = "Listings"
        Me.btnListings.UseVisualStyleBackColor = True
        '
        'formDashboard
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.btnListings)
        Me.Controls.Add(Me.btnAddModels)
        Me.Name = "formDashboard"
        Me.Text = "formDashboard"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnAddModels As Windows.Forms.Button
    Friend WithEvents btnListings As Windows.Forms.Button
End Class
