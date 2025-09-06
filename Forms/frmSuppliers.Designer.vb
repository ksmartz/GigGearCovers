<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSuppliers
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
        Me.btnOk = New System.Windows.Forms.Button()
        Me.txtCompanyName = New System.Windows.Forms.TextBox()
        Me.txtContact1 = New System.Windows.Forms.TextBox()
        Me.txtContact2 = New System.Windows.Forms.TextBox()
        Me.txtEmail1 = New System.Windows.Forms.TextBox()
        Me.txtEmail2 = New System.Windows.Forms.TextBox()
        Me.txtAddress1 = New System.Windows.Forms.TextBox()
        Me.txtCity = New System.Windows.Forms.TextBox()
        Me.txtState = New System.Windows.Forms.TextBox()
        Me.txtZipPostal = New System.Windows.Forms.TextBox()
        Me.txtWebsite = New System.Windows.Forms.TextBox()
        Me.txtAddress2 = New System.Windows.Forms.TextBox()
        Me.lblCompanyName = New System.Windows.Forms.Label()
        Me.lblContact1 = New System.Windows.Forms.Label()
        Me.lblContact2 = New System.Windows.Forms.Label()
        Me.lblPhone1 = New System.Windows.Forms.Label()
        Me.lblPhone2 = New System.Windows.Forms.Label()
        Me.lblEmail1 = New System.Windows.Forms.Label()
        Me.lblEmail2 = New System.Windows.Forms.Label()
        Me.lblAddress1 = New System.Windows.Forms.Label()
        Me.lblAddress2 = New System.Windows.Forms.Label()
        Me.lblCity = New System.Windows.Forms.Label()
        Me.lblState = New System.Windows.Forms.Label()
        Me.lblZipPostal = New System.Windows.Forms.Label()
        Me.lblWebsite = New System.Windows.Forms.Label()
        Me.cmbSuppliers = New System.Windows.Forms.ComboBox()
        Me.mtbPhone1 = New System.Windows.Forms.MaskedTextBox()
        Me.mtbPhone2 = New System.Windows.Forms.MaskedTextBox()
        Me.SuspendLayout()
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(183, 571)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(260, 77)
        Me.btnOk.TabIndex = 0
        Me.btnOk.Text = "OK"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'txtCompanyName
        '
        Me.txtCompanyName.Location = New System.Drawing.Point(22, 79)
        Me.txtCompanyName.Name = "txtCompanyName"
        Me.txtCompanyName.Size = New System.Drawing.Size(180, 26)
        Me.txtCompanyName.TabIndex = 1
        '
        'txtContact1
        '
        Me.txtContact1.Location = New System.Drawing.Point(22, 201)
        Me.txtContact1.Name = "txtContact1"
        Me.txtContact1.Size = New System.Drawing.Size(180, 26)
        Me.txtContact1.TabIndex = 2
        '
        'txtContact2
        '
        Me.txtContact2.Location = New System.Drawing.Point(36, 279)
        Me.txtContact2.Name = "txtContact2"
        Me.txtContact2.Size = New System.Drawing.Size(180, 26)
        Me.txtContact2.TabIndex = 3
        '
        'txtEmail1
        '
        Me.txtEmail1.Location = New System.Drawing.Point(505, 201)
        Me.txtEmail1.Name = "txtEmail1"
        Me.txtEmail1.Size = New System.Drawing.Size(180, 26)
        Me.txtEmail1.TabIndex = 6
        '
        'txtEmail2
        '
        Me.txtEmail2.Location = New System.Drawing.Point(516, 279)
        Me.txtEmail2.Name = "txtEmail2"
        Me.txtEmail2.Size = New System.Drawing.Size(180, 26)
        Me.txtEmail2.TabIndex = 7
        '
        'txtAddress1
        '
        Me.txtAddress1.Location = New System.Drawing.Point(226, 79)
        Me.txtAddress1.Name = "txtAddress1"
        Me.txtAddress1.Size = New System.Drawing.Size(180, 26)
        Me.txtAddress1.TabIndex = 8
        '
        'txtCity
        '
        Me.txtCity.Location = New System.Drawing.Point(652, 56)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(180, 26)
        Me.txtCity.TabIndex = 9
        '
        'txtState
        '
        Me.txtState.Location = New System.Drawing.Point(768, 125)
        Me.txtState.Name = "txtState"
        Me.txtState.Size = New System.Drawing.Size(180, 26)
        Me.txtState.TabIndex = 10
        '
        'txtZipPostal
        '
        Me.txtZipPostal.Location = New System.Drawing.Point(768, 238)
        Me.txtZipPostal.Name = "txtZipPostal"
        Me.txtZipPostal.Size = New System.Drawing.Size(180, 26)
        Me.txtZipPostal.TabIndex = 11
        '
        'txtWebsite
        '
        Me.txtWebsite.Location = New System.Drawing.Point(666, 361)
        Me.txtWebsite.Name = "txtWebsite"
        Me.txtWebsite.Size = New System.Drawing.Size(180, 26)
        Me.txtWebsite.TabIndex = 12
        '
        'txtAddress2
        '
        Me.txtAddress2.Location = New System.Drawing.Point(424, 79)
        Me.txtAddress2.Name = "txtAddress2"
        Me.txtAddress2.Size = New System.Drawing.Size(180, 26)
        Me.txtAddress2.TabIndex = 13
        '
        'lblCompanyName
        '
        Me.lblCompanyName.AutoSize = True
        Me.lblCompanyName.Location = New System.Drawing.Point(55, 38)
        Me.lblCompanyName.Name = "lblCompanyName"
        Me.lblCompanyName.Size = New System.Drawing.Size(122, 20)
        Me.lblCompanyName.TabIndex = 14
        Me.lblCompanyName.Text = "Company Name"
        '
        'lblContact1
        '
        Me.lblContact1.AutoSize = True
        Me.lblContact1.Location = New System.Drawing.Point(55, 155)
        Me.lblContact1.Name = "lblContact1"
        Me.lblContact1.Size = New System.Drawing.Size(78, 20)
        Me.lblContact1.TabIndex = 15
        Me.lblContact1.Text = "Contact 1"
        '
        'lblContact2
        '
        Me.lblContact2.AutoSize = True
        Me.lblContact2.Location = New System.Drawing.Point(42, 244)
        Me.lblContact2.Name = "lblContact2"
        Me.lblContact2.Size = New System.Drawing.Size(78, 20)
        Me.lblContact2.TabIndex = 17
        Me.lblContact2.Text = "Contact 2"
        '
        'lblPhone1
        '
        Me.lblPhone1.AutoSize = True
        Me.lblPhone1.Location = New System.Drawing.Point(304, 166)
        Me.lblPhone1.Name = "lblPhone1"
        Me.lblPhone1.Size = New System.Drawing.Size(68, 20)
        Me.lblPhone1.TabIndex = 18
        Me.lblPhone1.Text = "Phone 1"
        '
        'lblPhone2
        '
        Me.lblPhone2.AutoSize = True
        Me.lblPhone2.Location = New System.Drawing.Point(316, 244)
        Me.lblPhone2.Name = "lblPhone2"
        Me.lblPhone2.Size = New System.Drawing.Size(68, 20)
        Me.lblPhone2.TabIndex = 19
        Me.lblPhone2.Text = "Phone 2"
        '
        'lblEmail1
        '
        Me.lblEmail1.AutoSize = True
        Me.lblEmail1.Location = New System.Drawing.Point(543, 166)
        Me.lblEmail1.Name = "lblEmail1"
        Me.lblEmail1.Size = New System.Drawing.Size(61, 20)
        Me.lblEmail1.TabIndex = 20
        Me.lblEmail1.Text = "Email 1"
        '
        'lblEmail2
        '
        Me.lblEmail2.AutoSize = True
        Me.lblEmail2.Location = New System.Drawing.Point(550, 244)
        Me.lblEmail2.Name = "lblEmail2"
        Me.lblEmail2.Size = New System.Drawing.Size(61, 20)
        Me.lblEmail2.TabIndex = 21
        Me.lblEmail2.Text = "Email 2"
        '
        'lblAddress1
        '
        Me.lblAddress1.AutoSize = True
        Me.lblAddress1.Location = New System.Drawing.Point(282, 38)
        Me.lblAddress1.Name = "lblAddress1"
        Me.lblAddress1.Size = New System.Drawing.Size(81, 20)
        Me.lblAddress1.TabIndex = 22
        Me.lblAddress1.Text = "Address 1"
        '
        'lblAddress2
        '
        Me.lblAddress2.AutoSize = True
        Me.lblAddress2.Location = New System.Drawing.Point(457, 56)
        Me.lblAddress2.Name = "lblAddress2"
        Me.lblAddress2.Size = New System.Drawing.Size(81, 20)
        Me.lblAddress2.TabIndex = 23
        Me.lblAddress2.Text = "Address 2"
        '
        'lblCity
        '
        Me.lblCity.AutoSize = True
        Me.lblCity.Location = New System.Drawing.Point(704, 19)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(35, 20)
        Me.lblCity.TabIndex = 24
        Me.lblCity.Text = "City"
        '
        'lblState
        '
        Me.lblState.AutoSize = True
        Me.lblState.Location = New System.Drawing.Point(857, 85)
        Me.lblState.Name = "lblState"
        Me.lblState.Size = New System.Drawing.Size(48, 20)
        Me.lblState.TabIndex = 25
        Me.lblState.Text = "State"
        '
        'lblZipPostal
        '
        Me.lblZipPostal.AutoSize = True
        Me.lblZipPostal.Location = New System.Drawing.Point(808, 204)
        Me.lblZipPostal.Name = "lblZipPostal"
        Me.lblZipPostal.Size = New System.Drawing.Size(79, 20)
        Me.lblZipPostal.TabIndex = 26
        Me.lblZipPostal.Text = "Zip/Postal"
        '
        'lblWebsite
        '
        Me.lblWebsite.AutoSize = True
        Me.lblWebsite.Location = New System.Drawing.Point(739, 314)
        Me.lblWebsite.Name = "lblWebsite"
        Me.lblWebsite.Size = New System.Drawing.Size(67, 20)
        Me.lblWebsite.TabIndex = 27
        Me.lblWebsite.Text = "Website"
        '
        'cmbSuppliers
        '
        Me.cmbSuppliers.FormattingEnabled = True
        Me.cmbSuppliers.Location = New System.Drawing.Point(70, 346)
        Me.cmbSuppliers.Name = "cmbSuppliers"
        Me.cmbSuppliers.Size = New System.Drawing.Size(156, 28)
        Me.cmbSuppliers.TabIndex = 28
        '
        'mtbPhone1
        '
        Me.mtbPhone1.Location = New System.Drawing.Point(267, 201)
        Me.mtbPhone1.Name = "mtbPhone1"
        Me.mtbPhone1.Size = New System.Drawing.Size(176, 26)
        Me.mtbPhone1.TabIndex = 29
        '
        'mtbPhone2
        '
        Me.mtbPhone2.Location = New System.Drawing.Point(267, 279)
        Me.mtbPhone2.Name = "mtbPhone2"
        Me.mtbPhone2.Size = New System.Drawing.Size(188, 26)
        Me.mtbPhone2.TabIndex = 30
        '
        'frmSuppliers
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1132, 766)
        Me.Controls.Add(Me.mtbPhone2)
        Me.Controls.Add(Me.mtbPhone1)
        Me.Controls.Add(Me.cmbSuppliers)
        Me.Controls.Add(Me.lblWebsite)
        Me.Controls.Add(Me.lblZipPostal)
        Me.Controls.Add(Me.lblState)
        Me.Controls.Add(Me.lblCity)
        Me.Controls.Add(Me.lblAddress2)
        Me.Controls.Add(Me.lblAddress1)
        Me.Controls.Add(Me.lblEmail2)
        Me.Controls.Add(Me.lblEmail1)
        Me.Controls.Add(Me.lblPhone2)
        Me.Controls.Add(Me.lblPhone1)
        Me.Controls.Add(Me.lblContact2)
        Me.Controls.Add(Me.lblContact1)
        Me.Controls.Add(Me.lblCompanyName)
        Me.Controls.Add(Me.txtAddress2)
        Me.Controls.Add(Me.txtWebsite)
        Me.Controls.Add(Me.txtZipPostal)
        Me.Controls.Add(Me.txtState)
        Me.Controls.Add(Me.txtCity)
        Me.Controls.Add(Me.txtAddress1)
        Me.Controls.Add(Me.txtEmail2)
        Me.Controls.Add(Me.txtEmail1)
        Me.Controls.Add(Me.txtContact2)
        Me.Controls.Add(Me.txtContact1)
        Me.Controls.Add(Me.txtCompanyName)
        Me.Controls.Add(Me.btnOk)
        Me.Name = "frmSuppliers"
        Me.Text = "frmSuppliers"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnOk As Windows.Forms.Button
    Friend WithEvents txtCompanyName As Windows.Forms.TextBox
    Friend WithEvents txtContact1 As Windows.Forms.TextBox
    Friend WithEvents txtContact2 As Windows.Forms.TextBox
    Friend WithEvents txtEmail1 As Windows.Forms.TextBox
    Friend WithEvents txtEmail2 As Windows.Forms.TextBox
    Friend WithEvents txtAddress1 As Windows.Forms.TextBox
    Friend WithEvents txtCity As Windows.Forms.TextBox
    Friend WithEvents txtState As Windows.Forms.TextBox
    Friend WithEvents txtZipPostal As Windows.Forms.TextBox
    Friend WithEvents txtWebsite As Windows.Forms.TextBox
    Friend WithEvents txtAddress2 As Windows.Forms.TextBox
    Friend WithEvents lblCompanyName As Windows.Forms.Label
    Friend WithEvents lblContact1 As Windows.Forms.Label
    Friend WithEvents lblContact2 As Windows.Forms.Label
    Friend WithEvents lblPhone1 As Windows.Forms.Label
    Friend WithEvents lblPhone2 As Windows.Forms.Label
    Friend WithEvents lblEmail1 As Windows.Forms.Label
    Friend WithEvents lblEmail2 As Windows.Forms.Label
    Friend WithEvents lblAddress1 As Windows.Forms.Label
    Friend WithEvents lblAddress2 As Windows.Forms.Label
    Friend WithEvents lblCity As Windows.Forms.Label
    Friend WithEvents lblState As Windows.Forms.Label
    Friend WithEvents lblZipPostal As Windows.Forms.Label
    Friend WithEvents lblWebsite As Windows.Forms.Label
    Friend WithEvents cmbSuppliers As Windows.Forms.ComboBox
    Friend WithEvents mtbPhone1 As Windows.Forms.MaskedTextBox
    Friend WithEvents mtbPhone2 As Windows.Forms.MaskedTextBox
End Class
