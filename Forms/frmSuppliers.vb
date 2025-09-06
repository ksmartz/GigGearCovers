Imports System.Data.SqlClient
Imports System.Windows.Forms
Public Class frmSuppliers
    Private Sub frmSuppliers_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtWebsite.Text = "https://"
        LoadSuppliers()
    End Sub
    Private Sub LoadSuppliers()
        cmbSuppliers.Items.Clear()
        Using conn As SqlConnection = DbConnectionManager.GetConnection()
            conn.Open() ' <-- THIS LINE IS REQUIRED
            Dim cmd As New SqlCommand("SELECT PK_SupplierNameID, CompanyName FROM SupplierInformation", conn)
            Using reader As SqlDataReader = cmd.ExecuteReader()
                While reader.Read()
                    cmbSuppliers.Items.Add(New With {
                    .SupplierID = reader("PK_SupplierNameID"),
                    .CompanyName = reader("CompanyName").ToString()
                })
                End While
            End Using
        End Using
        cmbSuppliers.DisplayMember = "CompanyName"
        cmbSuppliers.ValueMember = "PK_SupplierNameID"
    End Sub

    Private Sub cmbSuppliers_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSuppliers.SelectedIndexChanged
        If cmbSuppliers.SelectedItem Is Nothing Then Exit Sub
        Dim selected = DirectCast(cmbSuppliers.SelectedItem, Object)
        Dim supplierId As Integer = selected.SupplierID
        LoadSupplierDetails(supplierId)
    End Sub

    Private Sub LoadSupplierDetails(supplierId As Integer)
        Using conn As SqlConnection = DbConnectionManager.GetConnection()
            Dim cmd As New SqlCommand("SELECT * FROM SupplierInformation WHERE PK_SupplierNameID = @PK_SupplierNameID", conn)
            cmd.Parameters.AddWithValue("@PK_SupplierNameID", supplierId)
            Using reader As SqlDataReader = cmd.ExecuteReader()
                If reader.Read() Then
                    txtCompanyName.Text = reader("CompanyName").ToString()
                    txtContact1.Text = reader("Contact1").ToString()
                    txtContact2.Text = reader("Contact2").ToString()
                    mtbPhone1.Text = reader("Phone1").ToString()
                    mtbPhone2.Text = reader("Phone2").ToString()
                    txtEmail1.Text = reader("Email1").ToString()
                    txtEmail2.Text = reader("Email2").ToString()
                    txtWebsite.Text = reader("Website").ToString()
                    txtAddress1.Text = reader("Address1").ToString()
                    txtAddress2.Text = reader("Address2").ToString()
                    txtCity.Text = reader("City").ToString()
                    txtState.Text = reader("State").ToString()
                    txtZipPostal.Text = reader("ZipPostal").ToString()
                End If
            End Using
        End Using
    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        Dim supplier As New SupplierInformation With {
        .CompanyName = txtCompanyName.Text,
        .Contact1 = txtContact1.Text,
        .Contact2 = txtContact2.Text,
        .Phone1 = mtbPhone1.Text,
        .Phone2 = mtbPhone2.Text,
        .Email1 = txtEmail1.Text,
        .Email2 = txtEmail2.Text,
        .Website = txtWebsite.Text,
        .Address1 = txtAddress1.Text,
        .Address2 = txtAddress2.Text,
        .City = txtCity.Text,
        .State = txtState.Text,
        .ZipPostal = txtZipPostal.Text
    }

        Using conn As SqlConnection = DbConnectionManager.GetConnection()
            conn.Open() ' <-- REQUIRED!
            If cmbSuppliers.SelectedItem IsNot Nothing Then
                ' Update existing supplier
                Dim selected = DirectCast(cmbSuppliers.SelectedItem, Object)
                Dim supplierId As Integer = selected.SupplierID

                ' Check for duplicate company name (excluding current supplier)
                Dim checkCmd As New SqlCommand("SELECT COUNT(*) FROM SupplierInformation WHERE CompanyName = @CompanyName AND PK_SupplierNameID <> @SupplierID", conn)
                checkCmd.Parameters.AddWithValue("@CompanyName", supplier.CompanyName)
                checkCmd.Parameters.AddWithValue("@SupplierID", supplierId)
                Dim exists As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                If exists > 0 Then
                    MessageBox.Show("A supplier with this company name already exists.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim updateCmd As New SqlCommand("UPDATE SupplierInformation SET CompanyName=@CompanyName, Contact1=@Contact1, Contact2=@Contact2, Phone1=@Phone1, Phone2=@Phone2, Email1=@Email1, Email2=@Email2, Website=@Website, Address1=@Address1, Address2=@Address2, City=@City, State=@State, ZipPostal=@ZipPostal WHERE PK_SupplierNameID=@SupplierID", conn)
                updateCmd.Parameters.AddWithValue("@CompanyName", supplier.CompanyName)
                updateCmd.Parameters.AddWithValue("@Contact1", supplier.Contact1)
                updateCmd.Parameters.AddWithValue("@Contact2", supplier.Contact2)
                updateCmd.Parameters.AddWithValue("@Phone1", supplier.Phone1)
                updateCmd.Parameters.AddWithValue("@Phone2", supplier.Phone2)
                updateCmd.Parameters.AddWithValue("@Email1", supplier.Email1)
                updateCmd.Parameters.AddWithValue("@Email2", supplier.Email2)
                updateCmd.Parameters.AddWithValue("@Website", supplier.Website)
                updateCmd.Parameters.AddWithValue("@Address1", supplier.Address1)
                updateCmd.Parameters.AddWithValue("@Address2", supplier.Address2)
                updateCmd.Parameters.AddWithValue("@City", supplier.City)
                updateCmd.Parameters.AddWithValue("@State", supplier.State)
                updateCmd.Parameters.AddWithValue("@ZipPostal", supplier.ZipPostal)
                updateCmd.Parameters.AddWithValue("@SupplierID", supplierId)
                updateCmd.ExecuteNonQuery()
                ' After MessageBox.Show("Supplier added successfully.")
                MessageBox.Show("Supplier added successfully.")
                ClearInputFields()
            Else
                ' Insert new supplier
                Dim checkCmd As New SqlCommand("SELECT COUNT(*) FROM SupplierInformation WHERE CompanyName = @CompanyName", conn)
                checkCmd.Parameters.AddWithValue("@CompanyName", supplier.CompanyName)
                Dim exists As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                If exists > 0 Then
                    MessageBox.Show("A supplier with this company name already exists.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim cmd As New SqlCommand("INSERT INTO SupplierInformation (CompanyName, Contact1, Contact2, Phone1, Phone2, Email1, Email2, Website, Address1, Address2, City, State, ZipPostal) VALUES (@CompanyName, @Contact1, @Contact2, @Phone1, @Phone2, @Email1, @Email2, @Website, @Address1, @Address2, @City, @State, @ZipPostal)", conn)
                cmd.Parameters.AddWithValue("@CompanyName", supplier.CompanyName)
                cmd.Parameters.AddWithValue("@Contact1", supplier.Contact1)
                cmd.Parameters.AddWithValue("@Contact2", supplier.Contact2)
                cmd.Parameters.AddWithValue("@Phone1", supplier.Phone1)
                cmd.Parameters.AddWithValue("@Phone2", supplier.Phone2)
                cmd.Parameters.AddWithValue("@Email1", supplier.Email1)
                cmd.Parameters.AddWithValue("@Email2", supplier.Email2)
                cmd.Parameters.AddWithValue("@Website", supplier.Website)
                cmd.Parameters.AddWithValue("@Address1", supplier.Address1)
                cmd.Parameters.AddWithValue("@Address2", supplier.Address2)
                cmd.Parameters.AddWithValue("@City", supplier.City)
                cmd.Parameters.AddWithValue("@State", supplier.State)
                cmd.Parameters.AddWithValue("@ZipPostal", supplier.ZipPostal)
                cmd.ExecuteNonQuery()
                ' After MessageBox.Show("Supplier added successfully.")
                MessageBox.Show("Supplier added successfully.")
                ClearInputFields()
            End If
        End Using

        LoadSuppliers()
    End Sub
    Private Sub ClearInputFields()
        txtCompanyName.Clear()
        txtContact1.Clear()
        txtContact2.Clear()
        mtbPhone1.Clear()
        mtbPhone2.Clear()
        txtEmail1.Clear()
        txtEmail2.Clear()
        txtWebsite.Text = "https://"
        txtAddress1.Clear()
        txtAddress2.Clear()
        txtCity.Clear()
        txtState.Clear()
        txtZipPostal.Clear()
        cmbSuppliers.SelectedIndex = -1
    End Sub

    Private Sub MaskedTextBox1_MaskInputRejected(sender As Object, e As MaskInputRejectedEventArgs) Handles mtbPhone1.MaskInputRejected

    End Sub

    Private Sub txtEmail1_TextChanged(sender As Object, e As EventArgs) Handles txtEmail1.TextChanged

    End Sub
End Class