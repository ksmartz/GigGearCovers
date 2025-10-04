' Purpose: Add/edit manufacturers and their series, with equipment type linkage.
' Dependencies: Imports System.Data.SqlClient, System.Windows.Forms, DbConnectionManager
' Current date: 2025-09-30

Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class FormAddManufacturerSeries
    Inherits Form




    Private dtEquipmentTypes As DataTable

    ' Purpose: Form constructor for Designer-created controls.
    ' Dependencies: Imports System.Windows.Forms
    ' Current date: 2025-09-30

    Public Sub New()
        InitializeComponent()
        AddHandler cmbManufacturer.KeyDown, AddressOf cmbManufacturer_KeyDown
        AddHandler cmbManufacturer.SelectedIndexChanged, AddressOf cmbManufacturer_SelectedIndexChanged ' >>> changed
    End Sub

    ' Purpose: Handles form load; loads manufacturers, equipment types, and sets up the series grid.
    ' Dependencies: Imports System.Data.SqlClient, System.Windows.Forms, DbConnectionManager
    ' Current date: 2025-09-30

    Private Sub FormAddManufacturerSeries_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadManufacturers() ' >>> changed
        LoadEquipmentTypes() ' >>> changed
        SetupSeriesGrid() ' >>> changed
    End Sub ' <<< end changed

    Private Sub LoadManufacturers()
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                Dim dt As New DataTable()
                Using cmd As New SqlCommand("SELECT PK_manufacturerId, manufacturerName FROM ModelManufacturers ORDER BY manufacturerName", conn)
                    dt.Load(cmd.ExecuteReader())
                End Using
                cmbManufacturer.DataSource = dt
                cmbManufacturer.DisplayMember = "manufacturerName"
                cmbManufacturer.ValueMember = "PK_manufacturerId"
                cmbManufacturer.SelectedIndex = -1
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading manufacturers: " & ex.Message)
        End Try
    End Sub
    ' Purpose: Loads all series for the selected manufacturer into dgvSeries, including PK_seriesId for updates.
    ' Dependencies: Imports System.Data.SqlClient, System.Windows.Forms, DbConnectionManager
    ' Current date: 2025-09-30
    Private Sub LoadSeriesForManufacturer()
        dgvSeries.Rows.Clear()
        If cmbManufacturer.SelectedIndex = -1 Then Return

        Dim manuId As Integer = Convert.ToInt32(CType(cmbManufacturer.SelectedItem, DataRowView)("PK_manufacturerId"))
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                Dim dt As New DataTable()
                ' >>> changed: Select PK_seriesId as well
                Using cmd As New SqlCommand("SELECT PK_seriesId, seriesName, FK_equipmentTypeId FROM ModelSeries WHERE FK_manufacturerId = @manuId ORDER BY seriesName", conn)
                    cmd.Parameters.AddWithValue("@manuId", manuId)
                    dt.Load(cmd.ExecuteReader())
                End Using
                For Each row As DataRow In dt.Rows
                    dgvSeries.Rows.Add(row("PK_seriesId"), row("seriesName").ToString(), row("FK_equipmentTypeId"))
                Next
                ' <<< end changed
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading series: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadEquipmentTypes()
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                dtEquipmentTypes = New DataTable()
                Using cmd As New SqlCommand("SELECT PK_equipmentTypeId, equipmentTypeName FROM ModelEquipmentTypes ORDER BY equipmentTypeName", conn)
                    dtEquipmentTypes.Load(cmd.ExecuteReader())
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading equipment types: " & ex.Message)
        End Try
    End Sub
    Private Sub cmbManufacturer_SelectedIndexChanged(sender As Object, e As EventArgs)
        LoadSeriesForManufacturer()
    End Sub
    ' Allow adding new manufacturer name directly in ComboBox
    Private Sub cmbManufacturer_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter AndAlso Not String.IsNullOrWhiteSpace(cmbManufacturer.Text) Then
            Dim newName = cmbManufacturer.Text.Trim()
            ' Check if already exists
            Dim exists = False
            For Each item As DataRowView In cmbManufacturer.Items
                If String.Equals(item("manufacturerName").ToString(), newName, StringComparison.OrdinalIgnoreCase) Then
                    exists = True
                    cmbManufacturer.SelectedItem = item
                    Exit For
                End If
            Next
            If Not exists Then
                ' Insert new manufacturer
                Try
                    Using conn = DbConnectionManager.CreateOpenConnection()
                        Using cmd As New SqlCommand("INSERT INTO ModelManufacturers (manufacturerName) OUTPUT INSERTED.PK_manufacturerId VALUES (@name)", conn)
                            cmd.Parameters.AddWithValue("@name", newName)
                            Dim newId = Convert.ToInt32(cmd.ExecuteScalar())
                            LoadManufacturers()
                            ' Select the new manufacturer
                            For i = 0 To cmbManufacturer.Items.Count - 1
                                Dim drv = TryCast(cmbManufacturer.Items(i), DataRowView)
                                If drv IsNot Nothing AndAlso Convert.ToInt32(drv("PK_manufacturerId")) = newId Then
                                    cmbManufacturer.SelectedIndex = i
                                    Exit For
                                End If
                            Next
                        End Using
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error adding manufacturer: " & ex.Message)
                End Try
            End If
            e.Handled = True
        End If
    End Sub

    ' Purpose: Set up the dgvSeries grid with columns for PK_seriesId (hidden), seriesName, and equipment type.
    ' Dependencies: Imports System.Data.SqlClient, System.Windows.Forms
    ' Current date: 2025-09-30
    Private Sub SetupSeriesGrid()
        ' >>> changed
        dgvSeries.Columns.Clear()
        ' Add hidden PK_seriesId column
        Dim idCol As New DataGridViewTextBoxColumn() With {
        .Name = "PK_seriesId",
        .HeaderText = "ID",
        .DataPropertyName = "PK_seriesId",
        .Visible = False
    }
        dgvSeries.Columns.Add(idCol)
        dgvSeries.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "seriesName", .HeaderText = "Series Name", .DataPropertyName = "seriesName"})
        Dim eqTypeCol As New DataGridViewComboBoxColumn() With {
        .Name = "FK_equipmentTypeId",
        .HeaderText = "Equipment Type",
        .DataPropertyName = "FK_equipmentTypeId",
        .DataSource = dtEquipmentTypes,
        .DisplayMember = "equipmentTypeName",
        .ValueMember = "PK_equipmentTypeId"
    }
        dgvSeries.Columns.Add(eqTypeCol)
        ' <<< end changed
    End Sub

    ' Save series for the selected manufacturer
    ' Purpose: Save new or updated series for the selected manufacturer, using INSERT for new and UPDATE for existing.
    ' Dependencies: Imports System.Data.SqlClient, System.Windows.Forms, DbConnectionManager
    ' Current date: 2025-09-30
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' >>> changed
        If cmbManufacturer.SelectedIndex = -1 Then
            MessageBox.Show("Please select or add a manufacturer name.")
            Return
        End If
        Dim manuId = CType((CType(cmbManufacturer.SelectedItem, DataRowView))("PK_manufacturerId"), Integer)
        Dim affectedCount As Integer = 0
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                For Each row As DataGridViewRow In dgvSeries.Rows
                    If row.IsNewRow Then Continue For
                    Dim seriesName = Convert.ToString(row.Cells("seriesName").Value)
                    Dim eqTypeId = row.Cells("FK_equipmentTypeId").Value
                    If String.IsNullOrWhiteSpace(seriesName) OrElse eqTypeId Is Nothing Then Continue For

                    Dim pkObj = row.Cells("PK_seriesId").Value
                    If pkObj IsNot Nothing AndAlso Not DBNull.Value.Equals(pkObj) AndAlso Convert.ToInt32(pkObj) > 0 Then
                        ' Existing: UPDATE
                        Using cmd As New SqlCommand("UPDATE ModelSeries SET seriesName=@seriesName, FK_equipmentTypeId=@eqTypeId WHERE PK_seriesId=@id", conn)
                            cmd.Parameters.AddWithValue("@seriesName", seriesName)
                            cmd.Parameters.AddWithValue("@eqTypeId", eqTypeId)
                            cmd.Parameters.AddWithValue("@id", pkObj)
                            affectedCount += cmd.ExecuteNonQuery()
                        End Using
                    Else
                        ' New: INSERT
                        Using cmd As New SqlCommand("INSERT INTO ModelSeries (seriesName, FK_manufacturerId, FK_equipmentTypeId) VALUES (@seriesName, @manuId, @eqTypeId)", conn)
                            cmd.Parameters.AddWithValue("@seriesName", seriesName)
                            cmd.Parameters.AddWithValue("@manuId", manuId)
                            cmd.Parameters.AddWithValue("@eqTypeId", eqTypeId)
                            affectedCount += cmd.ExecuteNonQuery()
                        End Using
                    End If
                Next
            End Using
            MessageBox.Show($"{affectedCount} series saved successfully!")
        Catch ex As Exception
            MessageBox.Show("Error saving series: " & ex.Message)
        End Try
        ' <<< end changed
    End Sub

    ' Close and return to formAddModels
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FormAddManufacturerSeries_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class