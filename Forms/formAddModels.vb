' Purpose: Allow user to select Manufacturer and Series, then add/edit multiple Model records in a DataGridView with ComboBox columns for lookup fields.
' Dependencies: Imports System.Data.SqlClient, System.Windows.Forms, DbConnectionManager
' Current date: 2025-09-30

Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class formAddModels
    ' DataTables for lookup values
    Private dtAngleTypes As DataTable
    Private dtAmpHandleLocations As DataTable
    Public Sub New()
        InitializeComponent()
        ' Ensure the Load event is wired up
        AddHandler Me.Load, AddressOf formAddModels_Load
        AddHandler cmbManufacturer.SelectedIndexChanged, AddressOf cmbManufacturer_SelectedIndexChanged
        AddHandler cmbSeries.SelectedIndexChanged, AddressOf cmbSeries_SelectedIndexChanged

    End Sub
    ' Loads lookup data and sets up the grid
    Private Sub formAddModels_Load(sender As Object, e As EventArgs)
        Try
            ' Load Manufacturers
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

            ' Load Angle Types
            Using conn = DbConnectionManager.CreateOpenConnection()
                dtAngleTypes = New DataTable()
                Using cmd As New SqlCommand("SELECT PK_AngleTypeId, AngleTypeName FROM AngleType ORDER BY AngleTypeName", conn)
                    dtAngleTypes.Load(cmd.ExecuteReader())
                End Using
            End Using

            ' Load Amp Handle Locations
            Using conn = DbConnectionManager.CreateOpenConnection()
                dtAmpHandleLocations = New DataTable()
                Using cmd As New SqlCommand("SELECT PK_AmpHandleLocationId, AmpHandleLocationName FROM AmpHandleLocation ORDER BY AmpHandleLocationName", conn)
                    dtAmpHandleLocations.Load(cmd.ExecuteReader())
                End Using
            End Using

            SetupModelGrid()
        Catch ex As Exception
            MessageBox.Show("Error loading dropdowns: " & ex.Message)
        End Try
    End Sub

    ' Loads Series for selected Manufacturer
    Private Sub cmbManufacturer_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmbManufacturer.SelectedIndex = -1 OrElse cmbManufacturer.SelectedValue Is Nothing Then
            cmbSeries.DataSource = Nothing
            Return
        End If
        Try
            ' >>> changed
            Dim manuId As Object = cmbManufacturer.SelectedValue
            If TypeOf manuId Is DataRowView Then
                manuId = DirectCast(manuId, DataRowView)("PK_manufacturerId")
            End If
            ' <<< end changed

            Using conn = DbConnectionManager.CreateOpenConnection()
                Dim dt As New DataTable()
                Using cmd As New SqlCommand("SELECT PK_seriesId, seriesName, FK_equipmentTypeId FROM ModelSeries WHERE FK_manufacturerId = @manuId ORDER BY seriesName", conn)
                    cmd.Parameters.AddWithValue("@manuId", manuId)
                    dt.Load(cmd.ExecuteReader())
                End Using
                cmbSeries.DataSource = dt
                cmbSeries.DisplayMember = "seriesName"
                cmbSeries.ValueMember = "PK_seriesId"
                cmbSeries.SelectedIndex = -1
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading series: " & ex.Message)
        End Try
    End Sub

    ' Loads equipment type for selected Series
    Private Sub cmbSeries_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cmbSeries.SelectedIndex = -1 OrElse cmbSeries.SelectedValue Is Nothing Then
            txtEquipmentType.Text = ""
            Return
        End If
        Try
            Dim drv As DataRowView = TryCast(cmbSeries.SelectedItem, DataRowView)
            If drv IsNot Nothing AndAlso Not IsDBNull(drv("FK_equipmentTypeId")) Then
                Dim eqId = Convert.ToInt32(drv("FK_equipmentTypeId"))
                Using conn = DbConnectionManager.CreateOpenConnection()
                    Using cmd As New SqlCommand("SELECT equipmentTypeName FROM ModelEquipmentTypes WHERE PK_equipmentTypeId = @eqId", conn)
                        cmd.Parameters.AddWithValue("@eqId", eqId)
                        Dim result = cmd.ExecuteScalar()
                        txtEquipmentType.Text = If(result IsNot Nothing, result.ToString(), "")
                    End Using
                End Using
            Else
                txtEquipmentType.Text = ""
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading equipment type: " & ex.Message)
        End Try
    End Sub

    ' Sets up DataGridView columns for model entry
    Private Sub SetupModelGrid()
        dgvModels.Columns.Clear()
        dgvModels.AllowUserToAddRows = True
        dgvModels.AllowUserToDeleteRows = True
        dgvModels.AutoGenerateColumns = False

        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "modelName", .HeaderText = "Model Name", .DataPropertyName = "modelName"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "width", .HeaderText = "Width", .DataPropertyName = "width"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "depth", .HeaderText = "Depth", .DataPropertyName = "depth"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "height", .HeaderText = "Height", .DataPropertyName = "height"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "optionalHeight", .HeaderText = "Optional Height", .DataPropertyName = "optionalHeight"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "optionalDepth", .HeaderText = "Optional Depth", .DataPropertyName = "optionalDepth"})

        Dim angleTypeCol As New DataGridViewComboBoxColumn() With {
            .Name = "FK_angleTypeId",
            .HeaderText = "Angle Type",
            .DataPropertyName = "FK_angleTypeId",
            .DataSource = dtAngleTypes,
            .DisplayMember = "AngleTypeName",
            .ValueMember = "PK_AngleTypeId"
        }
        dgvModels.Columns.Add(angleTypeCol)

        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "tahWidth", .HeaderText = "TAH Width", .DataPropertyName = "tahWidth"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "tahHeight", .HeaderText = "TAH Height", .DataPropertyName = "tahHeight"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "sahHeight", .HeaderText = "SAH Height", .DataPropertyName = "sahHeight"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "sahWidth", .HeaderText = "SAH Width", .DataPropertyName = "sahWidth"})

        Dim ampHandleCol As New DataGridViewComboBoxColumn() With {
            .Name = "FK_ampHandleLocationId",
            .HeaderText = "Amp Handle Location",
            .DataPropertyName = "FK_ampHandleLocationId",
            .DataSource = dtAmpHandleLocations,
            .DisplayMember = "AmpHandleLocationName",
            .ValueMember = "PK_AmpHandleLocationId"
        }
        dgvModels.Columns.Add(ampHandleCol)

        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "tahRearOffset", .HeaderText = "TAH Rear Offset", .DataPropertyName = "tahRearOffset"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "sahRearOffset", .HeaderText = "SAH Rear Offset", .DataPropertyName = "sahRearOffset"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "sahTopDownOffset", .HeaderText = "SAH TopDown Offset", .DataPropertyName = "sahTopDownOffset"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "musicRestDesign", .HeaderText = "Music Rest Design", .DataPropertyName = "musicRestDesign"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "chart_Template", .HeaderText = "Chart Template", .DataPropertyName = "chart_Template"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "notes", .HeaderText = "Notes", .DataPropertyName = "notes"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "totalFabricSquareInches", .HeaderText = "Total Fabric SqIn", .DataPropertyName = "totalFabricSquareInches"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "parentSku", .HeaderText = "Parent SKU", .DataPropertyName = "parentSku"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "WooProductId", .HeaderText = "Woo Product Id", .DataPropertyName = "WooProductId"})
        dgvModels.Columns.Add(New DataGridViewCheckBoxColumn() With {.Name = "onReverb", .HeaderText = "On Reverb", .DataPropertyName = "onReverb"})
    End Sub

    ' Saves all new models from the DataGridView to the database
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If cmbManufacturer.SelectedIndex = -1 OrElse cmbSeries.SelectedIndex = -1 Then
            MessageBox.Show("Please select a Manufacturer and Series.")
            Return
        End If

        If dgvModels.Rows.Count = 0 Then
            MessageBox.Show("Please add at least one model.")
            Return
        End If

        Dim insertedCount As Integer = 0
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                For Each row As DataGridViewRow In dgvModels.Rows
                    If row.IsNewRow Then Continue For
                    If String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("modelName").Value)) OrElse
                       String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("width").Value)) OrElse
                       String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("depth").Value)) OrElse
                       String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("height").Value)) Then
                        Continue For
                    End If

                    Dim sql As String = "INSERT INTO Model (FK_seriesId, modelName, width, depth, height, optionalHeight, optionalDepth, FK_angleTypeId, tahWidth, tahHeight, sahHeight, sahWidth, FK_ampHandleLocationId, tahRearOffset, sahRearOffset, sahTopDownOffset, musicRestDesign, chart_Template, notes, totalFabricSquareInches, parentSku, WooProductId, onReverb) " &
                                        "VALUES (@seriesId, @modelName, @width, @depth, @height, @optionalHeight, @optionalDepth, @FK_angleTypeId, @tahWidth, @tahHeight, @sahHeight, @sahWidth, @FK_ampHandleLocationId, @tahRearOffset, @sahRearOffset, @sahTopDownOffset, @musicRestDesign, @chart_Template, @notes, @totalFabricSquareInches, @parentSku, @WooProductId, @onReverb)"
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.Parameters.AddWithValue("@seriesId", cmbSeries.SelectedValue)
                        cmd.Parameters.AddWithValue("@modelName", Convert.ToString(row.Cells("modelName").Value))
                        cmd.Parameters.AddWithValue("@width", Convert.ToDecimal(row.Cells("width").Value))
                        cmd.Parameters.AddWithValue("@depth", Convert.ToDecimal(row.Cells("depth").Value))
                        cmd.Parameters.AddWithValue("@height", Convert.ToDecimal(row.Cells("height").Value))
                        cmd.Parameters.AddWithValue("@optionalHeight", If(IsDBNull(row.Cells("optionalHeight").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("optionalHeight").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("optionalHeight").Value)))
                        cmd.Parameters.AddWithValue("@optionalDepth", If(IsDBNull(row.Cells("optionalDepth").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("optionalDepth").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("optionalDepth").Value)))
                        cmd.Parameters.AddWithValue("@FK_angleTypeId", If(IsDBNull(row.Cells("FK_angleTypeId").Value), DBNull.Value, row.Cells("FK_angleTypeId").Value))
                        cmd.Parameters.AddWithValue("@tahWidth", If(IsDBNull(row.Cells("tahWidth").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("tahWidth").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("tahWidth").Value)))
                        cmd.Parameters.AddWithValue("@tahHeight", If(IsDBNull(row.Cells("tahHeight").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("tahHeight").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("tahHeight").Value)))
                        cmd.Parameters.AddWithValue("@sahHeight", If(IsDBNull(row.Cells("sahHeight").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("sahHeight").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("sahHeight").Value)))
                        cmd.Parameters.AddWithValue("@sahWidth", If(IsDBNull(row.Cells("sahWidth").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("sahWidth").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("sahWidth").Value)))
                        cmd.Parameters.AddWithValue("@FK_ampHandleLocationId", If(IsDBNull(row.Cells("FK_ampHandleLocationId").Value), DBNull.Value, row.Cells("FK_ampHandleLocationId").Value))
                        cmd.Parameters.AddWithValue("@tahRearOffset", If(IsDBNull(row.Cells("tahRearOffset").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("tahRearOffset").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("tahRearOffset").Value)))
                        cmd.Parameters.AddWithValue("@sahRearOffset", If(IsDBNull(row.Cells("sahRearOffset").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("sahRearOffset").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("sahRearOffset").Value)))
                        cmd.Parameters.AddWithValue("@sahTopDownOffset", If(IsDBNull(row.Cells("sahTopDownOffset").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("sahTopDownOffset").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("sahTopDownOffset").Value)))
                        cmd.Parameters.AddWithValue("@musicRestDesign", If(IsDBNull(row.Cells("musicRestDesign").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("musicRestDesign").Value)), DBNull.Value, row.Cells("musicRestDesign").Value))
                        cmd.Parameters.AddWithValue("@chart_Template", If(IsDBNull(row.Cells("chart_Template").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("chart_Template").Value)), DBNull.Value, row.Cells("chart_Template").Value))
                        cmd.Parameters.AddWithValue("@notes", If(IsDBNull(row.Cells("notes").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("notes").Value)), DBNull.Value, row.Cells("notes").Value))
                        cmd.Parameters.AddWithValue("@totalFabricSquareInches", If(IsDBNull(row.Cells("totalFabricSquareInches").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("totalFabricSquareInches").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("totalFabricSquareInches").Value)))
                        cmd.Parameters.AddWithValue("@parentSku", If(IsDBNull(row.Cells("parentSku").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("parentSku").Value)), DBNull.Value, row.Cells("parentSku").Value))
                        cmd.Parameters.AddWithValue("@WooProductId", If(IsDBNull(row.Cells("WooProductId").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("WooProductId").Value)), DBNull.Value, Convert.ToInt32(row.Cells("WooProductId").Value)))
                        cmd.Parameters.AddWithValue("@onReverb", If(IsDBNull(row.Cells("onReverb").Value), False, Convert.ToBoolean(row.Cells("onReverb").Value)))
                        insertedCount += cmd.ExecuteNonQuery()
                    End Using
                Next
            End Using
            MessageBox.Show($"{insertedCount} model(s) added successfully!")
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Error saving models: " & ex.Message)
        End Try
    End Sub

    ' Handles return to dashboard
    Private Sub btnReturnToDashboard_Click(sender As Object, e As EventArgs) Handles btnReturnToDashboard.Click
        Dim dashboard As New formDashboard()
        formDashboard.Show()
        Me.Close()
    End Sub
    ' Purpose: Opens FormAddManufacturerSeries and refreshes cmbManufacturer after closing.
    ' Dependencies: Imports System.Windows.Forms
    ' Current date: 2025-09-30

    Private Sub btnAddManufacturerSeriesInfo_Click(sender As Object, e As EventArgs) Handles btnAddManufacturerSeriesInfo.Click
        ' >>> changed
        Dim selectedManuId As Integer? = Nothing
        If cmbManufacturer.SelectedValue IsNot Nothing AndAlso Integer.TryParse(cmbManufacturer.SelectedValue.ToString(), selectedManuId) Then
            ' keep current selection
        End If
        Using frm As New FormAddManufacturerSeries()
            frm.ShowDialog(Me)
        End Using
        RefreshManufacturers(selectedManuId)
        ' <<< end changed
    End Sub

    ' Purpose: Reloads manufacturers into cmbManufacturer.
    ' Dependencies: Imports System.Data.SqlClient, DbConnectionManager
    ' Current date: 2025-09-30
    Private Sub RefreshManufacturers(Optional selectManufacturerId As Integer? = Nothing)
        ' >>> changed
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                Dim dt As New DataTable()
                Using cmd As New SqlCommand("SELECT PK_manufacturerId, manufacturerName FROM ModelManufacturers ORDER BY manufacturerName", conn)
                    dt.Load(cmd.ExecuteReader())
                End Using
                cmbManufacturer.DataSource = dt
                cmbManufacturer.DisplayMember = "manufacturerName"
                cmbManufacturer.ValueMember = "PK_manufacturerId"
                If selectManufacturerId.HasValue Then
                    cmbManufacturer.SelectedValue = selectManufacturerId.Value
                Else
                    cmbManufacturer.SelectedIndex = -1
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error refreshing manufacturers: " & ex.Message)
        End Try
        ' <<< end changed
    End Sub

    Private Sub formAddModels_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub dgvModels_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvModels.CellContentClick

    End Sub
End Class