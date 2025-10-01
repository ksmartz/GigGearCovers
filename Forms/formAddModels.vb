' Purpose: Allow user to select Manufacturer and Series, then add/edit multiple Model records in a DataGridView with ComboBox columns for lookup fields.
' Dependencies: Imports System.Data.SqlClient, System.Windows.Forms, DbConnectionManager
' Current date: 2025-09-30

Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class formAddModels
    ' DataTables for lookup values
    Private dtAngleTypes As DataTable
    Private dtAmpHandleLocations As DataTable

    Private dgvContextMenu As ContextMenuStrip ' <<< Add this line
    Public Sub New()
        InitializeComponent()
        ' Ensure the Load event is wired up
        AddHandler Me.Load, AddressOf formAddModels_Load
        AddHandler cmbManufacturer.SelectedIndexChanged, AddressOf cmbManufacturer_SelectedIndexChanged
        AddHandler cmbSeries.SelectedIndexChanged, AddressOf cmbSeries_SelectedIndexChanged
        ' Wire up dirty tracking for DataGridView edits
        AddHandler dgvModels.CellValueChanged, AddressOf dgvModels_CellValueChanged
        AddHandler dgvModels.DataError, AddressOf dgvModels_DataError
    End Sub
    ' Purpose: Loads lookup data for manufacturers, angle types, and amp handle locations. Ensures "None" is always present in angle types BEFORE grid columns/rows are created.
    ' Dependencies: Imports System.Data.SqlClient, System.Windows.Forms, DbConnectionManager
    ' Current date: 2025-09-30
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
            ' >>> changed
            ' Always ensure "None" (6) is present in dtAngleTypes BEFORE any grid columns/rows are created
            If dtAngleTypes.Rows.Cast(Of DataRow)().All(Function(r) Convert.ToInt32(r("PK_AngleTypeId")) <> 6) Then
                Dim noneRow As DataRow = dtAngleTypes.NewRow()
                noneRow("PK_AngleTypeId") = 6
                noneRow("AngleTypeName") = "None"
                dtAngleTypes.Rows.InsertAt(noneRow, 0)
            End If
            ' <<< end changed

            ' Load Amp Handle Locations
            Using conn = DbConnectionManager.CreateOpenConnection()
                dtAmpHandleLocations = New DataTable()
                Using cmd As New SqlCommand("SELECT PK_AmpHandleLocationId, AmpHandleLocationName FROM AmpHandleLocation ORDER BY AmpHandleLocationName", conn)
                    dtAmpHandleLocations.Load(cmd.ExecuteReader())
                End Using
            End Using

            ' Do NOT call SetupModelGrid here; wait until a series is selected
        Catch ex As Exception
            MessageBox.Show("Error loading dropdowns: " & ex.Message)
        End Try
    End Sub
    ' >>> changed
    ' Purpose: Selects the row under the mouse when right-clicking, so the context menu acts on the correct row.
    ' Dependencies: Imports System.Windows.Forms
    ' Current date: 2025-09-30
    Private Sub dgvModels_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            Dim hit As DataGridView.HitTestInfo = dgvModels.HitTest(e.X, e.Y)
            If hit.RowIndex >= 0 Then
                dgvModels.ClearSelection()
                dgvModels.Rows(hit.RowIndex).Selected = True
                dgvModels.CurrentCell = dgvModels.Rows(hit.RowIndex).Cells(0)
            End If
        End If
    End Sub
    ' >>> changed
    ' Purpose: Deletes the selected model from the database and refreshes the grid.
    ' Dependencies: Imports System.Data.SqlClient, DbConnectionManager
    ' Current date: 2025-09-30
    Private Sub DeleteSelectedModel(sender As Object, e As EventArgs)
        If dgvModels.SelectedRows.Count = 0 Then Return

        Dim row As DataGridViewRow = dgvModels.SelectedRows(0)
        Dim parentSku As String = TryCast(row.Cells("parentSku").Value, String)
        If String.IsNullOrWhiteSpace(parentSku) Then
            MessageBox.Show("Cannot delete: Model does not have a Parent SKU.")
            Return
        End If

        If MessageBox.Show($"Are you sure you want to delete model '{row.Cells("modelName").Value}'?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then
            Return
        End If

        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                Using cmd As New SqlCommand("DELETE FROM Model WHERE parentSku = @parentSku", conn)
                    cmd.Parameters.AddWithValue("@parentSku", parentSku)
                    Dim affected = cmd.ExecuteNonQuery()
                    If affected > 0 Then
                        MessageBox.Show("Model deleted.")
                    Else
                        MessageBox.Show("Model not found or already deleted.")
                    End If
                End Using
            End Using
            RefreshModelGridFromDatabase()
        Catch ex As Exception
            MessageBox.Show("Error deleting model: " & ex.Message)
        End Try
    End Sub
    ' Purpose: When manufacturer changes, clear the DataGridView, reset the series ComboBox, and clear equipment type.
    ' Dependencies: Imports System.Data.SqlClient, System.Windows.Forms, DbConnectionManager
    ' Current date: 2025-09-30
    Private Sub cmbManufacturer_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' >>> changed
        ' Always clear the grid and related UI when manufacturer changes
        dgvModels.Rows.Clear()
        dgvModels.Columns.Clear()
        cmbSeries.DataSource = Nothing
        txtEquipmentType.Text = ""
        ' <<< end changed

        If cmbManufacturer.SelectedIndex = -1 OrElse cmbManufacturer.SelectedValue Is Nothing Then
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

    Private Sub cmbSeries_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' >>> changed
        If cmbSeries.SelectedIndex = -1 OrElse cmbSeries.SelectedValue Is Nothing Then
            txtEquipmentType.Text = ""
            dgvModels.Rows.Clear()
            dgvModels.Columns.Clear() ' Clear columns as well
            Return
        End If

        Try
            ' Load equipment type as before
            Dim drv As DataRowView = TryCast(cmbSeries.SelectedItem, DataRowView)
            Dim eqId As Integer? = Nothing
            If drv IsNot Nothing AndAlso Not IsDBNull(drv("FK_equipmentTypeId")) Then
                eqId = Convert.ToInt32(drv("FK_equipmentTypeId"))
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

            ' Setup grid columns based on equipment type
            SetupModelGrid(eqId)

            ' Load all existing models for the selected series into the DataGridView
            Using conn = DbConnectionManager.CreateOpenConnection()
                Dim dt As New DataTable()
                Using cmd As New SqlCommand("SELECT modelName, width, depth, height, optionalHeight, optionalDepth, FK_angleTypeId, tahWidth, tahHeight, sahHeight, sahWidth, FK_ampHandleLocationId, tahRearOffset, sahRearOffset, sahTopDownOffset, musicRestDesign, chart_Template, notes, totalFabricSquareInches, parentSku, WooProductId, onReverb, lastUpdated FROM Model WHERE FK_seriesId = @seriesId ORDER BY modelName", conn)
                    cmd.Parameters.AddWithValue("@seriesId", cmbSeries.SelectedValue)
                    dt.Load(cmd.ExecuteReader())
                End Using
                dgvModels.Rows.Clear()
                For Each dr As DataRow In dt.Rows
                    Dim values As New List(Of Object)
                    For Each col As DataGridViewColumn In dgvModels.Columns
                        If dr.Table.Columns.Contains(col.Name) Then
                            Dim val = dr(col.Name)
                            ' Special handling for FK_angleTypeId: ensure integer and present in dtAngleTypes
                            If col.Name = "FK_angleTypeId" Then
                                Dim angleId As Integer = 6
                                If val IsNot Nothing AndAlso Not IsDBNull(val) AndAlso IsNumeric(val) Then
                                    angleId = Convert.ToInt32(val)
                                End If
                                ' Ensure angleId exists in dtAngleTypes, else set to 6
                                Dim found As Boolean = False
                                If dtAngleTypes IsNot Nothing Then
                                    found = dtAngleTypes.Rows.Cast(Of DataRow)().Any(Function(r) Convert.ToInt32(r("PK_AngleTypeId")) = angleId)
                                End If
                                If Not found Then angleId = 6
                                values.Add(angleId)
                            Else
                                values.Add(val)
                            End If
                        Else
                            values.Add(DBNull.Value)
                        End If
                    Next
                    dgvModels.Rows.Add(values.ToArray())
                Next
            End Using

        Catch ex As Exception
            MessageBox.Show("Error loading equipment type or models: " & ex.Message)
        End Try
        ' <<< end changed
    End Sub
    '##############################################################
    ' SetupModelGrid
    ' Purpose: Configure grid columns. Ensures ComboBox columns are
    '          type-safe and do not throw on new/blank rows.
    ' Dependencies: System.Windows.Forms
    ' Date: 2025-09-30
    '##############################################################
    Private Sub SetupModelGrid(Optional equipmentTypeId As Integer? = Nothing)
        Const MUSIC_KEYBOARD_EQUIPMENT_TYPE_ID As Integer = 2 ' adjust to your actual PK

        dgvModels.Columns.Clear()
        dgvModels.AllowUserToAddRows = True
        dgvModels.AllowUserToDeleteRows = True
        dgvModels.AutoGenerateColumns = False
        dgvModels.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvModels.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "modelName", .HeaderText = "Model Name", .DataPropertyName = "modelName"})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "width", .HeaderText = "Width", .DataPropertyName = "width", .Width = 50})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "depth", .HeaderText = "Depth", .DataPropertyName = "depth", .Width = 50})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "height", .HeaderText = "Height", .DataPropertyName = "height", .Width = 50})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "optionalHeight", .HeaderText = "Opt. Height", .DataPropertyName = "optionalHeight", .Width = 50})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "optionalDepth", .HeaderText = "Opt. Depth", .DataPropertyName = "optionalDepth", .Width = 50})

        ' Context menu + right-click row focus
        If dgvContextMenu Is Nothing Then
            dgvContextMenu = New ContextMenuStrip()
            dgvContextMenu.Items.Add("Delete Model", Nothing, AddressOf DeleteSelectedModel)
        End If
        dgvModels.ContextMenuStrip = dgvContextMenu
        RemoveHandler dgvModels.MouseDown, AddressOf dgvModels_MouseDown
        AddHandler dgvModels.MouseDown, AddressOf dgvModels_MouseDown

        Dim isMusicKeyboard As Boolean = (equipmentTypeId.HasValue AndAlso equipmentTypeId.Value = MUSIC_KEYBOARD_EQUIPMENT_TYPE_ID)
        If Not isMusicKeyboard AndAlso equipmentTypeId.HasValue Then
            ' Ensure "None" (6) exists in dtAngleTypes
            If dtAngleTypes IsNot Nothing Then
                Dim foundNone As Boolean = dtAngleTypes.Rows.Cast(Of DataRow)().
                Any(Function(r) Convert.ToInt32(r("PK_AngleTypeId")) = 6)
                If Not foundNone Then
                    Dim noneRow As DataRow = dtAngleTypes.NewRow()
                    noneRow("PK_AngleTypeId") = 6
                    noneRow("AngleTypeName") = "None"
                    dtAngleTypes.Rows.InsertAt(noneRow, 0)
                End If
            End If

            Dim angleTypeCol As New DataGridViewComboBoxColumn() With {
            .Name = "FK_angleTypeId",
            .HeaderText = "Angle Type",
            .DataPropertyName = "FK_angleTypeId",
            .DataSource = dtAngleTypes,
            .DisplayMember = "AngleTypeName",
            .ValueMember = "PK_AngleTypeId",
            .ValueType = GetType(Integer)
        }
            ' IMPORTANT: do NOT force 6 via NullValue; let the cell actually hold 6
            angleTypeCol.DefaultCellStyle = New DataGridViewCellStyle() With {
            .NullValue = Nothing
        }
            angleTypeCol.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            angleTypeCol.FlatStyle = FlatStyle.Flat
            angleTypeCol.DisplayStyleForCurrentCellOnly = False

            dgvModels.Columns.Add(angleTypeCol)

            dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "tahWidth", .HeaderText = "TAH Width", .DataPropertyName = "tahWidth", .Width = 50})
            dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "tahHeight", .HeaderText = "TAH Height", .DataPropertyName = "tahHeight", .Width = 50})
            dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "sahHeight", .HeaderText = "SAH Height", .DataPropertyName = "sahHeight", .Width = 50})
            dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "sahWidth", .HeaderText = "SAH Width", .DataPropertyName = "sahWidth", .Width = 50})

            dgvModels.Columns.Add(New DataGridViewComboBoxColumn() With {
            .Name = "FK_ampHandleLocationId",
            .HeaderText = "AH Location",
            .DataPropertyName = "FK_ampHandleLocationId",
            .DataSource = dtAmpHandleLocations,
            .DisplayMember = "AmpHandleLocationName",
            .ValueMember = "PK_AmpHandleLocationId"
        })

            dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "tahRearOffset", .HeaderText = "TAH Rear Offset", .DataPropertyName = "tahRearOffset", .Width = 50})
            dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "sahRearOffset", .HeaderText = "SAH Rear Offset", .DataPropertyName = "sahRearOffset", .Width = 50})
            dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "sahTopDownOffset", .HeaderText = "SAH TopDown Offset", .DataPropertyName = "sahTopDownOffset", .Width = 50})
        End If

        Dim yesNoSource As New List(Of KeyValuePair(Of String, Boolean)) From {
        New KeyValuePair(Of String, Boolean)("Yes", True),
        New KeyValuePair(Of String, Boolean)("No", False)
    }
        dgvModels.Columns.Add(New DataGridViewComboBoxColumn() With {
        .Name = "musicRestDesign",
        .HeaderText = "MRD?",
        .DataPropertyName = "musicRestDesign",
        .Width = 50,
        .DataSource = New BindingSource(yesNoSource, Nothing),
        .DisplayMember = "Key",
        .ValueMember = "Value"
    })
        dgvModels.Columns.Add(New DataGridViewComboBoxColumn() With {
        .Name = "chart_Template",
        .HeaderText = "Chart / Template",
        .DataPropertyName = "chart_Template",
        .Width = 50,
        .DataSource = New BindingSource(yesNoSource, Nothing),
        .DisplayMember = "Key",
        .ValueMember = "Value"
    })
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "notes", .HeaderText = "Notes", .DataPropertyName = "notes", .Width = 150})

        Dim totalFabricCol As New DataGridViewTextBoxColumn() With {.Name = "totalFabricSquareInches", .HeaderText = "Total Fabric", .DataPropertyName = "totalFabricSquareInches", .Width = 50}
        totalFabricCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvModels.Columns.Add(totalFabricCol)

        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "parentSku", .HeaderText = "Parent SKU", .DataPropertyName = "parentSku", .Width = 225})
        dgvModels.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "WooProductId", .HeaderText = "Woo Id", .DataPropertyName = "WooProductId", .Width = 75})
        dgvModels.Columns.Add(New DataGridViewComboBoxColumn() With {
        .Name = "onReverb",
        .HeaderText = "Reverb?",
        .DataPropertyName = "onReverb",
        .Width = 50,
        .DataSource = New BindingSource(yesNoSource, Nothing),
        .DisplayMember = "Key",
        .ValueMember = "Value"
    })

        Dim lastUpdatedCol As New DataGridViewTextBoxColumn() With {.Name = "lastUpdated", .HeaderText = "Last Update", .DataPropertyName = "lastUpdated", .Width = 75}
        lastUpdatedCol.DefaultCellStyle.Format = "MM/dd/yyyy"
        dgvModels.Columns.Add(lastUpdatedCol)

        If Not dgvModels.Columns.Contains("IsDirty") Then
            Dim isDirtyCol As New DataGridViewCheckBoxColumn() With {.Name = "IsDirty", .HeaderText = "IsDirty", .Visible = False}
            dgvModels.Columns.Add(isDirtyCol)
        End If
    End Sub


    '##############################################################
    ' dgvModels_DefaultValuesNeeded
    ' Purpose: Seed safe defaults for new rows without throwing when
    '          a column is not present (e.g., keyboard equipment type).
    ' Date: 2025-09-30
    '##############################################################
    Private Sub dgvModels_DefaultValuesNeeded(sender As Object, e As DataGridViewRowEventArgs) Handles dgvModels.DefaultValuesNeeded
        If dgvModels.Columns.Contains("FK_angleTypeId") Then
            e.Row.Cells("FK_angleTypeId").Value = CType(6, Integer) ' "None"
        End If
    End Sub
    '##############################################################
    ' dgvModels_RowsAdded
    ' Purpose: Ensure new placeholder row gets a valid combobox value.
    ' Date: 2025-09-30
    '##############################################################
    Private Sub dgvModels_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles dgvModels.RowsAdded
        For i As Integer = 0 To e.RowCount - 1
            Dim r = dgvModels.Rows(e.RowIndex + i)
            If Not r.IsNewRow AndAlso dgvModels.Columns.Contains("FK_angleTypeId") Then
                If r.Cells("FK_angleTypeId").Value Is Nothing OrElse IsDBNull(r.Cells("FK_angleTypeId").Value) Then
                    r.Cells("FK_angleTypeId").Value = CType(6, Integer)
                End If
            End If
        Next
    End Sub

    '##############################################################
    ' dgvModels_DataError
    ' Purpose: Swallow harmless paint-time ComboBox formatting errors.
    ' Date: 2025-09-30
    '##############################################################
    Private Sub dgvModels_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dgvModels.DataError
        Dim colName As String = dgvModels.Columns(e.ColumnIndex).Name
        Dim rowIdx As Integer = e.RowIndex
        System.Diagnostics.Debug.WriteLine($"[DGV DataError] col={colName}, row={rowIdx + 1}: {e.Exception?.Message}")
        e.ThrowException = False
        e.Cancel = True
    End Sub

    ' Purpose: Save all new and updated models from the DataGridView, rounding totalFabricSquareInches up to the next 1/4" before saving.
    ' Only new or edited rows are processed (dirty tracking).
    ' Dependencies: Imports System.Data.SqlClient, ModelSkuBuilder, System.Windows.Forms, System.Math
    ' Current date: 2025-09-30
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' >>> changed
        ' Ensure IsDirty column exists (hidden)
        If Not dgvModels.Columns.Contains("IsDirty") Then
            Dim isDirtyCol As New DataGridViewCheckBoxColumn() With {
            .Name = "IsDirty",
            .HeaderText = "IsDirty",
            .Visible = False
        }
            dgvModels.Columns.Add(isDirtyCol)
        End If

        ' Mark row dirty on edit
        RemoveHandler dgvModels.CellValueChanged, AddressOf dgvModels_CellValueChanged
        AddHandler dgvModels.CellValueChanged, AddressOf dgvModels_CellValueChanged

        If cmbManufacturer.SelectedIndex = -1 OrElse cmbSeries.SelectedIndex = -1 Then
            MessageBox.Show("Please select a Manufacturer and Series.")
            Return
        End If

        If dgvModels.Rows.Count = 0 Then
            MessageBox.Show("Please add at least one model.")
            Return
        End If

        Dim insertedCount As Integer = 0
        Dim updatedCount As Integer = 0

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

                    ' Only process if new (no parentSku) or dirty
                    Dim isDirty As Boolean = False
                    If row.Cells("IsDirty").Value IsNot Nothing AndAlso Convert.ToBoolean(row.Cells("IsDirty").Value) Then
                        isDirty = True
                    End If
                    Dim isNew As Boolean = String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("parentSku").Value))
                    If Not isDirty AndAlso Not isNew Then Continue For ' <<< Only process new or dirty rows

                    Dim manufacturerName As String = cmbManufacturer.Text
                    Dim seriesName As String = cmbSeries.Text
                    Dim modelName As String = Convert.ToString(row.Cells("modelName").Value)
                    Dim versionSuffix As String = "V1"
                    Dim uniqueId As Integer = 0
                    Dim parentSku As String = ModelSkuBuilder.GenerateParentSku(manufacturerName, seriesName, modelName, versionSuffix, uniqueId)

                    ' Get raw grid values for width, depth, height
                    Dim widthRaw As Decimal = Convert.ToDecimal(row.Cells("width").Value)
                    Dim depthRaw As Decimal = Convert.ToDecimal(row.Cells("depth").Value)
                    Dim heightRaw As Decimal = Convert.ToDecimal(row.Cells("height").Value)

                    ' Check if this is an existing model (has a parentSku in the DB)
                    Dim existingParentSku As String = TryCast(row.Cells("parentSku").Value, String)
                    Dim isExistingModel As Boolean = False
                    Dim dbValues As New Dictionary(Of String, Object)

                    If Not String.IsNullOrWhiteSpace(existingParentSku) Then
                        ' Check if parentSku exists in DB and get all fields
                        Using checkCmd As New SqlCommand("SELECT modelName, width, depth, height, optionalHeight, optionalDepth, FK_angleTypeId, tahWidth, tahHeight, sahHeight, sahWidth, FK_ampHandleLocationId, tahRearOffset, sahRearOffset, sahTopDownOffset, musicRestDesign, chart_Template, notes, totalFabricSquareInches, WooProductId, onReverb FROM Model WHERE parentSku = @parentSku", conn)
                            checkCmd.Parameters.AddWithValue("@parentSku", existingParentSku)
                            Using rdr = checkCmd.ExecuteReader()
                                If rdr.Read() Then
                                    isExistingModel = True
                                    For i As Integer = 0 To rdr.FieldCount - 1
                                        dbValues(rdr.GetName(i)) = rdr.GetValue(i)
                                    Next
                                End If
                            End Using
                        End Using
                    End If

                    ' Only recalculate if user changed width/depth/height
                    Dim widthRecalc As Decimal = widthRaw
                    Dim depthRecalc As Decimal = depthRaw
                    Dim heightRecalc As Decimal = heightRaw
                    Dim recalcNeeded As Boolean = False

                    If isExistingModel Then
                        If Not IsDBNull(dbValues("width")) AndAlso Convert.ToDecimal(dbValues("width")) <> widthRaw Then recalcNeeded = True
                        If Not IsDBNull(dbValues("depth")) AndAlso Convert.ToDecimal(dbValues("depth")) <> depthRaw Then recalcNeeded = True
                        If Not IsDBNull(dbValues("height")) AndAlso Convert.ToDecimal(dbValues("height")) <> heightRaw Then recalcNeeded = True
                    End If

                    If recalcNeeded OrElse Not isExistingModel Then
                        widthRecalc = Math.Ceiling(widthRaw * 8D) / 8D
                        depthRecalc = Math.Ceiling(depthRaw * 8D) / 8D
                        heightRecalc = Math.Ceiling(heightRaw * 8D) / 8D
                    End If

                    Dim adjWidth As Decimal = widthRecalc + 1.5D
                    Dim adjDepth As Decimal = depthRecalc + 1.5D
                    Dim adjHeight As Decimal = heightRecalc + 1.5D

                    Dim sidePanels As Decimal = (adjHeight * adjDepth) * 2D
                    Dim topBottomPanels As Decimal = (adjWidth * adjHeight) * 2D
                    Dim total As Decimal = sidePanels + topBottomPanels

                    ' Calculate totalWithWaste and round up to the next 1/4"
                    Dim totalWithWasteUnrounded As Decimal = total * 1.05D
                    Dim totalWithWaste As Decimal = Math.Ceiling(totalWithWasteUnrounded * 4D) / 4D

                    If isExistingModel Then
                        ' Compare all fields and build a change summary
                        Dim changes As New List(Of String)
                        Dim fieldNames As String() = {
                        "modelName", "width", "depth", "height", "optionalHeight", "optionalDepth", "FK_angleTypeId",
                        "tahWidth", "tahHeight", "sahHeight", "sahWidth", "FK_ampHandleLocationId", "tahRearOffset",
                        "sahRearOffset", "sahTopDownOffset", "musicRestDesign", "chart_Template", "notes",
                        "totalFabricSquareInches", "WooProductId", "onReverb"
                    }
                        For Each field In fieldNames
                            Dim oldVal = If(dbValues.ContainsKey(field), dbValues(field), Nothing)
                            Dim newVal As Object = Nothing
                            Select Case field
                                Case "modelName"
                                    newVal = modelName
                                Case "width"
                                    newVal = widthRaw
                                Case "depth"
                                    newVal = depthRaw
                                Case "height"
                                    newVal = heightRaw
                                Case "totalFabricSquareInches"
                                    newVal = totalWithWaste
                                Case Else
                                    newVal = row.Cells(field).Value
                            End Select
                            If IsDBNull(oldVal) Then oldVal = Nothing
                            If IsDBNull(newVal) Then newVal = Nothing
                            If (If(oldVal, "")).ToString() <> (If(newVal, "")).ToString() Then
                                changes.Add($"{field}: '{oldVal}' → '{newVal}'")
                            End If
                        Next

                        If changes.Count > 0 Then
                            Dim msg = "The following changes were made to this model:" & vbCrLf & String.Join(vbCrLf, changes) & vbCrLf & vbCrLf & "Do you want to save these changes?"
                            Dim result = MessageBox.Show(msg, "Confirm Model Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                            If result <> DialogResult.Yes Then
                                Continue For ' Skip update if not confirmed
                            End If
                        Else
                            Continue For ' No changes, skip update
                        End If

                        ' Update existing model
                        Dim updateSql As String =
                        "UPDATE Model SET modelName=@modelName, width=@width, depth=@depth, height=@height, optionalHeight=@optionalHeight, optionalDepth=@optionalDepth, FK_angleTypeId=@FK_angleTypeId, tahWidth=@tahWidth, tahHeight=@tahHeight, sahHeight=@sahHeight, sahWidth=@sahWidth, FK_ampHandleLocationId=@FK_ampHandleLocationId, tahRearOffset=@tahRearOffset, sahRearOffset=@sahRearOffset, sahTopDownOffset=@sahTopDownOffset, musicRestDesign=@musicRestDesign, chart_Template=@chart_Template, notes=@notes, totalFabricSquareInches=@totalFabricSquareInches, WooProductId=@WooProductId, onReverb=@onReverb, lastUpdated=@lastUpdated WHERE parentSku=@parentSku"
                        Using cmd As New SqlCommand(updateSql, conn)
                            cmd.Parameters.AddWithValue("@modelName", modelName)
                            cmd.Parameters.AddWithValue("@width", widthRecalc)
                            cmd.Parameters.AddWithValue("@depth", depthRecalc)
                            cmd.Parameters.AddWithValue("@height", heightRecalc)
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
                            cmd.Parameters.AddWithValue("@totalFabricSquareInches", totalWithWaste)
                            cmd.Parameters.AddWithValue("@WooProductId", If(IsDBNull(row.Cells("WooProductId").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("WooProductId").Value)), DBNull.Value, Convert.ToInt32(row.Cells("WooProductId").Value)))
                            cmd.Parameters.AddWithValue("@onReverb", If(IsDBNull(row.Cells("onReverb").Value), False, Convert.ToBoolean(row.Cells("onReverb").Value)))
                            cmd.Parameters.AddWithValue("@lastUpdated", DateTime.Now)
                            cmd.Parameters.AddWithValue("@parentSku", existingParentSku)
                            updatedCount += cmd.ExecuteNonQuery()
                        End Using

                        row.Cells("parentSku").Value = existingParentSku
                        row.Cells("totalFabricSquareInches").Value = totalWithWaste
                        row.Cells("lastUpdated").Value = DateTime.Now
                        row.Cells("IsDirty").Value = False ' <<< clear dirty flag after save
                    Else
                        ' Insert new model
                        row.Cells("parentSku").Value = parentSku
                        row.Cells("totalFabricSquareInches").Value = totalWithWaste
                        row.Cells("lastUpdated").Value = DateTime.Now

                        Dim sql As String = "INSERT INTO Model (FK_seriesId, modelName, width, depth, height, optionalHeight, optionalDepth, FK_angleTypeId, tahWidth, tahHeight, sahHeight, sahWidth, FK_ampHandleLocationId, tahRearOffset, sahRearOffset, sahTopDownOffset, musicRestDesign, chart_Template, notes, totalFabricSquareInches, parentSku, WooProductId, onReverb, lastUpdated) " &
                        "VALUES (@seriesId, @modelName, @width, @depth, @height, @optionalHeight, @optionalDepth, @FK_angleTypeId, @tahWidth, @tahHeight, @sahHeight, @sahWidth, @FK_ampHandleLocationId, @tahRearOffset, @sahRearOffset, @sahTopDownOffset, @musicRestDesign, @chart_Template, @notes, @totalFabricSquareInches, @parentSku, @WooProductId, @onReverb, @lastUpdated)"
                        Using cmd As New SqlCommand(sql, conn)
                            cmd.Parameters.AddWithValue("@seriesId", cmbSeries.SelectedValue)
                            cmd.Parameters.AddWithValue("@modelName", modelName)
                            cmd.Parameters.AddWithValue("@width", widthRecalc)
                            cmd.Parameters.AddWithValue("@depth", depthRecalc)
                            cmd.Parameters.AddWithValue("@height", heightRecalc)
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
                            cmd.Parameters.AddWithValue("@totalFabricSquareInches", totalWithWaste)
                            cmd.Parameters.AddWithValue("@parentSku", parentSku)
                            cmd.Parameters.AddWithValue("@WooProductId", If(IsDBNull(row.Cells("WooProductId").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("WooProductId").Value)), DBNull.Value, Convert.ToInt32(row.Cells("WooProductId").Value)))
                            cmd.Parameters.AddWithValue("@onReverb", If(IsDBNull(row.Cells("onReverb").Value), False, Convert.ToBoolean(row.Cells("onReverb").Value)))
                            cmd.Parameters.AddWithValue("@lastUpdated", DateTime.Now)
                            insertedCount += cmd.ExecuteNonQuery()
                        End Using
                        row.Cells("IsDirty").Value = False ' <<< clear dirty flag after save
                    End If
                Next
            End Using
            MessageBox.Show($"{insertedCount} model(s) added, {updatedCount} model(s) updated successfully!")
            RefreshModelGridFromDatabase()
        Catch ex As Exception
            MessageBox.Show("Error saving models: " & ex.Message)
        End Try
        ' <<< end changed
    End Sub

    ' Handles marking rows as dirty when edited
    Private Sub dgvModels_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        ' >>> changed
        If e.RowIndex >= 0 AndAlso Not dgvModels.Rows(e.RowIndex).IsNewRow Then
            dgvModels.Rows(e.RowIndex).Cells("IsDirty").Value = True
        End If
        ' <<< end changed
    End Sub

    ' Purpose: Reloads the DataGridView with all models for the currently selected Series, including lastUpdated.
    ' Ensures FK_angleTypeId is always set to a valid value for the ComboBox column.
    ' Dependencies: Imports System.Data.SqlClient, DbConnectionManager
    ' Current date: 2025-09-30
    Private Sub RefreshModelGridFromDatabase()
        ' >>> changed
        If cmbSeries.SelectedIndex = -1 OrElse cmbSeries.SelectedValue Is Nothing Then Exit Sub
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                Dim dt As New DataTable()
                Using cmd As New SqlCommand("SELECT modelName, width, depth, height, optionalHeight, optionalDepth, FK_angleTypeId, tahWidth, tahHeight, sahHeight, sahWidth, FK_ampHandleLocationId, tahRearOffset, sahRearOffset, sahTopDownOffset, musicRestDesign, chart_Template, notes, totalFabricSquareInches, parentSku, WooProductId, onReverb, lastUpdated FROM Model WHERE FK_seriesId = @seriesId ORDER BY modelName", conn)
                    cmd.Parameters.AddWithValue("@seriesId", cmbSeries.SelectedValue)
                    dt.Load(cmd.ExecuteReader())
                End Using
                dgvModels.Rows.Clear()
                For Each dr As DataRow In dt.Rows
                    Dim values As New List(Of Object)
                    For Each col As DataGridViewColumn In dgvModels.Columns
                        If dr.Table.Columns.Contains(col.Name) Then
                            Dim val = dr(col.Name)
                            ' Special handling for FK_angleTypeId: ensure integer and present in dtAngleTypes
                            If col.Name = "FK_angleTypeId" Then
                                Dim angleId As Integer = 6
                                If val IsNot Nothing AndAlso Not IsDBNull(val) AndAlso IsNumeric(val) Then
                                    angleId = Convert.ToInt32(val)
                                End If
                                ' Ensure angleId exists in dtAngleTypes, else set to 6
                                Dim found As Boolean = False
                                If dtAngleTypes IsNot Nothing Then
                                    found = dtAngleTypes.Rows.Cast(Of DataRow)().Any(Function(r) Convert.ToInt32(r("PK_AngleTypeId")) = angleId)
                                End If
                                If Not found Then angleId = 6
                                values.Add(angleId)
                            Else
                                values.Add(val)
                            End If
                        Else
                            values.Add(DBNull.Value)
                        End If
                    Next
                    dgvModels.Rows.Add(values.ToArray())
                Next
            End Using
        Catch ex As Exception
            MessageBox.Show("Error refreshing model grid: " & ex.Message)
        End Try
        ' <<< end changed
    End Sub
    'Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
    '    If cmbManufacturer.SelectedIndex = -1 OrElse cmbSeries.SelectedIndex = -1 Then
    '        MessageBox.Show("Please select a Manufacturer and Series.")
    '        Return
    '    End If

    '    If dgvModels.Rows.Count = 0 Then
    '        MessageBox.Show("Please add at least one model.")
    '        Return
    '    End If

    '    Dim insertedCount As Integer = 0
    '    Try
    '        Using conn = DbConnectionManager.CreateOpenConnection()
    '            For Each row As DataGridViewRow In dgvModels.Rows
    '                If row.IsNewRow Then Continue For
    '                If String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("modelName").Value)) OrElse
    '                   String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("width").Value)) OrElse
    '                   String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("depth").Value)) OrElse
    '                   String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("height").Value)) Then
    '                    Continue For
    '                End If

    '                Dim sql As String = "INSERT INTO Model (FK_seriesId, modelName, width, depth, height, optionalHeight, optionalDepth, FK_angleTypeId, tahWidth, tahHeight, sahHeight, sahWidth, FK_ampHandleLocationId, tahRearOffset, sahRearOffset, sahTopDownOffset, musicRestDesign, chart_Template, notes, totalFabricSquareInches, parentSku, WooProductId, onReverb) " &
    '                                    "VALUES (@seriesId, @modelName, @width, @depth, @height, @optionalHeight, @optionalDepth, @FK_angleTypeId, @tahWidth, @tahHeight, @sahHeight, @sahWidth, @FK_ampHandleLocationId, @tahRearOffset, @sahRearOffset, @sahTopDownOffset, @musicRestDesign, @chart_Template, @notes, @totalFabricSquareInches, @parentSku, @WooProductId, @onReverb)"
    '                Using cmd As New SqlCommand(sql, conn)
    '                    cmd.Parameters.AddWithValue("@seriesId", cmbSeries.SelectedValue)
    '                    cmd.Parameters.AddWithValue("@modelName", Convert.ToString(row.Cells("modelName").Value))
    '                    cmd.Parameters.AddWithValue("@width", Convert.ToDecimal(row.Cells("width").Value))
    '                    cmd.Parameters.AddWithValue("@depth", Convert.ToDecimal(row.Cells("depth").Value))
    '                    cmd.Parameters.AddWithValue("@height", Convert.ToDecimal(row.Cells("height").Value))
    '                    cmd.Parameters.AddWithValue("@optionalHeight", If(IsDBNull(row.Cells("optionalHeight").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("optionalHeight").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("optionalHeight").Value)))
    '                    cmd.Parameters.AddWithValue("@optionalDepth", If(IsDBNull(row.Cells("optionalDepth").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("optionalDepth").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("optionalDepth").Value)))
    '                    cmd.Parameters.AddWithValue("@FK_angleTypeId", If(IsDBNull(row.Cells("FK_angleTypeId").Value), DBNull.Value, row.Cells("FK_angleTypeId").Value))
    '                    cmd.Parameters.AddWithValue("@tahWidth", If(IsDBNull(row.Cells("tahWidth").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("tahWidth").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("tahWidth").Value)))
    '                    cmd.Parameters.AddWithValue("@tahHeight", If(IsDBNull(row.Cells("tahHeight").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("tahHeight").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("tahHeight").Value)))
    '                    cmd.Parameters.AddWithValue("@sahHeight", If(IsDBNull(row.Cells("sahHeight").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("sahHeight").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("sahHeight").Value)))
    '                    cmd.Parameters.AddWithValue("@sahWidth", If(IsDBNull(row.Cells("sahWidth").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("sahWidth").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("sahWidth").Value)))
    '                    cmd.Parameters.AddWithValue("@FK_ampHandleLocationId", If(IsDBNull(row.Cells("FK_ampHandleLocationId").Value), DBNull.Value, row.Cells("FK_ampHandleLocationId").Value))
    '                    cmd.Parameters.AddWithValue("@tahRearOffset", If(IsDBNull(row.Cells("tahRearOffset").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("tahRearOffset").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("tahRearOffset").Value)))
    '                    cmd.Parameters.AddWithValue("@sahRearOffset", If(IsDBNull(row.Cells("sahRearOffset").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("sahRearOffset").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("sahRearOffset").Value)))
    '                    cmd.Parameters.AddWithValue("@sahTopDownOffset", If(IsDBNull(row.Cells("sahTopDownOffset").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("sahTopDownOffset").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("sahTopDownOffset").Value)))
    '                    cmd.Parameters.AddWithValue("@musicRestDesign", If(IsDBNull(row.Cells("musicRestDesign").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("musicRestDesign").Value)), DBNull.Value, row.Cells("musicRestDesign").Value))
    '                    cmd.Parameters.AddWithValue("@chart_Template", If(IsDBNull(row.Cells("chart_Template").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("chart_Template").Value)), DBNull.Value, row.Cells("chart_Template").Value))
    '                    cmd.Parameters.AddWithValue("@notes", If(IsDBNull(row.Cells("notes").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("notes").Value)), DBNull.Value, row.Cells("notes").Value))
    '                    cmd.Parameters.AddWithValue("@totalFabricSquareInches", If(IsDBNull(row.Cells("totalFabricSquareInches").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("totalFabricSquareInches").Value)), DBNull.Value, Convert.ToDecimal(row.Cells("totalFabricSquareInches").Value)))
    '                    cmd.Parameters.AddWithValue("@parentSku", If(IsDBNull(row.Cells("parentSku").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("parentSku").Value)), DBNull.Value, row.Cells("parentSku").Value))
    '                    cmd.Parameters.AddWithValue("@WooProductId", If(IsDBNull(row.Cells("WooProductId").Value) OrElse String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("WooProductId").Value)), DBNull.Value, Convert.ToInt32(row.Cells("WooProductId").Value)))
    '                    cmd.Parameters.AddWithValue("@onReverb", If(IsDBNull(row.Cells("onReverb").Value), False, Convert.ToBoolean(row.Cells("onReverb").Value)))
    '                    insertedCount += cmd.ExecuteNonQuery()
    '                End Using
    '            Next
    '        End Using
    '        MessageBox.Show($"{insertedCount} model(s) added successfully!")
    '        Me.Close()
    '    Catch ex As Exception
    '        MessageBox.Show("Error saving models: " & ex.Message)
    '    End Try
    'End Sub

    ' Handles return to dashboard
    Private Sub btnReturnToDashboard_Click(sender As Object, e As EventArgs) Handles btnReturnToDashboard.Click
        Dim dashboard As New formDashboard()
        formDashboard.Show()
        Me.Close()
    End Sub
    ' Purpose: Opens FormAddManufacturerSeries and refreshes cmbManufacturer after closing, always clearing selection first.
    ' Dependencies: Imports System.Windows.Forms
    ' Current date: 2025-09-30
    Private Sub btnAddManufacturerSeriesInfo_Click(sender As Object, e As EventArgs) Handles btnAddManufacturerSeriesInfo.Click
        ' Always clear selection before opening the add form; do not track or reuse selected manufacturer
        cmbManufacturer.SelectedIndex = -1
        Using frm As New FormAddManufacturerSeries()
            frm.ShowDialog(Me)
        End Using
        RefreshManufacturers()

    End Sub

    ' Purpose: Reloads manufacturers into cmbManufacturer, safely handling selection after add.
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
                ' Only select if the value exists in the refreshed list
                If selectManufacturerId.HasValue Then
                    Dim found As Boolean = False
                    For Each row As DataRowView In dt.DefaultView
                        If Convert.ToInt32(row("PK_manufacturerId")) = selectManufacturerId.Value Then
                            found = True
                            Exit For
                        End If
                    Next
                    If found Then
                        cmbManufacturer.SelectedValue = selectManufacturerId.Value
                    Else
                        cmbManufacturer.SelectedIndex = -1
                    End If
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