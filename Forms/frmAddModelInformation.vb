Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports System.Linq
Imports System.Text.Json
Imports System.Windows.Forms

Public Class frmAddModelInformation


    '*******************************************************************************************************************************************************************
    '*******************************************************************************************************************************************************************
#Region "*************START***********FORM INITIALIZATION AND LOAD*****************"

    ' Form-level fields for state and data
    Private isFormLoaded As Boolean = False
    Private isUserSelectingSeries As Boolean = False
    Private isLoadingSeries As Boolean = False
    Private isLoadingModels As Boolean = False
    Private currentModelsTable As DataTable
    Private originalRowValues As Dictionary(Of Integer, Dictionary(Of String, Object)) = New Dictionary(Of Integer, Dictionary(Of String, Object))()

    ' Form Load event: sets up ComboBoxes, DataGridView, and handlers
    Private Sub AddModelInformation_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim db As New DbConnectionManager()

        ' Manufacturer ComboBox: table-driven
        cmbManufacturerName.DropDownStyle = ComboBoxStyle.DropDownList
        cmbManufacturerName.MaxDropDownItems = 10
        cmbManufacturerName.DataSource = Nothing
        cmbManufacturerName.DataSource = db.GetManufacturerNames()
        cmbManufacturerName.DisplayMember = "ManufacturerName"
        cmbManufacturerName.ValueMember = "PK_ManufacturerId"

        ' Equipment Type ComboBox: table-driven ONLY
        cmbEquipmentType.DataSource = Nothing
        cmbEquipmentType.DataSource = db.GetEquipmentTypes()
        cmbEquipmentType.DisplayMember = "EquipmentTypeName"
        cmbEquipmentType.ValueMember = "PK_EquipmentTypeId"
        cmbEquipmentType.DropDownStyle = ComboBoxStyle.DropDownList

        AddHandler dgvSeriesName.MouseDown, AddressOf dgvSeriesName_MouseDown
        AddHandler dgvSeriesName.KeyDown, AddressOf dgvSeriesName_KeyDown
        dgvModelInformation.AutoGenerateColumns = False

        isFormLoaded = True
        ' InitializeModelGridColumns() ' Uncomment if you want to initialize columns on load
    End Sub

#End Region


    '*******************************************************************************************************************************************************************
    '*******************************************************************************************************************************************************************
#Region "*************START***********COMBOBOX EVENT HANDLERS**********************************************************************************************************"

    ' Handles Enter key in Manufacturer ComboBox to add a new manufacturer
    Private Sub cmbManufacturerName_KeyDown(sender As Object, e As KeyEventArgs) Handles cmbManufacturerName.KeyDown
        If e.KeyCode = Keys.Enter Then
            CheckAndAddManufacturer()
            e.Handled = True
            e.SuppressKeyPress = True
        End If
    End Sub

    ' Handles leaving Manufacturer ComboBox to check/add manufacturer
    Private Sub cmbManufacturerName_Leave(sender As Object, e As EventArgs) Handles cmbManufacturerName.Leave
        CheckAndAddManufacturer()
    End Sub


    ' Handles Manufacturer selection change to update Series and Models
    Private Sub cmbManufacturerName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbManufacturerName.SelectedIndexChanged
        If Not isFormLoaded Then Return
        If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then Return

        Dim selectedManufacturerId As Integer
        If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then Return

        Dim db As New DbConnectionManager()
        Dim seriesList = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
        isLoadingSeries = True
        cmbSeriesName.DataSource = Nothing
        cmbSeriesName.DataSource = seriesList
        cmbSeriesName.DisplayMember = "SeriesName"
        cmbSeriesName.ValueMember = "PK_SeriesId"
        cmbSeriesName.SelectedIndex = -1
        isLoadingSeries = False

        LoadModelsForSelectedSeries()
        LoadSeriesGridForManufacturer()

    End Sub

    ' Handles Series selection change to update Equipment Type, columns, and models
    Private Sub cmbSeriesName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSeriesName.SelectedIndexChanged
        If isLoadingSeries OrElse Not isFormLoaded Then Return

        SetEquipmentTypeForSelectedSeries()

        ' Get the selected equipment type name
        Dim equipmentTypeName As String = cmbEquipmentType.Text.Trim()

        ' Initialize columns for the selected equipment type
        InitializeModelGridColumns(equipmentTypeName)

        LoadModelsForSelectedSeries()
        ShowListingPreview()

        If Not String.IsNullOrEmpty(cmbSeriesName.Text.Trim()) Then
            dgvModelInformation.Enabled = True
        Else
            dgvModelInformation.Enabled = False
        End If
    End Sub

    ' Handles leaving Series ComboBox to check/add new series
    Private Sub cmbSeriesName_Leave(sender As Object, e As EventArgs) Handles cmbSeriesName.Leave
        If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then Exit Sub
        Dim selectedManufacturerId As Integer
        If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then Exit Sub

        Dim enteredSeries As String = cmbSeriesName.Text.Trim()
        If String.IsNullOrEmpty(enteredSeries) Then Exit Sub

        Dim db As New DbConnectionManager()
        Dim existingSeries = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
        If Not existingSeries.AsEnumerable().Any(Function(row) String.Equals(row.Field(Of String)("SeriesName"), enteredSeries, StringComparison.OrdinalIgnoreCase)) Then
            Dim equipmentType As String = cmbEquipmentType.Text.Trim()
            If String.IsNullOrEmpty(equipmentType) Then
                MessageBox.Show("Please select or enter an equipment type for this series.", "Equipment Type Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbEquipmentType.Focus()
                Exit Sub
            End If

            If cmbEquipmentType.SelectedValue Is Nothing OrElse Not IsNumeric(cmbEquipmentType.SelectedValue) Then
                MessageBox.Show("Please select a valid equipment type.", "Invalid Equipment Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbEquipmentType.Focus()
                Return
            End If

            Dim equipmentTypeValue As Integer = Convert.ToInt32(cmbEquipmentType.SelectedValue)
            Dim msg As String = $"Do you want to save this new series?" & vbCrLf &
                        $"Manufacturer ID: {selectedManufacturerId}" & vbCrLf &
                        $"Series: {enteredSeries}" & vbCrLf &
                        $"Equipment Type: {equipmentType}"
            Dim result = MessageBox.Show(msg, "Save New Series", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                db.InsertSeries(enteredSeries, selectedManufacturerId, equipmentTypeValue)
                cmbSeriesName.DataSource = Nothing
                cmbSeriesName.DataSource = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
                cmbSeriesName.DisplayMember = "SeriesName"
                cmbSeriesName.ValueMember = "PK_SeriesId"
                cmbSeriesName.SelectedItem = enteredSeries
            End If
        End If

        SetEquipmentTypeForSelectedSeries()
    End Sub

    ' Handles leaving Equipment Type ComboBox to try saving new series
    Private Sub cmbEquipmentType_Leave(sender As Object, e As EventArgs) Handles cmbEquipmentType.Leave
        TrySaveNewSeries()
    End Sub

#End Region



    '*******************************************************************************************************************************************************************
    '*******************************************************************************************************************************************************************
#Region "*************START***********DATAGRIDVIEW EVENT HANDLERS*****************"

    Private Sub dgvSeriesName_MouseDown(sender As Object, e As MouseEventArgs)
        isUserSelectingSeries = True
    End Sub

    Private Sub dgvSeriesName_KeyDown(sender As Object, e As KeyEventArgs)
        isUserSelectingSeries = True
    End Sub

    Private Sub dgvSeriesName_SelectionChanged(sender As Object, e As EventArgs) Handles dgvSeriesName.SelectionChanged
        If Not isFormLoaded OrElse dgvSeriesName.SelectedRows.Count = 0 Then Return

        If Not isUserSelectingSeries Then Return
        isUserSelectingSeries = False

        Dim selectedRow = dgvSeriesName.SelectedRows(0)
        Dim selectedSeries As String = Convert.ToString(selectedRow.Cells("SeriesName").Value)
        Dim selectedManufacturer As String = cmbManufacturerName.Text.Trim()

        If String.IsNullOrEmpty(selectedManufacturer) OrElse String.IsNullOrEmpty(selectedSeries) Then Return

        Dim db As New DbConnectionManager()
        Dim dt = db.GetModelsByManufacturerAndSeries(selectedManufacturer, selectedSeries)
        If dt.Rows.Count = 0 Then
            MessageBox.Show("There are no models associated with this series yet.", "No Models Found", MessageBoxButtons.OK, MessageBoxIcon.Information)
            dgvModelInformation.DataSource = Nothing
        Else
            dgvModelInformation.DataSource = dt
        End If
    End Sub

    Private Sub dgvSeriesName_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSeriesName.CellContentClick
        ' (Empty or add logic as needed)
    End Sub

    Private Sub dgvModelInformation_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvModelInformation.CellContentClick
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return

        dgvModelInformation.CurrentCell = Nothing
        Application.DoEvents()

        If dgvModelInformation.Columns(e.ColumnIndex).Name = "Save" Then
            Dim row = dgvModelInformation.Rows(e.RowIndex)
            If row.IsNewRow AndAlso row.Cells.Cast(Of DataGridViewCell).All(Function(c) c.Value Is Nothing OrElse String.IsNullOrWhiteSpace(c.FormattedValue?.ToString())) Then
                MessageBox.Show("Please enter data before saving.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Use IDs, not names
            Dim selectedManufacturerId As Integer
            If cmbManufacturerName.SelectedValue Is Nothing OrElse Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
                MessageBox.Show("Please select a valid manufacturer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim selectedSeriesId As Integer
            If cmbSeriesName.SelectedValue Is Nothing OrElse Not Integer.TryParse(cmbSeriesName.SelectedValue.ToString(), selectedSeriesId) Then
                MessageBox.Show("Please select a valid series.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Get manufacturer, series, and model names
            Dim manufacturerName As String = cmbManufacturerName.Text.Trim()
            Dim seriesName As String = cmbSeriesName.Text.Trim()
            Dim modelName As String = If(dgvModelInformation.Columns.Contains("ModelName"), Convert.ToString(row.Cells("ModelName").Value)?.Trim(), Nothing)

            ' Gather all values (unchanged)
            Dim width As Decimal = If(dgvModelInformation.Columns.Contains("Width") AndAlso Not IsDBNull(row.Cells("Width").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Width").Value.ToString()), Convert.ToDecimal(row.Cells("Width").Value), 0D)
            Dim depth As Decimal = If(dgvModelInformation.Columns.Contains("Depth") AndAlso Not IsDBNull(row.Cells("Depth").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Depth").Value.ToString()), Convert.ToDecimal(row.Cells("Depth").Value), 0D)
            Dim height As Decimal = If(dgvModelInformation.Columns.Contains("Height") AndAlso Not IsDBNull(row.Cells("Height").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Height").Value.ToString()), Convert.ToDecimal(row.Cells("Height").Value), 0D)
            Dim optionalHeight As Decimal? = If(dgvModelInformation.Columns.Contains("OptionalHeight") AndAlso Not IsDBNull(row.Cells("OptionalHeight").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("OptionalHeight").Value.ToString()), Convert.ToDecimal(row.Cells("OptionalHeight").Value), Nothing)
            Dim optionalDepth As Decimal? = If(dgvModelInformation.Columns.Contains("OptionalDepth") AndAlso Not IsDBNull(row.Cells("OptionalDepth").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("OptionalDepth").Value.ToString()), Convert.ToDecimal(row.Cells("OptionalDepth").Value), Nothing)
            Dim angleType As String = If(dgvModelInformation.Columns.Contains("AngleType"), Convert.ToString(row.Cells("AngleType").Value), Nothing)
            Dim ampHandleLocation As String = If(dgvModelInformation.Columns.Contains("AmpHandleLocation"), Convert.ToString(row.Cells("AmpHandleLocation").Value), Nothing)
            Dim tahWidth As Decimal? = If(dgvModelInformation.Columns.Contains("TAHWidth") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHWidth").Value)), Convert.ToDecimal(row.Cells("TAHWidth").Value), Nothing)
            Dim tahHeight As Decimal? = If(dgvModelInformation.Columns.Contains("TAHHeight") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHHeight").Value)), Convert.ToDecimal(row.Cells("TAHHeight").Value), Nothing)
            Dim sahWidth As Decimal? = If(dgvModelInformation.Columns.Contains("SAHWidth") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHWidth").Value)), Convert.ToDecimal(row.Cells("SAHWidth").Value), Nothing)
            Dim sahHeight As Decimal? = If(dgvModelInformation.Columns.Contains("SAHHeight") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHHeight").Value)), Convert.ToDecimal(row.Cells("SAHHeight").Value), Nothing)
            Dim tahRearOffset As Decimal? = If(dgvModelInformation.Columns.Contains("TAHRearOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHRearOffset").Value)), Convert.ToDecimal(row.Cells("TAHRearOffset").Value), Nothing)
            Dim sahRearOffset As Decimal? = If(dgvModelInformation.Columns.Contains("SAHRearOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHRearOffset").Value)), Convert.ToDecimal(row.Cells("SAHRearOffset").Value), Nothing)
            Dim sahTopDownOffset As Decimal? = If(dgvModelInformation.Columns.Contains("SAHTopDownOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHTopDownOffset").Value)), Convert.ToDecimal(row.Cells("SAHTopDownOffset").Value), Nothing)
            Dim musicRestDesign As Boolean? = If(dgvModelInformation.Columns.Contains("MusicRestDesign") AndAlso row.Cells("MusicRestDesign").Value IsNot Nothing, Convert.ToBoolean(row.Cells("MusicRestDesign").Value), Nothing)
            Dim chartTemplate As Boolean? = If(dgvModelInformation.Columns.Contains("Chart_Template") AndAlso row.Cells("Chart_Template").Value IsNot Nothing, Convert.ToBoolean(row.Cells("Chart_Template").Value), Nothing)
            Dim notes As String = If(dgvModelInformation.Columns.Contains("Notes"), Convert.ToString(row.Cells("Notes").Value), Nothing)

            Dim db As New DbConnectionManager()
            Dim totalFabricSquareInches As Decimal = ((depth + 1.25D) * (height + 1.25D) * 2D) + ((depth + 1.25D + height + 1.25D + height + 1.25D) * (width + 1.25D))

            If dgvModelInformation.Columns.Contains("TotalFabricSquareInches") Then
                row.Cells("TotalFabricSquareInches").Value = totalFabricSquareInches
            End If

            Dim isInsert As Boolean = False
            If row.DataBoundItem IsNot Nothing AndAlso TypeOf row.DataBoundItem Is DataRowView Then
                Dim drv = DirectCast(row.DataBoundItem, DataRowView)
                If drv.Row.RowState = DataRowState.Added Then
                    isInsert = True
                End If
            End If

            If isInsert Then
                If db.ModelExists(selectedManufacturerId, selectedSeriesId, modelName) Then
                    MessageBox.Show("A model with this name already exists for the selected manufacturer and series.", "Duplicate Model", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                ' 1. Insert model WITHOUT ParentSKU, get modelId
                Dim modelId As Integer = db.InsertModel(selectedManufacturerId, selectedSeriesId, modelName, width, depth, height,
                totalFabricSquareInches, optionalHeight, optionalDepth, angleType, ampHandleLocation,
                tahWidth, tahHeight, sahWidth, sahHeight,
                tahRearOffset, sahRearOffset, sahTopDownOffset,
                musicRestDesign, chartTemplate, notes, Nothing)

                ' 2. Generate ParentSKU using the new modelId
                Dim parentSku As String = ModelSkuBuilder.GenerateParentSku(manufacturerName, seriesName, modelName, "V1", modelId)

                ' 3. Update the model record with the generated ParentSKU
                db.UpdateModelParentSku(modelId, parentSku)

                ' 4. Update the ParentSKU cell in the grid
                If dgvModelInformation.Columns.Contains("ParentSKU") Then
                    row.Cells("ParentSKU").Value = parentSku
                End If

                SaveModelHistoryCostRetailPricing(modelId, totalFabricSquareInches, notes)
                LoadModelsForSelectedSeries()
            Else
                Dim drv As DataRowView = DirectCast(row.DataBoundItem, DataRowView)
                Dim modelId As Integer = CInt(drv("PK_ModelId"))
                ' Generate ParentSKU using the modelId
                Dim parentSku As String = ModelSkuBuilder.GenerateParentSku(manufacturerName, seriesName, modelName, "V1", modelId)

                db.UpdateModel(modelId, modelName, width, depth, height, totalFabricSquareInches,
                optionalHeight, optionalDepth, angleType, ampHandleLocation, tahWidth, tahHeight,
                sahWidth, sahHeight, tahRearOffset, sahRearOffset, sahTopDownOffset,
                musicRestDesign, chartTemplate, notes, parentSku)

                If dgvModelInformation.Columns.Contains("ParentSKU") Then
                    row.Cells("ParentSKU").Value = parentSku
                End If

                SaveModelHistoryCostRetailPricing(modelId, totalFabricSquareInches, notes)
                LoadModelsForSelectedSeries()
            End If
        End If
    End Sub

    Private Sub dgvModelInformation_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvModelInformation.RowEnter
        Dim row = dgvModelInformation.Rows(e.RowIndex)
        If row.IsNewRow Then Exit Sub

        Dim values As New Dictionary(Of String, Object)
        For Each col As DataGridViewColumn In dgvModelInformation.Columns
            values(col.Name) = row.Cells(col.Name).Value
        Next
        originalRowValues(e.RowIndex) = values
    End Sub

    Private Sub dgvModelInformation_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvModelInformation.RowValidating
        If isLoadingModels Then Return

        Dim row = dgvModelInformation.Rows(e.RowIndex)
        If row.IsNewRow AndAlso row.Cells.Cast(Of DataGridViewCell).All(Function(c) c.Value Is Nothing OrElse String.IsNullOrWhiteSpace(c.FormattedValue?.ToString())) Then
            Return
        End If

        ' Use IDs, not names
        Dim selectedManufacturerId As Integer
        If cmbManufacturerName.SelectedValue Is Nothing OrElse Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
            MessageBox.Show("Please select a valid manufacturer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Return
        End If

        Dim selectedSeriesId As Integer
        If cmbSeriesName.SelectedValue Is Nothing OrElse Not Integer.TryParse(cmbSeriesName.SelectedValue.ToString(), selectedSeriesId) Then
            MessageBox.Show("Please select a valid series.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Return
        End If

        dgvModelInformation.EndEdit()

        ' Gather all values (unchanged)
        Dim manufacturerName As String = cmbManufacturerName.Text.Trim()
        Dim seriesName As String = cmbSeriesName.Text.Trim()
        Dim modelName As String = If(dgvModelInformation.Columns.Contains("ModelName"), Convert.ToString(row.Cells("ModelName").Value)?.Trim(), Nothing)
        Dim width As Decimal = If(dgvModelInformation.Columns.Contains("Width") AndAlso Not IsDBNull(row.Cells("Width").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Width").Value.ToString()), Convert.ToDecimal(row.Cells("Width").Value), 0D)
        Dim depth As Decimal = If(dgvModelInformation.Columns.Contains("Depth") AndAlso Not IsDBNull(row.Cells("Depth").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Depth").Value.ToString()), Convert.ToDecimal(row.Cells("Depth").Value), 0D)
        Dim height As Decimal = If(dgvModelInformation.Columns.Contains("Height") AndAlso Not IsDBNull(row.Cells("Height").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Height").Value.ToString()), Convert.ToDecimal(row.Cells("Height").Value), 0D)
        Dim optionalHeight As Decimal? = If(dgvModelInformation.Columns.Contains("OptionalHeight") AndAlso Not IsDBNull(row.Cells("OptionalHeight").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("OptionalHeight").Value.ToString()), Convert.ToDecimal(row.Cells("OptionalHeight").Value), Nothing)
        Dim optionalDepth As Decimal? = If(dgvModelInformation.Columns.Contains("OptionalDepth") AndAlso Not IsDBNull(row.Cells("OptionalDepth").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("OptionalDepth").Value.ToString()), Convert.ToDecimal(row.Cells("OptionalDepth").Value), Nothing)
        Dim angleType As String = If(dgvModelInformation.Columns.Contains("AngleType"), Convert.ToString(row.Cells("AngleType").Value), Nothing)
        Dim ampHandleLocation As String = If(dgvModelInformation.Columns.Contains("AmpHandleLocation"), Convert.ToString(row.Cells("AmpHandleLocation").Value), Nothing)
        Dim tahWidth As Decimal? = If(dgvModelInformation.Columns.Contains("TAHWidth") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHWidth").Value)), Convert.ToDecimal(row.Cells("TAHWidth").Value), Nothing)
        Dim tahHeight As Decimal? = If(dgvModelInformation.Columns.Contains("TAHHeight") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHHeight").Value)), Convert.ToDecimal(row.Cells("TAHHeight").Value), Nothing)
        Dim sahWidth As Decimal? = If(dgvModelInformation.Columns.Contains("SAHWidth") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHWidth").Value)), Convert.ToDecimal(row.Cells("SAHWidth").Value), Nothing)
        Dim sahHeight As Decimal? = If(dgvModelInformation.Columns.Contains("SAHHeight") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHHeight").Value)), Convert.ToDecimal(row.Cells("SAHHeight").Value), Nothing)
        Dim tahRearOffset As Decimal? = If(dgvModelInformation.Columns.Contains("TAHRearOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHRearOffset").Value)), Convert.ToDecimal(row.Cells("TAHRearOffset").Value), Nothing)
        Dim sahRearOffset As Decimal? = If(dgvModelInformation.Columns.Contains("SAHRearOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHRearOffset").Value)), Convert.ToDecimal(row.Cells("SAHRearOffset").Value), Nothing)
        Dim sahTopDownOffset As Decimal? = If(dgvModelInformation.Columns.Contains("SAHTopDownOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHTopDownOffset").Value)), Convert.ToDecimal(row.Cells("SAHTopDownOffset").Value), Nothing)
        Dim musicRestDesign As Boolean? = If(dgvModelInformation.Columns.Contains("MusicRestDesign") AndAlso row.Cells("MusicRestDesign").Value IsNot Nothing, Convert.ToBoolean(row.Cells("MusicRestDesign").Value), Nothing)
        Dim chartTemplate As Boolean? = If(dgvModelInformation.Columns.Contains("Chart_Template") AndAlso row.Cells("Chart_Template").Value IsNot Nothing, Convert.ToBoolean(row.Cells("Chart_Template").Value), Nothing)
        Dim notes As String = If(dgvModelInformation.Columns.Contains("Notes"), Convert.ToString(row.Cells("Notes").Value), Nothing)

        Dim db As New DbConnectionManager()
        Dim totalFabricSquareInches As Decimal = ((depth + 1.25D) * (height + 1.25D) * 2D) + ((depth + 1.25D + height + 1.25D + height + 1.25D) * (width + 1.25D))

        ' Update the cell in the grid
        If dgvModelInformation.Columns.Contains("TotalFabricSquareInches") Then
            row.Cells("TotalFabricSquareInches").Value = totalFabricSquareInches
        End If

        Dim originalModelName As String = ""
        If row.DataBoundItem IsNot Nothing AndAlso TypeOf row.DataBoundItem Is DataRowView Then
            originalModelName = Convert.ToString(DirectCast(row.DataBoundItem, DataRowView)("ModelName"))
        End If

        Dim isInsert As Boolean = String.IsNullOrEmpty(originalModelName)

        If isInsert Then
            If db.ModelExists(selectedManufacturerId, selectedSeriesId, modelName) Then
                MessageBox.Show("A model with this name already exists for the selected manufacturer and series.", "Duplicate Model", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                e.Cancel = True
                Return
            End If

            ' 1. Insert model WITHOUT ParentSKU, get modelId
            Dim modelId As Integer = db.InsertModel(selectedManufacturerId, selectedSeriesId, modelName, width, depth, height,
        totalFabricSquareInches, optionalHeight, optionalDepth, angleType, ampHandleLocation,
        tahWidth, tahHeight, sahWidth, sahHeight,
        tahRearOffset, sahRearOffset, sahTopDownOffset,
        musicRestDesign, chartTemplate, notes, Nothing)

            ' 2. Generate ParentSKU using the new modelId
            Dim parentSku As String = ModelSkuBuilder.GenerateParentSku(manufacturerName, seriesName, modelName, "V1", modelId)

            ' 3. Update the model record with the generated ParentSKU
            db.UpdateModelParentSku(modelId, parentSku)

            ' 4. Update the ParentSKU cell in the grid
            If dgvModelInformation.Columns.Contains("ParentSKU") Then
                row.Cells("ParentSKU").Value = parentSku
            End If

            SaveModelHistoryCostRetailPricing(modelId, totalFabricSquareInches, notes)
            LoadModelsForSelectedSeries()
        Else
            If row.DataBoundItem IsNot Nothing AndAlso TypeOf row.DataBoundItem Is DataRowView Then
                Dim drv As DataRowView = DirectCast(row.DataBoundItem, DataRowView)
                If Not IsDBNull(drv("PK_ModelId")) AndAlso IsNumeric(drv("PK_ModelId")) Then
                    Dim modelId As Integer = CInt(drv("PK_ModelId"))
                    ' Generate ParentSKU using the modelId
                    Dim parentSku As String = ModelSkuBuilder.GenerateParentSku(manufacturerName, seriesName, modelName, "V1", modelId)

                    db.UpdateModel(modelId, modelName, width, depth, height, totalFabricSquareInches,
                optionalHeight, optionalDepth, angleType, ampHandleLocation, tahWidth, tahHeight,
                sahWidth, sahHeight, tahRearOffset, sahRearOffset, sahTopDownOffset,
                musicRestDesign, chartTemplate, notes, parentSku)

                    ' Always update ParentSKU in DB on update
                    db.UpdateModelParentSku(modelId, parentSku)

                    If dgvModelInformation.Columns.Contains("ParentSKU") Then
                        row.Cells("ParentSKU").Value = parentSku
                    End If

                    SaveModelHistoryCostRetailPricing(modelId, totalFabricSquareInches, notes)
                End If
            End If
        End If
    End Sub

    Private Sub dgvModelInformation_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvModelInformation.CellValueChanged
        If e.RowIndex < 0 Then Exit Sub
        Dim row = dgvModelInformation.Rows(e.RowIndex)
        If dgvModelInformation.Columns(e.ColumnIndex).Name = "Width" OrElse
       dgvModelInformation.Columns(e.ColumnIndex).Name = "Depth" OrElse
       dgvModelInformation.Columns(e.ColumnIndex).Name = "Height" Then

            Dim width As Decimal = 0, depth As Decimal = 0, height As Decimal = 0
            Decimal.TryParse(Convert.ToString(row.Cells("Width").Value), width)
            Decimal.TryParse(Convert.ToString(row.Cells("Depth").Value), depth)
            Decimal.TryParse(Convert.ToString(row.Cells("Height").Value), height)
            Dim totalFabricSquareInches As Decimal = ((depth + 1.25D) * (height + 1.25D) * 2D) +
                                                 ((depth + 1.25D + height + 1.25D + height + 1.25D) * (width + 1.25D))
            row.Cells("TotalFabricSquareInches").Value = totalFabricSquareInches
        End If
    End Sub



    Private Sub dgvModelInformation_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvModelInformation.CellEndEdit
        dgvModelInformation.CommitEdit(DataGridViewDataErrorContexts.Commit)
        dgvModelInformation.EndEdit()
    End Sub

    Private Sub dgvModelInformation_Leave(sender As Object, e As EventArgs) Handles dgvModelInformation.Leave
        ' If the current row is the new row and has data, force commit and validation
        If dgvModelInformation.CurrentRow IsNot Nothing AndAlso dgvModelInformation.CurrentRow.IsNewRow Then
            Dim hasData = dgvModelInformation.CurrentRow.Cells.Cast(Of DataGridViewCell) _
            .Any(Function(c) c.Value IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(c.FormattedValue?.ToString()))
            If hasData Then
                ' Force the row to be validated and committed
                dgvModelInformation.CurrentRow.Selected = True
                dgvModelInformation.EndEdit()
                dgvModelInformation.CurrentCell = Nothing
                ' Explicitly call Validate to ensure RowValidating fires
                Me.Validate()
            End If
        ElseIf dgvModelInformation.IsCurrentRowDirty Then
            dgvModelInformation.EndEdit()
        End If
    End Sub

#End Region




    '*******************************************************************************************************************************************************************
    '*******************************************************************************************************************************************************************
#Region "*************START***********MODEL IMPORT/EXPORT***************************************************************************************************************"

    Private Sub btnUploadModels_Click(sender As Object, e As EventArgs) Handles btnUploadModels.Click
        Using ofd As New OpenFileDialog()
            ofd.Filter = "Excel or CSV Files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|All Files (*.*)|*.*"
            ofd.Title = "Select Model Information File"
            If ofd.ShowDialog() <> DialogResult.OK Then Return

            Dim filePath = ofd.FileName
            Dim dt As New DataTable()

            If Path.GetExtension(filePath).ToLower() = ".csv" Then
                ' --- CSV Handling ---
                Dim lines = File.ReadAllLines(filePath)
                If lines.Length < 2 Then
                    MessageBox.Show("The selected file does not contain enough data.", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
                Dim headers = lines(0).Split(","c).Select(Function(h) h.Trim()).ToArray()
                For Each header In headers
                    dt.Columns.Add(header)
                Next
                For i As Integer = 1 To lines.Length - 1
                    Dim values = lines(i).Split(","c).Select(Function(v) v.Trim()).ToArray()
                    If values.Length = headers.Length Then
                        dt.Rows.Add(values)
                    End If
                Next
            ElseIf Path.GetExtension(filePath).ToLower() = ".xlsx" OrElse Path.GetExtension(filePath).ToLower() = ".xls" Then
                ' --- Excel Handling ---
                Dim connStr As String
                If Path.GetExtension(filePath).ToLower() = ".xlsx" Then
                    connStr = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES';"
                Else
                    connStr = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath};Extended Properties='Excel 8.0;HDR=YES';"
                End If
                Try
                    Using conn As New OleDbConnection(connStr)
                        conn.Open()
                        Dim dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                        Dim sheetName = dtSchema.Rows(0)("TABLE_NAME").ToString()
                        Using cmd As New OleDbCommand($"SELECT * FROM [{sheetName}]", conn)
                            Using da As New OleDbDataAdapter(cmd)
                                da.Fill(dt)
                            End Using
                        End Using
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error reading Excel file: " & ex.Message, "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End Try
            Else
                MessageBox.Show("Unsupported file type.", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' --- Validate required columns ---
            Dim requiredCols = {"ManufacturerName", "SeriesName", "ModelName", "Width", "Depth", "Height"}
            For Each col In requiredCols
                If Not dt.Columns.Contains(col) Then
                    MessageBox.Show($"Missing required column: {col}", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            Next

            Dim db As New DbConnectionManager()
            Dim insertedCount As Integer = 0
            For Each row As DataRow In dt.Rows
                Dim manufacturerName = row("ManufacturerName").ToString().Trim()
                Dim seriesName = row("SeriesName").ToString().Trim()
                Dim modelName = row("ModelName").ToString().Trim()
                Dim widthStr = row("Width").ToString().Trim()
                Dim depthStr = row("Depth").ToString().Trim()
                Dim heightStr = row("Height").ToString().Trim()

                ' Skip rows missing any required value
                If String.IsNullOrEmpty(manufacturerName) OrElse String.IsNullOrEmpty(seriesName) OrElse
               String.IsNullOrEmpty(modelName) OrElse String.IsNullOrEmpty(widthStr) OrElse
               String.IsNullOrEmpty(depthStr) OrElse String.IsNullOrEmpty(heightStr) Then
                    Continue For
                End If

                Dim width As Decimal, depth As Decimal, height As Decimal
                If Not Decimal.TryParse(widthStr, width) OrElse Not Decimal.TryParse(depthStr, depth) OrElse Not Decimal.TryParse(heightStr, height) Then
                    Continue For
                End If

                ' --- Lookup ManufacturerId ---
                Dim manufacturerId As Integer = db.GetManufacturerIdByName(manufacturerName)
                If manufacturerId = 0 Then
                    manufacturerId = db.InsertManufacturer(manufacturerName)
                End If

                ' --- Lookup SeriesId ---
                Dim seriesId As Integer = db.GetSeriesIdByNameAndManufacturer(seriesName, manufacturerId)
                If seriesId = 0 Then
                    ' You may need to provide a default EquipmentTypeId if not present in the file
                    Dim defaultEquipmentTypeId As Integer = 1 ' Change as needed
                    seriesId = db.InsertSeries(seriesName, manufacturerId, defaultEquipmentTypeId)
                End If

                ' Optional fields (safe parsing)
                Dim optionalHeight As Decimal? = Nothing
                If dt.Columns.Contains("OptionalHeight") AndAlso Not IsDBNull(row("OptionalHeight")) Then
                    Dim val = row("OptionalHeight").ToString().Trim()
                    If val <> "" Then
                        Dim parsed As Decimal
                        If Decimal.TryParse(val, parsed) Then optionalHeight = parsed
                    End If
                End If

                Dim optionalDepth As Decimal? = Nothing
                If dt.Columns.Contains("OptionalDepth") AndAlso Not IsDBNull(row("OptionalDepth")) Then
                    Dim val = row("OptionalDepth").ToString().Trim()
                    If val <> "" Then
                        Dim parsed As Decimal
                        If Decimal.TryParse(val, parsed) Then optionalDepth = parsed
                    End If
                End If

                Dim angleType = If(dt.Columns.Contains("AngleType"), row("AngleType").ToString(), Nothing)
                Dim ampHandleLocation = If(dt.Columns.Contains("AmpHandleLocation"), row("AmpHandleLocation").ToString(), Nothing)

                Dim tahWidth As Decimal? = Nothing
                If dt.Columns.Contains("TAHWidth") AndAlso Not IsDBNull(row("TAHWidth")) Then
                    Dim val = row("TAHWidth").ToString().Trim()
                    If val <> "" Then
                        Dim parsed As Decimal
                        If Decimal.TryParse(val, parsed) Then tahWidth = parsed
                    End If
                End If

                Dim tahHeight As Decimal? = Nothing
                If dt.Columns.Contains("TAHHeight") AndAlso Not IsDBNull(row("TAHHeight")) Then
                    Dim val = row("TAHHeight").ToString().Trim()
                    If val <> "" Then
                        Dim parsed As Decimal
                        If Decimal.TryParse(val, parsed) Then tahHeight = parsed
                    End If
                End If

                Dim sahWidth As Decimal? = Nothing
                If dt.Columns.Contains("SAHWidth") AndAlso Not IsDBNull(row("SAHWidth")) Then
                    Dim val = row("SAHWidth").ToString().Trim()
                    If val <> "" Then
                        Dim parsed As Decimal
                        If Decimal.TryParse(val, parsed) Then sahWidth = parsed
                    End If
                End If

                Dim sahHeight As Decimal? = Nothing
                If dt.Columns.Contains("SAHHeight") AndAlso Not IsDBNull(row("SAHHeight")) Then
                    Dim val = row("SAHHeight").ToString().Trim()
                    If val <> "" Then
                        Dim parsed As Decimal
                        If Decimal.TryParse(val, parsed) Then sahHeight = parsed
                    End If
                End If

                Dim tahRearOffset As Decimal? = Nothing
                If dt.Columns.Contains("TAHRearOffset") AndAlso Not IsDBNull(row("TAHRearOffset")) Then
                    Dim val = row("TAHRearOffset").ToString().Trim()
                    If val <> "" Then
                        Dim parsed As Decimal
                        If Decimal.TryParse(val, parsed) Then tahRearOffset = parsed
                    End If
                End If

                Dim sahRearOffset As Decimal? = Nothing
                If dt.Columns.Contains("SAHRearOffset") AndAlso Not IsDBNull(row("SAHRearOffset")) Then
                    Dim val = row("SAHRearOffset").ToString().Trim()
                    If val <> "" Then
                        Dim parsed As Decimal
                        If Decimal.TryParse(val, parsed) Then sahRearOffset = parsed
                    End If
                End If

                Dim sahTopDownOffset As Decimal? = Nothing
                If dt.Columns.Contains("SAHTopDownOffset") AndAlso Not IsDBNull(row("SAHTopDownOffset")) Then
                    Dim val = row("SAHTopDownOffset").ToString().Trim()
                    If val <> "" Then
                        Dim parsed As Decimal
                        If Decimal.TryParse(val, parsed) Then sahTopDownOffset = parsed
                    End If
                End If

                Dim musicRestDesign As Boolean? = Nothing
                If dt.Columns.Contains("MusicRestDesign") AndAlso Not IsDBNull(row("MusicRestDesign")) Then
                    Dim val = row("MusicRestDesign").ToString().Trim()
                    If val <> "" Then
                        Dim parsed As Boolean
                        If Boolean.TryParse(val, parsed) Then musicRestDesign = parsed
                    End If
                End If

                Dim chartTemplate As Boolean? = Nothing
                If dt.Columns.Contains("Chart_Template") AndAlso Not IsDBNull(row("Chart_Template")) Then
                    Dim val = row("Chart_Template").ToString().Trim()
                    If val <> "" Then
                        Dim parsed As Boolean
                        If Boolean.TryParse(val, parsed) Then chartTemplate = parsed
                    End If
                End If

                Dim notes = If(dt.Columns.Contains("Notes"), row("Notes").ToString(), Nothing)

                Dim totalFabricSquareInches As Decimal = ((depth + 1.25D) * (height + 1.25D) * 2D) +
                                                     ((depth + 1.25D + height + 1.25D + height + 1.25D) * (width + 1.25D))

                If db.ModelExists(manufacturerId, seriesId, modelName) Then Continue For

                Try
                    ' 1. Insert model WITHOUT ParentSKU, get modelId
                    Dim modelId As Integer = db.InsertModel(manufacturerId, seriesId, modelName, width, depth, height,
                    totalFabricSquareInches, optionalHeight, optionalDepth, angleType, ampHandleLocation,
                    tahWidth, tahHeight, sahWidth, sahHeight,
                    tahRearOffset, sahRearOffset, sahTopDownOffset,
                    musicRestDesign, chartTemplate, notes, Nothing)

                    ' 2. Generate ParentSKU using the new modelId
                    Dim parentSku As String = ModelSkuBuilder.GenerateParentSku(manufacturerName, seriesName, modelName, "V1", modelId)

                    ' 3. Update the model record with the generated ParentSKU
                    db.UpdateModelParentSku(modelId, parentSku)

                    insertedCount += 1
                Catch ex As Exception
                    MessageBox.Show($"Error inserting model '{modelName}': {ex.Message}", "Insert Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Next

            LoadModelsForSelectedSeries()
            MessageBox.Show($"{insertedCount} models imported successfully.", "Import Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Using
    End Sub





    Private Async Sub btnUploadWooListings_Click(sender As Object, e As EventArgs) Handles btnUploadWooListings.Click
        If currentModelsTable Is Nothing OrElse currentModelsTable.Rows.Count = 0 Then
            MessageBox.Show("No models to upload.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim manufacturer As String = cmbManufacturerName.Text.Trim()
        Dim series As String = cmbSeriesName.Text.Trim()
        Dim equipmentType As String = cmbEquipmentType.Text.Trim()

        Dim rowsToUpload As New List(Of DataRow)
        If dgvModelInformation.SelectedRows.Count = 1 AndAlso Not dgvModelInformation.SelectedRows(0).IsNewRow Then
            Dim selectedIndex As Integer = dgvModelInformation.SelectedRows(0).Index
            If selectedIndex >= 0 AndAlso selectedIndex < currentModelsTable.Rows.Count Then
                rowsToUpload.Add(currentModelsTable.Rows(selectedIndex))
            End If
        Else
            For Each dr As DataRow In currentModelsTable.Rows
                rowsToUpload.Add(dr)
            Next
        End If

        Dim successCount As Integer = 0
        Dim failCount As Integer = 0

        For Each dr As DataRow In rowsToUpload
            Dim modelName As String = dr.Field(Of String)("ModelName")
            If String.IsNullOrWhiteSpace(modelName) Then Continue For

            Dim parentSku As String = ""
            If currentModelsTable.Columns.Contains("ParentSKU") AndAlso Not IsDBNull(dr("ParentSKU")) Then
                parentSku = dr.Field(Of String)("ParentSKU")
            End If

            ' --- FIX: Check for missing ParentSKU ---
            If String.IsNullOrWhiteSpace(parentSku) Then
                MessageBox.Show($"Model '{modelName}' is missing a ParentSKU. Fix before uploading.", "Missing SKU", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                failCount += 1
                Continue For
            End If

            Dim equipmentTypeForRow As String = equipmentType
            If currentModelsTable.Columns.Contains("EquipmentType") AndAlso Not IsDBNull(dr("EquipmentType")) Then
                equipmentTypeForRow = dr.Field(Of String)("EquipmentType")
            End If

            Dim categoryName As String = $"{equipmentTypeForRow} Cover"
            Dim db As New DbConnectionManager()
            Dim wooCategoryId As Integer = db.GetWooCategoryIdByName(categoryName)

            Dim productTitle As String = $"{manufacturer} {series} {modelName} {equipmentTypeForRow} Choice Waterproof Fabric or Premium Synthetic Leather Ultimate Custom Cover by GGC | Made in the U.S.A."

            Dim product As New MpWooCommerceProduct With {
            .name = productTitle,
            .type = "variable",
            .status = "publish",
            .catalog_visibility = "visible",
            .sku = parentSku,
            .categories = New List(Of MpWooCommerceProduct.MpWooCategory) From {
                New MpWooCommerceProduct.MpWooCategory With {.id = wooCategoryId}
            },
            .attributes = New List(Of MpWooAttribute) From {
                New MpWooAttribute With {
                    .name = "Fabric",
                    .variation = True,
                    .visible = True,
                    .options = New List(Of String) From {"Choice Waterproof Fabric", "Premium Synthetic Leather"}
                },
                New MpWooAttribute With {
                    .name = "Color",
                    .variation = True,
                    .visible = True,
                    .options = New List(Of String) From {"Pitch-Black", "Fire-Engine-Red", "Pacific-Blue", "Magnificent-White", "Royal-Purple", "Marbled-Brown", "Jade-Green"}
                },
                New MpWooAttribute With {
                    .name = "Manufacturer",
                    .variation = False,
                    .visible = True,
                    .options = New List(Of String) From {manufacturer}
                },
                New MpWooAttribute With {
                    .name = "Series",
                    .variation = False,
                    .visible = True,
                    .options = New List(Of String) From {series}
                },
                New MpWooAttribute With {
                    .name = "Model",
                    .variation = False,
                    .visible = True,
                    .options = New List(Of String) From {modelName}
                }
            }
        }

            ' --- ADD IMAGES HERE ---
            ' --- ADD IMAGES FROM DB ---
            Dim imagesDt = db.GetImagesForEquipmentTypeAndMarketplace(1, 6) ' 1 = Guitar Amplifier, 6 = WooCommerce
            product.images = New List(Of MpWooCommerceProduct.MpWooImage)()
            For Each imgRow As DataRow In imagesDt.Rows
                product.images.Add(New MpWooCommerceProduct.MpWooImage With {
        .src = imgRow.Field(Of String)("ImageUrl"),
        .name = imgRow.Field(Of String)("ImageType"),
        .alt = imgRow.Field(Of String)("AltText"),
        .position = If(imgRow.Field(Of Integer?)("Position"), 0)
    })
            Next

            ' --- END IMAGES SECTION ---

            ' Get the modelId from your DataRow
            Dim modelId As Integer = 0
            If dr.Table.Columns.Contains("PK_ModelId") AndAlso Not IsDBNull(dr("PK_ModelId")) Then
                modelId = CInt(dr("PK_ModelId"))
            Else
                failCount += 1
                Continue For
            End If

            Dim wooProductId As Integer = db.GetWooProductIdByModelId(modelId)

            Try
                If wooProductId = 0 Then
                    ' Create new product
                    Dim resultJson = Await WooCommerceApi.UploadProductAsync(product)
                    Dim newWooId = WooCommerceApi.ParseWooProductIdFromResult(resultJson)
                    db.SetWooProductIdForModel(modelId, newWooId)
                    WooCommerceApi.UpdateWooImageIds(resultJson)
                Else
                    ' Update existing product
                    Dim resultJson = Await WooCommerceApi.UpdateProductAsync(product, wooProductId)
                    WooCommerceApi.UpdateWooImageIds(resultJson)
                End If
                successCount += 1
            Catch ex As Exception
                MessageBox.Show(ex.ToString(), "WooCommerce Upload Error")
                failCount += 1
            End Try
        Next

        MessageBox.Show($"Upload complete! {successCount} products uploaded, {failCount} failed.", "WooCommerce Upload", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub


#End Region '----------------------------------------------------------------------------------------------------------------------------------------------------------


    '*******************************************************************************************************************************************************************
    '*******************************************************************************************************************************************************************
#Region "*************START***********MODEL CALCULATION HELPERS***************************************************************************************************************"

    Private Function CalculateModelMaterialCosts(totalFabricSquareInches As Decimal, Optional wastePercent As Decimal = 5D) As (costPerSqInch_ChoiceWaterproof As Decimal?, costPerSqInch_PremiumSyntheticLeather As Decimal?, costPerSqInch_Padding As Decimal?, baseCost_ChoiceWaterproof As Decimal?, baseCost_PremiumSyntheticLeather As Decimal?, baseCost_ChoiceWaterproof_Padded As Decimal?, baseCost_PremiumSyntheticLeather_Padded As Decimal?, baseCost_PaddingOnly As Decimal?)
        Dim db As New DbConnectionManager()
        Dim totalSqInchesWithWaste As Decimal = totalFabricSquareInches * (1 + wastePercent / 100D)

        Dim costPerSqInch_ChoiceWaterproof = db.GetActiveFabricCostPerSqInch("Choice Waterproof")
        Dim costPerSqInch_PremiumSyntheticLeather = db.GetActiveFabricCostPerSqInch("Premium Synthetic Leather")
        Dim costPerSqInch_Padding = db.GetActiveFabricCostPerSqInch("Padding")

        Dim baseCost_ChoiceWaterproof As Decimal? = If(costPerSqInch_ChoiceWaterproof.HasValue, totalSqInchesWithWaste * costPerSqInch_ChoiceWaterproof.Value, Nothing)
        Dim baseCost_PremiumSyntheticLeather As Decimal? = If(costPerSqInch_PremiumSyntheticLeather.HasValue, totalSqInchesWithWaste * costPerSqInch_PremiumSyntheticLeather.Value, Nothing)
        Dim baseCost_ChoiceWaterproof_Padded As Decimal? = If(costPerSqInch_ChoiceWaterproof.HasValue AndAlso costPerSqInch_Padding.HasValue, (costPerSqInch_ChoiceWaterproof.Value + costPerSqInch_Padding.Value) * totalSqInchesWithWaste, Nothing)
        Dim baseCost_PremiumSyntheticLeather_Padded As Decimal? = If(costPerSqInch_PremiumSyntheticLeather.HasValue AndAlso costPerSqInch_Padding.HasValue, (costPerSqInch_PremiumSyntheticLeather.Value + costPerSqInch_Padding.Value) * totalSqInchesWithWaste, Nothing)
        Dim baseCost_PaddingOnly As Decimal? = If(costPerSqInch_Padding.HasValue, costPerSqInch_Padding.Value * totalSqInchesWithWaste, Nothing)

        Return (costPerSqInch_ChoiceWaterproof, costPerSqInch_PremiumSyntheticLeather, costPerSqInch_Padding, baseCost_ChoiceWaterproof, baseCost_PremiumSyntheticLeather, baseCost_ChoiceWaterproof_Padded, baseCost_PremiumSyntheticLeather_Padded, baseCost_PaddingOnly)
    End Function



    Private Sub SaveModelHistoryCostRetailPricing(modelId As Integer, totalFabricSquareInches As Decimal, notes As String)
        Dim wastePercent As Decimal = 5D
        Dim costs = CalculateModelMaterialCosts(totalFabricSquareInches, wastePercent)
        Dim weights = CalculateModelMaterialWeights(totalFabricSquareInches)
        Dim db As New DbConnectionManager()

        ' Shipping
        Dim shipping_Choice As Decimal = db.GetShippingCostByWeight(If(weights.weight_ChoiceWaterproof, 0D))
        Dim shipping_ChoicePadded As Decimal = db.GetShippingCostByWeight(If(weights.weight_ChoiceWaterproof_Padded, 0D))
        Dim shipping_Leather As Decimal = db.GetShippingCostByWeight(If(weights.weight_PremiumSyntheticLeather, 0D))
        Dim shipping_LeatherPadded As Decimal = db.GetShippingCostByWeight(If(weights.weight_PremiumSyntheticLeather_Padded, 0D))

        ' Base fabric cost + shipping
        Dim baseFabricCost_Choice_Weight As Decimal? = If(costs.baseCost_ChoiceWaterproof, 0D) + shipping_Choice
        Dim baseFabricCost_ChoicePadding_Weight As Decimal? = If(costs.baseCost_ChoiceWaterproof_Padded, 0D) + shipping_ChoicePadded
        Dim baseFabricCost_Leather_Weight As Decimal? = If(costs.baseCost_PremiumSyntheticLeather, 0D) + shipping_Leather
        Dim baseFabricCost_LeatherPadding_Weight As Decimal? = If(costs.baseCost_PremiumSyntheticLeather_Padded, 0D) + shipping_LeatherPadded

        ' Labor
        Dim hourlyRate As Decimal = 17D
        Dim CostLaborNoPadding As Decimal = 0.5D * hourlyRate
        Dim CostLaborWithPadding As Decimal = 1D * hourlyRate

        Dim BaseCost_Choice_Labor As Decimal? = If(baseFabricCost_Choice_Weight, 0D) + CostLaborNoPadding
        Dim BaseCost_ChoicePadding_Labor As Decimal? = If(baseFabricCost_ChoicePadding_Weight, 0D) + CostLaborWithPadding
        Dim BaseCost_Leather_Labor As Decimal? = If(baseFabricCost_Leather_Weight, 0D) + CostLaborNoPadding
        Dim BaseCost_LeatherPadding_Labor As Decimal? = If(baseFabricCost_LeatherPadding_Weight, 0D) + CostLaborWithPadding

        ' Profits
        Dim profit_Choice As Decimal = db.GetLatestProfitValue("Choice")
        Dim profit_ChoicePadded As Decimal = db.GetLatestProfitValue("ChoicePadded")
        Dim profit_Leather As Decimal = db.GetLatestProfitValue("Leather")
        Dim profit_LeatherPadded As Decimal = db.GetLatestProfitValue("LeatherPadded")

        ' Marketplace fees
        Dim amazonFeePct As Decimal = db.GetMarketplaceFeePercentage("Amazon")
        Dim reverbFeePct As Decimal = db.GetMarketplaceFeePercentage("Reverb")
        Dim ebayFeePct As Decimal = db.GetMarketplaceFeePercentage("eBay")
        Dim etsyFeePct As Decimal = db.GetMarketplaceFeePercentage("Etsy")

        ' Retail (for fee calculation only)
        Dim retail_Choice = BaseCost_Choice_Labor + profit_Choice
        Dim retail_ChoicePadded = BaseCost_ChoicePadding_Labor + profit_ChoicePadded
        Dim retail_Leather = BaseCost_Leather_Labor + profit_Leather
        Dim retail_LeatherPadded = BaseCost_LeatherPadding_Labor + profit_LeatherPadded

        ' Marketplace fees
        Dim AmazonFee_Choice = retail_Choice * (amazonFeePct / 100D)
        Dim AmazonFee_ChoicePadded = retail_ChoicePadded * (amazonFeePct / 100D)
        Dim AmazonFee_Leather = retail_Leather * (amazonFeePct / 100D)
        Dim AmazonFee_LeatherPadded = retail_LeatherPadded * (amazonFeePct / 100D)

        Dim ReverbFee_Choice = retail_Choice * (reverbFeePct / 100D)
        Dim ReverbFee_ChoicePadded = retail_ChoicePadded * (reverbFeePct / 100D)
        Dim ReverbFee_Leather = retail_Leather * (reverbFeePct / 100D)
        Dim ReverbFee_LeatherPadded = retail_LeatherPadded * (reverbFeePct / 100D)

        Dim eBayFee_Choice = retail_Choice * (ebayFeePct / 100D)
        Dim eBayFee_ChoicePadded = retail_ChoicePadded * (ebayFeePct / 100D)
        Dim eBayFee_Leather = retail_Leather * (ebayFeePct / 100D)
        Dim eBayFee_LeatherPadded = retail_LeatherPadded * (ebayFeePct / 100D)

        Dim etsyFee_Choice = retail_Choice * (etsyFeePct / 100D)
        Dim etsyFee_ChoicePadded = retail_ChoicePadded * (etsyFeePct / 100D)
        Dim etsyFee_Leather = retail_Leather * (etsyFeePct / 100D)
        Dim etsyFee_LeatherPadded = retail_LeatherPadded * (etsyFeePct / 100D)

        ' Grand totals (actual cost + marketplace fee, profit NOT included)
        Dim BaseCost_GrandTotal_Choice_Amazon = BaseCost_Choice_Labor + AmazonFee_Choice
        Dim BaseCost_GrandTotal_ChoicePadded_Amazon = BaseCost_ChoicePadding_Labor + AmazonFee_ChoicePadded
        Dim BaseCost_GrandTotal_Leather_Amazon = BaseCost_Leather_Labor + AmazonFee_Leather
        Dim BaseCost_GrandTotal_LeatherPadded_Amazon = BaseCost_LeatherPadding_Labor + AmazonFee_LeatherPadded

        Dim BaseCost_GrandTotal_Choice_Reverb = BaseCost_Choice_Labor + ReverbFee_Choice
        Dim BaseCost_GrandTotal_ChoicePadded_Reverb = BaseCost_ChoicePadding_Labor + ReverbFee_ChoicePadded
        Dim BaseCost_GrandTotal_Leather_Reverb = BaseCost_Leather_Labor + ReverbFee_Leather
        Dim BaseCost_GrandTotal_LeatherPadded_Reverb = BaseCost_LeatherPadding_Labor + ReverbFee_LeatherPadded

        Dim BaseCost_GrandTotal_Choice_eBay = BaseCost_Choice_Labor + eBayFee_Choice
        Dim BaseCost_GrandTotal_ChoicePadded_eBay = BaseCost_ChoicePadding_Labor + eBayFee_ChoicePadded
        Dim BaseCost_GrandTotal_Leather_eBay = BaseCost_Leather_Labor + eBayFee_Leather
        Dim BaseCost_GrandTotal_LeatherPadded_eBay = BaseCost_LeatherPadding_Labor + eBayFee_LeatherPadded

        Dim BaseCost_GrandTotal_Choice_Etsy = BaseCost_Choice_Labor + etsyFee_Choice
        Dim BaseCost_GrandTotal_ChoicePadded_Etsy = BaseCost_ChoicePadding_Labor + etsyFee_ChoicePadded
        Dim BaseCost_GrandTotal_Leather_Etsy = BaseCost_Leather_Labor + etsyFee_Leather
        Dim BaseCost_GrandTotal_LeatherPadded_Etsy = BaseCost_LeatherPadding_Labor + etsyFee_LeatherPadded

        ' === Calculate Retail Prices ===
        Dim RetailPrice_Choice_Amazon = CalculateRetailPrice(BaseCost_GrandTotal_Choice_Amazon, profit_Choice)
        Dim RetailPrice_ChoicePadded_Amazon = CalculateRetailPrice(BaseCost_GrandTotal_ChoicePadded_Amazon, profit_ChoicePadded)
        Dim RetailPrice_Leather_Amazon = CalculateRetailPrice(BaseCost_GrandTotal_Leather_Amazon, profit_Leather)
        Dim RetailPrice_LeatherPadded_Amazon = CalculateRetailPrice(BaseCost_GrandTotal_LeatherPadded_Amazon, profit_LeatherPadded)

        Dim RetailPrice_Choice_Reverb = CalculateRetailPrice(BaseCost_GrandTotal_Choice_Reverb, profit_Choice)
        Dim RetailPrice_ChoicePadded_Reverb = CalculateRetailPrice(BaseCost_GrandTotal_ChoicePadded_Reverb, profit_ChoicePadded)
        Dim RetailPrice_Leather_Reverb = CalculateRetailPrice(BaseCost_GrandTotal_Leather_Reverb, profit_Leather)
        Dim RetailPrice_LeatherPadded_Reverb = CalculateRetailPrice(BaseCost_GrandTotal_LeatherPadded_Reverb, profit_LeatherPadded)

        Dim RetailPrice_Choice_eBay = CalculateRetailPrice(BaseCost_GrandTotal_Choice_eBay, profit_Choice)
        Dim RetailPrice_ChoicePadded_eBay = CalculateRetailPrice(BaseCost_GrandTotal_ChoicePadded_eBay, profit_ChoicePadded)
        Dim RetailPrice_Leather_eBay = CalculateRetailPrice(BaseCost_GrandTotal_Leather_eBay, profit_Leather)
        Dim RetailPrice_LeatherPadded_eBay = CalculateRetailPrice(BaseCost_GrandTotal_LeatherPadded_eBay, profit_LeatherPadded)

        Dim RetailPrice_Choice_Etsy = CalculateRetailPrice(BaseCost_GrandTotal_Choice_Etsy, profit_Choice)
        Dim RetailPrice_ChoicePadded_Etsy = CalculateRetailPrice(BaseCost_GrandTotal_ChoicePadded_Etsy, profit_ChoicePadded)
        Dim RetailPrice_Leather_Etsy = CalculateRetailPrice(BaseCost_GrandTotal_Leather_Etsy, profit_Leather)
        Dim RetailPrice_LeatherPadded_Etsy = CalculateRetailPrice(BaseCost_GrandTotal_LeatherPadded_Etsy, profit_LeatherPadded)



        Dim latestRow = db.GetLatestModelHistoryCostRetailPricing(modelId)
        If latestRow IsNot Nothing Then
            If ModelHistoryRowsAreEquivalent(latestRow, totalFabricSquareInches, wastePercent, costs, weights,
            baseFabricCost_Choice_Weight, baseFabricCost_ChoicePadding_Weight, baseFabricCost_Leather_Weight,
            baseFabricCost_LeatherPadding_Weight, BaseCost_Choice_Labor, BaseCost_ChoicePadding_Labor,
            BaseCost_Leather_Labor, BaseCost_LeatherPadding_Labor, profit_Choice, profit_ChoicePadded,
            profit_Leather, profit_LeatherPadded, AmazonFee_Choice, AmazonFee_ChoicePadded, AmazonFee_Leather,
            AmazonFee_LeatherPadded, ReverbFee_Choice, ReverbFee_ChoicePadded, ReverbFee_Leather,
            ReverbFee_LeatherPadded, eBayFee_Choice, eBayFee_ChoicePadded, eBayFee_Leather, eBayFee_LeatherPadded,
            etsyFee_Choice, etsyFee_ChoicePadded, etsyFee_Leather, etsyFee_LeatherPadded,
            BaseCost_GrandTotal_Choice_Amazon, BaseCost_GrandTotal_ChoicePadded_Amazon,
            BaseCost_GrandTotal_Leather_Amazon, BaseCost_GrandTotal_LeatherPadded_Amazon,
            BaseCost_GrandTotal_Choice_Reverb, BaseCost_GrandTotal_ChoicePadded_Reverb,
            BaseCost_GrandTotal_Leather_Reverb, BaseCost_GrandTotal_LeatherPadded_Reverb,
            BaseCost_GrandTotal_Choice_eBay, BaseCost_GrandTotal_ChoicePadded_eBay,
            BaseCost_GrandTotal_Leather_eBay, BaseCost_GrandTotal_LeatherPadded_eBay,
            BaseCost_GrandTotal_Choice_Etsy, BaseCost_GrandTotal_ChoicePadded_Etsy,
            BaseCost_GrandTotal_Leather_Etsy, BaseCost_GrandTotal_LeatherPadded_Etsy,
            RetailPrice_Choice_Amazon, RetailPrice_ChoicePadded_Amazon, RetailPrice_Leather_Amazon,
            RetailPrice_LeatherPadded_Amazon, RetailPrice_Choice_Reverb, RetailPrice_ChoicePadded_Reverb,
            RetailPrice_Leather_Reverb, RetailPrice_LeatherPadded_Reverb, RetailPrice_Choice_eBay,
            RetailPrice_ChoicePadded_eBay, RetailPrice_Leather_eBay, RetailPrice_LeatherPadded_eBay,
            RetailPrice_Choice_Etsy, RetailPrice_ChoicePadded_Etsy, RetailPrice_Leather_Etsy,
            RetailPrice_LeatherPadded_Etsy, notes) Then
                ' No changes, do not insert
                Exit Sub
            End If
        End If


        db.InsertModelHistoryCostRetailPricing(
        modelId,
        costs.costPerSqInch_ChoiceWaterproof,
        costs.costPerSqInch_PremiumSyntheticLeather,
        costs.costPerSqInch_Padding,
        totalFabricSquareInches,
        wastePercent,
        costs.baseCost_ChoiceWaterproof,
        costs.baseCost_PremiumSyntheticLeather,
        costs.baseCost_ChoiceWaterproof_Padded,
        costs.baseCost_PremiumSyntheticLeather_Padded,
        costs.baseCost_PaddingOnly,
        weights.weight_PaddingOnly,
        weights.weight_ChoiceWaterproof,
        weights.weight_ChoiceWaterproof_Padded,
        weights.weight_PremiumSyntheticLeather,
        weights.weight_PremiumSyntheticLeather_Padded,
        shipping_Choice,
        shipping_ChoicePadded,
        shipping_Leather,
        shipping_LeatherPadded,
        baseFabricCost_Choice_Weight,
        baseFabricCost_ChoicePadding_Weight,
        baseFabricCost_Leather_Weight,
        baseFabricCost_LeatherPadding_Weight,
        BaseCost_Choice_Labor,
        BaseCost_ChoicePadding_Labor,
        BaseCost_Leather_Labor,
        BaseCost_LeatherPadding_Labor,
        profit_Choice,
        profit_ChoicePadded,
        profit_Leather,
        profit_LeatherPadded,
        AmazonFee_Choice,
        AmazonFee_ChoicePadded,
        AmazonFee_Leather,
        AmazonFee_LeatherPadded,
        ReverbFee_Choice,
        ReverbFee_ChoicePadded,
        ReverbFee_Leather,
        ReverbFee_LeatherPadded,
        eBayFee_Choice,
        eBayFee_ChoicePadded,
        eBayFee_Leather,
        eBayFee_LeatherPadded,
        etsyFee_Choice,
        etsyFee_ChoicePadded,
        etsyFee_Leather,
        etsyFee_LeatherPadded,
        BaseCost_GrandTotal_Choice_Amazon,
        BaseCost_GrandTotal_ChoicePadded_Amazon,
        BaseCost_GrandTotal_Leather_Amazon,
        BaseCost_GrandTotal_LeatherPadded_Amazon,
        BaseCost_GrandTotal_Choice_Reverb,
        BaseCost_GrandTotal_ChoicePadded_Reverb,
        BaseCost_GrandTotal_Leather_Reverb,
        BaseCost_GrandTotal_LeatherPadded_Reverb,
        BaseCost_GrandTotal_Choice_eBay,
        BaseCost_GrandTotal_ChoicePadded_eBay,
        BaseCost_GrandTotal_Leather_eBay,
        BaseCost_GrandTotal_LeatherPadded_eBay,
        BaseCost_GrandTotal_Choice_Etsy,
        BaseCost_GrandTotal_ChoicePadded_Etsy,
        BaseCost_GrandTotal_Leather_Etsy,
        BaseCost_GrandTotal_LeatherPadded_Etsy,
        RetailPrice_Choice_Amazon,
        RetailPrice_ChoicePadded_Amazon,
        RetailPrice_Leather_Amazon,
        RetailPrice_LeatherPadded_Amazon,
        RetailPrice_Choice_Reverb,
        RetailPrice_ChoicePadded_Reverb,
        RetailPrice_Leather_Reverb,
        RetailPrice_LeatherPadded_Reverb,
        RetailPrice_Choice_eBay,
        RetailPrice_ChoicePadded_eBay,
        RetailPrice_Leather_eBay,
        RetailPrice_LeatherPadded_eBay,
        RetailPrice_Choice_Etsy,
        RetailPrice_ChoicePadded_Etsy,
        RetailPrice_Leather_Etsy,
        RetailPrice_LeatherPadded_Etsy,
        notes
    )
    End Sub


    Private Function ModelHistoryRowsAreEquivalent(latestRow As DataRow,
    totalFabricSquareInches As Decimal, wastePercent As Decimal,
    costs As Object, weights As Object, baseFabricCost_Choice_Weight As Decimal?,
    baseFabricCost_ChoicePadding_Weight As Decimal?, baseFabricCost_Leather_Weight As Decimal?,
    baseFabricCost_LeatherPadding_Weight As Decimal?, BaseCost_Choice_Labor As Decimal?,
    BaseCost_ChoicePadding_Labor As Decimal?, BaseCost_Leather_Labor As Decimal?,
    BaseCost_LeatherPadding_Labor As Decimal?, profit_Choice As Decimal, profit_ChoicePadded As Decimal,
    profit_Leather As Decimal, profit_LeatherPadded As Decimal, AmazonFee_Choice As Decimal,
    AmazonFee_ChoicePadded As Decimal, AmazonFee_Leather As Decimal, AmazonFee_LeatherPadded As Decimal,
    ReverbFee_Choice As Decimal, ReverbFee_ChoicePadded As Decimal, ReverbFee_Leather As Decimal,
    ReverbFee_LeatherPadded As Decimal, eBayFee_Choice As Decimal, eBayFee_ChoicePadded As Decimal,
    eBayFee_Leather As Decimal, eBayFee_LeatherPadded As Decimal, etsyFee_Choice As Decimal,
    etsyFee_ChoicePadded As Decimal, etsyFee_Leather As Decimal, etsyFee_LeatherPadded As Decimal,
    BaseCost_GrandTotal_Choice_Amazon As Decimal, BaseCost_GrandTotal_ChoicePadded_Amazon As Decimal,
    BaseCost_GrandTotal_Leather_Amazon As Decimal, BaseCost_GrandTotal_LeatherPadded_Amazon As Decimal,
    BaseCost_GrandTotal_Choice_Reverb As Decimal, BaseCost_GrandTotal_ChoicePadded_Reverb As Decimal,
    BaseCost_GrandTotal_Leather_Reverb As Decimal, BaseCost_GrandTotal_LeatherPadded_Reverb As Decimal,
    BaseCost_GrandTotal_Choice_eBay As Decimal, BaseCost_GrandTotal_ChoicePadded_eBay As Decimal,
    BaseCost_GrandTotal_Leather_eBay As Decimal, BaseCost_GrandTotal_LeatherPadded_eBay As Decimal,
    BaseCost_GrandTotal_Choice_Etsy As Decimal, BaseCost_GrandTotal_ChoicePadded_Etsy As Decimal,
    BaseCost_GrandTotal_Leather_Etsy As Decimal, BaseCost_GrandTotal_LeatherPadded_Etsy As Decimal,
    RetailPrice_Choice_Amazon As Decimal, RetailPrice_ChoicePadded_Amazon As Decimal,
    RetailPrice_Leather_Amazon As Decimal, RetailPrice_LeatherPadded_Amazon As Decimal,
    RetailPrice_Choice_Reverb As Decimal, RetailPrice_ChoicePadded_Reverb As Decimal,
    RetailPrice_Leather_Reverb As Decimal, RetailPrice_LeatherPadded_Reverb As Decimal,
    RetailPrice_Choice_eBay As Decimal, RetailPrice_ChoicePadded_eBay As Decimal,
    RetailPrice_Leather_eBay As Decimal, RetailPrice_LeatherPadded_eBay As Decimal,
    RetailPrice_Choice_Etsy As Decimal, RetailPrice_ChoicePadded_Etsy As Decimal,
    RetailPrice_Leather_Etsy As Decimal, RetailPrice_LeatherPadded_Etsy As Decimal,
    notes As String) As Boolean

        ' Compare all relevant fields, return False if any are different
        If Not ValuesAreEquivalent(latestRow("TotalFabricSquareInches"), totalFabricSquareInches) Then Return False
        If Not ValuesAreEquivalent(latestRow("WastePercent"), wastePercent) Then Return False
        If Not ValuesAreEquivalent(latestRow("CostPerSqInch_ChoiceWaterproof"), costs.costPerSqInch_ChoiceWaterproof) Then Return False
        If Not ValuesAreEquivalent(latestRow("CostPerSqInch_PremiumSyntheticLeather"), costs.costPerSqInch_PremiumSyntheticLeather) Then Return False
        If Not ValuesAreEquivalent(latestRow("CostPerSqInch_Padding"), costs.costPerSqInch_Padding) Then Return False
        ' ...repeat for all other fields you want to compare...
        ' For brevity, only a few are shown. You should compare all fields that are calculated and stored.

        ' Example for notes:
        If Not ValuesAreEquivalent(latestRow("Notes"), notes) Then Return False

        ' If all fields are equivalent
        Return True
    End Function




#End Region '----------END---------------MODEL CALCULATION HELPERS----------------------------------------------------------------------------------------------------


    '*****************************************************************************************************************************************************************
#Region "*************START**************DATABASE HELPERS*************************************************************************************************************"
    '*****************************************************************************************************************************************************************
    Private Sub CheckAndAddManufacturer()
        Dim enteredName As String = cmbManufacturerName.Text.Trim()
        If String.IsNullOrEmpty(enteredName) Then Exit Sub

        Dim db As New DbConnectionManager()
        Dim existingNames = db.GetManufacturerNames()
        If Not existingNames.AsEnumerable().Any(Function(row) String.Equals(row.Field(Of String)("ManufacturerName"), enteredName, StringComparison.OrdinalIgnoreCase)) Then
            Dim result = MessageBox.Show(
            $"'{enteredName}' is not in the list. Do you want to add it as a new manufacturer?",
            "Add Manufacturer",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )
            If result = DialogResult.Yes Then
                db.InsertManufacturer(enteredName)
                ' Refresh ComboBox
                cmbManufacturerName.DataSource = Nothing
                cmbManufacturerName.DataSource = db.GetManufacturerNames()
                cmbManufacturerName.SelectedItem = enteredName
            End If
        End If
    End Sub

    Private Sub TrySaveNewSeries()
        ' Get the selected manufacturer ID (as Integer)
        Dim selectedManufacturerId As Integer
        If cmbManufacturerName.SelectedValue Is Nothing OrElse Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
            MessageBox.Show("Please select a valid manufacturer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim enteredSeries As String = cmbSeriesName.Text.Trim()
        If String.IsNullOrEmpty(enteredSeries) Then Exit Sub

        Dim db As New DbConnectionManager()
        Dim existingSeries = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
        If Not existingSeries.AsEnumerable().Any(Function(row) String.Equals(row.Field(Of String)("SeriesName"), enteredSeries, StringComparison.OrdinalIgnoreCase)) Then
            Dim equipmentType As String = cmbEquipmentType.Text.Trim()
            If String.IsNullOrEmpty(equipmentType) Then
                MessageBox.Show("Please select or enter an equipment type for this series.", "Equipment Type Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbEquipmentType.Focus()
                Exit Sub
            End If

            If cmbEquipmentType.SelectedValue Is Nothing OrElse Not IsNumeric(cmbEquipmentType.SelectedValue) Then
                MessageBox.Show("Please select a valid equipment type.", "Invalid Equipment Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbEquipmentType.Focus()
                Return
            End If

            Dim equipmentTypeValue As Integer = Convert.ToInt32(cmbEquipmentType.SelectedValue)
            Dim msg As String = $"Do you want to save this new series?" & vbCrLf &
                        $"Manufacturer ID: {selectedManufacturerId}" & vbCrLf &
                        $"Series: {enteredSeries}" & vbCrLf &
                        $"Equipment Type: {equipmentType}"
            Dim result = MessageBox.Show(msg, "Save New Series", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                db.InsertSeries(enteredSeries, selectedManufacturerId, equipmentTypeValue)
                cmbSeriesName.DataSource = Nothing
                cmbSeriesName.DataSource = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
                cmbSeriesName.DisplayMember = "SeriesName"
                cmbSeriesName.ValueMember = "PK_SeriesId"
                cmbSeriesName.SelectedItem = enteredSeries
            End If
        End If

        SetEquipmentTypeForSelectedSeries()
    End Sub

    Private Sub TrySaveOrUpdateSeries()
        ' Get the selected manufacturer ID (as Integer)
        Dim selectedManufacturerId As Integer
        If cmbManufacturerName.SelectedValue Is Nothing OrElse Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
            MessageBox.Show("Please select a valid manufacturer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim enteredSeries As String = cmbSeriesName.Text.Trim()
        If String.IsNullOrEmpty(enteredSeries) Then Exit Sub

        Dim db As New DbConnectionManager()
        Dim existingSeries = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
        If Not existingSeries.AsEnumerable().Any(Function(row) String.Equals(row.Field(Of String)("SeriesName"), enteredSeries, StringComparison.OrdinalIgnoreCase)) Then
            ' Require equipment type selection
            If cmbEquipmentType.SelectedValue Is Nothing OrElse Not IsNumeric(cmbEquipmentType.SelectedValue) Then
                MessageBox.Show("Please select a valid equipment type.", "Invalid Equipment Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbEquipmentType.Focus()
                Return
            End If

            Dim equipmentTypeValue As Integer = Convert.ToInt32(cmbEquipmentType.SelectedValue)
            Dim msg As String = $"Do you want to save this new series?" & vbCrLf &
                            $"Manufacturer ID: {selectedManufacturerId}" & vbCrLf &
                            $"Series: {enteredSeries}" & vbCrLf &
                            $"Equipment Type: {cmbEquipmentType.Text.Trim()}"
            Dim result = MessageBox.Show(msg, "Save New Series", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                db.InsertSeries(enteredSeries, selectedManufacturerId, equipmentTypeValue)
                cmbSeriesName.DataSource = Nothing
                cmbSeriesName.DataSource = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
                cmbSeriesName.DisplayMember = "SeriesName"
                cmbSeriesName.ValueMember = "PK_SeriesId"
                cmbSeriesName.SelectedItem = enteredSeries
            End If
        Else
            ' Update equipment type for existing series
            If cmbEquipmentType.SelectedValue Is Nothing OrElse Not IsNumeric(cmbEquipmentType.SelectedValue) Then
                MessageBox.Show("Please select a valid equipment type.", "Invalid Equipment Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbEquipmentType.Focus()
                Return
            End If
            Dim equipmentTypeValue As Integer = Convert.ToInt32(cmbEquipmentType.SelectedValue)
            Dim seriesId As Integer = CInt(existingSeries.AsEnumerable().First(Function(row) String.Equals(row.Field(Of String)("SeriesName"), enteredSeries, StringComparison.OrdinalIgnoreCase))("PK_SeriesId"))
            db.UpdateSeriesEquipmentType(seriesId, equipmentTypeValue)
            MessageBox.Show("Series equipment type updated.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        SetEquipmentTypeForSelectedSeries()
    End Sub

    Private Sub LoadModelsForSelectedSeries()
        isLoadingModels = True
        If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then
            dgvModelInformation.DataSource = Nothing
            isLoadingModels = False
            Return
        End If
        If cmbSeriesName.SelectedIndex = -1 OrElse cmbSeriesName.SelectedValue Is Nothing Then
            dgvModelInformation.DataSource = Nothing
            isLoadingModels = False
            Return
        End If

        Dim selectedManufacturerId As Integer
        If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
            dgvModelInformation.DataSource = Nothing
            isLoadingModels = False
            Return
        End If
        Dim selectedSeriesId As Integer
        If Not Integer.TryParse(cmbSeriesName.SelectedValue.ToString(), selectedSeriesId) Then
            dgvModelInformation.DataSource = Nothing
            isLoadingModels = False
            Return
        End If

        Dim db As New DbConnectionManager()
        Dim dt = db.GetModelsByManufacturerAndSeries(selectedManufacturerId, selectedSeriesId)
        currentModelsTable = dt
        dgvModelInformation.DataSource = currentModelsTable
        isLoadingModels = False
    End Sub

    Private Sub LoadSeriesGridForManufacturer()
        If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then
            dgvSeriesName.DataSource = Nothing
            Return
        End If

        Dim selectedManufacturerId As Integer
        If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
            dgvSeriesName.DataSource = Nothing
            Return
        End If

        Dim db As New DbConnectionManager()
        Dim dt = db.GetSeriesAndEquipmentTypeByManufacturer(selectedManufacturerId)
        dgvSeriesName.DataSource = dt
        dgvSeriesName.ClearSelection()

        ' Hide PK_SeriesId column
        If dgvSeriesName.Columns.Contains("PK_SeriesId") Then
            dgvSeriesName.Columns("PK_SeriesId").Visible = False
        End If
    End Sub


    ''' <summary>
    ''' Inserts a new model into the Model table, generating and saving the ParentSKU.
    ''' </summary>
    ''' <param name="connection">An open SqlConnection.</param>
    ''' <param name="manufacturerName">Manufacturer name.</param>
    ''' <param name="seriesName">Series name.</param>
    ''' <param name="modelName">Model name.</param>
    ''' <param name="versionSuffix">Optional version suffix (default "V1").</param>
    Public Sub InsertModelWithSku(connection As SqlConnection,
                              manufacturerName As String,
                              seriesName As String,
                              modelName As String,
                              Optional versionSuffix As String = "V1")
        ' 1. Insert the model WITHOUT ParentSKU, get the new modelId
        Dim modelId As Integer
        Using insertCmd As New SqlCommand("INSERT INTO Model (ManufacturerName, SeriesName, ModelName) VALUES (@ManufacturerName, @SeriesName, @ModelName); SELECT CAST(SCOPE_IDENTITY() AS int);", connection)
            insertCmd.Parameters.AddWithValue("@ManufacturerName", manufacturerName)
            insertCmd.Parameters.AddWithValue("@SeriesName", seriesName)
            insertCmd.Parameters.AddWithValue("@ModelName", modelName)
            modelId = CInt(insertCmd.ExecuteScalar())
        End Using

        ' 2. Generate ParentSKU using the new modelId
        Dim parentSku As String = ModelSkuBuilder.GenerateParentSku(manufacturerName, seriesName, modelName, versionSuffix, modelId)

        ' 3. Update the model record with the generated ParentSKU
        Using updateCmd As New SqlCommand("UPDATE Model SET ParentSKU = @ParentSKU WHERE PK_ModelId = @ModelId", connection)
            updateCmd.Parameters.AddWithValue("@ParentSKU", parentSku)
            updateCmd.Parameters.AddWithValue("@ModelId", modelId)
            updateCmd.ExecuteNonQuery()
        End Using
    End Sub

#End Region '---------END---------------DATABASE HELPERS--------------------------------------------------------------------------------------------------------------

    '*****************************************************************************************************************************************************************
#Region "*************START*************UI HELPERS********************************************************************************************************************"
    '*****************************************************************************************************************************************************************


    Private Sub SetEquipmentTypeForSelectedSeries()
        If isLoadingSeries Then Return
        If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then
            cmbEquipmentType.SelectedIndex = -1
            Return
        End If
        If cmbSeriesName.SelectedIndex = -1 OrElse cmbSeriesName.SelectedValue Is Nothing Then
            cmbEquipmentType.SelectedIndex = -1
            Return
        End If

        Dim selectedManufacturerId As Integer
        If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
            cmbEquipmentType.SelectedIndex = -1
            Return
        End If
        Dim selectedSeriesId As Integer
        If Not Integer.TryParse(cmbSeriesName.SelectedValue.ToString(), selectedSeriesId) Then
            cmbEquipmentType.SelectedIndex = -1
            Return
        End If

        Dim db As New DbConnectionManager()
        Dim equipmentTypeId = db.GetEquipmentTypeIdByManufacturerAndSeries(selectedManufacturerId, selectedSeriesId)

        If equipmentTypeId.HasValue Then
            cmbEquipmentType.SelectedValue = equipmentTypeId.Value
        Else
            cmbEquipmentType.SelectedIndex = -1
        End If
    End Sub

    Private Sub InitializeModelGridColumns(Optional equipmentTypeName As String = Nothing)
        dgvModelInformation.Columns.Clear()

        ' Always visible columns
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "ModelName", .HeaderText = "Model Name", .DataPropertyName = "ModelName"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "ParentSKU", .HeaderText = "Parent SKU", .DataPropertyName = "ParentSKU", .ReadOnly = True, .Width = 180})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "Width", .HeaderText = "Width", .DataPropertyName = "Width"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "Depth", .HeaderText = "Depth", .DataPropertyName = "Depth"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "Height", .HeaderText = "Height", .DataPropertyName = "Height"})

        ' Add TotalFabricSquareInches column (read-only, calculated)
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {
        .Name = "TotalFabricSquareInches",
        .HeaderText = "Total Fabric Sq In",
        .DataPropertyName = "TotalFabricSquareInches",
        .ReadOnly = True
    })
        ' Add equipment-type-specific columns
        If equipmentTypeName = "Guitar Amplifier" OrElse equipmentTypeName = "Bass Cabinet" Then
            AddGuitarAmplifierColumns()
        ElseIf equipmentTypeName = "Music Keyboard" Then
            AddMusicKeyboardColumns()
        End If

        ' Always visible for all types
        dgvModelInformation.Columns.Add(New DataGridViewCheckBoxColumn With {.Name = "Chart_Template", .HeaderText = "Chart/Template"})
        dgvModelInformation.Columns.Add("Notes", "Notes")

        Dim saveButtonCol As New DataGridViewButtonColumn With {
    .Name = "Save",
    .HeaderText = "Save",
    .Text = "Save",
    .UseColumnTextForButtonValue = True,
    .Width = 60
}
        dgvModelInformation.Columns.Add(saveButtonCol)
        FormatModelGridColumns()
    End Sub

    Private Sub AddGuitarAmplifierColumns()
        Dim db As New DbConnectionManager()

        ' Optional Height/Depth
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "OptionalHeight", .HeaderText = "Opt. Height", .DataPropertyName = "OptionalHeight"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "OptionalDepth", .HeaderText = "Opt. Depth", .DataPropertyName = "OptionalDepth"})

        ' AngleType ComboBox
        Dim angleTypeCol As New DataGridViewComboBoxColumn With {
        .Name = "AngleType",
        .HeaderText = "Angle Type",
        .DataPropertyName = "AngleType",
        .Width = 120,
        .DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
    }
        angleTypeCol.DataSource = db.GetAllAngleTypesTable()
        angleTypeCol.DisplayMember = "AngleTypeName"
        angleTypeCol.ValueMember = "AngleTypeName"
        angleTypeCol.FlatStyle = FlatStyle.Standard
        angleTypeCol.ReadOnly = False
        dgvModelInformation.Columns.Add(angleTypeCol)

        ' TAH/SAH columns
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "TAHWidth", .HeaderText = "TAH Width", .DataPropertyName = "TAHWidth"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "TAHHeight", .HeaderText = "TAH Height", .DataPropertyName = "TAHHeight"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "SAHWidth", .HeaderText = "SAH Width", .DataPropertyName = "SAHWidth"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "SAHHeight", .HeaderText = "SAH Height", .DataPropertyName = "SAHHeight"})

        ' AmpHandleLocation ComboBox
        Dim ampHandleCol As New DataGridViewComboBoxColumn With {
        .Name = "AmpHandleLocation",
        .HeaderText = "Amp Handle Location",
        .DataPropertyName = "AmpHandleLocation",
        .Width = 125,
        .DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
    }
        ampHandleCol.DataSource = db.GetAllAmpHandleLocationsTable()
        ampHandleCol.DisplayMember = "AmpHandleLocationName"
        ampHandleCol.ValueMember = "AmpHandleLocationName"
        ampHandleCol.FlatStyle = FlatStyle.Standard
        ampHandleCol.ReadOnly = False
        dgvModelInformation.Columns.Add(ampHandleCol)

        ' Rear/Offset columns
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "TAHRearOffset", .HeaderText = "TAH Rear Offset", .DataPropertyName = "TAHRearOffset"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "SAHRearOffset", .HeaderText = "SAH Rear Offset", .DataPropertyName = "SAHRearOffset"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "SAHTopDownOffset", .HeaderText = "SAH TopDown Offset", .DataPropertyName = "SAHTopDownOffset"})
    End Sub

    Private Sub AddMusicKeyboardColumns()
        dgvModelInformation.Columns.Add(New DataGridViewCheckBoxColumn With {.Name = "MusicRestDesign", .HeaderText = "Music Rest Design", .DataPropertyName = "MusicRestDesign"})
    End Sub


    Private Sub FormatModelGridColumns()
        ' Center header text and wrap
        For Each col As DataGridViewColumn In dgvModelInformation.Columns
            col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            col.HeaderCell.Style.WrapMode = DataGridViewTriState.True
        Next

        dgvModelInformation.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
        dgvModelInformation.ColumnHeadersHeight = 44 ' Adjust as needed

        ' Center cell text for specific columns
        Dim centerStyle As New DataGridViewCellStyle()
        centerStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        For Each colName In New String() {"Width", "Depth", "Height", "OptionalHeight", "OptionalDepth"}
            If dgvModelInformation.Columns.Contains(colName) Then
                dgvModelInformation.Columns(colName).DefaultCellStyle = centerStyle
            End If
        Next

        ' Set column widths if needed
        Dim widths As New Dictionary(Of String, Integer) From {
        {"ModelName", 150},
        {"Width", 50},
        {"Depth", 50},
        {"Height", 50},
        {"OptionalHeight", 50},
        {"OptionalDepth", 50},
        {"AngleType", 75},
        {"AmpHandleLocation", 125},
        {"TAHWidth", 50},
        {"TAHHeight", 50},
        {"SAHWidth", 50},
        {"SAHHeight", 50},
        {"TAHRearOffset", 80},
        {"SAHRearOffset", 80},
        {"SAHTopDownOffset", 90},
        {"Chart_Template", 90},
        {"Notes", 175}
    }
        For Each kvp In widths
            If dgvModelInformation.Columns.Contains(kvp.Key) Then
                dgvModelInformation.Columns(kvp.Key).Width = kvp.Value
            End If
        Next
    End Sub



#End Region '-------- -END---------------UI HELPERS--------------------------------------------------------------------------------------------------------------------



    '*****************************************************************************************************************************************************************
#Region "*************START*************BUTTON EVENT HANDLERS*********************************************************************************************************"
    '*****************************************************************************************************************************************************************

    Private Sub btnSaveSeries_Click(sender As Object, e As EventArgs) Handles btnSaveSeries.Click
        If String.IsNullOrEmpty(cmbManufacturerName.Text.Trim()) OrElse cmbManufacturerName.SelectedIndex = -1 Then
            MessageBox.Show("Please select a manufacturer before saving or updating a series.", "Manufacturer Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        TrySaveOrUpdateSeries()
        LoadSeriesGridForManufacturer()
    End Sub

    Private Sub btn_UpdateCosts_Click(sender As Object, e As EventArgs) Handles btn_UpdateCosts.Click
        Dim db As New DbConnectionManager()
        Dim dtModels = db.GetAllModelsWithDetails()
        Dim results As New List(Of ModelCostUpdateResults)()

        For Each row As DataRow In dtModels.Rows
            Dim modelId As Integer = CInt(row("PK_ModelId"))
            Dim modelName As String = row("ModelName").ToString()
            Dim seriesName As String = row("SeriesName").ToString()
            Dim manufacturerName As String = row("ManufacturerName").ToString()
            Dim totalFabricSquareInches As Decimal = If(IsDBNull(row("TotalFabricSquareInches")), 0D, Convert.ToDecimal(row("TotalFabricSquareInches")))
            Dim notes As String = row("Notes").ToString()

            If totalFabricSquareInches > 0 Then
                Dim dbMgr As New DbConnectionManager()
                Dim latestRow = dbMgr.GetLatestModelHistoryCostRetailPricing(modelId)
                SaveModelHistoryCostRetailPricing(modelId, totalFabricSquareInches, notes)
                Dim newRow = dbMgr.GetLatestModelHistoryCostRetailPricing(modelId)
                Dim updated As Boolean = (latestRow Is Nothing) OrElse (Not ValuesAreEquivalent(latestRow("DateCalculated"), newRow("DateCalculated")))
                Dim msg As String = If(updated, "Updated", "No change needed")
                results.Add(New ModelCostUpdateResults(modelId, modelName, seriesName, manufacturerName, updated, msg))
            Else
                results.Add(New ModelCostUpdateResults(modelId, modelName, seriesName, manufacturerName, False, "Skipped (no size)"))
            End If
        Next

        Dim updatedCount = results.Where(Function(r) r.Updated).Count()
        Dim skippedCount = results.Count - updatedCount

        MessageBox.Show(
        $"Cost update complete. {updatedCount} models updated, {skippedCount} skipped." & vbCrLf &
        String.Join(vbCrLf, results.Select(Function(r) $"{r.ManufacturerName} - {r.SeriesName} - {r.ModelName}: {r.Message}")),
        "Update Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub btnsavenewmanufacturer_Click(sender As Object, e As EventArgs) Handles btnSaveNewManufacturer.Click
        Dim enteredName As String = txtAddManufacturer.Text.Trim()
        If String.IsNullOrEmpty(enteredName) Then Exit Sub

        Dim db As New DbConnectionManager()
        Dim existingNames = db.GetManufacturerNames()
        If Not existingNames.AsEnumerable().Any(Function(row) String.Equals(row.Field(Of String)("ManufacturerName"), enteredName, StringComparison.OrdinalIgnoreCase)) Then
            Dim result = MessageBox.Show(
            $"'{enteredName}' is not in the list. Do you want to add it as a new manufacturer?",
            "Add Manufacturer",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )
            If result = DialogResult.Yes Then
                db.InsertManufacturer(enteredName)
                MessageBox.Show($"'{enteredName}' has been added.", "Manufacturer Added", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ' Refresh the manufacturer ComboBox
                cmbManufacturerName.DataSource = Nothing
                cmbManufacturerName.DataSource = db.GetManufacturerNames()
                cmbManufacturerName.DisplayMember = "ManufacturerName"
                cmbManufacturerName.ValueMember = "PK_ManufacturerId"
                ' Optionally select the new manufacturer
                For i As Integer = 0 To cmbManufacturerName.Items.Count - 1
                    Dim drv = TryCast(cmbManufacturerName.Items(i), DataRowView)
                    If drv IsNot Nothing AndAlso String.Equals(drv("ManufacturerName").ToString(), enteredName, StringComparison.OrdinalIgnoreCase) Then
                        cmbManufacturerName.SelectedIndex = i
                        Exit For
                    End If
                Next
                txtAddManufacturer.Clear()
            Else
                txtAddManufacturer.Clear()
            End If
        Else
            MessageBox.Show($"'{enteredName}' already exists.", "Duplicate Manufacturer", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtAddManufacturer.Clear()
        End If
    End Sub

    Private Sub btnFixMissingParentSkus_Click(sender As Object, e As EventArgs) Handles btnFixMissingParentSkus.Click
        Dim db As New DbConnectionManager()
        db.UpdateMissingParentSKUs()
        LoadModelsForSelectedSeries() ' <-- Ensure UI/data is refreshed
        MessageBox.Show("Missing ParentSKUs have been generated and updated.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

#End Region '---------END---------------BUTTON EVENT HANDLERS---------------------------------------------------------------------------------------------------------






    '*****************************************************************************************************************************************************************
#Region "*************START*************SAMPLE PRODUCT LISTING INFORMATION********************************************************************************************"
    '*****************************************************************************************************************************************************************


    Private Sub ShowListingPreview()
        Dim sampleTitle As String
        If dgvModelInformation.Rows.Count = 0 OrElse dgvModelInformation.Rows(0).IsNewRow Then
            sampleTitle = "Sample Product Title: (no models loaded)"
        Else
            Dim manufacturer As String = cmbManufacturerName.Text.Trim()
            Dim series As String = cmbSeriesName.Text.Trim()
            Dim equipmentType As String = cmbEquipmentType.Text.Trim()
            Dim modelName As String = Convert.ToString(dgvModelInformation.Rows(0).Cells("ModelName").Value)
            sampleTitle = $"{manufacturer} {series} {modelName} {equipmentType} Ultimate Custom Cover by GGC"
        End If

        Dim sampleDescription As String = GetSampleProductDescription()
        Dim previewForm As New MpListingsPreview()
        previewForm.SetSampleProductInfo(sampleTitle, sampleDescription)
        previewForm.Show()
    End Sub
    Private Function GetSampleProductDescription() As String
        Dim manufacturer As String = cmbManufacturerName.Text.Trim()
        Dim series As String = cmbSeriesName.Text.Trim()
        Dim modelName As String = ""
        If dgvModelInformation.Rows.Count > 0 AndAlso Not dgvModelInformation.Rows(0).IsNewRow Then
            modelName = Convert.ToString(dgvModelInformation.Rows(0).Cells("ModelName").Value)
        End If
        Dim equipmentType As String = cmbEquipmentType.Text.Trim()

        Dim template As String = "<p>Your <b>[MANUFACTURERNAME] [SERIESNAME] [MODELNAME] [EQUIPMENTTYPE] </b>deserves the best protection, and that's exactly what our covers deliver. Gig Gear Covers hand-craft these covers in the U.S., using FIRST QUALITY materials, including Cordura/Magnatuff or other premium waterproof fabric. We use Denali, Enduratex or similar quality Synthetic Leather.</p>
<p>Elevate your stage presence while keeping your equipment in top-notch condition.</p><ul><li>?? <b>[Enhance your [EQUIPMENTTYPE]]</b> lifespan and performance with the durable and stylish GGC [EQUIPMENTTYPE] Cover.</li>
<li>?? <b>[Premium Materials:] </b>
<p><b>Our 'Choice'  Fabric:</b>  Built to last, our 'Choice' [EQUIPMENTTYPE] covers are made from a high denier Cordura/Magnatuff fabric that is tough, waterproof & extremely tear resistant. </p>
<p>Also, fade resistant it will protect your [EQUIPMENTTYPE] from the Sun ??(UV Rays). In other words, your equipment won't turn Yellow or have its natural color fade if left with the sun shining on it. </p>
<p><b>Polyester excels because: </b>
*It resists pilling & is wrinkle resistant.
*It expels water & dries faster. </p>
<p><b>Premium Synthetic Leather: </b>Our Synthetic Leather is beautiful, water-resistant and oozes class.</p></li>
<li><b>?? [Commitment to Quality & Custom Made!]</b> - Made in the U.S.A.</li>
<p>Sewn w/ industrial sewing machines using Industrial strength materials and thread. This creates strong, durable seams & stitches for covers designed to last!</p>
<li>?? <b>[No Hassle Maintenance!]</b> Just wipe down with a damp cloth and/or non-abrasive cleaning solution.</li></ul>
<p>
Options you can choose from:
<ul>
	<li><b>Padding:</b> Soft, Brushed Black Padding for that awesome feel. It helps even more to lessen damage your amp might get from being hauled around.</li>	
	<li><b>2-in-1 Cargo 'Pick-Pocket':</b> Store your phone, glasses wallet, keys, cords & picks. Magnetic clasp secures the pocket. Pockets opens to approximately 3"" - 4"" deep!</li>
	<li><b>Zippered [EQUIPMENTTYPE] Handle:</b> Protect your amp even more by choosing a Zippered [EQUIPMENTTYPE] Handle Opening.</li>
	<li><b>Premium Synthetic Fabric Colors:</b> Pitch-Black, Pacific-Blue,  Fire-Engine-Red, Orange-Tea, Cloudy Gray, Marbled Brown, Jade-Green, Royal Purple, Magnificent-White</li>
	<li><b>'Choice' Waterproof Fabric Colors:</b> Pitch-Black, Pacific-Blue</li>
</ul>
***********************************************************************
<p><b>HOW TO ORDER A [EQUIPMENTTYPE] COVER & ? [Protect Your Investment]</br> 
**************************************************************</b></p>
<p><b><i>?? See above pictures or description for colors available in your fabric type.</i></b></p>
<p><i>Choose from the colors below and send a message with your order about the color you would like.</i></p>

<p><b>Returns/Refunds:</b></p>
<p>As items are custom-made upon order, we do not accept returns or refunds. But, if there is something wrong with the cover, we ship the wrong one or make it the wrong size, then of course we will take care of you and get you a replacement cover right away.
<p><b>Please do not order a cover if you do not actually have the [MANUFACTURERNAME] [SERIESNAME] [MODELNAME] [EQUIPMENTTYPE] that is named in each listing.</b> If you can't find your model, message us, we can help you.</p>
<p>And finally, we know sometimes mistakes can be made, you order the wrong one, etc . . . While we don't accept returns, we are willing to work with you to get you what you need.</p>
<p><b>Shipping:</b></p>
<p>Standard shipping is via USPS. </p>
<p>All items are custom made and can take approximately 3-4 business days before shipping. </br>Upgrade to expedited shipping and speed up the shipping time and to prioritize your order.</p>
<p><b>Disclaimers:</b></p>
<p>Images ARE NOT INTENDED to represent the product you are buying the dust cover for. Your cover will come for the item you are purchasing in the title of the listing.</p> 
<p>We do not make any guarantee regarding the ability of this cover to protect your equipment/instrument from harm.</p>"

        Return template.Replace("[MANUFACTURERNAME]", manufacturer) _
                       .Replace("[SERIESNAME]", series) _
                       .Replace("[MODELNAME]", modelName) _
                       .Replace("[EQUIPMENTTYPE]", equipmentType)
    End Function



#End Region '-------- -END---------------SAMPLE PRODUCT LISTING INFORMATION------------------------------------------------------------------------------------------------











    'Private Sub dgvSeriesName_MouseDown(sender As Object, e As MouseEventArgs)
    '    isUserSelectingSeries = True
    'End Sub

    'Private Sub dgvSeriesName_KeyDown(sender As Object, e As KeyEventArgs)
    '    isUserSelectingSeries = True
    'End Sub

    'Private Sub UpdateSampleProductTitle()
    '    If dgvModelInformation.Rows.Count = 0 OrElse dgvModelInformation.Rows(0).IsNewRow Then
    '        lblSampleProductTitle.Text = "Sample Product Title: (no models loaded)"
    '        Return
    '    End If

    '    Dim manufacturer As String = cmbManufacturerName.Text.Trim()
    '    Dim series As String = cmbSeriesName.Text.Trim()
    '    Dim equipmentType As String = cmbEquipmentType.Text.Trim()
    '    Dim modelName As String = Convert.ToString(dgvModelInformation.Rows(0).Cells("ModelName").Value)

    '    Dim sampleTitle As String = $"{manufacturer} {series} {modelName} {equipmentType} Ultimate Custom Cover by GGC"
    '    lblSampleProductTitle.Text = "Sample Product Title: " & sampleTitle
    'End Sub

    Private Function ValuesAreEquivalent(val1 As Object, val2 As Object) As Boolean
        ' Treat Nothing, empty string, 0, and False as equivalent if both sides are "empty"
        If IsNothing(val1) AndAlso IsNothing(val2) Then Return True
        If IsNothing(val1) AndAlso (val2 Is "" OrElse val2 Is DBNull.Value) Then Return True
        If IsNothing(val2) AndAlso (val1 Is "" OrElse val1 Is DBNull.Value) Then Return True

        ' Numeric zero equivalence
        If (IsNothing(val1) OrElse val1 Is "" OrElse val1 Is DBNull.Value) AndAlso IsNumeric(val2) AndAlso Convert.ToDecimal(val2) = 0 Then Return True
        If (IsNothing(val2) OrElse val2 Is "" OrElse val2 Is DBNull.Value) AndAlso IsNumeric(val1) AndAlso Convert.ToDecimal(val1) = 0 Then Return True

        ' Boolean false equivalence
        If (IsNothing(val1) OrElse val1 Is "" OrElse val1 Is DBNull.Value) AndAlso TypeOf val2 Is Boolean AndAlso CBool(val2) = False Then Return True
        If (IsNothing(val2) OrElse val2 Is "" OrElse val2 Is DBNull.Value) AndAlso TypeOf val1 Is Boolean AndAlso CBool(val1) = False Then Return True

        ' Otherwise, use normal equality
        Return Object.Equals(val1, val2)
    End Function


    'Private Sub CheckAndAddManufacturer()
    '    Dim enteredName As String = cmbManufacturerName.Text.Trim()
    '    If String.IsNullOrEmpty(enteredName) Then Exit Sub

    '    Dim db As New DbConnectionManager()
    '    Dim existingNames = db.GetManufacturerNames()
    '    If Not existingNames.AsEnumerable().Any(Function(row) String.Equals(row.Field(Of String)("ManufacturerName"), enteredName, StringComparison.OrdinalIgnoreCase)) Then
    '        Dim result = MessageBox.Show(
    '        $"'{enteredName}' is not in the list. Do you want to add it as a new manufacturer?",
    '        "Add Manufacturer",
    '        MessageBoxButtons.YesNo,
    '        MessageBoxIcon.Question
    '    )
    '        If result = DialogResult.Yes Then
    '            db.InsertManufacturer(enteredName)
    '            ' Refresh ComboBox
    '            cmbManufacturerName.DataSource = Nothing
    '            cmbManufacturerName.DataSource = db.GetManufacturerNames()
    '            cmbManufacturerName.SelectedItem = enteredName
    '        End If
    '    End If
    'End Sub



    'Private Sub cmbSeriesName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSeriesName.SelectedIndexChanged
    '    If isLoadingSeries OrElse Not isFormLoaded Then Return

    '    SetEquipmentTypeForSelectedSeries()

    '    ' Get the selected equipment type name
    '    Dim equipmentTypeName As String = cmbEquipmentType.Text.Trim()

    '    ' Initialize columns for the selected equipment type
    '    InitializeModelGridColumns(equipmentTypeName)

    '    LoadModelsForSelectedSeries()
    '    UpdateSampleProductTitle()
    '    lblSampleDescription.Text = GetSampleProductDescription()

    '    If Not String.IsNullOrEmpty(cmbSeriesName.Text.Trim()) Then
    '        dgvModelInformation.Enabled = True
    '    Else
    '        dgvModelInformation.Enabled = False
    '    End If
    'End Sub

    'Private Sub cmbSeriesName_Leave(sender As Object, e As EventArgs) Handles cmbSeriesName.Leave
    '    If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then Exit Sub
    '    Dim selectedManufacturerId As Integer
    '    If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then Exit Sub

    '    Dim enteredSeries As String = cmbSeriesName.Text.Trim()
    '    If String.IsNullOrEmpty(enteredSeries) Then Exit Sub

    '    Dim db As New DbConnectionManager()
    '    Dim existingSeries = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
    '    If Not existingSeries.AsEnumerable().Any(Function(row) String.Equals(row.Field(Of String)("SeriesName"), enteredSeries, StringComparison.OrdinalIgnoreCase)) Then
    '        Dim equipmentType As String = cmbEquipmentType.Text.Trim()
    '        If String.IsNullOrEmpty(equipmentType) Then
    '            MessageBox.Show("Please select or enter an equipment type for this series.", "Equipment Type Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '            cmbEquipmentType.Focus()
    '            Exit Sub
    '        End If

    '        If cmbEquipmentType.SelectedValue Is Nothing OrElse Not IsNumeric(cmbEquipmentType.SelectedValue) Then
    '            MessageBox.Show("Please select a valid equipment type.", "Invalid Equipment Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '            cmbEquipmentType.Focus()
    '            Return
    '        End If

    '        Dim equipmentTypeValue As Integer = Convert.ToInt32(cmbEquipmentType.SelectedValue)
    '        Dim msg As String = $"Do you want to save this new series?" & vbCrLf &
    '                        $"Manufacturer ID: {selectedManufacturerId}" & vbCrLf &
    '                        $"Series: {enteredSeries}" & vbCrLf &
    '                        $"Equipment Type: {equipmentType}"
    '        Dim result = MessageBox.Show(msg, "Save New Series", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
    '        If result = DialogResult.Yes Then
    '            db.InsertSeries(enteredSeries, selectedManufacturerId, equipmentTypeValue)
    '            cmbSeriesName.DataSource = Nothing
    '            cmbSeriesName.DataSource = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
    '            cmbSeriesName.DisplayMember = "SeriesName"
    '            cmbSeriesName.ValueMember = "PK_SeriesId"
    '            cmbSeriesName.SelectedItem = enteredSeries
    '        End If
    '    End If

    '    SetEquipmentTypeForSelectedSeries()
    'End Sub

    'Private Sub SetEquipmentTypeForSelectedSeries()
    '    If isLoadingSeries Then Return
    '    If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then
    '        cmbEquipmentType.SelectedIndex = -1
    '        Return
    '    End If
    '    If cmbSeriesName.SelectedIndex = -1 OrElse cmbSeriesName.SelectedValue Is Nothing Then
    '        cmbEquipmentType.SelectedIndex = -1
    '        Return
    '    End If

    '    Dim selectedManufacturerId As Integer
    '    If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
    '        cmbEquipmentType.SelectedIndex = -1
    '        Return
    '    End If
    '    Dim selectedSeriesId As Integer
    '    If Not Integer.TryParse(cmbSeriesName.SelectedValue.ToString(), selectedSeriesId) Then
    '        cmbEquipmentType.SelectedIndex = -1
    '        Return
    '    End If

    '    Dim db As New DbConnectionManager()
    '    Dim equipmentTypeId = db.GetEquipmentTypeIdByManufacturerAndSeries(selectedManufacturerId, selectedSeriesId)

    '    If equipmentTypeId.HasValue Then
    '        cmbEquipmentType.SelectedValue = equipmentTypeId.Value
    '    Else
    '        cmbEquipmentType.SelectedIndex = -1
    '    End If
    'End Sub

    'Private Sub cmbEquipmentType_Leave(sender As Object, e As EventArgs) Handles cmbEquipmentType.Leave
    '    TrySaveNewSeries()
    'End Sub
    'Private Function CalculateModelMaterialCosts(totalFabricSquareInches As Decimal, Optional wastePercent As Decimal = 5D) As (costPerSqInch_ChoiceWaterproof As Decimal?, costPerSqInch_PremiumSyntheticLeather As Decimal?, costPerSqInch_Padding As Decimal?, baseCost_ChoiceWaterproof As Decimal?, baseCost_PremiumSyntheticLeather As Decimal?, baseCost_ChoiceWaterproof_Padded As Decimal?, baseCost_PremiumSyntheticLeather_Padded As Decimal?, baseCost_PaddingOnly As Decimal?)
    '    Dim db As New DbConnectionManager()
    '    Dim totalSqInchesWithWaste As Decimal = totalFabricSquareInches * (1 + wastePercent / 100D)

    '    Dim costPerSqInch_ChoiceWaterproof = db.GetActiveFabricCostPerSqInch("Choice Waterproof")
    '    Dim costPerSqInch_PremiumSyntheticLeather = db.GetActiveFabricCostPerSqInch("Premium Synthetic Leather")
    '    Dim costPerSqInch_Padding = db.GetActiveFabricCostPerSqInch("Padding")

    '    Dim baseCost_ChoiceWaterproof As Decimal? = If(costPerSqInch_ChoiceWaterproof.HasValue, totalSqInchesWithWaste * costPerSqInch_ChoiceWaterproof.Value, Nothing)
    '    Dim baseCost_PremiumSyntheticLeather As Decimal? = If(costPerSqInch_PremiumSyntheticLeather.HasValue, totalSqInchesWithWaste * costPerSqInch_PremiumSyntheticLeather.Value, Nothing)
    '    Dim baseCost_ChoiceWaterproof_Padded As Decimal? = If(costPerSqInch_ChoiceWaterproof.HasValue AndAlso costPerSqInch_Padding.HasValue, (costPerSqInch_ChoiceWaterproof.Value + costPerSqInch_Padding.Value) * totalSqInchesWithWaste, Nothing)
    '    Dim baseCost_PremiumSyntheticLeather_Padded As Decimal? = If(costPerSqInch_PremiumSyntheticLeather.HasValue AndAlso costPerSqInch_Padding.HasValue, (costPerSqInch_PremiumSyntheticLeather.Value + costPerSqInch_Padding.Value) * totalSqInchesWithWaste, Nothing)
    '    Dim baseCost_PaddingOnly As Decimal? = If(costPerSqInch_Padding.HasValue, costPerSqInch_Padding.Value * totalSqInchesWithWaste, Nothing)

    '    Return (costPerSqInch_ChoiceWaterproof, costPerSqInch_PremiumSyntheticLeather, costPerSqInch_Padding, baseCost_ChoiceWaterproof, baseCost_PremiumSyntheticLeather, baseCost_ChoiceWaterproof_Padded, baseCost_PremiumSyntheticLeather_Padded, baseCost_PaddingOnly)
    'End Function

    'Private Sub LoadModelsForSelectedSeries()
    '    isLoadingModels = True
    '    If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then
    '        dgvModelInformation.DataSource = Nothing
    '        isLoadingModels = False
    '        Return
    '    End If
    '    If cmbSeriesName.SelectedIndex = -1 OrElse cmbSeriesName.SelectedValue Is Nothing Then
    '        dgvModelInformation.DataSource = Nothing
    '        isLoadingModels = False
    '        Return
    '    End If

    '    Dim selectedManufacturerId As Integer
    '    If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
    '        dgvModelInformation.DataSource = Nothing
    '        isLoadingModels = False
    '        Return
    '    End If
    '    Dim selectedSeriesId As Integer
    '    If Not Integer.TryParse(cmbSeriesName.SelectedValue.ToString(), selectedSeriesId) Then
    '        dgvModelInformation.DataSource = Nothing
    '        isLoadingModels = False
    '        Return
    '    End If

    '    Dim db As New DbConnectionManager()
    '    Dim dt = db.GetModelsByManufacturerAndSeries(selectedManufacturerId, selectedSeriesId)
    '    currentModelsTable = dt
    '    dgvModelInformation.DataSource = currentModelsTable
    '    UpdateSampleProductTitle()
    '    isLoadingModels = False
    'End Sub



    ' --- UPDATED: dgvModelInformation_CellContentClick ---
    'Private Sub dgvModelInformation_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvModelInformation.CellContentClick
    '    If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return

    '    dgvModelInformation.CurrentCell = Nothing
    '    Application.DoEvents()

    '    If dgvModelInformation.Columns(e.ColumnIndex).Name = "Save" Then
    '        Dim row = dgvModelInformation.Rows(e.RowIndex)
    '        If row.IsNewRow AndAlso row.Cells.Cast(Of DataGridViewCell).All(Function(c) c.Value Is Nothing OrElse String.IsNullOrWhiteSpace(c.FormattedValue?.ToString())) Then
    '            MessageBox.Show("Please enter data before saving.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '            Return
    '        End If

    '        ' Use IDs, not names
    '        Dim selectedManufacturerId As Integer
    '        If cmbManufacturerName.SelectedValue Is Nothing OrElse Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
    '            MessageBox.Show("Please select a valid manufacturer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Return
    '        End If

    '        Dim selectedSeriesId As Integer
    '        If cmbSeriesName.SelectedValue Is Nothing OrElse Not Integer.TryParse(cmbSeriesName.SelectedValue.ToString(), selectedSeriesId) Then
    '            MessageBox.Show("Please select a valid series.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Return
    '        End If

    '        ' Gather all values (unchanged)
    '        Dim modelName As String = If(dgvModelInformation.Columns.Contains("ModelName"), Convert.ToString(row.Cells("ModelName").Value)?.Trim(), Nothing)
    '        Dim width As Decimal = If(dgvModelInformation.Columns.Contains("Width") AndAlso Not IsDBNull(row.Cells("Width").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Width").Value.ToString()), Convert.ToDecimal(row.Cells("Width").Value), 0D)
    '        Dim depth As Decimal = If(dgvModelInformation.Columns.Contains("Depth") AndAlso Not IsDBNull(row.Cells("Depth").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Depth").Value.ToString()), Convert.ToDecimal(row.Cells("Depth").Value), 0D)
    '        Dim height As Decimal = If(dgvModelInformation.Columns.Contains("Height") AndAlso Not IsDBNull(row.Cells("Height").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Height").Value.ToString()), Convert.ToDecimal(row.Cells("Height").Value), 0D)
    '        Dim optionalHeight As Decimal? = If(dgvModelInformation.Columns.Contains("OptionalHeight") AndAlso Not IsDBNull(row.Cells("OptionalHeight").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("OptionalHeight").Value.ToString()), Convert.ToDecimal(row.Cells("OptionalHeight").Value), Nothing)
    '        Dim optionalDepth As Decimal? = If(dgvModelInformation.Columns.Contains("OptionalDepth") AndAlso Not IsDBNull(row.Cells("OptionalDepth").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("OptionalDepth").Value.ToString()), Convert.ToDecimal(row.Cells("OptionalDepth").Value), Nothing)
    '        Dim angleType As String = If(dgvModelInformation.Columns.Contains("AngleType"), Convert.ToString(row.Cells("AngleType").Value), Nothing)
    '        Dim ampHandleLocation As String = If(dgvModelInformation.Columns.Contains("AmpHandleLocation"), Convert.ToString(row.Cells("AmpHandleLocation").Value), Nothing)
    '        Dim tahWidth As Decimal? = If(dgvModelInformation.Columns.Contains("TAHWidth") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHWidth").Value)), Convert.ToDecimal(row.Cells("TAHWidth").Value), Nothing)
    '        Dim tahHeight As Decimal? = If(dgvModelInformation.Columns.Contains("TAHHeight") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHHeight").Value)), Convert.ToDecimal(row.Cells("TAHHeight").Value), Nothing)
    '        Dim sahWidth As Decimal? = If(dgvModelInformation.Columns.Contains("SAHWidth") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHWidth").Value)), Convert.ToDecimal(row.Cells("SAHWidth").Value), Nothing)
    '        Dim sahHeight As Decimal? = If(dgvModelInformation.Columns.Contains("SAHHeight") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHHeight").Value)), Convert.ToDecimal(row.Cells("SAHHeight").Value), Nothing)
    '        Dim tahRearOffset As Decimal? = If(dgvModelInformation.Columns.Contains("TAHRearOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHRearOffset").Value)), Convert.ToDecimal(row.Cells("TAHRearOffset").Value), Nothing)
    '        Dim sahRearOffset As Decimal? = If(dgvModelInformation.Columns.Contains("SAHRearOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHRearOffset").Value)), Convert.ToDecimal(row.Cells("SAHRearOffset").Value), Nothing)
    '        Dim sahTopDownOffset As Decimal? = If(dgvModelInformation.Columns.Contains("SAHTopDownOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHTopDownOffset").Value)), Convert.ToDecimal(row.Cells("SAHTopDownOffset").Value), Nothing)
    '        Dim musicRestDesign As Boolean? = If(dgvModelInformation.Columns.Contains("MusicRestDesign") AndAlso row.Cells("MusicRestDesign").Value IsNot Nothing, Convert.ToBoolean(row.Cells("MusicRestDesign").Value), Nothing)
    '        Dim chartTemplate As Boolean? = If(dgvModelInformation.Columns.Contains("Chart_Template") AndAlso row.Cells("Chart_Template").Value IsNot Nothing, Convert.ToBoolean(row.Cells("Chart_Template").Value), Nothing)
    '        Dim notes As String = If(dgvModelInformation.Columns.Contains("Notes"), Convert.ToString(row.Cells("Notes").Value), Nothing)

    '        Dim db As New DbConnectionManager()
    '        Dim totalFabricSquareInches As Decimal = ((depth + 1.25D) * (height + 1.25D) * 2D) + ((depth + 1.25D + height + 1.25D + height + 1.25D) * (width + 1.25D))

    '        If dgvModelInformation.Columns.Contains("TotalFabricSquareInches") Then
    '            row.Cells("TotalFabricSquareInches").Value = totalFabricSquareInches
    '        End If

    '        Dim isInsert As Boolean = False
    '        If row.DataBoundItem IsNot Nothing AndAlso TypeOf row.DataBoundItem Is DataRowView Then
    '            Dim drv = DirectCast(row.DataBoundItem, DataRowView)
    '            If drv.Row.RowState = DataRowState.Added Then
    '                isInsert = True
    '            End If
    '        End If

    '        If isInsert Then
    '            If db.ModelExists(selectedManufacturerId, selectedSeriesId, modelName) Then
    '                MessageBox.Show("A model with this name already exists for the selected manufacturer and series.", "Duplicate Model", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '                Return
    '            End If
    '            Dim modelId As Integer = db.InsertModel(selectedManufacturerId, selectedSeriesId, modelName, width, depth, height,
    '            totalFabricSquareInches, optionalHeight, optionalDepth, angleType, ampHandleLocation,
    '            tahWidth, tahHeight, sahWidth, sahHeight,
    '            tahRearOffset, sahRearOffset, sahTopDownOffset,
    '            musicRestDesign, chartTemplate, notes)
    '            SaveModelHistoryCostRetailPricing(modelId, totalFabricSquareInches, notes)
    '            LoadModelsForSelectedSeries()
    '        Else
    '            Dim drv As DataRowView = DirectCast(row.DataBoundItem, DataRowView)
    '            Dim modelId As Integer = CInt(drv("PK_ModelId"))
    '            db.UpdateModel(modelId, modelName, width, depth, height, totalFabricSquareInches,
    '            optionalHeight, optionalDepth, angleType, ampHandleLocation, tahWidth, tahHeight,
    '            sahWidth, sahHeight, tahRearOffset, sahRearOffset, sahTopDownOffset,
    '            musicRestDesign, chartTemplate, notes)
    '            SaveModelHistoryCostRetailPricing(modelId, totalFabricSquareInches, notes)
    '            LoadModelsForSelectedSeries()
    '        End If
    '    End If
    'End Sub
    'Private Sub TrySaveNewSeries()
    '    ' Get the selected manufacturer ID (as Integer)
    '    Dim selectedManufacturerId As Integer
    '    If cmbManufacturerName.SelectedValue Is Nothing OrElse Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
    '        MessageBox.Show("Please select a valid manufacturer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Exit Sub
    '    End If

    '    Dim enteredSeries As String = cmbSeriesName.Text.Trim()
    '    If String.IsNullOrEmpty(enteredSeries) Then Exit Sub

    '    Dim db As New DbConnectionManager()
    '    Dim existingSeries = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
    '    If Not existingSeries.AsEnumerable().Any(Function(row) String.Equals(row.Field(Of String)("SeriesName"), enteredSeries, StringComparison.OrdinalIgnoreCase)) Then
    '        Dim equipmentType As String = cmbEquipmentType.Text.Trim()
    '        If String.IsNullOrEmpty(equipmentType) Then
    '            MessageBox.Show("Please select or enter an equipment type for this series.", "Equipment Type Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '            cmbEquipmentType.Focus()
    '            Exit Sub
    '        End If

    '        If cmbEquipmentType.SelectedValue Is Nothing OrElse Not IsNumeric(cmbEquipmentType.SelectedValue) Then
    '            MessageBox.Show("Please select a valid equipment type.", "Invalid Equipment Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '            cmbEquipmentType.Focus()
    '            Return
    '        End If

    '        Dim equipmentTypeValue As Integer = Convert.ToInt32(cmbEquipmentType.SelectedValue)
    '        Dim msg As String = $"Do you want to save this new series?" & vbCrLf &
    '                    $"Manufacturer ID: {selectedManufacturerId}" & vbCrLf &
    '                    $"Series: {enteredSeries}" & vbCrLf &
    '                    $"Equipment Type: {equipmentType}"
    '        Dim result = MessageBox.Show(msg, "Save New Series", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
    '        If result = DialogResult.Yes Then
    '            db.InsertSeries(enteredSeries, selectedManufacturerId, equipmentTypeValue)
    '            cmbSeriesName.DataSource = Nothing
    '            cmbSeriesName.DataSource = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
    '            cmbSeriesName.DisplayMember = "SeriesName"
    '            cmbSeriesName.ValueMember = "PK_SeriesId"
    '            cmbSeriesName.SelectedItem = enteredSeries
    '        End If
    '    End If

    '    SetEquipmentTypeForSelectedSeries()
    'End Sub
    'Private Sub ImportModelsFromDataTable(dt As DataTable, selectedManufacturer As String, selectedSeries As String)
    '    Dim db As New DbConnectionManager()
    '    Dim insertedCount As Integer = 0

    '    For Each row As DataRow In dt.Rows
    '        If String.IsNullOrWhiteSpace(row("ModelName").ToString()) Then Continue For
    '        Dim modelName = row("ModelName").ToString().Trim()
    '        Dim width = Convert.ToDecimal(row("Width"))
    '        Dim depth = Convert.ToDecimal(row("Depth"))
    '        Dim height = Convert.ToDecimal(row("Height"))

    '        ' Optional fields
    '        Dim optionalHeight = If(dt.Columns.Contains("OptionalHeight") AndAlso Not IsDBNull(row("OptionalHeight")) AndAlso row("OptionalHeight").ToString() <> "", CType(row("OptionalHeight"), Decimal?), Nothing)
    '        Dim optionalDepth = If(dt.Columns.Contains("OptionalDepth") AndAlso Not IsDBNull(row("OptionalDepth")) AndAlso row("OptionalDepth").ToString() <> "", CType(row("OptionalDepth"), Decimal?), Nothing)
    '        Dim angleType = If(dt.Columns.Contains("AngleType"), row("AngleType").ToString(), Nothing)
    '        Dim ampHandleLocation = If(dt.Columns.Contains("AmpHandleLocation"), row("AmpHandleLocation").ToString(), Nothing)
    '        Dim tahWidth = If(dt.Columns.Contains("TAHWidth") AndAlso Not IsDBNull(row("TAHWidth")) AndAlso row("TAHWidth").ToString() <> "", CType(row("TAHWidth"), Decimal?), Nothing)
    '        Dim tahHeight = If(dt.Columns.Contains("TAHHeight") AndAlso Not IsDBNull(row("TAHHeight")) AndAlso row("TAHHeight").ToString() <> "", CType(row("TAHHeight"), Decimal?), Nothing)
    '        Dim sahWidth = If(dt.Columns.Contains("SAHWidth") AndAlso Not IsDBNull(row("SAHWidth")) AndAlso row("SAHWidth").ToString() <> "", CType(row("SAHWidth"), Decimal?), Nothing)
    '        Dim sahHeight = If(dt.Columns.Contains("SAHHeight") AndAlso Not IsDBNull(row("SAHHeight")) AndAlso row("SAHHeight").ToString() <> "", CType(row("SAHHeight"), Decimal?), Nothing)
    '        Dim tahRearOffset = If(dt.Columns.Contains("TAHRearOffset") AndAlso Not IsDBNull(row("TAHRearOffset")) AndAlso row("TAHRearOffset").ToString() <> "", CType(row("TAHRearOffset"), Decimal?), Nothing)
    '        Dim sahRearOffset = If(dt.Columns.Contains("SAHRearOffset") AndAlso Not IsDBNull(row("SAHRearOffset")) AndAlso row("SAHRearOffset").ToString() <> "", CType(row("SAHRearOffset"), Decimal?), Nothing)
    '        Dim sahTopDownOffset = If(dt.Columns.Contains("SAHTopDownOffset") AndAlso Not IsDBNull(row("SAHTopDownOffset")) AndAlso row("SAHTopDownOffset").ToString() <> "", CType(row("SAHTopDownOffset"), Decimal?), Nothing)
    '        Dim musicRestDesign = If(dt.Columns.Contains("MusicRestDesign") AndAlso Not IsDBNull(row("MusicRestDesign")) AndAlso row("MusicRestDesign").ToString() <> "", CBool(row("MusicRestDesign")), CType(Nothing, Boolean?))
    '        Dim chartTemplate = If(dt.Columns.Contains("Chart_Template") AndAlso Not IsDBNull(row("Chart_Template")) AndAlso row("Chart_Template").ToString() <> "", CBool(row("Chart_Template")), CType(Nothing, Boolean?))
    '        Dim notes = If(dt.Columns.Contains("Notes"), row("Notes").ToString(), Nothing)

    '        Dim totalFabricSquareInches As Decimal = ((depth + 1.25D) * (height + 1.25D) * 2D) +
    '                                             ((depth + 1.25D + height + 1.25D + height + 1.25D) * (width + 1.25D))

    '        If db.ModelExists(selectedManufacturer, selectedSeries, modelName) Then Continue For

    '        Try
    '            db.InsertModel(selectedManufacturer, selectedSeries, modelName, width, depth, height,
    '            totalFabricSquareInches, optionalHeight, optionalDepth, angleType, ampHandleLocation,
    '            tahWidth, tahHeight, sahWidth, sahHeight,
    '            tahRearOffset, sahRearOffset, sahTopDownOffset,
    '            musicRestDesign, chartTemplate, notes)
    '            insertedCount += 1
    '        Catch ex As Exception
    '            MessageBox.Show($"Error inserting model '{modelName}': {ex.Message}", "Insert Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End Try
    '    Next

    '    LoadModelsForSelectedSeries()
    '    MessageBox.Show($"{insertedCount} models imported successfully.", "Import Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
    'End Sub
    'Private Sub btnSaveSeries_Click(sender As Object, e As EventArgs) Handles btnSaveSeries.Click
    '    If String.IsNullOrEmpty(cmbManufacturerName.Text.Trim()) OrElse cmbManufacturerName.SelectedIndex = -1 Then
    '        MessageBox.Show("Please select a manufacturer before saving or updating a series.", "Manufacturer Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '        Return
    '    End If
    '    TrySaveOrUpdateSeries()
    '    LoadSeriesGridForManufacturer()
    'End Sub


    'Private Sub LoadSeriesGridForManufacturer()
    '    If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then
    '        dgvSeriesName.DataSource = Nothing
    '        Return
    '    End If

    '    Dim selectedManufacturerId As Integer
    '    If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
    '        dgvSeriesName.DataSource = Nothing
    '        Return
    '    End If

    '    Dim db As New DbConnectionManager()
    '    Dim dt = db.GetSeriesAndEquipmentTypeByManufacturer(selectedManufacturerId)
    '    dgvSeriesName.DataSource = dt
    '    dgvSeriesName.ClearSelection()

    '    ' Hide PK_SeriesId column
    '    If dgvSeriesName.Columns.Contains("PK_SeriesId") Then
    '        dgvSeriesName.Columns("PK_SeriesId").Visible = False
    '    End If
    'End Sub
    'Private Sub dgvSeriesName_SelectionChanged(sender As Object, e As EventArgs) Handles dgvSeriesName.SelectionChanged
    '    If Not isFormLoaded OrElse dgvSeriesName.SelectedRows.Count = 0 Then Return

    '    If Not isUserSelectingSeries Then Return
    '    isUserSelectingSeries = False

    '    Dim selectedRow = dgvSeriesName.SelectedRows(0)
    '    Dim selectedSeries As String = Convert.ToString(selectedRow.Cells("SeriesName").Value)
    '    Dim selectedManufacturer As String = cmbManufacturerName.Text.Trim()

    '    If String.IsNullOrEmpty(selectedManufacturer) OrElse String.IsNullOrEmpty(selectedSeries) Then Return

    '    Dim db As New DbConnectionManager()
    '    Dim dt = db.GetModelsByManufacturerAndSeries(selectedManufacturer, selectedSeries)
    '    If dt.Rows.Count = 0 Then
    '        MessageBox.Show("There are no models associated with this series yet.", "No Models Found", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '        dgvModelInformation.DataSource = Nothing
    '    Else
    '        dgvModelInformation.DataSource = dt
    '    End If
    'End Sub
    '    Private Sub InitializeModelGridColumns(Optional equipmentTypeName As String = Nothing)
    '        dgvModelInformation.Columns.Clear()

    '        ' Always visible columns
    '        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "ModelName", .HeaderText = "Model Name", .DataPropertyName = "ModelName"})
    '        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "Width", .HeaderText = "Width", .DataPropertyName = "Width"})
    '        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "Depth", .HeaderText = "Depth", .DataPropertyName = "Depth"})
    '        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "Height", .HeaderText = "Height", .DataPropertyName = "Height"})

    '        ' Add TotalFabricSquareInches column (read-only, calculated)
    '        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {
    '        .Name = "TotalFabricSquareInches",
    '        .HeaderText = "Total Fabric Sq In",
    '        .DataPropertyName = "TotalFabricSquareInches",
    '        .ReadOnly = True
    '    })
    '        ' Add equipment-type-specific columns
    '        If equipmentTypeName = "Guitar Amplifier" OrElse equipmentTypeName = "Bass Cabinet" Then
    '            AddGuitarAmplifierColumns()
    '        ElseIf equipmentTypeName = "Music Keyboard" Then
    '            AddMusicKeyboardColumns()
    '        End If

    '        ' Always visible for all types
    '        dgvModelInformation.Columns.Add(New DataGridViewCheckBoxColumn With {.Name = "Chart_Template", .HeaderText = "Chart/Template"})
    '        dgvModelInformation.Columns.Add("Notes", "Notes")

    '        Dim saveButtonCol As New DataGridViewButtonColumn With {
    '    .Name = "Save",
    '    .HeaderText = "Save",
    '    .Text = "Save",
    '    .UseColumnTextForButtonValue = True,
    '    .Width = 60
    '}
    '        dgvModelInformation.Columns.Add(saveButtonCol)
    '        FormatModelGridColumns()
    '    End Sub

    '    Private Sub AddGuitarAmplifierColumns()
    '        Dim db As New DbConnectionManager()

    '        ' Optional Height/Depth
    '        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "OptionalHeight", .HeaderText = "Opt. Height", .DataPropertyName = "OptionalHeight"})
    '        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "OptionalDepth", .HeaderText = "Opt. Depth", .DataPropertyName = "OptionalDepth"})

    '        ' AngleType ComboBox
    '        Dim angleTypeCol As New DataGridViewComboBoxColumn With {
    '        .Name = "AngleType",
    '        .HeaderText = "Angle Type",
    '        .DataPropertyName = "AngleType",
    '        .Width = 120,
    '        .DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
    '    }
    '        angleTypeCol.DataSource = db.GetAllAngleTypesTable()
    '        angleTypeCol.DisplayMember = "AngleTypeName"
    '        angleTypeCol.ValueMember = "AngleTypeName"
    '        angleTypeCol.FlatStyle = FlatStyle.Standard
    '        angleTypeCol.ReadOnly = False
    '        dgvModelInformation.Columns.Add(angleTypeCol)

    '        ' TAH/SAH columns
    '        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "TAHWidth", .HeaderText = "TAH Width", .DataPropertyName = "TAHWidth"})
    '        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "TAHHeight", .HeaderText = "TAH Height", .DataPropertyName = "TAHHeight"})
    '        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "SAHWidth", .HeaderText = "SAH Width", .DataPropertyName = "SAHWidth"})
    '        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "SAHHeight", .HeaderText = "SAH Height", .DataPropertyName = "SAHHeight"})

    '        ' AmpHandleLocation ComboBox
    '        Dim ampHandleCol As New DataGridViewComboBoxColumn With {
    '        .Name = "AmpHandleLocation",
    '        .HeaderText = "Amp Handle Location",
    '        .DataPropertyName = "AmpHandleLocation",
    '        .Width = 125,
    '        .DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
    '    }
    '        ampHandleCol.DataSource = db.GetAllAmpHandleLocationsTable()
    '        ampHandleCol.DisplayMember = "AmpHandleLocationName"
    '        ampHandleCol.ValueMember = "AmpHandleLocationName"
    '        ampHandleCol.FlatStyle = FlatStyle.Standard
    '        ampHandleCol.ReadOnly = False
    '        dgvModelInformation.Columns.Add(ampHandleCol)

    '        ' Rear/Offset columns
    '        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "TAHRearOffset", .HeaderText = "TAH Rear Offset", .DataPropertyName = "TAHRearOffset"})
    '        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "SAHRearOffset", .HeaderText = "SAH Rear Offset", .DataPropertyName = "SAHRearOffset"})
    '        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "SAHTopDownOffset", .HeaderText = "SAH TopDown Offset", .DataPropertyName = "SAHTopDownOffset"})
    '    End Sub

    '    Private Sub AddMusicKeyboardColumns()
    '        dgvModelInformation.Columns.Add(New DataGridViewCheckBoxColumn With {.Name = "MusicRestDesign", .HeaderText = "Music Rest Design", .DataPropertyName = "MusicRestDesign"})
    '    End Sub

    '    'Private Sub UpdateModelGridColumnVisibility(equipmentTypeName As String)
    '    '    ' No column-adding code here anymore!
    '    '    ' You can use this for show/hide logic if needed, but not for adding columns.
    '    'End Sub
    '    Private Sub FormatModelGridColumns()
    '        ' Center header text and wrap
    '        For Each col As DataGridViewColumn In dgvModelInformation.Columns
    '            col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
    '            col.HeaderCell.Style.WrapMode = DataGridViewTriState.True
    '        Next

    '        dgvModelInformation.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
    '        dgvModelInformation.ColumnHeadersHeight = 44 ' Adjust as needed

    '        ' Center cell text for specific columns
    '        Dim centerStyle As New DataGridViewCellStyle()
    '        centerStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

    '        For Each colName In New String() {"Width", "Depth", "Height", "OptionalHeight", "OptionalDepth"}
    '            If dgvModelInformation.Columns.Contains(colName) Then
    '                dgvModelInformation.Columns(colName).DefaultCellStyle = centerStyle
    '            End If
    '        Next

    '        ' Set column widths if needed
    '        Dim widths As New Dictionary(Of String, Integer) From {
    '        {"ModelName", 150},
    '        {"Width", 50},
    '        {"Depth", 50},
    '        {"Height", 50},
    '        {"OptionalHeight", 50},
    '        {"OptionalDepth", 50},
    '        {"AngleType", 75},
    '        {"AmpHandleLocation", 125},
    '        {"TAHWidth", 50},
    '        {"TAHHeight", 50},
    '        {"SAHWidth", 50},
    '        {"SAHHeight", 50},
    '        {"TAHRearOffset", 80},
    '        {"SAHRearOffset", 80},
    '        {"SAHTopDownOffset", 90},
    '        {"Chart_Template", 90},
    '        {"Notes", 175}
    '    }
    '        For Each kvp In widths
    '            If dgvModelInformation.Columns.Contains(kvp.Key) Then
    '                dgvModelInformation.Columns(kvp.Key).Width = kvp.Value
    '            End If
    '        Next
    '    End Sub

    ' Handle when a row begins editing to store original values
    'Private Sub dgvModelInformation_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvModelInformation.RowEnter
    '    Dim row = dgvModelInformation.Rows(e.RowIndex)
    '    If row.IsNewRow Then Exit Sub

    '    Dim values As New Dictionary(Of String, Object)
    '    For Each col As DataGridViewColumn In dgvModelInformation.Columns
    '        values(col.Name) = row.Cells(col.Name).Value
    '    Next
    '    originalRowValues(e.RowIndex) = values
    'End Sub

    ' --- UPDATED: dgvModelInformation_RowValidating ---
    'Private Sub dgvModelInformation_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvModelInformation.RowValidating
    '    If isLoadingModels Then Return

    '    Dim row = dgvModelInformation.Rows(e.RowIndex)
    '    If row.IsNewRow AndAlso row.Cells.Cast(Of DataGridViewCell).All(Function(c) c.Value Is Nothing OrElse String.IsNullOrWhiteSpace(c.FormattedValue?.ToString())) Then
    '        Return
    '    End If

    '    ' Use IDs, not names
    '    Dim selectedManufacturerId As Integer
    '    If cmbManufacturerName.SelectedValue Is Nothing OrElse Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
    '        MessageBox.Show("Please select a valid manufacturer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        e.Cancel = True
    '        Return
    '    End If

    '    Dim selectedSeriesId As Integer
    '    If cmbSeriesName.SelectedValue Is Nothing OrElse Not Integer.TryParse(cmbSeriesName.SelectedValue.ToString(), selectedSeriesId) Then
    '        MessageBox.Show("Please select a valid series.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        e.Cancel = True
    '        Return
    '    End If

    '    dgvModelInformation.EndEdit()

    '    ' Gather all values (unchanged)
    '    Dim modelName As String = If(dgvModelInformation.Columns.Contains("ModelName"), Convert.ToString(row.Cells("ModelName").Value)?.Trim(), Nothing)
    '    Dim width As Decimal = If(dgvModelInformation.Columns.Contains("Width") AndAlso Not IsDBNull(row.Cells("Width").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Width").Value.ToString()), Convert.ToDecimal(row.Cells("Width").Value), 0D)
    '    Dim depth As Decimal = If(dgvModelInformation.Columns.Contains("Depth") AndAlso Not IsDBNull(row.Cells("Depth").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Depth").Value.ToString()), Convert.ToDecimal(row.Cells("Depth").Value), 0D)
    '    Dim height As Decimal = If(dgvModelInformation.Columns.Contains("Height") AndAlso Not IsDBNull(row.Cells("Height").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Height").Value.ToString()), Convert.ToDecimal(row.Cells("Height").Value), 0D)
    '    Dim optionalHeight As Decimal? = If(dgvModelInformation.Columns.Contains("OptionalHeight") AndAlso Not IsDBNull(row.Cells("OptionalHeight").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("OptionalHeight").Value.ToString()), Convert.ToDecimal(row.Cells("OptionalHeight").Value), Nothing)
    '    Dim optionalDepth As Decimal? = If(dgvModelInformation.Columns.Contains("OptionalDepth") AndAlso Not IsDBNull(row.Cells("OptionalDepth").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("OptionalDepth").Value.ToString()), Convert.ToDecimal(row.Cells("OptionalDepth").Value), Nothing)
    '    Dim angleType As String = If(dgvModelInformation.Columns.Contains("AngleType"), Convert.ToString(row.Cells("AngleType").Value), Nothing)
    '    Dim ampHandleLocation As String = If(dgvModelInformation.Columns.Contains("AmpHandleLocation"), Convert.ToString(row.Cells("AmpHandleLocation").Value), Nothing)
    '    Dim tahWidth As Decimal? = If(dgvModelInformation.Columns.Contains("TAHWidth") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHWidth").Value)), Convert.ToDecimal(row.Cells("TAHWidth").Value), Nothing)
    '    Dim tahHeight As Decimal? = If(dgvModelInformation.Columns.Contains("TAHHeight") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHHeight").Value)), Convert.ToDecimal(row.Cells("TAHHeight").Value), Nothing)
    '    Dim sahWidth As Decimal? = If(dgvModelInformation.Columns.Contains("SAHWidth") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHWidth").Value)), Convert.ToDecimal(row.Cells("SAHWidth").Value), Nothing)
    '    Dim sahHeight As Decimal? = If(dgvModelInformation.Columns.Contains("SAHHeight") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHHeight").Value)), Convert.ToDecimal(row.Cells("SAHHeight").Value), Nothing)
    '    Dim tahRearOffset As Decimal? = If(dgvModelInformation.Columns.Contains("TAHRearOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHRearOffset").Value)), Convert.ToDecimal(row.Cells("TAHRearOffset").Value), Nothing)
    '    Dim sahRearOffset As Decimal? = If(dgvModelInformation.Columns.Contains("SAHRearOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHRearOffset").Value)), Convert.ToDecimal(row.Cells("SAHRearOffset").Value), Nothing)
    '    Dim sahTopDownOffset As Decimal? = If(dgvModelInformation.Columns.Contains("SAHTopDownOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHTopDownOffset").Value)), Convert.ToDecimal(row.Cells("SAHTopDownOffset").Value), Nothing)
    '    Dim musicRestDesign As Boolean? = If(dgvModelInformation.Columns.Contains("MusicRestDesign") AndAlso row.Cells("MusicRestDesign").Value IsNot Nothing, Convert.ToBoolean(row.Cells("MusicRestDesign").Value), Nothing)
    '    Dim chartTemplate As Boolean? = If(dgvModelInformation.Columns.Contains("Chart_Template") AndAlso row.Cells("Chart_Template").Value IsNot Nothing, Convert.ToBoolean(row.Cells("Chart_Template").Value), Nothing)
    '    Dim notes As String = If(dgvModelInformation.Columns.Contains("Notes"), Convert.ToString(row.Cells("Notes").Value), Nothing)

    '    Dim db As New DbConnectionManager()
    '    Dim totalFabricSquareInches As Decimal = ((depth + 1.25D) * (height + 1.25D) * 2D) + ((depth + 1.25D + height + 1.25D + height + 1.25D) * (width + 1.25D))

    '    ' Update the cell in the grid
    '    If dgvModelInformation.Columns.Contains("TotalFabricSquareInches") Then
    '        row.Cells("TotalFabricSquareInches").Value = totalFabricSquareInches
    '    End If

    '    Dim originalModelName As String = ""
    '    If row.DataBoundItem IsNot Nothing AndAlso TypeOf row.DataBoundItem Is DataRowView Then
    '        originalModelName = Convert.ToString(DirectCast(row.DataBoundItem, DataRowView)("ModelName"))
    '    End If

    '    Dim isInsert As Boolean = String.IsNullOrEmpty(originalModelName)

    '    If isInsert Then
    '        If db.ModelExists(selectedManufacturerId, selectedSeriesId, modelName) Then
    '            MessageBox.Show("A model with this name already exists for the selected manufacturer and series.", "Duplicate Model", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '            e.Cancel = True
    '            Return
    '        End If
    '        Dim modelId As Integer = db.InsertModel(selectedManufacturerId, selectedSeriesId, modelName, width, depth, height,
    '        totalFabricSquareInches, optionalHeight, optionalDepth, angleType, ampHandleLocation,
    '        tahWidth, tahHeight, sahWidth, sahHeight,
    '        tahRearOffset, sahRearOffset, sahTopDownOffset,
    '        musicRestDesign, chartTemplate, notes)
    '        SaveModelHistoryCostRetailPricing(modelId, totalFabricSquareInches, notes)
    '        LoadModelsForSelectedSeries()
    '    Else
    '        If row.DataBoundItem IsNot Nothing AndAlso TypeOf row.DataBoundItem Is DataRowView Then
    '            Dim drv As DataRowView = DirectCast(row.DataBoundItem, DataRowView)
    '            If Not IsDBNull(drv("PK_ModelId")) AndAlso IsNumeric(drv("PK_ModelId")) Then
    '                Dim modelId As Integer = CInt(drv("PK_ModelId"))
    '                db.UpdateModel(modelId, modelName, width, depth, height, totalFabricSquareInches,
    '                optionalHeight, optionalDepth, angleType, ampHandleLocation, tahWidth, tahHeight,
    '                sahWidth, sahHeight, tahRearOffset, sahRearOffset, sahTopDownOffset,
    '                musicRestDesign, chartTemplate, notes)
    '                SaveModelHistoryCostRetailPricing(modelId, totalFabricSquareInches, notes)
    '            End If
    '        End If
    '    End If
    'End Sub

    'Private Sub dgvModelInformation_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvModelInformation.CellValueChanged
    '    If e.RowIndex < 0 Then Exit Sub
    '    Dim row = dgvModelInformation.Rows(e.RowIndex)
    '    If dgvModelInformation.Columns(e.ColumnIndex).Name = "Width" OrElse
    '   dgvModelInformation.Columns(e.ColumnIndex).Name = "Depth" OrElse
    '   dgvModelInformation.Columns(e.ColumnIndex).Name = "Height" Then

    '        Dim width As Decimal = 0, depth As Decimal = 0, height As Decimal = 0
    '        Decimal.TryParse(Convert.ToString(row.Cells("Width").Value), width)
    '        Decimal.TryParse(Convert.ToString(row.Cells("Depth").Value), depth)
    '        Decimal.TryParse(Convert.ToString(row.Cells("Height").Value), height)
    '        Dim totalFabricSquareInches As Decimal = ((depth + 1.25D) * (height + 1.25D) * 2D) +
    '                                             ((depth + 1.25D + height + 1.25D + height + 1.25D) * (width + 1.25D))
    '        row.Cells("TotalFabricSquareInches").Value = totalFabricSquareInches
    '    End If
    'End Sub
    'Private Function ValuesAreEquivalent(val1 As Object, val2 As Object) As Boolean
    '    ' Treat Nothing, empty string, 0, and False as equivalent if both sides are "empty"
    '    If IsNothing(val1) AndAlso IsNothing(val2) Then Return True
    '    If IsNothing(val1) AndAlso (val2 Is "" OrElse val2 Is DBNull.Value) Then Return True
    '    If IsNothing(val2) AndAlso (val1 Is "" OrElse val1 Is DBNull.Value) Then Return True

    '    ' Numeric zero equivalence
    '    If (IsNothing(val1) OrElse val1 Is "" OrElse val1 Is DBNull.Value) AndAlso IsNumeric(val2) AndAlso Convert.ToDecimal(val2) = 0 Then Return True
    '    If (IsNothing(val2) OrElse val2 Is "" OrElse val2 Is DBNull.Value) AndAlso IsNumeric(val1) AndAlso Convert.ToDecimal(val1) = 0 Then Return True

    '    ' Boolean false equivalence
    '    If (IsNothing(val1) OrElse val1 Is "" OrElse val1 Is DBNull.Value) AndAlso TypeOf val2 Is Boolean AndAlso CBool(val2) = False Then Return True
    '    If (IsNothing(val2) OrElse val2 Is "" OrElse val2 Is DBNull.Value) AndAlso TypeOf val1 Is Boolean AndAlso CBool(val1) = False Then Return True

    '    ' Otherwise, use normal equality
    '    Return Object.Equals(val1, val2)
    'End Function

    'Private Sub dgvModelInformation_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvModelInformation.CellEndEdit
    '    dgvModelInformation.CommitEdit(DataGridViewDataErrorContexts.Commit)
    '    dgvModelInformation.EndEdit()
    'End Sub
    'Private Sub dgvModelInformation_Leave(sender As Object, e As EventArgs) Handles dgvModelInformation.Leave
    '    ' If the current row is the new row and has data, force commit and validation
    '    If dgvModelInformation.CurrentRow IsNot Nothing AndAlso dgvModelInformation.CurrentRow.IsNewRow Then
    '        Dim hasData = dgvModelInformation.CurrentRow.Cells.Cast(Of DataGridViewCell) _
    '        .Any(Function(c) c.Value IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(c.FormattedValue?.ToString()))
    '        If hasData Then
    '            ' Force the row to be validated and committed
    '            dgvModelInformation.CurrentRow.Selected = True
    '            dgvModelInformation.EndEdit()
    '            dgvModelInformation.CurrentCell = Nothing
    '            ' Explicitly call Validate to ensure RowValidating fires
    '            Me.Validate()
    '        End If
    '    ElseIf dgvModelInformation.IsCurrentRowDirty Then
    '        dgvModelInformation.EndEdit()
    '    End If
    'End Sub
    'Private Sub SaveModelHistoryCostRetailPricing(modelId As Integer, totalFabricSquareInches As Decimal, notes As String)
    '    Dim wastePercent As Decimal = 5D
    '    Dim costs = CalculateModelMaterialCosts(totalFabricSquareInches, wastePercent)
    '    Dim weights = CalculateModelMaterialWeights(totalFabricSquareInches)
    '    Dim db As New DbConnectionManager()

    '    ' Shipping
    '    Dim shipping_Choice As Decimal = db.GetShippingCostByWeight(If(weights.weight_ChoiceWaterproof, 0D))
    '    Dim shipping_ChoicePadded As Decimal = db.GetShippingCostByWeight(If(weights.weight_ChoiceWaterproof_Padded, 0D))
    '    Dim shipping_Leather As Decimal = db.GetShippingCostByWeight(If(weights.weight_PremiumSyntheticLeather, 0D))
    '    Dim shipping_LeatherPadded As Decimal = db.GetShippingCostByWeight(If(weights.weight_PremiumSyntheticLeather_Padded, 0D))

    '    ' Base fabric cost + shipping
    '    Dim baseFabricCost_Choice_Weight As Decimal? = If(costs.baseCost_ChoiceWaterproof, 0D) + shipping_Choice
    '    Dim baseFabricCost_ChoicePadding_Weight As Decimal? = If(costs.baseCost_ChoiceWaterproof_Padded, 0D) + shipping_ChoicePadded
    '    Dim baseFabricCost_Leather_Weight As Decimal? = If(costs.baseCost_PremiumSyntheticLeather, 0D) + shipping_Leather
    '    Dim baseFabricCost_LeatherPadding_Weight As Decimal? = If(costs.baseCost_PremiumSyntheticLeather_Padded, 0D) + shipping_LeatherPadded

    '    ' Labor
    '    Dim hourlyRate As Decimal = 17D
    '    Dim CostLaborNoPadding As Decimal = 0.5D * hourlyRate
    '    Dim CostLaborWithPadding As Decimal = 1D * hourlyRate

    '    Dim BaseCost_Choice_Labor As Decimal? = If(baseFabricCost_Choice_Weight, 0D) + CostLaborNoPadding
    '    Dim BaseCost_ChoicePadding_Labor As Decimal? = If(baseFabricCost_ChoicePadding_Weight, 0D) + CostLaborWithPadding
    '    Dim BaseCost_Leather_Labor As Decimal? = If(baseFabricCost_Leather_Weight, 0D) + CostLaborNoPadding
    '    Dim BaseCost_LeatherPadding_Labor As Decimal? = If(baseFabricCost_LeatherPadding_Weight, 0D) + CostLaborWithPadding

    '    ' Profits
    '    Dim profit_Choice As Decimal = db.GetLatestProfitValue("Choice")
    '    Dim profit_ChoicePadded As Decimal = db.GetLatestProfitValue("ChoicePadded")
    '    Dim profit_Leather As Decimal = db.GetLatestProfitValue("Leather")
    '    Dim profit_LeatherPadded As Decimal = db.GetLatestProfitValue("LeatherPadded")

    '    ' Marketplace fees
    '    Dim amazonFeePct As Decimal = db.GetMarketplaceFeePercentage("Amazon")
    '    Dim reverbFeePct As Decimal = db.GetMarketplaceFeePercentage("Reverb")
    '    Dim ebayFeePct As Decimal = db.GetMarketplaceFeePercentage("eBay")
    '    Dim etsyFeePct As Decimal = db.GetMarketplaceFeePercentage("Etsy")

    '    ' Retail (for fee calculation only)
    '    Dim retail_Choice = BaseCost_Choice_Labor + profit_Choice
    '    Dim retail_ChoicePadded = BaseCost_ChoicePadding_Labor + profit_ChoicePadded
    '    Dim retail_Leather = BaseCost_Leather_Labor + profit_Leather
    '    Dim retail_LeatherPadded = BaseCost_LeatherPadding_Labor + profit_LeatherPadded

    '    ' Marketplace fees
    '    Dim AmazonFee_Choice = retail_Choice * (amazonFeePct / 100D)
    '    Dim AmazonFee_ChoicePadded = retail_ChoicePadded * (amazonFeePct / 100D)
    '    Dim AmazonFee_Leather = retail_Leather * (amazonFeePct / 100D)
    '    Dim AmazonFee_LeatherPadded = retail_LeatherPadded * (amazonFeePct / 100D)

    '    Dim ReverbFee_Choice = retail_Choice * (reverbFeePct / 100D)
    '    Dim ReverbFee_ChoicePadded = retail_ChoicePadded * (reverbFeePct / 100D)
    '    Dim ReverbFee_Leather = retail_Leather * (reverbFeePct / 100D)
    '    Dim ReverbFee_LeatherPadded = retail_LeatherPadded * (reverbFeePct / 100D)

    '    Dim eBayFee_Choice = retail_Choice * (ebayFeePct / 100D)
    '    Dim eBayFee_ChoicePadded = retail_ChoicePadded * (ebayFeePct / 100D)
    '    Dim eBayFee_Leather = retail_Leather * (ebayFeePct / 100D)
    '    Dim eBayFee_LeatherPadded = retail_LeatherPadded * (ebayFeePct / 100D)

    '    Dim etsyFee_Choice = retail_Choice * (etsyFeePct / 100D)
    '    Dim etsyFee_ChoicePadded = retail_ChoicePadded * (etsyFeePct / 100D)
    '    Dim etsyFee_Leather = retail_Leather * (etsyFeePct / 100D)
    '    Dim etsyFee_LeatherPadded = retail_LeatherPadded * (etsyFeePct / 100D)

    '    ' Grand totals (actual cost + marketplace fee, profit NOT included)
    '    Dim BaseCost_GrandTotal_Choice_Amazon = BaseCost_Choice_Labor + AmazonFee_Choice
    '    Dim BaseCost_GrandTotal_ChoicePadded_Amazon = BaseCost_ChoicePadding_Labor + AmazonFee_ChoicePadded
    '    Dim BaseCost_GrandTotal_Leather_Amazon = BaseCost_Leather_Labor + AmazonFee_Leather
    '    Dim BaseCost_GrandTotal_LeatherPadded_Amazon = BaseCost_LeatherPadding_Labor + AmazonFee_LeatherPadded

    '    Dim BaseCost_GrandTotal_Choice_Reverb = BaseCost_Choice_Labor + ReverbFee_Choice
    '    Dim BaseCost_GrandTotal_ChoicePadded_Reverb = BaseCost_ChoicePadding_Labor + ReverbFee_ChoicePadded
    '    Dim BaseCost_GrandTotal_Leather_Reverb = BaseCost_Leather_Labor + ReverbFee_Leather
    '    Dim BaseCost_GrandTotal_LeatherPadded_Reverb = BaseCost_LeatherPadding_Labor + ReverbFee_LeatherPadded

    '    Dim BaseCost_GrandTotal_Choice_eBay = BaseCost_Choice_Labor + eBayFee_Choice
    '    Dim BaseCost_GrandTotal_ChoicePadded_eBay = BaseCost_ChoicePadding_Labor + eBayFee_ChoicePadded
    '    Dim BaseCost_GrandTotal_Leather_eBay = BaseCost_Leather_Labor + eBayFee_Leather
    '    Dim BaseCost_GrandTotal_LeatherPadded_eBay = BaseCost_LeatherPadding_Labor + eBayFee_LeatherPadded

    '    Dim BaseCost_GrandTotal_Choice_Etsy = BaseCost_Choice_Labor + etsyFee_Choice
    '    Dim BaseCost_GrandTotal_ChoicePadded_Etsy = BaseCost_ChoicePadding_Labor + etsyFee_ChoicePadded
    '    Dim BaseCost_GrandTotal_Leather_Etsy = BaseCost_Leather_Labor + etsyFee_Leather
    '    Dim BaseCost_GrandTotal_LeatherPadded_Etsy = BaseCost_LeatherPadding_Labor + etsyFee_LeatherPadded

    '    ' === Calculate Retail Prices ===
    '    Dim RetailPrice_Choice_Amazon = CalculateRetailPrice(BaseCost_GrandTotal_Choice_Amazon, profit_Choice)
    '    Dim RetailPrice_ChoicePadded_Amazon = CalculateRetailPrice(BaseCost_GrandTotal_ChoicePadded_Amazon, profit_ChoicePadded)
    '    Dim RetailPrice_Leather_Amazon = CalculateRetailPrice(BaseCost_GrandTotal_Leather_Amazon, profit_Leather)
    '    Dim RetailPrice_LeatherPadded_Amazon = CalculateRetailPrice(BaseCost_GrandTotal_LeatherPadded_Amazon, profit_LeatherPadded)

    '    Dim RetailPrice_Choice_Reverb = CalculateRetailPrice(BaseCost_GrandTotal_Choice_Reverb, profit_Choice)
    '    Dim RetailPrice_ChoicePadded_Reverb = CalculateRetailPrice(BaseCost_GrandTotal_ChoicePadded_Reverb, profit_ChoicePadded)
    '    Dim RetailPrice_Leather_Reverb = CalculateRetailPrice(BaseCost_GrandTotal_Leather_Reverb, profit_Leather)
    '    Dim RetailPrice_LeatherPadded_Reverb = CalculateRetailPrice(BaseCost_GrandTotal_LeatherPadded_Reverb, profit_LeatherPadded)

    '    Dim RetailPrice_Choice_eBay = CalculateRetailPrice(BaseCost_GrandTotal_Choice_eBay, profit_Choice)
    '    Dim RetailPrice_ChoicePadded_eBay = CalculateRetailPrice(BaseCost_GrandTotal_ChoicePadded_eBay, profit_ChoicePadded)
    '    Dim RetailPrice_Leather_eBay = CalculateRetailPrice(BaseCost_GrandTotal_Leather_eBay, profit_Leather)
    '    Dim RetailPrice_LeatherPadded_eBay = CalculateRetailPrice(BaseCost_GrandTotal_LeatherPadded_eBay, profit_LeatherPadded)

    '    Dim RetailPrice_Choice_Etsy = CalculateRetailPrice(BaseCost_GrandTotal_Choice_Etsy, profit_Choice)
    '    Dim RetailPrice_ChoicePadded_Etsy = CalculateRetailPrice(BaseCost_GrandTotal_ChoicePadded_Etsy, profit_ChoicePadded)
    '    Dim RetailPrice_Leather_Etsy = CalculateRetailPrice(BaseCost_GrandTotal_Leather_Etsy, profit_Leather)
    '    Dim RetailPrice_LeatherPadded_Etsy = CalculateRetailPrice(BaseCost_GrandTotal_LeatherPadded_Etsy, profit_LeatherPadded)



    '    Dim latestRow = db.GetLatestModelHistoryCostRetailPricing(modelId)
    '    If latestRow IsNot Nothing Then
    '        If ModelHistoryRowsAreEquivalent(latestRow, totalFabricSquareInches, wastePercent, costs, weights,
    '        baseFabricCost_Choice_Weight, baseFabricCost_ChoicePadding_Weight, baseFabricCost_Leather_Weight,
    '        baseFabricCost_LeatherPadding_Weight, BaseCost_Choice_Labor, BaseCost_ChoicePadding_Labor,
    '        BaseCost_Leather_Labor, BaseCost_LeatherPadding_Labor, profit_Choice, profit_ChoicePadded,
    '        profit_Leather, profit_LeatherPadded, AmazonFee_Choice, AmazonFee_ChoicePadded, AmazonFee_Leather,
    '        AmazonFee_LeatherPadded, ReverbFee_Choice, ReverbFee_ChoicePadded, ReverbFee_Leather,
    '        ReverbFee_LeatherPadded, eBayFee_Choice, eBayFee_ChoicePadded, eBayFee_Leather, eBayFee_LeatherPadded,
    '        etsyFee_Choice, etsyFee_ChoicePadded, etsyFee_Leather, etsyFee_LeatherPadded,
    '        BaseCost_GrandTotal_Choice_Amazon, BaseCost_GrandTotal_ChoicePadded_Amazon,
    '        BaseCost_GrandTotal_Leather_Amazon, BaseCost_GrandTotal_LeatherPadded_Amazon,
    '        BaseCost_GrandTotal_Choice_Reverb, BaseCost_GrandTotal_ChoicePadded_Reverb,
    '        BaseCost_GrandTotal_Leather_Reverb, BaseCost_GrandTotal_LeatherPadded_Reverb,
    '        BaseCost_GrandTotal_Choice_eBay, BaseCost_GrandTotal_ChoicePadded_eBay,
    '        BaseCost_GrandTotal_Leather_eBay, BaseCost_GrandTotal_LeatherPadded_eBay,
    '        BaseCost_GrandTotal_Choice_Etsy, BaseCost_GrandTotal_ChoicePadded_Etsy,
    '        BaseCost_GrandTotal_Leather_Etsy, BaseCost_GrandTotal_LeatherPadded_Etsy,
    '        RetailPrice_Choice_Amazon, RetailPrice_ChoicePadded_Amazon, RetailPrice_Leather_Amazon,
    '        RetailPrice_LeatherPadded_Amazon, RetailPrice_Choice_Reverb, RetailPrice_ChoicePadded_Reverb,
    '        RetailPrice_Leather_Reverb, RetailPrice_LeatherPadded_Reverb, RetailPrice_Choice_eBay,
    '        RetailPrice_ChoicePadded_eBay, RetailPrice_Leather_eBay, RetailPrice_LeatherPadded_eBay,
    '        RetailPrice_Choice_Etsy, RetailPrice_ChoicePadded_Etsy, RetailPrice_Leather_Etsy,
    '        RetailPrice_LeatherPadded_Etsy, notes) Then
    '            ' No changes, do not insert
    '            Exit Sub
    '        End If
    '    End If


    '    db.InsertModelHistoryCostRetailPricing(
    '    modelId,
    '    costs.costPerSqInch_ChoiceWaterproof,
    '    costs.costPerSqInch_PremiumSyntheticLeather,
    '    costs.costPerSqInch_Padding,
    '    totalFabricSquareInches,
    '    wastePercent,
    '    costs.baseCost_ChoiceWaterproof,
    '    costs.baseCost_PremiumSyntheticLeather,
    '    costs.baseCost_ChoiceWaterproof_Padded,
    '    costs.baseCost_PremiumSyntheticLeather_Padded,
    '    costs.baseCost_PaddingOnly,
    '    weights.weight_PaddingOnly,
    '    weights.weight_ChoiceWaterproof,
    '    weights.weight_ChoiceWaterproof_Padded,
    '    weights.weight_PremiumSyntheticLeather,
    '    weights.weight_PremiumSyntheticLeather_Padded,
    '    shipping_Choice,
    '    shipping_ChoicePadded,
    '    shipping_Leather,
    '    shipping_LeatherPadded,
    '    baseFabricCost_Choice_Weight,
    '    baseFabricCost_ChoicePadding_Weight,
    '    baseFabricCost_Leather_Weight,
    '    baseFabricCost_LeatherPadding_Weight,
    '    BaseCost_Choice_Labor,
    '    BaseCost_ChoicePadding_Labor,
    '    BaseCost_Leather_Labor,
    '    BaseCost_LeatherPadding_Labor,
    '    profit_Choice,
    '    profit_ChoicePadded,
    '    profit_Leather,
    '    profit_LeatherPadded,
    '    AmazonFee_Choice,
    '    AmazonFee_ChoicePadded,
    '    AmazonFee_Leather,
    '    AmazonFee_LeatherPadded,
    '    ReverbFee_Choice,
    '    ReverbFee_ChoicePadded,
    '    ReverbFee_Leather,
    '    ReverbFee_LeatherPadded,
    '    eBayFee_Choice,
    '    eBayFee_ChoicePadded,
    '    eBayFee_Leather,
    '    eBayFee_LeatherPadded,
    '    etsyFee_Choice,
    '    etsyFee_ChoicePadded,
    '    etsyFee_Leather,
    '    etsyFee_LeatherPadded,
    '    BaseCost_GrandTotal_Choice_Amazon,
    '    BaseCost_GrandTotal_ChoicePadded_Amazon,
    '    BaseCost_GrandTotal_Leather_Amazon,
    '    BaseCost_GrandTotal_LeatherPadded_Amazon,
    '    BaseCost_GrandTotal_Choice_Reverb,
    '    BaseCost_GrandTotal_ChoicePadded_Reverb,
    '    BaseCost_GrandTotal_Leather_Reverb,
    '    BaseCost_GrandTotal_LeatherPadded_Reverb,
    '    BaseCost_GrandTotal_Choice_eBay,
    '    BaseCost_GrandTotal_ChoicePadded_eBay,
    '    BaseCost_GrandTotal_Leather_eBay,
    '    BaseCost_GrandTotal_LeatherPadded_eBay,
    '    BaseCost_GrandTotal_Choice_Etsy,
    '    BaseCost_GrandTotal_ChoicePadded_Etsy,
    '    BaseCost_GrandTotal_Leather_Etsy,
    '    BaseCost_GrandTotal_LeatherPadded_Etsy,
    '    RetailPrice_Choice_Amazon,
    '    RetailPrice_ChoicePadded_Amazon,
    '    RetailPrice_Leather_Amazon,
    '    RetailPrice_LeatherPadded_Amazon,
    '    RetailPrice_Choice_Reverb,
    '    RetailPrice_ChoicePadded_Reverb,
    '    RetailPrice_Leather_Reverb,
    '    RetailPrice_LeatherPadded_Reverb,
    '    RetailPrice_Choice_eBay,
    '    RetailPrice_ChoicePadded_eBay,
    '    RetailPrice_Leather_eBay,
    '    RetailPrice_LeatherPadded_eBay,
    '    RetailPrice_Choice_Etsy,
    '    RetailPrice_ChoicePadded_Etsy,
    '    RetailPrice_Leather_Etsy,
    '    RetailPrice_LeatherPadded_Etsy,
    '    notes
    ')
    'End Sub



    'Private Sub btnUploadModels_Click(sender As Object, e As EventArgs) Handles btnUploadModels.Click
    '    Using ofd As New OpenFileDialog()
    '        ofd.Filter = "Excel or CSV Files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|All Files (*.*)|*.*"
    '        ofd.Title = "Select Model Information File"
    '        If ofd.ShowDialog() <> DialogResult.OK Then Return

    '        Dim filePath = ofd.FileName
    '        Dim dt As New DataTable()

    '        If Path.GetExtension(filePath).ToLower() = ".csv" Then
    '            ' --- CSV Handling ---
    '            Dim lines = File.ReadAllLines(filePath)
    '            If lines.Length < 2 Then
    '                MessageBox.Show("The selected file does not contain enough data.", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '                Return
    '            End If
    '            Dim headers = lines(0).Split(","c).Select(Function(h) h.Trim()).ToArray()
    '            For Each header In headers
    '                dt.Columns.Add(header)
    '            Next
    '            For i As Integer = 1 To lines.Length - 1
    '                Dim values = lines(i).Split(","c).Select(Function(v) v.Trim()).ToArray()
    '                If values.Length = headers.Length Then
    '                    dt.Rows.Add(values)
    '                End If
    '            Next
    '        ElseIf Path.GetExtension(filePath).ToLower() = ".xlsx" OrElse Path.GetExtension(filePath).ToLower() = ".xls" Then
    '            ' --- Excel Handling ---
    '            Dim connStr As String
    '            If Path.GetExtension(filePath).ToLower() = ".xlsx" Then
    '                connStr = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES';"
    '            Else
    '                connStr = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath};Extended Properties='Excel 8.0;HDR=YES';"
    '            End If
    '            Try
    '                Using conn As New OleDbConnection(connStr)
    '                    conn.Open()
    '                    Dim dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
    '                    Dim sheetName = dtSchema.Rows(0)("TABLE_NAME").ToString()
    '                    Using cmd As New OleDbCommand($"SELECT * FROM [{sheetName}]", conn)
    '                        Using da As New OleDbDataAdapter(cmd)
    '                            da.Fill(dt)
    '                        End Using
    '                    End Using
    '                End Using
    '            Catch ex As Exception
    '                MessageBox.Show("Error reading Excel file: " & ex.Message, "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                Return
    '            End Try
    '        Else
    '            MessageBox.Show("Unsupported file type.", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Return
    '        End If

    '        ' --- Validate required columns ---
    '        Dim requiredCols = {"ManufacturerName", "SeriesName", "ModelName", "Width", "Depth", "Height"}
    '        For Each col In requiredCols
    '            If Not dt.Columns.Contains(col) Then
    '                MessageBox.Show($"Missing required column: {col}", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                Return
    '            End If
    '        Next

    '        Dim db As New DbConnectionManager()
    '        Dim insertedCount As Integer = 0
    '        For Each row As DataRow In dt.Rows
    '            Dim manufacturerName = row("ManufacturerName").ToString().Trim()
    '            Dim seriesName = row("SeriesName").ToString().Trim()
    '            Dim modelName = row("ModelName").ToString().Trim()
    '            Dim widthStr = row("Width").ToString().Trim()
    '            Dim depthStr = row("Depth").ToString().Trim()
    '            Dim heightStr = row("Height").ToString().Trim()

    '            ' Skip rows missing any required value
    '            If String.IsNullOrEmpty(manufacturerName) OrElse String.IsNullOrEmpty(seriesName) OrElse
    '           String.IsNullOrEmpty(modelName) OrElse String.IsNullOrEmpty(widthStr) OrElse
    '           String.IsNullOrEmpty(depthStr) OrElse String.IsNullOrEmpty(heightStr) Then
    '                Continue For
    '            End If

    '            Dim width As Decimal, depth As Decimal, height As Decimal
    '            If Not Decimal.TryParse(widthStr, width) OrElse Not Decimal.TryParse(depthStr, depth) OrElse Not Decimal.TryParse(heightStr, height) Then
    '                Continue For
    '            End If

    '            ' --- Lookup ManufacturerId ---
    '            Dim manufacturerId As Integer = db.GetManufacturerIdByName(manufacturerName)
    '            If manufacturerId = 0 Then
    '                manufacturerId = db.InsertManufacturer(manufacturerName)
    '            End If

    '            ' --- Lookup SeriesId ---
    '            Dim seriesId As Integer = db.GetSeriesIdByNameAndManufacturer(seriesName, manufacturerId)
    '            If seriesId = 0 Then
    '                ' You may need to provide a default EquipmentTypeId if not present in the file
    '                Dim defaultEquipmentTypeId As Integer = 1 ' Change as needed
    '                seriesId = db.InsertSeries(seriesName, manufacturerId, defaultEquipmentTypeId)
    '            End If

    '            ' Optional fields (safe parsing)
    '            Dim optionalHeight As Decimal? = Nothing
    '            If dt.Columns.Contains("OptionalHeight") AndAlso Not IsDBNull(row("OptionalHeight")) Then
    '                Dim val = row("OptionalHeight").ToString().Trim()
    '                If val <> "" Then
    '                    Dim parsed As Decimal
    '                    If Decimal.TryParse(val, parsed) Then optionalHeight = parsed
    '                End If
    '            End If

    '            Dim optionalDepth As Decimal? = Nothing
    '            If dt.Columns.Contains("OptionalDepth") AndAlso Not IsDBNull(row("OptionalDepth")) Then
    '                Dim val = row("OptionalDepth").ToString().Trim()
    '                If val <> "" Then
    '                    Dim parsed As Decimal
    '                    If Decimal.TryParse(val, parsed) Then optionalDepth = parsed
    '                End If
    '            End If

    '            Dim angleType = If(dt.Columns.Contains("AngleType"), row("AngleType").ToString(), Nothing)
    '            Dim ampHandleLocation = If(dt.Columns.Contains("AmpHandleLocation"), row("AmpHandleLocation").ToString(), Nothing)

    '            Dim tahWidth As Decimal? = Nothing
    '            If dt.Columns.Contains("TAHWidth") AndAlso Not IsDBNull(row("TAHWidth")) Then
    '                Dim val = row("TAHWidth").ToString().Trim()
    '                If val <> "" Then
    '                    Dim parsed As Decimal
    '                    If Decimal.TryParse(val, parsed) Then tahWidth = parsed
    '                End If
    '            End If

    '            Dim tahHeight As Decimal? = Nothing
    '            If dt.Columns.Contains("TAHHeight") AndAlso Not IsDBNull(row("TAHHeight")) Then
    '                Dim val = row("TAHHeight").ToString().Trim()
    '                If val <> "" Then
    '                    Dim parsed As Decimal
    '                    If Decimal.TryParse(val, parsed) Then tahHeight = parsed
    '                End If
    '            End If

    '            Dim sahWidth As Decimal? = Nothing
    '            If dt.Columns.Contains("SAHWidth") AndAlso Not IsDBNull(row("SAHWidth")) Then
    '                Dim val = row("SAHWidth").ToString().Trim()
    '                If val <> "" Then
    '                    Dim parsed As Decimal
    '                    If Decimal.TryParse(val, parsed) Then sahWidth = parsed
    '                End If
    '            End If

    '            Dim sahHeight As Decimal? = Nothing
    '            If dt.Columns.Contains("SAHHeight") AndAlso Not IsDBNull(row("SAHHeight")) Then
    '                Dim val = row("SAHHeight").ToString().Trim()
    '                If val <> "" Then
    '                    Dim parsed As Decimal
    '                    If Decimal.TryParse(val, parsed) Then sahHeight = parsed
    '                End If
    '            End If

    '            Dim tahRearOffset As Decimal? = Nothing
    '            If dt.Columns.Contains("TAHRearOffset") AndAlso Not IsDBNull(row("TAHRearOffset")) Then
    '                Dim val = row("TAHRearOffset").ToString().Trim()
    '                If val <> "" Then
    '                    Dim parsed As Decimal
    '                    If Decimal.TryParse(val, parsed) Then tahRearOffset = parsed
    '                End If
    '            End If

    '            Dim sahRearOffset As Decimal? = Nothing
    '            If dt.Columns.Contains("SAHRearOffset") AndAlso Not IsDBNull(row("SAHRearOffset")) Then
    '                Dim val = row("SAHRearOffset").ToString().Trim()
    '                If val <> "" Then
    '                    Dim parsed As Decimal
    '                    If Decimal.TryParse(val, parsed) Then sahRearOffset = parsed
    '                End If
    '            End If

    '            Dim sahTopDownOffset As Decimal? = Nothing
    '            If dt.Columns.Contains("SAHTopDownOffset") AndAlso Not IsDBNull(row("SAHTopDownOffset")) Then
    '                Dim val = row("SAHTopDownOffset").ToString().Trim()
    '                If val <> "" Then
    '                    Dim parsed As Decimal
    '                    If Decimal.TryParse(val, parsed) Then sahTopDownOffset = parsed
    '                End If
    '            End If

    '            Dim musicRestDesign As Boolean? = Nothing
    '            If dt.Columns.Contains("MusicRestDesign") AndAlso Not IsDBNull(row("MusicRestDesign")) Then
    '                Dim val = row("MusicRestDesign").ToString().Trim()
    '                If val <> "" Then
    '                    Dim parsed As Boolean
    '                    If Boolean.TryParse(val, parsed) Then musicRestDesign = parsed
    '                End If
    '            End If

    '            Dim chartTemplate As Boolean? = Nothing
    '            If dt.Columns.Contains("Chart_Template") AndAlso Not IsDBNull(row("Chart_Template")) Then
    '                Dim val = row("Chart_Template").ToString().Trim()
    '                If val <> "" Then
    '                    Dim parsed As Boolean
    '                    If Boolean.TryParse(val, parsed) Then chartTemplate = parsed
    '                End If
    '            End If

    '            Dim notes = If(dt.Columns.Contains("Notes"), row("Notes").ToString(), Nothing)

    '            Dim totalFabricSquareInches As Decimal = ((depth + 1.25D) * (height + 1.25D) * 2D) +
    '                                                 ((depth + 1.25D + height + 1.25D + height + 1.25D) * (width + 1.25D))

    '            If db.ModelExists(manufacturerId, seriesId, modelName) Then Continue For

    '            Try
    '                db.InsertModel(manufacturerId, seriesId, modelName, width, depth, height,
    '                totalFabricSquareInches, optionalHeight, optionalDepth, angleType, ampHandleLocation,
    '                tahWidth, tahHeight, sahWidth, sahHeight,
    '                tahRearOffset, sahRearOffset, sahTopDownOffset,
    '                musicRestDesign, chartTemplate, notes)
    '                insertedCount += 1
    '            Catch ex As Exception
    '                MessageBox.Show($"Error inserting model '{modelName}': {ex.Message}", "Insert Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            End Try
    '        Next

    '        LoadModelsForSelectedSeries()
    '        MessageBox.Show($"{insertedCount} models imported successfully.", "Import Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '    End Using
    'End Sub

    'Private Function ModelHistoryRowsAreEquivalent(latestRow As DataRow,
    'totalFabricSquareInches As Decimal, wastePercent As Decimal,
    'costs As Object, weights As Object, baseFabricCost_Choice_Weight As Decimal?,
    'baseFabricCost_ChoicePadding_Weight As Decimal?, baseFabricCost_Leather_Weight As Decimal?,
    'baseFabricCost_LeatherPadding_Weight As Decimal?, BaseCost_Choice_Labor As Decimal?,
    'BaseCost_ChoicePadding_Labor As Decimal?, BaseCost_Leather_Labor As Decimal?,
    'BaseCost_LeatherPadding_Labor As Decimal?, profit_Choice As Decimal, profit_ChoicePadded As Decimal,
    'profit_Leather As Decimal, profit_LeatherPadded As Decimal, AmazonFee_Choice As Decimal,
    'AmazonFee_ChoicePadded As Decimal, AmazonFee_Leather As Decimal, AmazonFee_LeatherPadded As Decimal,
    'ReverbFee_Choice As Decimal, ReverbFee_ChoicePadded As Decimal, ReverbFee_Leather As Decimal,
    'ReverbFee_LeatherPadded As Decimal, eBayFee_Choice As Decimal, eBayFee_ChoicePadded As Decimal,
    'eBayFee_Leather As Decimal, eBayFee_LeatherPadded As Decimal, etsyFee_Choice As Decimal,
    'etsyFee_ChoicePadded As Decimal, etsyFee_Leather As Decimal, etsyFee_LeatherPadded As Decimal,
    'BaseCost_GrandTotal_Choice_Amazon As Decimal, BaseCost_GrandTotal_ChoicePadded_Amazon As Decimal,
    'BaseCost_GrandTotal_Leather_Amazon As Decimal, BaseCost_GrandTotal_LeatherPadded_Amazon As Decimal,
    'BaseCost_GrandTotal_Choice_Reverb As Decimal, BaseCost_GrandTotal_ChoicePadded_Reverb As Decimal,
    'BaseCost_GrandTotal_Leather_Reverb As Decimal, BaseCost_GrandTotal_LeatherPadded_Reverb As Decimal,
    'BaseCost_GrandTotal_Choice_eBay As Decimal, BaseCost_GrandTotal_ChoicePadded_eBay As Decimal,
    'BaseCost_GrandTotal_Leather_eBay As Decimal, BaseCost_GrandTotal_LeatherPadded_eBay As Decimal,
    'BaseCost_GrandTotal_Choice_Etsy As Decimal, BaseCost_GrandTotal_ChoicePadded_Etsy As Decimal,
    'BaseCost_GrandTotal_Leather_Etsy As Decimal, BaseCost_GrandTotal_LeatherPadded_Etsy As Decimal,
    'RetailPrice_Choice_Amazon As Decimal, RetailPrice_ChoicePadded_Amazon As Decimal,
    'RetailPrice_Leather_Amazon As Decimal, RetailPrice_LeatherPadded_Amazon As Decimal,
    'RetailPrice_Choice_Reverb As Decimal, RetailPrice_ChoicePadded_Reverb As Decimal,
    'RetailPrice_Leather_Reverb As Decimal, RetailPrice_LeatherPadded_Reverb As Decimal,
    'RetailPrice_Choice_eBay As Decimal, RetailPrice_ChoicePadded_eBay As Decimal,
    'RetailPrice_Leather_eBay As Decimal, RetailPrice_LeatherPadded_eBay As Decimal,
    'RetailPrice_Choice_Etsy As Decimal, RetailPrice_ChoicePadded_Etsy As Decimal,
    'RetailPrice_Leather_Etsy As Decimal, RetailPrice_LeatherPadded_Etsy As Decimal,
    'notes As String) As Boolean

    '    ' Compare all relevant fields, return False if any are different
    '    If Not ValuesAreEquivalent(latestRow("TotalFabricSquareInches"), totalFabricSquareInches) Then Return False
    '    If Not ValuesAreEquivalent(latestRow("WastePercent"), wastePercent) Then Return False
    '    If Not ValuesAreEquivalent(latestRow("CostPerSqInch_ChoiceWaterproof"), costs.costPerSqInch_ChoiceWaterproof) Then Return False
    '    If Not ValuesAreEquivalent(latestRow("CostPerSqInch_PremiumSyntheticLeather"), costs.costPerSqInch_PremiumSyntheticLeather) Then Return False
    '    If Not ValuesAreEquivalent(latestRow("CostPerSqInch_Padding"), costs.costPerSqInch_Padding) Then Return False
    '    ' ...repeat for all other fields you want to compare...
    '    ' For brevity, only a few are shown. You should compare all fields that are calculated and stored.

    '    ' Example for notes:
    '    If Not ValuesAreEquivalent(latestRow("Notes"), notes) Then Return False

    '    ' If all fields are equivalent
    '    Return True
    'End Function
    'Private Sub TrySaveOrUpdateSeries()
    '    ' Get the selected manufacturer ID (as Integer)
    '    Dim selectedManufacturerId As Integer
    '    If cmbManufacturerName.SelectedValue Is Nothing OrElse Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
    '        MessageBox.Show("Please select a valid manufacturer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Exit Sub
    '    End If

    '    Dim enteredSeries As String = cmbSeriesName.Text.Trim()
    '    If String.IsNullOrEmpty(enteredSeries) Then Exit Sub

    '    Dim db As New DbConnectionManager()
    '    Dim existingSeries = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
    '    If Not existingSeries.AsEnumerable().Any(Function(row) String.Equals(row.Field(Of String)("SeriesName"), enteredSeries, StringComparison.OrdinalIgnoreCase)) Then
    '        ' Require equipment type selection
    '        If cmbEquipmentType.SelectedValue Is Nothing OrElse Not IsNumeric(cmbEquipmentType.SelectedValue) Then
    '            MessageBox.Show("Please select a valid equipment type.", "Invalid Equipment Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '            cmbEquipmentType.Focus()
    '            Return
    '        End If

    '        Dim equipmentTypeValue As Integer = Convert.ToInt32(cmbEquipmentType.SelectedValue)
    '        Dim msg As String = $"Do you want to save this new series?" & vbCrLf &
    '                        $"Manufacturer ID: {selectedManufacturerId}" & vbCrLf &
    '                        $"Series: {enteredSeries}" & vbCrLf &
    '                        $"Equipment Type: {cmbEquipmentType.Text.Trim()}"
    '        Dim result = MessageBox.Show(msg, "Save New Series", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
    '        If result = DialogResult.Yes Then
    '            db.InsertSeries(enteredSeries, selectedManufacturerId, equipmentTypeValue)
    '            cmbSeriesName.DataSource = Nothing
    '            cmbSeriesName.DataSource = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
    '            cmbSeriesName.DisplayMember = "SeriesName"
    '            cmbSeriesName.ValueMember = "PK_SeriesId"
    '            cmbSeriesName.SelectedItem = enteredSeries
    '        End If
    '    Else
    '        ' Update equipment type for existing series
    '        If cmbEquipmentType.SelectedValue Is Nothing OrElse Not IsNumeric(cmbEquipmentType.SelectedValue) Then
    '            MessageBox.Show("Please select a valid equipment type.", "Invalid Equipment Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '            cmbEquipmentType.Focus()
    '            Return
    '        End If
    '        Dim equipmentTypeValue As Integer = Convert.ToInt32(cmbEquipmentType.SelectedValue)
    '        Dim seriesId As Integer = CInt(existingSeries.AsEnumerable().First(Function(row) String.Equals(row.Field(Of String)("SeriesName"), enteredSeries, StringComparison.OrdinalIgnoreCase))("PK_SeriesId"))
    '        db.UpdateSeriesEquipmentType(seriesId, equipmentTypeValue)
    '        MessageBox.Show("Series equipment type updated.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '    End If

    '    SetEquipmentTypeForSelectedSeries()
    'End Sub

    'Private Sub dgvSeriesName_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSeriesName.CellContentClick

    'End Sub

    'Private Sub btn_UpdateCosts_Click(sender As Object, e As EventArgs) Handles btn_UpdateCosts.Click
    '    Dim db As New DbConnectionManager()
    '    Dim dtModels = db.GetAllModelsWithDetails()
    '    Dim results As New List(Of ModelCostUpdateResults)()

    '    For Each row As DataRow In dtModels.Rows
    '        Dim modelId As Integer = CInt(row("PK_ModelId"))
    '        Dim modelName As String = row("ModelName").ToString()
    '        Dim seriesName As String = row("SeriesName").ToString()
    '        Dim manufacturerName As String = row("ManufacturerName").ToString()
    '        Dim totalFabricSquareInches As Decimal = If(IsDBNull(row("TotalFabricSquareInches")), 0D, Convert.ToDecimal(row("TotalFabricSquareInches")))
    '        Dim notes As String = row("Notes").ToString()

    '        If totalFabricSquareInches > 0 Then
    '            Dim dbMgr As New DbConnectionManager()
    '            Dim latestRow = dbMgr.GetLatestModelHistoryCostRetailPricing(modelId)
    '            SaveModelHistoryCostRetailPricing(modelId, totalFabricSquareInches, notes)
    '            Dim newRow = dbMgr.GetLatestModelHistoryCostRetailPricing(modelId)
    '            Dim updated As Boolean = (latestRow Is Nothing) OrElse (Not ValuesAreEquivalent(latestRow("DateCalculated"), newRow("DateCalculated")))
    '            Dim msg As String = If(updated, "Updated", "No change needed")
    '            results.Add(New ModelCostUpdateResults(modelId, modelName, seriesName, manufacturerName, updated, msg))
    '        Else
    '            results.Add(New ModelCostUpdateResults(modelId, modelName, seriesName, manufacturerName, False, "Skipped (no size)"))
    '        End If
    '    Next

    '    Dim updatedCount = results.Where(Function(r) r.Updated).Count()
    '    Dim skippedCount = results.Count - updatedCount

    '    MessageBox.Show(
    '    $"Cost update complete. {updatedCount} models updated, {skippedCount} skipped." & vbCrLf &
    '    String.Join(vbCrLf, results.Select(Function(r) $"{r.ManufacturerName} - {r.SeriesName} - {r.ModelName}: {r.Message}")),
    '    "Update Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
    'End Sub

    'Private Sub btnsavenewmanufacturer_Click(sender As Object, e As EventArgs) Handles btnSaveNewManufacturer.Click
    '    Dim enteredName As String = txtAddManufacturer.Text.Trim()
    '    If String.IsNullOrEmpty(enteredName) Then Exit Sub

    '    Dim db As New DbConnectionManager()
    '    Dim existingNames = db.GetManufacturerNames()
    '    If Not existingNames.AsEnumerable().Any(Function(row) String.Equals(row.Field(Of String)("ManufacturerName"), enteredName, StringComparison.OrdinalIgnoreCase)) Then
    '        Dim result = MessageBox.Show(
    '        $"'{enteredName}' is not in the list. Do you want to add it as a new manufacturer?",
    '        "Add Manufacturer",
    '        MessageBoxButtons.YesNo,
    '        MessageBoxIcon.Question
    '    )
    '        If result = DialogResult.Yes Then
    '            db.InsertManufacturer(enteredName)
    '            MessageBox.Show($"'{enteredName}' has been added.", "Manufacturer Added", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '            ' Refresh the manufacturer ComboBox
    '            cmbManufacturerName.DataSource = Nothing
    '            cmbManufacturerName.DataSource = db.GetManufacturerNames()
    '            cmbManufacturerName.DisplayMember = "ManufacturerName"
    '            cmbManufacturerName.ValueMember = "PK_ManufacturerId"
    '            ' Optionally select the new manufacturer
    '            For i As Integer = 0 To cmbManufacturerName.Items.Count - 1
    '                Dim drv = TryCast(cmbManufacturerName.Items(i), DataRowView)
    '                If drv IsNot Nothing AndAlso String.Equals(drv("ManufacturerName").ToString(), enteredName, StringComparison.OrdinalIgnoreCase) Then
    '                    cmbManufacturerName.SelectedIndex = i
    '                    Exit For
    '                End If
    '            Next
    '            txtAddManufacturer.Clear()
    '        Else
    '            txtAddManufacturer.Clear()
    '        End If
    '    Else
    '        MessageBox.Show($"'{enteredName}' already exists.", "Duplicate Manufacturer", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '        txtAddManufacturer.Clear()
    '    End If
    'End Sub

    'Private Async Sub btnUploadWooListings_Click(sender As Object, e As EventArgs) Handles btnUploadWooListings.Click
    '    If currentModelsTable Is Nothing OrElse currentModelsTable.Rows.Count = 0 Then
    '        MessageBox.Show("No models to upload.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '        Return
    '    End If

    '    Dim manufacturer As String = cmbManufacturerName.Text.Trim()
    '    Dim series As String = cmbSeriesName.Text.Trim()
    '    Dim equipmentType As String = cmbEquipmentType.Text.Trim()

    '    Dim successCount As Integer = 0
    '    Dim failCount As Integer = 0

    '    For Each dr As DataRow In currentModelsTable.Rows
    '        Dim modelName As String = dr.Field(Of String)("ModelName")
    '        If String.IsNullOrWhiteSpace(modelName) Then Continue For

    '        Dim productTitle As String = $"{manufacturer} {series} {modelName} {equipmentType} Ultimate Custom Cover by GGC"

    '        Dim product As New MpWooCommerceProduct With {
    '        .name = productTitle,
    '        .type = "variable",
    '        .status = "publish",
    '        .attributes = New List(Of MpWooAttribute) From {
    '            New MpWooAttribute With {
    '                .name = "Fabric",
    '                .variation = True,
    '                .options = New List(Of String) From {"Choice Waterproof", "Premium Synthetic Leather"}
    '            },
    '            New MpWooAttribute With {
    '                .name = "Color",
    '                .variation = True,
    '                .options = New List(Of String) From {"Pitch-Black", "Fire-Engine-Red", "Pacific-Blue", "Magnificent-White", "Royal-Purple", "Marbled-Brown", "Jade-Green"}
    '            }
    '        }
    '    }

    '        Dim parentProductJson As String = System.Text.Json.JsonSerializer.Serialize(product)
    '        Try
    '            Await WooCommerceApi.UploadProductAsync(parentProductJson)
    '            successCount += 1
    '        Catch ex As Exception
    '            failCount += 1
    '        End Try
    '    Next

    '    MessageBox.Show($"Upload complete! {successCount} products uploaded, {failCount} failed.", "WooCommerce Upload", MessageBoxButtons.OK, MessageBoxIcon.Information)
    'End Sub

    '    Private Function GetSampleProductDescription() As String
    '        Dim manufacturer As String = cmbManufacturerName.Text.Trim()
    '        Dim series As String = cmbSeriesName.Text.Trim()
    '        Dim modelName As String = ""
    '        If dgvModelInformation.Rows.Count > 0 AndAlso Not dgvModelInformation.Rows(0).IsNewRow Then
    '            modelName = Convert.ToString(dgvModelInformation.Rows(0).Cells("ModelName").Value)
    '        End If
    '        Dim equipmentType As String = cmbEquipmentType.Text.Trim()

    '        Dim template As String = "<p>Your <b>[MANUFACTURERNAME] [SERIESNAME] [MODELNAME] [EQUIPMENTTYPE] </b>deserves the best protection, and that's exactly what our covers deliver. Gig Gear Covers hand-craft these covers in the U.S., using FIRST QUALITY materials, including Cordura/Magnatuff or other premium waterproof fabric. We use Denali, Enduratex or similar quality Synthetic Leather.</p>
    '<p>Elevate your stage presence while keeping your equipment in top-notch condition.</p><ul><li>?? <b>[Enhance your [EQUIPMENTTYPE]]</b> lifespan and performance with the durable and stylish GGC [EQUIPMENTTYPE] Cover.</li>
    '<li>?? <b>[Premium Materials:] </b>
    '<p><b>Our 'Choice'  Fabric:</b>  Built to last, our 'Choice' [EQUIPMENTTYPE] covers are made from a high denier Cordura/Magnatuff fabric that is tough, waterproof & extremely tear resistant. </p>
    '<p>Also, fade resistant it will protect your [EQUIPMENTTYPE] from the Sun ??(UV Rays). In other words, your equipment won't turn Yellow or have its natural color fade if left with the sun shining on it. </p>
    '<p><b>Polyester excels because: </b>
    '*It resists pilling & is wrinkle resistant.
    '*It expels water & dries faster. </p>
    '<p><b>Premium Synthetic Leather: </b>Our Synthetic Leather is beautiful, water-resistant and oozes class.</p></li>
    '<li><b>?? [Commitment to Quality & Custom Made!]</b> - Made in the U.S.A.</li>
    '<p>Sewn w/ industrial sewing machines using Industrial strength materials and thread. This creates strong, durable seams & stitches for covers designed to last!</p>
    '<li>?? <b>[No Hassle Maintenance!]</b> Just wipe down with a damp cloth and/or non-abrasive cleaning solution.</li></ul>
    '<p>
    'Options you can choose from:
    '<ul>
    '	<li><b>Padding:</b> Soft, Brushed Black Padding for that awesome feel. It helps even more to lessen damage your amp might get from being hauled around.</li>	
    '	<li><b>2-in-1 Cargo 'Pick-Pocket':</b> Store your phone, glasses wallet, keys, cords & picks. Magnetic clasp secures the pocket. Pockets opens to approximately 3"" - 4"" deep!</li>
    '	<li><b>Zippered [EQUIPMENTTYPE] Handle:</b> Protect your amp even more by choosing a Zippered [EQUIPMENTTYPE] Handle Opening.</li>
    '	<li><b>Premium Synthetic Fabric Colors:</b> Pitch-Black, Pacific-Blue,  Fire-Engine-Red, Orange-Tea, Cloudy Gray, Marbled Brown, Jade-Green, Royal Purple, Magnificent-White</li>
    '	<li><b>'Choice' Waterproof Fabric Colors:</b> Pitch-Black, Pacific-Blue</li>
    '</ul>
    '***********************************************************************
    '<p><b>HOW TO ORDER A [EQUIPMENTTYPE] COVER & ? [Protect Your Investment]</br> 
    '**************************************************************</b></p>
    '<p><b><i>?? See above pictures or description for colors available in your fabric type.</i></b></p>
    '<p><i>Choose from the colors below and send a message with your order about the color you would like.</i></p>

    '<p><b>Returns/Refunds:</b></p>
    '<p>As items are custom-made upon order, we do not accept returns or refunds. But, if there is something wrong with the cover, we ship the wrong one or make it the wrong size, then of course we will take care of you and get you a replacement cover right away.
    '<p><b>Please do not order a cover if you do not actually have the [MANUFACTURERNAME] [SERIESNAME] [MODELNAME] [EQUIPMENTTYPE] that is named in each listing.</b> If you can't find your model, message us, we can help you.</p>
    '<p>And finally, we know sometimes mistakes can be made, you order the wrong one, etc . . . While we don't accept returns, we are willing to work with you to get you what you need.</p>
    '<p><b>Shipping:</b></p>
    '<p>Standard shipping is via USPS. </p>
    '<p>All items are custom made and can take approximately 3-4 business days before shipping. </br>Upgrade to expedited shipping and speed up the shipping time and to prioritize your order.</p>
    '<p><b>Disclaimers:</b></p>
    '<p>Images ARE NOT INTENDED to represent the product you are buying the dust cover for. Your cover will come for the item you are purchasing in the title of the listing.</p> 
    '<p>We do not make any guarantee regarding the ability of this cover to protect your equipment/instrument from harm.</p>"

    '        Return template.Replace("[MANUFACTURERNAME]", manufacturer) _
    '                       .Replace("[SERIESNAME]", series) _
    '                       .Replace("[MODELNAME]", modelName) _
    '                       .Replace("[EQUIPMENTTYPE]", equipmentType)
    '    End Function
End Class