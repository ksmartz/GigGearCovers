Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Threading.Tasks
Imports System.Windows.Forms


Public Class frmAddModelInformation

    Private isFormLoaded As Boolean = False
    Private _isKeyboardLayout As Boolean = False

    ' ========= utilities =========


    Private Function FindTextBox(name As String) As TextBox
        Dim arr = Me.Controls.Find(name, True)
        If arr Is Nothing OrElse arr.Length = 0 Then Return Nothing
        Return TryCast(arr(0), TextBox)
    End Function

    Private Function GetCellValue(Of T)(row As DataGridViewRow, colName As String) As T
        Dim v = row.Cells(colName).Value
        If v Is Nothing OrElse v Is DBNull.Value Then Return Nothing
        Return DirectCast(Convert.ChangeType(v, GetType(T)), T)
    End Function

    Private Function GetCellDecimal(row As DataGridViewRow, colName As String) As Decimal?
        Dim v = row.Cells(colName).Value
        If v Is Nothing OrElse v Is DBNull.Value Then Return Nothing
        Dim d As Decimal
        If Decimal.TryParse(v.ToString(), Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture, d) Then
            Return d
        End If
        If Decimal.TryParse(v.ToString(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d) Then
            Return d
        End If
        Return Nothing
    End Function

    Private Function FirstExistingColumn(dt As DataTable, ParamArray candidates() As String) As String
        For Each c In candidates
            If dt.Columns.Contains(c) Then Return c
        Next
        ' if none found, just return the first candidate (will throw at runtime if truly absent)
        Return candidates(0)
    End Function

    Private Sub BindComboSmart(cmb As ComboBox, dt As DataTable, displayCandidates() As String, valueCandidates() As String)
        If cmb Is Nothing Then Return
        cmb.BeginUpdate()
        cmb.DataSource = Nothing

        If dt Is Nothing OrElse dt.Columns.Count = 0 Then
            cmb.DisplayMember = Nothing
            cmb.ValueMember = Nothing
            cmb.EndUpdate()
            Exit Sub
        End If

        Dim disp = FirstExistingColumn(dt, displayCandidates)
        Dim val = FirstExistingColumn(dt, valueCandidates)

        cmb.DisplayMember = disp
        cmb.ValueMember = val
        cmb.DataSource = dt
        cmb.SelectedIndex = -1
        cmb.EndUpdate()
    End Sub



    Private Function ParseDecimal(tb As TextBox, Optional fallback As Decimal = 0D) As Decimal
        If tb Is Nothing Then Return fallback
        Dim d As Decimal
        If Decimal.TryParse(tb.Text, d) Then Return d
        Return fallback
    End Function

    ' ========= load =========
    Private Sub frmAddModelInformation_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        isFormLoaded = False

        BindManufacturers()
        BindEquipmentTypes()
        BindOptionalSupplierArea()

        WireEvents()

        isFormLoaded = True
    End Sub
    ' =========================
    ' EVENT WIRING (no helpers)
    ' =========================
    Private Sub WireEvents()
        Dim arr() As Control
        Dim cmbManufacturer As ComboBox = Nothing
        Dim cmbSeries As ComboBox = Nothing
        Dim cmbSupplier As ComboBox = Nothing
        Dim cmbBrand As ComboBox = Nothing
        Dim btnUpload As Button = Nothing
        Dim btnSave As Button = Nothing
        Dim btnUpdate As Button = Nothing

        arr = Me.Controls.Find("cmbManufacturerName", True)
        If arr IsNot Nothing AndAlso arr.Length > 0 Then cmbManufacturer = TryCast(arr(0), ComboBox)

        arr = Me.Controls.Find("cmbSeriesName", True)
        If arr IsNot Nothing AndAlso arr.Length > 0 Then cmbSeries = TryCast(arr(0), ComboBox)

        arr = Me.Controls.Find("cmbSupplier", True)
        If arr IsNot Nothing AndAlso arr.Length > 0 Then cmbSupplier = TryCast(arr(0), ComboBox)

        arr = Me.Controls.Find("cmbBrand", True)
        If arr IsNot Nothing AndAlso arr.Length > 0 Then cmbBrand = TryCast(arr(0), ComboBox)

        arr = Me.Controls.Find("btnUploadWooListings", True)
        If arr IsNot Nothing AndAlso arr.Length > 0 Then btnUpload = TryCast(arr(0), Button)

        arr = Me.Controls.Find("btnSavePricing", True)
        If arr IsNot Nothing AndAlso arr.Length > 0 Then btnSave = TryCast(arr(0), Button)

        arr = Me.Controls.Find("btnUpdateProductInfo", True)
        If arr IsNot Nothing AndAlso arr.Length > 0 Then btnUpdate = TryCast(arr(0), Button)

        If cmbManufacturer IsNot Nothing Then
            RemoveHandler cmbManufacturer.SelectedIndexChanged, AddressOf cmbManufacturerName_SelectedIndexChanged
            AddHandler cmbManufacturer.SelectedIndexChanged, AddressOf cmbManufacturerName_SelectedIndexChanged
        End If

        If cmbSeries IsNot Nothing Then
            RemoveHandler cmbSeries.SelectedIndexChanged, AddressOf cmbSeriesName_SelectedIndexChanged
            AddHandler cmbSeries.SelectedIndexChanged, AddressOf cmbSeriesName_SelectedIndexChanged

            RemoveHandler cmbSeries.SelectionChangeCommitted, AddressOf cmbSeriesName_SelectionChangeCommitted
            AddHandler cmbSeries.SelectionChangeCommitted, AddressOf cmbSeriesName_SelectionChangeCommitted
        End If

        If cmbSupplier IsNot Nothing Then
            RemoveHandler cmbSupplier.SelectedIndexChanged, AddressOf cmbSupplier_SelectedIndexChanged
            AddHandler cmbSupplier.SelectedIndexChanged, AddressOf cmbSupplier_SelectedIndexChanged
        End If

        If cmbBrand IsNot Nothing Then
            RemoveHandler cmbBrand.SelectedIndexChanged, AddressOf cmbBrand_SelectedIndexChanged
            AddHandler cmbBrand.SelectedIndexChanged, AddressOf cmbBrand_SelectedIndexChanged
        End If

        If btnUpload IsNot Nothing Then
            RemoveHandler btnUpload.Click, AddressOf btnUploadWooListings_Click
            AddHandler btnUpload.Click, AddressOf btnUploadWooListings_Click
        End If

        If btnSave IsNot Nothing Then
            RemoveHandler btnSave.Click, AddressOf btnSavePricing_Click
            AddHandler btnSave.Click, AddressOf btnSavePricing_Click
        End If

        If btnUpdate IsNot Nothing Then
            RemoveHandler btnUpdate.Click, AddressOf btnUpdateProductInfo_Click
            AddHandler btnUpdate.Click, AddressOf btnUpdateProductInfo_Click
        End If
    End Sub





    Private Sub ConfigureModelGridColumns(equipmentTypeName As String)
        Dim isKeyboard As Boolean =
        Not String.IsNullOrWhiteSpace(equipmentTypeName) AndAlso
        equipmentTypeName.Trim().ToLowerInvariant().Contains("keyboard")
        _isKeyboardLayout = isKeyboard

        Dim g = dgvModelInformation
        g.RowHeadersVisible = False
        g.AutoGenerateColumns = False
        g.AllowUserToAddRows = False
        g.EditMode = DataGridViewEditMode.EditOnEnter
        g.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        g.Columns.Clear()

        ' Hidden PK
        g.Columns.Add(New DataGridViewTextBoxColumn With {
        .Name = "PK_ModelId",
        .HeaderText = "Id",
        .DataPropertyName = "PK_ModelId",
        .Visible = False
    })

        ' ModelName BEFORE ParentSku (read-only, wide)
        Dim colModel = MakeTextCol("ModelName", "Model", fill:=True, isReadOnly:=True)
        colModel.MinimumWidth = 180
        colModel.FillWeight = 220
        g.Columns.Add(colModel)

        ' Parent SKU
        Dim colParent = MakeTextCol("ParentSku", "Parent SKU")
        colParent.MinimumWidth = 140
        colParent.FillWeight = 160
        g.Columns.Add(colParent)

        ' Common dims
        g.Columns.Add(MakeTextCol("Width", "Width"))
        g.Columns.Add(MakeTextCol("Depth", "Depth"))
        g.Columns.Add(MakeTextCol("Height", "Height"))
        g.Columns.Add(MakeTextCol("TotalFabricSquareInches", "Total Sq In"))

        If isKeyboard Then
            g.Columns.Add(MakeTextCol("MusicRestDesign", "Music Rest Design"))
            g.Columns.Add(MakeTextCol("Chart_Template", "Chart Template"))
            g.Columns.Add(MakeTextCol("Notes", "Notes"))
            g.Columns.Add(MakeTextCol("WooProductId", "Woo Product Id", isReadOnly:=True))
        Else
            ' Amp/Cab layout includes combo for AmpHandleLocation
            g.Columns.Add(MakeAmpHandleComboCol())
            g.Columns.Add(MakeTextCol("TAHWidth", "TAH Width"))
            g.Columns.Add(MakeTextCol("TAHHeight", "TAH Height"))
            g.Columns.Add(MakeTextCol("TAHRearOffset", "TAH Rear Offset"))
            g.Columns.Add(MakeTextCol("SAHHeight", "SAH Height"))
            g.Columns.Add(MakeTextCol("SAHWidth", "SAH Width"))
            g.Columns.Add(MakeTextCol("SAHRearOffset", "SAH Rear Offset"))
            g.Columns.Add(MakeTextCol("SAHTopDownOffset", "SAH TopDown Offset"))
            g.Columns.Add(MakeTextCol("Chart_Template", "Chart Template"))
            g.Columns.Add(MakeTextCol("Notes", "Notes"))
            g.Columns.Add(MakeTextCol("WooProductId", "Woo Product Id", isReadOnly:=True))
        End If

        ' Last column = Save button
        g.Columns.Add(MakeSaveButtonCol())
    End Sub




    Private Function MakeTextCol(prop As String,
                             header As String,
                             Optional fill As Boolean = False,
                             Optional isReadOnly As Boolean = False) As DataGridViewTextBoxColumn
        Dim c As New DataGridViewTextBoxColumn With {
        .Name = prop,
        .HeaderText = header,
        .DataPropertyName = prop,
        .ReadOnly = isReadOnly
    }
        If fill Then c.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Return c
    End Function

    ' frmAddModelInformation.vb
    Private Sub LoadModelsGrid(mid As Integer, sid As Integer, equipmentTypeName As String)
        Dim dt = DbConnectionManager.GetModelsForGrid(mid, sid)

        Dim g = dgvModelInformation
        g.AutoGenerateColumns = False
        g.RowHeadersVisible = False           ' hides the fixed row-header strip with the arrow
        g.Columns.Clear()

        ' Hidden PK
        g.Columns.Add(New DataGridViewTextBoxColumn With {
        .Name = "PK_ModelId", .HeaderText = "Id", .DataPropertyName = "PK_ModelId", .Visible = False
    })

        ' ModelName BEFORE ParentSku
        Dim colModel = MakeTextCol("ModelName", "Model", fill:=True, isReadOnly:=True)
        colModel.MinimumWidth = 180
        colModel.FillWeight = 200
        g.Columns.Add(colModel)

        g.Columns.Add(MakeTextCol("ParentSku", "Parent SKU"))
        g.Columns.Add(MakeTextCol("Width", "Width"))
        g.Columns.Add(MakeTextCol("Depth", "Depth"))
        g.Columns.Add(MakeTextCol("Height", "Height"))
        g.Columns.Add(MakeTextCol("TotalFabricSquareInches", "Total Sq In"))

        Dim isKeyboard = equipmentTypeName IsNot Nothing AndAlso
                     equipmentTypeName.Trim().ToLowerInvariant().Contains("keyboard")

        If isKeyboard Then
            g.Columns.Add(MakeTextCol("MusicRestDesign", "Music Rest Design"))
            g.Columns.Add(MakeTextCol("Chart_Template", "Chart Template"))
            g.Columns.Add(MakeTextCol("Notes", "Notes"))
            g.Columns.Add(MakeTextCol("WooProductId", "Woo Product Id", isReadOnly:=True))
        Else
            g.Columns.Add(MakeAmpHandleComboCol())
            g.Columns.Add(MakeTextCol("TAHWidth", "TAH Width"))
            g.Columns.Add(MakeTextCol("TAHHeight", "TAH Height"))
            g.Columns.Add(MakeTextCol("TAHRearOffset", "TAH Rear Offset"))
            g.Columns.Add(MakeTextCol("SAHHeight", "SAH Height"))
            g.Columns.Add(MakeTextCol("SAHWidth", "SAH Width"))
            g.Columns.Add(MakeTextCol("SAHRearOffset", "SAH Rear Offset"))
            g.Columns.Add(MakeTextCol("SAHTopDownOffset", "SAH TopDown Offset"))
            g.Columns.Add(MakeTextCol("Chart_Template", "Chart Template"))
            g.Columns.Add(MakeTextCol("Notes", "Notes"))
            g.Columns.Add(MakeTextCol("WooProductId", "Woo Product Id", isReadOnly:=True))
        End If

        ' Save button column
        g.Columns.Add(MakeSaveButtonCol())

        g.DataSource = dt
    End Sub




    ' ========= binding helpers =========
    Private Sub BindManufacturers()
        Dim cmbManufacturer = FindCombo("cmbManufacturerName")
        If cmbManufacturer Is Nothing Then Return

        Dim mfg As DataTable = DbConnectionManager.GetManufacturers()
        ' manufacturerName / ManufacturerName ; PK_manufacturerId / PK_ManufacturerId
        BindComboSmart(
            cmbManufacturer,
            mfg,
            {"manufacturerName", "ManufacturerName"},
            {"PK_manufacturerId", "PK_ManufacturerId"}
        )

        Dim cmbSeries = FindCombo("cmbSeriesName")
        If cmbSeries IsNot Nothing Then cmbSeries.DataSource = Nothing
    End Sub

    Private Sub BindEquipmentTypes()
        Dim cmbEquipmentType = FindCombo("cmbEquipmentType")
        If cmbEquipmentType Is Nothing Then Return

        Dim dt As DataTable = DbConnectionManager.GetEquipmentTypes()
        ' EquipmentTypeName / EquipmentType ; PK_equipmentTypeId / PK_EquipmentTypeId
        BindComboSmart(
            cmbEquipmentType,
            dt,
            {"EquipmentTypeName", "EquipmentType"},
            {"PK_equipmentTypeId", "PK_EquipmentTypeId"}
        )
    End Sub

    Private Sub BindOptionalSupplierArea()
        Dim cmbSupplier = FindCombo("cmbSupplier")
        Dim cmbColor = FindCombo("cmbColor")
        Dim cmbFabricType = FindCombo("cmbFabricType")
        Dim cmbBrand = FindCombo("cmbBrand")
        Dim cmbProduct = FindCombo("cmbProduct")

        If cmbSupplier Is Nothing AndAlso cmbColor Is Nothing AndAlso cmbFabricType Is Nothing AndAlso cmbBrand Is Nothing AndAlso cmbProduct Is Nothing Then
            Return
        End If

        If cmbSupplier IsNot Nothing Then
            Dim sup = DbConnectionManager.GetAllSuppliers() ' SupplierName, SupplierID
            BindComboSmart(cmbSupplier, sup, {"SupplierName"}, {"SupplierID"})
        End If

        If cmbColor IsNot Nothing Then
            Dim colors = DbConnectionManager.GetAllMaterialColors() ' ColorNameFriendly, PK_ColorNameID
            BindComboSmart(cmbColor, colors, {"ColorNameFriendly", "ColorName"}, {"PK_ColorNameID", "ColorId"})
        End If

        If cmbFabricType IsNot Nothing Then
            Dim types = DbConnectionManager.GetAllFabricTypes() ' FabricType, PK_FabricTypeNameId
            BindComboSmart(cmbFabricType, types, {"FabricType"}, {"PK_FabricTypeNameId"})
        End If

        If cmbBrand IsNot Nothing Then cmbBrand.DataSource = Nothing
        If cmbProduct IsNot Nothing Then cmbProduct.DataSource = Nothing
    End Sub

    ' ========= change handlers (NO Handles clauses) =========
    ' Manufacturer => Series
    Private Sub cmbManufacturerName_SelectedIndexChanged(sender As Object, e As EventArgs)
        If Not isFormLoaded Then Exit Sub

        Dim cmbManufacturer = TryCast(sender, ComboBox)
        Dim cmbSeries = FindCombo("cmbSeriesName")
        Dim cmbEquipmentType = FindCombo("cmbEquipmentType")

        Dim mid As Integer
        If cmbManufacturer Is Nothing OrElse cmbManufacturer.SelectedValue Is Nothing OrElse
           Not Integer.TryParse(cmbManufacturer.SelectedValue.ToString(), mid) Then

            If cmbSeries IsNot Nothing Then
                cmbSeries.BeginUpdate()
                cmbSeries.DataSource = Nothing
                cmbSeries.EndUpdate()
            End If
            If cmbEquipmentType IsNot Nothing Then cmbEquipmentType.SelectedIndex = -1
            Exit Sub
        End If

        Dim series As DataTable = DbConnectionManager.GetSeriesByManufacturer(mid) ' SeriesName / PK_SeriesId
        BindComboSmart(cmbSeries, series, {"SeriesName", "seriesName"}, {"PK_SeriesId", "PK_seriesId"})
        If cmbEquipmentType IsNot Nothing Then cmbEquipmentType.SelectedIndex = -1
    End Sub




    Private Sub SyncEquipmentTypeFromSeries()
        Dim cmbSeries = FindCombo("cmbSeriesName")
        Dim cmbEquipmentType = FindCombo("cmbEquipmentType")
        If cmbSeries Is Nothing OrElse cmbEquipmentType Is Nothing Then Exit Sub

        Dim sid? As Integer = GetSelectedInt(cmbSeries)
        If Not sid.HasValue Then
            cmbEquipmentType.SelectedIndex = -1
            Exit Sub
        End If

        Dim t = DbConnectionManager.GetEquipmentTypeForSeries(sid.Value) ' (Id As Integer?, Name As String)
        If t.Id.HasValue Then
            cmbEquipmentType.SelectedValue = t.Id.Value
        Else
            cmbEquipmentType.SelectedIndex = -1
        End If
    End Sub

    ' Supplier => Brands
    Private Sub cmbSupplier_SelectedIndexChanged(sender As Object, e As EventArgs)
        If Not isFormLoaded Then Exit Sub

        Dim cmbSupplier = TryCast(sender, ComboBox)
        Dim cmbBrand = FindCombo("cmbBrand")
        Dim cmbProduct = FindCombo("cmbProduct")

        Dim sid As Integer
        If cmbSupplier Is Nothing OrElse cmbSupplier.SelectedValue Is Nothing OrElse
           Not Integer.TryParse(cmbSupplier.SelectedValue.ToString(), sid) Then

            If cmbBrand IsNot Nothing Then
                cmbBrand.BeginUpdate()
                cmbBrand.DataSource = Nothing
                cmbBrand.EndUpdate()
            End If
            If cmbProduct IsNot Nothing Then
                cmbProduct.BeginUpdate()
                cmbProduct.DataSource = Nothing
                cmbProduct.EndUpdate()
            End If
            Exit Sub
        End If

        Dim brands As DataTable = DbConnectionManager.GetBrandsForSupplier(sid) ' BrandName / PK_FabricBrandNameId
        BindComboSmart(cmbBrand, brands, {"BrandName"}, {"PK_FabricBrandNameId"})

        If cmbProduct IsNot Nothing Then
            cmbProduct.BeginUpdate()
            cmbProduct.DataSource = Nothing
            cmbProduct.EndUpdate()
        End If
    End Sub

    ' Brand => Products
    Private Sub cmbBrand_SelectedIndexChanged(sender As Object, e As EventArgs)
        If Not isFormLoaded Then Exit Sub

        Dim cmbBrand = TryCast(sender, ComboBox)
        Dim cmbProduct = FindCombo("cmbProduct")

        Dim bid As Integer
        If cmbBrand Is Nothing OrElse cmbBrand.SelectedValue Is Nothing OrElse
           Not Integer.TryParse(cmbBrand.SelectedValue.ToString(), bid) Then
            If cmbProduct IsNot Nothing Then
                cmbProduct.BeginUpdate()
                cmbProduct.DataSource = Nothing
                cmbProduct.EndUpdate()
            End If
            Exit Sub
        End If

        Dim prods As DataTable = DbConnectionManager.GetProductsByBrandId(bid) ' BrandProductName / PK_FabricBrandProductNameId
        BindComboSmart(cmbProduct, prods, {"BrandProductName", "ProductName"}, {"PK_FabricBrandProductNameId", "ProductId"})
    End Sub

    '==============================================================================
    ' Sub: btnUploadWooListings_Click
    ' Form: frmAddModelInformation
    ' Purpose:
    '   Batch publish all models in dgvModelInformation to WooCommerce with a live
    '   progress dialog (log + progress bar + cancel).
    '
    ' Dependencies:
    '   - DataGridView named dgvModelInformation with column "PK_ModelId" (Integer).
    '   - Optional checkbox column "Include" to filter rows.
    '   - WooVariationPublisher.PublishModelAsync(modelId) As Task(Of PublishResult)
    '   - Helper ExtractModelIdsFromGrid (provided below).
    '
    ' Date: 2025-09-15
    '------------------------------------------------------------------------------
    Private Async Sub btnUploadWooListings_Click(sender As Object, e As EventArgs) Handles btnUploadWooListings.Click
        If dgvModelInformation Is Nothing OrElse dgvModelInformation.Rows.Count = 0 Then
            MessageBox.Show("No rows in the grid to upload.", "Nothing to do", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim modelIds As List(Of Integer) = ExtractModelIdsFromGrid(dgvModelInformation, "PK_ModelId", optionalIncludeColName:="Include")
        If modelIds Is Nothing OrElse modelIds.Count = 0 Then
            MessageBox.Show("No models selected. Check your grid (and 'Include' if present).", "Nothing selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim ok As New List(Of String)
        Dim bad As New List(Of String)

        Dim btn = TryCast(sender, Control)
        If btn IsNot Nothing Then btn.Enabled = False
        Cursor.Current = Cursors.WaitCursor

        Dim dlg As New frmPublishProgress()
        Try
            ' Show modeless so UI stays responsive
            dlg.Show(Me)
            dlg.StartBatch(modelIds.Count, "Publishing models to WooCommerce…")

            For i = 0 To modelIds.Count - 1
                Dim modelId = modelIds(i)

                If dlg.IsCancelRequested Then
                    dlg.LogLine("Batch canceled by user.")
                    Exit For
                End If

                dlg.SetStatus($"Publishing model {modelId} ({i + 1}/{modelIds.Count})…")
                dlg.LogLine($"Model {modelId}: starting")

                Dim result As PublishResult = Nothing
                Try
                    result = Await WooVariationPublisher.PublishModelAsync(modelId)
                Catch ex As Exception
                    result = New PublishResult With {.Success = False, .Message = ex.Message}
                End Try

                If result IsNot Nothing AndAlso result.Success Then
                    ok.Add($"Model {modelId}: Parent WooID {result.ProductId}, Variations={result.VariationResults.Count}")
                    dlg.LogLine($"Model {modelId}: ✅ success (WooID {result.ProductId}, {result.VariationResults.Count} variations)")
                Else
                    Dim msg As String = If(result?.Message, "Unknown error")
                    bad.Add($"Model {modelId}: {msg}")
                    dlg.LogLine($"Model {modelId}: ❌ {msg}")
                End If

                dlg.Advance($"Publishing models… ({i + 1}/{modelIds.Count})")
            Next

            ' Final summary
            dlg.LogLine("Batch complete.")
            Dim summary As String =
            $"Processed {ok.Count + bad.Count} of {modelIds.Count}." &
            If(ok.Count > 0, Environment.NewLine & $"Successes ({ok.Count}):" & Environment.NewLine & "- " & String.Join(Environment.NewLine & "- ", ok), "") &
            If(bad.Count > 0, Environment.NewLine & $"Failures ({bad.Count}):" & Environment.NewLine & "- " & String.Join(Environment.NewLine & "- ", bad), "")

            MessageBox.Show(summary, If(bad.Count = 0, "WooCommerce Upload Results", "WooCommerce Upload Results — Some Failures"),
                        MessageBoxButtons.OK, If(bad.Count = 0, MessageBoxIcon.Information, MessageBoxIcon.Warning))

        Finally
            dlg.Close()
            dlg.Dispose()
            Cursor.Current = Cursors.Default
            If btn IsNot Nothing Then btn.Enabled = True
        End Try
    End Sub




    Private Sub btnSavePricing_Click(sender As Object, e As EventArgs)
        Dim cmbSupplier = FindCombo("cmbSupplier")
        Dim cmbProduct = FindCombo("cmbProduct")
        Dim cmbColor = FindCombo("cmbColor")
        Dim cmbFabricType = FindCombo("cmbFabricType")
        Dim tbShipping = FindTextBox("txtShippingCost")
        Dim tbYard = FindTextBox("txtCostPerLinearYard")
        Dim tbSqIn = FindTextBox("txtCostPerSquareInch")
        Dim tbWtSqIn = FindTextBox("txtWeightPerSquareInch")

        If cmbSupplier Is Nothing OrElse cmbProduct Is Nothing OrElse cmbColor Is Nothing OrElse cmbFabricType Is Nothing _
           OrElse tbShipping Is Nothing OrElse tbYard Is Nothing OrElse tbSqIn Is Nothing OrElse tbWtSqIn Is Nothing Then
            Return
        End If

        If cmbSupplier.SelectedValue Is Nothing OrElse cmbProduct.SelectedValue Is Nothing OrElse
           cmbColor.SelectedValue Is Nothing OrElse cmbFabricType.SelectedValue Is Nothing Then
            MessageBox.Show("Please pick Supplier, Product, Color, and Fabric Type first.")
            Exit Sub
        End If

        Dim supplierId As Integer = CInt(cmbSupplier.SelectedValue)
        Dim productId As Integer = CInt(cmbProduct.SelectedValue)
        Dim colorId As Integer = CInt(cmbColor.SelectedValue)
        Dim fabricTypeId As Integer = CInt(cmbFabricType.SelectedValue)

        Dim row = DbConnectionManager.GetSupplierProductNameData(supplierId, productId, colorId, fabricTypeId)
        If row Is Nothing Then
            MessageBox.Show("No supplier/product mapping exists for that combination.")
            Exit Sub
        End If

        Dim spndId As Integer = CInt(row("PK_SupplierProductNameDataId"))
        Dim shipping As Decimal = ParseDecimal(tbShipping)
        Dim costPerYard As Decimal = ParseDecimal(tbYard)
        Dim costPerSqIn As Decimal = ParseDecimal(tbSqIn)
        Dim weightPerSqIn As Decimal = ParseDecimal(tbWtSqIn)

        DbConnectionManager.InsertFabricPricingHistory(spndId, shipping, costPerYard, costPerSqIn, weightPerSqIn)
        MessageBox.Show("Pricing history saved.")
    End Sub

    Private Sub btnUpdateProductInfo_Click(sender As Object, e As EventArgs)
        Dim cmbProduct = FindCombo("cmbProduct")
        Dim tbWtYard = FindTextBox("txtWeightPerLinearYard")
        Dim tbRollWidth = FindTextBox("txtFabricRollWidth")

        If cmbProduct Is Nothing OrElse tbWtYard Is Nothing OrElse tbRollWidth Is Nothing Then Return
        If cmbProduct.SelectedValue Is Nothing Then
            MessageBox.Show("Pick a product first.")
            Exit Sub
        End If

        Dim productId As Integer = CInt(cmbProduct.SelectedValue)
        Dim weight As Decimal = ParseDecimal(tbWtYard)
        Dim rollWidth As Decimal = ParseDecimal(tbRollWidth)

        DbConnectionManager.UpdateFabricProductInfo(productId, weight, rollWidth)
        MessageBox.Show("Product info updated.")
    End Sub


    '==============================================================================
    ' Function: ExtractModelIdsFromGrid
    ' Purpose:
    '   Return distinct positive PK_ModelId values from the grid.
    '   If optionalIncludeColName exists, only rows with True in that column are used.
    ' Date: 2025-09-15
    '------------------------------------------------------------------------------
    Private Function ExtractModelIdsFromGrid(grid As DataGridView, idColName As String, Optional optionalIncludeColName As String = Nothing) As List(Of Integer)
        Dim result As New List(Of Integer)
        If grid Is Nothing OrElse grid.Rows.Count = 0 Then Return result
        If Not grid.Columns.Contains(idColName) Then Throw New InvalidOperationException($"Grid is missing required column '{idColName}'.")

        Dim useInclude As Boolean = Not String.IsNullOrWhiteSpace(optionalIncludeColName) AndAlso grid.Columns.Contains(optionalIncludeColName)

        For Each r As DataGridViewRow In grid.Rows
            If r.IsNewRow Then Continue For

            If useInclude Then
                Dim inc As Boolean = False
                Dim incObj = r.Cells(optionalIncludeColName).Value
                If incObj IsNot Nothing AndAlso incObj IsNot DBNull.Value Then Boolean.TryParse(incObj.ToString(), inc)
                If Not inc Then Continue For
            End If

            Dim idObj = r.Cells(idColName).Value
            Dim id As Integer
            If idObj IsNot Nothing AndAlso idObj IsNot DBNull.Value AndAlso Integer.TryParse(idObj.ToString(), id) AndAlso id > 0 Then
                result.Add(id)
            End If
        Next

        Return result.Distinct().ToList()
    End Function

    Private Sub LoadModelsForSelectedSeries()
        If Not isFormLoaded Then Exit Sub

        Dim cmbManufacturer = FindCombo("cmbManufacturerName")
        Dim cmbSeries = FindCombo("cmbSeriesName")
        Dim grid = FindGrid("dgvModelInformation")
        If cmbManufacturer Is Nothing OrElse cmbSeries Is Nothing OrElse grid Is Nothing Then Exit Sub

        Dim mid As Integer
        Dim sid As Integer
        If cmbManufacturer.SelectedValue Is Nothing _
       OrElse Not Integer.TryParse(cmbManufacturer.SelectedValue.ToString(), mid) _
       OrElse cmbSeries.SelectedValue Is Nothing _
       OrElse Not Integer.TryParse(cmbSeries.SelectedValue.ToString(), sid) Then

            grid.DataSource = Nothing
            grid.Rows.Clear()
            Exit Sub
        End If

        Dim models As DataTable = DbConnectionManager.GetModelsForSeries(mid, sid)

        grid.AutoGenerateColumns = True ' or False if you use custom columns
        grid.DataSource = models

        ' Optional: make it neat
        If grid.Columns.Contains("PK_ModelId") Then
            grid.Columns("PK_ModelId").HeaderText = "Model ID"
            grid.Columns("PK_ModelId").Width = 90
        End If
        If grid.Columns.Contains("ModelName") Then
            grid.Columns("ModelName").HeaderText = "Model Name"
            grid.Columns("ModelName").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End If
    End Sub

    ' NO Handles clause – you’re wiring this via AddHandler in WireEvents()
    Private isLoadingSeries As Boolean = False  ' keep if you already use this flag

    Private Sub cmbSeriesName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSeriesName.SelectedIndexChanged
        If isLoadingSeries Then Return

        ' Get manufacturer & series IDs from the bound combos
        If cmbManufacturerName.SelectedValue Is Nothing OrElse cmbSeriesName.SelectedValue Is Nothing Then
            dgvModelInformation.DataSource = Nothing
            dgvModelInformation.Rows.Clear()
            dgvModelInformation.Columns.Clear()
            Return
        End If

        Dim manufacturerId As Integer
        Dim seriesId As Integer
        If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), manufacturerId) Then
            dgvModelInformation.DataSource = Nothing : Return
        End If
        If Not Integer.TryParse(cmbSeriesName.SelectedValue.ToString(), seriesId) Then
            dgvModelInformation.DataSource = Nothing : Return
        End If

        ' 1) What equipment type is this series?
        Dim et = DbConnectionManager.GetEquipmentTypeForSeries(seriesId)
        Dim equipmentTypeName As String = et.Name

        ' 2) Configure the visible columns in the requested order
        ConfigureModelGridColumns(equipmentTypeName)

        ' 3) Fetch the rows and bind
        Dim dt As DataTable = DbConnectionManager.GetModelGridRows(manufacturerId, seriesId, equipmentTypeName)
        dgvModelInformation.DataSource = dt
    End Sub

    'Private Sub cmbSeriesName_SelectedIndexChanged(sender As Object, e As EventArgs)
    '    If Not isFormLoaded Then Exit Sub
    '    SyncEquipmentTypeFromSeries()
    '    LoadModelsForSelectedSeries()
    'End Sub

    ' NO Handles clause – wired dynamically
    Private Sub cmbSeriesName_SelectionChangeCommitted(sender As Object, e As EventArgs)
        If Not isFormLoaded Then Exit Sub
        SyncEquipmentTypeFromSeries()
        LoadModelsForSelectedSeries()
    End Sub

    ' ===== Helpers (place inside frmAddModelInformation class) =====
    Private Function FindCombo(name As String) As ComboBox
        Dim arr = Me.Controls.Find(name, True)
        If arr Is Nothing OrElse arr.Length = 0 Then Return Nothing
        Return TryCast(arr(0), ComboBox)
    End Function

    Private Function FindGrid(name As String) As DataGridView
        Dim arr = Me.Controls.Find(name, True)
        If arr Is Nothing OrElse arr.Length = 0 Then Return Nothing
        Return TryCast(arr(0), DataGridView)
    End Function

    Private Function FindButton(name As String) As Button
        Dim arr = Me.Controls.Find(name, True)
        If arr Is Nothing OrElse arr.Length = 0 Then Return Nothing
        Return TryCast(arr(0), Button)
    End Function

    ' Safely read SelectedValue as Integer?
    Private Function GetSelectedInt(cmb As ComboBox) As Integer?
        If cmb Is Nothing OrElse cmb.SelectedValue Is Nothing Then Return Nothing
        Dim i As Integer
        If Integer.TryParse(cmb.SelectedValue.ToString(), i) Then Return i
        Return Nothing
    End Function



    Private Function MakeSaveButtonCol() As DataGridViewButtonColumn
        Return New DataGridViewButtonColumn With {
        .Name = "colSave",
        .HeaderText = "",
        .Text = "Save",
        .UseColumnTextForButtonValue = True
    }
    End Function

    Private Function MakeAmpHandleComboCol() As DataGridViewComboBoxColumn
        ' expects a lookup table of handle locations
        Dim src As DataTable = DbConnectionManager.GetAmpHandleLocations()
        Return New DataGridViewComboBoxColumn With {
        .Name = "AmpHandleLocation",
        .HeaderText = "Amp Handle Location",
        .DataPropertyName = "AmpHandleLocation",
        .DisplayMember = "HandleLocationName",
        .ValueMember = "HandleLocationName",
        .DataSource = src,
        .FlatStyle = FlatStyle.Standard
    }
    End Function


    Private Sub RefreshModelsGridFromSelections()
        Dim cmbManufacturer = FindCombo("cmbManufacturerName")
        Dim cmbSeries = FindCombo("cmbSeriesName")
        Dim grid = dgvModelInformation
        If grid Is Nothing Then Exit Sub

        Dim mid? As Integer = GetSelectedInt(cmbManufacturer)
        Dim sid? As Integer = GetSelectedInt(cmbSeries)
        If Not mid.HasValue OrElse Not sid.HasValue Then
            grid.DataSource = Nothing
            Exit Sub
        End If

        ' Get equipment type name so we know which layout to build
        Dim etInfo = DbConnectionManager.GetEquipmentTypeForSeries(sid.Value)
        ConfigureModelGridColumns(etInfo.Name)

        ' Pull full rows
        Dim dt = DbConnectionManager.GetModelsForGrid(mid.Value, sid.Value)

        ' Normalize column name so the grid can bind to "ModelName"
        If dt.Columns.Contains("modelName") AndAlso Not dt.Columns.Contains("ModelName") Then
            dt.Columns("modelName").ColumnName = "ModelName"
        End If

        ' Bind
        dgvModelInformation.RowHeadersVisible = False
        dgvModelInformation.DataSource = dt

        grid.DataSource = dt
    End Sub

    Private Sub dgvModelInformation_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvModelInformation.CellContentClick
        If e.RowIndex < 0 Then Return
        Dim g = dgvModelInformation
        If g.Columns(e.ColumnIndex).Name <> "colSave" Then Return
        g.EndEdit()

        Dim r = g.Rows(e.RowIndex)

        Dim id = CInt(r.Cells("PK_ModelId").Value)
        Dim parentSku = GetCellValue(Of String)(r, "ParentSku")
        Dim width? = GetCellDecimal(r, "Width")
        Dim depth? = GetCellDecimal(r, "Depth")
        Dim height? = GetCellDecimal(r, "Height")
        Dim tfsi? = GetCellDecimal(r, "TotalFabricSquareInches")

        Dim ampLoc As String = Nothing
        Dim tahW? As Decimal = Nothing, tahH? As Decimal = Nothing, tahRear? As Decimal = Nothing
        Dim sahH? As Decimal = Nothing, sahW? As Decimal = Nothing, sahRear? As Decimal = Nothing, sahTD? As Decimal = Nothing
        Dim mrDesign As String = Nothing

        If Not _isKeyboardLayout Then
            ampLoc = GetCellValue(Of String)(r, "AmpHandleLocation")
            tahW = GetCellDecimal(r, "TAHWidth")
            tahH = GetCellDecimal(r, "TAHHeight")
            tahRear = GetCellDecimal(r, "TAHRearOffset")
            sahH = GetCellDecimal(r, "SAHHeight")
            sahW = GetCellDecimal(r, "SAHWidth")
            sahRear = GetCellDecimal(r, "SAHRearOffset")
            sahTD = GetCellDecimal(r, "SAHTopDownOffset")
        Else
            mrDesign = GetCellValue(Of String)(r, "MusicRestDesign")
        End If

        Dim chartTempl = GetCellValue(Of String)(r, "Chart_Template")
        Dim notes = GetCellValue(Of String)(r, "Notes")

        Try
            DbConnectionManager.UpdateModelDisplayFields(
            id, parentSku, width, depth, height, tfsi,
            ampLoc, tahW, tahH, tahRear,
            sahH, sahW, sahRear, sahTD,
            mrDesign, chartTempl, notes
        )
            MessageBox.Show("Row saved.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Save failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgvModelInformation_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dgvModelInformation.DataError
        ' Avoid hard crashes on parse/combobox issues; show a friendly message if needed
        e.ThrowException = False
    End Sub

    '==============================================================================
    ' Sub: btnPublishSelectedSeries_Click
    ' Form: frmAddModelInformation
    ' Purpose:
    '   Publish all models for the *currently selected* Manufacturer + Series
    '   (from cmbManufacturerName, cmbSeriesName) to WooCommerce.
    '   - Does NOT read anything from dgvModelInformation
    '   - Loads models directly from DB via joins (robust to schema quirks)
    '   - Streams progress with frmPublishProgress
    '
    ' Dependencies:
    '   - ComboBoxes: cmbManufacturerName (SelectedValue = PK_ManufacturerId),
    '                 cmbSeriesName (SelectedValue = PK_SeriesId)
    '   - DbConnectionManager.GetConnection(), EnsureOpen(conn)
    '   - WooVariationPublisher.PublishModelAsync(modelId)
    '   - frmPublishProgress (provided earlier)
    '
    ' Date: 2025-09-15
    '------------------------------------------------------------------------------
    Private Async Sub btnPublishSelectedSeries_Click(sender As Object, e As EventArgs) Handles btnPublishSelectedSeries.Click
        ' 1) Read selections safely
        Dim manufacturerId As Integer = 0
        Dim seriesId As Integer = 0

        If cmbManufacturerName Is Nothing OrElse cmbManufacturerName.SelectedValue Is Nothing _
           OrElse Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), manufacturerId) _
           OrElse manufacturerId <= 0 Then
            MessageBox.Show("Pick a Manufacturer first.", "Missing selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If cmbSeriesName Is Nothing OrElse cmbSeriesName.SelectedValue Is Nothing _
           OrElse Not Integer.TryParse(cmbSeriesName.SelectedValue.ToString(), seriesId) _
           OrElse seriesId <= 0 Then
            MessageBox.Show("Pick a Series first.", "Missing selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        ' 2) Load models for this Manufacturer + Series (direct from DB, not the grid)
        Dim dt As DataTable = LoadModelsForSelection(manufacturerId, seriesId)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            MessageBox.Show("No models found for that Manufacturer/Series.", "Nothing to publish", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        ' Pull the list of ModelIds we will publish
        Dim modelIds As List(Of Integer) =
            (From r As DataRow In dt.Rows
             Let v = If(r("PK_ModelId"), 0)
             Let mid = If(v Is DBNull.Value, 0, Convert.ToInt32(v))
             Where mid > 0
             Select mid).Distinct().ToList()

        If modelIds.Count = 0 Then
            MessageBox.Show("No valid Model IDs found.", "Nothing to publish", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        ' Optional: nice label for the dialog title
        Dim manuName As String = SafeField(dt, "ManufacturerName")
        Dim seriesName As String = SafeField(dt, "SeriesName")

        ' 3) Show progress dialog and publish
        Dim ok As New List(Of String)
        Dim bad As New List(Of String)

        Dim btn = TryCast(sender, Control)
        If btn IsNot Nothing Then btn.Enabled = False
        Cursor.Current = Cursors.WaitCursor

        Dim dlg As New frmPublishProgress()
        Try
            dlg.Show(Me)
            Dim title As String = If(String.IsNullOrWhiteSpace(manuName) AndAlso String.IsNullOrWhiteSpace(seriesName),
                                     "Publishing selected series…",
                                     $"Publishing: {manuName} • {seriesName}")
            dlg.StartBatch(modelIds.Count, title)

            For i = 0 To modelIds.Count - 1
                Dim mid = modelIds(i)

                If dlg.IsCancelRequested Then
                    dlg.LogLine("Batch canceled by user.")
                    Exit For
                End If

                dlg.SetStatus($"Publishing model {mid} ({i + 1}/{modelIds.Count})…")
                dlg.LogLine($"Model {mid}: starting")

                Dim result As PublishResult = Nothing
                Try
                    result = Await WooVariationPublisher.PublishModelAsync(mid)
                Catch ex As Exception
                    result = New PublishResult With {.Success = False, .Message = ex.Message}
                End Try

                If result IsNot Nothing AndAlso result.Success Then
                    ok.Add($"Model {mid}: Parent WooID {result.ProductId}, Variations={result.VariationResults.Count}")
                    dlg.LogLine($"Model {mid}: ✅ success (WooID {result.ProductId}, {result.VariationResults.Count} variations)")
                Else
                    Dim msg As String = If(result?.Message, "Unknown error")
                    bad.Add($"Model {mid}: {msg}")
                    dlg.LogLine($"Model {mid}: ❌ {msg}")
                End If

                dlg.Advance($"Publishing models… ({i + 1}/{modelIds.Count})")
            Next

            ' 4) Final summary
            dlg.LogLine("Batch complete.")
            Dim summary As String =
                $"Processed {ok.Count + bad.Count} of {modelIds.Count}." &
                If(ok.Count > 0, Environment.NewLine & $"Successes ({ok.Count}):" & Environment.NewLine & "- " & String.Join(Environment.NewLine & "- ", ok), "") &
                If(bad.Count > 0, Environment.NewLine & $"Failures ({bad.Count}):" & Environment.NewLine & "- " & String.Join(Environment.NewLine & "- ", bad), "")

            MessageBox.Show(summary,
                            If(bad.Count = 0, "WooCommerce Upload Results", "WooCommerce Upload Results — Some Failures"),
                            MessageBoxButtons.OK,
                            If(bad.Count = 0, MessageBoxIcon.Information, MessageBoxIcon.Warning))

        Finally
            dlg.Close()
            dlg.Dispose()
            Cursor.Current = Cursors.Default
            If btn IsNot Nothing Then btn.Enabled = True
        End Try
    End Sub
    '==============================================================================
    ' Function: LoadModelsForSelection
    ' Form: frmAddModelInformation
    ' Purpose:
    '   Return a DataTable of models for a given Manufacturer + Series, including
    '   only the columns the publisher truly needs:
    '     - PK_ModelId
    '     - ModelName
    '     - ParentSKU
    '     - FK_SeriesId  (so PublishModelAsync can resolve category)
    '     - ManufacturerName, SeriesName (nice-to-have for dialog labeling)
    '
    ' Notes:
    '   - We do NOT touch Model.FK_ManufacturerId (which caused your error).
    '     Instead, we filter via the Series → Manufacturer chain:
    '       Model (FK_SeriesId) → ModelSeries (FK_ManufacturerId)
    '
    ' Date: 2025-09-15
    '------------------------------------------------------------------------------
    Private Function LoadModelsForSelection(manufacturerId As Integer, seriesId As Integer) As DataTable
        Dim dt As New DataTable()

        Using conn As SqlConnection = DbConnectionManager.GetConnection()
            DbConnectionManager.EnsureOpen(conn)

            ' We filter by SeriesId and ManufacturerId using joins.
            ' Adjust table/column names if your actual schema differs.
            Dim sql As String = "
            SELECT
                m.PK_ModelId,
                m.ModelName,
                m.ParentSKU,
                m.FK_SeriesId,
                mf.ManufacturerName,
                s.SeriesName
            FROM Model              AS m
            INNER JOIN ModelSeries  AS s  ON s.PK_SeriesId        = m.FK_SeriesId
            INNER JOIN ModelManufacturers AS mf ON mf.PK_ManufacturerId = s.FK_ManufacturerId
            WHERE s.PK_SeriesId = @sid
              AND mf.PK_ManufacturerId = @mid
            ORDER BY m.ModelName;"

            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.Add("@sid", SqlDbType.Int).Value = seriesId
                cmd.Parameters.Add("@mid", SqlDbType.Int).Value = manufacturerId

                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using

        Return dt
    End Function
    '==============================================================================
    ' Function: SafeField
    ' Form: frmAddModelInformation
    ' Purpose:
    '   Pull a single string value from the first row of a DataTable if present.
    '------------------------------------------------------------------------------
    Private Function SafeField(dt As DataTable, colName As String) As String
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return ""
        If Not dt.Columns.Contains(colName) Then Return ""
        Dim v = dt.Rows(0)(colName)
        If v Is Nothing OrElse v Is DBNull.Value Then Return ""
        Return v.ToString().Trim()
    End Function

End Class

