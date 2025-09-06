Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Windows.Forms

Partial Public Class frmFabricEntryForm
    Inherits Form

    Private isFormLoading As Boolean = False
    Private suppressProductSelectionEvent As Boolean = False
    Private fabricTypes As DataTable
    Private originalRowValues As New Dictionary(Of Integer, Dictionary(Of String, Object))()
    Private Sub frmFabricEntryForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        isFormLoading = True
        LoadAllBrandsToCombo()
        Dim db As New DbConnectionManager()
        Dim suppliers = db.GetAllSuppliers() ' Returns DataTable or List(Of SupplierInformation)
        cmbSupplier.DataSource = suppliers
        cmbSupplier.DisplayMember = "SupplierName" ' Adjust to your actual column/property name
        cmbSupplier.ValueMember = "SupplierID"     ' Adjust to your actual column/property name

        ' Add this line to ensure no supplier is selected by default
        cmbSupplier.SelectedIndex = -1


        ' Load all colors into cmbColor
        LoadAllColorsToCombo()

        ' Optionally, clear related controls
        cmbBrand.SelectedIndex = -1
        cmbProduct.DataSource = Nothing
        cmbFabricType.DataSource = Nothing

        ' Make calculated fields read-only
        txtSquareInchesPerLinearYard.ReadOnly = True
        txtCostPerSquareInch.ReadOnly = True
        txtWeightPerSquareInch.ReadOnly = True

        txtShippingCost.ReadOnly = True
        txtCostPerLinearYard.ReadOnly = True
        txtTotalYards.ReadOnly = True

        If cmbSupplier.SelectedValue IsNot Nothing AndAlso TypeOf cmbSupplier.SelectedValue Is Integer Then
            InitializeAssignFabricsGrid(CInt(cmbSupplier.SelectedValue))
        End If
        isFormLoading = False
    End Sub

    Private Sub LoadAllColorsToCombo()
        Dim db As New DbConnectionManager()
        Dim colors As DataTable = db.GetAllMaterialColors() ' You need to implement this method
        cmbColor.DataSource = colors
        cmbColor.DataSource = db.GetAllMaterialColors()
        cmbColor.DisplayMember = "ColorNameFriendly"
        cmbColor.ValueMember = "PK_ColorNameID"
        cmbColor.SelectedIndex = -1
    End Sub
    ' 4. Add new brand if not exists (Validating event)
    Private Sub cmbBrand_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cmbBrand.Validating
        Dim enteredBrand As String = cmbBrand.Text.Trim()
        If String.IsNullOrWhiteSpace(enteredBrand) Then Return

        Dim exists As Boolean = False
        For Each item As DataRowView In cmbBrand.Items
            If String.Equals(item("BrandName").ToString(), enteredBrand, StringComparison.OrdinalIgnoreCase) Then
                exists = True
                cmbBrand.SelectedValue = item("PK_FabricBrandNameId")
                Exit For
            End If
        Next

        If Not exists Then
            Dim db As New DbConnectionManager()
            Dim newId As Integer
            Using conn = DbConnectionManager.GetConnection()
                If conn.State <> ConnectionState.Open Then conn.Open()
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "INSERT INTO FabricBrandName (BrandName) VALUES (@BrandName); SELECT CAST(SCOPE_IDENTITY() AS int);"
                    cmd.Parameters.AddWithValue("@BrandName", enteredBrand)
                    newId = CInt(cmd.ExecuteScalar())
                End Using
            End Using

            Dim dt = CType(cmbBrand.DataSource, DataTable)
            Dim newRow = dt.NewRow()
            newRow("PK_FabricBrandNameId") = newId
            newRow("BrandName") = enteredBrand
            dt.Rows.Add(newRow)
            cmbBrand.SelectedValue = newId
        End If
    End Sub
    ' 2. Load all products for a brand into cmbProduct as DataTable
    Private Sub LoadAllProductsForBrand(brandId As Integer)



        Dim db As New DbConnectionManager()
        Dim products As DataTable = db.GetProductsByBrandId(brandId) ' Must return DataTable
        cmbProduct.DataSource = products
        cmbProduct.DisplayMember = "BrandProductName"
        cmbProduct.ValueMember = "PK_FabricBrandProductNameId"
        cmbProduct.SelectedIndex = -1
        Dim dt As DataTable = CType(cmbProduct.DataSource, DataTable)
        Dim allNames = String.Join(", ", dt.AsEnumerable().Select(Function(r) r.Field(Of String)("BrandProductName")))
        ' MessageBox.Show("Products loaded: " & allNames)
    End Sub
    ' 5. Add new product if not exists (Validating event)
    Private Sub cmbProduct_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cmbProduct.Validating
        Dim enteredProduct As String = cmbProduct.Text.Trim()
        If String.IsNullOrWhiteSpace(enteredProduct) Then
            txtFabricBrandProductName.Clear()
            Return
        End If

        ' Update the textbox with whatever is in the ComboBox, even if not in the list
        txtFabricBrandProductName.Text = enteredProduct

        ' Optionally, select the existing product if it matches
        For Each item As DataRowView In cmbProduct.Items
            If String.Equals(item("BrandProductName").ToString(), enteredProduct, StringComparison.OrdinalIgnoreCase) Then
                cmbProduct.SelectedValue = item("PK_FabricBrandProductNameId")
                Exit For
            End If
        Next

        ' Do not insert or validate fabric type here.
        ' Product will be added (if needed) in the Save button logic.
    End Sub

    Private Sub dgvAssignFabrics_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAssignFabrics.RowEnter
        Dim row = dgvAssignFabrics.Rows(e.RowIndex)
        If row.IsNewRow Then Exit Sub

        Dim values As New Dictionary(Of String, Object)
        For Each col As DataGridViewColumn In dgvAssignFabrics.Columns
            values(col.Name) = row.Cells(col.Name).Value
        Next
        originalRowValues(e.RowIndex) = values
    End Sub
    Private Sub dgvAssignFabrics_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAssignFabrics.CellContentClick
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return

        If dgvAssignFabrics.Columns(e.ColumnIndex).Name = "Save" Then
            Dim row = dgvAssignFabrics.Rows(e.RowIndex)
            If row.IsNewRow Then Return

            ' Gather current values
            Dim changedFields As New List(Of String)
            Dim changes As New List(Of String)
            Dim currentValues As New Dictionary(Of String, Object)
            For Each col As DataGridViewColumn In dgvAssignFabrics.Columns
                currentValues(col.Name) = row.Cells(col.Name).Value
            Next

            Dim originalValues As Dictionary(Of String, Object) = Nothing
            If originalRowValues.TryGetValue(e.RowIndex, originalValues) Then
                For Each kv In currentValues
                    If originalValues.ContainsKey(kv.Key) AndAlso Not Object.Equals(originalValues(kv.Key), kv.Value) Then
                        changedFields.Add(kv.Key)
                        changes.Add($"{kv.Key}: {originalValues(kv.Key)} → {kv.Value}")
                    End If
                Next
            End If

            If changes.Count = 0 Then
                MessageBox.Show("No changes to save for this row.", "No Changes", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim msg = "You are about to update this row. Changes:" & vbCrLf & String.Join(vbCrLf, changes) & vbCrLf & "Do you want to save these changes?"
            If MessageBox.Show(msg, "Confirm Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                ' --- Update DB here ---
                Dim db As New DbConnectionManager()
                Dim supplierId = CInt(cmbSupplier.SelectedValue)
                Dim productId = CInt(row.Cells("PK_FabricBrandProductNameId").Value)
                Dim colorId = CInt(row.Cells("FK_ColorNameID").Value)
                Dim fabricTypeId = CInt(row.Cells("FabricType").Value)
                Dim supplierProduct = db.GetSupplierProductNameData(supplierId, productId, colorId, fabricTypeId)
                If supplierProduct Is Nothing Then
                    MessageBox.Show("Could not find supplier-product record.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
                Dim supplierProductNameDataId = CInt(supplierProduct("PK_SupplierProductNameDataId"))

                Dim shippingCost = Convert.ToDecimal(row.Cells("ShippingCost").Value)
                Dim costPerLinearYard = Convert.ToDecimal(row.Cells("CostPerLinearYard").Value)
                Dim weightPerLinearYard = Convert.ToDecimal(row.Cells("WeightPerLinearYard").Value)
                Dim fabricRollWidth = Convert.ToDecimal(row.Cells("FabricRollWidth").Value)
                Dim totalYards = Convert.ToDecimal(row.Cells("TotalYards").Value)

                Dim totalCost As Decimal = (costPerLinearYard * totalYards) + shippingCost
                Dim totalCostPerLinearYard As Decimal = If(totalYards > 0, totalCost / totalYards, 0D)
                Dim squareInchesPerLinearYard As Decimal = fabricRollWidth * 36D
                Dim costPerSquareInch As Decimal = If(squareInchesPerLinearYard > 0, Math.Round(totalCostPerLinearYard / squareInchesPerLinearYard, 5), 0D)
                Dim weightPerSquareInch As Decimal = If(squareInchesPerLinearYard > 0, Math.Round(weightPerLinearYard / squareInchesPerLinearYard, 5), 0D)

                row.Cells("CostPerSquareInch").Value = costPerSquareInch
                row.Cells("WeightPerSquareInch").Value = weightPerSquareInch

                ' Update pricing history if needed
                If changedFields.Contains("ShippingCost") OrElse changedFields.Contains("CostPerLinearYard") Then
                    db.InsertFabricPricingHistory(supplierProductNameDataId, shippingCost, costPerLinearYard, costPerSquareInch, weightPerSquareInch)
                End If
                ' Update product info if needed
                If changedFields.Contains("WeightPerLinearYard") OrElse changedFields.Contains("FabricRollWidth") Then
                    db.UpdateFabricProductInfo(productId, weightPerLinearYard, fabricRollWidth)
                End If

                MessageBox.Show("Changes saved.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ' Optionally, refresh the grid or update originalRowValues
                originalRowValues(e.RowIndex) = New Dictionary(Of String, Object)(currentValues)
            End If
        End If
    End Sub

    Private Sub InitializeAssignFabricsGrid(supplierId As Integer)
        dgvAssignFabrics.Columns.Clear()
        Dim db As New DbConnectionManager()

        ' Only brands linked to this supplier
        Dim brandNames = db.GetBrandsForSupplier(supplierId)
        fabricTypes = db.GetAllFabricTypes()


        For Each row As DataRow In fabricTypes.Rows
            row("PK_FabricTypeNameId") = Convert.ToInt32(row("PK_FabricTypeNameId"))
        Next
        Dim allProducts = db.GetProductsForSupplier(supplierId)

        '    dgvAssignFabrics.Columns.Add(New DataGridViewComboBoxColumn With {
        '    .Name = "BrandName",
        '    .HeaderText = "Brand",
        '    .DataSource = brandNames,
        '    .DisplayMember = "BrandName",
        '    .ValueMember = "PK_FabricBrandNameId",
        '    .DataPropertyName = "PK_FabricBrandNameId"
        '})
        ' Add this TextBox column instead:
        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {
    .Name = "BrandName",
    .HeaderText = "Brand",
    .DataPropertyName = "BrandName",
    .ReadOnly = True
})
        '    dgvAssignFabrics.Columns.Add(New DataGridViewComboBoxColumn With {
        '    .Name = "ProductName",
        '    .HeaderText = "Fabric",
        '    .DataSource = allProducts,
        '    .DisplayMember = "BrandProductName",
        '    .ValueMember = "PK_FabricBrandProductNameId",
        '    .DataPropertyName = "PK_FabricBrandProductNameId"
        '})


        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {
    .Name = "ProductName",
    .HeaderText = "Fabric",
    .DataPropertyName = "BrandProductName",
    .ReadOnly = True
})
        ' Add ColorNameFriendly as a read-only text column
        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {
        .Name = "ColorNameFriendly",
        .HeaderText = "Color",
        .DataPropertyName = "ColorNameFriendly",
        .ReadOnly = True
    })

        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {
    .Name = "FK_ColorNameID",
    .HeaderText = "Color ID",
    .DataPropertyName = "FK_ColorNameID",
    .Visible = False
})
        dgvAssignFabrics.Columns.Add(New DataGridViewComboBoxColumn With {
        .Name = "FabricType",
        .HeaderText = "Fabric Type",
        .DataSource = fabricTypes,
        .DisplayMember = "FabricType",
        .ValueMember = "PK_FabricTypeNameId",
        .DataPropertyName = "FK_FabricTypeNameId"
    })
        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "ShippingCost", .HeaderText = "Shipping", .DataPropertyName = "ShippingCost", .ValueType = GetType(Decimal)})
        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "CostPerLinearYard", .HeaderText = "Cost/Linear Yard", .DataPropertyName = "CostPerLinearYard", .ValueType = GetType(Decimal)})
        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "CostPerSquareInch", .HeaderText = "Cost/Square Inch", .DataPropertyName = "CostPerSquareInch", .ValueType = GetType(Decimal)})
        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "WeightPerSquareInch", .HeaderText = "Weight/Square Inch", .DataPropertyName = "WeightPerSquareInch", .ValueType = GetType(Decimal)})
        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "WeightPerLinearYard", .HeaderText = "Weight/Linear Yard", .DataPropertyName = "WeightPerLinearYard", .ValueType = GetType(Decimal)})
        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "FabricRollWidth", .HeaderText = "Roll Width", .DataPropertyName = "FabricRollWidth", .ValueType = GetType(Decimal)})
        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "TotalYards", .HeaderText = "Total Yards", .DataPropertyName = "TotalYards", .ValueType = GetType(Decimal)})

        Dim saveButtonCol As New DataGridViewButtonColumn With {
        .Name = "Save",
        .HeaderText = "Save",
        .Text = "Save",
        .UseColumnTextForButtonValue = True,
        .Width = 60
    }
        dgvAssignFabrics.Columns.Add(saveButtonCol)
        FormatAssignFabricsGrid()
    End Sub
    ' 1. Load all brands into cmbBrand as DataTable
    Private Sub LoadAllBrandsToCombo()
        Dim db As New DbConnectionManager()
        Dim brands As DataTable = db.GetAllFabricBrandNames() ' Must return DataTable
        cmbBrand.DataSource = brands
        cmbBrand.DisplayMember = "BrandName"
        cmbBrand.ValueMember = "PK_FabricBrandNameId"
        cmbBrand.SelectedIndex = -1
    End Sub
    ' 2. Update LoadSupplierProductsToGrid to load TotalYards from the DB
    Private Sub LoadSupplierProductsToGrid(supplierId As Integer)
        Dim dt As New DataTable()
        dt.Columns.Add("PK_FabricBrandNameId", GetType(Integer))
        dt.Columns.Add("BrandName", GetType(String))
        dt.Columns.Add("PK_FabricBrandProductNameId", GetType(Integer))
        dt.Columns.Add("BrandProductName", GetType(String))
        dt.Columns.Add("FK_ColorNameID", GetType(Integer)) ' <-- ADD THIS
        dt.Columns.Add("ColorNameFriendly", GetType(String))
        dt.Columns.Add("FK_FabricTypeNameId", GetType(Integer))
        dt.Columns.Add("ShippingCost", GetType(Decimal))
        dt.Columns.Add("CostPerLinearYard", GetType(Decimal))
        dt.Columns.Add("CostPerSquareInch", GetType(Decimal))
        dt.Columns.Add("WeightPerSquareInch", GetType(Decimal))
        dt.Columns.Add("WeightPerLinearYard", GetType(Decimal))
        dt.Columns.Add("FabricRollWidth", GetType(Decimal))
        dt.Columns.Add("TotalYards", GetType(Decimal))
        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {
        .Name = "PK_FabricBrandProductNameId",
        .HeaderText = "Product ID",
        .DataPropertyName = "PK_FabricBrandProductNameId",
        .Visible = False
    })

        Using conn = DbConnectionManager.GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT 
                    b.PK_FabricBrandNameId AS PK_FabricBrandNameId,
                    b.BrandName,
                    p.PK_FabricBrandProductNameId AS PK_FabricBrandProductNameId,
                    p.BrandProductName AS BrandProductName,
                    j.FK_ColorNameID, -- <-- ADD THIS
                    j.FK_FabricTypeNameId,
                    c.ColorNameFriendly,
                    fp.ShippingCost,
                    fp.CostPerLinearYard,
                    fp.CostPerSquareInch,
                    fp.WeightPerSquareInch,
                    p.WeightPerLinearYard,
                    p.FabricRollWidth,
                    s.TotalYards
                FROM SupplierProductNameData s
                INNER JOIN JoinProductColorFabricType j ON s.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                INNER JOIN FabricBrandProductName p ON j.FK_FabricBrandProductNameId = p.PK_FabricBrandProductNameId
                INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                INNER JOIN FabricColor c ON j.FK_ColorNameID = c.PK_ColorNameID
                OUTER APPLY (
                    SELECT TOP 1 
                        fph.ShippingCost,
                        fph.CostPerLinearYard,
                        fph.CostPerSquareInch,
                        fph.WeightPerSquareInch
                    FROM FabricPricingHistory fph
                    WHERE fph.FK_SupplierProductNameDataId = s.PK_SupplierProductNameDataId
                    ORDER BY fph.DateFrom DESC
                ) fp
                WHERE s.FK_SupplierNameId = @SupplierId"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim row = dt.NewRow()
                        row("PK_FabricBrandNameId") = reader("PK_FabricBrandNameId")
                        row("BrandName") = If(IsDBNull(reader("BrandName")), "", reader("BrandName").ToString())
                        row("PK_FabricBrandProductNameId") = reader("PK_FabricBrandProductNameId")
                        row("BrandProductName") = reader("BrandProductName")
                        row("FK_ColorNameID") = reader("FK_ColorNameID") ' <-- ADD THIS
                        row("ColorNameFriendly") = reader("ColorNameFriendly")
                        row("FK_FabricTypeNameId") = If(IsDBNull(reader("FK_FabricTypeNameId")), DBNull.Value, reader("FK_FabricTypeNameId"))
                        row("FK_FabricTypeNameId") = Convert.ToInt32(row("FK_FabricTypeNameId"))
                        row("ShippingCost") = If(IsDBNull(reader("ShippingCost")), 0D, reader("ShippingCost"))
                        row("CostPerLinearYard") = If(IsDBNull(reader("CostPerLinearYard")), 0D, reader("CostPerLinearYard"))
                        row("CostPerSquareInch") = If(IsDBNull(reader("CostPerSquareInch")), 0D, reader("CostPerSquareInch"))
                        row("WeightPerSquareInch") = If(IsDBNull(reader("WeightPerSquareInch")), 0D, reader("WeightPerSquareInch"))
                        row("WeightPerLinearYard") = If(IsDBNull(reader("WeightPerLinearYard")), 0D, reader("WeightPerLinearYard"))
                        row("FabricRollWidth") = If(IsDBNull(reader("FabricRollWidth")), 0D, reader("FabricRollWidth"))
                        row("TotalYards") = If(IsDBNull(reader("TotalYards")), 0D, reader("TotalYards"))
                        dt.Rows.Add(row)
                    End While
                End Using
            End Using
        End Using

        dgvAssignFabrics.DataSource = Nothing
        dgvAssignFabrics.AutoGenerateColumns = False
        dgvAssignFabrics.DataSource = dt
    End Sub
    Private Sub FormatAssignFabricsGrid()
        ' Center header text and wrap
        For Each col As DataGridViewColumn In dgvAssignFabrics.Columns
            col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            col.HeaderCell.Style.WrapMode = DataGridViewTriState.True
        Next

        dgvAssignFabrics.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
        dgvAssignFabrics.ColumnHeadersHeight = 44 ' Adjust as needed

        ' Center cell text for all columns
        Dim centerStyle As New DataGridViewCellStyle()
        centerStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        For Each col As DataGridViewColumn In dgvAssignFabrics.Columns
            col.DefaultCellStyle = centerStyle
        Next

        ' Set specific column widths (adjust as needed)
        Dim widths As New Dictionary(Of String, Integer) From {
        {"BrandName", 120},
        {"ProductName", 120},
        {"ColorNameFriendly", 75},
        {"FabricType", 175},
        {"ShippingCost", 50},
        {"CostPerLinearYard", 70},
        {"CostPerSquareInch", 70},
        {"WeightPerSquareInch", 80},
        {"WeightPerLinearYard", 75},
        {"FabricRollWidth", 70},
        {"TotalYards", 70},
        {"Save", 50}
    }
        For Each kvp In widths
            If dgvAssignFabrics.Columns.Contains(kvp.Key) Then
                dgvAssignFabrics.Columns(kvp.Key).Width = kvp.Value
            End If
        Next
    End Sub
    Private Sub dgvAssignFabrics_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dgvAssignFabrics.DataError
        MessageBox.Show("Please enter a valid numeric value for this field.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        e.Cancel = False
    End Sub
    Private Sub ParseFabricRowValues(row As DataGridViewRow, ByRef shippingCost As Decimal, ByRef costPerLinearYard As Decimal, ByRef weightPerLinearYard As Decimal, ByRef fabricRollWidth As Decimal)
        ' Safely parse fabric row values, defaulting to 0D if parsing fails
        If Not Decimal.TryParse(row.Cells("ShippingCost").Value?.ToString(), shippingCost) Then
            shippingCost = 0D
        End If
        If Not Decimal.TryParse(row.Cells("CostPerLinearYard").Value?.ToString(), costPerLinearYard) Then
            costPerLinearYard = 0D
        End If
        If Not Decimal.TryParse(row.Cells("WeightPerLinearYard").Value?.ToString(), weightPerLinearYard) Then
            weightPerLinearYard = 0D
        End If
        If Not Decimal.TryParse(row.Cells("FabricRollWidth").Value?.ToString(), fabricRollWidth) Then
            fabricRollWidth = 0D
        End If
    End Sub
    Private Sub dgvAssignFabrics_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgvAssignFabrics.EditingControlShowing
        If dgvAssignFabrics.CurrentCell.ColumnIndex = dgvAssignFabrics.Columns("ProductName").Index Then
            Dim combo As ComboBox = TryCast(e.Control, ComboBox)
            If combo IsNot Nothing Then
                RemoveHandler combo.DropDown, AddressOf ProductNameDropDown
                AddHandler combo.DropDown, AddressOf ProductNameDropDown
            End If
        End If
    End Sub
    '*************************************************
    ' 3. Update dgvAssignFabrics_CellValueChanged to handle TotalYards
    Private Sub dgvAssignFabrics_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAssignFabrics.CellValueChanged
        If e.RowIndex < 0 Then Return

        Dim row = dgvAssignFabrics.Rows(e.RowIndex)
        Dim colName = dgvAssignFabrics.Columns(e.ColumnIndex).Name

        ' --- NEW: Handle IsActiveForMarketplace logic ---
        If colName = "IsActiveForMarketplace" Then
            Dim isActive As Boolean = CBool(row.Cells("IsActiveForMarketplace").Value)
            If isActive Then
                ' Get the supplier for this row (from the grid, not the combo)
                Dim activeSupplierId As Integer = CInt(row.Cells("PK_SupplierNameId").Value)
                ' Get the current combination from the form
                Dim brandId As Integer = CInt(cmbBrand.SelectedValue)
                Dim productId As Integer = CInt(cmbProduct.SelectedValue)
                Dim colorId As Integer = CInt(cmbColor.SelectedValue)
                Dim fabricTypeId As Integer = CInt(cmbFabricType.SelectedValue)

                ' Update DB: set all to inactive, then set this one to active
                Dim db As New DbConnectionManager()
                db.SetActiveForMarketplaceForCombination(brandId, productId, colorId, fabricTypeId, activeSupplierId)

                ' Uncheck all other rows in the grid
                For Each otherRow As DataGridViewRow In dgvAssignFabrics.Rows
                    If otherRow.Index <> e.RowIndex Then
                        otherRow.Cells("IsActiveForMarketplace").Value = False
                    End If
                Next
            End If
            ' No need to process further for this column
            Return
        End If

        ' --- EXISTING LOGIC: Pricing, product info, etc. ---
        Dim supplierId = cmbSupplier.SelectedValue
        Dim brandName = row.Cells("BrandName").Value?.ToString()
        Dim productName = row.Cells("ProductName").Value?.ToString()

        Dim db2 As New DbConnectionManager()
        Dim pricing As DataRow = Nothing
        Dim productInfo As DataRow = Nothing

        If Not String.IsNullOrEmpty(brandName) AndAlso Not String.IsNullOrEmpty(productName) AndAlso supplierId IsNot Nothing Then
            pricing = db2.GetFabricPricingHistory(supplierId, brandName, productName)
            productInfo = db2.GetFabricProductInfo(brandName, productName)
        End If

        ' Only act when ProductName changes
        If colName = "ProductName" Then
            If productInfo IsNot Nothing AndAlso Not IsDBNull(productInfo("FK_FabricTypeNameId")) Then
                Dim fabricTypeId = Convert.ToInt32(productInfo("FK_FabricTypeNameId"))
                Dim found = fabricTypes.AsEnumerable().Any(Function(r) r.Field(Of Integer)("PK_FabricTypeNameId") = fabricTypeId)
                If found Then
                    row.Cells("FabricType").Value = fabricTypeId
                Else
                    row.Cells("FabricType").Value = DBNull.Value
                End If
            Else
                row.Cells("FabricType").Value = DBNull.Value
            End If

            row.Cells("ShippingCost").Value = If(pricing IsNot Nothing AndAlso Not IsDBNull(pricing("ShippingCost")), pricing("ShippingCost"), 0D)
            row.Cells("CostPerLinearYard").Value = If(pricing IsNot Nothing AndAlso Not IsDBNull(pricing("CostPerLinearYard")), pricing("CostPerLinearYard"), 0D)
            row.Cells("CostPerSquareInch").Value = If(pricing IsNot Nothing AndAlso Not IsDBNull(pricing("CostPerSquareInch")), pricing("CostPerSquareInch"), 0D)
            row.Cells("WeightPerSquareInch").Value = If(pricing IsNot Nothing AndAlso Not IsDBNull(pricing("WeightPerSquareInch")), pricing("WeightPerSquareInch"), 0D)
            row.Cells("WeightPerLinearYard").Value = If(productInfo IsNot Nothing AndAlso Not IsDBNull(productInfo("WeightPerLinearYard")), productInfo("WeightPerLinearYard"), 0D)
            row.Cells("FabricRollWidth").Value = If(productInfo IsNot Nothing AndAlso Not IsDBNull(productInfo("FabricRollWidth")), productInfo("FabricRollWidth"), 0D)
            row.Cells("TotalYards").Value = If(pricing IsNot Nothing AndAlso Not IsDBNull(pricing.Table.Columns.Contains("TotalYards") AndAlso pricing("TotalYards")), pricing("TotalYards"), 0D)
        End If

        ' Recalculate and update dependent columns and database when relevant values change
        If colName = "CostPerLinearYard" OrElse colName = "WeightPerLinearYard" OrElse colName = "FabricRollWidth" OrElse colName = "TotalYards" Then
            Dim costPerLinearYard As Decimal = If(IsDBNull(row.Cells("CostPerLinearYard").Value), 0D, Convert.ToDecimal(row.Cells("CostPerLinearYard").Value))
            Dim weightPerLinearYard As Decimal = If(IsDBNull(row.Cells("WeightPerLinearYard").Value), 0D, Convert.ToDecimal(row.Cells("WeightPerLinearYard").Value))
            Dim fabricRollWidth As Decimal = If(IsDBNull(row.Cells("FabricRollWidth").Value), 0D, Convert.ToDecimal(row.Cells("FabricRollWidth").Value))
            Dim shippingCost As Decimal = If(IsDBNull(row.Cells("ShippingCost").Value), 0D, Convert.ToDecimal(row.Cells("ShippingCost").Value))
            Dim totalYards As Decimal = If(IsDBNull(row.Cells("TotalYards").Value), 0D, Convert.ToDecimal(row.Cells("TotalYards").Value))

            Dim totalCost As Decimal = (costPerLinearYard * totalYards) + shippingCost
            Dim totalCostPerLinearYard As Decimal = If(totalYards > 0, totalCost / totalYards, 0D)
            Dim squareInchesPerLinearYard As Decimal = fabricRollWidth * 36D
            Dim costPerSquareInch As Decimal = If(squareInchesPerLinearYard > 0, Math.Round(totalCostPerLinearYard / squareInchesPerLinearYard, 5), 0D)
            Dim weightPerSquareInch As Decimal = If(squareInchesPerLinearYard > 0, Math.Round(weightPerLinearYard / squareInchesPerLinearYard, 5), 0D)

            row.Cells("CostPerSquareInch").Value = costPerSquareInch
            row.Cells("WeightPerSquareInch").Value = weightPerSquareInch

            If supplierId IsNot Nothing AndAlso Not String.IsNullOrEmpty(productName) Then
                Dim productId As Integer = -1
                If Not IsDBNull(row.Cells("ProductName").Value) Then
                    productId = Convert.ToInt32(row.Cells("ProductName").Value)
                End If
                Dim colorId As Integer = CInt(row.Cells("FK_ColorNameID").Value)
                Dim fabricTypeId As Integer = CInt(row.Cells("FK_FabricTypeNameId").Value)
                Dim supplierProduct = db2.GetSupplierProductNameData(supplierId, productId, colorId, fabricTypeId)
                If supplierProduct IsNot Nothing Then
                    Dim supplierProductNameDataId = CInt(supplierProduct("PK_SupplierProductNameDataId"))
                    db2.InsertFabricPricingHistory(supplierProductNameDataId, shippingCost, costPerLinearYard, costPerSquareInch, weightPerSquareInch)
                    db2.UpdateFabricProductInfo(productId, weightPerLinearYard, fabricRollWidth)
                    db2.UpdateSupplierProductTotalYards(supplierProductNameDataId, totalYards)
                End If
            End If
        End If
        CheckAndDisplaySuppliersForCombination()
    End Sub

    Private Sub ProductNameDropDown(sender As Object, e As EventArgs)
        Dim rowIndex = dgvAssignFabrics.CurrentCell.RowIndex
        Dim brandName = dgvAssignFabrics.Rows(rowIndex).Cells("BrandName").Value?.ToString()
        If String.IsNullOrEmpty(brandName) Then Return

        Dim db As New DbConnectionManager()
        Dim products = db.GetProductsByBrandName(brandName) ' Should return DataTable with PK_FabricBrandProductNameId, BrandProductName
        Dim combo As ComboBox = CType(dgvAssignFabrics.EditingControl, ComboBox)
        combo.DataSource = products
        combo.DisplayMember = "BrandProductName"
        combo.ValueMember = "PK_FabricBrandProductNameId"
    End Sub

    Private Sub LoadAllBrands() '************REFACTOR TO DB MANAGER
        Dim db As New DbConnectionManager()
        Dim brands = db.GetAllFabricBrandNames() ' Should return a List(Of FabricBrandName) or DataTable
        cmbBrand.DataSource = brands
        cmbBrand.DisplayMember = "BrandName"
        cmbBrand.ValueMember = "PK_FabricBrandNameId"
        cmbBrand.SelectedIndex = -1
    End Sub
    ' 3. Load brands for a supplier into cmbBrand as DataTable
    Private Sub LoadBrandsForSupplier(supplierId As Integer, Optional selectBrandId As Integer = -1)
        Dim db As New DbConnectionManager()
        Dim brands As DataTable = db.GetBrandsForSupplier(supplierId) ' Must return DataTable
        cmbBrand.DataSource = brands
        cmbBrand.DisplayMember = "BrandName"
        cmbBrand.ValueMember = "PK_FabricBrandNameId"
        If selectBrandId <> -1 Then
            cmbBrand.SelectedValue = selectBrandId
        ElseIf brands.Rows.Count > 0 Then
            cmbBrand.SelectedIndex = 0
        End If
        cmbBrand.Refresh()
    End Sub

    Private Sub LoadProductWeightAndWidth(productId As Integer) '************REFACTOR TO DB MANAGER
        Dim db As New DbConnectionManager()
        Dim productInfo = db.GetFabricProductInfoById(productId) ' You may need to add this method
        If productInfo IsNot Nothing Then
            txtWeightPerLinearYard.Text = productInfo("WeightPerLinearYard").ToString()
            txtFabricRollWidth.Text = productInfo("FabricRollWidth").ToString()
        Else
            txtWeightPerLinearYard.Clear()
            txtFabricRollWidth.Clear()
        End If
    End Sub

    Private Sub cmbSupplier_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSupplier.SelectedIndexChanged
        If isFormLoading Then Return

        ' Update supplier fields enabled state when supplier changes
        Dim enableSupplierFields As Boolean =
        chkAssignToSupplier.Checked AndAlso
        cmbSupplier.SelectedValue IsNot Nothing AndAlso
        TypeOf cmbSupplier.SelectedValue Is Integer

        txtShippingCost.Enabled = enableSupplierFields
        txtCostPerLinearYard.Enabled = enableSupplierFields
        txtTotalYards.Enabled = enableSupplierFields

        If cmbSupplier.SelectedValue Is Nothing OrElse Not TypeOf cmbSupplier.SelectedValue Is Integer Then
            dgvAssignFabrics.DataSource = Nothing
            dgvAssignFabrics.Rows.Clear()
            ' Only clear combos if NOT assigning to supplier
            If Not chkAssignToSupplier.Checked Then
                cmbBrand.DataSource = Nothing
                cmbProduct.DataSource = Nothing
                cmbFabricType.DataSource = Nothing
            End If
            Return
        End If

        Dim supplierId As Integer = CInt(cmbSupplier.SelectedValue)

        ' 1. Initialize grid columns with only brands for this supplier
        InitializeAssignFabricsGrid(supplierId)

        ' 2. Load all supplier's products into the grid
        LoadSupplierProductsToGrid(supplierId)

        ' 3. Reset and reload main form brand/product combos
        If Not chkAssignToSupplier.Checked Then
            Dim db As New DbConnectionManager()
            Dim brands = db.GetBrandsForSupplier(supplierId)
            cmbBrand.DataSource = brands
            cmbBrand.DisplayMember = "BrandName"
            cmbBrand.ValueMember = "PK_FabricBrandNameId"
            cmbBrand.SelectedIndex = -1
            cmbProduct.DataSource = Nothing
            LoadFabricTypeCombo(-1)
            cmbFabricType.SelectedIndex = -1
        End If
    End Sub

    Private Sub cmbBrand_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBrand.SelectedIndexChanged
        If isFormLoading Then Return


        ' Autofill txtFabricBrandName with the selected brand's name
        If cmbBrand.SelectedItem IsNot Nothing Then
            If TypeOf cmbBrand.SelectedItem Is DataRowView Then
                txtFabricBrandName.Text = CType(cmbBrand.SelectedItem, DataRowView)("BrandName").ToString()
            ElseIf TypeOf cmbBrand.SelectedItem Is BrandDisplayItem Then
                txtFabricBrandName.Text = CType(cmbBrand.SelectedItem, BrandDisplayItem).BrandName
            Else
                txtFabricBrandName.Text = cmbBrand.Text
            End If
        Else
            txtFabricBrandName.Clear()
        End If

        suppressProductSelectionEvent = True

        cmbProduct.DataSource = Nothing
        LoadFabricTypeCombo(-1)
        cmbFabricType.SelectedIndex = -1

        ClearTextBoxes()

        If cmbBrand.SelectedValue Is Nothing OrElse Not TypeOf cmbBrand.SelectedValue Is Integer Then
            suppressProductSelectionEvent = False
            Return
        End If

        Dim brandId As Integer = CInt(cmbBrand.SelectedValue)

        ' --- Place your block here ---
        If cmbSupplier.SelectedValue Is Nothing OrElse Not TypeOf cmbSupplier.SelectedValue Is Integer Then
            ' MessageBox.Show("Calling LoadAllProductsForBrand")
            LoadAllProductsForBrand(brandId)
            suppressProductSelectionEvent = False
            Return
        End If

        ' --- The rest of your logic for when a supplier is selected ---
        If chkAssignToSupplier.Checked Then
            LoadAllProductsForBrand(brandId)
        Else
            Dim supplierId As Integer = CInt(cmbSupplier.SelectedValue)
            Dim db As New DbConnectionManager()
            Dim products As DataTable = db.GetProductsForSupplierAndBrand(supplierId, brandId)
            cmbProduct.DataSource = products
            cmbProduct.DisplayMember = "BrandProductName"
            cmbProduct.ValueMember = "PK_FabricBrandProductNameId"
        End If

        cmbProduct.SelectedIndex = -1
        suppressProductSelectionEvent = False
        CheckAndDisplaySuppliersForCombination()
    End Sub
    Private Sub cmbProduct_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbProduct.SelectedIndexChanged
        If suppressProductSelectionEvent Then Return

        ' Autofill product name
        If cmbProduct.SelectedItem IsNot Nothing Then
            If TypeOf cmbProduct.SelectedItem Is DataRowView Then
                txtFabricBrandProductName.Text = CType(cmbProduct.SelectedItem, DataRowView)("BrandProductName").ToString()
            ElseIf TypeOf cmbProduct.SelectedItem Is ProductDisplayItem Then
                txtFabricBrandProductName.Text = CType(cmbProduct.SelectedItem, ProductDisplayItem).BrandProductName
            Else
                txtFabricBrandProductName.Text = cmbProduct.Text
            End If
        Else
            txtFabricBrandProductName.Clear()
        End If

        ClearTextBoxes()

        If cmbProduct.SelectedValue Is Nothing OrElse Not TypeOf cmbProduct.SelectedValue Is Integer Then Return
        Dim productId As Integer = CInt(cmbProduct.SelectedValue)
        Dim db As New DbConnectionManager()

        ' 1. Get all fabric types and bind ComboBox
        Dim types As DataTable = db.GetAllFabricTypes()
        cmbFabricType.DataSource = types
        cmbFabricType.DisplayMember = "FabricType"
        cmbFabricType.ValueMember = "PK_FabricTypeNameId"

        ' 2. Get the fabric type for this product
        Dim fabricTypeId As Integer = -1
        Dim productInfo = db.GetFabricProductInfoById(productId)
        If productInfo IsNot Nothing AndAlso Not IsDBNull(productInfo("FK_FabricTypeNameId")) Then
            fabricTypeId = CInt(productInfo("FK_FabricTypeNameId"))
        End If

        ' 3. Only set SelectedValue if the value exists
        Dim found As Boolean = types.AsEnumerable().Any(Function(r) r.Field(Of Integer)("PK_FabricTypeNameId") = fabricTypeId)
        If found AndAlso fabricTypeId <> -1 Then
            cmbFabricType.SelectedValue = fabricTypeId
        Else
            cmbFabricType.SelectedIndex = -1
        End If

        ' 4. Fill product-specific fields (weight and roll width)
        LoadProductWeightAndWidth(productId)

        ' 5. Force UI update
        cmbFabricType.Refresh()
        cmbFabricType.Invalidate()

        ' 6. Debug output
        Debug.WriteLine("cmbFabricType.DataSource count: " & types.Rows.Count)
        Debug.WriteLine("Set fabricTypeId: " & fabricTypeId)
        Debug.WriteLine("cmbFabricType.SelectedValue: " & cmbFabricType.SelectedValue)
        Debug.WriteLine("cmbFabricType.SelectedIndex: " & cmbFabricType.SelectedIndex)
        Debug.WriteLine("cmbFabricType.Text: " & cmbFabricType.Text)
        CheckAndDisplaySuppliersForCombination()
    End Sub
    'Private Sub cmbProduct_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbProduct.SelectedIndexChanged
    '    If suppressProductSelectionEvent Then Return

    '    ' Autofill txtFabricBrandProductName with the selected product's name
    '    If cmbProduct.SelectedItem IsNot Nothing Then
    '        If TypeOf cmbProduct.SelectedItem Is DataRowView Then
    '            txtFabricBrandProductName.Text = CType(cmbProduct.SelectedItem, DataRowView)("BrandProductName").ToString()
    '        ElseIf TypeOf cmbProduct.SelectedItem Is ProductDisplayItem Then
    '            txtFabricBrandProductName.Text = CType(cmbProduct.SelectedItem, ProductDisplayItem).BrandProductName
    '        Else
    '            txtFabricBrandProductName.Text = cmbProduct.Text
    '        End If
    '    Else
    '        txtFabricBrandProductName.Clear()
    '    End If

    '    ClearTextBoxes()

    '    If cmbProduct.SelectedValue Is Nothing OrElse Not TypeOf cmbProduct.SelectedValue Is Integer Then Return
    '    Dim productId As Integer = CInt(cmbProduct.SelectedValue)
    '    Dim db As New DbConnectionManager()

    '    ' Always reload fabric types before setting SelectedValue
    '    Dim types As DataTable = db.GetAllFabricTypes()
    '    cmbFabricType.DataSource = types
    '    cmbFabricType.DisplayMember = "FabricType"
    '    cmbFabricType.ValueMember = "PK_FabricTypeNameId"

    '    ' --- DEBUG: Show all fabric types in Output window ---
    '    Debug.WriteLine("---- Fabric Types in Combo ----")
    '    For Each row As DataRow In types.Rows
    '        Debug.WriteLine("PK_FabricTypeNameId: " & row("PK_FabricTypeNameId").ToString() & " | FabricType: " & row("FabricType").ToString())
    '    Next
    '    Debug.WriteLine("ComboBox Items Count: " & cmbFabricType.Items.Count)
    '    Debug.WriteLine("SelectedValue (before): " & cmbFabricType.SelectedValue)
    '    Debug.WriteLine("SelectedIndex (before): " & cmbFabricType.SelectedIndex)
    '    Debug.WriteLine("Text (before): " & cmbFabricType.Text)

    '    ' Get the fabric type for this product
    '    Dim fabricTypeId As Integer = -1
    '    Dim productInfo = db.GetFabricProductInfoById(productId)
    '    If productInfo IsNot Nothing AndAlso Not IsDBNull(productInfo("FK_FabricTypeNameId")) Then
    '        fabricTypeId = CInt(productInfo("FK_FabricTypeNameId"))
    '    End If

    '    Debug.WriteLine("ProductId: " & productId & " | FK_FabricTypeNameId: " & fabricTypeId)

    '    ' Only set SelectedValue if the value exists in the DataTable
    '    Dim found As Boolean = types.AsEnumerable().Any(Function(r) r.Field(Of Integer)("PK_FabricTypeNameId") = fabricTypeId)
    '    If found AndAlso fabricTypeId <> -1 Then
    '        cmbFabricType.SelectedValue = fabricTypeId
    '        Debug.WriteLine("cmbFabricType.SelectedValue set to: " & fabricTypeId)
    '        Debug.WriteLine("AFTER SET SelectedValue: " & cmbFabricType.SelectedValue)
    '        Debug.WriteLine("AFTER SET SelectedIndex: " & cmbFabricType.SelectedIndex)
    '        Debug.WriteLine("AFTER SET Text: " & cmbFabricType.Text)
    '    Else
    '        cmbFabricType.SelectedIndex = -1
    '        Debug.WriteLine("Fabric type id not found in types table or invalid id.")
    '    End If

    '    cmbFabricType.Refresh()
    '    cmbFabricType.Invalidate()

    '    ' Clear supplier-specific fields ONLY
    '    txtTotalYards.Clear()
    '    txtShippingCost.Clear()
    '    txtCostPerLinearYard.Clear()
    '    txtCostPerSquareInch.Clear()
    '    ' DO NOT clear txtWeightPerSquareInch or txtSquareInchesPerLinearYard or txtFabricRollWidth

    '    If chkAssignToSupplier.Checked Then
    '        LoadProductWeightAndWidth(productId)
    '        If productInfo IsNot Nothing Then
    '            txtFabricBrandName.Text = productInfo("BrandName").ToString()
    '            txtFabricBrandProductName.Text = productInfo("BrandProductName").ToString()
    '        End If
    '        Return
    '    End If

    '    ' Not in assign mode: Load supplier-specific data as before
    '    Dim supplierProductNameDataId As Integer = -1
    '    Dim squareInchesPerLinearYard As Decimal = 0D
    '    Dim totalYards As Decimal = 0D

    '    If cmbSupplier.SelectedValue IsNot Nothing AndAlso TypeOf cmbSupplier.SelectedValue Is Integer Then
    '        Dim supplierId As Integer = CInt(cmbSupplier.SelectedValue)

    '        ' Use helper to get supplier-product data
    '        Dim supplierProduct = db.GetSupplierProductNameData(supplierId, productId)
    '        If supplierProduct IsNot Nothing Then
    '            If Not IsDBNull(supplierProduct("PK_SupplierProductNameDataId")) Then supplierProductNameDataId = CInt(supplierProduct("PK_SupplierProductNameDataId"))
    '            If Not IsDBNull(supplierProduct("FK_FabricTypeNameId")) Then fabricTypeId = CInt(supplierProduct("FK_FabricTypeNameId"))
    '            If Not IsDBNull(supplierProduct("SquareInchesPerLinearYard")) Then squareInchesPerLinearYard = CDec(supplierProduct("SquareInchesPerLinearYard"))
    '            If Not IsDBNull(supplierProduct("TotalYards")) Then totalYards = CDec(supplierProduct("TotalYards"))
    '            txtSquareInchesPerLinearYard.Text = squareInchesPerLinearYard.ToString()
    '            txtTotalYards.Text = totalYards.ToString()
    '            found = types.AsEnumerable().Any(Function(r) r.Field(Of Integer)("PK_FabricTypeNameId") = fabricTypeId)
    '            If found AndAlso fabricTypeId <> -1 Then
    '                cmbFabricType.SelectedValue = fabricTypeId
    '                Debug.WriteLine("cmbFabricType.SelectedValue set to (supplier): " & fabricTypeId)
    '                Debug.WriteLine("AFTER SET SelectedValue (supplier): " & cmbFabricType.SelectedValue)
    '                Debug.WriteLine("AFTER SET SelectedIndex (supplier): " & cmbFabricType.SelectedIndex)
    '                Debug.WriteLine("AFTER SET Text (supplier): " & cmbFabricType.Text)
    '            Else
    '                cmbFabricType.SelectedIndex = -1
    '                Debug.WriteLine("Fabric type id (supplier) not found in types table or invalid id.")
    '            End If
    '        Else
    '            cmbFabricType.SelectedIndex = -1
    '        End If

    '        ' Load weight and width from the product table
    '        Dim productInfo2 = db.GetFabricProductInfoById(productId)
    '        If productInfo2 IsNot Nothing Then
    '            txtWeightPerLinearYard.Text = productInfo2("WeightPerLinearYard").ToString()
    '            txtFabricRollWidth.Text = productInfo2("FabricRollWidth").ToString()
    '            txtFabricBrandName.Text = productInfo2("BrandName").ToString()
    '            txtFabricBrandProductName.Text = productInfo2("BrandProductName").ToString()
    '            Dim prodFabricTypeId As Integer = -1
    '            If Not IsDBNull(productInfo2("FK_FabricTypeNameId")) Then
    '                prodFabricTypeId = CInt(productInfo2("FK_FabricTypeNameId"))
    '            End If
    '            found = types.AsEnumerable().Any(Function(r) r.Field(Of Integer)("PK_FabricTypeNameId") = prodFabricTypeId)
    '            If found AndAlso prodFabricTypeId <> -1 Then
    '                cmbFabricType.SelectedValue = prodFabricTypeId
    '                Debug.WriteLine("cmbFabricType.SelectedValue set to (productInfo2): " & prodFabricTypeId)
    '                Debug.WriteLine("AFTER SET SelectedValue (productInfo2): " & cmbFabricType.SelectedValue)
    '                Debug.WriteLine("AFTER SET SelectedIndex (productInfo2): " & cmbFabricType.SelectedIndex)
    '                Debug.WriteLine("AFTER SET Text (productInfo2): " & cmbFabricType.Text)
    '            End If
    '            ' Calculate and fill weight per square inch
    '            Dim weight As Decimal = 0D
    '            Dim rollWidth As Decimal = 0D
    '            If Not IsDBNull(productInfo2("WeightPerLinearYard")) Then weight = CDec(productInfo2("WeightPerLinearYard"))
    '            If Not IsDBNull(productInfo2("FabricRollWidth")) Then rollWidth = CDec(productInfo2("FabricRollWidth"))
    '            Dim sqInches = rollWidth * 36D
    '            If sqInches > 0 Then
    '                txtWeightPerSquareInch.Text = Math.Round(weight / sqInches, 5).ToString()
    '            Else
    '                txtWeightPerSquareInch.Clear()
    '            End If
    '        End If

    '        ' Load latest pricing from FabricPricingHistory using helper
    '        Dim pricing = db.GetFabricPricingHistoryByProductId(supplierId, productId)
    '        If pricing IsNot Nothing Then
    '            txtShippingCost.Text = pricing("ShippingCost").ToString()
    '            txtCostPerLinearYard.Text = pricing("CostPerLinearYard").ToString()
    '            txtCostPerSquareInch.Text = pricing("CostPerSquareInch").ToString()
    '            txtWeightPerSquareInch.Text = pricing("WeightPerSquareInch").ToString()
    '        End If
    '    Else
    '        cmbFabricType.SelectedIndex = -1
    '    End If
    'End Sub

    Private Sub LoadFabricTypeCombo(selectedId As Integer) '************REFACTOR TO DB MANAGER
        Dim db As New DbConnectionManager()
        Dim types = db.GetAllFabricTypes()
        cmbFabricType.DataSource = types
        cmbFabricType.DisplayMember = "FabricType"
        cmbFabricType.ValueMember = "PK_FabricTypeNameId"
        cmbFabricType.SelectedValue = selectedId
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Dim assignToSupplier = chkAssignToSupplier.Checked

            ' --- Validate required fields ---
            If String.IsNullOrWhiteSpace(txtFabricBrandName.Text) Then
                MessageBox.Show("Please enter a brand name.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If String.IsNullOrWhiteSpace(txtFabricBrandProductName.Text) Then
                MessageBox.Show("Please enter a product name.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If cmbFabricType.SelectedValue Is Nothing OrElse Not TypeOf cmbFabricType.SelectedValue Is Integer Then
                MessageBox.Show("Please select a fabric type.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If cmbColor.SelectedValue Is Nothing OrElse Not TypeOf cmbColor.SelectedValue Is Integer Then
                MessageBox.Show("Please select a color.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If String.IsNullOrWhiteSpace(txtFabricRollWidth.Text) Then
                MessageBox.Show("Please enter the fabric roll width.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If String.IsNullOrWhiteSpace(txtWeightPerLinearYard.Text) Then
                MessageBox.Show("Please enter the weight per linear yard.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' --- Explicit numeric validation for product fields ---
            Dim WeightPerLinearYard As Decimal
            If Not Decimal.TryParse(txtWeightPerLinearYard.Text, WeightPerLinearYard) Then
                MessageBox.Show("Please enter a valid numeric value for Weight Per Linear Yard.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtWeightPerLinearYard.Focus()
                Return
            End If
            Dim FabricRollWidth As Decimal
            If Not Decimal.TryParse(txtFabricRollWidth.Text, FabricRollWidth) Then
                MessageBox.Show("Please enter a valid numeric value for Fabric Roll Width.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtFabricRollWidth.Focus()
                Return
            End If

            Dim fabricTypeId As Integer = CInt(cmbFabricType.SelectedValue)
            Dim colorId As Integer = CInt(cmbColor.SelectedValue)

            If assignToSupplier Then
                If cmbSupplier.SelectedValue Is Nothing OrElse Not TypeOf cmbSupplier.SelectedValue Is Integer Then
                    MessageBox.Show("Please select a supplier.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
                If String.IsNullOrWhiteSpace(txtShippingCost.Text) Then
                    MessageBox.Show("Please enter the shipping cost.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
                If String.IsNullOrWhiteSpace(txtCostPerLinearYard.Text) Then
                    MessageBox.Show("Please enter the cost per linear yard.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
                If String.IsNullOrWhiteSpace(txtTotalYards.Text) Then
                    MessageBox.Show("Please enter the total yards.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                ' --- Explicit numeric validation for supplier fields ---
                Dim shippingCost As Decimal
                If Not Decimal.TryParse(txtShippingCost.Text, shippingCost) Then
                    MessageBox.Show("Please enter a valid numeric value for Shipping Cost.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtShippingCost.Focus()
                    Return
                End If
                Dim costPerLinearYard As Decimal
                If Not Decimal.TryParse(txtCostPerLinearYard.Text, costPerLinearYard) Then
                    MessageBox.Show("Please enter a valid numeric value for Cost Per Linear Yard.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtCostPerLinearYard.Focus()
                    Return
                End If
                Dim totalYards As Decimal
                If Not Decimal.TryParse(txtTotalYards.Text, totalYards) Then
                    MessageBox.Show("Please enter a valid numeric value for Total Yards.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtTotalYards.Focus()
                    Return
                End If
            End If

            Dim brandName As String = txtFabricBrandName.Text.Trim()
            Dim productName As String = txtFabricBrandProductName.Text.Trim()
            Dim brandId As Integer
            Dim productId As Integer
            Dim joinId As Integer

            Using conn = DbConnectionManager.GetConnection()
                If conn.State <> ConnectionState.Open Then conn.Open()
                Using trans = conn.BeginTransaction()
                    Try
                        ' 1. Ensure brand exists
                        Using cmd = conn.CreateCommand()
                            cmd.Transaction = trans
                            cmd.CommandText = "SELECT PK_FabricBrandNameId FROM FabricBrandName WHERE LOWER(BrandName) = @BrandName"
                            cmd.Parameters.AddWithValue("@BrandName", brandName.ToLower())
                            Dim result = cmd.ExecuteScalar()
                            If result IsNot Nothing Then
                                brandId = CInt(result)
                            Else
                                cmd.CommandText = "INSERT INTO FabricBrandName (BrandName) VALUES (@BrandNameInsert); SELECT CAST(SCOPE_IDENTITY() AS int);"
                                cmd.Parameters.Clear()
                                cmd.Parameters.AddWithValue("@BrandNameInsert", brandName)
                                brandId = CInt(cmd.ExecuteScalar())
                            End If
                        End Using

                        ' 2. Ensure product exists for brand
                        Using cmd = conn.CreateCommand()
                            cmd.Transaction = trans
                            cmd.CommandText = "SELECT PK_FabricBrandProductNameId FROM FabricBrandProductName WHERE BrandProductName = @ProductName AND FK_FabricBrandNameId = @BrandId"
                            cmd.Parameters.AddWithValue("@ProductName", productName)
                            cmd.Parameters.AddWithValue("@BrandId", brandId)
                            Dim result = cmd.ExecuteScalar()
                            If result IsNot Nothing Then
                                productId = CInt(result)
                            Else
                                cmd.CommandText = "INSERT INTO FabricBrandProductName (BrandProductName, FK_FabricBrandNameId) VALUES (@ProductNameInsert, @BrandIdInsert); SELECT CAST(SCOPE_IDENTITY() AS int);"
                                cmd.Parameters.Clear()
                                cmd.Parameters.AddWithValue("@ProductNameInsert", productName)
                                cmd.Parameters.AddWithValue("@BrandIdInsert", brandId)
                                productId = CInt(cmd.ExecuteScalar())
                            End If
                        End Using

                        ' 3. Ensure JoinProductColorFabricType row exists
                        Using cmd = conn.CreateCommand()
                            cmd.Transaction = trans
                            cmd.CommandText = "SELECT PK_JoinProductColorFabricTypeId FROM JoinProductColorFabricType WHERE FK_FabricBrandProductNameId = @ProductId AND FK_ColorNameID = @ColorId AND FK_FabricTypeNameId = @FabricTypeId"
                            cmd.Parameters.AddWithValue("@ProductId", productId)
                            cmd.Parameters.AddWithValue("@ColorId", colorId)
                            cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                            Dim result = cmd.ExecuteScalar()
                            If result IsNot Nothing Then
                                joinId = CInt(result)
                            Else
                                cmd.CommandText = "INSERT INTO JoinProductColorFabricType (FK_FabricBrandProductNameId, FK_ColorNameID, FK_FabricTypeNameId) VALUES (@ProductId, @ColorId, @FabricTypeId); SELECT CAST(SCOPE_IDENTITY() AS int);"
                                cmd.Parameters.Clear()
                                cmd.Parameters.AddWithValue("@ProductId", productId)
                                cmd.Parameters.AddWithValue("@ColorId", colorId)
                                cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                                joinId = CInt(cmd.ExecuteScalar())
                            End If
                        End Using

                        If Not assignToSupplier Then
                            ' Only update base product info if NOT assigning to supplier
                            Using cmd = conn.CreateCommand()
                                cmd.Transaction = trans
                                cmd.CommandText = "UPDATE FabricBrandProductName SET WeightPerLinearYard = @Weight, FabricRollWidth = @RollWidth, FK_FabricTypeNameId = @FabricTypeId WHERE PK_FabricBrandProductNameId = @ProductId"
                                cmd.Parameters.AddWithValue("@Weight", WeightPerLinearYard)
                                cmd.Parameters.AddWithValue("@RollWidth", FabricRollWidth)
                                cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                                cmd.Parameters.AddWithValue("@ProductId", productId)
                                cmd.ExecuteNonQuery()
                            End Using
                            trans.Commit()
                            MessageBox.Show("Product information and color/fabric combination updated.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Return
                        End If

                        ' --- Supplier-specific logic ---
                        Dim supplierId As Integer = CInt(cmbSupplier.SelectedValue)
                        Dim shippingCost As Decimal = 0D, costPerLinearYard As Decimal = 0D, totalYards As Decimal = 0D
                        If Not Decimal.TryParse(txtShippingCost.Text, shippingCost) Then shippingCost = 0D
                        If Not Decimal.TryParse(txtCostPerLinearYard.Text, costPerLinearYard) Then costPerLinearYard = 0D
                        If Not Decimal.TryParse(txtTotalYards.Text, totalYards) Then totalYards = 0D

                        ' Calculated fields
                        Dim squareInchesPerLinearYard As Decimal = FabricRollWidth * 36D
                        Dim totalCost As Decimal = (costPerLinearYard * totalYards) + shippingCost
                        Dim totalCostPerLinearYard As Decimal = If(totalYards > 0, totalCost / totalYards, 0D)
                        Dim costPerSquareInch As Decimal = If(squareInchesPerLinearYard > 0, Math.Round(totalCostPerLinearYard / squareInchesPerLinearYard, 5), 0D)
                        Dim weightPerSquareInch As Decimal = If(squareInchesPerLinearYard > 0, Math.Round(WeightPerLinearYard / squareInchesPerLinearYard, 5), 0D)

                        ' Update calculated fields on the form
                        txtSquareInchesPerLinearYard.Text = squareInchesPerLinearYard.ToString("F2")
                        txtCostPerSquareInch.Text = costPerSquareInch.ToString("F5")
                        txtWeightPerSquareInch.Text = weightPerSquareInch.ToString("F5")

                        ' --- Enforce only one active per supplier/fabric type ---
                        Dim isActiveForMarketplace As Boolean = chkIsActiveForMarketPlace.Checked
                        Dim supplierProductId As Integer

                        If isActiveForMarketplace Then
                            ' Set all other products for this supplier/fabric type to inactive
                            Using cmd = conn.CreateCommand()
                                cmd.Transaction = trans
                                cmd.CommandText = "
                            UPDATE s
                            SET s.IsActiveForMarketplace = 0
                            FROM SupplierProductNameData s
                            INNER JOIN JoinProductColorFabricType j ON s.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                            WHERE s.FK_SupplierNameId = @SupplierId
                              AND j.FK_FabricTypeNameId = @FabricTypeId"
                                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                                cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If

                        ' 1. Insert or update SupplierProductNameData (now using joinId)
                        Using cmd = conn.CreateCommand()
                            cmd.Transaction = trans
                            cmd.CommandText = "SELECT PK_SupplierProductNameDataId FROM SupplierProductNameData WHERE FK_SupplierNameId = @SupplierId AND FK_JoinProductColorFabricTypeId = @JoinId"
                            cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                            cmd.Parameters.AddWithValue("@JoinId", joinId)
                            Dim result = cmd.ExecuteScalar()
                            If result IsNot Nothing Then
                                supplierProductId = CInt(result)
                                ' --- UPDATE: Only update supplier-specific fields ---
                                cmd.Parameters.Clear()
                                cmd.CommandText = "UPDATE SupplierProductNameData SET SquareInchesPerLinearYard = @SquareInches, TotalYards = @TotalYards, IsActiveForMarketplace = @IsActive WHERE PK_SupplierProductNameDataId = @SupplierProductId"
                                cmd.Parameters.AddWithValue("@SquareInches", squareInchesPerLinearYard)
                                cmd.Parameters.AddWithValue("@TotalYards", totalYards)
                                cmd.Parameters.AddWithValue("@IsActive", isActiveForMarketplace)
                                cmd.Parameters.AddWithValue("@SupplierProductId", supplierProductId)
                                cmd.ExecuteNonQuery()
                            Else
                                ' --- INSERT: Only insert supplier-specific fields ---
                                cmd.Parameters.Clear()
                                cmd.CommandText = "INSERT INTO SupplierProductNameData (FK_SupplierNameId, FK_JoinProductColorFabricTypeId, SquareInchesPerLinearYard, TotalYards, IsActiveForMarketplace) VALUES (@SupplierId, @JoinId, @SquareInches, @TotalYards, @IsActive); SELECT CAST(SCOPE_IDENTITY() AS int);"
                                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                                cmd.Parameters.AddWithValue("@JoinId", joinId)
                                cmd.Parameters.AddWithValue("@SquareInches", squareInchesPerLinearYard)
                                cmd.Parameters.AddWithValue("@TotalYards", totalYards)
                                cmd.Parameters.AddWithValue("@IsActive", isActiveForMarketplace)
                                supplierProductId = CInt(cmd.ExecuteScalar())
                            End If
                        End Using

                        ' 2. Insert new pricing history record for this supplier-product
                        Using cmd = conn.CreateCommand()
                            cmd.Transaction = trans
                            cmd.CommandText = "INSERT INTO FabricPricingHistory (FK_SupplierProductNameDataId, DateFrom, ShippingCost, CostPerLinearYard, CostPerSquareInch, WeightPerSquareInch, Quantity) VALUES (@SupplierProductId, @DateFrom, @ShippingCost, @CostPerLinearYard, @CostPerSquareInch, @WeightPerSquareInch, @Quantity)"
                            cmd.Parameters.AddWithValue("@SupplierProductId", supplierProductId)
                            cmd.Parameters.AddWithValue("@DateFrom", Date.Now)
                            cmd.Parameters.AddWithValue("@ShippingCost", shippingCost)
                            cmd.Parameters.AddWithValue("@CostPerLinearYard", costPerLinearYard)
                            cmd.Parameters.AddWithValue("@CostPerSquareInch", costPerSquareInch)
                            cmd.Parameters.AddWithValue("@WeightPerSquareInch", weightPerSquareInch)
                            cmd.Parameters.AddWithValue("@Quantity", totalYards)
                            cmd.ExecuteNonQuery()
                        End Using

                        trans.Commit()
                        MessageBox.Show("Supplier-specific fabric information saved.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ' Refresh the grid to show the new/updated fabric
                        If cmbSupplier.SelectedValue IsNot Nothing AndAlso TypeOf cmbSupplier.SelectedValue Is Integer Then
                            LoadSupplierProductsToGrid(CInt(cmbSupplier.SelectedValue))
                        End If
                    Catch exTrans As Exception
                        trans.Rollback()
                        Throw
                    End Try
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error saving data: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CheckAndDisplaySuppliersForCombination()
        If chkAssignToSupplier.Checked Then Return
        If cmbSupplier.SelectedValue IsNot Nothing AndAlso TypeOf cmbSupplier.SelectedValue Is Integer Then Return

        If cmbBrand.SelectedValue Is Nothing OrElse Not TypeOf cmbBrand.SelectedValue Is Integer Then Return
        If cmbProduct.SelectedValue Is Nothing OrElse Not TypeOf cmbProduct.SelectedValue Is Integer Then Return
        If cmbColor.SelectedValue Is Nothing OrElse Not TypeOf cmbColor.SelectedValue Is Integer Then Return
        If cmbFabricType.SelectedValue Is Nothing OrElse Not TypeOf cmbFabricType.SelectedValue Is Integer Then Return

        Dim brandId As Integer = CInt(cmbBrand.SelectedValue)
        Dim productId As Integer = CInt(cmbProduct.SelectedValue)
        Dim colorId As Integer = CInt(cmbColor.SelectedValue)
        Dim fabricTypeId As Integer = CInt(cmbFabricType.SelectedValue)

        Dim db As New DbConnectionManager()
        Dim suppliers As DataTable = db.GetSuppliersForCombination(brandId, productId, colorId, fabricTypeId)

        If suppliers.Rows.Count > 0 Then
            dgvAssignFabrics.DataSource = Nothing
            dgvAssignFabrics.Columns.Clear()
            dgvAssignFabrics.AutoGenerateColumns = False

            dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "PK_SupplierNameId",
            .HeaderText = "Supplier ID",
            .DataPropertyName = "PK_SupplierNameId",
            .ReadOnly = True
        })
            dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "CompanyName",
            .HeaderText = "Supplier Name",
            .DataPropertyName = "CompanyName",
            .ReadOnly = True
        })
            dgvAssignFabrics.Columns.Add(New DataGridViewCheckBoxColumn With {
            .Name = "IsActiveForMarketplace",
            .HeaderText = "Active for Marketplace",
            .DataPropertyName = "IsActiveForMarketplace",
            .ReadOnly = False ' <-- Make editable
        })
            dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "CostPerSquareInch",
            .HeaderText = "Cost/Square Inch",
            .DataPropertyName = "CostPerSquareInch",
            .ReadOnly = True,
            .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N5"}
        })

            dgvAssignFabrics.DataSource = suppliers
        Else
            dgvAssignFabrics.DataSource = Nothing
            dgvAssignFabrics.Rows.Clear()
            dgvAssignFabrics.Columns.Clear()
        End If
    End Sub




    'Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
    '    Try
    '        ' --- Validate required fields ---
    '        If cmbSupplier.SelectedValue Is Nothing OrElse Not TypeOf cmbSupplier.SelectedValue Is Integer Then
    '            MessageBox.Show("Please select a supplier.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '            Return
    '        End If
    '        If cmbFabricType.SelectedValue Is Nothing OrElse Not TypeOf cmbFabricType.SelectedValue Is Integer Then
    '            MessageBox.Show("Please select a fabric type.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '            Return
    '        End If

    '        Dim supplierId As Integer = CInt(cmbSupplier.SelectedValue)
    '        Dim fabricTypeId As Integer = CInt(cmbFabricType.SelectedValue)
    '        Dim isActive As Boolean = chkIsActiveForMarketPlace.Checked

    '        ' --- Brand ---
    '        Dim brandName As String = txtFabricBrandName.Text.Trim()
    '        If String.IsNullOrWhiteSpace(brandName) Then
    '            MessageBox.Show("Please enter a brand name.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '            Return
    '        End If

    '        ' --- Product ---
    '        Dim productName As String = txtFabricBrandProductName.Text.Trim()
    '        If String.IsNullOrWhiteSpace(productName) Then
    '            MessageBox.Show("Please enter a product name.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '            Return
    '        End If

    '        ' --- Parse numeric fields ---
    '        Dim weightPerLinearYard As Decimal
    '        Dim fabricRollWidth As Decimal
    '        Dim totalYards As Decimal
    '        Dim shippingCost As Decimal
    '        Dim costPerLinearYard As Decimal

    '        If Not Decimal.TryParse(txtWeightPerLinearYard.Text, weightPerLinearYard) Then weightPerLinearYard = 0D
    '        If Not Decimal.TryParse(txtFabricRollWidth.Text, fabricRollWidth) Then fabricRollWidth = 0D
    '        If Not Decimal.TryParse(txtTotalYards.Text, totalYards) Then totalYards = 0D
    '        If Not Decimal.TryParse(txtShippingCost.Text, shippingCost) Then shippingCost = 0D
    '        If Not Decimal.TryParse(txtCostPerLinearYard.Text, costPerLinearYard) Then costPerLinearYard = 0D

    '        ' --- Calculations ---
    '        Dim totalCost As Decimal = (costPerLinearYard * totalYards) + shippingCost
    '        Dim totalCostPerLinearYard As Decimal = If(totalYards > 0, totalCost / totalYards, 0D)
    '        Dim squareInchesPerLinearYard As Decimal = fabricRollWidth * 36D
    '        Dim costPerSquareInch As Decimal = If(squareInchesPerLinearYard > 0, Math.Round(totalCostPerLinearYard / squareInchesPerLinearYard, 5), 0D)
    '        Dim weightPerSquareInch As Decimal = If(squareInchesPerLinearYard > 0, Math.Round(weightPerLinearYard / squareInchesPerLinearYard, 5), 0D)

    '        ' --- Update calculated textboxes ---
    '        txtCostPerSquareInch.Text = costPerSquareInch.ToString("F5")
    '        txtWeightPerSquareInch.Text = weightPerSquareInch.ToString("F5")
    '        txtSquareInchesPerLinearYard.Text = squareInchesPerLinearYard.ToString("F2")

    '        ' --- DB Operations ---
    '        Dim brandId As Integer
    '        Dim productId As Integer
    '        Dim newSupplierProductId As Integer

    '        Using conn = DbConnectionManager.GetConnection()
    '            If conn.State <> ConnectionState.Open Then conn.Open()
    '            Using trans = conn.BeginTransaction()
    '                Try
    '                    ' 1. Ensure brand exists
    '                    Using cmd = conn.CreateCommand()
    '                        cmd.Transaction = trans
    '                        cmd.CommandText = "SELECT PK_FabricBrandNameId FROM FabricBrandName WHERE LOWER(BrandName) = @BrandName"
    '                        cmd.Parameters.AddWithValue("@BrandName", brandName.ToLower())
    '                        Dim result = cmd.ExecuteScalar()
    '                        If result IsNot Nothing Then
    '                            brandId = CInt(result)
    '                        Else
    '                            cmd.CommandText = "INSERT INTO FabricBrandName (BrandName) VALUES (@BrandNameInsert); SELECT CAST(SCOPE_IDENTITY() AS int);"
    '                            cmd.Parameters.Clear()
    '                            cmd.Parameters.AddWithValue("@BrandNameInsert", brandName)
    '                            brandId = CInt(cmd.ExecuteScalar())
    '                        End If
    '                    End Using

    '                    ' 2. Ensure product exists for brand
    '                    Using cmd = conn.CreateCommand()
    '                        cmd.Transaction = trans
    '                        cmd.CommandText = "SELECT PK_FabricBrandProductNameId FROM FabricBrandProductName WHERE BrandProductName = @ProductName AND FK_FabricBrandNameId = @BrandId"
    '                        cmd.Parameters.AddWithValue("@ProductName", productName)
    '                        cmd.Parameters.AddWithValue("@BrandId", brandId)
    '                        Dim result = cmd.ExecuteScalar()
    '                        If result IsNot Nothing Then
    '                            productId = CInt(result)
    '                        Else
    '                            cmd.CommandText = "INSERT INTO FabricBrandProductName (BrandProductName, FK_FabricBrandNameId) VALUES (@ProductNameInsert, @BrandIdInsert); SELECT CAST(SCOPE_IDENTITY() AS int);"
    '                            cmd.Parameters.Clear()
    '                            cmd.Parameters.AddWithValue("@ProductNameInsert", productName)
    '                            cmd.Parameters.AddWithValue("@BrandIdInsert", brandId)
    '                            productId = CInt(cmd.ExecuteScalar())
    '                        End If
    '                    End Using

    '                    ' 3. Unset IsActiveForMarketplace for all other records with this product if needed
    '                    If isActive Then
    '                        Using cmd = conn.CreateCommand()
    '                            cmd.Transaction = trans
    '                            cmd.CommandText = "UPDATE SupplierProductNameData SET IsActiveForMarketplace = 0 WHERE FK_FabricBrandProductNameId = @ProductId"
    '                            cmd.Parameters.AddWithValue("@ProductId", productId)
    '                            cmd.ExecuteNonQuery()
    '                        End Using
    '                    End If

    '                    ' 4. Insert into SupplierProductNameData
    '                    Using cmd = conn.CreateCommand()
    '                        cmd.Transaction = trans
    '                        cmd.CommandText =
    '                        "INSERT INTO SupplierProductNameData " &
    '                        "(FK_SupplierNameId, FK_FabricBrandProductNameId, FK_FabricTypeNameId, WeightPerLinearYard, SquareInchesPerLinearYard, FabricRollWidth, TotalYards, IsActiveForMarketplace) " &
    '                        "VALUES (@SupplierId, @ProductId, @FabricTypeId, @Weight, @SquareInches, @RollWidth, @TotalYards, @IsActive); " &
    '                        "SELECT CAST(SCOPE_IDENTITY() AS int);"
    '                        cmd.Parameters.AddWithValue("@SupplierId", supplierId)
    '                        cmd.Parameters.AddWithValue("@ProductId", productId)
    '                        cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
    '                        cmd.Parameters.AddWithValue("@Weight", weightPerLinearYard)
    '                        cmd.Parameters.AddWithValue("@SquareInches", squareInchesPerLinearYard)
    '                        cmd.Parameters.AddWithValue("@RollWidth", fabricRollWidth)
    '                        cmd.Parameters.AddWithValue("@TotalYards", totalYards)
    '                        cmd.Parameters.AddWithValue("@IsActive", isActive)
    '                        newSupplierProductId = CInt(cmd.ExecuteScalar())
    '                    End Using

    '                    ' 5. Insert into FabricPricingHistory
    '                    Using cmd = conn.CreateCommand()
    '                        cmd.Transaction = trans
    '                        cmd.CommandText =
    '                        "INSERT INTO FabricPricingHistory " &
    '                        "(FK_SupplierProductNameDataId, DateFrom, ShippingCost, CostPerLinearYard, CostPerSquareInch, WeightPerSquareInch, Quantity) " &
    '                        "VALUES (@SupplierProductNameDataId, @DateFrom, @ShippingCost, @CostPerLinearYard, @CostPerSquareInch, @WeightPerSquareInch, @Quantity)"
    '                        cmd.Parameters.AddWithValue("@SupplierProductNameDataId", newSupplierProductId)
    '                        cmd.Parameters.AddWithValue("@DateFrom", Date.Now)
    '                        cmd.Parameters.AddWithValue("@ShippingCost", shippingCost)
    '                        cmd.Parameters.AddWithValue("@CostPerLinearYard", costPerLinearYard)
    '                        cmd.Parameters.AddWithValue("@CostPerSquareInch", costPerSquareInch)
    '                        cmd.Parameters.AddWithValue("@WeightPerSquareInch", weightPerSquareInch)
    '                        cmd.Parameters.AddWithValue("@Quantity", totalYards)
    '                        cmd.ExecuteNonQuery()
    '                    End Using

    '                    trans.Commit()
    '                Catch exTrans As Exception
    '                    trans.Rollback()
    '                    Throw
    '                End Try
    '            End Using
    '        End Using

    '        MessageBox.Show("Saved and calculated values updated.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

    '        ' Optionally, reload brands/products for the supplier
    '        If cmbSupplier.SelectedValue IsNot Nothing AndAlso TypeOf cmbSupplier.SelectedValue Is Integer Then
    '            LoadBrandsForSupplier(CInt(cmbSupplier.SelectedValue), brandId)
    '        End If

    '    Catch ex As Exception
    '        MessageBox.Show("Error saving data: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

    Private Sub btnAddFabricType_Click(sender As Object, e As EventArgs) Handles btnAddFabricType.Click
        Dim newTypeName As String = txtAddFabricType.Text.Trim()
        Dim newAbbreviation As String = txtFabricTypeNameAbbreviation.Text.Trim()

        If String.IsNullOrWhiteSpace(newTypeName) Then
            MessageBox.Show("Please enter a fabric type name.")
            Return
        End If

        ' Check if the fabric type already exists (case-insensitive)
        Dim exists As Boolean = False
        Using conn = DbConnectionManager.GetConnection()
            If conn.State <> ConnectionState.Open Then
                conn.Open()
            End If
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT COUNT(*) FROM FabricTypeName WHERE LOWER(FabricType) = @TypeName"
                cmd.Parameters.AddWithValue("@TypeName", newTypeName.ToLower())
                exists = CInt(cmd.ExecuteScalar()) > 0
            End Using
        End Using

        If exists Then
            MessageBox.Show("This fabric type already exists.")
            Return
        End If

        ' Insert the new fabric type
        Using conn = DbConnectionManager.GetConnection()
            If conn.State <> ConnectionState.Open Then
                conn.Open()
            End If
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "INSERT INTO FabricTypeName (FabricType, FabricTypeNameAbbreviation) VALUES (@TypeName, @Abbreviation)"
                cmd.Parameters.AddWithValue("@TypeName", newTypeName)
                cmd.Parameters.AddWithValue("@Abbreviation", newAbbreviation)
                cmd.ExecuteNonQuery()
            End Using
        End Using

        txtAddFabricType.Clear()
        txtFabricTypeNameAbbreviation.Clear()
        MessageBox.Show("Fabric type added.")
    End Sub

    Private Sub chkAssignToSupplier_CheckedChanged(sender As Object, e As EventArgs) Handles chkAssignToSupplier.CheckedChanged
        Dim enableSupplierFields As Boolean =
        chkAssignToSupplier.Checked AndAlso
        cmbSupplier.SelectedValue IsNot Nothing AndAlso
        TypeOf cmbSupplier.SelectedValue Is Integer



        If chkAssignToSupplier.CheckState = CheckState.Unchecked Then
            ClearTextBoxes()

        Else
            txtShippingCost.ReadOnly = False
            txtCostPerLinearYard.ReadOnly = False
            txtTotalYards.ReadOnly = False
        End If


        'If chkAssignToSupplier.Checked = True Then
        '    txtShippingCost.ReadOnly = False
        '    txtCostPerLinearYard.ReadOnly = False
        '    txtTotalYards.ReadOnly = False
        'ElseIf chkAssignToSupplier.Checked = False Then
        '    txtShippingCost.ReadOnly = True
        '    txtCostPerLinearYard.ReadOnly = True
        '    txtTotalYards.ReadOnly = True
        '    ' Also set ReadOnly for extra safety
        '    txtShippingCost.ReadOnly = Not enableSupplierFields
        '    txtCostPerLinearYard.ReadOnly = Not enableSupplierFields
        '    txtTotalYards.ReadOnly = Not enableSupplierFields

        'End If


        'txtShippingCost.Enabled = enableSupplierFields
        'txtCostPerLinearYard.Enabled = enableSupplierFields
        'txtTotalYards.Enabled = enableSupplierFields



        'If Not enableSupplierFields Then
        '    txtShippingCost.Clear()

        '    txtCostPerLinearYard.Clear()
        '    txtTotalYards.Clear()
        'End If
    End Sub
    Private Sub ClearTextBoxes()
        txtTotalYards.Clear()
        txtShippingCost.Clear()
        txtCostPerLinearYard.Clear()
        txtCostPerSquareInch.Clear()
        ' Do NOT clear product-specific fields here!
    End Sub



    Private Function InsertSupplierProductNameData(
    supplierId As Integer,
    productId As Integer,
    fabricTypeId As Integer,
    weight As Decimal,
    squareInchesPerLinearYard As Decimal,
    rollWidth As Decimal,
    totalYards As Decimal,
    isActive As Boolean
) As Integer
        Dim newSupplierProductId As Integer
        Using conn = DbConnectionManager.GetConnection()
            If conn.State <> ConnectionState.Open Then
                conn.Open()
            End If
            Using cmd = conn.CreateCommand()
                cmd.CommandText =
                    "INSERT INTO SupplierProductNameData " &
                    "(FK_SupplierNameId, FK_FabricBrandProductNameId, FK_FabricTypeNameId, WeightPerLinearYard, SquareInchesPerLinearYard, FabricRollWidth, TotalYards, IsActiveForMarketplace) " &
                    "VALUES (@SupplierId, @ProductId, @FabricTypeId, @Weight, @SquareInches, @RollWidth, @TotalYards, @IsActive); " &
                    "SELECT CAST(SCOPE_IDENTITY() AS int);"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                cmd.Parameters.AddWithValue("@ProductId", productId)
                cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                cmd.Parameters.AddWithValue("@Weight", weight)
                cmd.Parameters.AddWithValue("@SquareInches", squareInchesPerLinearYard)
                cmd.Parameters.AddWithValue("@RollWidth", rollWidth)
                cmd.Parameters.AddWithValue("@TotalYards", totalYards)
                cmd.Parameters.AddWithValue("@IsActive", isActive)
                newSupplierProductId = CInt(cmd.ExecuteScalar())
            End Using
        End Using
        Return newSupplierProductId
    End Function

    Private Sub LoadAllBrandsWithSupplierMark(supplierId As Integer)
        Dim allBrands As New List(Of BrandDisplayItem)
        Dim supplierBrandIds As New HashSet(Of Integer)

        ' Get all brands associated with the supplier
        If supplierId > 0 Then
            Using conn = DbConnectionManager.GetConnection()
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT DISTINCT b.PK_FabricBrandNameId
                                   FROM SupplierProductNameData s
                                   INNER JOIN FabricBrandProductName p ON s.FK_FabricBrandProductNameId = p.PK_FabricBrandProductNameId
                                   INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                                   WHERE s.FK_SupplierNameId = @SupplierId"
                    cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            supplierBrandIds.Add(reader.GetInt32(0))
                        End While
                    End Using
                End Using
            End Using
        End If

        ' Get all brands
        Using conn = DbConnectionManager.GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_FabricBrandNameId, BrandName FROM FabricBrandName"
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim id = reader.GetInt32(0)
                        allBrands.Add(New BrandDisplayItem With {
                        .PK_FabricBrandNameId = id,
                        .BrandName = reader.GetString(1),
                        .IsSupplierBrand = supplierBrandIds.Contains(id)
                    })
                    End While
                End Using
            End Using
        End Using

        cmbBrand.DataSource = allBrands
        cmbBrand.DisplayMember = "DisplayText"
        cmbBrand.ValueMember = "PK_FabricBrandNameId"
        cmbBrand.SelectedIndex = -1
    End Sub

    Private Sub lblAddEditProductName_Click(sender As Object, e As EventArgs) Handles lblAddEditProductName.Click

    End Sub

    Private Sub txtShippingCost_TextChanged(sender As Object, e As EventArgs) Handles txtShippingCost.TextChanged

    End Sub

    Private Sub txtAddFabricType_TextChanged(sender As Object, e As EventArgs) Handles txtAddFabricType.TextChanged

    End Sub



    Private Sub btnClearSupplier_Click(sender As Object, e As EventArgs) Handles btnClearSupplier.Click
        isFormLoading = True

        ' Reset supplier combo and related controls
        cmbSupplier.SelectedIndex = -1
        cmbProduct.DataSource = Nothing
        cmbProduct.SelectedIndex = -1
        cmbFabricType.DataSource = Nothing
        cmbFabricType.SelectedIndex = -1

        cmbColor.SelectedIndex = -1

        ' Reload all brands into cmbBrand
        LoadAllBrandsToCombo()
        cmbBrand.SelectedIndex = -1

        ' Clear all textboxes
        txtFabricBrandName.Clear()
        txtFabricBrandProductName.Clear()
        txtWeightPerLinearYard.Clear()
        txtFabricRollWidth.Clear()
        txtSquareInchesPerLinearYard.Clear()
        txtWeightPerSquareInch.Clear()
        txtTotalYards.Clear()
        txtShippingCost.Clear()
        txtCostPerLinearYard.Clear()
        txtCostPerSquareInch.Clear()

        ' Reset checkboxes
        chkAssignToSupplier.Checked = False
        ' If you have chkIsActiveForMarketPlace or others, reset as needed:
        ' chkIsActiveForMarketPlace.Checked = False

        ' Reset DataGridView
        dgvAssignFabrics.DataSource = Nothing
        dgvAssignFabrics.Rows.Clear()
        dgvAssignFabrics.Columns.Clear()

        ' Restore read-only/enabled states as in form load
        txtSquareInchesPerLinearYard.ReadOnly = True
        txtCostPerSquareInch.ReadOnly = True
        txtWeightPerSquareInch.ReadOnly = True
        txtShippingCost.ReadOnly = True
        txtCostPerLinearYard.ReadOnly = True
        txtTotalYards.ReadOnly = True

        isFormLoading = False
    End Sub

    Private Sub cmbColor_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cmbColor.Validating
        Dim enteredColor As String = cmbColor.Text.Trim()
        If String.IsNullOrWhiteSpace(enteredColor) Then Return

        ' Check if color already exists in the ComboBox (by ColorNameFriendly)
        Dim exists As Boolean = False
        For Each item As DataRowView In cmbColor.Items
            If String.Equals(item("ColorNameFriendly").ToString(), enteredColor, StringComparison.OrdinalIgnoreCase) Then
                exists = True
                cmbColor.SelectedValue = item("PK_ColorNameID")
                Exit For
            End If
        Next

        If Not exists Then
            ' Only insert if it truly does not exist
            Dim db As New DbConnectionManager()
            Dim newId As Integer
            Using conn = DbConnectionManager.GetConnection()
                If conn.State <> ConnectionState.Open Then conn.Open()
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "INSERT INTO FabricColor (ColorName, ColorNameFriendly, ColorNameAbbreviation) VALUES (@ColorName, @ColorNameFriendly, @ColorNameAbbreviation); SELECT CAST(SCOPE_IDENTITY() AS int);"
                    cmd.Parameters.AddWithValue("@ColorName", enteredColor)
                    cmd.Parameters.AddWithValue("@ColorNameFriendly", enteredColor)
                    cmd.Parameters.AddWithValue("@ColorNameAbbreviation", enteredColor.Substring(0, Math.Min(20, enteredColor.Length)))
                    newId = CInt(cmd.ExecuteScalar())
                End Using
            End Using

            ' Add new color to ComboBox DataSource
            Dim dt = CType(cmbColor.DataSource, DataTable)
            Dim newRow = dt.NewRow()
            newRow("PK_ColorNameID") = newId
            newRow("ColorName") = enteredColor
            newRow("ColorNameFriendly") = enteredColor
            dt.Rows.Add(newRow)
            cmbColor.SelectedValue = newId
        End If
    End Sub

    Private Sub txtFabricRollWidth_TextChanged(sender As Object, e As EventArgs) Handles txtFabricRollWidth.TextChanged

    End Sub

    Private Sub lblFabricTypeNameAbbreviation_Click(sender As Object, e As EventArgs) Handles lblFabricTypeNameAbbreviation.Click

    End Sub
    Private Sub cmbColor_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbColor.SelectedIndexChanged
        CheckAndDisplaySuppliersForCombination()
    End Sub

    Private Sub cmbFabricType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFabricType.SelectedIndexChanged
        CheckAndDisplaySuppliersForCombination()
    End Sub
End Class
Public Class ProductDisplayItem
    Public Property PK_FabricBrandProductNameId As Integer
    Public Property BrandProductName As String
    Public Property IsSupplierProduct As Boolean
    Public ReadOnly Property DisplayText As String
        Get
            If IsSupplierProduct Then
                Return "★ " & BrandProductName
            Else
                Return BrandProductName
            End If
        End Get
    End Property
End Class