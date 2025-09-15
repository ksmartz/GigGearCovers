Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Linq
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

        ' Suppliers
        Dim suppliers = DbConnectionManager.GetAllSuppliers()
        cmbSupplier.DataSource = suppliers
        cmbSupplier.DisplayMember = "SupplierName"
        cmbSupplier.ValueMember = "SupplierID"
        cmbSupplier.SelectedIndex = -1

        ' Colors
        LoadAllColorsToCombo()

        ' Clear dependent combos
        cmbBrand.SelectedIndex = -1
        cmbProduct.DataSource = Nothing
        cmbFabricType.DataSource = Nothing

        ' Calculated fields are read-only
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
        Dim colors As DataTable = DbConnectionManager.GetAllMaterialColors()
        cmbColor.DataSource = colors
        cmbColor.DisplayMember = "ColorNameFriendly"
        cmbColor.ValueMember = "PK_ColorNameID"
        cmbColor.SelectedIndex = -1
    End Sub

    ' Add new brand if not exists (Validating)
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

    ' Load all products for a brand
    Private Sub LoadAllProductsForBrand(brandId As Integer)
        Dim products As DataTable = DbConnectionManager.GetProductsByBrandId(brandId)
        cmbProduct.DataSource = products
        cmbProduct.DisplayMember = "BrandProductName"
        cmbProduct.ValueMember = "PK_FabricBrandProductNameId"
        cmbProduct.SelectedIndex = -1
    End Sub

    ' Add new product if not exists (Validating)
    Private Sub cmbProduct_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cmbProduct.Validating
        Dim enteredProduct As String = cmbProduct.Text.Trim()
        If String.IsNullOrWhiteSpace(enteredProduct) Then
            txtFabricBrandProductName.Clear()
            Return
        End If

        ' Mirror combo text to textbox
        txtFabricBrandProductName.Text = enteredProduct

        ' If exists, select it
        For Each item As DataRowView In cmbProduct.Items
            If String.Equals(item("BrandProductName").ToString(), enteredProduct, StringComparison.OrdinalIgnoreCase) Then
                cmbProduct.SelectedValue = item("PK_FabricBrandProductNameId")
                Exit For
            End If
        Next
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
        If dgvAssignFabrics.Columns(e.ColumnIndex).Name <> "Save" Then Return

        Dim row = dgvAssignFabrics.Rows(e.RowIndex)
        If row.IsNewRow Then Return

        ' Detect changes since RowEnter
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
        If MessageBox.Show(msg, "Confirm Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return

        ' Lookup SPND
        Dim supplierId = CInt(cmbSupplier.SelectedValue)
        Dim productId = CInt(row.Cells("PK_FabricBrandProductNameId").Value) ' <-- ID comes from hidden column
        Dim colorId = CInt(row.Cells("FK_ColorNameID").Value)
        Dim fabricTypeId = CInt(row.Cells("FabricType").Value) ' <-- column Name is "FabricType"

        Dim supplierProduct = DbConnectionManager.GetSupplierProductNameData(supplierId, productId, colorId, fabricTypeId)
        If supplierProduct Is Nothing Then
            MessageBox.Show("Could not find supplier-product record.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim supplierProductNameDataId = CInt(supplierProduct("PK_SupplierProductNameDataId"))

        ' Parse numeric inputs
        Dim shippingCost = Convert.ToDecimal(row.Cells("ShippingCost").Value)
        Dim costPerLinearYard = Convert.ToDecimal(row.Cells("CostPerLinearYard").Value)
        Dim weightPerLinearYard = Convert.ToDecimal(row.Cells("WeightPerLinearYard").Value)
        Dim fabricRollWidth = Convert.ToDecimal(row.Cells("FabricRollWidth").Value)
        Dim totalYards = Convert.ToDecimal(row.Cells("TotalYards").Value)

        ' Recalculate dependent
        Dim totalCost As Decimal = (costPerLinearYard * totalYards) + shippingCost
        Dim totalCostPerLinearYard As Decimal = If(totalYards > 0, totalCost / totalYards, 0D)
        Dim squareInchesPerLinearYard As Decimal = fabricRollWidth * 36D
        Dim costPerSquareInch As Decimal = If(squareInchesPerLinearYard > 0, Math.Round(totalCostPerLinearYard / squareInchesPerLinearYard, 5), 0D)
        Dim weightPerSquareInch As Decimal = If(squareInchesPerLinearYard > 0, Math.Round(weightPerLinearYard / squareInchesPerLinearYard, 5), 0D)

        row.Cells("CostPerSquareInch").Value = costPerSquareInch
        row.Cells("WeightPerSquareInch").Value = weightPerSquareInch

        ' Save changes
        If changedFields.Contains("ShippingCost") OrElse changedFields.Contains("CostPerLinearYard") Then
            DbConnectionManager.InsertFabricPricingHistory(supplierProductNameDataId, shippingCost, costPerLinearYard, costPerSquareInch, weightPerSquareInch)
        End If
        If changedFields.Contains("WeightPerLinearYard") OrElse changedFields.Contains("FabricRollWidth") Then
            DbConnectionManager.UpdateFabricProductInfo(productId, weightPerLinearYard, fabricRollWidth)
        End If
        DbConnectionManager.UpdateSupplierProductTotalYards(supplierProductNameDataId, totalYards)

        MessageBox.Show("Changes saved.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
        originalRowValues(e.RowIndex) = New Dictionary(Of String, Object)(currentValues)
    End Sub

    Private Sub InitializeAssignFabricsGrid(supplierId As Integer)
        dgvAssignFabrics.Columns.Clear()

        ' Only brands linked to this supplier
        Dim brandNames = DbConnectionManager.GetBrandsForSupplier(supplierId)
        fabricTypes = DbConnectionManager.GetAllFabricTypes()

        ' Ensure PK column type is Int32 in the DataTable
        For Each row As DataRow In fabricTypes.Rows
            row("PK_FabricTypeNameId") = Convert.ToInt32(row("PK_FabricTypeNameId"))
        Next

        ' All products for supplier (used only for initial data shape)
        Dim allProducts = DbConnectionManager.GetProductsForSupplier(supplierId)

        ' Brand (read-only text)
        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "BrandName",
            .HeaderText = "Brand",
            .DataPropertyName = "BrandName",
            .ReadOnly = True
        })

        ' Product name (read-only text)
        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "ProductName",
            .HeaderText = "Fabric",
            .DataPropertyName = "BrandProductName",
            .ReadOnly = True
        })

        ' Color text
        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "ColorNameFriendly",
            .HeaderText = "Color",
            .DataPropertyName = "ColorNameFriendly",
            .ReadOnly = True
        })

        ' Hidden Color ID
        dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "FK_ColorNameID",
            .HeaderText = "Color ID",
            .DataPropertyName = "FK_ColorNameID",
            .Visible = False
        })

        ' Fabric Type (combobox bound to fabricTypes)
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

    ' Load all brands (DataTable) to cmbBrand
    Private Sub LoadAllBrandsToCombo()
        Dim brands As DataTable = DbConnectionManager.GetAllFabricBrandNames()
        cmbBrand.DataSource = brands
        cmbBrand.DisplayMember = "BrandName"
        cmbBrand.ValueMember = "PK_FabricBrandNameId"
        cmbBrand.SelectedIndex = -1
    End Sub

    ' Load supplier's products to grid (includes PK_FabricBrandProductNameId hidden column)
    Private Sub LoadSupplierProductsToGrid(supplierId As Integer)
        Dim dt As New DataTable()
        dt.Columns.Add("PK_FabricBrandNameId", GetType(Integer))
        dt.Columns.Add("BrandName", GetType(String))
        dt.Columns.Add("PK_FabricBrandProductNameId", GetType(Integer))   ' hidden column added to grid below
        dt.Columns.Add("BrandProductName", GetType(String))
        dt.Columns.Add("FK_ColorNameID", GetType(Integer))
        dt.Columns.Add("ColorNameFriendly", GetType(String))
        dt.Columns.Add("FK_FabricTypeNameId", GetType(Integer))
        dt.Columns.Add("ShippingCost", GetType(Decimal))
        dt.Columns.Add("CostPerLinearYard", GetType(Decimal))
        dt.Columns.Add("CostPerSquareInch", GetType(Decimal))
        dt.Columns.Add("WeightPerSquareInch", GetType(Decimal))
        dt.Columns.Add("WeightPerLinearYard", GetType(Decimal))
        dt.Columns.Add("FabricRollWidth", GetType(Decimal))
        dt.Columns.Add("TotalYards", GetType(Decimal))

        ' Ensure hidden ID column exists in grid (once)
        If Not dgvAssignFabrics.Columns.Contains("PK_FabricBrandProductNameId") Then
            dgvAssignFabrics.Columns.Add(New DataGridViewTextBoxColumn With {
                .Name = "PK_FabricBrandProductNameId",
                .HeaderText = "Product ID",
                .DataPropertyName = "PK_FabricBrandProductNameId",
                .Visible = False
            })
        End If

        Using conn = DbConnectionManager.GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText =
"SELECT 
    b.PK_FabricBrandNameId AS PK_FabricBrandNameId,
    b.BrandName,
    p.PK_FabricBrandProductNameId AS PK_FabricBrandProductNameId,
    p.BrandProductName AS BrandProductName,
    j.FK_ColorNameID,
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
                        row("FK_ColorNameID") = reader("FK_ColorNameID")
                        row("ColorNameFriendly") = reader("ColorNameFriendly")
                        row("FK_FabricTypeNameId") = If(IsDBNull(reader("FK_FabricTypeNameId")), DBNull.Value, reader("FK_FabricTypeNameId"))
                        If Not IsDBNull(row("FK_FabricTypeNameId")) Then
                            row("FK_FabricTypeNameId") = Convert.ToInt32(row("FK_FabricTypeNameId"))
                        End If
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
        For Each col As DataGridViewColumn In dgvAssignFabrics.Columns
            col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            col.HeaderCell.Style.WrapMode = DataGridViewTriState.True
            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next

        dgvAssignFabrics.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
        dgvAssignFabrics.ColumnHeadersHeight = 44

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
        If Not Decimal.TryParse(TryCast(row.Cells("ShippingCost").Value, String), shippingCost) Then shippingCost = 0D
        If Not Decimal.TryParse(TryCast(row.Cells("CostPerLinearYard").Value, String), costPerLinearYard) Then costPerLinearYard = 0D
        If Not Decimal.TryParse(TryCast(row.Cells("WeightPerLinearYard").Value, String), weightPerLinearYard) Then weightPerLinearYard = 0D
        If Not Decimal.TryParse(TryCast(row.Cells("FabricRollWidth").Value, String), fabricRollWidth) Then fabricRollWidth = 0D
    End Sub

    Private Sub dgvAssignFabrics_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgvAssignFabrics.EditingControlShowing
        ' Only attach when the editing control is a ComboBox (e.g., if you later switch ProductName back to a ComboBox column)
        If dgvAssignFabrics.CurrentCell Is Nothing Then Return
        If dgvAssignFabrics.Columns(dgvAssignFabrics.CurrentCell.ColumnIndex).Name <> "ProductName" Then Return
        If Not TypeOf e.Control Is ComboBox Then Return

        Dim combo As ComboBox = TryCast(e.Control, ComboBox)
        If combo IsNot Nothing Then
            RemoveHandler combo.DropDown, AddressOf ProductNameDropDown
            AddHandler combo.DropDown, AddressOf ProductNameDropDown
        End If
    End Sub

    Private Sub dgvAssignFabrics_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAssignFabrics.CellValueChanged
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return

        Dim row = dgvAssignFabrics.Rows(e.RowIndex)
        Dim colName = dgvAssignFabrics.Columns(e.ColumnIndex).Name

        ' Toggle Active for marketplace (when viewing suppliers list for a combination)
        If colName = "IsActiveForMarketplace" Then
            Dim isActive As Boolean = CBool(row.Cells("IsActiveForMarketplace").Value)
            If isActive Then
                Dim activeSupplierId As Integer = CInt(row.Cells("PK_SupplierNameId").Value)

                Dim brandId As Integer = CInt(cmbBrand.SelectedValue)
                Dim productId As Integer = CInt(cmbProduct.SelectedValue)
                Dim colorId As Integer = CInt(cmbColor.SelectedValue)
                Dim fabricTypeId As Integer = CInt(cmbFabricType.SelectedValue)

                DbConnectionManager.SetActiveForMarketplaceForCombination(brandId, productId, colorId, fabricTypeId, activeSupplierId)

                For Each otherRow As DataGridViewRow In dgvAssignFabrics.Rows
                    If otherRow.Index <> e.RowIndex Then
                        otherRow.Cells("IsActiveForMarketplace").Value = False
                    End If
                Next
            End If
            Return
        End If

        ' Load pricing/product info snapshot (best-effort)
        Dim supplierIdObj = cmbSupplier.SelectedValue
        Dim brandName = TryCast(row.Cells("BrandName").Value, String)
        Dim productName = TryCast(row.Cells("ProductName").Value, String)

        Dim pricing As DataRow = Nothing
        Dim productInfo As DataRow = Nothing

        If supplierIdObj IsNot Nothing AndAlso TypeOf supplierIdObj Is Integer AndAlso Not String.IsNullOrEmpty(brandName) AndAlso Not String.IsNullOrEmpty(productName) Then
            pricing = DbConnectionManager.GetFabricPricingHistory(CInt(supplierIdObj), brandName, productName)
            productInfo = DbConnectionManager.GetFabricProductInfo(brandName, productName)
        End If

        ' Only act when ProductName text changes (fill other fields if available)
        If colName = "ProductName" Then
            If productInfo IsNot Nothing AndAlso Not IsDBNull(productInfo("FK_FabricTypeNameId")) Then
                Dim fabricTypeId2 = Convert.ToInt32(productInfo("FK_FabricTypeNameId"))
                Dim found = fabricTypes.AsEnumerable().Any(Function(r) r.Field(Of Integer)("PK_FabricTypeNameId") = fabricTypeId2)
                If found Then
                    row.Cells("FabricType").Value = fabricTypeId2
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
            row.Cells("TotalYards").Value = If(pricing IsNot Nothing AndAlso pricing.Table.Columns.Contains("TotalYards") AndAlso Not IsDBNull(pricing("TotalYards")), pricing("TotalYards"), 0D)
        End If

        ' When any of these change, recompute and persist
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

            If supplierIdObj IsNot Nothing AndAlso TypeOf supplierIdObj Is Integer AndAlso Not String.IsNullOrEmpty(productName) Then
                ' Use the ID column for productId
                Dim productId As Integer = If(IsDBNull(row.Cells("PK_FabricBrandProductNameId").Value), -1, Convert.ToInt32(row.Cells("PK_FabricBrandProductNameId").Value))
                Dim colorId As Integer = CInt(row.Cells("FK_ColorNameID").Value)
                Dim fabricTypeId As Integer = CInt(row.Cells("FabricType").Value)

                Dim supplierProduct = DbConnectionManager.GetSupplierProductNameData(CInt(supplierIdObj), productId, colorId, fabricTypeId)
                If supplierProduct IsNot Nothing Then
                    Dim supplierProductNameDataId = CInt(supplierProduct("PK_SupplierProductNameDataId"))
                    DbConnectionManager.InsertFabricPricingHistory(supplierProductNameDataId, shippingCost, costPerLinearYard, costPerSquareInch, weightPerSquareInch)
                    DbConnectionManager.UpdateFabricProductInfo(productId, weightPerLinearYard, fabricRollWidth)
                    DbConnectionManager.UpdateSupplierProductTotalYards(supplierProductNameDataId, totalYards)
                End If
            End If
        End If

        CheckAndDisplaySuppliersForCombination()
    End Sub

    Private Sub ProductNameDropDown(sender As Object, e As EventArgs)
        Dim rowIndex = dgvAssignFabrics.CurrentCell.RowIndex
        Dim brandName = TryCast(dgvAssignFabrics.Rows(rowIndex).Cells("BrandName").Value, String)
        If String.IsNullOrEmpty(brandName) Then Return

        Dim products = DbConnectionManager.GetProductsByBrandName(brandName)
        Dim combo As ComboBox = TryCast(dgvAssignFabrics.EditingControl, ComboBox)
        If combo Is Nothing Then Return

        combo.DataSource = products
        combo.DisplayMember = "BrandProductName"
        combo.ValueMember = "PK_FabricBrandProductNameId"
    End Sub

    Private Sub LoadAllBrands() ' (unused helper – keep if you need it)
        Dim brands = DbConnectionManager.GetAllFabricBrandNames()
        cmbBrand.DataSource = brands
        cmbBrand.DisplayMember = "BrandName"
        cmbBrand.ValueMember = "PK_FabricBrandNameId"
        cmbBrand.SelectedIndex = -1
    End Sub

    Private Sub LoadBrandsForSupplier(supplierId As Integer, Optional selectBrandId As Integer = -1)
        Dim brands As DataTable = DbConnectionManager.GetBrandsForSupplier(supplierId)
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

    Private Sub LoadProductWeightAndWidth(productId As Integer)
        Dim productInfo = DbConnectionManager.GetFabricProductInfoById(productId)
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
            If Not chkAssignToSupplier.Checked Then
                cmbBrand.DataSource = Nothing
                cmbProduct.DataSource = Nothing
                cmbFabricType.DataSource = Nothing
            End If
            Return
        End If

        Dim supplierId As Integer = CInt(cmbSupplier.SelectedValue)

        InitializeAssignFabricsGrid(supplierId)
        LoadSupplierProductsToGrid(supplierId)

        If Not chkAssignToSupplier.Checked Then
            Dim brands = DbConnectionManager.GetBrandsForSupplier(supplierId)
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

        ' Autofill brand name textbox
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

        ' If no supplier picked yet, show all products for the brand
        If cmbSupplier.SelectedValue Is Nothing OrElse Not TypeOf cmbSupplier.SelectedValue Is Integer Then
            LoadAllProductsForBrand(brandId)
            suppressProductSelectionEvent = False
            Return
        End If

        ' Supplier picked
        If chkAssignToSupplier.Checked Then
            LoadAllProductsForBrand(brandId)
        Else
            Dim supplierId As Integer = CInt(cmbSupplier.SelectedValue)
            Dim products As DataTable = DbConnectionManager.GetProductsForSupplierAndBrand(supplierId, brandId)
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

        ' Mirror product name to textbox
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

        ' 1. Fabric types for combo
        Dim types As DataTable = DbConnectionManager.GetAllFabricTypes()
        cmbFabricType.DataSource = types
        cmbFabricType.DisplayMember = "FabricType"
        cmbFabricType.ValueMember = "PK_FabricTypeNameId"

        ' 2. Fabric type for this product
        Dim fabricTypeId As Integer = -1
        Dim productInfo = DbConnectionManager.GetFabricProductInfoById(productId)
        If productInfo IsNot Nothing AndAlso Not IsDBNull(productInfo("FK_FabricTypeNameId")) Then
            fabricTypeId = CInt(productInfo("FK_FabricTypeNameId"))
        End If

        Dim found As Boolean = types.AsEnumerable().Any(Function(r) r.Field(Of Integer)("PK_FabricTypeNameId") = fabricTypeId)
        If found AndAlso fabricTypeId <> -1 Then
            cmbFabricType.SelectedValue = fabricTypeId
        Else
            cmbFabricType.SelectedIndex = -1
        End If

        ' 4. Fill product-specific fields (weight and width)
        LoadProductWeightAndWidth(productId)

        ' 5. Force UI update
        cmbFabricType.Refresh()
        cmbFabricType.Invalidate()

        CheckAndDisplaySuppliersForCombination()
    End Sub

    Private Sub LoadFabricTypeCombo(selectedId As Integer)
        Dim types = DbConnectionManager.GetAllFabricTypes()
        cmbFabricType.DataSource = types
        cmbFabricType.DisplayMember = "FabricType"
        cmbFabricType.ValueMember = "PK_FabricTypeNameId"
        cmbFabricType.SelectedValue = selectedId
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Dim assignToSupplier = chkAssignToSupplier.Checked

            ' --- validate ---
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

            ' numeric validation
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
                        ' 1. Ensure brand
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

                        ' 2. Ensure product
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

                        ' 3. Ensure Join row
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
                            ' Update base product info only
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

                        ' Supplier-specific
                        Dim supplierId As Integer = CInt(cmbSupplier.SelectedValue)
                        Dim shippingCost As Decimal
                        Dim costPerLinearYard As Decimal
                        Dim totalYards As Decimal
                        Decimal.TryParse(txtShippingCost.Text, shippingCost)
                        Decimal.TryParse(txtCostPerLinearYard.Text, costPerLinearYard)
                        Decimal.TryParse(txtTotalYards.Text, totalYards)

                        Dim squareInchesPerLinearYard As Decimal = FabricRollWidth * 36D
                        Dim totalCost As Decimal = (costPerLinearYard * totalYards) + shippingCost
                        Dim totalCostPerLinearYard As Decimal = If(totalYards > 0, totalCost / totalYards, 0D)
                        Dim costPerSquareInch As Decimal = If(squareInchesPerLinearYard > 0, Math.Round(totalCostPerLinearYard / squareInchesPerLinearYard, 5), 0D)
                        Dim weightPerSquareInch As Decimal = If(squareInchesPerLinearYard > 0, Math.Round(WeightPerLinearYard / squareInchesPerLinearYard, 5), 0D)

                        txtSquareInchesPerLinearYard.Text = squareInchesPerLinearYard.ToString("F2")
                        txtCostPerSquareInch.Text = costPerSquareInch.ToString("F5")
                        txtWeightPerSquareInch.Text = weightPerSquareInch.ToString("F5")

                        Dim isActiveForMarketplace As Boolean = chkIsActiveForMarketPlace.Checked
                        Dim supplierProductId As Integer

                        If isActiveForMarketplace Then
                            Using cmd = conn.CreateCommand()
                                cmd.Transaction = trans
                                cmd.CommandText =
"UPDATE s
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

                        Using cmd = conn.CreateCommand()
                            cmd.Transaction = trans
                            cmd.CommandText = "SELECT PK_SupplierProductNameDataId FROM SupplierProductNameData WHERE FK_SupplierNameId = @SupplierId AND FK_JoinProductColorFabricTypeId = @JoinId"
                            cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                            cmd.Parameters.AddWithValue("@JoinId", joinId)
                            Dim result = cmd.ExecuteScalar()
                            If result IsNot Nothing Then
                                supplierProductId = CInt(result)
                                cmd.Parameters.Clear()
                                cmd.CommandText = "UPDATE SupplierProductNameData SET SquareInchesPerLinearYard = @SquareInches, TotalYards = @TotalYards, IsActiveForMarketplace = @IsActive WHERE PK_SupplierProductNameDataId = @SupplierProductId"
                                cmd.Parameters.AddWithValue("@SquareInches", squareInchesPerLinearYard)
                                cmd.Parameters.AddWithValue("@TotalYards", totalYards)
                                cmd.Parameters.AddWithValue("@IsActive", isActiveForMarketplace)
                                cmd.Parameters.AddWithValue("@SupplierProductId", supplierProductId)
                                cmd.ExecuteNonQuery()
                            Else
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

        Dim suppliers As DataTable = DbConnectionManager.GetSuppliersForCombination(brandId, productId, colorId, fabricTypeId)

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
                .ReadOnly = False
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

    Private Sub btnAddFabricType_Click(sender As Object, e As EventArgs) Handles btnAddFabricType.Click
        Dim newTypeName As String = txtAddFabricType.Text.Trim()
        Dim newAbbreviation As String = txtFabricTypeNameAbbreviation.Text.Trim()

        If String.IsNullOrWhiteSpace(newTypeName) Then
            MessageBox.Show("Please enter a fabric type name.")
            Return
        End If

        ' check exists
        Dim exists As Boolean = False
        Using conn = DbConnectionManager.GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
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

        ' insert
        Using conn = DbConnectionManager.GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
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
            If conn.State <> ConnectionState.Open Then conn.Open()
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

    ' trivial handlers kept (empty)
    Private Sub lblAddEditProductName_Click(sender As Object, e As EventArgs) Handles lblAddEditProductName.Click
    End Sub
    Private Sub txtShippingCost_TextChanged(sender As Object, e As EventArgs) Handles txtShippingCost.TextChanged
    End Sub
    Private Sub txtAddFabricType_TextChanged(sender As Object, e As EventArgs) Handles txtAddFabricType.TextChanged
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





' ===== helper view models you referenced =====
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


