'------------------------------------------------------------------------------
' Purpose: Handles saving a new supplier using DbConnectionManager for SQL connection.
' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager, System.Windows.Forms
' Current date: 2025-09-30
'------------------------------------------------------------------------------

Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Windows.Forms
Imports Data


Public Class FormMaterialsSuppliers
    Private isLoadingProducts As Boolean = False
    Private dgvContextMenu As ContextMenuStrip
    Private WithEvents mnuDeleteRow As ToolStripMenuItem
    Private selectedSupplierId As Integer? = Nothing
    Private selectedBrandId As Integer? = Nothing
    Private selectedProductType As String = Nothing
    ' Add this at the top of your Form class:

    '------------------------------------------------------------------------------
    ' Purpose: Loads all lookup data and initializes the DataGridView, including loading colors into cmbColor on form load.
    ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager, System.Windows.Forms
    ' Current date: 2025-10-03
    '------------------------------------------------------------------------------
    Private Sub FormMaterialsSuppliers_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Lookup data
        LoadUnitTypes()
        LoadProductTypes()
        LoadBrands()
        BindSuppliers()           ' <-- use SupplierInformation, not old LoadSuppliers()
        LoadAllProductsToSelection()
        LoadColors()

        ' UI setup
        cmbProductType.DrawMode = DrawMode.OwnerDrawFixed
        RemoveHandler cmbProductType.DrawItem, AddressOf cmbProductType_DrawItem
        AddHandler cmbProductType.DrawItem, AddressOf cmbProductType_DrawItem

        InitializeDgvContextMenu()

        ' Start with empty grid; columns will be built after first data bind
        dgvSelectedProducts.VirtualMode = False
        dgvSelectedProducts.DataSource = Nothing
        dgvSelectedProducts.Columns.Clear()
        dgvSelectedProducts.AutoGenerateColumns = False

        ' Wire selection changed (avoid double-wiring)
        RemoveHandler cmbSupplier.SelectedIndexChanged, AddressOf cmbSupplier_SelectedIndexChanged
        AddHandler cmbSupplier.SelectedIndexChanged, AddressOf cmbSupplier_SelectedIndexChanged

        ' Optional: also wire DataError once here
        RemoveHandler dgvSelectedProducts.DataError, AddressOf dgvSelectedProducts_DataError
        AddHandler dgvSelectedProducts.DataError, AddressOf dgvSelectedProducts_DataError
    End Sub

    Private Sub BindSuppliers()
        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                Dim dt As New DataTable()
                Using cmd As New SqlCommand("
                SELECT PK_SupplierNameId, CompanyName
                FROM SupplierInformation
                WHERE ISNULL(CompanyName,'') <> ''
                ORDER BY CompanyName;", conn)
                    Using da As New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                cmbSupplier.DisplayMember = "CompanyName"
                cmbSupplier.ValueMember = "PK_SupplierNameId"
                cmbSupplier.DataSource = dt
                cmbSupplier.DropDownStyle = ComboBoxStyle.DropDownList
                cmbSupplier.SelectedIndex = -1
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading suppliers: " & ex.Message)
        End Try
    End Sub

    '------------------------------------------------------------------------------
    ' Purpose: When a supplier is selected, loads all associated brands, product types, and product names into the DataGridView.
    ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager, System.Windows.Forms
    ' Current date: 2025-10-03
    '------------------------------------------------------------------------------
    ' Class-level if you use it elsewhere:
    ' Private selectedSupplierId As Integer?

    Private Sub cmbSupplier_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSupplier.SelectedIndexChanged
        If cmbSupplier.SelectedIndex >= 0 AndAlso cmbSupplier.SelectedValue IsNot Nothing AndAlso IsNumeric(cmbSupplier.SelectedValue) Then
            selectedSupplierId = CInt(cmbSupplier.SelectedValue)   ' this is PK_SupplierNameId
        Else
            selectedSupplierId = Nothing
        End If

        ' This binds the grid and THEN calls InitializeDgvSelectedProducts internally
        LoadFilteredSupplierProducts()
    End Sub


    'Call this from your Form's Load event (after InitializeComponent):
    Private Sub InitializeDgvContextMenu()
        dgvContextMenu = New ContextMenuStrip()
        mnuDeleteRow = New ToolStripMenuItem("Delete Row")
        dgvContextMenu.Items.Add(mnuDeleteRow)
        dgvSelectedProducts.ContextMenuStrip = dgvContextMenu

    End Sub
    ' Handles saving a new supplier, using DbConnectionManager for all DB access.
    Private Sub btnSaveSupplier_Click(sender As Object, e As EventArgs) Handles btnSaveSupplier.Click
        ' >>> changed
        ' Validate required fields
        If String.IsNullOrWhiteSpace(txtCompanyName.Text) Then
            MessageBox.Show("Company Name is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCompanyName.Focus()
            Exit Sub
        End If
        If String.IsNullOrWhiteSpace(txtWebsite1.Text) Then
            MessageBox.Show("Website is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtWebsite1.Focus()
            Exit Sub
        End If

        ' Check for duplicate company name
        Dim exists As Boolean = False
        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                Using cmd As New SqlCommand("SELECT COUNT(*) FROM SupplierInformation WHERE CompanyName = @CompanyName", conn)
                    cmd.Parameters.AddWithValue("@CompanyName", txtCompanyName.Text.Trim())
                    exists = CInt(cmd.ExecuteScalar()) > 0
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error checking for duplicate company: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try

        If exists Then
            MessageBox.Show("A supplier with this company name already exists.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCompanyName.Focus()
            Exit Sub
        End If

        ' Insert new supplier
        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                Using cmd As New SqlCommand("
                    INSERT INTO SupplierInformation
                    (CompanyName, Contact1FirstName, Contact1LastName, Contact1Phone1, Contact1Phone2, Contact1Email1, Contact1Email2, Website1,
                     Contact2FirstName, Contact2LastName, Contact2Phone1, Contact2Phone2, Contact2Email1, Contact2Email2,
                     Address11, Address21, City1, State1, ZipPostal1,
                     Address12, Address22, City2, State2, ZipPostal2)
                    VALUES
                    (@CompanyName, @Contact1FirstName, @Contact1LastName, @Contact1Phone1, @Contact1Phone2, @Contact1Email1, @Contact1Email2, @Website1,
                     @Contact2FirstName, @Contact2LastName, @Contact2Phone1, @Contact2Phone2, @Contact2Email1, @Contact2Email2,
                     @Address11, @Address21, @City1, @State1, @ZipPostal1,
                     @Address12, @Address22, @City2, @State2, @ZipPostal2)
                ", conn)
                    ' Map controls to parameters
                    cmd.Parameters.AddWithValue("@CompanyName", txtCompanyName.Text.Trim())
                    cmd.Parameters.AddWithValue("@Contact1FirstName", txtContact1FirstName.Text.Trim())
                    cmd.Parameters.AddWithValue("@Contact1LastName", txtContact1LastName.Text.Trim())
                    cmd.Parameters.AddWithValue("@Contact1Phone1", txtContact1Phone1.Text.Trim())
                    cmd.Parameters.AddWithValue("@Contact1Phone2", txtContact1Phone2.Text.Trim())
                    cmd.Parameters.AddWithValue("@Contact1Email1", txtContact1Email1.Text.Trim())
                    cmd.Parameters.AddWithValue("@Contact1Email2", txtContact1Email2.Text.Trim())
                    cmd.Parameters.AddWithValue("@Website1", txtWebsite1.Text.Trim())
                    cmd.Parameters.AddWithValue("@Contact2FirstName", txtContact2FirstName.Text.Trim())
                    cmd.Parameters.AddWithValue("@Contact2LastName", txtContact2LastName.Text.Trim())
                    cmd.Parameters.AddWithValue("@Contact2Phone1", txtContact2Phone1.Text.Trim())
                    cmd.Parameters.AddWithValue("@Contact2Phone2", txtContact2Phone2.Text.Trim())
                    cmd.Parameters.AddWithValue("@Contact2Email1", txtContact2Email1.Text.Trim())
                    cmd.Parameters.AddWithValue("@Contact2Email2", txtContact2Email2.Text.Trim())
                    cmd.Parameters.AddWithValue("@Address11", txtAddress1.Text.Trim())
                    cmd.Parameters.AddWithValue("@Address21", txtAddress2.Text.Trim())
                    cmd.Parameters.AddWithValue("@City1", txtCity1.Text.Trim())
                    cmd.Parameters.AddWithValue("@State1", txtState1.Text.Trim())
                    cmd.Parameters.AddWithValue("@ZipPostal1", txtZipPostal1.Text.Trim())
                    cmd.Parameters.AddWithValue("@Address12", txtAddress12.Text.Trim())
                    cmd.Parameters.AddWithValue("@Address22", txtAddress22.Text.Trim())
                    cmd.Parameters.AddWithValue("@City2", txtCity2.Text.Trim())
                    cmd.Parameters.AddWithValue("@State2", txtState2.Text.Trim())
                    cmd.Parameters.AddWithValue("@ZipPostal2", txtZipPostal2.Text.Trim())

                    cmd.ExecuteNonQuery()
                End Using
            End Using
            MessageBox.Show("Supplier saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ' Optionally clear fields or switch tab here
        Catch ex As Exception
            MessageBox.Show("Error saving supplier: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    '------------------------------------------------------------------------------
    ' Purpose: Handles new brand entry in cmbBrand. If user types a new brand and leaves the ComboBox,
    '          prompts to add it to ProductBrands, saves if confirmed, and selects it.
    ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager, System.Windows.Forms
    ' Current date: 2025-10-03
    '------------------------------------------------------------------------------

    Private Sub cmbBrand_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cmbBrand.Validating
        ' >>> changed
        ' Only proceed if the text is not empty and not already in the list
        Dim newBrandName As String = cmbBrand.Text.Trim()
        If newBrandName = "" Then Exit Sub

        ' Check if the brand already exists in the ComboBox list (case-insensitive)
        Dim existsInList As Boolean = False
        For Each item As DataRowView In cmbBrand.Items
            If String.Equals(item("BrandName").ToString(), newBrandName, StringComparison.OrdinalIgnoreCase) Then
                existsInList = True
                Exit For
            End If
        Next
        If existsInList Then Exit Sub

        ' Prompt user to add new brand
        Dim result = MessageBox.Show($"New brand name: {newBrandName}.{Environment.NewLine}Press Yes to continue or No to exit.", "Add New Brand", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Try
                Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                    ' Check if brand already exists in DB (case-insensitive)
                    Using checkCmd As New SqlCommand("SELECT BrandID FROM ProductBrands WHERE LOWER(BrandName) = LOWER(@BrandName)", conn)
                        checkCmd.Parameters.AddWithValue("@BrandName", newBrandName)
                        Dim dbResult = checkCmd.ExecuteScalar()
                        If dbResult IsNot Nothing Then
                            MessageBox.Show("This brand already exists.", "Duplicate Brand", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            cmbBrand.SelectedValue = CInt(dbResult)
                            Exit Sub
                        End If
                    End Using
                    ' Insert new brand
                    Using insertCmd As New SqlCommand("INSERT INTO ProductBrands (BrandName) OUTPUT INSERTED.BrandID VALUES (@BrandName)", conn)
                        insertCmd.Parameters.AddWithValue("@BrandName", newBrandName)
                        Dim newBrandId As Integer = CInt(insertCmd.ExecuteScalar())
                        MessageBox.Show($"Brand '{newBrandName}' added successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ' Reload brands and select the new one
                        LoadBrands()
                        cmbBrand.SelectedValue = newBrandId
                        ' Optionally, trigger product name reload
                        If cmbProductType.Text <> "" Then
                            LoadProductNames(newBrandId, cmbProductType.Text)
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Error adding new brand: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            ' User cancelled; reset ComboBox
            LoadBrands()
            cmbBrand.SelectedIndex = -1
            e.Cancel = True
        End If
        ' <<< end changed
    End Sub
    ' Purpose: Handles brand selection or entry. If a new brand is typed, prompts to add it to ProductBrands, saves if confirmed, and selects it.
    ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager, System.Windows.Forms
    ' Current date: 2025-10-03
    Private Sub cmbBrand_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBrand.SelectedIndexChanged
        ' >>> changed

        ' Always reload product types for the selected brand
        LoadProductTypes() ' <<< THIS LINE ADDED
        ' Check if the user typed a new brand (not in the list)
        If cmbBrand.Focused AndAlso cmbBrand.SelectedIndex = -1 AndAlso Not String.IsNullOrWhiteSpace(cmbBrand.Text) Then
            Dim newBrandName As String = cmbBrand.Text.Trim()
            Dim result = MessageBox.Show($"New brand name: {newBrandName}.{Environment.NewLine}Press Yes to continue or No to exit.", "Add New Brand", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Try
                    Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                        ' Check if brand already exists (case-insensitive)
                        Using checkCmd As New SqlCommand("SELECT COUNT(*) FROM ProductBrands WHERE LOWER(BrandName) = LOWER(@BrandName)", conn)
                            checkCmd.Parameters.AddWithValue("@BrandName", newBrandName)
                            Dim exists As Boolean = CInt(checkCmd.ExecuteScalar()) > 0
                            If exists Then
                                MessageBox.Show("This brand already exists.", "Duplicate Brand", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Else
                                ' Insert new brand
                                Using insertCmd As New SqlCommand("INSERT INTO ProductBrands (BrandName) OUTPUT INSERTED.BrandID VALUES (@BrandName)", conn)
                                    insertCmd.Parameters.AddWithValue("@BrandName", newBrandName)
                                    Dim newBrandId As Integer = CInt(insertCmd.ExecuteScalar())
                                    MessageBox.Show($"Brand '{newBrandName}' added successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                    ' Reload brands and select the new one
                                    LoadBrands()
                                    cmbBrand.SelectedValue = newBrandId
                                    ' Optionally, trigger product name reload
                                    If cmbProductType.Text <> "" Then
                                        LoadProductNames(newBrandId, cmbProductType.Text)
                                    End If
                                    Exit Sub
                                End Using
                            End If
                        End Using
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error adding new brand: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Else
                ' User cancelled; reset ComboBox
                LoadBrands()
                cmbBrand.SelectedIndex = -1
                Exit Sub
            End If
        End If

        ' Existing logic for loading product names
        If cmbBrand.SelectedValue IsNot Nothing AndAlso cmbProductType.Text <> "" Then
            Dim brandId As Integer
            If TypeOf cmbBrand.SelectedValue Is DataRowView Then
                brandId = CInt(DirectCast(cmbBrand.SelectedValue, DataRowView)("BrandID"))
            Else
                brandId = CInt(cmbBrand.SelectedValue)
            End If
            LoadProductNames(brandId, cmbProductType.Text)
        End If
        ' <<< end changed
    End Sub

    '------------------------------------------------------------------------------
    ' Purpose: Custom-draws cmbProductType items, bolding those associated with the selected brand and drawing a separator line.
    ' Dependencies: Imports System.Drawing, System.Windows.Forms
    ' Current date: 2025-10-03
    '------------------------------------------------------------------------------
    Private Sub cmbProductType_DrawItem(sender As Object, e As DrawItemEventArgs) Handles cmbProductType.DrawItem
        ' >>> changed
        ' Defensive: skip if no items
        If e.Index < 0 OrElse cmbProductType.Items.Count = 0 Then
            Return
        End If

        ' Get DataRowView for this item
        Dim drv As DataRowView = TryCast(cmbProductType.Items(e.Index), DataRowView)
        If drv Is Nothing Then Return

        Dim text As String = drv("ProductTypeName").ToString()
        Dim sortOrder As Integer = 1
        If drv.Row.Table.Columns.Contains("SortOrder") Then
            sortOrder = Convert.ToInt32(drv("SortOrder"))
        End If

        ' Determine if this is the first "not associated" item for separator
        Dim drawSeparator As Boolean = False
        If e.Index > 0 Then
            Dim prevDrv As DataRowView = TryCast(cmbProductType.Items(e.Index - 1), DataRowView)
            If prevDrv IsNot Nothing AndAlso Convert.ToInt32(prevDrv("SortOrder")) = 0 AndAlso sortOrder = 1 Then
                drawSeparator = True
            End If
        End If

        ' Draw background
        e.DrawBackground()

        ' Choose font: bold for associated, regular for others
        Dim fontToUse As Font = If(sortOrder = 0, New Font(e.Font, FontStyle.Bold), e.Font)
        Dim textColor As Brush = If((e.State And DrawItemState.Selected) = DrawItemState.Selected, Brushes.White, Brushes.Black)

        ' Draw text
        e.Graphics.DrawString(text, fontToUse, textColor, e.Bounds.Left + 2, e.Bounds.Top + 2)

        ' Draw separator line if needed
        If drawSeparator Then
            Dim y As Integer = e.Bounds.Top
            Using pen As New Pen(Color.Gray)
                e.Graphics.DrawLine(pen, e.Bounds.Left, y, e.Bounds.Right, y)
            End Using
        End If

        e.DrawFocusRectangle()
        ' <<< end changed
    End Sub

    Private Sub cmbProductType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbProductType.SelectedIndexChanged
        ' >>> changed
        ' Reload product names for the selected brand and type
        If cmbBrand.SelectedValue IsNot Nothing AndAlso cmbProductType.Text <> "" Then
            Dim brandId As Integer
            If TypeOf cmbBrand.SelectedValue Is DataRowView Then
                brandId = CInt(DirectCast(cmbBrand.SelectedValue, DataRowView)("BrandID"))
            Else
                brandId = CInt(cmbBrand.SelectedValue)
            End If
            LoadProductNames(brandId, cmbProductType.Text)
        End If
        ' Also reload the product selection ComboBox to filter by product type
        If cmbProductType.Text <> "" Then
            LoadAllProductsToSelection(cmbProductType.Text)
        End If
        ' <<< end changed
    End Sub
    '------------------------------------------------------------------------------
    ' Purpose: Handles adding/updating a supplier price, then refreshes the DataGridView to show the new/updated row.
    ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager, WinForms controls
    ' Current date: 2025-10-03
    '------------------------------------------------------------------------------
    '------------------------------------------------------------------------------
    ' Purpose: Add/update a supplier price then refresh the grid.
    ' Notes:
    '   - SupplierID comes from SupplierInformation (cmbSupplier.SelectedValue)
    '   - MERGE still writes IsPreferred as "Yes"/"No" to match your current table
    '------------------------------------------------------------------------------
    Private Sub btnAddUpdatePrice_Click(sender As Object, e As EventArgs) Handles btnAddUpdatePrice.Click
        ' Validate input
        If cmbBrand.SelectedItem Is Nothing _
        OrElse String.IsNullOrWhiteSpace(cmbProductName.Text) _
        OrElse cmbSupplier.SelectedItem Is Nothing _
        OrElse String.IsNullOrWhiteSpace(txtPricePerUnit.Text) _
        OrElse cmbUnitType.SelectedItem Is Nothing Then

            MessageBox.Show("Please fill in all required fields (Brand, Product Name, Supplier, Price, Unit Type).")
            Exit Sub
        End If

        Dim brandName As String = cmbBrand.Text.Trim()
        Dim productName As String = cmbProductName.Text.Trim()
        Dim productType As String = cmbProductType.Text.Trim()
        Dim fabricWidthInches As Decimal = If(numFabricWidth.Visible, numFabricWidth.Value, 0D)
        Dim supplierName As String = cmbSupplier.Text.Trim()

        Dim pricePerUnit As Decimal
        If Not Decimal.TryParse(txtPricePerUnit.Text, pricePerUnit) OrElse pricePerUnit <= 0D Then
            MessageBox.Show("Please enter a valid price per unit.")
            Exit Sub
        End If

        Dim minQty As Integer = CInt(numQuantity.Value)
        Dim shippingCost As Decimal
        If Not Decimal.TryParse(txtShippingCost.Text, shippingCost) Then shippingCost = 0D

        Dim unitType As String = cmbUnitType.Text.Trim()

        ' Color must be selected from list (your rule)
        Dim colorId As Integer?
        If cmbColor.SelectedItem IsNot Nothing AndAlso cmbColor.SelectedValue IsNot Nothing AndAlso IsNumeric(cmbColor.SelectedValue) Then
            colorId = CInt(cmbColor.SelectedValue)
        Else
            MessageBox.Show("Please select a color from the list. Do not type a new color here.", "Color Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Weight per square inch (only for fabric-like types)
        Dim weightPerLinearYard As Decimal = If(numWeightPerLinearYard.Visible, numWeightPerLinearYard.Value, 0D)
        Dim weightPerSquareInch As Decimal? = Nothing
        If (productType = "Choice Fabric" OrElse productType = "Synthetic Leather" OrElse productType = "Headliner-Padding") _
       AndAlso fabricWidthInches > 0D AndAlso weightPerLinearYard > 0D Then
            weightPerSquareInch = Math.Round(weightPerLinearYard / (36D * fabricWidthInches), 5)
        End If

        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                ' --- Brand ---
                Dim brandId As Integer = GetOrInsertId(conn, "ProductBrands", "BrandName", brandName)

                ' --- Product (unique per brand) ---
                Dim productId As Integer
                Using cmd As New SqlCommand("SELECT ProductID FROM Products WHERE BrandID=@BrandID AND ProductName=@ProductName", conn)
                    cmd.Parameters.AddWithValue("@BrandID", brandId)
                    cmd.Parameters.AddWithValue("@ProductName", productName)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing Then
                        productId = CInt(result)
                    Else
                        Using insertCmd As New SqlCommand("
                        INSERT INTO Products (BrandID, ProductName, ProductType, FabricWidthInches)
                        OUTPUT INSERTED.ProductID
                        VALUES (@BrandID, @ProductName, @ProductType, @FabricWidthInches);", conn)
                            insertCmd.Parameters.AddWithValue("@BrandID", brandId)
                            insertCmd.Parameters.AddWithValue("@ProductName", productName)
                            insertCmd.Parameters.AddWithValue("@ProductType", productType)
                            insertCmd.Parameters.AddWithValue("@FabricWidthInches", fabricWidthInches)
                            productId = CInt(insertCmd.ExecuteScalar())
                        End Using
                    End If
                End Using

                ' --- Supplier (from SupplierInformation) ---
                Dim supplierId As Integer = 0
                If cmbSupplier.SelectedValue IsNot Nothing AndAlso IsNumeric(cmbSupplier.SelectedValue) Then
                    supplierId = CInt(cmbSupplier.SelectedValue) ' <- PK_SupplierNameId
                Else
                    ' Fallback lookup if somehow not bound
                    Using sCmd As New SqlCommand("SELECT PK_SupplierNameId FROM SupplierInformation WHERE CompanyName = @CompanyName", conn)
                        sCmd.Parameters.AddWithValue("@CompanyName", supplierName)
                        Dim sid = sCmd.ExecuteScalar()
                        If sid Is Nothing Then
                            MessageBox.Show("Please pick a supplier from the list.")
                            Exit Sub
                        End If
                        supplierId = CInt(sid)
                    End Using
                End If

                ' --- Upsert price into SupplierProductPrices (your schema) ---
                Using cmd As New SqlCommand("
                MERGE SupplierProductPrices AS target
                USING (SELECT @SupplierID AS SupplierID, @ProductID AS ProductID, @ColorID AS ColorID, @MinQty AS MinQty) AS source
                ON (target.SupplierID = source.SupplierID
                    AND target.ProductID = source.ProductID
                    AND ((target.ColorID IS NULL AND source.ColorID IS NULL) OR (target.ColorID = source.ColorID))
                    AND target.MinQty = source.MinQty)
                WHEN MATCHED THEN
                    UPDATE SET PricePerUnit=@PricePerUnit,
                               UnitType=@UnitType,
                               ShippingCost=@ShippingCost,
                               LastUpdated=GETDATE(),
                               WeightPerSquareInch=@WeightPerSquareInch,
                               IsPreferred=@IsPreferred
                WHEN NOT MATCHED THEN
                    INSERT (SupplierID, ProductID, ColorID, PricePerUnit, UnitType, MinQty, ShippingCost, LastUpdated, WeightPerSquareInch, IsPreferred)
                    VALUES (@SupplierID, @ProductID, @ColorID, @PricePerUnit, @UnitType, @MinQty, @ShippingCost, GETDATE(), @WeightPerSquareInch, @IsPreferred);", conn)

                    cmd.Parameters.AddWithValue("@SupplierID", supplierId) ' <- from SupplierInformation
                    cmd.Parameters.AddWithValue("@ProductID", productId)
                    cmd.Parameters.AddWithValue("@ColorID", If(colorId.HasValue, CType(colorId, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@PricePerUnit", pricePerUnit)
                    cmd.Parameters.AddWithValue("@UnitType", unitType)
                    cmd.Parameters.AddWithValue("@MinQty", minQty)
                    cmd.Parameters.AddWithValue("@ShippingCost", shippingCost)
                    cmd.Parameters.AddWithValue("@WeightPerSquareInch", If(weightPerSquareInch.HasValue, CType(weightPerSquareInch, Object), DBNull.Value))
                    ' You currently store "Yes"/"No" in IsPreferred — keep consistent
                    cmd.Parameters.AddWithValue("@IsPreferred", If(chkPreferredOption.Checked, "Yes", "No"))

                    cmd.ExecuteNonQuery()
                End Using

                MessageBox.Show("Product and price saved successfully.")

                ' Refresh grid for the currently selected supplier
                LoadFilteredSupplierProducts()
            End Using
        Catch ex As Exception
            MessageBox.Show("Error saving product and price: " & ex.Message)
        End Try
    End Sub



    '------------------------------------------------------------------------------
    Private Sub LoadProductPricesToGrid(productId As Integer)
        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()

                'Debug.WriteLine("SQL: " & sql)
                'For Each p As SqlParameter In cmd.Parameters
                '    Debug.WriteLine($"{p.ParameterName} = {p.Value}")
                'Next



                Dim rows As New List(Of SupplierProductPriceRow)()

                Dim sql As String =
"SELECT p.PriceID, p.ProductID, p.SupplierID, s.CompanyName AS Supplier, " & ' >>> changed
"ISNULL(c.ColorName, 'N/A') AS ColorName, " &
"ISNULL(c.ColorNameFriendly, 'N/A') AS FriendlyName, " &
"p.PricePerUnit, p.UnitType, p.MinQty, p.ShippingCost, p.LastUpdated, " &
"pr.WeightPerLinearYard, " &
"CAST(pr.FabricWidthInches AS decimal(18,5)) AS FabricWidthInches, " &
"p.IsPreferred, " &
"(CAST((p.PricePerUnit * p.MinQty + ISNULL(p.ShippingCost,0)) AS decimal(18,5)) / NULLIF(p.MinQty,0)) AS FinalPricePerUnit, " &
"CASE WHEN pr.FabricWidthInches > 0 THEN " &
"    ROUND((CAST((p.PricePerUnit * p.MinQty + ISNULL(p.ShippingCost,0)) AS decimal(18,10)) / (36 * pr.FabricWidthInches * p.MinQty)), 5) " &
"ELSE NULL END AS PricePerSquareInch, " &
"CASE WHEN pr.FabricWidthInches > 0 AND pr.WeightPerLinearYard > 0 THEN " &
"    ROUND((pr.WeightPerLinearYard / (36 * pr.FabricWidthInches)), 5) " &
"ELSE NULL END AS WeightPerSquareInch " &
"FROM SupplierProductPrices p " &
"INNER JOIN SupplierInformation s ON p.SupplierID = s.PK_SupplierNameId " &
"LEFT JOIN ProductColor c ON p.ColorID = c.PK_ColorNameID " &
"INNER JOIN Products pr ON p.ProductID = pr.ProductID " &
"WHERE p.ProductID = @ProductID "

                If selectedSupplierId.HasValue Then
                    sql &= " AND p.SupplierID = @SupplierID"
                End If
                If selectedBrandId.HasValue Then
                    sql &= " AND pr.BrandID = @BrandID"
                End If
                If Not String.IsNullOrWhiteSpace(selectedProductType) Then
                    sql &= " AND pr.ProductType = @ProductType"
                End If
                sql &= " ORDER BY s.CompanyName, c.ColorName, p.MinQty"

                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ProductID", productId)
                    If selectedSupplierId.HasValue Then
                        cmd.Parameters.AddWithValue("@SupplierID", selectedSupplierId.Value)
                    End If
                    If selectedBrandId.HasValue Then
                        cmd.Parameters.AddWithValue("@BrandID", selectedBrandId.Value)
                    End If
                    If Not String.IsNullOrWhiteSpace(selectedProductType) Then
                        cmd.Parameters.AddWithValue("@ProductType", selectedProductType)
                    End If
                    Using rdr As SqlDataReader = cmd.ExecuteReader()
                        While rdr.Read()
                            Dim row As New SupplierProductPriceRow With {
                            .PriceID = If(IsDBNull(rdr("PriceID")), 0, CInt(rdr("PriceID"))),
                            .ProductID = If(IsDBNull(rdr("ProductID")), 0, CInt(rdr("ProductID"))),
                            .SupplierID = If(IsDBNull(rdr("SupplierID")), 0, CInt(rdr("SupplierID"))), ' >>> changed
                            .Supplier = If(IsDBNull(rdr("Supplier")), "", rdr("Supplier").ToString()),
                            .ColorName = If(IsDBNull(rdr("ColorName")), "", rdr("ColorName").ToString()),
                            .FriendlyName = If(IsDBNull(rdr("FriendlyName")), "", rdr("FriendlyName").ToString()),
                            .PricePerUnit = If(IsDBNull(rdr("PricePerUnit")), 0D, Convert.ToDecimal(rdr("PricePerUnit"))),
                            .UnitType = If(IsDBNull(rdr("UnitType")), "", rdr("UnitType").ToString()),
                            .MinQty = If(IsDBNull(rdr("MinQty")), 0, CInt(rdr("MinQty"))),
                            .ShippingCost = If(IsDBNull(rdr("ShippingCost")), 0D, Convert.ToDecimal(rdr("ShippingCost"))),
                            .LastUpdated = If(IsDBNull(rdr("LastUpdated")), Date.MinValue, Convert.ToDateTime(rdr("LastUpdated"))),
                            .WeightPerLinearYard = If(IsDBNull(rdr("WeightPerLinearYard")), 0D, Convert.ToDecimal(rdr("WeightPerLinearYard"))),
                            .FabricWidthInches = If(IsDBNull(rdr("FabricWidthInches")), 0D, Convert.ToDecimal(rdr("FabricWidthInches"))),
                            .IsPreferred = Not IsDBNull(rdr("IsPreferred")) AndAlso rdr("IsPreferred").ToString().Trim().Equals("Yes", StringComparison.OrdinalIgnoreCase),
                            .FinalPricePerUnit = If(IsDBNull(rdr("FinalPricePerUnit")), 0D, Convert.ToDecimal(rdr("FinalPricePerUnit"))),
                            .PricePerSquareInch = If(IsDBNull(rdr("PricePerSquareInch")), 0D, Convert.ToDecimal(rdr("PricePerSquareInch"))),
                            .WeightPerSquareInch = If(IsDBNull(rdr("WeightPerSquareInch")), 0D, Convert.ToDecimal(rdr("WeightPerSquareInch")))
                        }
                            rows.Add(row)
                        End While
                    End Using
                End Using
                Debug.WriteLine($"[DEBUG] Rows loaded in LoadProductPricesToGrid: {rows.Count}")
                dgvSelectedProducts.VirtualMode = False
                dgvSelectedProducts.DataSource = Nothing
                dgvSelectedProducts.AutoGenerateColumns = False
                dgvSelectedProducts.DataSource = New BindingList(Of SupplierProductPriceRow)(rows)
                dgvSelectedProducts.Refresh()
            End Using
            dgvSelectedProducts.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            dgvSelectedProducts.RowTemplate.Height = 75
            For Each row As DataGridViewRow In dgvSelectedProducts.Rows
                row.Height = 75
            Next
            dgvSelectedProducts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        Catch ex As Exception
            MessageBox.Show("Error loading product prices: " & ex.Message)
        End Try

        FormatDgvSelectedProducts()
    End Sub





    '------------------------------------------------------------------------------
    ' Purpose: Build columns AFTER DataSource is set.
    '          - ProductName is a ComboBox using Value/Display mapping (includes "" -> "(none)")
    '          - IsPreferred is a Yes/No ComboBox bound to the Boolean property
    '------------------------------------------------------------------------------
    Private Sub InitializeDgvSelectedProducts()
        dgvSelectedProducts.Columns.Clear()
        dgvSelectedProducts.AutoGenerateColumns = False

        ' ---------- Plain text columns ----------
        AddTextCol(dgvSelectedProducts, "Supplier", "Supplier", 120, isReadOnly:=True)
        AddTextCol(dgvSelectedProducts, "BrandName", "BrandName", 100, isReadOnly:=True)
        AddTextCol(dgvSelectedProducts, "ProductType", "ProductType", 100, isReadOnly:=True)

        ' ---------- ProductName ComboBox with Value/Display mapping ----------
        Dim pnItems As New List(Of KeyValuePair(Of String, String))()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        ' Map empty value to "(none)"
        pnItems.Add(New KeyValuePair(Of String, String)("(none)", ""))
        seen.Add("")

        ' Current grid values
        Dim bl = TryCast(dgvSelectedProducts.DataSource, BindingList(Of SupplierProductPriceRow))
        If bl IsNot Nothing Then
            For Each r In bl
                Dim v As String = If(r.ProductName, "").Trim()
                If v.Length > 0 AndAlso seen.Add(v) Then
                    pnItems.Add(New KeyValuePair(Of String, String)(v, v))
                End If
            Next
        End If

        ' DB values (filtered by brand/type if present)
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                Dim sql As String = "SELECT ProductName FROM Products"
                Dim whereAdded As Boolean = False
                If cmbBrand.SelectedValue IsNot Nothing AndAlso IsNumeric(cmbBrand.SelectedValue) Then
                    sql &= " WHERE BrandID=@BrandID" : whereAdded = True
                End If
                If Not String.IsNullOrWhiteSpace(cmbProductType.Text) Then
                    sql &= If(whereAdded, " AND", " WHERE") & " ProductType=@ProductType"
                End If
                sql &= " ORDER BY ProductName"

                Using cmd As New SqlCommand(sql, conn)
                    If cmbBrand.SelectedValue IsNot Nothing AndAlso IsNumeric(cmbBrand.SelectedValue) Then
                        cmd.Parameters.AddWithValue("@BrandID", CInt(cmbBrand.SelectedValue))
                    End If
                    If Not String.IsNullOrWhiteSpace(cmbProductType.Text) Then
                        cmd.Parameters.AddWithValue("@ProductType", cmbProductType.Text)
                    End If
                    Using rd = cmd.ExecuteReader()
                        While rd.Read()
                            Dim pn As String = rd("ProductName").ToString().Trim()
                            If pn <> "" AndAlso seen.Add(pn) Then
                                pnItems.Add(New KeyValuePair(Of String, String)(pn, pn))
                            End If
                        End While
                    End Using
                End Using
            End Using
        Catch
            ' best effort; list still valid
        End Try

        Dim pnBinding As New BindingSource(pnItems, Nothing)
        Dim colProductName As New DataGridViewComboBoxColumn() With {
        .Name = "ProductName",
        .HeaderText = "Product Name",
        .DataPropertyName = "ProductName",        ' value in the row
        .DataSource = pnBinding,
        .DisplayMember = "Key",                   ' shown text
        .ValueMember = "Value",                   ' stored in cell
        .ValueType = GetType(String),
        .DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton,
        .DisplayStyleForCurrentCellOnly = True,
        .FlatStyle = FlatStyle.Flat,
        .Width = 150
    }
        colProductName.DefaultCellStyle.NullValue = ""  ' show blank for Nothing
        dgvSelectedProducts.Columns.Add(colProductName)

        ' ---------- Remaining text/number columns ----------
        AddTextCol(dgvSelectedProducts, "ColorName", "ColorName", 80, isReadOnly:=True)
        AddTextCol(dgvSelectedProducts, "FriendlyName", "FriendlyName", 90, isReadOnly:=True)
        AddTextCol(dgvSelectedProducts, "PricePerUnit", "PricePerUnit", 70)
        AddTextCol(dgvSelectedProducts, "UnitType", "UnitType", 70)
        AddTextCol(dgvSelectedProducts, "MinQty", "MinQty", 50)
        AddTextCol(dgvSelectedProducts, "ShippingCost", "ShippingCost", 70)
        AddTextCol(dgvSelectedProducts, "LastUpdated", "LastUpdated", 95, isReadOnly:=True, fmt:="yyyy-MM-dd")
        AddTextCol(dgvSelectedProducts, "WeightPerLinearYard", "WeightPerLinearYard", 90)
        AddTextCol(dgvSelectedProducts, "FabricWidthInches", "FabricWidthInches", 90)
        AddTextCol(dgvSelectedProducts, "FinalPricePerUnit", "FinalPricePerUnit", 90, isReadOnly:=True)
        AddTextCol(dgvSelectedProducts, "PricePerSquareInch", "PricePerSquareInch", 90, isReadOnly:=True)
        AddTextCol(dgvSelectedProducts, "WeightPerSquareInch", "WeightPerSquareInch", 90, isReadOnly:=True)

        ' ---------- IsPreferred as a Yes/No ComboBox bound to Boolean ----------
        Dim preferredItems As New List(Of KeyValuePair(Of String, Boolean)) From {
        New KeyValuePair(Of String, Boolean)("Yes", True),
        New KeyValuePair(Of String, Boolean)("No", False)
    }

        Dim colPreferred As New DataGridViewComboBoxColumn() With {
        .Name = "IsPreferred",
        .HeaderText = "Preferred",
        .DataPropertyName = "IsPreferred",                 ' Boolean on your row
        .DataSource = New BindingSource(preferredItems, Nothing),
        .DisplayMember = "Key",
        .ValueMember = "Value",
        .ValueType = GetType(Boolean),
        .DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton,
        .DisplayStyleForCurrentCellOnly = True,
        .FlatStyle = FlatStyle.Flat,
        .Width = 70
    }
        colPreferred.DefaultCellStyle.NullValue = False        ' default to "No" visually
        dgvSelectedProducts.Columns.Add(colPreferred)

        ' ---------- Optional Save button ----------
        If dgvSelectedProducts.Columns("SaveButton") Is Nothing Then
            Dim btnCol As New DataGridViewButtonColumn() With {
            .Name = "SaveButton",
            .HeaderText = "",
            .Text = "Save",
            .UseColumnTextForButtonValue = True,
            .Width = 60
        }
            dgvSelectedProducts.Columns.Add(btnCol)
        End If

        ' ---------- Handlers ----------
        RemoveHandler dgvSelectedProducts.DataError, AddressOf dgvSelectedProducts_DataError
        AddHandler dgvSelectedProducts.DataError, AddressOf dgvSelectedProducts_DataError

        RemoveHandler dgvSelectedProducts.EditingControlShowing, AddressOf dgvSelectedProducts_EditingControlShowing
        AddHandler dgvSelectedProducts.EditingControlShowing, AddressOf dgvSelectedProducts_EditingControlShowing
    End Sub



    '------------------------------------------------------------------------------
    ' Purpose: Stop "value not in list" popups. If ProductName value isn't in the
    '          ComboBox list, add it dynamically so the cell can display.
    '------------------------------------------------------------------------------
    Private Sub dgvSelectedProducts_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) _
    Handles dgvSelectedProducts.DataError

        e.ThrowException = False : e.Cancel = True
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Exit Sub

        Dim grid = DirectCast(sender, DataGridView)
        Dim col = grid.Columns(e.ColumnIndex)

        If String.Equals(col.Name, "IsPreferred", StringComparison.OrdinalIgnoreCase) Then
            ' Default any odd/NULL values to False so the combo can render "No"
            Dim cell = grid.Rows(e.RowIndex).Cells(e.ColumnIndex)
            If cell.Value Is Nothing OrElse cell.Value Is DBNull.Value Then
                cell.Value = False
            End If
        End If
    End Sub


    '------------------------------------------------------------------------------
    ' Purpose: Turn on AutoComplete for the ProductName ComboBox when editing.
    '------------------------------------------------------------------------------
    Private Sub dgvSelectedProducts_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs)
        Dim grid = DirectCast(sender, DataGridView)
        If grid.CurrentCell Is Nothing Then Exit Sub
        If grid.Columns(grid.CurrentCell.ColumnIndex).Name <> "ProductName" Then Exit Sub

        Dim cb = TryCast(e.Control, ComboBox)
        If cb IsNot Nothing Then
            cb.DropDownStyle = ComboBoxStyle.DropDown
            cb.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            cb.AutoCompleteSource = AutoCompleteSource.ListItems
        End If
    End Sub

    '------------------------------------------------------------------------------
    ' Purpose: Populates the ProductName ComboBox with only product names for the selected Brand and ProductType in the current row.
    ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager, System.Windows.Forms
    ' Current date: 2025-10-04
    '------------------------------------------------------------------------------
    Private Sub ProductNameComboBox_DropDown(sender As Object, e As EventArgs)
        ' >>> changed
        Dim combo As ComboBox = TryCast(sender, ComboBox)
        If combo Is Nothing Then Exit Sub

        Dim rowIndex As Integer = dgvSelectedProducts.CurrentCell.RowIndex
        Dim productType As String = ""
        Dim brandName As String = ""
        If dgvSelectedProducts.Rows(rowIndex).Cells("ProductType").Value IsNot Nothing Then
            productType = dgvSelectedProducts.Rows(rowIndex).Cells("ProductType").Value.ToString()
        End If
        If dgvSelectedProducts.Rows(rowIndex).Cells("BrandName").Value IsNot Nothing Then
            brandName = dgvSelectedProducts.Rows(rowIndex).Cells("BrandName").Value.ToString()
        End If

        Dim dtProductNames As New DataTable()
        Dim brandId As Integer = 0
        Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
            Using cmdBrand As New SqlCommand("SELECT BrandID FROM ProductBrands WHERE BrandName = @BrandName", conn)
                cmdBrand.Parameters.AddWithValue("@BrandName", brandName)
                Dim result = cmdBrand.ExecuteScalar()
                If result IsNot Nothing Then brandId = CInt(result)
            End Using
            Using cmd As New SqlCommand("SELECT ProductName FROM Products WHERE BrandID = @BrandID AND ProductType = @ProductType ORDER BY ProductName", conn)
                cmd.Parameters.AddWithValue("@BrandID", brandId)
                cmd.Parameters.AddWithValue("@ProductType", productType)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dtProductNames.Load(reader)
                End Using
            End Using
        End Using

        combo.DataSource = dtProductNames
        combo.DisplayMember = "ProductName"
        combo.ValueMember = "ProductName"
        ' <<< end changed
    End Sub
    '------------------------------------------------------------------------------
    ' Purpose: Handles product selection and loads related prices into dgvSelectedProducts.
    '          FIX: Ensures ComboBox selection is valid and ProductID is not zero.
    ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager, WinForms controls
    ' Current date: 2025-10-03
    '------------------------------------------------------------------------------
    Private Sub cmbProductSelection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbProductSelection.SelectedIndexChanged
        ' >>> changed
        If isLoadingProducts Then Exit Sub

        ' Defensive: ensure ComboBox is actually selected and not empty
        If cmbProductSelection.SelectedIndex = -1 OrElse cmbProductSelection.SelectedValue Is Nothing OrElse cmbProductSelection.Text.Trim() = "" Then
            Exit Sub
        End If

        Dim productId As Integer
        If Integer.TryParse(cmbProductSelection.SelectedValue.ToString(), productId) AndAlso productId > 0 Then
            ' OK
        Else
            Try
                Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                    Using cmd As New SqlCommand("SELECT ProductID FROM Products WHERE ProductName = @ProductName", conn)
                        cmd.Parameters.AddWithValue("@ProductName", cmbProductSelection.Text.Trim())
                        Dim result = cmd.ExecuteScalar()
                        If result IsNot Nothing Then
                            productId = CInt(result)
                        Else
                            MessageBox.Show("Product not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Error looking up product: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try
        End If

        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                Using cmd As New SqlCommand("SELECT BrandID, ProductType, ProductName, FabricWidthInches FROM Products WHERE ProductID = @ProductID", conn)
                    cmd.Parameters.AddWithValue("@ProductID", productId)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            cmbBrand.SelectedValue = reader("BrandID")
                            cmbProductType.Text = reader("ProductType").ToString()
                            cmbProductName.Text = reader("ProductName").ToString()
                            If numFabricWidth IsNot Nothing Then
                                numFabricWidth.Value = If(IsDBNull(reader("FabricWidthInches")), 0D, Convert.ToDecimal(reader("FabricWidthInches")))
                            End If
                        End If
                    End Using
                End Using
            End Using

            LoadColors()
            selectedSupplierId = Nothing
            selectedBrandId = Nothing
            selectedProductType = Nothing

            Debug.WriteLine($"[DEBUG] Loading prices for ProductID={productId}, selectedSupplierId={selectedSupplierId}, selectedBrandId={selectedBrandId}, selectedProductType={selectedProductType}")

            LoadProductPricesToGrid(productId)
        Catch ex As Exception
            MessageBox.Show("Error loading product details: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        ' <<< end changed
    End Sub
    '------------------------------------------------------------------------------
    ' Purpose: Formats DataGridView columns, headers, and ensures only calculated fields are read-only.
    ' Dependencies: System.Windows.Forms, System.Drawing
    ' Current date: 2025-10-03
    '------------------------------------------------------------------------------
    Private Sub FormatDgvSelectedProducts()
        ' >>> changed
        ' Center headers and non-checkbox cells
        For Each col As DataGridViewColumn In dgvSelectedProducts.Columns
            col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            If Not TypeOf col Is DataGridViewCheckBoxColumn Then
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        ' Friendly headers + widths
        Dim headerMap As New Dictionary(Of String, (Header As String, Width As Integer)) From {
        {"Supplier", ("Supplier", 120)},
        {"ColorName", ("Color", 50)},
        {"FriendlyName", ("Friendly Name", 80)},
        {"PricePerUnit", ("Price/Unit", 70)},
        {"UnitType", ("Unit Type", 70)},
        {"MinQty", ("Min Qty", 50)},
        {"ShippingCost", ("Shipping", 60)},
        {"LastUpdated", ("Last Updated", 70)},
        {"WeightPerLinearYard", ("Weight/Linear Yard", 70)},
        {"FabricWidthInches", ("Fabric Width (in)", 70)},
        {"FinalPricePerUnit", ("Final Price/Unit", 70)},
        {"PricePerSquareInch", ("Price / Sq. Inch", 70)},
        {"WeightPerSquareInch", ("Weight / Sq. Inch", 70)},
        {"IsPreferred", ("Preferred", 60)}
    }

        For Each col As DataGridViewColumn In dgvSelectedProducts.Columns
            If headerMap.ContainsKey(col.Name) Then
                col.HeaderText = headerMap(col.Name).Header
                col.Width = headerMap(col.Name).Width
            End If

            ' Only calculated fields are read-only
            If col.Name = "FinalPricePerUnit" OrElse col.Name = "PricePerSquareInch" OrElse col.Name = "WeightPerSquareInch" OrElse col.Name = "Supplier" OrElse col.Name = "ColorName" OrElse col.Name = "FriendlyName" OrElse col.Name = "LastUpdated" OrElse col.Name = "PriceID" OrElse col.Name = "ProductID" Then
                col.ReadOnly = True
            Else
                col.ReadOnly = False
                col.DefaultCellStyle.BackColor = Color.LightYellow
            End If

            ' Formats
            If col.Name = "LastUpdated" Then
                col.DefaultCellStyle.Format = "yyyy-MM-dd"
            ElseIf col.Name = "FabricWidthInches" Then
                col.DefaultCellStyle.Format = "N0"
            ElseIf col.Name = "FinalPricePerUnit" Then
                col.DefaultCellStyle.Format = "N5"
            End If
        Next

        ' >>> changed: Remove old editable highlights loop, now handled above

        ' Row/column sizing
        dgvSelectedProducts.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        dgvSelectedProducts.RowTemplate.Height = 75
        For Each row As DataGridViewRow In dgvSelectedProducts.Rows
            row.Height = 75
        Next
        dgvSelectedProducts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        ' Your existing helper
        EnableDgvHeaderWordWrap()
        ' <<< end changed
    End Sub




    '------------------------------------------------------------------------------
    ' Purpose: Enables word wrap for DataGridView header rows and adjusts header height.
    ' Dependencies: Imports System.Windows.Forms, System.Drawing
    ' Current date: 2025-10-01
    '------------------------------------------------------------------------------
    Private Sub EnableDgvHeaderWordWrap()
        ' >>> changed
        ' Enable word wrap for header cells
        dgvSelectedProducts.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True

        ' Optionally, increase header height for better appearance
        dgvSelectedProducts.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllHeaders
        dgvSelectedProducts.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize

        ' Optionally, force a minimum header height
        dgvSelectedProducts.ColumnHeadersHeight = 40

        ' Refresh to apply changes
        dgvSelectedProducts.Refresh()
        ' <<< end changed
    End Sub

    ' Purpose: Get the preferred supplier price for a given product and color.
    ' Dependencies: Imports System.Data.SqlClient
    ' Current date: 2025-10-02
    Public Function GetPreferredSupplierPrice(productId As Integer, colorId As Integer) As Decimal
        ' >>> changed
        Try
            Using conn As New SqlConnection("your-connection-string")
                conn.Open()
                Using cmd As New SqlCommand("SELECT TOP 1 PricePerUnit FROM SupplierProductPrices WHERE ProductID = @ProductID AND ColorID = @ColorID AND IsPreferred = 1", conn)
                    cmd.Parameters.AddWithValue("@ProductID", productId)
                    cmd.Parameters.AddWithValue("@ColorID", colorId)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return Convert.ToDecimal(result)
                    End If
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception("Error retrieving preferred supplier price: " & ex.Message, ex)
        End Try
        Return 0D
        ' <<< end changed
    End Function
    '------------------------------------------------------------------------------
    ' Purpose: Handles Save button click in dgvSelectedProducts, updates DB, logs price history, and sets preferred supplier price.
    '          Now checks for valid row index to prevent out-of-range errors.
    ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager, System.Windows.Forms
    ' Current date: 2025-10-03
    '------------------------------------------------------------------------------
    Private Sub dgvSelectedProducts_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSelectedProducts.CellClick
        ' >>> changed
        ' Only handle clicks on valid, non-new rows and the SaveButton column
        If e.RowIndex >= 0 AndAlso e.RowIndex < dgvSelectedProducts.Rows.Count - 1 AndAlso dgvSelectedProducts.Columns(e.ColumnIndex).Name = "SaveButton" Then
            Dim row = dgvSelectedProducts.Rows(e.RowIndex)
            Dim priceId = CInt(row.Cells("PriceID").Value)
            Dim newPrice As Decimal = CDec(row.Cells("PricePerUnit").Value)
            Dim newMinQty As Integer = CInt(row.Cells("MinQty").Value)
            Dim newShipping As Decimal = CDec(row.Cells("ShippingCost").Value)
            Dim newWeightPerLinearYard As Decimal = CDec(row.Cells("WeightPerLinearYard").Value)
            Dim newFabricWidthInches As Decimal = CDec(row.Cells("FabricWidthInches").Value)
            Dim productId As Integer = CInt(row.Cells("ProductID").Value)
            Dim colorId As Integer = 0
            If dgvSelectedProducts.Columns.Contains("ColorID") AndAlso row.Cells("ColorID").Value IsNot Nothing AndAlso IsNumeric(row.Cells("ColorID").Value) Then
                colorId = CInt(row.Cells("ColorID").Value)
            End If
            Dim isPreferred As Boolean = False
            If dgvSelectedProducts.Columns.Contains("IsPreferred") AndAlso row.Cells("IsPreferred").Value IsNot Nothing Then
                Dim val = row.Cells("IsPreferred").Value.ToString()
                isPreferred = (val = "Yes")
            End If

            Try
                Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                    ' Get old values for history
                    Dim oldPrice As Decimal, oldMinQty As Integer, oldShipping As Decimal
                    Using cmd As New SqlCommand("SELECT PricePerUnit, MinQty, ShippingCost FROM SupplierProductPrices WHERE PriceID=@PriceID", conn)
                        cmd.Parameters.AddWithValue("@PriceID", priceId)
                        Using rdr = cmd.ExecuteReader()
                            If rdr.Read() Then
                                oldPrice = CDec(rdr("PricePerUnit"))
                                oldMinQty = CInt(rdr("MinQty"))
                                oldShipping = CDec(rdr("ShippingCost"))
                            End If
                        End Using
                    End Using

                    ' Only log if something changed
                    If oldPrice <> newPrice OrElse oldMinQty <> newMinQty OrElse oldShipping <> newShipping Then
                        Using histCmd As New SqlCommand("INSERT INTO SupplierProductPriceHistory (PriceID, OldPricePerUnit, NewPricePerUnit, OldMinQty, NewMinQty, OldShippingCost, NewShippingCost, ChangedBy, ChangedAt) VALUES (@PriceID, @OldPrice, @NewPrice, @OldMinQty, @NewMinQty, @OldShipping, @NewShipping, @ChangedBy, GETDATE())", conn)
                            histCmd.Parameters.AddWithValue("@PriceID", priceId)
                            histCmd.Parameters.AddWithValue("@OldPrice", oldPrice)
                            histCmd.Parameters.AddWithValue("@NewPrice", newPrice)
                            histCmd.Parameters.AddWithValue("@OldMinQty", oldMinQty)
                            histCmd.Parameters.AddWithValue("@NewMinQty", newMinQty)
                            histCmd.Parameters.AddWithValue("@OldShipping", oldShipping)
                            histCmd.Parameters.AddWithValue("@NewShipping", newShipping)
                            histCmd.Parameters.AddWithValue("@ChangedBy", Environment.UserName)
                            histCmd.ExecuteNonQuery()
                        End Using
                    End If

                    ' Update main table
                    Using updCmd As New SqlCommand("UPDATE SupplierProductPrices SET PricePerUnit=@Price, MinQty=@MinQty, ShippingCost=@Shipping, LastUpdated=GETDATE() WHERE PriceID=@PriceID", conn)
                        updCmd.Parameters.AddWithValue("@Price", newPrice)
                        updCmd.Parameters.AddWithValue("@MinQty", newMinQty)
                        updCmd.Parameters.AddWithValue("@Shipping", newShipping)
                        updCmd.Parameters.AddWithValue("@PriceID", priceId)
                        updCmd.ExecuteNonQuery()
                    End Using

                    ' >>> changed
                    ' Update WeightPerLinearYard and FabricWidthInches in the Products table for the associated ProductID
                    Using updProductCmd As New SqlCommand("UPDATE Products SET WeightPerLinearYard=@WeightPerLinearYard, FabricWidthInches=@FabricWidthInches WHERE ProductID=@ProductID", conn)
                        updProductCmd.Parameters.AddWithValue("@WeightPerLinearYard", newWeightPerLinearYard)
                        updProductCmd.Parameters.AddWithValue("@FabricWidthInches", newFabricWidthInches)
                        updProductCmd.Parameters.AddWithValue("@ProductID", productId)
                        updProductCmd.ExecuteNonQuery()
                    End Using
                    ' <<< end changed

                    ' Set preferred supplier price if checked
                    If isPreferred Then
                        ' Unset all others for this product/color
                        Using unsetCmd As New SqlCommand("UPDATE SupplierProductPrices SET IsPreferred = 0 WHERE ProductID = @ProductID AND (ColorID = @ColorID OR (@ColorID = 0 AND ColorID IS NULL))", conn)
                            unsetCmd.Parameters.AddWithValue("@ProductID", productId)
                            unsetCmd.Parameters.AddWithValue("@ColorID", colorId)
                            unsetCmd.ExecuteNonQuery()
                        End Using
                    End If
                    ' Set this row's IsPreferred
                    Using setCmd As New SqlCommand("UPDATE SupplierProductPrices SET IsPreferred = @IsPreferred WHERE PriceID = @PriceID", conn)
                        setCmd.Parameters.AddWithValue("@IsPreferred", If(isPreferred, 1, 0))
                        setCmd.Parameters.AddWithValue("@PriceID", priceId)
                        setCmd.ExecuteNonQuery()
                    End Using
                End Using
                MessageBox.Show("Row updated, history logged, and preferred supplier set.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Error saving row: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            ' Refresh grid
            If row.Cells("ProductID").Value IsNot Nothing Then
                LoadProductPricesToGrid(CInt(row.Cells("ProductID").Value))
            End If
        End If
        ' <<< end changed
    End Sub
    ' <<< end changed
    Private Sub dgvSelectedProducts_MouseDown(sender As Object, e As MouseEventArgs) Handles dgvSelectedProducts.MouseDown
        If e.Button = MouseButtons.Right Then
            Dim hit As DataGridView.HitTestInfo = dgvSelectedProducts.HitTest(e.X, e.Y)
            If hit.Type = DataGridViewHitTestType.Cell Then
                dgvSelectedProducts.ClearSelection()
                dgvSelectedProducts.Rows(hit.RowIndex).Selected = True
                dgvSelectedProducts.CurrentCell = dgvSelectedProducts.Rows(hit.RowIndex).Cells(0)
            End If
        End If
    End Sub

    ' Handle the context menu "Delete Row" click:
    Private Sub mnuDeleteRow_Click(sender As Object, e As EventArgs) Handles mnuDeleteRow.Click
        If dgvSelectedProducts.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a row to delete.", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim row As DataGridViewRow = dgvSelectedProducts.SelectedRows(0)
        ' Assumes your DataGridView has a PriceID column (hidden or visible)
        Dim priceIdObj As Object = row.Cells("PriceID").Value
        If priceIdObj Is Nothing OrElse Not IsNumeric(priceIdObj) Then
            MessageBox.Show("Could not determine the PriceID for the selected row.", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim priceId As Integer = CInt(priceIdObj)

        If MessageBox.Show("Are you sure you want to delete this price entry?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
            Try
                Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                    Using cmd As New SqlCommand("DELETE FROM SupplierProductPrices WHERE PriceID = @PriceID", conn)
                        cmd.Parameters.AddWithValue("@PriceID", priceId)
                        Dim rowsAffected = cmd.ExecuteNonQuery()
                        If rowsAffected > 0 Then
                            MessageBox.Show("Row deleted.", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Refresh grid (get ProductID from row or current context)
                            Dim productId As Integer
                            If row.Cells("ProductID").Value IsNot Nothing AndAlso IsNumeric(row.Cells("ProductID").Value) Then
                                productId = CInt(row.Cells("ProductID").Value)
                            ElseIf cmbProductSelection.SelectedValue IsNot Nothing AndAlso IsNumeric(cmbProductSelection.SelectedValue) Then
                                productId = CInt(cmbProductSelection.SelectedValue)
                            Else
                                productId = 0
                            End If
                            If productId > 0 Then
                                LoadProductPricesToGrid(productId)
                            End If
                        Else
                            MessageBox.Show("Delete failed. Row not found.", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Error deleting row: " & ex.Message, "Delete", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub btnCloseForm_Click(sender As Object, e As EventArgs) Handles btnCloseForm.Click
        Dim dashboard As New formDashboard()
        formDashboard.Show()
        Me.Close()
    End Sub


    Private Sub LoadProductNames(brandId As Integer, productType As String)
        '------------------------------------------------------------------------------
        ' Purpose: Loads product names for the selected brand and type into cmbProductName.
        '          If the product type is "Headliner Padding" and no product names exist,
        '          adds a "Not set" entry (ProductTypeID 18) to the ComboBox.
        ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager
        ' Current date: 2025-10-03
        '------------------------------------------------------------------------------
        ' >>> changed
        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                Dim dt As New DataTable()
                ' Get all product names for this brand and type
                Using cmd As New SqlCommand("SELECT ProductName FROM Products WHERE BrandID = @BrandID AND ProductType = @ProductType ORDER BY ProductName", conn)
                    cmd.Parameters.AddWithValue("@BrandID", brandId)
                    cmd.Parameters.AddWithValue("@ProductType", productType)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using

                ' If "Headliner Padding" and no product names, add "Not set"
                If productType.Trim().Equals("Headliner Padding", StringComparison.OrdinalIgnoreCase) AndAlso dt.Rows.Count = 0 Then
                    dt.Columns.Add("ProductTypeID", GetType(Integer))
                    Dim row As DataRow = dt.NewRow()
                    row("ProductName") = "Not set"
                    row("ProductTypeID") = 18
                    dt.Rows.Add(row)
                End If

                cmbProductName.DataSource = dt
                cmbProductName.DisplayMember = "ProductName"
                cmbProductName.ValueMember = "ProductName"
                cmbProductName.Text = ""
            End Using

            If cmbProductName.SelectedValue IsNot Nothing Then

            End If
        Catch ex As Exception
            MessageBox.Show("Error loading product names: " & ex.Message)
        End Try
        ' <<< end changed
    End Sub


    Private Sub cmbProductName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbProductName.SelectedIndexChanged
        LoadColors()
    End Sub
    Private Sub LoadProductTypes()
        ' >>> changed
        ' Loads product types, putting those associated with the selected brand at the top (alphabetically), then the rest (alphabetically).
        ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager
        ' Current date: 2025-10-03

        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                Dim dt As New DataTable()

                ' Get the selected brand ID (if any)
                Dim brandId As Integer? = Nothing
                If cmbBrand IsNot Nothing AndAlso cmbBrand.SelectedValue IsNot Nothing AndAlso IsNumeric(cmbBrand.SelectedValue) Then
                    brandId = CInt(cmbBrand.SelectedValue)
                End If

                Dim sql As String
                If brandId.HasValue Then
                    ' Use UNION ALL and explicit column names/types for correct ordering
                    sql =
                "SELECT pt.ProductTypeID, pt.ProductTypeName, 0 AS SortOrder " &
                "FROM ProductTypes pt " &
                "WHERE pt.ProductTypeName IN (SELECT DISTINCT ProductType FROM Products WHERE BrandID = @BrandID) " &
                "UNION ALL " &
                "SELECT pt.ProductTypeID, pt.ProductTypeName, 1 AS SortOrder " &
                "FROM ProductTypes pt " &
                "WHERE pt.ProductTypeName NOT IN (SELECT DISTINCT ProductType FROM Products WHERE BrandID = @BrandID) " &
                "ORDER BY SortOrder, ProductTypeName"
                Else
                    sql = "SELECT ProductTypeID, ProductTypeName, 1 AS SortOrder FROM ProductTypes ORDER BY ProductTypeName"
                End If

                Using cmd As New SqlCommand(sql, conn)
                    If brandId.HasValue Then
                        cmd.Parameters.AddWithValue("@BrandID", brandId.Value)
                    End If
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dt.Load(reader)
                    End Using
                End Using

                cmbProductType.DataSource = dt
                cmbProductType.DisplayMember = "ProductTypeName"
                cmbProductType.ValueMember = "ProductTypeID"
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading product types: " & ex.Message)
        End Try
        ' <<< end changed
    End Sub

    '------------------------------------------------------------------------------
    ' Purpose: Loads brands from ProductBrands table into cmbBrand ComboBox.
    ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager
    ' Current date: 2025-10-01
    '------------------------------------------------------------------------------
    Private Sub LoadBrands()
        ' >>> changed
        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                Using cmd As New SqlCommand("SELECT BrandID, BrandName FROM ProductBrands ORDER BY BrandName", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        Dim dt As New DataTable()
                        dt.Load(reader)
                        cmbBrand.DataSource = dt
                        cmbBrand.DisplayMember = "BrandName"
                        cmbBrand.ValueMember = "BrandID"
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading brands: " & ex.Message)
        End Try
        ' <<< end changed
    End Sub

    '------------------------------------------------------------------------------
    ' Purpose: Loads suppliers from Suppliers table into cmbSupplier ComboBox.
    ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager
    ' Current date: 2025-10-01
    '------------------------------------------------------------------------------
    Private Sub LoadSuppliers()
        ' >>> changed
        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                Using cmd As New SqlCommand("SELECT PK_SupplierNameId, CompanyName FROM SupplierInformation ORDER BY CompanyName", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        Dim dt As New DataTable()
                        dt.Load(reader)
                        cmbSupplier.DataSource = dt
                        cmbSupplier.DisplayMember = "CompanyName"
                        cmbSupplier.ValueMember = "PK_SupplierNameId"
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading suppliers: " & ex.Message)
        End Try
        ' <<< end changed
    End Sub

    '------------------------------------------------------------------------------
    ' Purpose: Loads all available colors from ProductColor into cmbColor.
    '          Uses ColorName as display, PK_ColorNameID as value, and ColorNameFriendly for the friendly name.
    ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager, WinForms controls
    ' Current date: 2025-10-03
    '------------------------------------------------------------------------------
    Private Sub LoadColors()
        ' >>> changed
        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                Dim sql As String = "SELECT PK_ColorNameID, ColorName, ColorNameFriendly FROM ProductColor ORDER BY ColorName"
                Using cmd As New SqlCommand(sql, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        Dim dt As New DataTable()
                        dt.Load(reader)
                        cmbColor.DataSource = dt
                        cmbColor.DisplayMember = "ColorName"
                        cmbColor.ValueMember = "PK_ColorNameID"
                    End Using
                End Using
            End Using
            ' Update txtColorFriendlyName if a color is selected
            If cmbColor.SelectedItem IsNot Nothing AndAlso TypeOf cmbColor.SelectedItem Is DataRowView Then
                Dim drv As DataRowView = DirectCast(cmbColor.SelectedItem, DataRowView)
                txtColorFriendlyName.Text = drv("ColorNameFriendly").ToString()
            ElseIf cmbColor.SelectedIndex >= 0 AndAlso cmbColor.Items.Count > 0 Then
                Dim drv As DataRowView = DirectCast(cmbColor.Items(cmbColor.SelectedIndex), DataRowView)
                txtColorFriendlyName.Text = drv("ColorNameFriendly").ToString()
            Else
                txtColorFriendlyName.Text = ""
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading colors: " & ex.Message)
        End Try
        ' <<< end changed
    End Sub

    '------------------------------------------------------------------------------
    ' Purpose: Updates txtColorFriendlyName when a color is selected in cmbColor.
    ' Dependencies: None
    ' Current date: 2025-10-03
    '------------------------------------------------------------------------------
    Private Sub cmbColor_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbColor.SelectedIndexChanged
        ' >>> changed
        If cmbColor.SelectedItem IsNot Nothing AndAlso TypeOf cmbColor.SelectedItem Is DataRowView Then
            Dim drv As DataRowView = DirectCast(cmbColor.SelectedItem, DataRowView)
            txtColorFriendlyName.Text = drv("ColorNameFriendly").ToString()
        ElseIf cmbColor.SelectedIndex >= 0 AndAlso cmbColor.Items.Count > 0 Then
            Dim drv As DataRowView = DirectCast(cmbColor.Items(cmbColor.SelectedIndex), DataRowView)
            txtColorFriendlyName.Text = drv("ColorNameFriendly").ToString()
        Else
            txtColorFriendlyName.Text = ""
        End If
        ' <<< end changed
    End Sub
    '------------------------------------------------------------------------------
    ' Purpose: Inserts a new fabric product with weight per linear yard and square yard.
    ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager
    ' Current date: 2025-10-01
    '------------------------------------------------------------------------------
    Private Sub AddFabricProduct(brandId As Integer, productName As String, productType As String, fabricWidthInches As Decimal)
        ' >>> changed
        ' Insert a new product with weight columns
        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                Using insertCmd As New SqlCommand("INSERT INTO Products (BrandID, ProductName, ProductType, FabricWidthInches, WeightPerLinearYard, WeightPerSquareYard) OUTPUT INSERTED.ProductID VALUES (@BrandID, @ProductName, @ProductType, @FabricWidthInches, @WeightPerLinearYard, @WeightPerSquareYard)", conn)
                    insertCmd.Parameters.AddWithValue("@BrandID", brandId)
                    insertCmd.Parameters.AddWithValue("@ProductName", productName)
                    insertCmd.Parameters.AddWithValue("@ProductType", productType)
                    insertCmd.Parameters.AddWithValue("@FabricWidthInches", fabricWidthInches)
                    insertCmd.Parameters.AddWithValue("@WeightPerLinearYard", numWeightPerLinearYard.Value)
                    insertCmd.Parameters.AddWithValue("@WeightPerSquareYard", numWeightPerSquareYard.Value)
                    Dim productId As Integer = CInt(insertCmd.ExecuteScalar())
                    MessageBox.Show("Fabric product added with weights.", "Success")
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error adding fabric product: " & ex.Message)
        End Try
        ' <<< end changed
    End Sub

    Private Sub LoadUnitTypes()
        ' >>> changed
        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                Using cmd As New SqlCommand("SELECT UnitTypeID, UnitTypeName FROM ProductUnitTypes ORDER BY UnitTypeName", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        Dim dt As New DataTable()
                        dt.Load(reader)
                        cmbUnitType.DataSource = dt
                        cmbUnitType.DisplayMember = "UnitTypeName"
                        cmbUnitType.ValueMember = "UnitTypeID"
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading unit types: " & ex.Message)
        End Try
        ' <<< end changed
    End Sub




    '------------------------------------------------------------------------------
    ' Purpose: Loads products for the selected brand and product type into cmbProductSelection.
    '          Only products matching both are shown.
    ' Dependencies: Imports System.Data.SqlClient, Data.DbConnectionManager, WinForms controls
    ' Current date: 2025-10-03
    '------------------------------------------------------------------------------
    Private Sub LoadAllProductsToSelection(Optional productType As String = "")
        ' >>> changed
        isLoadingProducts = True ' Suppress event while loading
        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                ' Build SQL to filter by both BrandID and ProductType if available
                Dim sql As String = "SELECT ProductID, ProductName FROM Products"
                Dim hasBrand As Boolean = cmbBrand.SelectedValue IsNot Nothing AndAlso IsNumeric(cmbBrand.SelectedValue)
                Dim hasType As Boolean = Not String.IsNullOrWhiteSpace(productType)
                If hasBrand OrElse hasType Then
                    sql &= " WHERE"
                    If hasBrand Then
                        sql &= " BrandID = @BrandID"
                    End If
                    If hasType Then
                        If hasBrand Then sql &= " AND"
                        sql &= " ProductType = @ProductType"
                    End If
                End If
                sql &= " ORDER BY ProductName"
                Using cmd As New SqlCommand(sql, conn)
                    If hasBrand Then
                        cmd.Parameters.AddWithValue("@BrandID", CInt(cmbBrand.SelectedValue))
                    End If
                    If hasType Then
                        cmd.Parameters.AddWithValue("@ProductType", productType)
                    End If
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        Dim dt As New DataTable()
                        dt.Load(reader)
                        cmbProductSelection.DataSource = dt
                        cmbProductSelection.DisplayMember = "ProductName"
                        cmbProductSelection.ValueMember = "ProductID"
                        cmbProductSelection.SelectedIndex = -1 ' No selection by default
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading products: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            isLoadingProducts = False ' Allow event after loading
        End Try
        ' <<< end changed
    End Sub

    '------------------------------------------------------------------------------
    ' Purpose: Loads supplier product prices into the grid and binds a BindingList.
    '          Trims strings and converts IsPreferred to Boolean.
    '------------------------------------------------------------------------------
    Private Sub LoadFilteredSupplierProducts()
        Try
            Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                Dim sql As String =
"SELECT p.PriceID, p.ProductID, pr.BrandID, s.CompanyName AS Supplier, pr.ProductType, pr.ProductName, " &
"b.BrandName, " &
"ISNULL(c.ColorName, 'N/A') AS ColorName, ISNULL(c.ColorNameFriendly, 'N/A') AS FriendlyName, " &
"p.PricePerUnit, p.UnitType, p.MinQty, p.ShippingCost, p.LastUpdated, " &
"pr.WeightPerLinearYard, CAST(pr.FabricWidthInches AS decimal(18,5)) AS FabricWidthInches, " &
"p.IsPreferred, " &
"(CAST((p.PricePerUnit * p.MinQty + ISNULL(p.ShippingCost,0)) AS decimal(18,5)) / NULLIF(p.MinQty,0)) AS FinalPricePerUnit, " &
"CASE WHEN pr.FabricWidthInches > 0 THEN " &
"  ROUND((CAST((p.PricePerUnit * p.MinQty + ISNULL(p.ShippingCost,0)) AS decimal(18,10)) / (36 * pr.FabricWidthInches * p.MinQty)), 5) " &
"ELSE NULL END AS PricePerSquareInch, " &
"CASE WHEN pr.FabricWidthInches > 0 AND pr.WeightPerLinearYard > 0 THEN " &
"  ROUND((pr.WeightPerLinearYard / (36 * pr.FabricWidthInches)), 5) " &
"ELSE NULL END AS WeightPerSquareInch " &
"FROM SupplierProductPrices p " &
"INNER JOIN SupplierInformation s ON p.SupplierID = s.PK_SupplierNameId " &
"LEFT JOIN ProductColor c ON p.ColorID = c.PK_ColorNameID " &
"INNER JOIN Products pr ON p.ProductID = pr.ProductID " &
"INNER JOIN ProductBrands b ON pr.BrandID = b.BrandID " &
"WHERE 1=1"

                If selectedSupplierId.HasValue Then sql &= " AND p.SupplierID = @SupplierID"
                If selectedBrandId.HasValue Then sql &= " AND pr.BrandID = @BrandID"
                If Not String.IsNullOrWhiteSpace(selectedProductType) Then sql &= " AND pr.ProductType = @ProductType"
                sql &= " ORDER BY s.CompanyName, c.ColorName, p.MinQty"

                Using cmd As New SqlCommand(sql, conn)
                    If selectedSupplierId.HasValue Then cmd.Parameters.AddWithValue("@SupplierID", selectedSupplierId.Value)
                    If selectedBrandId.HasValue Then cmd.Parameters.AddWithValue("@BrandID", selectedBrandId.Value)
                    If Not String.IsNullOrWhiteSpace(selectedProductType) Then cmd.Parameters.AddWithValue("@ProductType", selectedProductType)

                    Using rdr As SqlDataReader = cmd.ExecuteReader()
                        Dim rows As New List(Of SupplierProductPriceRow)
                        While rdr.Read()
                            Dim prefBool As Boolean = False
                            If Not IsDBNull(rdr("IsPreferred")) Then
                                Dim raw = rdr("IsPreferred")
                                If TypeOf raw Is Boolean Then
                                    prefBool = CBool(raw)
                                Else
                                    Dim s = raw.ToString().Trim()
                                    If IsNumeric(raw) Then prefBool = (Convert.ToInt32(raw) = 1) _
                                Else prefBool = (String.Equals(s, "Yes", StringComparison.OrdinalIgnoreCase) OrElse s = "True")
                                End If
                            End If

                            rows.Add(New SupplierProductPriceRow With {
                            .PriceID = If(IsDBNull(rdr("PriceID")), 0, CInt(rdr("PriceID"))),
                            .ProductID = If(IsDBNull(rdr("ProductID")), 0, CInt(rdr("ProductID"))),
                            .Supplier = If(IsDBNull(rdr("Supplier")), "", rdr("Supplier").ToString().Trim()),
                            .BrandName = If(IsDBNull(rdr("BrandName")), "", rdr("BrandName").ToString().Trim()),
                            .ProductType = If(IsDBNull(rdr("ProductType")), "", rdr("ProductType").ToString().Trim()),
                            .ProductName = If(IsDBNull(rdr("ProductName")), "", rdr("ProductName").ToString().Trim()),
                            .ColorName = If(IsDBNull(rdr("ColorName")), "", rdr("ColorName").ToString().Trim()),
                            .FriendlyName = If(IsDBNull(rdr("FriendlyName")), "", rdr("FriendlyName").ToString().Trim()),
                            .PricePerUnit = If(IsDBNull(rdr("PricePerUnit")), 0D, Convert.ToDecimal(rdr("PricePerUnit"))),
                            .UnitType = If(IsDBNull(rdr("UnitType")), "", rdr("UnitType").ToString().Trim()),
                            .MinQty = If(IsDBNull(rdr("MinQty")), 0, CInt(rdr("MinQty"))),
                            .ShippingCost = If(IsDBNull(rdr("ShippingCost")), 0D, Convert.ToDecimal(rdr("ShippingCost"))),
                            .LastUpdated = If(IsDBNull(rdr("LastUpdated")), Date.MinValue, Convert.ToDateTime(rdr("LastUpdated"))),
                            .WeightPerLinearYard = If(IsDBNull(rdr("WeightPerLinearYard")), 0D, Convert.ToDecimal(rdr("WeightPerLinearYard"))),
                            .FabricWidthInches = If(IsDBNull(rdr("FabricWidthInches")), 0D, Convert.ToDecimal(rdr("FabricWidthInches"))),
                            .IsPreferred = prefBool,
                            .FinalPricePerUnit = If(IsDBNull(rdr("FinalPricePerUnit")), 0D, Convert.ToDecimal(rdr("FinalPricePerUnit"))),
                            .PricePerSquareInch = If(IsDBNull(rdr("PricePerSquareInch")), 0D, Convert.ToDecimal(rdr("PricePerSquareInch"))),
                            .WeightPerSquareInch = If(IsDBNull(rdr("WeightPerSquareInch")), 0D, Convert.ToDecimal(rdr("WeightPerSquareInch")))
                        })
                        End While

                        dgvSelectedProducts.VirtualMode = False
                        dgvSelectedProducts.DataSource = Nothing
                        dgvSelectedProducts.AutoGenerateColumns = False

                        ' Bind first
                        Dim bl As New BindingList(Of SupplierProductPriceRow)(rows)
                        dgvSelectedProducts.DataSource = bl

                        ' Then build columns
                        InitializeDgvSelectedProducts()
                    End Using
                End Using
            End Using

            FormatDgvSelectedProducts()
        Catch ex As Exception
            MessageBox.Show("Error loading supplier products: " & ex.Message)
        End Try
    End Sub

    '------------------------------------------------------------------------------
    ' Purpose: Helper to get or insert an ID for a value (optionally with a parent foreign key).
    '          Uses explicit mapping for ID column names to avoid SQL errors.
    ' Dependencies: Imports System.Data.SqlClient
    ' Current date: 2025-10-01
    '------------------------------------------------------------------------------
    Private Function GetOrInsertId(conn As SqlConnection, tableName As String, columnName As String, value As String, Optional parentColumn As String = Nothing, Optional parentId As Integer = 0) As Integer
        ' >>> changed
        ' Map table names to their correct ID column names
        Dim idColumn As String
        Select Case tableName
            Case "ProductBrands"
                idColumn = "BrandID"
            Case "Products"
                idColumn = "ProductID"
            Case "ProductColors"
                idColumn = "ColorID"
            Case "SupplierInformation"
                idColumn = "PK_SupplierNameId"
            Case Else
                idColumn = tableName.Substring(0, tableName.Length - 1) & "ID"
        End Select

        Dim id As Integer = 0
        Dim sql As String = $"SELECT {idColumn} FROM {tableName} WHERE {columnName}=@Value"
        If parentColumn IsNot Nothing Then
            sql &= $" AND {parentColumn}=@ParentId"
        End If
        Using cmd As New SqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@Value", value)
            If parentColumn IsNot Nothing Then
                cmd.Parameters.AddWithValue("@ParentId", parentId)
            End If
            Dim result = cmd.ExecuteScalar()
            If result IsNot Nothing Then
                id = CInt(result)
            Else
                sql = $"INSERT INTO {tableName} ({columnName}"
                If parentColumn IsNot Nothing Then sql &= $", {parentColumn}"
                sql &= $") OUTPUT INSERTED.{idColumn} VALUES (@Value"
                If parentColumn IsNot Nothing Then sql &= ", @ParentId"
                sql &= ")"
                Using insertCmd As New SqlCommand(sql, conn)
                    insertCmd.Parameters.AddWithValue("@Value", value)
                    If parentColumn IsNot Nothing Then
                        insertCmd.Parameters.AddWithValue("@ParentId", parentId)
                    End If
                    id = CInt(insertCmd.ExecuteScalar())
                End Using
            End If
        End Using
        Return id
        ' <<< end changed
    End Function


    '------------------------------------------------------------------------------
    ' Purpose: Adds a checkbox column to dgvTest and populates it with sample data.
    ' Dependencies: Imports System.Windows.Forms, System.Data
    ' Current date: 2025-10-02
    '------------------------------------------------------------------------------

    Private Sub AddCheckboxToDgvTest()
        ' >>> changed
        ' Create a DataTable with a Boolean column
        Dim dt As New DataTable()
        dt.Columns.Add("Name", GetType(String))
        dt.Columns.Add("IsChecked", GetType(Boolean))

        ' Add sample rows
        dt.Rows.Add("Row 1", True)
        dt.Rows.Add("Row 2", False)
        dt.Rows.Add("Row 3", True)

        ' Clear and set up the DataGridView
        dgvTest.Columns.Clear()
        dgvTest.DataSource = Nothing
        dgvTest.AutoGenerateColumns = False

        ' Add a text column
        Dim colText As New DataGridViewTextBoxColumn() With {
        .Name = "Name",
        .HeaderText = "Name",
        .DataPropertyName = "Name",
        .ReadOnly = True
    }
        dgvTest.Columns.Add(colText)

        ' Add a checkbox column bound to the Boolean column
        Dim colCheck As New DataGridViewCheckBoxColumn() With {
        .Name = "IsChecked",
        .HeaderText = "Checked?",
        .DataPropertyName = "IsChecked",
        .ThreeState = False,
        .Width = 60
    }
        dgvTest.Columns.Add(colCheck)

        ' Bind the data
        dgvTest.DataSource = dt
        ' <<< end changed
    End Sub

    '------------------------------------------------------------------------------
    ' AddTextCol
    '------------------------------------------------------------------------------
    Private Sub AddTextCol(grid As DataGridView,
                       colName As String,
                       dataProp As String,
                       width As Integer,
                       Optional isReadOnly As Boolean = False,
                       Optional isVisible As Boolean = True,
                       Optional fmt As String = Nothing)
        Dim col As New DataGridViewTextBoxColumn() With {
        .Name = colName,
        .HeaderText = colName,
        .DataPropertyName = dataProp,
        .Width = width,
        .ReadOnly = isReadOnly,
        .Visible = isVisible,
        .SortMode = DataGridViewColumnSortMode.NotSortable
    }
        If Not String.IsNullOrEmpty(fmt) Then col.DefaultCellStyle.Format = fmt
        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        grid.Columns.Add(col)
    End Sub

    Private Sub cmbSupplier_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles cmbSupplier.SelectedIndexChanged

    End Sub

    ''------------------------------------------------------------------------------
    '' Purpose: Adds a DataGridViewTextBoxColumn to the specified DataGridView with common options.
    '' Dependencies: Imports System.Windows.Forms
    '' Current date: 2025-10-03
    ''------------------------------------------------------------------------------
    'Private Sub AddTextCol(
    '    dgv As DataGridView,
    '    name As String,
    '    dataPropertyName As String,
    '    width As Integer,
    '    Optional isReadOnly As Boolean = False,
    '    Optional isVisible As Boolean = True,
    '    Optional fmt As String = Nothing
    ')
    '    ' >>> changed
    '    Dim col As New DataGridViewTextBoxColumn() With {
    '        .Name = name,
    '        .HeaderText = name,
    '        .DataPropertyName = dataPropertyName,
    '        .Width = width,
    '        .ReadOnly = isReadOnly,
    '        .Visible = isVisible
    '    }
    '    If Not String.IsNullOrEmpty(fmt) Then
    '        col.DefaultCellStyle.Format = fmt
    '    End If
    '    dgv.Columns.Add(col)
    '    ' <<< end changed
    'End Sub
End Class
'------------------------------------------------------------------------------
' Purpose: Represents a row in the supplier product price grid, with IsPreferred as Boolean.
' Dependencies: None (part of FormMaterialsSuppliers)
' Current date: 2025-10-03
'------------------------------------------------------------------------------
Public Class SupplierProductPriceRow
    Public Property PriceID As Integer
    Public Property ProductID As Integer
    Public Property BrandName As String
    Public Property Supplier As String
    Public Property ProductType As String
    Public Property ProductName As String
    Public Property ColorName As String
    Public Property FriendlyName As String
    Public Property PricePerUnit As Decimal
    Public Property UnitType As String
    Public Property MinQty As Integer
    Public Property ShippingCost As Decimal
    Public Property LastUpdated As Date
    Public Property WeightPerLinearYard As Decimal
    Public Property FabricWidthInches As Decimal
    Public Property IsPreferred As Boolean
    Public Property FinalPricePerUnit As Decimal
    Public Property PricePerSquareInch As Decimal
    Public Property WeightPerSquareInch As Decimal
    Public Property SupplierID As Integer

    Public Property IsPreferredDisplay As String
        Get
            Return If(IsPreferred, "Yes", "No")
        End Get
        Set(value As String)
            IsPreferred = (value = "Yes")
        End Set
    End Property
End Class