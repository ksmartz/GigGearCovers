Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Linq
Public Class DbConnectionManager
    ' Adjust your connection string as needed
    Private Shared ReadOnly Property ConnectionString As String
        Get
            Return "Data Source=MYPC\SQLEXPRESS;Initial Catalog=GigGearCoversDb;Integrated Security=True"
        End Get
    End Property

    Public Shared Function GetConnection() As SqlConnection
        Return New SqlConnection(ConnectionString)
    End Function
    ' Add this to DbConnectionManager.vb:
    Public Function GetProductsForSupplier(supplierId As Integer) As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT DISTINCT 
                    p.PK_FabricBrandProductNameId, 
                    p.BrandProductName
                FROM SupplierProductNameData s
                INNER JOIN JoinProductColorFabricType j ON s.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                INNER JOIN FabricBrandProductName p ON j.FK_FabricBrandProductNameId = p.PK_FabricBrandProductNameId
                WHERE s.FK_SupplierNameId = @SupplierId
                ORDER BY p.BrandProductName"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                Using adapter As New SqlDataAdapter(CType(cmd, SqlCommand))
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function
    Public Function GetOrCreateJoinProductColorFabricTypeId(productId As Integer, colorId As Integer, fabricTypeId As Integer) As Integer
        Using conn = GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT PK_JoinProductColorFabricTypeId
                FROM JoinProductColorFabricType
                WHERE FK_FabricBrandProductNameId = @ProductId
                  AND FK_ColorNameID = @ColorId
                  AND FK_FabricTypeNameId = @FabricTypeId"
                cmd.Parameters.AddWithValue("@ProductId", productId)
                cmd.Parameters.AddWithValue("@ColorId", colorId)
                cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing Then
                    Return CInt(result)
                End If

                ' Insert if not found
                cmd.CommandText = "
                INSERT INTO JoinProductColorFabricType (FK_FabricBrandProductNameId, FK_ColorNameID, FK_FabricTypeNameId)
                VALUES (@ProductId, @ColorId, @FabricTypeId);
                SELECT CAST(SCOPE_IDENTITY() AS int);"
                Return CInt(cmd.ExecuteScalar())
            End Using
        End Using
    End Function
    Public Function GetAllMaterialColors() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_ColorNameID, ColorName, ColorNameFriendly FROM FabricColor ORDER BY ColorNameFriendly"
                Dim dt As New DataTable()
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
                Return dt
            End Using
        End Using
    End Function

    Public Sub UpdateSupplierProductTotalYards(supplierProductNameDataId As Integer, totalYards As Decimal)
        Using conn = GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "UPDATE SupplierProductNameData SET TotalYards = @TotalYards WHERE PK_SupplierProductNameDataId = @Id"
                cmd.Parameters.AddWithValue("@TotalYards", totalYards)
                cmd.Parameters.AddWithValue("@Id", supplierProductNameDataId)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Public Sub SetActiveForMarketplaceForCombination(brandId As Integer, productId As Integer, colorId As Integer, fabricTypeId As Integer, activeSupplierId As Integer)
        Using conn = GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using trans = conn.BeginTransaction()
                ' Set all to inactive
                Using cmd = conn.CreateCommand()
                    cmd.Transaction = trans
                    cmd.CommandText = "
                    UPDATE spnd
                    SET spnd.IsActiveForMarketplace = 0
                    FROM SupplierProductNameData spnd
                    INNER JOIN JoinProductColorFabricType j ON spnd.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                    WHERE j.FK_FabricBrandProductNameId = @ProductId
                      AND j.FK_ColorNameID = @ColorId
                      AND j.FK_FabricTypeNameId = @FabricTypeId
                      AND EXISTS (
                          SELECT 1 FROM FabricBrandProductName p
                          WHERE p.PK_FabricBrandProductNameId = j.FK_FabricBrandProductNameId
                            AND p.FK_FabricBrandNameId = @BrandId
                      )"
                    cmd.Parameters.AddWithValue("@BrandId", brandId)
                    cmd.Parameters.AddWithValue("@ProductId", productId)
                    cmd.Parameters.AddWithValue("@ColorId", colorId)
                    cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                    cmd.ExecuteNonQuery()
                End Using
                ' Set selected to active
                Using cmd = conn.CreateCommand()
                    cmd.Transaction = trans
                    cmd.CommandText = "
                    UPDATE spnd
                    SET spnd.IsActiveForMarketplace = 1
                    FROM SupplierProductNameData spnd
                    INNER JOIN JoinProductColorFabricType j ON spnd.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                    WHERE j.FK_FabricBrandProductNameId = @ProductId
                      AND j.FK_ColorNameID = @ColorId
                      AND j.FK_FabricTypeNameId = @FabricTypeId
                      AND EXISTS (
                          SELECT 1 FROM FabricBrandProductName p
                          WHERE p.PK_FabricBrandProductNameId = j.FK_FabricBrandProductNameId
                            AND p.FK_FabricBrandNameId = @BrandId
                      )
                      AND spnd.FK_SupplierNameId = @SupplierId"
                    cmd.Parameters.AddWithValue("@BrandId", brandId)
                    cmd.Parameters.AddWithValue("@ProductId", productId)
                    cmd.Parameters.AddWithValue("@ColorId", colorId)
                    cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                    cmd.Parameters.AddWithValue("@SupplierId", activeSupplierId)
                    cmd.ExecuteNonQuery()
                End Using
                trans.Commit()
            End Using
        End Using
    End Sub
    Public Function GetSuppliersForCombination(brandId As Integer, productId As Integer, colorId As Integer, fabricTypeId As Integer) As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT 
                    s.PK_SupplierNameId, 
                    s.CompanyName,
                    spnd.IsActiveForMarketplace,
                    fp.CostPerSquareInch
                FROM SupplierInformation s
                INNER JOIN SupplierProductNameData spnd ON s.PK_SupplierNameId = spnd.FK_SupplierNameId
                INNER JOIN JoinProductColorFabricType j ON spnd.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                OUTER APPLY (
                    SELECT TOP 1 fph.CostPerSquareInch
                    FROM FabricPricingHistory fph
                    WHERE fph.FK_SupplierProductNameDataId = spnd.PK_SupplierProductNameDataId
                    ORDER BY fph.DateFrom DESC
                ) fp
                WHERE j.FK_FabricBrandProductNameId = @ProductId
                  AND j.FK_ColorNameID = @ColorId
                  AND j.FK_FabricTypeNameId = @FabricTypeId
                  AND EXISTS (
                      SELECT 1 FROM FabricBrandProductName p
                      WHERE p.PK_FabricBrandProductNameId = j.FK_FabricBrandProductNameId
                        AND p.FK_FabricBrandNameId = @BrandId
                  )"
                cmd.Parameters.AddWithValue("@BrandId", brandId)
                cmd.Parameters.AddWithValue("@ProductId", productId)
                cmd.Parameters.AddWithValue("@ColorId", colorId)
                cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                Using reader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return dt
    End Function
    Public Function GetProductsForSupplierAndBrand(supplierId As Integer, brandId As Integer) As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT DISTINCT p.PK_FabricBrandProductNameId, p.BrandProductName
                FROM SupplierProductNameData s
                INNER JOIN JoinProductColorFabricType j ON s.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                INNER JOIN FabricBrandProductName p ON j.FK_FabricBrandProductNameId = p.PK_FabricBrandProductNameId
                WHERE s.FK_SupplierNameId = @SupplierId AND p.FK_FabricBrandNameId = @BrandId
                ORDER BY p.BrandProductName"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                cmd.Parameters.AddWithValue("@BrandId", brandId)
                Using reader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return dt
    End Function
    ' Inserts a new pricing history record for a supplier-product
    Public Sub InsertFabricPricingHistory(
    supplierProductNameDataId As Integer,
    shippingCost As Decimal,
    costPerLinearYard As Decimal,
    costPerSquareInch As Decimal,
    weightPerSquareInch As Decimal
)
        Using conn = GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                INSERT INTO FabricPricingHistory
                (FK_SupplierProductNameDataId, DateFrom, ShippingCost, CostPerLinearYard, CostPerSquareInch, WeightPerSquareInch)
                VALUES (@SupplierProductNameDataId, @DateFrom, @ShippingCost, @CostPerLinearYard, @CostPerSquareInch, @WeightPerSquareInch)"
                cmd.Parameters.AddWithValue("@SupplierProductNameDataId", supplierProductNameDataId)
                cmd.Parameters.AddWithValue("@DateFrom", Date.Now)
                cmd.Parameters.AddWithValue("@ShippingCost", shippingCost)
                cmd.Parameters.AddWithValue("@CostPerLinearYard", costPerLinearYard)
                cmd.Parameters.AddWithValue("@CostPerSquareInch", costPerSquareInch)
                cmd.Parameters.AddWithValue("@WeightPerSquareInch", weightPerSquareInch)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    ' Updates WeightPerLinearYard and FabricRollWidth for a product
    Public Sub UpdateFabricProductInfo(
    productId As Integer,
    weightPerLinearYard As Decimal,
    fabricRollWidth As Decimal
)
        Using conn = GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                UPDATE FabricBrandProductName
                SET WeightPerLinearYard = @WeightPerLinearYard,
                    FabricRollWidth = @FabricRollWidth
                WHERE PK_FabricBrandProductNameId = @ProductId"
                cmd.Parameters.AddWithValue("@WeightPerLinearYard", weightPerLinearYard)
                cmd.Parameters.AddWithValue("@FabricRollWidth", fabricRollWidth)
                cmd.Parameters.AddWithValue("@ProductId", productId)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Public Function GetBrandsForSupplier(supplierId As Integer) As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT DISTINCT b.PK_FabricBrandNameId, b.BrandName
                FROM SupplierProductNameData s
                INNER JOIN JoinProductColorFabricType j ON s.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                INNER JOIN FabricBrandProductName p ON j.FK_FabricBrandProductNameId = p.PK_FabricBrandProductNameId
                INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                WHERE s.FK_SupplierNameId = @SupplierId"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                Using adapter As New SqlDataAdapter(CType(cmd, SqlCommand))
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function
    ' 1. Get all suppliers
    Public Function GetAllSuppliers() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_SupplierNameId AS SupplierID, CompanyName AS SupplierName FROM SupplierInformation"
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 2. Get all fabric brand names
    Public Function GetAllFabricBrandNames() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_FabricBrandNameId, BrandName FROM FabricBrandName"
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 3. Get all fabric types
    Public Function GetAllFabricTypes() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_FabricTypeNameId, FabricType FROM FabricTypeName"
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 4. Get products by brand name (for DataGridView dropdown)
    Public Function GetProductsByBrandName(brandName As String) As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT p.PK_FabricBrandProductNameId, p.BrandProductName AS ProductName
                                   FROM FabricBrandProductName p
                                   INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                                   WHERE b.BrandName = @BrandName"
                cmd.Parameters.AddWithValue("@BrandName", brandName)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 5. Get products by brand ID (for ComboBox)
    Public Function GetProductsByBrandId(brandId As Integer) As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_FabricBrandProductNameId, BrandProductName FROM FabricBrandProductName WHERE FK_FabricBrandNameId = @BrandId"
                cmd.Parameters.AddWithValue("@BrandId", brandId)
                Using reader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' 6. Get product info by brand and product name (for DataGridView)
    Public Function GetFabricProductInfo(brandName As String, productName As String) As DataRow
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT p.*, b.BrandName
                                   FROM FabricBrandProductName p
                                   INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                                   WHERE b.BrandName = @BrandName AND p.BrandProductName = @ProductName"
                cmd.Parameters.AddWithValue("@BrandName", brandName)
                cmd.Parameters.AddWithValue("@ProductName", productName)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    If dt.Rows.Count > 0 Then Return dt.Rows(0) Else Return Nothing
                End Using
            End Using
        End Using
    End Function

    ' 7. Get product info by product ID (for ComboBox and details)
    Public Function GetFabricProductInfoById(productId As Integer) As DataRow
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT p.*, b.BrandName
                                   FROM FabricBrandProductName p
                                   INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                                   WHERE p.PK_FabricBrandProductNameId = @ProductId"
                cmd.Parameters.AddWithValue("@ProductId", productId)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    If dt.Rows.Count > 0 Then Return dt.Rows(0) Else Return Nothing
                End Using
            End Using
        End Using
    End Function

    ' 8. Get latest pricing history by supplier, brand, and product name (for DataGridView)
    Public Function GetFabricPricingHistory(supplierId As Integer, brandName As String, productName As String) As DataRow
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT TOP 1 h.*
                FROM FabricPricingHistory h
                INNER JOIN SupplierProductNameData s ON h.FK_SupplierProductNameDataId = s.PK_SupplierProductNameDataId
                INNER JOIN JoinProductColorFabricType j ON s.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                INNER JOIN FabricBrandProductName p ON j.FK_FabricBrandProductNameId = p.PK_FabricBrandProductNameId
                INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                WHERE s.FK_SupplierNameId = @SupplierId AND b.BrandName = @BrandName AND p.BrandProductName = @ProductName
                ORDER BY h.DateFrom DESC"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                cmd.Parameters.AddWithValue("@BrandName", brandName)
                cmd.Parameters.AddWithValue("@ProductName", productName)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    If dt.Rows.Count > 0 Then Return dt.Rows(0) Else Return Nothing
                End Using
            End Using
        End Using
    End Function

    ' 9. Get latest pricing history by supplier and product ID (for ComboBox/details)
    Public Function GetFabricPricingHistoryByProductId(supplierId As Integer, productId As Integer, colorId As Integer, fabricTypeId As Integer) As DataRow
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT TOP 1 h.*
                FROM FabricPricingHistory h
                INNER JOIN SupplierProductNameData s ON h.FK_SupplierProductNameDataId = s.PK_SupplierProductNameDataId
                INNER JOIN JoinProductColorFabricType j ON s.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                WHERE s.FK_SupplierNameId = @SupplierId
                  AND j.FK_FabricBrandProductNameId = @ProductId
                  AND j.FK_ColorNameID = @ColorId
                  AND j.FK_FabricTypeNameId = @FabricTypeId
                ORDER BY h.DateFrom DESC"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                cmd.Parameters.AddWithValue("@ProductId", productId)
                cmd.Parameters.AddWithValue("@ColorId", colorId)
                cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    If dt.Rows.Count > 0 Then Return dt.Rows(0) Else Return Nothing
                End Using
            End Using
        End Using
    End Function

    ' 10. Get supplier-product data (for ComboBox/details)
    Public Function GetSupplierProductNameData(supplierId As Integer, productId As Integer, colorId As Integer, fabricTypeId As Integer) As DataRow
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT s.*
                FROM SupplierProductNameData s
                INNER JOIN JoinProductColorFabricType j ON s.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                WHERE s.FK_SupplierNameId = @SupplierId
                  AND j.FK_FabricBrandProductNameId = @ProductId
                  AND j.FK_ColorNameID = @ColorId
                  AND j.FK_FabricTypeNameId = @FabricTypeId"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                cmd.Parameters.AddWithValue("@ProductId", productId)
                cmd.Parameters.AddWithValue("@ColorId", colorId)
                cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    If dt.Rows.Count > 0 Then Return dt.Rows(0) Else Return Nothing
                End Using
            End Using
        End Using
    End Function

    Public Function GetLatestModelHistoryCostRetailPricing(modelId As Integer) As DataRow
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT TOP 1 *
                FROM ModelHistoryCostRetailPricing
                WHERE FK_ModelId = @ModelId
                ORDER BY DateCalculated DESC"
                cmd.Parameters.AddWithValue("@ModelId", modelId)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    If dt.Rows.Count > 0 Then Return dt.Rows(0) Else Return Nothing
                End Using
            End Using
        End Using
    End Function
    ' Add these methods to your DbConnectionManager class
    Public Function GetActiveCostPerSquareInch(materialType As String) As Decimal
        Using conn = GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT TOP 1 fp.CostPerSquareInch
                FROM SupplierProductNameData s
                INNER JOIN JoinProductColorFabricType j ON s.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                INNER JOIN FabricBrandProductName p ON j.FK_FabricBrandProductNameId = p.PK_FabricBrandProductNameId
                INNER JOIN FabricTypeName ft ON j.FK_FabricTypeNameId = ft.PK_FabricTypeNameId
                OUTER APPLY (
                    SELECT TOP 1 fph.CostPerSquareInch
                    FROM FabricPricingHistory fph
                    WHERE fph.FK_SupplierProductNameDataId = s.PK_SupplierProductNameDataId
                    ORDER BY fph.DateFrom DESC
                ) fp
                WHERE ft.FabricType = @MaterialType
                  AND s.IsActiveForMarketplace = 1"
                cmd.Parameters.AddWithValue("@MaterialType", materialType)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    Return Convert.ToDecimal(result)
                End If
            End Using
        End Using
        Return 0D
    End Function
    ' 1. Get all equipment types
    Public Function GetEquipmentTypes() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_EquipmentTypeId, EquipmentTypeName FROM ModelEquipmentTypes"
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 2. Get all manufacturer names
    Public Function GetManufacturerNames() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_ManufacturerId, ManufacturerName FROM ModelManufacturers ORDER BY ManufacturerName"
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    ' Set the primary key for easier DataRow lookup by ID
                    If dt.Columns.Contains("PK_ManufacturerId") Then
                        dt.PrimaryKey = New DataColumn() {dt.Columns("PK_ManufacturerId")}
                    End If
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 3. Insert a new manufacturer
    Public Function InsertManufacturer(manufacturerName As String) As Integer
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "INSERT INTO ModelManufacturers (ManufacturerName) VALUES (@Name); SELECT CAST(SCOPE_IDENTITY() AS int);"
                cmd.Parameters.AddWithValue("@Name", manufacturerName)
                Return CInt(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    ' 4. Get all series names for a manufacturer
    Public Function GetSeriesNamesByManufacturer(manufacturerId As Integer) As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_SeriesId, SeriesName FROM ModelSeries WHERE FK_ManufacturerId = @ManufacturerId"
                cmd.Parameters.AddWithValue("@ManufacturerId", manufacturerId)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 5. Insert a new series
    Public Function InsertSeries(seriesName As String, manufacturerId As Integer, equipmentTypeId As Integer) As Integer
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "INSERT INTO ModelSeries (SeriesName, FK_ManufacturerId, FK_EquipmentTypeId) VALUES (@Name, @ManufacturerId, @EquipmentTypeId); SELECT CAST(SCOPE_IDENTITY() AS int);"
                cmd.Parameters.AddWithValue("@Name", seriesName)
                cmd.Parameters.AddWithValue("@ManufacturerId", manufacturerId)
                cmd.Parameters.AddWithValue("@EquipmentTypeId", equipmentTypeId)
                Return CInt(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    ' 6. Get equipment type ID by manufacturer and series
    Public Function GetEquipmentTypeIdByManufacturerAndSeries(manufacturerId As Integer, seriesId As Integer) As Integer?
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT FK_EquipmentTypeId FROM ModelSeries WHERE PK_SeriesId = @SeriesId AND FK_ManufacturerId = @ManufacturerId"
                cmd.Parameters.AddWithValue("@SeriesId", seriesId)
                cmd.Parameters.AddWithValue("@ManufacturerId", manufacturerId)
                conn.Open() ' <-- REQUIRED!
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    Return CInt(result)
                Else
                    Return Nothing
                End If
            End Using
        End Using
    End Function

    ' 7. Get models by manufacturer and series
    Public Function GetModelsByManufacturerAndSeries(manufacturerId As Integer, seriesId As Integer) As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT m.*
                FROM Model m
                INNER JOIN ModelSeries s ON m.FK_SeriesId = s.PK_SeriesId
                WHERE m.FK_SeriesId = @SeriesId AND s.FK_ManufacturerId = @ManufacturerId"
                cmd.Parameters.AddWithValue("@SeriesId", seriesId)
                cmd.Parameters.AddWithValue("@ManufacturerId", manufacturerId)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 8. Check if a model exists
    Public Function ModelExists(manufacturerId As Integer, seriesId As Integer, modelName As String) As Boolean
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT COUNT(*) FROM Model WHERE FK_seriesId = @SeriesId AND modelName = @ModelName"
                cmd.Parameters.AddWithValue("@SeriesId", seriesId)
                cmd.Parameters.AddWithValue("@ModelName", modelName)
                Return CInt(cmd.ExecuteScalar()) > 0
            End Using
        End Using
    End Function

    ' 9. Insert a new model
    Public Function InsertModel(
    manufacturerId As Integer,
    seriesId As Integer,
    modelName As String,
    width As Decimal,
    depth As Decimal,
    height As Decimal,
    totalFabricSquareInches As Decimal,
    OptionalHeight As Object,
    OptionalDepth As Object,
    AngleType As Object,
    AmpHandleLocation As Object,
    TAHWidth As Object,
    TAHHeight As Object,
    SAHWidth As Object,
    SAHHeight As Object,
    TAHRearOffset As Object,
    SAHRearOffset As Object,
    SAHTopDownOffset As Object,
    MusicRestDesign As Object,
    Chart_Template As Object,
    Notes As Object,
    parentSku As String
) As Integer
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "INSERT INTO Model (FK_SeriesId, ModelName, Width, Depth, Height, TotalFabricSquareInches, OptionalHeight, OptionalDepth, AngleType, AmpHandleLocation, TAHWidth, TAHHeight, SAHWidth, SAHHeight, TAHRearOffset, SAHRearOffset, SAHTopDownOffset, MusicRestDesign, Chart_Template, Notes, ParentSKU) " &
                              "VALUES (@SeriesId, @ModelName, @Width, @Depth, @Height, @TotalFabricSquareInches, @OptionalHeight, @OptionalDepth, @AngleType, @AmpHandleLocation, @TAHWidth, @TAHHeight, @SAHWidth, @SAHHeight, @TAHRearOffset, @SAHRearOffset, @SAHTopDownOffset, @MusicRestDesign, @Chart_Template, @Notes, @ParentSKU); " &
                              "SELECT CAST(SCOPE_IDENTITY() AS int);"
                cmd.Parameters.AddWithValue("@SeriesId", seriesId)
                cmd.Parameters.AddWithValue("@ModelName", modelName)
                cmd.Parameters.AddWithValue("@Width", width)
                cmd.Parameters.AddWithValue("@Depth", depth)
                cmd.Parameters.AddWithValue("@Height", height)
                cmd.Parameters.AddWithValue("@TotalFabricSquareInches", totalFabricSquareInches)
                cmd.Parameters.AddWithValue("@OptionalHeight", If(OptionalHeight Is Nothing, DBNull.Value, OptionalHeight))
                cmd.Parameters.AddWithValue("@OptionalDepth", If(OptionalDepth Is Nothing, DBNull.Value, OptionalDepth))
                cmd.Parameters.AddWithValue("@AngleType", If(AngleType Is Nothing, DBNull.Value, AngleType))
                cmd.Parameters.AddWithValue("@AmpHandleLocation", If(AmpHandleLocation Is Nothing, DBNull.Value, AmpHandleLocation))
                cmd.Parameters.AddWithValue("@TAHWidth", If(TAHWidth Is Nothing, DBNull.Value, TAHWidth))
                cmd.Parameters.AddWithValue("@TAHHeight", If(TAHHeight Is Nothing, DBNull.Value, TAHHeight))
                cmd.Parameters.AddWithValue("@SAHWidth", If(SAHWidth Is Nothing, DBNull.Value, SAHWidth))
                cmd.Parameters.AddWithValue("@SAHHeight", If(SAHHeight Is Nothing, DBNull.Value, SAHHeight))
                cmd.Parameters.AddWithValue("@TAHRearOffset", If(TAHRearOffset Is Nothing, DBNull.Value, TAHRearOffset))
                cmd.Parameters.AddWithValue("@SAHRearOffset", If(SAHRearOffset Is Nothing, DBNull.Value, SAHRearOffset))
                cmd.Parameters.AddWithValue("@SAHTopDownOffset", If(SAHTopDownOffset Is Nothing, DBNull.Value, SAHTopDownOffset))
                cmd.Parameters.AddWithValue("@MusicRestDesign", If(MusicRestDesign Is Nothing, DBNull.Value, MusicRestDesign))
                cmd.Parameters.AddWithValue("@Chart_Template", If(Chart_Template Is Nothing, DBNull.Value, Chart_Template))
                cmd.Parameters.AddWithValue("@Notes", If(Notes Is Nothing, DBNull.Value, Notes))
                cmd.Parameters.AddWithValue("@ParentSKU", parentSku)
                Return CInt(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    ' Update an existing model, including TotalFabricSquareInches and Notes
    ' Update an existing model, including ParentSKU
    Public Function UpdateModel(
    modelId As Integer,
    modelName As String,
    width As Decimal,
    depth As Decimal,
    height As Decimal,
    totalFabricSquareInches As Decimal,
    OptionalHeight As Object,
    OptionalDepth As Object,
    AngleType As Object,
    AmpHandleLocation As Object,
    TAHWidth As Object,
    TAHHeight As Object,
    SAHWidth As Object,
    SAHHeight As Object,
    TAHRearOffset As Object,
    SAHRearOffset As Object,
    SAHTopDownOffset As Object,
    MusicRestDesign As Object,
    Chart_Template As Object,
    Notes As Object,
    parentSku As String
) As Boolean
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "UPDATE Model SET " &
                "ModelName=@ModelName, Width=@Width, Depth=@Depth, Height=@Height, " &
                "TotalFabricSquareInches=@TotalFabricSquareInches, OptionalHeight=@OptionalHeight, OptionalDepth=@OptionalDepth, " &
                "AngleType=@AngleType, AmpHandleLocation=@AmpHandleLocation, TAHWidth=@TAHWidth, TAHHeight=@TAHHeight, " &
                "SAHWidth=@SAHWidth, SAHHeight=@SAHHeight, TAHRearOffset=@TAHRearOffset, SAHRearOffset=@SAHRearOffset, " &
                "SAHTopDownOffset=@SAHTopDownOffset, MusicRestDesign=@MusicRestDesign, Chart_Template=@Chart_Template, " &
                "Notes=@Notes, ParentSKU=@ParentSKU " &
                "WHERE PK_ModelId=@ModelId"
                cmd.Parameters.AddWithValue("@ModelId", modelId)
                cmd.Parameters.AddWithValue("@ModelName", modelName)
                cmd.Parameters.AddWithValue("@Width", width)
                cmd.Parameters.AddWithValue("@Depth", depth)
                cmd.Parameters.AddWithValue("@Height", height)
                cmd.Parameters.AddWithValue("@TotalFabricSquareInches", totalFabricSquareInches)
                cmd.Parameters.AddWithValue("@OptionalHeight", If(OptionalHeight Is Nothing, DBNull.Value, OptionalHeight))
                cmd.Parameters.AddWithValue("@OptionalDepth", If(OptionalDepth Is Nothing, DBNull.Value, OptionalDepth))
                cmd.Parameters.AddWithValue("@AngleType", If(AngleType Is Nothing, DBNull.Value, AngleType))
                cmd.Parameters.AddWithValue("@AmpHandleLocation", If(AmpHandleLocation Is Nothing, DBNull.Value, AmpHandleLocation))
                cmd.Parameters.AddWithValue("@TAHWidth", If(TAHWidth Is Nothing, DBNull.Value, TAHWidth))
                cmd.Parameters.AddWithValue("@TAHHeight", If(TAHHeight Is Nothing, DBNull.Value, TAHHeight))
                cmd.Parameters.AddWithValue("@SAHWidth", If(SAHWidth Is Nothing, DBNull.Value, SAHWidth))
                cmd.Parameters.AddWithValue("@SAHHeight", If(SAHHeight Is Nothing, DBNull.Value, SAHHeight))
                cmd.Parameters.AddWithValue("@TAHRearOffset", If(TAHRearOffset Is Nothing, DBNull.Value, TAHRearOffset))
                cmd.Parameters.AddWithValue("@SAHRearOffset", If(SAHRearOffset Is Nothing, DBNull.Value, SAHRearOffset))
                cmd.Parameters.AddWithValue("@SAHTopDownOffset", If(SAHTopDownOffset Is Nothing, DBNull.Value, SAHTopDownOffset))
                cmd.Parameters.AddWithValue("@MusicRestDesign", If(MusicRestDesign Is Nothing, DBNull.Value, MusicRestDesign))
                cmd.Parameters.AddWithValue("@Chart_Template", If(Chart_Template Is Nothing, DBNull.Value, Chart_Template))
                cmd.Parameters.AddWithValue("@Notes", If(Notes Is Nothing, DBNull.Value, Notes))
                cmd.Parameters.AddWithValue("@ParentSKU", parentSku)
                Return cmd.ExecuteNonQuery() > 0
            End Using
        End Using
    End Function


    Public Function GetImagesForEquipmentTypeAndMarketplace(equipmentTypeId As Integer, marketplaceNameId As Integer) As DataTable
        Dim dt As New DataTable()
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Using cmd As New SqlCommand("SELECT ImageUrl, ImageType, AltText, Position FROM ModelEquipmentTypeImage WHERE FK_equipmentTypeId = @equipmentTypeId AND FK_marketplaceNameId = @marketplaceNameId AND IsActive = 1 ORDER BY Position", conn)
                cmd.Parameters.AddWithValue("@equipmentTypeId", equipmentTypeId)
                cmd.Parameters.AddWithValue("@marketplaceNameId", marketplaceNameId)

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return dt
    End Function
    Public Sub UpdateMissingParentSKUs()
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            ' 1. Get all models missing a ParentSKU
            Dim selectCmd As New SqlCommand("
        SELECT m.PK_ModelId, m.ModelName, s.SeriesName, mf.ManufacturerName
        FROM Model m
        INNER JOIN ModelSeries s ON m.FK_SeriesId = s.PK_SeriesId
        INNER JOIN ModelManufacturers mf ON s.FK_ManufacturerId = mf.PK_ManufacturerId
        WHERE m.ParentSKU IS NULL OR m.ParentSKU = ''", conn)

            Using reader = selectCmd.ExecuteReader()
                Dim updates As New List(Of (Integer, String))
                While reader.Read()
                    Dim modelId = CInt(reader("PK_ModelId"))
                    Dim manufacturerName = reader("ManufacturerName").ToString()
                    Dim seriesName = reader("SeriesName").ToString()
                    Dim modelName = reader("ModelName").ToString()
                    ' Always include modelId in SKU
                    Dim parentSku = ModelSkuBuilder.GenerateParentSku(manufacturerName, seriesName, modelName, "V1", modelId)
                    updates.Add((modelId, parentSku))
                End While
                reader.Close()

                ' 2. Update each model with the generated ParentSKU
                For Each tup In updates
                    Dim updateCmd As New SqlCommand("UPDATE Model SET ParentSKU = @ParentSKU WHERE PK_ModelId = @ModelId", conn)
                    updateCmd.Parameters.AddWithValue("@ParentSKU", tup.Item2)
                    updateCmd.Parameters.AddWithValue("@ModelId", tup.Item1)
                    updateCmd.ExecuteNonQuery()
                Next
            End Using
        End Using
    End Sub
    Public Function UpdateSeriesEquipmentType(seriesId As Integer, equipmentTypeId As Integer) As Boolean
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "UPDATE ModelSeries SET FK_EquipmentTypeId = @EquipmentTypeId WHERE PK_SeriesId = @SeriesId"
                cmd.Parameters.AddWithValue("@EquipmentTypeId", equipmentTypeId)
                cmd.Parameters.AddWithValue("@SeriesId", seriesId)
                Return cmd.ExecuteNonQuery() > 0
            End Using
        End Using
    End Function

    Public Function GetSeriesAndEquipmentTypeByManufacturer(manufacturerId As Integer) As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT 
                    s.PK_SeriesId, 
                    s.SeriesName, 
                    e.EquipmentTypeName
                FROM 
                    ModelSeries s
                    INNER JOIN ModelEquipmentTypes e ON s.FK_EquipmentTypeId = e.PK_EquipmentTypeId
                WHERE 
                    s.FK_ManufacturerId = @ManufacturerId
                ORDER BY s.SeriesName"
                cmd.Parameters.AddWithValue("@ManufacturerId", manufacturerId)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    Public Function GetAllAngleTypesTable() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_AngleTypeId, AngleTypeName FROM AngleType"
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    Public Function GetAllAmpHandleLocationsTable() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_AmpHandleLocationId, AmpHandleLocationName FROM AmpHandleLocation"
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    Public Function GetActiveChoiceWaterproofCostPerSqInch() As Decimal?
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT TOP 1 fp.CostPerSquareInch
                FROM SupplierProductNameData spnd
                INNER JOIN JoinProductColorFabricType j ON spnd.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                INNER JOIN FabricTypeName ft ON j.FK_FabricTypeNameId = ft.PK_FabricTypeNameId
                OUTER APPLY (
                    SELECT TOP 1 fph.CostPerSquareInch
                    FROM FabricPricingHistory fph
                    WHERE fph.FK_SupplierProductNameDataId = spnd.PK_SupplierProductNameDataId
                    ORDER BY fph.DateFrom DESC
                ) fp
                WHERE ft.FabricType = 'Choice Waterproof'
                  AND spnd.IsActiveForMarketplace = 1
                ORDER BY fp.CostPerSquareInch"
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    Return Convert.ToDecimal(result)
                End If
            End Using
        End Using
        Return Nothing
    End Function
    ' Returns the first FabricBrandProductName row for the given fabric type name.
    Public Function GetActiveFabricBrandProductName(fabricTypeName As String) As DataRow
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT TOP 1 fbp.*
                FROM FabricBrandProductName fbp
                INNER JOIN FabricTypeName ftn ON fbp.FK_FabricTypeNameId = ftn.PK_FabricTypeNameId
                WHERE ftn.FabricType = @FabricTypeName
                ORDER BY fbp.PK_FabricBrandProductNameId"
                cmd.Parameters.AddWithValue("@FabricTypeName", fabricTypeName)
                Using da As New SqlClient.SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    If dt.Rows.Count > 0 Then Return dt.Rows(0)
                End Using
            End Using
        End Using
        Return Nothing
    End Function


    Public Function GetShippingCostByWeight(weightOz As Decimal) As Decimal
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT TOP 1 ShippingCost
                FROM CostShipping
                WHERE @WeightOz >= MinWeight AND @WeightOz <= MaxWeight
                ORDER BY MinWeight ASC"
                cmd.Parameters.AddWithValue("@WeightOz", weightOz)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    Return Convert.ToDecimal(result)
                End If
            End Using
        End Using
        Return 0D
    End Function

    Public Function GetLatestProfitValue(coverType As String) As Decimal
        Dim sql = "SELECT TOP 1 profitValue FROM Profit WHERE coverType = @coverType ORDER BY effectiveDate DESC"
        Using conn = GetConnection()
            conn.Open()
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@coverType", coverType)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    Return Convert.ToDecimal(result)
                Else
                    Return 0D
                End If
            End Using
        End Using
    End Function

    Public Function GetMarketplaceFeePercentage(marketplaceName As String) As Decimal
        Dim sql = "
        SELECT TOP 1 mf.listingFeePercent
        FROM MarketplaceFees mf
        INNER JOIN MarketplaceName mn ON mf.FK_marketplaceNameId = mn.PK_marketplaceNameId
        WHERE mn.marketplaceName = @MarketplaceName
        ORDER BY mf.effectiveDate DESC"
        Using conn = GetConnection()
            conn.Open()
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@MarketplaceName", marketplaceName)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    Return Convert.ToDecimal(result)
                Else
                    Return 0D
                End If
            End Using
        End Using
    End Function
    ' Returns the ManufacturerId for a given name, or 0 if not found
    Public Function GetManufacturerIdByName(manufacturerName As String) As Integer
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_ManufacturerId FROM ModelManufacturers WHERE ManufacturerName = @Name"
                cmd.Parameters.AddWithValue("@Name", manufacturerName)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    Return CInt(result)
                Else
                    Return 0
                End If
            End Using
        End Using
    End Function
    ' Returns the SeriesId for a given name and manufacturer, or 0 if not found
    Public Function GetSeriesIdByNameAndManufacturer(seriesName As String, manufacturerId As Integer) As Integer
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_SeriesId FROM ModelSeries WHERE SeriesName = @SeriesName AND FK_ManufacturerId = @ManufacturerId"
                cmd.Parameters.AddWithValue("@SeriesName", seriesName)
                cmd.Parameters.AddWithValue("@ManufacturerId", manufacturerId)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    Return CInt(result)
                Else
                    Return 0
                End If
            End Using
        End Using
    End Function
    Public Sub InsertModelHistoryCostRetailPricing(
    modelId As Integer,
    costPerSqInch_ChoiceWaterproof As Decimal?,
    costPerSqInch_PremiumSyntheticLeather As Decimal?,
    costPerSqInch_Padding As Decimal?,
    totalFabricSquareInches As Decimal,
    wastePercent As Decimal,
    baseCost_ChoiceWaterproof As Decimal?,
    baseCost_PremiumSyntheticLeather As Decimal?,
    baseCost_ChoiceWaterproof_Padded As Decimal?,
    baseCost_PremiumSyntheticLeather_Padded As Decimal?,
    baseCost_PaddingOnly As Decimal?,
    baseWeight_PaddingOnly As Decimal?,
    baseWeight_ChoiceWaterproof As Decimal?,
    baseWeight_ChoiceWaterproof_Padded As Decimal?,
    baseWeight_PremiumSyntheticLeather As Decimal?,
    baseWeight_PremiumSyntheticLeather_Padded As Decimal?,
    FabricWeight_Choice_Cost As Decimal?,
    FabricWeight_ChoicePadding_Cost As Decimal?,
    FabricWeight_Leather_Cost As Decimal?,
    FabricWeight_LeatherPadding_Cost As Decimal?,
    baseFabricCost_Choice_Weight As Decimal?,
    baseFabricCost_ChoicePadding_Weight As Decimal?,
    baseFabricCost_Leather_Weight As Decimal?,
    baseFabricCost_LeatherPadding_Weight As Decimal?,
    BaseCost_Choice_Labor As Decimal?,
    BaseCost_ChoicePadding_Labor As Decimal?,
    BaseCost_Leather_Labor As Decimal?,
    BaseCost_LeatherPadding_Labor As Decimal?,
    profit_Choice As Decimal,
    profit_ChoicePadded As Decimal,
    profit_Leather As Decimal,
    profit_LeatherPadded As Decimal,
    AmazonFee_Choice As Decimal,
    AmazonFee_ChoicePadded As Decimal,
    AmazonFee_Leather As Decimal,
    AmazonFee_LeatherPadded As Decimal,
    ReverbFee_Choice As Decimal,
    ReverbFee_ChoicePadded As Decimal,
    ReverbFee_Leather As Decimal,
    ReverbFee_LeatherPadded As Decimal,
    eBayFee_Choice As Decimal,
    eBayFee_ChoicePadded As Decimal,
    eBayFee_Leather As Decimal,
    eBayFee_LeatherPadded As Decimal,
    EtsyFee_Choice As Decimal,
    EtsyFee_ChoicePadded As Decimal,
    EtsyFee_Leather As Decimal,
    EtsyFee_LeatherPadded As Decimal,
    BaseCost_GrandTotal_Choice_Amazon As Decimal,
    BaseCost_GrandTotal_ChoicePadded_Amazon As Decimal,
    BaseCost_GrandTotal_Leather_Amazon As Decimal,
    BaseCost_GrandTotal_LeatherPadded_Amazon As Decimal,
    BaseCost_GrandTotal_Choice_Reverb As Decimal,
    BaseCost_GrandTotal_ChoicePadded_Reverb As Decimal,
    BaseCost_GrandTotal_Leather_Reverb As Decimal,
    BaseCost_GrandTotal_LeatherPadded_Reverb As Decimal,
    BaseCost_GrandTotal_Choice_eBay As Decimal,
    BaseCost_GrandTotal_ChoicePadded_eBay As Decimal,
    BaseCost_GrandTotal_Leather_eBay As Decimal,
    BaseCost_GrandTotal_LeatherPadded_eBay As Decimal,
    BaseCost_GrandTotal_Choice_Etsy As Decimal,
    BaseCost_GrandTotal_ChoicePadded_Etsy As Decimal,
    BaseCost_GrandTotal_Leather_Etsy As Decimal,
    BaseCost_GrandTotal_LeatherPadded_Etsy As Decimal,
    RetailPrice_Choice_Amazon As Decimal,
    RetailPrice_ChoicePadded_Amazon As Decimal,
    RetailPrice_Leather_Amazon As Decimal,
    RetailPrice_LeatherPadded_Amazon As Decimal,
    RetailPrice_Choice_Reverb As Decimal,
    RetailPrice_ChoicePadded_Reverb As Decimal,
    RetailPrice_Leather_Reverb As Decimal,
    RetailPrice_LeatherPadded_Reverb As Decimal,
    RetailPrice_Choice_eBay As Decimal,
    RetailPrice_ChoicePadded_eBay As Decimal,
    RetailPrice_Leather_eBay As Decimal,
    RetailPrice_LeatherPadded_eBay As Decimal,
    RetailPrice_Choice_Etsy As Decimal,
    RetailPrice_ChoicePadded_Etsy As Decimal,
    RetailPrice_Leather_Etsy As Decimal,
    RetailPrice_LeatherPadded_Etsy As Decimal,
    notes As String
)
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                INSERT INTO ModelHistoryCostRetailPricing
                (
                    FK_ModelId,
                    CostPerSqInch_ChoiceWaterproof,
                    CostPerSqInch_PremiumSyntheticLeather,
                    CostPerSqInch_Padding,
                    TotalFabricSquareInches,
                    WastePercent,
                    BaseFabricCost_ChoiceWaterproof,
                    BaseFabricCost_PremiumSyntheticLeather,
                    BaseFabricCost_ChoiceWaterproof_Padded,
                    BaseFabricCost_PremiumSyntheticLeather_Padded,
                    BaseFabricCost_PaddingOnly,
                    BaseWeight_PaddingOnly,
                    BaseWeight_ChoiceWaterproof,
                    BaseWeight_ChoiceWaterproof_Padded,
                    BaseWeight_PremiumSyntheticLeather,
                    BaseWeight_PremiumSyntheticLeather_Padded,
                    FabricWeight_Choice_Cost,
                    FabricWeight_ChoicePadding_Cost,
                    FabricWeight_Leather_Cost,
                    FabricWeight_LeatherPadding_Cost,
                    BaseFabricCost_Choice_Weight,
                    BaseFabricCost_ChoicePadding_Weight,
                    BaseFabricCost_Leather_Weight,
                    BaseFabricCost_LeatherPadding_Weight,
                    BaseCost_Choice_Labor,
                    BaseCost_ChoicePadding_Labor,
                    BaseCost_Leather_Labor,
                    BaseCost_LeatherPadding_Labor,
                    Profit_Choice,
                    Profit_ChoicePadded,
                    Profit_Leather,
                    Profit_LeatherPadded,
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
                    EtsyFee_Choice,
                    EtsyFee_ChoicePadded,
                    EtsyFee_Leather,
                    EtsyFee_LeatherPadded,
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
                    DateCalculated,
                    Notes
                )
                VALUES (
                    @ModelId,
                    @CWSqInch,
                    @SLSqInch,
                    @PadSqInch,
                    @TotalSqIn,
                    @Waste,
                    @CWCost,
                    @SLCost,
                    @CWPadCost,
                    @SLPadCost,
                    @PadOnlyCost,
                    @PadOnlyWeight,
                    @CWWeight,
                    @CWPadWeight,
                    @SLWeight,
                    @SLPadWeight,
                    @FWChoice,
                    @FWChoicePad,
                    @FWLeather,
                    @FWLeatherPad,
                    @BaseCWWeight,
                    @BaseCWPadWeight,
                    @BaseSLWeight,
                    @BaseSLPadWeight,
                    @BaseCost_Choice_Labor,
                    @BaseCost_ChoicePadding_Labor,
                    @BaseCost_Leather_Labor,
                    @BaseCost_LeatherPadding_Labor,
                    @Profit_Choice,
                    @Profit_ChoicePadded,
                    @Profit_Leather,
                    @Profit_LeatherPadded,
                    @AmazonFee_Choice,
                    @AmazonFee_ChoicePadded,
                    @AmazonFee_Leather,
                    @AmazonFee_LeatherPadded,
                    @ReverbFee_Choice,
                    @ReverbFee_ChoicePadded,
                    @ReverbFee_Leather,
                    @ReverbFee_LeatherPadded,
                    @eBayFee_Choice,
                    @eBayFee_ChoicePadded,
                    @eBayFee_Leather,
                    @eBayFee_LeatherPadded,
                    @EtsyFee_Choice,
                    @EtsyFee_ChoicePadded,
                    @EtsyFee_Leather,
                    @EtsyFee_LeatherPadded,
                    @BaseCost_GrandTotal_Choice_Amazon,
                    @BaseCost_GrandTotal_ChoicePadded_Amazon,
                    @BaseCost_GrandTotal_Leather_Amazon,
                    @BaseCost_GrandTotal_LeatherPadded_Amazon,
                    @BaseCost_GrandTotal_Choice_Reverb,
                    @BaseCost_GrandTotal_ChoicePadded_Reverb,
                    @BaseCost_GrandTotal_Leather_Reverb,
                    @BaseCost_GrandTotal_LeatherPadded_Reverb,
                    @BaseCost_GrandTotal_Choice_eBay,
                    @BaseCost_GrandTotal_ChoicePadded_eBay,
                    @BaseCost_GrandTotal_Leather_eBay,
                    @BaseCost_GrandTotal_LeatherPadded_eBay,
                    @BaseCost_GrandTotal_Choice_Etsy,
                    @BaseCost_GrandTotal_ChoicePadded_Etsy,
                    @BaseCost_GrandTotal_Leather_Etsy,
                    @BaseCost_GrandTotal_LeatherPadded_Etsy,
                    @RetailPrice_Choice_Amazon,
                    @RetailPrice_ChoicePadded_Amazon,
                    @RetailPrice_Leather_Amazon,
                    @RetailPrice_LeatherPadded_Amazon,
                    @RetailPrice_Choice_Reverb,
                    @RetailPrice_ChoicePadded_Reverb,
                    @RetailPrice_Leather_Reverb,
                    @RetailPrice_LeatherPadded_Reverb,
                    @RetailPrice_Choice_eBay,
                    @RetailPrice_ChoicePadded_eBay,
                    @RetailPrice_Leather_eBay,
                    @RetailPrice_LeatherPadded_eBay,
                    @RetailPrice_Choice_Etsy,
                    @RetailPrice_ChoicePadded_Etsy,
                    @RetailPrice_Leather_Etsy,
                    @RetailPrice_LeatherPadded_Etsy,
                    GETDATE(),
                    @Notes
                )"
                cmd.Parameters.AddWithValue("@ModelId", modelId)
                cmd.Parameters.AddWithValue("@CWSqInch", If(costPerSqInch_ChoiceWaterproof, DBNull.Value))
                cmd.Parameters.AddWithValue("@SLSqInch", If(costPerSqInch_PremiumSyntheticLeather, DBNull.Value))
                cmd.Parameters.AddWithValue("@PadSqInch", If(costPerSqInch_Padding, DBNull.Value))
                cmd.Parameters.AddWithValue("@TotalSqIn", totalFabricSquareInches)
                cmd.Parameters.AddWithValue("@Waste", wastePercent)
                cmd.Parameters.AddWithValue("@CWCost", If(baseCost_ChoiceWaterproof, DBNull.Value))
                cmd.Parameters.AddWithValue("@SLCost", If(baseCost_PremiumSyntheticLeather, DBNull.Value))
                cmd.Parameters.AddWithValue("@CWPadCost", If(baseCost_ChoiceWaterproof_Padded, DBNull.Value))
                cmd.Parameters.AddWithValue("@SLPadCost", If(baseCost_PremiumSyntheticLeather_Padded, DBNull.Value))
                cmd.Parameters.AddWithValue("@PadOnlyCost", If(baseCost_PaddingOnly, DBNull.Value))
                cmd.Parameters.AddWithValue("@PadOnlyWeight", If(baseWeight_PaddingOnly, DBNull.Value))
                cmd.Parameters.AddWithValue("@CWWeight", If(baseWeight_ChoiceWaterproof, DBNull.Value))
                cmd.Parameters.AddWithValue("@CWPadWeight", If(baseWeight_ChoiceWaterproof_Padded, DBNull.Value))
                cmd.Parameters.AddWithValue("@SLWeight", If(baseWeight_PremiumSyntheticLeather, DBNull.Value))
                cmd.Parameters.AddWithValue("@SLPadWeight", If(baseWeight_PremiumSyntheticLeather_Padded, DBNull.Value))
                cmd.Parameters.AddWithValue("@FWChoice", If(FabricWeight_Choice_Cost, DBNull.Value))
                cmd.Parameters.AddWithValue("@FWChoicePad", If(FabricWeight_ChoicePadding_Cost, DBNull.Value))
                cmd.Parameters.AddWithValue("@FWLeather", If(FabricWeight_Leather_Cost, DBNull.Value))
                cmd.Parameters.AddWithValue("@FWLeatherPad", If(FabricWeight_LeatherPadding_Cost, DBNull.Value))
                cmd.Parameters.AddWithValue("@BaseCWWeight", If(baseFabricCost_Choice_Weight, DBNull.Value))
                cmd.Parameters.AddWithValue("@BaseCWPadWeight", If(baseFabricCost_ChoicePadding_Weight, DBNull.Value))
                cmd.Parameters.AddWithValue("@BaseSLWeight", If(baseFabricCost_Leather_Weight, DBNull.Value))
                cmd.Parameters.AddWithValue("@BaseSLPadWeight", If(baseFabricCost_LeatherPadding_Weight, DBNull.Value))
                cmd.Parameters.AddWithValue("@BaseCost_Choice_Labor", If(BaseCost_Choice_Labor, DBNull.Value))
                cmd.Parameters.AddWithValue("@BaseCost_ChoicePadding_Labor", If(BaseCost_ChoicePadding_Labor, DBNull.Value))
                cmd.Parameters.AddWithValue("@BaseCost_Leather_Labor", If(BaseCost_Leather_Labor, DBNull.Value))
                cmd.Parameters.AddWithValue("@BaseCost_LeatherPadding_Labor", If(BaseCost_LeatherPadding_Labor, DBNull.Value))
                cmd.Parameters.AddWithValue("@Profit_Choice", profit_Choice)
                cmd.Parameters.AddWithValue("@Profit_ChoicePadded", profit_ChoicePadded)
                cmd.Parameters.AddWithValue("@Profit_Leather", profit_Leather)
                cmd.Parameters.AddWithValue("@Profit_LeatherPadded", profit_LeatherPadded)
                cmd.Parameters.AddWithValue("@AmazonFee_Choice", AmazonFee_Choice)
                cmd.Parameters.AddWithValue("@AmazonFee_ChoicePadded", AmazonFee_ChoicePadded)
                cmd.Parameters.AddWithValue("@AmazonFee_Leather", AmazonFee_Leather)
                cmd.Parameters.AddWithValue("@AmazonFee_LeatherPadded", AmazonFee_LeatherPadded)
                cmd.Parameters.AddWithValue("@ReverbFee_Choice", ReverbFee_Choice)
                cmd.Parameters.AddWithValue("@ReverbFee_ChoicePadded", ReverbFee_ChoicePadded)
                cmd.Parameters.AddWithValue("@ReverbFee_Leather", ReverbFee_Leather)
                cmd.Parameters.AddWithValue("@ReverbFee_LeatherPadded", ReverbFee_LeatherPadded)
                cmd.Parameters.AddWithValue("@eBayFee_Choice", eBayFee_Choice)
                cmd.Parameters.AddWithValue("@eBayFee_ChoicePadded", eBayFee_ChoicePadded)
                cmd.Parameters.AddWithValue("@eBayFee_Leather", eBayFee_Leather)
                cmd.Parameters.AddWithValue("@eBayFee_LeatherPadded", eBayFee_LeatherPadded)
                cmd.Parameters.AddWithValue("@EtsyFee_Choice", EtsyFee_Choice)
                cmd.Parameters.AddWithValue("@EtsyFee_ChoicePadded", EtsyFee_ChoicePadded)
                cmd.Parameters.AddWithValue("@EtsyFee_Leather", EtsyFee_Leather)
                cmd.Parameters.AddWithValue("@EtsyFee_LeatherPadded", EtsyFee_LeatherPadded)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_Choice_Amazon", BaseCost_GrandTotal_Choice_Amazon)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_ChoicePadded_Amazon", BaseCost_GrandTotal_ChoicePadded_Amazon)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_Leather_Amazon", BaseCost_GrandTotal_Leather_Amazon)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_LeatherPadded_Amazon", BaseCost_GrandTotal_LeatherPadded_Amazon)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_Choice_Reverb", BaseCost_GrandTotal_Choice_Reverb)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_ChoicePadded_Reverb", BaseCost_GrandTotal_ChoicePadded_Reverb)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_Leather_Reverb", BaseCost_GrandTotal_Leather_Reverb)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_LeatherPadded_Reverb", BaseCost_GrandTotal_LeatherPadded_Reverb)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_Choice_eBay", BaseCost_GrandTotal_Choice_eBay)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_ChoicePadded_eBay", BaseCost_GrandTotal_ChoicePadded_eBay)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_Leather_eBay", BaseCost_GrandTotal_Leather_eBay)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_LeatherPadded_eBay", BaseCost_GrandTotal_LeatherPadded_eBay)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_Choice_Etsy", BaseCost_GrandTotal_Choice_Etsy)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_ChoicePadded_Etsy", BaseCost_GrandTotal_ChoicePadded_Etsy)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_Leather_Etsy", BaseCost_GrandTotal_Leather_Etsy)
                cmd.Parameters.AddWithValue("@BaseCost_GrandTotal_LeatherPadded_Etsy", BaseCost_GrandTotal_LeatherPadded_Etsy)
                cmd.Parameters.AddWithValue("@RetailPrice_Choice_Amazon", RetailPrice_Choice_Amazon)
                cmd.Parameters.AddWithValue("@RetailPrice_ChoicePadded_Amazon", RetailPrice_ChoicePadded_Amazon)
                cmd.Parameters.AddWithValue("@RetailPrice_Leather_Amazon", RetailPrice_Leather_Amazon)
                cmd.Parameters.AddWithValue("@RetailPrice_LeatherPadded_Amazon", RetailPrice_LeatherPadded_Amazon)
                cmd.Parameters.AddWithValue("@RetailPrice_Choice_Reverb", RetailPrice_Choice_Reverb)
                cmd.Parameters.AddWithValue("@RetailPrice_ChoicePadded_Reverb", RetailPrice_ChoicePadded_Reverb)
                cmd.Parameters.AddWithValue("@RetailPrice_Leather_Reverb", RetailPrice_Leather_Reverb)
                cmd.Parameters.AddWithValue("@RetailPrice_LeatherPadded_Reverb", RetailPrice_LeatherPadded_Reverb)
                cmd.Parameters.AddWithValue("@RetailPrice_Choice_eBay", RetailPrice_Choice_eBay)
                cmd.Parameters.AddWithValue("@RetailPrice_ChoicePadded_eBay", RetailPrice_ChoicePadded_eBay)
                cmd.Parameters.AddWithValue("@RetailPrice_Leather_eBay", RetailPrice_Leather_eBay)
                cmd.Parameters.AddWithValue("@RetailPrice_LeatherPadded_eBay", RetailPrice_LeatherPadded_eBay)
                cmd.Parameters.AddWithValue("@RetailPrice_Choice_Etsy", RetailPrice_Choice_Etsy)
                cmd.Parameters.AddWithValue("@RetailPrice_ChoicePadded_Etsy", RetailPrice_ChoicePadded_Etsy)
                cmd.Parameters.AddWithValue("@RetailPrice_Leather_Etsy", RetailPrice_Leather_Etsy)
                cmd.Parameters.AddWithValue("@RetailPrice_LeatherPadded_Etsy", RetailPrice_LeatherPadded_Etsy)
                cmd.Parameters.AddWithValue("@Notes", If(notes Is Nothing, DBNull.Value, notes))
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Public Function UpdateModelParentSku(modelId As Integer, parentSku As String) As Integer
        Using conn As New SqlConnection(DbConnectionManager.ConnectionString)
            conn.Open()
            Using cmd As New SqlCommand("UPDATE Model SET ParentSKU = @ParentSKU WHERE PK_ModelId = @ModelId", conn)
                cmd.Parameters.AddWithValue("@ParentSKU", parentSku)
                cmd.Parameters.AddWithValue("@ModelId", modelId)
                Return cmd.ExecuteNonQuery() ' Returns number of rows updated
            End Using
        End Using
    End Function
    Public Function GetWooCategoryIdByName(categoryName As String) As Integer
        Dim query As String = "SELECT wooCategoryId FROM MarketplaceWOOCOMMERCECategoryID WHERE wooCategoryName = @name"
        Using conn = GetConnection()
            conn.Open()
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@name", categoryName)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) AndAlso IsNumeric(result) Then
                    Return Convert.ToInt32(result)
                Else
                    Throw New Exception("WooCommerce category not found for: " & categoryName)
                End If
            End Using
        End Using
    End Function

    ' Gets the WooProductId for a given model (returns 0 if not set)
    Public Function GetWooProductIdByModelId(modelId As Integer) As Integer
        Dim query As String = "SELECT WooProductId FROM Model WHERE PK_ModelId = @modelId"
        Using conn = GetConnection()
            conn.Open()
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@modelId", modelId)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) AndAlso IsNumeric(result) Then
                    Return Convert.ToInt32(result)
                Else
                    Return 0
                End If
            End Using
        End Using
    End Function

    ' Saves the WooProductId for a given model
    Public Sub SetWooProductIdForModel(modelId As Integer, wooProductId As Integer)
        Dim query As String = "UPDATE Model SET WooProductId = @wooProductId WHERE PK_ModelId = @modelId"
        Using conn = GetConnection()
            conn.Open()
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@wooProductId", wooProductId)
                cmd.Parameters.AddWithValue("@modelId", modelId)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Public Function GetAllModelsWithDetails() As DataTable
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT 
                    m.PK_ModelId, 
                    m.ModelName, 
                    s.SeriesName, 
                    mf.ManufacturerName, 
                    m.TotalFabricSquareInches, 
                    ISNULL(m.Notes, '') AS Notes
                FROM Model m
                INNER JOIN ModelSeries s ON m.FK_SeriesId = s.PK_SeriesId
                INNER JOIN ModelManufacturers mf ON s.FK_ManufacturerId = mf.PK_ManufacturerId"
                Using da As New SqlClient.SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function
    Public Function GetActiveFabricCostPerSqInch(fabricTypeName As String) As Decimal?
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT TOP 1 fp.CostPerSquareInch
                FROM SupplierProductNameData spnd
                INNER JOIN JoinProductColorFabricType j ON spnd.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                INNER JOIN FabricTypeName ft ON j.FK_FabricTypeNameId = ft.PK_FabricTypeNameId
                OUTER APPLY (
                    SELECT TOP 1 fph.CostPerSquareInch
                    FROM FabricPricingHistory fph
                    WHERE fph.FK_SupplierProductNameDataId = spnd.PK_SupplierProductNameDataId
                    ORDER BY fph.DateFrom DESC
                ) fp
                WHERE ft.FabricType = @FabricTypeName
                  AND spnd.IsActiveForMarketplace = 1
                ORDER BY fp.CostPerSquareInch"
                cmd.Parameters.AddWithValue("@FabricTypeName", fabricTypeName)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    Return Convert.ToDecimal(result)
                End If
            End Using
        End Using
        Return Nothing
    End Function
End Class