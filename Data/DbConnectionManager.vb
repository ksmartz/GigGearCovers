Option Strict On
Option Explicit On
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports System.Linq


Public NotInheritable Class DbConnectionManager
    Private Sub New()
    End Sub

    ' Local fallback so the app still runs if App.config is missing/empty.
    Private Shared ReadOnly Fallback As String =
        "Data Source=MYPC\SQLEXPRESS;Initial Catalog=GigGearCoversDb;Integrated Security=True;Encrypt=True;TrustServerCertificate=True"

    Private Shared Function ResolveConnectionString() As String
        Dim cs = ConfigurationManager.ConnectionStrings("GigGearCoversDb")
        If cs IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(cs.ConnectionString) Then
            Return cs.ConnectionString
        End If
        Return Fallback
    End Function

    Public Shared ReadOnly Property ConnectionString As String
        Get
            Return ResolveConnectionString()
        End Get
    End Property

    Public Shared Function GetConnection() As SqlConnection
        Return New SqlConnection(ConnectionString)
    End Function

    Public Shared Sub EnsureOpen(conn As SqlConnection)
        If conn.State <> ConnectionState.Open Then conn.Open()
    End Sub
    ' === BRAND LIST ===
    Public Shared Function GetAllFabricBrandNames() As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT PK_FabricBrandNameId, BrandName
                FROM FabricBrandName
                ORDER BY BrandName;"
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' === PRODUCTS BY BRAND (brand name text) ===
    Public Shared Function GetProductsByBrandName(brandName As String) As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT p.PK_FabricBrandProductNameId, p.BrandProductName
                FROM FabricBrandProductName p
                INNER JOIN FabricBrandName b ON b.PK_FabricBrandNameId = p.FK_FabricBrandNameId
                WHERE LOWER(b.BrandName) = LOWER(@BrandName)
                ORDER BY p.BrandProductName;"
                cmd.Parameters.AddWithValue("@BrandName", brandName)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function


    ' === PRODUCT INFO (brand+product names) ===
    ' Returns WeightPerLinearYard, FabricRollWidth, FK_FabricTypeNameId, BrandName, BrandProductName
    Public Shared Function GetFabricProductInfo(brandName As String, productName As String) As DataRow
        Dim dt As New DataTable()
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT TOP 1 
                    p.PK_FabricBrandProductNameId,
                    p.BrandProductName,
                    b.BrandName,
                    p.WeightPerLinearYard,
                    p.FabricRollWidth,
                    p.FK_FabricTypeNameId
                FROM FabricBrandProductName p
                INNER JOIN FabricBrandName b ON b.PK_FabricBrandNameId = p.FK_FabricBrandNameId
                WHERE LOWER(b.BrandName) = LOWER(@BrandName)
                  AND LOWER(p.BrandProductName) = LOWER(@ProductName);"
                cmd.Parameters.AddWithValue("@BrandName", brandName)
                cmd.Parameters.AddWithValue("@ProductName", productName)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        If dt.Rows.Count > 0 Then Return dt.Rows(0)
        Return Nothing
    End Function

    ' === PRODUCT INFO (by product id) ===
    Public Shared Function GetFabricProductInfoById(productId As Integer) As DataRow
        Dim dt As New DataTable()
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT TOP 1
                    p.PK_FabricBrandProductNameId,
                    p.BrandProductName,
                    b.BrandName,
                    p.WeightPerLinearYard,
                    p.FabricRollWidth,
                    p.FK_FabricTypeNameId
                FROM FabricBrandProductName p
                INNER JOIN FabricBrandName b ON b.PK_FabricBrandNameId = p.FK_FabricBrandNameId
                WHERE p.PK_FabricBrandProductNameId = @ProductId;"
                cmd.Parameters.AddWithValue("@ProductId", productId)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        If dt.Rows.Count > 0 Then Return dt.Rows(0)
        Return Nothing
    End Function

    ' === LATEST PRICING (supplier + brand/product names) ===
    ' Returns the latest row from FabricPricingHistory for that supplier/product
    Public Shared Function GetFabricPricingHistory(supplierId As Integer, brandName As String, productName As String) As DataRow
        Dim dt As New DataTable()
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                ' Find latest pricing for any color/fabric-type of that product/supplier
                cmd.CommandText = "
                WITH P AS (
                    SELECT p.PK_FabricBrandProductNameId
                    FROM FabricBrandProductName p
                    INNER JOIN FabricBrandName b ON b.PK_FabricBrandNameId = p.FK_FabricBrandNameId
                    WHERE LOWER(b.BrandName) = LOWER(@BrandName)
                      AND LOWER(p.BrandProductName) = LOWER(@ProductName)
                )
                SELECT TOP 1 fph.*
                FROM SupplierProductNameData s
                INNER JOIN JoinProductColorFabricType j ON j.PK_JoinProductColorFabricTypeId = s.FK_JoinProductColorFabricTypeId
                INNER JOIN P ON P.PK_FabricBrandProductNameId = j.FK_FabricBrandProductNameId
                INNER JOIN FabricPricingHistory fph ON fph.FK_SupplierProductNameDataId = s.PK_SupplierProductNameDataId
                WHERE s.FK_SupplierNameId = @SupplierId
                ORDER BY fph.DateFrom DESC;"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                cmd.Parameters.AddWithValue("@BrandName", brandName)
                cmd.Parameters.AddWithValue("@ProductName", productName)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        If dt.Rows.Count > 0 Then Return dt.Rows(0)
        Return Nothing
    End Function

    ' === UPDATE TOTAL YARDS FOR A SUPPLIER-PRODUCT ROW ===
    Public Shared Sub UpdateSupplierProductTotalYards(supplierProductNameDataId As Integer, totalYards As Decimal)
        Using conn = GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                UPDATE SupplierProductNameData
                SET TotalYards = @TotalYards
                WHERE PK_SupplierProductNameDataId = @Id;"
                cmd.Parameters.AddWithValue("@TotalYards", totalYards)
                cmd.Parameters.AddWithValue("@Id", supplierProductNameDataId)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ' === GET SUPPLIERS FOR A SPECIFIC BRAND/PRODUCT/COLOR/FABRIC-TYPE COMBINATION ===
    ' Returns PK_SupplierNameId, CompanyName, IsActiveForMarketplace, CostPerSquareInch (latest)
    Public Shared Function GetSuppliersForCombination(brandId As Integer, productId As Integer, colorId As Integer, fabricTypeId As Integer) As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT DISTINCT
                    s.FK_SupplierNameId AS PK_SupplierNameId,
                    sn.CompanyName,
                    s.IsActiveForMarketplace,
                    ISNULL(fp.CostPerSquareInch, 0) AS CostPerSquareInch
                FROM SupplierProductNameData s
                INNER JOIN SupplierName sn ON sn.PK_SupplierNameId = s.FK_SupplierNameId
                INNER JOIN JoinProductColorFabricType j ON j.PK_JoinProductColorFabricTypeId = s.FK_JoinProductColorFabricTypeId
                OUTER APPLY (
                    SELECT TOP 1 fph.CostPerSquareInch
                    FROM FabricPricingHistory fph
                    WHERE fph.FK_SupplierProductNameDataId = s.PK_SupplierProductNameDataId
                    ORDER BY fph.DateFrom DESC
                ) fp
                WHERE j.FK_FabricBrandProductNameId = @ProductId
                  AND j.FK_ColorNameID = @ColorId
                  AND j.FK_FabricTypeNameId = @FabricTypeId
                ORDER BY sn.CompanyName;"
                cmd.Parameters.AddWithValue("@ProductId", productId)
                cmd.Parameters.AddWithValue("@ColorId", colorId)
                cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' === ENFORCE SINGLE ACTIVE SUPPLIER FOR A COMBINATION ===
    Public Shared Sub SetActiveForMarketplaceForCombination(brandId As Integer, productId As Integer, colorId As Integer, fabricTypeId As Integer, activeSupplierId As Integer)
        Using conn = GetConnection()
            If conn.State <> ConnectionState.Open Then conn.Open()
            Using tx = conn.BeginTransaction()
                Try
                    ' 1) Deactivate all suppliers for THAT combination
                    Using cmd = conn.CreateCommand()
                        cmd.Transaction = tx
                        cmd.CommandText = "
                        UPDATE s SET s.IsActiveForMarketplace = 0
                        FROM SupplierProductNameData s
                        INNER JOIN JoinProductColorFabricType j 
                            ON j.PK_JoinProductColorFabricTypeId = s.FK_JoinProductColorFabricTypeId
                        WHERE j.FK_FabricBrandProductNameId = @ProductId
                          AND j.FK_ColorNameID = @ColorId
                          AND j.FK_FabricTypeNameId = @FabricTypeId;"
                        cmd.Parameters.AddWithValue("@ProductId", productId)
                        cmd.Parameters.AddWithValue("@ColorId", colorId)
                        cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                        cmd.ExecuteNonQuery()
                    End Using

                    ' 2) Activate the chosen supplier for THAT combination
                    Using cmd = conn.CreateCommand()
                        cmd.Transaction = tx
                        cmd.CommandText = "
                        UPDATE s SET s.IsActiveForMarketplace = 1
                        FROM SupplierProductNameData s
                        INNER JOIN JoinProductColorFabricType j 
                            ON j.PK_JoinProductColorFabricTypeId = s.FK_JoinProductColorFabricTypeId
                        WHERE j.FK_FabricBrandProductNameId = @ProductId
                          AND j.FK_ColorNameID = @ColorId
                          AND j.FK_FabricTypeNameId = @FabricTypeId
                          AND s.FK_SupplierNameId = @SupplierId;"
                        cmd.Parameters.AddWithValue("@ProductId", productId)
                        cmd.Parameters.AddWithValue("@ColorId", colorId)
                        cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                        cmd.Parameters.AddWithValue("@SupplierId", activeSupplierId)
                        cmd.ExecuteNonQuery()
                    End Using

                    tx.Commit()
                Catch
                    tx.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    ' === PRODUCTS FOR SUPPLIER & BRAND (ids) ===
    Public Shared Function GetProductsForSupplierAndBrand(supplierId As Integer, brandId As Integer) As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT DISTINCT p.PK_FabricBrandProductNameId, p.BrandProductName
                FROM SupplierProductNameData s
                INNER JOIN JoinProductColorFabricType j ON j.PK_JoinProductColorFabricTypeId = s.FK_JoinProductColorFabricTypeId
                INNER JOIN FabricBrandProductName p ON p.PK_FabricBrandProductNameId = j.FK_FabricBrandProductNameId
                WHERE s.FK_SupplierNameId = @SupplierId
                  AND p.FK_FabricBrandNameId = @BrandId
                ORDER BY p.BrandProductName;"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                cmd.Parameters.AddWithValue("@BrandId", brandId)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function
    ' --------------------- COMMON LOOKUPS USED BY FORMS ---------------------

    ' Manufacturers for frmAddModelInformation
    Public Shared Function GetManufacturers() As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd As New SqlCommand("
                SELECT PK_ManufacturerId, ManufacturerName
                FROM ModelManufacturers
                ORDER BY ManufacturerName;", conn)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' Series for a manufacturer (frmAddModelInformation)
    Public Shared Function GetSeriesByManufacturer(manufacturerId As Integer) As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd As New SqlCommand("
                SELECT PK_SeriesId, SeriesName
                FROM ModelSeries
                WHERE FK_ManufacturerId = @mid
                ORDER BY SeriesName;", conn)
                cmd.Parameters.Add("@mid", SqlDbType.Int).Value = manufacturerId
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' Equipment types (frmAddModelInformation)
    Public Shared Function GetEquipmentTypes() As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd As New SqlCommand("
                SELECT PK_EquipmentTypeId, EquipmentTypeName
                FROM ModelEquipmentTypes
                ORDER BY EquipmentTypeName;", conn)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' Equipment type for a series (frmAddModelInformation)
    Public Shared Function GetEquipmentTypeForSeries(seriesId As Integer) As (Id As Integer?, Name As String)
        Const sql As String = "
            SELECT s.FK_EquipmentTypeId, et.EquipmentTypeName
            FROM ModelSeries s
            LEFT JOIN ModelEquipmentTypes et ON et.PK_EquipmentTypeId = s.FK_EquipmentTypeId
            WHERE s.PK_SeriesId = @sid;"
        Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
            cmd.Parameters.Add("@sid", SqlDbType.Int).Value = seriesId
            EnsureOpen(cn)
            Using r = cmd.ExecuteReader(CommandBehavior.SingleRow)
                If Not r.Read() Then Return (Nothing, Nothing)
                Dim id As Integer? = If(r.IsDBNull(0), CType(Nothing, Integer?), r.GetInt32(0))
                Dim name As String = If(r.IsDBNull(1), Nothing, r.GetString(1))
                Return (id, name)
            End Using
        End Using
    End Function

    ' --------------------- FABRIC ENTRY FORM LOOKUPS & ACTIONS ---------------------

    ' 1) All suppliers
    Public Shared Function GetAllSuppliers() As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = New SqlCommand("
                SELECT PK_SupplierNameId AS SupplierID, CompanyName AS SupplierName
                FROM SupplierInformation
                ORDER BY CompanyName;", conn)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' 2) All colors
    Public Shared Function GetAllMaterialColors() As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = New SqlCommand("
                SELECT PK_ColorNameID, ColorName, ColorNameFriendly
                FROM FabricColor
                ORDER BY ColorNameFriendly;", conn)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' 3) All fabric types
    Public Shared Function GetAllFabricTypes() As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = New SqlCommand("
                SELECT PK_FabricTypeNameId, FabricType
                FROM FabricTypeName
                ORDER BY FabricType;", conn)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' 4) Brands available for a supplier
    Public Shared Function GetBrandsForSupplier(supplierId As Integer) As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = New SqlCommand("
                SELECT DISTINCT b.PK_FabricBrandNameId, b.BrandName
                FROM SupplierProductNameData s
                INNER JOIN JoinProductColorFabricType j ON s.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                INNER JOIN FabricBrandProductName p ON j.FK_FabricBrandProductNameId = p.PK_FabricBrandProductNameId
                INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                WHERE s.FK_SupplierNameId = @SupplierId
                ORDER BY b.BrandName;", conn)
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' 5) Products for a supplier (all brands)
    Public Shared Function GetProductsForSupplier(supplierId As Integer) As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = New SqlCommand("
                SELECT DISTINCT 
                    p.PK_FabricBrandProductNameId, 
                    p.BrandProductName
                FROM SupplierProductNameData s
                INNER JOIN JoinProductColorFabricType j ON s.FK_JoinProductColorFabricTypeId = j.PK_JoinProductColorFabricTypeId
                INNER JOIN FabricBrandProductName p ON j.FK_FabricBrandProductNameId = p.PK_FabricBrandProductNameId
                WHERE s.FK_SupplierNameId = @SupplierId
                ORDER BY p.BrandProductName;", conn)
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' 6) Products by brand (brandId) for ComboBoxes
    Public Shared Function GetProductsByBrandId(brandId As Integer) As DataTable
        Dim dt As New DataTable()
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = New SqlCommand("
                SELECT PK_FabricBrandProductNameId, BrandProductName
                FROM FabricBrandProductName
                WHERE FK_FabricBrandNameId = @BrandId
                ORDER BY BrandProductName;", conn)
                cmd.Parameters.AddWithValue("@BrandId", brandId)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' 7) SupplierProduct row for (supplier, product, color, fabricType)
    Public Shared Function GetSupplierProductNameData(supplierId As Integer, productId As Integer, colorId As Integer, fabricTypeId As Integer) As DataRow
        Dim dt As New DataTable()
        Using conn = GetConnection()
            EnsureOpen(conn)
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
                    da.Fill(dt)
                End Using
            End Using
        End Using
        If dt.Rows.Count > 0 Then Return dt.Rows(0) Else Return Nothing
    End Function

    ' 8) Insert pricing history for a supplier-product
    Public Shared Sub InsertFabricPricingHistory(
        supplierProductNameDataId As Integer,
        shippingCost As Decimal,
        costPerLinearYard As Decimal,
        costPerSquareInch As Decimal,
        weightPerSquareInch As Decimal
    )
        Using conn = GetConnection()
            EnsureOpen(conn)
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

    ' 9) Update product weight/roll width
    Public Shared Sub UpdateFabricProductInfo(
        productId As Integer,
        weightPerLinearYard As Decimal,
        fabricRollWidth As Decimal
    )
        Using conn = GetConnection()
            EnsureOpen(conn)
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






    ' === Marketplace fee percentage (0.10 = 10%) ===
    Public Shared Function GetMarketplaceFeePercentage() As Decimal
        Dim pct As Decimal
        Dim raw = ConfigurationManager.AppSettings("MarketplaceFeePercent")
        If raw IsNot Nothing AndAlso Decimal.TryParse(raw, pct) AndAlso pct >= 0D AndAlso pct <= 1D Then
            Return pct
        End If
        ' Fallback default:
        Return 0.1D
    End Function

    ' === Record pricing snapshot for a model you just published to Woo ===
    ' Adjust table/column names to your schema if needed.
    Public Shared Sub InsertModelHistoryCostRetailPricing(
    modelId As Integer,
    wooParentProductId As Integer,
    wooParentSku As String,
    totalWeightOunces As Decimal,
    shippingCost As Decimal,
    materialCost As Decimal,
    marketplaceFeePercent As Decimal,
    retailPrice As Decimal,
    Optional note As String = Nothing
)
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                INSERT INTO ModelHistoryCostRetailPricing
                (FK_ModelId, WooParentProductId, WooParentSku, TotalWeightOunces, ShippingCost, MaterialCost,
                 MarketplaceFeePercent, RetailPrice, CreatedOn, Note)
                VALUES
                (@ModelId, @WooParentProductId, @WooParentSku, @TotalWeightOunces, @ShippingCost, @MaterialCost,
                 @MarketplaceFeePercent, @RetailPrice, @CreatedOn, @Note);"

                cmd.Parameters.AddWithValue("@ModelId", modelId)
                cmd.Parameters.AddWithValue("@WooParentProductId", wooParentProductId)
                cmd.Parameters.AddWithValue("@WooParentSku", If(wooParentSku Is Nothing, String.Empty, wooParentSku))
                cmd.Parameters.AddWithValue("@TotalWeightOunces", totalWeightOunces)
                cmd.Parameters.AddWithValue("@ShippingCost", shippingCost)
                cmd.Parameters.AddWithValue("@MaterialCost", materialCost)
                cmd.Parameters.AddWithValue("@MarketplaceFeePercent", marketplaceFeePercent)
                cmd.Parameters.AddWithValue("@RetailPrice", retailPrice)
                cmd.Parameters.AddWithValue("@CreatedOn", Date.UtcNow)

                ' >>> FIX: declare the parameter type and set Value explicitly
                Dim pNote = cmd.Parameters.Add("@Note", SqlDbType.NVarChar, -1) ' -1 = NVARCHAR(MAX)
                If note Is Nothing OrElse note.Length = 0 Then
                    pNote.Value = DBNull.Value
                Else
                    pNote.Value = note
                End If

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ' --- LOOK UP A PRODUCT BY NAME (used for weights/widths) ---
    Public Shared Function GetActiveFabricBrandProductName(productName As String) As DataRow
        Dim dt As New DataTable()
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                ' If you actually need "active for marketplace", join SupplierProductNameData and filter there.
                cmd.CommandText = "
                SELECT TOP 1 
                    p.PK_FabricBrandProductNameId,
                    p.BrandProductName,
                    p.WeightPerLinearYard,
                    p.FabricRollWidth
                FROM FabricBrandProductName p
                WHERE LOWER(p.BrandProductName) = LOWER(@pname)
                ORDER BY p.PK_FabricBrandProductNameId;"
                cmd.Parameters.AddWithValue("@pname", productName)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        If dt.Rows.Count > 0 Then Return dt.Rows(0)
        Return Nothing
    End Function

    ' --- SHIPPING BY WEIGHT (ounces) ---
    Public Shared Function GetShippingCostByWeight(totalOunces As Decimal) As Decimal
        ' If you have a table with tiers, replace the query below accordingly.
        ' This implementation tries a tiered lookup and falls back to 0 on any error.
        Try
            Using conn = GetConnection()
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "
                    SELECT TOP 1 Cost
                    FROM ShippingRateByOunces
                    WHERE @oz BETWEEN MinOunces AND MaxOunces
                    ORDER BY MaxOunces;"
                    cmd.Parameters.AddWithValue("@oz", totalOunces)
                    Dim obj = cmd.ExecuteScalar()
                    If obj IsNot Nothing AndAlso obj IsNot DBNull.Value Then
                        Return Convert.ToDecimal(obj)
                    End If
                End Using
            End Using
        Catch
            ' swallow if table doesn't exist yet; treat as 0
        End Try
        Return 0D
    End Function

    ' --- MARKETPLACE FEE PERCENT (e.g., "eBay", "Etsy") ---
    Public Shared Function GetMarketplaceFeePercentage(marketplace As String) As Decimal
        Try
            Using conn = GetConnection()
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "
                    SELECT TOP 1 FeePercent
                    FROM MarketplaceFees
                    WHERE LOWER(MarketplaceName) = LOWER(@m);"
                    cmd.Parameters.AddWithValue("@m", marketplace)
                    Dim obj = cmd.ExecuteScalar()
                    If obj IsNot Nothing AndAlso obj IsNot DBNull.Value Then
                        Return Convert.ToDecimal(obj)
                    End If
                End Using
            End Using
        Catch
        End Try
        Return 0D
    End Function
    Public Shared Sub InsertModelHistoryCostRetailPricing(
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
    weight_PaddingOnly As Decimal?,
    weight_ChoiceWaterproof As Decimal?,
    weight_ChoiceWaterproof_Padded As Decimal?,
    weight_PremiumSyntheticLeather As Decimal?,
    weight_PremiumSyntheticLeather_Padded As Decimal?,
    shipping_Choice As Decimal,
    shipping_ChoicePadded As Decimal,
    shipping_Leather As Decimal,
    shipping_LeatherPadded As Decimal,
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
        ' Summaries for the short columns (so we can reuse the existing table)
        Dim totalWeightOunces As Decimal =
        weight_ChoiceWaterproof.GetValueOrDefault() +
        weight_ChoiceWaterproof_Padded.GetValueOrDefault() +
        weight_PremiumSyntheticLeather.GetValueOrDefault() +
        weight_PremiumSyntheticLeather_Padded.GetValueOrDefault() +
        weight_PaddingOnly.GetValueOrDefault()

        Dim shippingCost As Decimal = shipping_Choice + shipping_ChoicePadded + shipping_Leather + shipping_LeatherPadded

        Dim materialCost As Decimal =
        baseCost_ChoiceWaterproof.GetValueOrDefault() +
        baseCost_ChoiceWaterproof_Padded.GetValueOrDefault() +
        baseCost_PremiumSyntheticLeather.GetValueOrDefault() +
        baseCost_PremiumSyntheticLeather_Padded.GetValueOrDefault() +
        baseCost_PaddingOnly.GetValueOrDefault()

        ' We’ll keep these as placeholders; you can compute a blended fee/retail if you prefer
        Dim marketplaceFeePercent As Decimal = 0D
        Dim retailPrice As Decimal = 0D

        ' Serialize detail as JSON into Note (so nothing is lost)
        Dim sb As New StringBuilder(2048)
        sb.Append("{""detail"":{")
        sb.Append("""costPerSqIn"":{")
        sb.Append("""choice"":").Append(If(costPerSqInch_ChoiceWaterproof.HasValue, costPerSqInch_ChoiceWaterproof.Value.ToString(System.Globalization.CultureInfo.InvariantCulture), "null")).Append(","c)
        sb.Append("""leather"":").Append(If(costPerSqInch_PremiumSyntheticLeather.HasValue, costPerSqInch_PremiumSyntheticLeather.Value.ToString(System.Globalization.CultureInfo.InvariantCulture), "null")).Append(","c)
        sb.Append("""padding"":").Append(If(costPerSqInch_Padding.HasValue, costPerSqInch_Padding.Value.ToString(System.Globalization.CultureInfo.InvariantCulture), "null"))
        sb.Append("},""wastePercent"":").Append(wastePercent.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append(","c)
        sb.Append("""notes"":").Append("""").Append((If(notes, String.Empty)).Replace("""", "\""")).Append("""")
        sb.Append("}}")
        Dim noteJson As String = sb.ToString()

        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                INSERT INTO ModelHistoryCostRetailPricing
                (FK_ModelId, WooParentProductId, WooParentSku, TotalWeightOunces, ShippingCost, MaterialCost,
                 MarketplaceFeePercent, RetailPrice, CreatedOn, Note)
                VALUES
                (@ModelId, @WooParentProductId, @WooParentSku, @TotalWeightOunces, @ShippingCost, @MaterialCost,
                 @MarketplaceFeePercent, @RetailPrice, @CreatedOn, @Note);"
                ' Re-use existing columns; we don’t have Woo product info here
                cmd.Parameters.AddWithValue("@ModelId", modelId)
                cmd.Parameters.AddWithValue("@WooParentProductId", 0)
                cmd.Parameters.AddWithValue("@WooParentSku", "")
                cmd.Parameters.AddWithValue("@TotalWeightOunces", totalWeightOunces)
                cmd.Parameters.AddWithValue("@ShippingCost", shippingCost)
                cmd.Parameters.AddWithValue("@MaterialCost", materialCost)
                cmd.Parameters.AddWithValue("@MarketplaceFeePercent", marketplaceFeePercent)
                cmd.Parameters.AddWithValue("@RetailPrice", retailPrice)
                cmd.Parameters.AddWithValue("@CreatedOn", Date.UtcNow)
                cmd.Parameters.AddWithValue("@Note", If(noteJson Is Nothing, CType(DBNull.Value, Object), CType(noteJson, Object)))
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    ' -----------------------------------------------
    ' MODEL LOOKUP
    ' -----------------------------------------------
    Public Shared Function GetModelRow(modelId As Integer) As DataRow
        Dim dt As New DataTable()
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = conn.CreateCommand()
                ' Adjust table/column names if your schema differs
                cmd.CommandText = "
                SELECT TOP 1 *
                FROM Model
                WHERE PK_ModelId = @Id;"
                cmd.Parameters.Add("@Id", SqlDbType.Int).Value = modelId
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        If dt.Rows.Count > 0 Then Return dt.Rows(0)
        Return Nothing
    End Function

    ' -----------------------------------------------
    ' STATIC LISTS USED BY WOO PUBLISHER
    ' -----------------------------------------------
    Public Shared Function GetAllFabricTypeNames() As List(Of String)
        Dim list As New List(Of String)()
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd As New SqlCommand("
            SELECT FabricType
            FROM FabricTypeName
            ORDER BY FabricType;", conn)
                Using r = cmd.ExecuteReader()
                    While r.Read()
                        If Not r.IsDBNull(0) Then list.Add(r.GetString(0))
                    End While
                End Using
            End Using
        End Using
        Return list
    End Function

    Public Shared Function GetAllColorNames() As List(Of String)
        Dim list As New List(Of String)()
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd As New SqlCommand("
            SELECT ColorNameFriendly
            FROM FabricColor
            ORDER BY ColorNameFriendly;", conn)
                Using r = cmd.ExecuteReader()
                    While r.Read()
                        If Not r.IsDBNull(0) Then list.Add(r.GetString(0))
                    End While
                End Using
            End Using
        End Using
        Return list
    End Function

    ' -----------------------------------------------
    ' IMAGES (PARENT GALLERY + VARIATION IMAGE REF)
    ' -----------------------------------------------
    ' Returns list of (WooMediaId, Url, Position) for a parent product gallery
    Public Shared Function GetParentGalleryImages(equipmentTypeId As Integer) _
    As List(Of Tuple(Of Integer?, String, Integer))

        Dim result As New List(Of Tuple(Of Integer?, String, Integer))()
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT mei.WooMediaId, mei.Url, mei.Position
                FROM ModelEquipmentTypeImage AS mei
                WHERE mei.FK_EquipmentTypeId = @et
                ORDER BY COALESCE(mei.Position, 0), mei.PK_ModelEquipmentTypeImageId;"
                cmd.Parameters.Add("@et", SqlDbType.Int).Value = equipmentTypeId
                Using r = cmd.ExecuteReader()
                    While r.Read()
                        Dim mediaId As Integer? = If(r.IsDBNull(0), CType(Nothing, Integer?), r.GetInt32(0))
                        Dim url As String = If(r.IsDBNull(1), Nothing, r.GetString(1))
                        Dim pos As Integer = If(r.IsDBNull(2), 0, r.GetInt32(2))
                        result.Add(Tuple.Create(mediaId, url, pos))
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    ' Returns a single image ref for a variation (WooMediaId, Url)
    Public Shared Function GetVariationImageRef(equipmentTypeId As Integer,
                                            fabricTypeName As String,
                                            colorNameFriendly As String) _
                                            As Tuple(Of Integer?, String)
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT TOP 1 vm.WooMediaId, vm.Url
                FROM VariationMedia vm
                WHERE vm.FK_EquipmentTypeId = @et
                  AND LOWER(vm.FabricTypeName) = LOWER(@ft)
                  AND LOWER(vm.ColorNameFriendly) = LOWER(@color)
                ORDER BY vm.Priority ASC, vm.PK_VariationMediaId ASC;"
                cmd.Parameters.Add("@et", SqlDbType.Int).Value = equipmentTypeId
                cmd.Parameters.Add("@ft", SqlDbType.NVarChar, 100).Value = fabricTypeName
                cmd.Parameters.Add("@color", SqlDbType.NVarChar, 100).Value = colorNameFriendly
                Using r = cmd.ExecuteReader(CommandBehavior.SingleRow)
                    If Not r.Read() Then Return Tuple.Create(CType(Nothing, Integer?), CType(Nothing, String))
                    Dim mediaId As Integer? = If(r.IsDBNull(0), CType(Nothing, Integer?), r.GetInt32(0))
                    Dim url As String = If(r.IsDBNull(1), Nothing, r.GetString(1))
                    Return Tuple.Create(mediaId, url)
                End Using
            End Using
        End Using
    End Function

    ' -----------------------------------------------
    ' LOGGING (SYNC LOG)
    ' -----------------------------------------------
    ' Writes a normalized log row for any Woo operation.
    Public Shared Sub InsertMpWooSyncLog(operation As String,
                                     refSku As String,
                                     refWooId As Integer?,
                                     success As Boolean,
                                     httpStatus As String,
                                     requestJson As String,
                                     responseJson As String,
                                     notes As String)
        ' Normalize fields for DB
        Dim op As String = If(operation, String.Empty).Trim()
        Dim action As String =
        If(op.IndexOf("create", StringComparison.OrdinalIgnoreCase) >= 0, "Create",
        If(op.IndexOf("update", StringComparison.OrdinalIgnoreCase) >= 0, "Update",
        If(op.IndexOf("delete", StringComparison.OrdinalIgnoreCase) >= 0, "Delete", "Other")))

        Dim entityType As String =
        If(op.IndexOf("variation", StringComparison.OrdinalIgnoreCase) >= 0, "Variation",
        If(op.IndexOf("parent", StringComparison.OrdinalIgnoreCase) >= 0, "Product", "Unknown"))

        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
INSERT INTO dbo.MpWooSyncLog
    ([EntityType],[Action],[Operation],[RefSKU],[RefWooId],
     [Status],[Success],[HttpStatus],
     [RequestJson],[ResponseJson],[Notes],
     [DateLogged],[CreatedAtUtc])
VALUES
    (@EntityType,@Action,@Operation,@RefSKU,@RefWooId,
     @Status,@Success,@HttpStatus,
     @RequestJson,@ResponseJson,@Notes,
     SYSDATETIME(), SYSUTCDATETIME());"

                cmd.Parameters.Add("@EntityType", SqlDbType.NVarChar, 50).Value = entityType
                cmd.Parameters.Add("@Action", SqlDbType.NVarChar, 50).Value = action
                cmd.Parameters.Add("@Operation", SqlDbType.NVarChar, 100).Value = op
                cmd.Parameters.Add("@RefSKU", SqlDbType.NVarChar, 100).Value = If(refSku, String.Empty)
                If refWooId.HasValue Then
                    cmd.Parameters.Add("@RefWooId", SqlDbType.Int).Value = refWooId.Value
                Else
                    cmd.Parameters.Add("@RefWooId", SqlDbType.Int).Value = DBNull.Value
                End If
                cmd.Parameters.Add("@Status", SqlDbType.NVarChar, 50).Value = If(success, "OK", "ERROR")
                cmd.Parameters.Add("@Success", SqlDbType.Bit).Value = If(success, 1, 0)

                If String.IsNullOrWhiteSpace(httpStatus) Then
                    cmd.Parameters.Add("@HttpStatus", SqlDbType.NVarChar, 20).Value = DBNull.Value
                Else
                    cmd.Parameters.Add("@HttpStatus", SqlDbType.NVarChar, 20).Value = httpStatus
                End If

                cmd.Parameters.Add("@RequestJson", SqlDbType.NVarChar, -1).Value =
                If(requestJson, String.Empty)
                cmd.Parameters.Add("@ResponseJson", SqlDbType.NVarChar, -1).Value =
                If(responseJson, String.Empty)
                cmd.Parameters.Add("@Notes", SqlDbType.NVarChar, -1).Value =
                If(notes, String.Empty)

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ' -----------------------------------------------
    ' UPSERTS FOR PRODUCT / VARIATION ROWS
    ' -----------------------------------------------
    ' Returns PK from MpWooProduct (upsert on (ModelId, ParentSku))
    Public Shared Function UpsertMpWooProduct(modelId As Integer,
                                          parentSku As String,
                                          wooProductId As Integer?,
                                          wooCategoryId As Integer?,
                                          status As String,
                                          message As String) As Integer
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
MERGE dbo.MpWooProduct AS T
USING (SELECT @ModelId AS ModelId, @ParentSku AS ParentSku) AS S
ON (T.FK_ModelId = S.ModelId AND T.ParentSku = S.ParentSku)
WHEN MATCHED THEN
    UPDATE SET
        T.WooProductId   = COALESCE(@WooProductId, T.WooProductId),
        T.WooCategoryId  = COALESCE(@WooCategoryId, T.WooCategoryId),
        T.SyncStatus     = @Status,
        T.LastMessage    = @Message,
        T.UpdatedAtUtc   = SYSUTCDATETIME()
WHEN NOT MATCHED THEN
    INSERT (FK_ModelId, ParentSku, WooProductId, WooCategoryId, SyncStatus, LastMessage, CreatedAtUtc, UpdatedAtUtc)
    VALUES (@ModelId, @ParentSku, @WooProductId, @WooCategoryId, @Status, @Message, SYSUTCDATETIME(), SYSUTCDATETIME())
OUTPUT inserted.PK_MpWooProductId;"  ' <-- matches the (renamed) PK
                cmd.Parameters.Add("@ModelId", SqlDbType.Int).Value = modelId
                cmd.Parameters.Add("@ParentSku", SqlDbType.NVarChar, 64).Value = If(parentSku, String.Empty)

                If wooProductId.HasValue Then
                    cmd.Parameters.Add("@WooProductId", SqlDbType.Int).Value = wooProductId.Value
                Else
                    cmd.Parameters.Add("@WooProductId", SqlDbType.Int).Value = DBNull.Value
                End If

                If wooCategoryId.HasValue Then
                    cmd.Parameters.Add("@WooCategoryId", SqlDbType.Int).Value = wooCategoryId.Value
                Else
                    cmd.Parameters.Add("@WooCategoryId", SqlDbType.Int).Value = DBNull.Value
                End If

                cmd.Parameters.Add("@Status", SqlDbType.NVarChar, 32).Value = If(status, String.Empty)
                cmd.Parameters.Add("@Message", SqlDbType.NVarChar, -1).Value = If(message, String.Empty)

                Dim id As Object = cmd.ExecuteScalar()
                Return Convert.ToInt32(id)
            End Using
        End Using
    End Function

    ' DbConnectionManager.vb

    ' --- Fabric: name -> 1-char abbreviation (dbo.fabrictypename) ---
    Public Shared Function GetFabricAbbreviationMap() As Dictionary(Of String, String)
        Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT FabricTypeName, FabricTypeNameAbbreviation " &
                              "FROM dbo.fabrictypename " &
                              "WHERE FabricTypeName IS NOT NULL " &
                              "  AND FabricTypeNameAbbreviation IS NOT NULL " &
                              "  AND LTRIM(RTRIM(FabricTypeNameAbbreviation)) <> '';"
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        Dim name As String = If(rd.IsDBNull(0), "", rd.GetString(0)).Trim()
                        Dim code As String = If(rd.IsDBNull(1), "", rd.GetString(1)).Trim().ToUpperInvariant()
                        If code.Length > 1 Then code = code.Substring(0, 1)
                        If code.Length < 1 Then code = "X"          ' or: code = code.PadRight(1, "X"c)
                        If name <> "" Then dict(name) = code
                    End While
                End Using
            End Using
        End Using
        Return dict
    End Function

    ' --- Color: name -> 3-char abbreviation (dbo.fabriccolor) ---
    Public Shared Function GetColorAbbreviationMap() As Dictionary(Of String, String)
        Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT ColorName, ColorNameAbbreviation " &
                              "FROM dbo.fabriccolor " &
                              "WHERE ColorName IS NOT NULL " &
                              "  AND ColorNameAbbreviation IS NOT NULL " &
                              "  AND LTRIM(RTRIM(ColorNameAbbreviation)) <> '';"
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        Dim name As String = If(rd.IsDBNull(0), "", rd.GetString(0)).Trim()
                        Dim code As String = If(rd.IsDBNull(1), "", rd.GetString(1)).Trim().ToUpperInvariant()
                        If code.Length > 3 Then code = code.Substring(0, 3)
                        If code.Length < 3 Then code = code.PadRight(3, "X"c)    ' ✅ correct char literal
                        If name <> "" Then dict(name) = code
                    End While
                End Using
            End Using
        End Using
        Return dict
    End Function



    ' --- Color: name -> 3-char abbreviation (dbo.ColorName; falls back to dbo.Color) ---



    ' Optional: track the Woo media used for a variation image
    Public Shared Sub UpsertMpWooVariationMedia(mpWooProductId As Integer,
                                            variationSku As String,
                                            wooMediaId As Integer?,
                                            imageUrl As String)
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
MERGE dbo.MpWooVariationMedia AS T
USING (SELECT @MpWooProductId AS PId, @Sku AS Sku) AS S
ON (T.FK_MpWooProductId = S.PId AND T.VariationSku = S.Sku)
WHEN MATCHED THEN
    UPDATE SET T.WooMediaId = @WooMediaId, T.ImageUrl = @ImageUrl, T.UpdatedAtUtc = SYSUTCDATETIME()
WHEN NOT MATCHED THEN
    INSERT (FK_MpWooProductId, VariationSku, WooMediaId, ImageUrl, CreatedAtUtc, UpdatedAtUtc)
    VALUES (@MpWooProductId, @Sku, @WooMediaId, @ImageUrl, SYSUTCDATETIME(), SYSUTCDATETIME());"
                cmd.Parameters.Add("@MpWooProductId", SqlDbType.Int).Value = mpWooProductId
                cmd.Parameters.Add("@Sku", SqlDbType.NVarChar, 200).Value = variationSku
                If wooMediaId.HasValue Then
                    cmd.Parameters.Add("@WooMediaId", SqlDbType.Int).Value = wooMediaId.Value
                Else
                    cmd.Parameters.Add("@WooMediaId", SqlDbType.Int).Value = DBNull.Value
                End If
                cmd.Parameters.Add("@ImageUrl", SqlDbType.NVarChar, -1).Value = If(imageUrl, String.Empty)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ' === UPSERT A CHILD VARIATION MAPPING (Woo <-> DB) ===
    ' Table (suggested):
    '   MpWooVariation(
    '       PK_MpWooVariationId INT IDENTITY(1,1) PRIMARY KEY,
    '       FK_MpWooProductId   INT NOT NULL,
    '       VariationSku        NVARCHAR(200) NOT NULL,
    '       WooVariationId      INT NULL,
    '       FabricOption        NVARCHAR(200) NULL,
    '       ColorOption         NVARCHAR(200) NULL,
    '       LastSyncStatus      NVARCHAR(50)  NULL,
    '       LastSyncNote        NVARCHAR(4000) NULL,
    '       LastSyncedOn        DATETIME2 NOT NULL
    '   )
    Public Shared Function UpsertMpWooVariation(mpWooProductId As Integer,
                                            childSku As String,
                                            wooVariationId As Integer?,
                                            fabricOption As String,
                                            colorOption As String,
                                            status As String,
                                            note As String) As Integer
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = conn.CreateCommand()

                Dim child As String = If(childSku, "").Trim()
                If child.Length = 0 Then
                    Throw New ArgumentException("ChildSKU must not be empty.", NameOf(childSku))
                End If

                cmd.CommandText = "
MERGE dbo.MpWooVariation AS T
USING (SELECT @FK AS FK, @ChildSku AS ChildSKU) AS S
ON (T.FK_MpWooProductId = S.FK AND T.ChildSKU = S.ChildSKU)
WHEN MATCHED THEN
    UPDATE SET
        T.WooVariationId  = COALESCE(@WooVarId, T.WooVariationId),
        T.FabricOption    = @FabricOption,
        T.ColorOption     = @ColorOption,
        T.LastSyncStatus  = @Status,
        T.LastSyncNote    = @Note,
        T.LastSyncedOn    = SYSUTCDATETIME()
WHEN NOT MATCHED THEN
    INSERT (FK_MpWooProductId, ChildSKU, WooVariationId, FabricOption, ColorOption, LastSyncStatus, LastSyncNote, LastSyncedOn)
    VALUES (@FK, @ChildSku, @WooVarId, @FabricOption, @ColorOption, @Status, @Note, SYSUTCDATETIME())
OUTPUT inserted.PK_MpWooVariationId;
"
                cmd.Parameters.Add("@FK", SqlDbType.Int).Value = mpWooProductId
                cmd.Parameters.Add("@ChildSku", SqlDbType.NVarChar, 128).Value = child

                If wooVariationId.HasValue Then
                    cmd.Parameters.Add("@WooVarId", SqlDbType.Int).Value = wooVariationId.Value
                Else
                    cmd.Parameters.Add("@WooVarId", SqlDbType.Int).Value = DBNull.Value
                End If

                cmd.Parameters.Add("@FabricOption", SqlDbType.NVarChar, 200).Value = If(fabricOption, "")
                cmd.Parameters.Add("@ColorOption", SqlDbType.NVarChar, 200).Value = If(colorOption, "")
                cmd.Parameters.Add("@Status", SqlDbType.NVarChar, 64).Value = If(status, "")
                cmd.Parameters.Add("@Note", SqlDbType.NVarChar, -1).Value = If(note, "")

                Dim idObj = cmd.ExecuteScalar()
                Return Convert.ToInt32(idObj)
            End Using
        End Using
    End Function





    ' === MODELS FOR A GIVEN MANUFACTURER + SERIES ===
    ' Returns: PK_ModelId, ModelName (add more columns if you want to display more)
    Public Shared Function GetModelsForSeriesAndManufacturer(manufacturerId As Integer, seriesId As Integer) As DataTable
        Dim dt As New DataTable()
        Const sql As String = "
        SELECT m.PK_ModelId, m.ModelName
        FROM Models m
        INNER JOIN ModelSeries s ON s.PK_SeriesId = m.FK_SeriesId
        WHERE m.FK_SeriesId = @sid
          AND s.FK_ManufacturerId = @mid
        ORDER BY m.ModelName;"
        Using conn = GetConnection()
            EnsureOpen(conn)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = sql
                cmd.Parameters.Add("@sid", SqlDbType.Int).Value = seriesId
                cmd.Parameters.Add("@mid", SqlDbType.Int).Value = manufacturerId
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' === MODELS FOR A GIVEN MANUFACTURER + SERIES ===
    Public Shared Function GetModelsForSeries(manufacturerId As Integer, seriesId As Integer) As DataTable
        ' Adjust table/column names if your model table is named differently.
        ' This assumes:
        '   ModelInformation (PK_ModelId, ModelName, FK_SeriesId, ...)
        '   ModelSeries (PK_SeriesId, FK_ManufacturerId, ...)
        '   ModelManufacturers (PK_ManufacturerId, ...)
        Dim dt As New DataTable()
        Const sql As String = "
        SELECT m.PK_ModelId, m.ModelName
        FROM Model m
        INNER JOIN ModelSeries s ON s.PK_SeriesId = m.FK_SeriesId
        WHERE s.PK_SeriesId = @sid
          AND s.FK_ManufacturerId = @mid
        ORDER BY m.ModelName;"

        Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
            cmd.Parameters.Add("@sid", SqlDbType.Int).Value = seriesId
            cmd.Parameters.Add("@mid", SqlDbType.Int).Value = manufacturerId
            EnsureOpen(cn)
            Using da As New SqlDataAdapter(cmd)
                da.Fill(dt)
            End Using
        End Using
        Return dt
    End Function
    ' Returns a DataTable ready to bind to dgvModelInformation, with columns tailored by equipment type.
    Public Shared Function GetModelGridRows(manufacturerId As Integer,
                                        seriesId As Integer,
                                        equipmentTypeName As String) As DataTable
        Dim isKeyboard As Boolean =
        equipmentTypeName IsNot Nothing AndAlso
        equipmentTypeName.Trim().ToLowerInvariant().Contains("keyboard")

        ' NOTE: Adjust table/column names if your schema differs.
        ' This assumes:
        '   ModelInformation m (PK_ModelId, ModelName, ParentSku, Width, Depth, Height,
        '                       TotalFabricSquareInches, AmpHandleLocation,
        '                       TAHWidth, TAHHeight, TAHRearOffset,
        '                       SAHHeight, SAHWidth, SAHRearOffset, SAHTopDownOffset,
        '                       MusicRestDesign, Chart_Template, Notes, FK_SeriesId)
        '   ModelSeries s (PK_SeriesId, FK_ManufacturerId)
        '   MpWooProduct p (FK_ModelId, WooProductId)
        Dim sqlAmplifier As String =
"SELECT  m.PK_ModelId,
         m.ModelName,
         m.ParentSku,
         m.Width,
         m.Depth,
         m.Height,
         m.TotalFabricSquareInches,
         m.AmpHandleLocation,
         m.TAHWidth,
         m.TAHHeight,
         m.TAHRearOffset,
         m.SAHHeight,
         m.SAHWidth,
         m.SAHRearOffset,
         m.SAHTopDownOffset,
         m.Chart_Template,
         m.Notes,
         p.WooProductId
FROM     Model m
JOIN     ModelSeries s ON s.PK_SeriesId = m.FK_SeriesId
LEFT JOIN MpWooProduct p ON p.FK_ModelId = m.PK_ModelId
WHERE    s.PK_SeriesId = @sid
  AND    s.FK_ManufacturerId = @mid
ORDER BY m.ModelName;"

        Dim sqlKeyboard As String =
"SELECT  m.PK_ModelId,
         m.ModelName,
         m.ParentSku,
         m.Width,
         m.Depth,
         m.Height,
         m.TotalFabricSquareInches,
         m.MusicRestDesign,
         m.Chart_Template,
         m.Notes,
         p.WooProductId
FROM     Model m
JOIN     ModelSeries s ON s.PK_SeriesId = m.FK_SeriesId
LEFT JOIN MpWooProduct p ON p.FK_ModelId = m.PK_ModelId
WHERE    s.PK_SeriesId = @sid
  AND    s.FK_ManufacturerId = @mid
ORDER BY m.ModelName;"

        Dim dt As New DataTable()
        Using cn = GetConnection()
            EnsureOpen(cn)
            Using cmd As New SqlCommand(If(isKeyboard, sqlKeyboard, sqlAmplifier), cn)
                cmd.Parameters.Add("@sid", SqlDbType.Int).Value = seriesId
                cmd.Parameters.Add("@mid", SqlDbType.Int).Value = manufacturerId
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' DbConnectionManager.vb
    Public Shared Function GetAmpHandleLocations() As DataTable
        Dim dt As New DataTable()
        Using cn = GetConnection()
            EnsureOpen(cn)
            Using cmd As New SqlCommand("
            SELECT PK_AmpHandleLocationId, AmpHandleLocationName AS HandleLocationName
            FROM AmpHandleLocation
            ORDER BY AmpHandleLocationName;", cn)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function



    ' 3) Save edits from one row (handles both amp/cab and keyboard layouts)
    Public Shared Sub UpdateModelDisplayFields(
    modelId As Integer,
    parentSku As String,
    width As Decimal?, depth As Decimal?, height As Decimal?,
    totalFabricSquareInches As Decimal?,
    ampHandleLocation As String,
    tahWidth As Decimal?, tahHeight As Decimal?, tahRearOffset As Decimal?,
    sahHeight As Decimal?, sahWidth As Decimal?, sahRearOffset As Decimal?, sahTopDownOffset As Decimal?,
    musicRestDesign As String,
    chartTemplate As String,
    notes As String
)
        Using cn = GetConnection()
            EnsureOpen(cn)
            Using cmd = cn.CreateCommand()
                cmd.CommandText = "
                UPDATE Model
                SET ParentSku               = @ParentSku,
                    Width                   = @Width,
                    Depth                   = @Depth,
                    Height                  = @Height,
                    TotalFabricSquareInches = @TFSI,
                    AmpHandleLocation       = @AmpHandleLocation,
                    TAHWidth                = @TAHWidth,
                    TAHHeight               = @TAHHeight,
                    TAHRearOffset           = @TAHRearOffset,
                    SAHHeight               = @SAHHeight,
                    SAHWidth                = @SAHWidth,
                    SAHRearOffset           = @SAHRearOffset,
                    SAHTopDownOffset        = @SAHTopDownOffset,
                    MusicRestDesign         = @MusicRestDesign,
                    Chart_Template          = @ChartTemplate,
                    Notes                   = @Notes
                WHERE PK_ModelId = @Id;"
                cmd.Parameters.Add("@Id", SqlDbType.Int).Value = modelId
                cmd.Parameters.Add("@ParentSku", SqlDbType.NVarChar, 200).Value = If(parentSku, String.Empty)

                Dim p = cmd.Parameters
                p.Add("@Width", SqlDbType.Decimal).Value = If(width.HasValue, CType(width, Object), DBNull.Value)
                p.Add("@Depth", SqlDbType.Decimal).Value = If(depth.HasValue, CType(depth, Object), DBNull.Value)
                p.Add("@Height", SqlDbType.Decimal).Value = If(height.HasValue, CType(height, Object), DBNull.Value)
                p.Add("@TFSI", SqlDbType.Decimal).Value = If(totalFabricSquareInches.HasValue, CType(totalFabricSquareInches, Object), DBNull.Value)

                p.Add("@AmpHandleLocation", SqlDbType.NVarChar, 100).Value = If(ampHandleLocation, CType(DBNull.Value, Object))
                p.Add("@TAHWidth", SqlDbType.Decimal).Value = If(tahWidth.HasValue, CType(tahWidth, Object), DBNull.Value)
                p.Add("@TAHHeight", SqlDbType.Decimal).Value = If(tahHeight.HasValue, CType(tahHeight, Object), DBNull.Value)
                p.Add("@TAHRearOffset", SqlDbType.Decimal).Value = If(tahRearOffset.HasValue, CType(tahRearOffset, Object), DBNull.Value)

                p.Add("@SAHHeight", SqlDbType.Decimal).Value = If(sahHeight.HasValue, CType(sahHeight, Object), DBNull.Value)
                p.Add("@SAHWidth", SqlDbType.Decimal).Value = If(sahWidth.HasValue, CType(sahWidth, Object), DBNull.Value)
                p.Add("@SAHRearOffset", SqlDbType.Decimal).Value = If(sahRearOffset.HasValue, CType(sahRearOffset, Object), DBNull.Value)
                p.Add("@SAHTopDownOffset", SqlDbType.Decimal).Value = If(sahTopDownOffset.HasValue, CType(sahTopDownOffset, Object), DBNull.Value)

                p.Add("@MusicRestDesign", SqlDbType.NVarChar, 200).Value = If(musicRestDesign, CType(DBNull.Value, Object))
                p.Add("@ChartTemplate", SqlDbType.NVarChar, 200).Value = If(chartTemplate, CType(DBNull.Value, Object))
                p.Add("@Notes", SqlDbType.NVarChar, -1).Value = If(notes, CType(DBNull.Value, Object))

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ' DbConnectionManager.vb
    Public Shared Function GetModelsForGrid(manufacturerId As Integer, seriesId As Integer) As DataTable
        Dim dt As New DataTable()
        Const sql As String = "
    SELECT
        m.PK_modelId                 AS PK_ModelId,
        m.modelName                  AS ModelName,            -- <== alias fixes the case
        m.parentSku                  AS ParentSku,
        m.width                      AS Width,
        m.depth                      AS Depth,
        m.height                     AS Height,
        m.totalFabricSquareInches    AS TotalFabricSquareInches,
        m.ampHandleLocation          AS AmpHandleLocation,
        m.tahWidth                   AS TAHWidth,
        m.tahHeight                  AS TAHHeight,
        m.tahRearOffset              AS TAHRearOffset,
        m.sahHeight                  AS SAHHeight,
        m.sahWidth                   AS SAHWidth,
        m.sahRearOffset              AS SAHRearOffset,
        m.sahTopDownOffset           AS SAHTopDownOffset,
        m.musicRestDesign            AS MusicRestDesign,
        m.chart_Template             AS Chart_Template,
        m.notes                      AS Notes,
        m.WooProductId               AS WooProductId
    FROM model m
    INNER JOIN ModelSeries s ON s.PK_SeriesId = m.FK_seriesId
    WHERE m.FK_seriesId = @sid AND s.FK_ManufacturerId = @mid
    ORDER BY m.modelName;"

        Using cn = GetConnection(), cmd As New SqlClient.SqlCommand(sql, cn)
            cmd.Parameters.Add("@sid", SqlDbType.Int).Value = seriesId
            cmd.Parameters.Add("@mid", SqlDbType.Int).Value = manufacturerId
            EnsureOpen(cn)
            Using da As New SqlClient.SqlDataAdapter(cmd)
                da.Fill(dt)
            End Using
        End Using
        Return dt
    End Function

End Class
