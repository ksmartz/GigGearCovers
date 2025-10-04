' Purpose: Handles navigation from Dashboard to AddModels and Listings forms.
' Dependencies: Imports System.Windows.Forms
' Current date: 2025-09-30

Imports System.Windows.Forms

Public Class formDashboard

    ' Handles Add Models button click
    Private Sub btnAddModels_Click(sender As Object, e As EventArgs) Handles btnAddModels.Click
        Dim addModelsForm As New formAddModels()
        formAddModels.Show()
        ' Me.Close() ' Close dashboard
    End Sub

    ' Handles Listings button click
    Private Sub btnListings_Click(sender As Object, e As EventArgs) Handles btnListings.Click
        Dim listingsForm As New formListings()
        formListings.Show()
        'Me.Close() ' Close dashboard
    End Sub

    Private Sub formDashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    '------------------------------------------------------------------------------
    ' Purpose: Handles Materials Suppliers button click to open FormMaterialsSuppliers.
    ' Dependencies: Imports System.Windows.Forms
    ' Current date: 2025-09-30
    '------------------------------------------------------------------------------

    Private Sub btnMaterialsSuppliers_Click(sender As Object, e As EventArgs) Handles btnMaterialsSuppliers.Click
        ' Open FormMaterialsSuppliers and close the dashboard
        Dim suppliersForm As New FormMaterialsSuppliers()
        suppliersForm.Show()
        ' Me.Close() ' Close dashboard
    End Sub
End Class