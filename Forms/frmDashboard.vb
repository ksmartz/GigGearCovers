Imports System.Text.Json
Imports System.Windows.Forms

Public Class frmDashboard
    Private Sub btnAddSupplier_Click(sender As Object, e As EventArgs) Handles btnAddSupplier.Click
        Dim supplierForm As New frmSuppliers()
        supplierForm.ShowDialog()
    End Sub

    Private Sub btnMaterialsDataEntry_Click(sender As Object, e As EventArgs) Handles btnMaterialsDataEntry.Click
        Dim fabricEntryForm As New frmFabricEntryForm()
        fabricEntryForm.ShowDialog() ' Use Show() if you want it non-modal
    End Sub

    Private Sub btnAddModels_Click(sender As Object, e As EventArgs) Handles btnAddModels.Click
        Dim ModelInformationForm As New frmAddModelInformation()
        ModelInformationForm.ShowDialog()
    End Sub



    Private Sub frmDashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub


End Class