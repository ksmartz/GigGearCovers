Imports System.Windows.Forms

' Purpose: Displays a selectable, scrollable JSON payload for review before upload.
' Dependencies: Imports System.Windows.Forms
' Current date: 2025-09-27
Public Class JsonReviewForm
    Inherits Form

    Private txtJson As TextBox
    Private btnContinue As Button
    Private btnCancel As Button

    Public Sub New(jsonText As String)
        Me.Text = "Review Listing Payload"
        Me.Size = New Drawing.Size(700, 500)
        Me.StartPosition = FormStartPosition.CenterParent

        txtJson = New TextBox() With {
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = ScrollBars.Both,
            .WordWrap = False,
            .Dock = DockStyle.Top,
            .Font = New Drawing.Font("Consolas", 10),
            .Height = 400,
            .Text = jsonText
        }

        btnContinue = New Button() With {
            .Text = "Continue",
            .DialogResult = DialogResult.OK,
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right,
            .Left = 480,
            .Top = 410,
            .Width = 100
        }
        btnCancel = New Button() With {
            .Text = "Cancel",
            .DialogResult = DialogResult.Cancel,
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right,
            .Left = 590,
            .Top = 410,
            .Width = 100
        }

        Me.Controls.Add(txtJson)
        Me.Controls.Add(btnContinue)
        Me.Controls.Add(btnCancel)

        Me.AcceptButton = btnContinue
        Me.CancelButton = btnCancel
    End Sub
End Class
