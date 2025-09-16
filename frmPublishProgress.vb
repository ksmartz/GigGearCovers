Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms


'==============================================================================
' Class: frmPublishProgress
' Purpose:
'   Lightweight progress window for long-running WooCommerce publish batches.
'   - Shows a progress bar + count label
'   - Streams log lines into a scrolling ListBox
'   - Supports Cancel (cancels the batch BEFORE starting the next item)
'
' Dependencies: None (pure WinForms). Create and use from any form.
' Date: 2025-09-15
'------------------------------------------------------------------------------
Public Class frmPublishProgress
    Inherits Form

    Private ReadOnly _lblTitle As New Label()
    Private ReadOnly _progress As New ProgressBar()
    Private ReadOnly _lblCount As New Label()
    Private ReadOnly _lst As New ListBox()
    Private ReadOnly _btnCancel As New Button()
    Private _total As Integer = 0
    Private _current As Integer = 0
    Private _cancelRequested As Boolean = False

    Public Sub New()
        Me.Text = "Publishing to WooCommerce..."
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Width = 720
        Me.Height = 520

        _lblTitle.AutoSize = True
        _lblTitle.Top = 12
        _lblTitle.Left = 12
        _lblTitle.Text = "Starting..."

        _progress.Left = 12
        _progress.Top = _lblTitle.Bottom + 8
        _progress.Width = Me.ClientSize.Width - 24
        _progress.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        _progress.Minimum = 0
        _progress.Step = 1

        _lblCount.AutoSize = True
        _lblCount.Left = 12
        _lblCount.Top = _progress.Bottom + 8
        _lblCount.Text = "0 / 0"

        _lst.Left = 12
        _lst.Top = _lblCount.Bottom + 8
        _lst.Width = Me.ClientSize.Width - 24
        _lst.Height = Me.ClientSize.Height - _lst.Top - 56
        _lst.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right

        _btnCancel.Text = "Cancel"
        _btnCancel.Width = 100
        _btnCancel.Height = 28
        _btnCancel.Left = Me.ClientSize.Width - _btnCancel.Width - 12
        _btnCancel.Top = Me.ClientSize.Height - _btnCancel.Height - 12
        _btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        AddHandler _btnCancel.Click, AddressOf OnCancelClicked

        Me.Controls.AddRange(New Control() {_lblTitle, _progress, _lblCount, _lst, _btnCancel})

        ' Ensure layout updates when window is resized
        AddHandler Me.Resize, Sub(sender, e)
                                  _progress.Width = Me.ClientSize.Width - 24
                                  _lst.Width = Me.ClientSize.Width - 24
                                  _lst.Height = Me.ClientSize.Height - _lst.Top - 56
                                  _btnCancel.Left = Me.ClientSize.Width - _btnCancel.Width - 12
                                  _btnCancel.Top = Me.ClientSize.Height - _btnCancel.Height - 12
                              End Sub
    End Sub

    ' --- Public API -----------------------------------------------------------

    ''' <summary>Initialize total count and reset UI.</summary>
    Public Sub StartBatch(total As Integer, Optional title As String = Nothing)
        If Me.InvokeRequired Then
            Me.Invoke(New Action(Of Integer, String)(AddressOf StartBatch), total, title)
            Return
        End If
        _total = Math.Max(0, total)
        _current = 0
        _cancelRequested = False
        _progress.Maximum = If(_total <= 0, 1, _total)
        _progress.Value = 0
        _lblCount.Text = $"0 / {_total}"
        _lblTitle.Text = If(String.IsNullOrWhiteSpace(title), "Publishing models…", title)
        _lst.Items.Clear()
    End Sub

    ''' <summary>Append a log line (threadsafe).</summary>
    Public Sub LogLine(message As String)
        If Me.InvokeRequired Then
            Me.Invoke(New Action(Of String)(AddressOf LogLine), message)
            Return
        End If
        If String.IsNullOrEmpty(message) Then Return
        _lst.Items.Add($"{DateTime.Now:HH:mm:ss}  {message}")
        _lst.TopIndex = _lst.Items.Count - 1
    End Sub

    ''' <summary>Advance progress by one and update the count label.</summary>
    Public Sub Advance(Optional currentItemLabel As String = Nothing)
        If Me.InvokeRequired Then
            Me.Invoke(New Action(Of String)(AddressOf Advance), currentItemLabel)
            Return
        End If
        _current = Math.Min(_current + 1, _total)
        If _progress.Value < _progress.Maximum Then
            _progress.PerformStep()
        End If
        _lblCount.Text = $"{_current} / {_total}"
        If Not String.IsNullOrWhiteSpace(currentItemLabel) Then
            _lblTitle.Text = currentItemLabel
        End If
    End Sub

    ''' <summary>True if the user clicked Cancel.</summary>
    Public ReadOnly Property IsCancelRequested As Boolean
        Get
            If Me.InvokeRequired Then
                Dim result As Boolean = False
                Me.Invoke(New Action(Sub() result = _cancelRequested))
                Return result
            End If
            Return _cancelRequested
        End Get
    End Property

    ''' <summary>Set a custom title/status line.</summary>
    Public Sub SetStatus(title As String)
        If Me.InvokeRequired Then
            Me.Invoke(New Action(Of String)(AddressOf SetStatus), title)
            Return
        End If
        _lblTitle.Text = title
    End Sub

    ' --- Internals ------------------------------------------------------------

    Private Sub OnCancelClicked(sender As Object, e As EventArgs)
        _cancelRequested = True
        _btnCancel.Enabled = False
        LogLine("Cancel requested. Finishing current item, then stopping…")
        _lblTitle.Text = "Cancel requested…"
    End Sub
End Class

