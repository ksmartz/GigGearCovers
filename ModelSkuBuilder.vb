Option Strict On
Option Infer On

Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Linq

Module ModelSkuBuilder

    Private Const SKU_MAX As Integer = 32

#Region "PUBLIC FUNCTIONS"

    ' Parent builder: keep as you already have it (unchanged)
    Public Function GenerateParentSku(manufacturerName As String,
                                      seriesName As String,
                                      modelName As String,
                                      Optional versionSuffix As String = "V1",
                                      Optional uniqueId As Integer = 0) As String
        Dim manSeg = New String(If(manufacturerName, "").Where(Function(c) Char.IsLetterOrDigit(c)).ToArray()).ToUpperInvariant()
        manSeg = manSeg.PadRight(6, "X"c).Substring(0, 6) & "-"

        Dim serSeg = New String(If(seriesName, "").Where(Function(c) Char.IsLetterOrDigit(c)).ToArray()).ToUpperInvariant()
        serSeg = serSeg.PadRight(6, "X"c).Substring(0, 6) & "-"

        Dim modelSegLen As Integer = SKU_MAX - manSeg.Length - serSeg.Length - 2 - 5
        Dim modSeg = New String(If(modelName, "").Where(Function(c) Char.IsLetterOrDigit(c)).ToArray()).ToUpperInvariant()
        If modSeg.Length > modelSegLen Then
            modSeg = modSeg.Substring(0, modelSegLen)
        ElseIf modSeg.Length < modelSegLen Then
            modSeg = modSeg.PadRight(modelSegLen, "X"c)
        End If

        Dim ver = If(String.IsNullOrEmpty(versionSuffix), "V1", versionSuffix.ToUpperInvariant())
        If ver.Length > 2 Then ver = ver.Substring(0, 2)

        Dim last5 = (uniqueId Mod 100000).ToString().PadLeft(5, "0"c)

        Return (manSeg & serSeg & modSeg & ver & last5)
    End Function

    ''' <summary>
    ''' Child SKU from codes:
    ''' base = cleaned parent minus last 7 chars (ver2 + last5); result = base & "-F-CCC"
    ''' Always returns a non-empty, <=32 char string.
    ''' </summary>
    Public Function GenerateChildSkuFromParentCodes(parentSku As String,
                                                    fabricCode As String,
                                                    colorCode As String) As String
        ' Clean the parent (uppercase, keep A-Z/0-9/- only)
        Dim cleanParent As String = CleanSku(If(parentSku, ""))

        ' Base is parent without the reserved tail (2 + 5)
        Dim basePart As String = If(cleanParent.Length >= 7,
                                    cleanParent.Substring(0, cleanParent.Length - 7),
                                    String.Empty)

        ' Normalize codes
        Dim f As String = SanitizeLen(fabricCode, 1)  ' 1 char
        Dim c As String = SanitizeLen(colorCode, 3)   ' 3 chars

        ' Build suffix
        Dim suffix As String = "-" & f & "-" & c

        ' Avoid double hyphens like "...-" + "-F-CCC"
        basePart = basePart.TrimEnd("-"c)

        ' Compose
        Dim sku As String = basePart & suffix

        ' If parent was missing/too short, sku starts with "-", still valid but non-empty.
        ' Clamp to 32
        If sku.Length > SKU_MAX Then
            Dim maxBase = Math.Max(0, SKU_MAX - suffix.Length)
            basePart = If(basePart.Length > maxBase, basePart.Substring(0, maxBase), basePart)
            sku = basePart & suffix
        End If

        ' Final safety: never let blank slip out
        If String.IsNullOrWhiteSpace(sku) Then
            sku = "X" & suffix    ' e.g., "X-C-BLK"
            If sku.Length > SKU_MAX Then sku = sku.Substring(0, SKU_MAX)
        End If

        Return sku
    End Function

#End Region

#Region "HELPERS (unchanged + sanitize)"
    Private Function CleanSku(s As String) As String
        If String.IsNullOrEmpty(s) Then Return ""
        ' Keep A-Z, 0-9, and hyphen; uppercase; collapse multiple hyphens
        Dim up = s.ToUpperInvariant()
        Dim filtered = New String(up.Where(Function(ch) Char.IsLetterOrDigit(ch) OrElse ch = "-"c).ToArray())
        ' collapse sequences of '-' to a single '-'
        filtered = Regex.Replace(filtered, "-{2,}", "-")
        Return filtered.Trim()
    End Function

    Private Function SplitWords(text As String) As List(Of String)
        Dim cleaned = Regex.Replace(If(text, "").ToUpperInvariant(), "[^A-Z0-9\s\-]", " ")
        Dim parts = Regex.Split(cleaned, "[^A-Z0-9]+").Where(Function(s) s.Length > 0).ToList()
        If parts.Count = 0 Then parts.Add("X")
        Return parts
    End Function

    Private Function SplitTokensAlnum(text As String) As List(Of String)
        Dim parts = Regex.Split(If(text, ""), "[^A-Z0-9]+").Where(Function(s) s.Length > 0).ToList()
        If parts.Count = 0 Then parts.Add("X")
        Return parts
    End Function

    Private Function FirstAlpha(token As String) As Char
        For Each ch In If(token, "").ToUpperInvariant()
            If ch >= "A"c AndAlso ch <= "Z"c Then Return ch
        Next
        Return ChrW(0)
    End Function

    Private Function FirstNAlpha(token As String, n As Integer) As String
        Dim sb As New StringBuilder()
        For Each ch In If(token, "").ToUpperInvariant()
            If ch >= "A"c AndAlso ch <= "Z"c Then
                sb.Append(ch)
                If sb.Length = n Then Exit For
            End If
        Next
        Return sb.ToString()
    End Function

    Private Function PadOrTrim(s As String, length As Integer, pad As Char) As String
        Dim t = Regex.Replace(If(s, "").ToUpperInvariant(), "[^A-Z0-9]", "")
        If t.Length > length Then Return t.Substring(0, length)
        If t.Length < length Then Return t.PadRight(length, pad)
        Return t
    End Function

    ''' <summary>Uppercase alnum, then trim/pad to fixed length with 'X'.</summary>
    Private Function SanitizeLen(s As String, n As Integer) As String
        Dim t = If(s, "").ToUpperInvariant()
        t = New String(t.Where(Function(ch) Char.IsLetterOrDigit(ch)).ToArray())
        If t.Length > n Then t = t.Substring(0, n)
        If t.Length < n Then t = t.PadRight(n, "X"c)
        Return t
    End Function
#End Region

End Module
