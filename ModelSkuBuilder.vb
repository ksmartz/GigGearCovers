Option Strict On
Option Infer On

Imports System.Text
Imports System.Text.RegularExpressions

Module ModelSkuBuilder









    '*****************************************************************************************************************************************************************
#Region "*************START*************PUBLIC FUNCTIONS**************************************************************************************************************"
    '*****************************************************************************************************************************************************************
    ''' <summary>
    ''' Generates a 32-character parent SKU using manufacturer, series, model, and optional version suffix.
    ''' </summary>
    Public Function GenerateParentSku(manufacturerName As String,
                                  seriesName As String,
                                  modelName As String,
                                  Optional versionSuffix As String = "V1",
                                  Optional uniqueId As Integer = 0) As String
        ' 1. Manufacturer segment (6 chars + "-")
        Dim manSeg = New String(manufacturerName.Where(Function(c) Char.IsLetterOrDigit(c)).ToArray()).ToUpperInvariant()
        manSeg = manSeg.PadRight(6, "X"c).Substring(0, 6) & "-"

        ' 2. Series segment (6 chars + "-")
        Dim serSeg = New String(seriesName.Where(Function(c) Char.IsLetterOrDigit(c)).ToArray()).ToUpperInvariant()
        serSeg = serSeg.PadRight(6, "X"c).Substring(0, 6) & "-"

        ' 3. Model segment (fill up to 32 - 6 - 1 - 6 - 1 - 2 - 5 = 11 chars)
        Dim modelSegLen As Integer = 32 - manSeg.Length - serSeg.Length - 2 - 5
        Dim modSeg = New String(modelName.Where(Function(c) Char.IsLetterOrDigit(c)).ToArray()).ToUpperInvariant()
        If modSeg.Length > modelSegLen Then
            modSeg = modSeg.Substring(0, modelSegLen)
        ElseIf modSeg.Length < modelSegLen Then
            modSeg = modSeg.PadRight(modelSegLen, "X"c)
        End If

        ' 4. Version (2 chars)
        Dim ver = If(String.IsNullOrEmpty(versionSuffix), "V1", versionSuffix.ToUpperInvariant())
        If ver.Length > 2 Then ver = ver.Substring(0, 2)

        ' 5. Last 5 digits of modelId, zero-padded
        Dim last5 = (uniqueId Mod 100000).ToString().PadLeft(5, "0"c)

        ' 6. Combine
        Dim sku = manSeg & serSeg & modSeg & ver & last5
        Return sku
    End Function




#End Region '---------END---------------PUBLIC FUNCTIONS--------------------------------------------------------------------------------------------------------------




    '*****************************************************************************************************************************************************************
#Region "*************START*************SEGMENT BUILDERS**************************************************************************************************************"
    '*****************************************************************************************************************************************************************
    Private Function MakeMaker3(name As String) As String
        Dim words = SplitWords(name)
        If words.Count >= 2 Then
            Dim a = FirstNAlpha(words(0), 1)
            Dim b = FirstNAlpha(words(1), 2)
            Dim code = (a & b).ToUpperInvariant()
            Return PadOrTrim(code, 3, "X"c)
        Else
            Dim code = FirstNAlpha(words(0), 3).ToUpperInvariant()
            Return PadOrTrim(code, 3, "X"c)
        End If
    End Function

    Private Function MakeSeries4(name As String) As String
        Dim words = SplitWords(name)
        Dim sb As New StringBuilder()
        For i = 0 To Math.Min(3, words.Count - 1)
            Dim ch = FirstAlpha(words(i))
            If ch <> ChrW(0) Then sb.Append(ch)
        Next
        If sb.Length < 4 AndAlso words.Count > 0 Then
            Dim w = Regex.Replace(words(0).ToUpperInvariant(), "[^A-Z0-9]", "")
            Dim pos = 1
            While sb.Length < 4 AndAlso pos < w.Length
                sb.Append(w(pos))
                pos += 1
            End While
        End If
        Dim code = sb.ToString().ToUpperInvariant()
        Return PadOrTrim(code, 4, "X"c)
    End Function

    Private Function MakeMakerInitials2(name As String) As String
        Dim words = SplitWords(name)
        If words.Count >= 2 Then
            Dim a = FirstAlpha(words(0))
            Dim b = FirstAlpha(words(1))
            Return $"{a}{b}".ToUpperInvariant()
        Else
            Dim w = Regex.Replace(words(0).ToUpperInvariant(), "[^A-Z0-9]", "")
            Return PadOrTrim(w, 2, "X"c)
        End If
    End Function

    Private Function MakeSeriesInitial1(name As String) As String
        Dim words = SplitWords(name)
        Dim ch = If(words.Count > 0, FirstAlpha(words(0)), ChrW(0))
        If ch = ChrW(0) Then ch = "X"c
        Return ch.ToString().ToUpperInvariant()
    End Function

    Private Function LeadingDigits2FromModelStart(model As String) As String
        Dim m = Regex.Match(model.Trim(), "^\s*(\d+)")
        If m.Success Then
            Dim d = m.Groups(1).Value
            If d.Length >= 2 Then
                Return d.Substring(0, 2)
            Else
                Return d.PadRight(2, "0"c)
            End If
        End If
        Return "00"
    End Function

    ''' <summary>
    ''' Generates a Parent SKU for a model.
    ''' </summary>
    ''' <param name="manufacturerName">The manufacturer name.</param>
    ''' <param name="seriesName">The series name.</param>
    ''' <param name="modelName">The model name.</param>
    ''' <param name="versionSuffix">The version suffix.</param>
    ''' <param name="modelId">The model ID.</param>
    ''' <returns>The generated Parent SKU.</returns>
    ' Builds a 12-character model segment, reserving the last 2 for versioning.
    Private Function BuildModel12(model As String, mi2 As String, ld2 As String, si1 As String, versionSuffix As String) As String
        Dim baseModelLen As Integer = 10 ' was 8 in BuildModel10, now 10 for a total of 12 with version
        Dim tokens = SplitTokensAlnum(model.ToUpperInvariant())
        Dim idx As Integer = 0

        Dim baseDigits As String = ""
        If idx < tokens.Count AndAlso Regex.IsMatch(tokens(0), "^\d+$") Then
            baseDigits = tokens(0)
            idx += 1
        End If
        If baseDigits.Length > 2 Then baseDigits = baseDigits.Substring(0, 2)

        Dim neededLetters As Integer = 7 - baseDigits.Length ' was 5, now 7
        Dim baseLetters As New StringBuilder()
        While baseLetters.Length < neededLetters AndAlso idx < tokens.Count
            Dim t = tokens(idx)
            If Regex.IsMatch(t, "^[A-Z]+$") Then
                baseLetters.Append(t(0))
            ElseIf Regex.IsMatch(t, "^[A-Z0-9]+$") Then
                Dim c = t.FirstOrDefault(Function(ch) Char.IsLetter(ch))
                If c <> ChrW(0) Then baseLetters.Append(c)
            End If
            idx += 1
        End While
        If baseLetters.Length < neededLetters AndAlso idx > 0 Then
            Dim lastWord = tokens(Math.Max(0, idx - 1))
            Dim pos As Integer = 1
            While baseLetters.Length < neededLetters AndAlso pos < lastWord.Length
                If Char.IsLetter(lastWord(pos)) Then baseLetters.Append(lastWord(pos))
                pos += 1
            End While
        End If

        Dim base10 As String = (baseDigits & baseLetters.ToString()).PadRight(baseModelLen, "X"c)
        If base10.Length > baseModelLen Then base10 = base10.Substring(0, baseModelLen)

        ' Always append versionSuffix (should be 2 chars)
        Dim mod12 As String = base10 & versionSuffix.PadRight(2, "X"c)
        If mod12.Length > 12 Then mod12 = mod12.Substring(0, 12)

        Return mod12
    End Function


#End Region '---------END---------------SEGMENT BUILDERS--------------------------------------------------------------------------------------------------------------



    '*****************************************************************************************************************************************************************
#Region "*************START*************HELPERS**************************************************************************************************************"
    '*****************************************************************************************************************************************************************
    Private Function SplitWords(text As String) As List(Of String)
        Dim cleaned = Regex.Replace(text.ToUpperInvariant(), "[^A-Z0-9\s\-]", " ")
        Dim parts = Regex.Split(cleaned, "[^A-Z0-9]+").Where(Function(s) s.Length > 0).ToList()
        If parts.Count = 0 Then parts.Add("X")
        Return parts
    End Function

    Private Function SplitTokensAlnum(text As String) As List(Of String)
        Dim parts = Regex.Split(text, "[^A-Z0-9]+").Where(Function(s) s.Length > 0).ToList()
        If parts.Count = 0 Then parts.Add("X")
        Return parts
    End Function

    Private Function FirstAlpha(token As String) As Char
        For Each ch In token.ToUpperInvariant()
            If ch >= "A"c AndAlso ch <= "Z"c Then Return ch
        Next
        Return ChrW(0)
    End Function

    Private Function FirstNAlpha(token As String, n As Integer) As String
        Dim sb As New StringBuilder()
        For Each ch In token.ToUpperInvariant()
            If ch >= "A"c AndAlso ch <= "Z"c Then
                sb.Append(ch)
                If sb.Length = n Then Exit For
            End If
        Next
        Return sb.ToString()
    End Function

    Private Function PadOrTrim(s As String, length As Integer, pad As Char) As String
        Dim t = Regex.Replace(s.ToUpperInvariant(), "[^A-Z0-9]", "")
        If t.Length > length Then Return t.Substring(0, length)
        If t.Length < length Then Return t.PadRight(length, pad)
        Return t
    End Function


#End Region '---------END---------------HELPERS--------------------------------------------------------------------------------------------------------------






End Module
