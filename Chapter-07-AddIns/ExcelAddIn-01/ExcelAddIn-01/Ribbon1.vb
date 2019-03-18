Imports Microsoft.Office.Tools.Ribbon


Public Class AtifNaseem

    
    Private Sub BtnLCase_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnLCase.Click
        Dim RNG As Excel.Range
        For Each RNG In Globals.ThisAddIn.Application.Selection
            If RNG.Value <> "" And Not (IsNumeric(RNG.Value)) Then
                RNG.Value = LCase(RNG.Value)
            End If
        Next RNG
    End Sub

    Private Sub BtnUCase_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnUCase.Click
        Dim RNG As Excel.Range
        For Each RNG In Globals.ThisAddIn.Application.Selection
            If RNG.Value <> "" And Not (IsNumeric(RNG.Value)) Then
                RNG.Value = UCase(RNG.Value)
            End If
        Next RNG
    End Sub

    Private Sub BtnPCase_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnPCase.Click
        Dim RNG As Excel.Range
        For Each RNG In Globals.ThisAddIn.Application.Selection
            If RNG.Value <> "" And Not (IsNumeric(RNG.Value)) Then
                RNG.Value = StrConv(RNG.Value, vbProperCase)
            End If
        Next RNG
    End Sub

    Private Sub BtnSCase_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnSCase.Click
        Dim RNG As Excel.Range
        For Each RNG In Globals.ThisAddIn.Application.Selection
            If RNG.Value <> "" And Not (IsNumeric(RNG.Value)) Then
                RNG.Value = UCase(Left(RNG.Value, 1)) & LCase(Mid(RNG.Value, 2))
            End If
        Next RNG
    End Sub

    Private Sub BtnTxt2Num_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnTxt2Num.Click
        Dim RNG As Excel.Range
        For Each RNG In Globals.ThisAddIn.Application.Selection
            If RNG.Value <> "" And Val(RNG.Value) <> 0 Then
                RNG.Value = Val(RNG.Value)
            End If
        Next RNG
    End Sub

    Private Sub BtnNum2Txt_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnNum2Txt.Click
        Dim RNG As Excel.Range
        For Each RNG In Globals.ThisAddIn.Application.Selection
            If IsNumeric(RNG.Value) Then
                RNG.Formula = "'" & RNG.Value
            End If
        Next RNG
    End Sub
End Class
