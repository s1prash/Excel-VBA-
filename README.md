' Excel-VBA code for ASTM E2709-11: Acceptance Limit Table for Single Stage written in SAS

Option Explicit
Sub astm()
'acceptance limit table for single stage
Const D As Integer = 1
Const cilevel As Integer = 95
Const lbond As Integer = 95
Const numb As Integer = 30
Dim n1 As Integer
n1 = 5
Dim z As Double
z = Application.WorksheetFunction.Norm_S_Inv((1 + (cilevel / 100) ^ 0.5) / 2)
Dim chi As Double
chi = Application.WorksheetFunction.ChiSq_Inv(1 - (cilevel / 100) ^ 0.5, numb - 1)
Dim sdold As Double
sdold = 0
Dim startsd As Double
startsd = 0.01
Dim mean As Double
For mean = 96 To 104
    Dim sampsd As Double
    For sampsd = startsd To 20 Step 0.001
        Dim sig As Double
        sig = ((numb - 1) * sampsd * sampsd / chi) ^ 0.5
        Dim llu As Double
        llu = mean - z * sig / (numb) ^ 0.5
        Dim p1L As Double
        p1L = ((Application.WorksheetFunction.NormSDist((105 - llu) / sig)) - Application.WorksheetFunction.NormSDist((95 - llu) / sig)) ^ n1
        Dim ulu As Double
        ulu = mean + z * sig / (numb) ^ 0.5
        Dim p1U As Double
        p1U = ((Application.WorksheetFunction.NormSDist((105 - ulu) / sig)) - Application.WorksheetFunction.NormSDist((95 - ulu) / sig)) ^ n1
        Dim overbd As Double
        overbd = Application.WorksheetFunction.Min(p1L, p1U)
            If overbd < lbond / 100 Then
                sampsd = sampsd - 0.001
                Dim cv As Double
                cv = 100 * sampsd / mean
                Range("a1").Value = "Mean"
                Range("b1").Value = "Standard deviation"
                Range("c1").Value = "%RSD"
                Range("d1").Value = "Lower Bound"
                Range("e1").Value = "Upper Bound"
                Range("a" & Rows.Count).End(xlUp).Offset(1, 0) = mean
                Range("b" & Rows.Count).End(xlUp).Offset(1, 0) = sampsd
                Range("c" & Rows.Count).End(xlUp).Offset(1, 0) = cv
                Range("d" & Rows.Count).End(xlUp).Offset(1, 0) = p1L
                Range("e" & Rows.Count).End(xlUp).Offset(1, 0) = p1U
                  sampsd = 20
            End If
        Next sampsd
    Next mean
End Sub
