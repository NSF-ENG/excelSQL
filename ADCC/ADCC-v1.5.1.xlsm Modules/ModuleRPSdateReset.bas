Attribute VB_Name = "ModuleRPSdateReset"
Sub resetDates()
Attribute resetDates.VB_ProcData.VB_Invoke_Func = " \n14"
'
' resetDates Macro
'
    Range("N10").Select
    ActiveCell.FormulaR1C1 = "=rps_end_date-5*365"
    Range("O10").Select
    ActiveCell.FormulaR1C1 = "=from_date"
    Range("P10").Select
    ActiveCell.FormulaR1C1 = "=to_date+TIME(17,,)"
End Sub
