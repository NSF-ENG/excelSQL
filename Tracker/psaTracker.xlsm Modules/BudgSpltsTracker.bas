Attribute VB_Name = "BudgSpltsTracker"


Sub RefreshActiveSheetBudgSplts() ' Everytime user clicks the refresh button
Dim dateWhere As String ' nsf_rcvd_date range
Dim dateDDWhere As String 'dd_rcm_date range
Dim budgYear As String 'budg_yr
Dim dateLast As String 'last update
Dim datePRC As String ' nsf_rcvd_date or budg_yr
Dim addBudgSplts As String ' additional proposals
Dim start As String ' True or false to start where
Dim CRNL As String ' concatenate
CRNL = Chr(13) & Chr(10)

Call handlePwd

'-----------DATES------------------------


If hasValue("from_date") Then dateWhere = "AND p.nsf_rcvd_date >= {ts '" & Format(ActiveSheet.Range("from_date"), "yyyy-mm-dd hh:mm:ss") & "'}"
If hasValue("to_date") Then dateWhere = dateWhere & " AND p.nsf_rcvd_date <= {ts '" & Format(ActiveSheet.Range("to_date"), "yyyy-mm-dd hh:mm:ss") & "'} "
If hasValue("dd_from_date") Then dateDDWhere = "AND p.dd_rcom_date >= {ts '" & Format(ActiveSheet.Range("dd_from_date"), "yyyy-mm-dd hh:mm:ss") & "'}"
If hasValue("dd_to_date") Then dateDDWhere = dateDDWhere & " AND p.dd_rcom_date <= {ts '" & Format(ActiveSheet.Range("dd_to_date"), "yyyy-mm-dd hh:mm:ss") & "'} "
If hasValue("budg_yr") Then budgYear = andWhere("b", "budg_yr") & vbNewLine ', isIntField:=True
If hasValue("last_updt_tmsp") Then dateLast = "AND b.last_updt_tmsp >= {ts '" & Format(ActiveSheet.Range("last_updt_tmsp"), "yyyy-mm-dd hh:mm:ss") & "'}"
 
If (hasValue("from_date") Or hasValue("dd_from_date") Or hasValue("budg_yr") Or hasValue("last_updt_tmsp")) Then
     start = "(1=1)" 'True
Else
    start = "(0=1)" 'False
    If Len(ActiveSheet.Range("add_Budg_Splts").Rows.count) < 1 Then ' no specific prop_id
        MsgBox "Include at least one from date, or include proposal numbers in the Add table."
        End
    End If
End If


Dim myBSplit As String ' Main query to restrict resuls with optional parameters
Dim myProp As String ' important information related proposals
Dim myBudgPRCs As String ' query used to get PRCS by budget splits
Dim myPropPRCs As String ' query used to get PRCS by proposal
Dim myResults As String ' showing all fields that are going to be shown in spreadsheet, tip(check option in external properies in excel to perserve sort, column, order of data)
Dim dropTables As String
Dim budg_spltsInclude  As String
Dim budg_spltsExclude  As String
'------------------------------------------Include/Exclude Budget Splits----------------------------------------------
'Table created  from add_budg_splts
budg_spltsInclude = "SET NOCOUNT ON CREATE TABLE #AddBudgSplts(" & CRNL _
& "prop_id char(7)," & CRNL _
& "budg_yr smallint null," & CRNL _
& "splt_id char(2) null" & CRNL _
& ") " & CRNL

'Table created  from omit_budg_splts
budg_spltsExclude = "CREATE TABLE #OmitBudgSplts(" & CRNL _
& "prop_id char(7)," & CRNL _
& "budg_yr smallint null," & CRNL _
& "splt_id char(2) null" & CRNL _
& ") " & CRNL

'Include
Set rng = ActiveSheet.Range("add_budg_splts") ' get range of tables add_budg_splts

For Each Row In rng.Rows 'value loop to populate table
 If Row.Cells(1).Value = "" And (Row.Cells(2).Value <> "" Or Row.Cells(3).Value <> "") Then
 MsgBox "Include prop_id in row " & Row.Address
 End
 ElseIf Row.Cells(1).Value = "" And (hasValue("from_date") = False And hasValue("dd_from_date") = False And hasValue("budg_yr") = False And hasValue("last_updt_tmsp") = False) Then
  MsgBox "Include dates or prop_id" ' user has to use parameters
    End
 ElseIf Row.Cells(1).Value <> "" Then
 ' Include budget splits from add_budg_splts table
 
  budg_spltsInclude = budg_spltsInclude _
 & " Insert into #AddBudgSplts(prop_id ,budg_yr,splt_id) " & CRNL
 
 If Row.Cells(2).Value = "" And Row.Cells(3).Value <> "" Then
   budg_spltsInclude = budg_spltsInclude _
 & "VALUES('" & Row.Cells(1).Value & "',NULL,'" & Row.Cells(3).Value & "') " & CRNL
 
 ElseIf Row.Cells(2).Value <> "" And Row.Cells(3).Value = "" Then
   budg_spltsInclude = budg_spltsInclude _
 & "VALUES('" & Row.Cells(1).Value & "'," & Row.Cells(2).Value & ",NULL) " & CRNL
 
 ElseIf Row.Cells(2).Value = "" And Row.Cells(3).Value = "" Then
   budg_spltsInclude = budg_spltsInclude _
 & "VALUES('" & Row.Cells(1).Value & "',NULL,NULL) " & CRNL
 
 Else
   budg_spltsInclude = budg_spltsInclude _
 & "VALUES('" & Row.Cells(1).Value & "'," & Row.Cells(2).Value & ",'" & Row.Cells(3).Value & "') " & CRNL
 End If
 
End If

Next Row

'Exclude
Set rng2 = ActiveSheet.Range("omit_budg_splts") ' get range of tables  omit_budg_splts

For Each Row In rng2.Rows 'value loop to populate table to exlude budget splits
  ' Include budget splits from omit_budg_splts table
   If Row.Cells(1).Value = "" And (Row.Cells(2).Value <> "" Or Row.Cells(3).Value <> "") Then
 MsgBox "Include prop_id"
 End
 ElseIf Row.Cells(1).Value = "" And (hasValue("from_date") = False And hasValue("dd_from_date") = False And hasValue("budg_yr") = False And hasValue("last_updt_tmsp") = False) Then
  MsgBox "Include dates or prop_id" ' user has to use parameters
    End
  ElseIf Row.Cells(1).Value <> "" Then
 ' Include budget splits from add_budg_splts table
 
 
  budg_spltsExclude = budg_spltsExclude _
 & " Insert into #OmitBudgSplts(prop_id ,budg_yr,splt_id) " & CRNL
 
 If Row.Cells(2).Value = "" And Row.Cells(3).Value <> "" Then
  budg_spltsExclude = budg_spltsExclude _
 & "VALUES('" & Row.Cells(1).Value & "',NULL,'" & Row.Cells(3).Value & "') " & CRNL
 
 ElseIf Row.Cells(2).Value <> "" And Row.Cells(3).Value = "" Then
       budg_spltsExclude = budg_spltsExclude _
 & "VALUES('" & Row.Cells(1).Value & "'," & Row.Cells(2).Value & ",NULL) " & CRNL
 
 ElseIf Row.Cells(2).Value = "" And Row.Cells(3).Value = "" Then
  budg_spltsExclude = budg_spltsExclude _
 & "VALUES('" & Row.Cells(1).Value & "',NULL,NULL) " & CRNL
 
 Else
     budg_spltsExclude = budg_spltsExclude _
 & "VALUES('" & Row.Cells(1).Value & "'," & Row.Cells(2).Value & ",'" & Row.Cells(3).Value & "') " & CRNL
 End If
 
End If

 
Next Row

'---------------------------------------------------------------------------------------------------------------
'QUERY GETS BUDG_SPLIT MAIN FIELDS
myBSplit = "SELECT b.prop_id,b.budg_yr,b.splt_id,b.budg_splt_tot_dol, b.org_code as Bdg_Org_Code," & CRNL _
& "b.pgm_ele_code+' - '+pe.pgm_ele_name as PEC_bdg_splt" & CRNL _
& "INTO #myBSplit" & CRNL _
& "from csd.budg_splt b" & CRNL _
& "JOIN csd.prop p on p.prop_id=b.prop_id" & CRNL _
& "JOIN csd.pgm_ele  as pe ON pe.pgm_ele_code=b.pgm_ele_code" & CRNL _
& "JOIN csd.prop_stts ps on ps.prop_stts_code=p.prop_stts_code" & CRNL _
& "JOIN csd.natr_rqst nr on nr.natr_rqst_code = p.natr_rqst_code" & CRNL _
& "JOIN csd.awd_istr ai on p.rcom_awd_istr = ai.awd_istr_code" & CRNL _
& "JOIN csd.org  as og on og.org_code=p.org_code" & CRNL

myBSplit = myBSplit _
& "WHERE ((" & start & dateWhere & dateDDWhere & budgYear & dateLast & CRNL _
& andWhere("p", "pgm_annc_id") & CRNL _
& andWhere("", "b.org_code") & andWhere("p", "org_code") & CRNL _
& andWhere("", "b.pgm_ele_code") & andWhere("p", "pgm_ele_code") & CRNL _
& andWhere("ps", "prop_stts_abbr") & andWhere("p", "obj_clas_code") & andWhere("", "b.obj_clas_code") & CRNL _
& andWhere("nr", "natr_rqst_abbr") & CRNL _
& andWhere("ai", "awd_istr_abbr") & CRNL _
& andWhere("p", "pm_ibm_logn_id") & andWhere("og", "dir_div_abbr") & andWhere("", "b.pm_ibm_logn_id") & CRNL _
& andWhere("p", "prop_titl_txt") & CRNL

'-----------CASE FOR PROP PRCS INCLUDE/EXCLUDE--------------------------------------
field = Trim(ActiveSheet.Range("pa.prop_atr_code").Value)
If Left(field, 1) = "~" Then  ' have negation,Prop PRCS.
myBSplit = myBSplit _
    & excludePRCS("csd.prop_atr pa ", "", "pa.prop_atr_code", " and pa.prop_id=b.prop_id AND pa.prop_atr_type_code='PRC' ") & CRNL
Else ' Include Budg PRCS

myBSplit = myBSplit _
    & includePRCS("csd.prop_atr pa ", "", "pa.prop_atr_code", " and pa.prop_id=b.prop_id AND pa.prop_atr_type_code='PRC' ") & CRNL
End If
'-----------CASE FOR BUDG PRCS INCLUDE-----------------------------------------------
field = Trim(ActiveSheet.Range("pgm_ref_code").Value) ' Budg PRCS Include

If Left(field, 1) = "~" Then ' It should not be negative
    MsgBox "Negation will be ignored, if you want to exclude budg PRCS, use the budg PRCS exclude input" ' includePRCs will get rid of ~
    myBSplit = myBSplit _
    & includePRCS("csd.budg_pgm_ref bpr ", "bpr", "pgm_ref_code", " and bpr.prop_id=b.prop_id AND bpr.splt_id=b.splt_id AND bpr.budg_yr=b.budg_yr ") & CRNL
Else ' Include Budg PRCS, with no problem of ~
    myBSplit = myBSplit _
    & includePRCS("csd.budg_pgm_ref bpr ", "bpr", "pgm_ref_code", " and bpr.prop_id=b.prop_id AND bpr.splt_id=b.splt_id AND bpr.budg_yr=b.budg_yr ") & CRNL
End If

'-----------CASE FOR BUDG PRCS EXCLUDE-----------------------------------------------
'excludePRCS gets rid of ~
field = Trim(ActiveSheet.Range("bpr.pgm_ref_code").Value)
If Left(field, 1) = "~" Then
    MsgBox "Tilde will be removed, PRCS are going to be excluded with/without tilde ~"
    myBSplit = myBSplit _
    & excludePRCS("csd.budg_pgm_ref bpr ", "", "bpr.pgm_ref_code", " and bpr.prop_id=b.prop_id AND bpr.splt_id=b.splt_id AND bpr.budg_yr=b.budg_yr ") & CRNL
Else
     myBSplit = myBSplit _
    & excludePRCS("csd.budg_pgm_ref bpr ", "", "bpr.pgm_ref_code", " and bpr.prop_id=b.prop_id AND bpr.splt_id=b.splt_id AND bpr.budg_yr=b.budg_yr ") & CRNL
End If


'-----------Include(addBudgSplts)-----------------------------------------------
 myBSplit = myBSplit _
 & "))" & CRNL
 ' Insert is used to add extra budget splits, it has to have same # of fields as #myBSplit
   myBSplit = myBSplit _
 & "INSERT INTO #myBSplit " & CRNL _
 & "select  bs.prop_id," & CRNL _
 & "bs.budg_yr, bs.splt_id," & CRNL _
 & "bs.budg_splt_tot_dol, bs.org_code as Bdg_Org_Code,bs.pgm_ele_code+' - '+pe.pgm_ele_name as PEC_bdg_splt" & CRNL _
 & "from #AddBudgSplts t" & CRNL _
 & "JOIN CSD.budg_splt bs ON bs.prop_id=t.prop_id  " & CRNL _
 & "and isnull(t.budg_yr,bs.budg_yr) = bs.budg_yr " & CRNL _
 & "and isnull(t.splt_id,bs.splt_id) = bs.splt_id " & CRNL _
& "JOIN csd.pgm_ele  as pe ON pe.pgm_ele_code=bs.pgm_ele_code" & CRNL

   'Delete is used to remove  budget splits based on prop_id, budg_yr, and splt_id
  myBSplit = myBSplit _
 & "delete from #myBSplit" & CRNL _
 & "from  #myBSplit b" & CRNL _
 & "Join  #OmitBudgSplts as t" & CRNL _
 & "ON b.prop_id=t.prop_id " & CRNL _
 & "and isnull(t.budg_yr,b.budg_yr) = b.budg_yr " & CRNL _
 & "and isnull(t.splt_id,b.splt_id) = b.splt_id " & CRNL

'isnull will check if users only entered prop_id and left budg_yr or splt_id blank, if blank, it will substitue for all

'-----QUERY TO GET BUDGET PRCS
myBudgPRCs = "SELECT bpr.prop_id,bpr.budg_yr,bpr.splt_id, bpr.pgm_ref_code,id=identity(18), 0 as 'seq'" & CRNL _
& "INTO #myBudgPRCs FROM #myBSplit mbs,csd.budg_pgm_ref  bpr" & CRNL _
& "WHERE mbs.prop_id = bpr.prop_id  AND mbs.splt_id= bpr.splt_id AND mbs.budg_yr=bpr.budg_yr" & CRNL _
& "order by bpr.prop_id,bpr.budg_yr,bpr.splt_id, bpr.pgm_ref_code" & CRNL

myBudgPRCs = myBudgPRCs _
& "SELECT prop_id,budg_yr,splt_id, MIN(id) as 'start'" & CRNL _
& "INTO #mySt2 " & CRNL _
& "FROM #myBudgPRCS " & CRNL _
& "GROUP BY prop_id,budg_yr,splt_id" & CRNL _
& "UPDATE #myBudgPRCs set seq = id-M.start FROM #myBudgPRCs rb, #mySt2 M" & CRNL _
& "WHERE rb.prop_id = M.prop_id AND rb.budg_yr=M.budg_yr AND rb.splt_id= M.splt_id " & CRNL

'-----QUERY TO GET PROPOSAL PRCS
myPropPRCs = "SELECT p.prop_id, pa.prop_atr_code,id=identity(18), 0 as 'seq'" & CRNL _
& "INTO #myPropPRCs" & CRNL _
& "FROM (select distinct prop_id from #myBSplit mbs) as p, csd.prop_atr pa" & CRNL _
& "WHERE pa.prop_id = p.prop_id  AND pa.prop_atr_type_code = 'PRC'" & CRNL _
& "order by p.prop_id, pa.prop_atr_code" & CRNL _


myPropPRCs = myPropPRCs _
& "SELECT prop_id, MIN(id) as 'start'" & CRNL _
& "INTO #mySt3 " & CRNL _
& "FROM #myPropPRCs group by prop_id" & CRNL _
& "UPDATE #myPropPRCs set seq = id-M.start FROM #myPropPRCs r, #mySt3 M" & CRNL _
& "WHERE r.prop_id = M.prop_id " & CRNL

'-----QUERY TO GET PROPOSAL INFORMATION
myProp = "select p.prop_id,isnull(p.lead_prop_id,p.prop_id) AS lead_id," & CRNL _
& "CASE WHEN p.lead_prop_id IS NULL THEN 'I' WHEN p.lead_prop_id = p.prop_id THEN 'L' ELSE 'N' END AS ILN," & CRNL _
& "p.nsf_rcvd_date, p.dd_rcom_date, p.pgm_annc_id,p.org_code as Prop_Org_Code," & CRNL _
& "p.pgm_ele_code+' - '+pe.pgm_ele_name as PEC_prop," & CRNL _
& "p.pm_ibm_logn_id as Prop_Pm_ibm_logn_id," & CRNL _
& "p.obj_clas_code as Prop_Obj_Clas_Code," & CRNL _
& "ps.prop_stts_abbr,nr.natr_rqst_abbr," & CRNL _
& "pi.pi_last_name, pi.pi_frst_name," & CRNL _
& "i.inst_shrt_name as inst_name,p.rqst_dol,p.rcom_awd_istr,ai.awd_istr_txt, ai.awd_istr_abbr,ai.awd_istr_abbr as rcom_istr_abbr," & CRNL _
& "p.prop_titl_txt,og.dir_div_abbr,id=identity(18), 0 as 'seq'" & CRNL _
& "into #myProp" & CRNL _
& "FROM (select distinct prop_id from #myBSplit mbs) as prop" & CRNL _
& "JOIN csd.prop p ON p.prop_id = prop.prop_id" & CRNL _
& "JOIN csd.awd_istr ai on p.rcom_awd_istr = ai.awd_istr_code" & CRNL _
& "JOIN csd.pgm_ele  as pe ON pe.pgm_ele_code=p.pgm_ele_code" & CRNL _
& "JOIN csd.prop_stts ps on ps.prop_stts_code=p.prop_stts_code" & CRNL _
& "JOIN csd.natr_rqst nr on nr.natr_rqst_code = p.natr_rqst_code" & CRNL _
& "JOIN csd.pi_vw pi on p.pi_id=pi.pi_id" & CRNL _
& "JOIN csd.inst as i on i.inst_id=p.inst_id" & CRNL _
& "JOIN csd.org as og on og.org_code=p.org_code" & CRNL
'& "JOIN  #myBSplit mbs on mbs.prop_id=p.prop_id" & CRNL

'-----QUERY TO DISPLAY RESULTS
myResults = "SELECT distinct getdate() as run_date,mp.nsf_rcvd_date,mp.dd_rcom_date,mp.Prop_Org_Code," & CRNL _
& "b.org_code as Budg_Org_Code,mp.PEC_prop,mbs.PEC_bdg_splt," & CRNL _
& "mp.Prop_Pm_ibm_logn_id,b.pm_ibm_logn_id as Budg_Pm_ibm_logn_id, mp.Prop_Obj_Clas_Code,b.obj_clas_code as Budg_Obj_Clas_Code," & CRNL _
& "mp.pgm_annc_id,mp.prop_stts_abbr,mp.natr_rqst_abbr," & CRNL _
& " mp.dir_div_abbr,mbs.prop_id, mp.ILN,mp.lead_id,b.awd_id, mp.prop_titl_txt," & CRNL _
& "mbs.splt_id,mbs.budg_yr,b.budg_splt_tot_dol,mp.rqst_dol," & CRNL _
& "(SELECT MAX( CASE pa.seq WHEN 0 THEN rtrim(pa.prop_atr_code) END)+" & CRNL _
& "MAX( CASE pa.seq WHEN 1 THEN ' '+rtrim(pa.prop_atr_code) END)+" & CRNL _
& "MAX( CASE pa.seq WHEN 2 THEN ' '+rtrim(pa.prop_atr_code) END)+" & CRNL _
& "MAX( CASE pa.seq WHEN 3 THEN ' '+rtrim(pa.prop_atr_code) END)+" & CRNL _
& "MAX( CASE pa.seq WHEN 4 THEN ' '+rtrim(pa.prop_atr_code) END)+" & CRNL _
& "MAX( CASE pa.seq WHEN 5 THEN ' '+rtrim(pa.prop_atr_code) END)+" & CRNL _
& "MAX( CASE pa.seq WHEN 6 THEN ' '+rtrim(pa.prop_atr_code) END)+" & CRNL _
& "MAX( CASE pa.seq WHEN 7 THEN ' '+rtrim(pa.prop_atr_code) END)+" & CRNL _
& "MAX( CASE pa.seq WHEN 8 THEN ' '+rtrim(pa.prop_atr_code) END)+" & CRNL _
& "MAX( CASE pa.seq WHEN 9 THEN ' '+rtrim(pa.prop_atr_code) END)+" & CRNL _
& "MAX( CASE pa.seq WHEN 10 THEN ' '+rtrim(pa.prop_atr_code) END)" & CRNL _
& "FROM #myPropPRCs pa WHERE pa.prop_id = mbs.prop_id) AS 'Prop PRCs'," & CRNL
myResults = myResults _
& "(SELECT MAX( CASE bp.seq WHEN 0 THEN rtrim(bp.pgm_ref_code) END)+" & CRNL _
& "MAX( CASE bp.seq WHEN 1 THEN ' ' + rtrim(bp.pgm_ref_code) END)+" & CRNL _
& "MAX( CASE bp.seq WHEN 2 THEN ' ' + rtrim(bp.pgm_ref_code) END)+" & CRNL _
& "MAX( CASE bp.seq WHEN 3 THEN ' ' + rtrim(bp.pgm_ref_code) END)+" & CRNL _
& "MAX( CASE bp.seq WHEN 4 THEN ' ' + rtrim(bp.pgm_ref_code) END)+" & CRNL _
& "MAX( CASE bp.seq WHEN 5 THEN ' ' + rtrim(bp.pgm_ref_code) END)+" & CRNL _
& "MAX( CASE bp.seq WHEN 6 THEN ' ' + rtrim(bp.pgm_ref_code) END)+" & CRNL _
& "MAX( CASE bp.seq WHEN 7 THEN ' ' + rtrim(bp.pgm_ref_code) END)+" & CRNL _
& "MAX( CASE bp.seq WHEN 8 THEN ' ' + rtrim(bp.pgm_ref_code) END)+" & CRNL _
& "MAX( CASE bp.seq WHEN 9 THEN ' ' + rtrim(bp.pgm_ref_code) END)+" & CRNL _
& "MAX( CASE bp.seq WHEN 10 THEN ' ' + rtrim(bp.pgm_ref_code) END)" & CRNL _
& "FROM #myBudgPRCs bp" & CRNL _
& "WHERE bp.prop_id=mbs.prop_id and bp.budg_yr=mbs.budg_yr and bp.splt_id = mbs.splt_id group by bp.prop_id, bp.budg_yr, bp.splt_id) AS 'Budg PRCs'," & CRNL _
& "mp.pi_last_name, mp.pi_frst_name,mp.inst_name,b.last_updt_user, b.last_updt_tmsp,ast.awd_stts_abbr, ast.awd_stts_txt" & CRNL _
& ", cs.cgi_stts_txt,mp.awd_istr_txt, mp.awd_istr_abbr,mp.rcom_istr_abbr from #myBSplit mbs" & CRNL _
& "JOIN csd.budg_splt b on mbs.prop_id=b.prop_id and mbs.budg_yr=b.budg_yr and mbs.splt_id=b.splt_id" & CRNL _
& "LEFT JOIN csd.awd a on b.awd_id=a.awd_id" & CRNL _
& "LEFT JOIN csd.awd_stts ast on a.awd_stts_code=ast.awd_stts_code" & CRNL _
& "LEFT JOIN csd.cgi c on  b.awd_id=c.awd_id and b.budg_yr= c.cgi_yr" & CRNL _
& "LEFT JOIN csd.cgi_stts cs on c.cgi_stts_code=cs.cgi_stts_code" & CRNL _
& "JOIN #myProp mp on mp.prop_id=mbs.prop_id" & CRNL


    Dim QT As QueryTable
    Dim LO As ListObject
    For Each LO In ActiveSheet.ListObjects
      If (Left(LO.Name, 15) = "SplitQueryTable") Then Set QT = LO.QueryTable 'excel adds a number; we ignore.
    Next
    
    With QT
       .CommandText = budg_spltsInclude & budg_spltsExclude & myBSplit & myBudgPRCs & myPropPRCs & myProp & myResults
       .Refresh (False)
    End With
    Call RefreshPivotTables(ActiveSheet, QT)
End Sub




