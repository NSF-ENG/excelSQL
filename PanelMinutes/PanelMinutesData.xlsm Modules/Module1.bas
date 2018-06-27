Attribute VB_Name = "Module1"
Sub MinutesMerge()

Dim CRNL As String
CRNL = Chr(13) & Chr(10)
Dim ids As String
Dim nrows As Long
nrows = Range("PanelTable[panl_id]").Rows.Count
If nrows < 2 Then
   If nrows < 1 Then
     MsgBox "No panels listed under panl_id column; aborting"
     Exit Sub
   Else
     ids = Range("PanelTable[panl_id]").Value
   End If
Else ' we have at least two
  ids = Join(Application.Transpose(Range("PanelTable[panl_id]").Value), "','")
End If
' strip spaces (visible and invisible) and ,'' from string.
ids = "(panl.panl_id In " & Replace(Replace(Replace("('" & ids & "'))", " ", ""), Chr(63), ""), ",''", "") & CRNL

'MsgBox ids

queryText = "SELECT panl.panl_id, panl.panl_name, panl.panl_bgn_date, panl.panl_end_date, panl.panl_loc, panl.org_code, panl.pgm_ele_code, panl.pm_logn_id," & _
        "panl.meet_type_code, panl.meet_fmt, panl.fund_org_code, panl.fund_pgm_ele_code, panl.fund_app_code, panl_stts.panl_stts_txt, panl.oblg_flag, org.org_long_name, pgm_ele.pgm_ele_long_name," & CRNL & _
        "(SELECT COUNT(panl_prop.prop_id) FROM csd.panl_prop panl_prop, csd.prop prop WHERE panl.panl_id = panl_prop.panl_id AND panl_prop.prop_id = prop.prop_id AND (prop.prop_id = isnull( prop.lead_prop_id, prop.prop_id) ) ) AS 'Nproj'," & CRNL & _
        "(SELECT COUNT(panl_prop.prop_id) FROM csd.panl_prop panl_prop WHERE panl.panl_id = panl_prop.panl_id) AS 'Nprop'," & CRNL & _
        "nullif((SELECT COUNT(panl_revr.revr_id) FROM csd.panl_revr panl_revr WHERE panl.panl_id = panl_revr.panl_id AND (panl_revr.tele_conf_part_flag = 'N')),0) AS 'Nrevr'," & CRNL & _
        "nullif((SELECT COUNT(panl_revr.revr_id) FROM csd.panl_revr panl_revr WHERE panl.panl_id = panl_revr.panl_id AND (panl_revr.tele_conf_part_flag = 'Y')),0) AS 'Nvirt_revr'" & CRNL & _
        "FROM csd.org org, csd.panl panl, csd.panl_stts panl_stts, csd.pgm_ele pgm_ele" & CRNL & _
        "WHERE panl.panl_stts_code = panl_stts.panl_stts_code AND panl.pgm_ele_code = pgm_ele.pgm_ele_code AND panl.org_code = org.org_code AND " & CRNL & _
        ids & CRNL

'MsgBox queryText
     
    Dim QT As QueryTable
    Set QT = Worksheets("MinutesMerge").ListObjects.Item(1).QueryTable
    With QT
        .CommandText = queryText
        .Refresh
    End With
End Sub
