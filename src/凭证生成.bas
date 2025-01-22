Attribute VB_Name = "凭证生成"

Public Sub FKHZXJ()   '''''''''付款汇总--现金
On Error Resume Next

Adodc8.RecordSource = "SELECT * FROM CLFKPZ"
Adodc8.Refresh
Adodc7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLFKPZ WHERE INSTR(凭证号,'2-')>0 AND 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "2-1"
If Not Adodc7.Recordset.EOF Then
PZH = "2-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
Adodc9.RecordSource = "SELECT * FROM JGZCX1 where val(本期现付款)>0"
Adodc9.Refresh
If Adodc9.Recordset.EOF Then
Exit Sub
Else
Adodc9.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc9.Recordset.EOF
For i = 1 To 5
Adodc8.Recordset.AddNew
Adodc8.Recordset.Fields(0) = "购材料"     '''''摘要
If InStr(Adodc9.Recordset.Fields(4), "-") > 0 Then
Adodc8.Recordset.Fields(1) = Left(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") - 1)      ''''借方总账科目
Adodc8.Recordset.Fields(2) = Mid(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") + 1)         ''''借方明细科目
Else
Adodc8.Recordset.Fields(1) = Adodc9.Recordset.Fields(4)
End If
If InStr(Adodc9.Recordset.Fields(0), "-") > 0 Then
Adodc8.Recordset.Fields(3) = Left(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") - 1)     '''''''''''贷方总账科目
Adodc8.Recordset.Fields(4) = Mid(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") + 1)       '''贷方明细科目
Else
Adodc8.Recordset.Fields(3) = Adodc9.Recordset.Fields(0)
End If
Adodc8.Recordset.Fields(5) = Adodc9.Recordset.Fields(6)     ''''''''发生金额
Adodc8.Recordset.Fields(6) = PZH                           '''''''''''''凭证号
Adodc8.Recordset.Fields(7) = DTPicker3.Value               '''''''操作日期
Adodc8.Recordset.Fields(8) = Adodc9.Recordset.Fields(2)
Adodc8.Recordset.Fields(9) = ""                           ''''''''''记账
Adodc8.Recordset.Fields(10) = ""                          ''''''''''''复核
Adodc8.Recordset.Fields(11) = adodcCombo3.Text       ''''''''''''''制单
Adodc8.Recordset.Update
Adodc9.Recordset.Edit    '''''''''''''''''''''''''''''''
Adodc9.Recordset.Fields(12) = "已"
Adodc9.Recordset.Update    ''''''''''''''''''''''''''''''
Adodc9.Recordset.MoveNext
If Adodc9.Recordset.EOF Then
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "现金付款凭证")
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "现金付款单"
Adodc2.Recordset.Fields(3) = "付款凭证"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
Exit Sub
End If
Next
Adodc7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLFKPZ WHERE INSTR(凭证号,'2-')>0 AND 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "2-1"
If Not Adodc7.Recordset.EOF Then
PZH = "2-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "现金付款凭证")
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "现金付款单"
Adodc2.Recordset.Fields(3) = "付款凭证"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
End If
End Sub
''''''''''''''
Public Sub FKHZYH()   '''''''''付款汇总---银行存款
On Error Resume Next
Adodc7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLFKPZ WHERE INSTR(凭证号,'4-')>0 AND 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "4-1"
If Not Adodc7.Recordset.EOF Then
PZH = "4-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
Adodc9.RecordSource = "SELECT * FROM TZJZMX WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) AND (审核确认<>'已' OR 审核确认=NULL) AND 贷方金额<>'0' AND 借方金额='0' AND INSTR(类别,'银行存款')>0 AND (审核<>'已' OR 审核=NULL)"
Adodc9.Refresh
If Adodc9.Recordset.EOF Then
Exit Sub
Else
Adodc9.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc9.Recordset.EOF
For i = 1 To 5
Adodc8.Recordset.AddNew
Adodc8.Recordset.Fields(0) = Adodc9.Recordset.Fields(3)
If InStr(Adodc9.Recordset.Fields(4), "-") > 0 Then
Adodc8.Recordset.Fields(1) = Left(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") - 1)
Adodc8.Recordset.Fields(2) = Mid(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") + 1)
Else
Adodc8.Recordset.Fields(1) = Adodc9.Recordset.Fields(4)
End If
If InStr(Adodc9.Recordset.Fields(0), "-") > 0 Then
Adodc8.Recordset.Fields(3) = Left(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") - 1)
Adodc8.Recordset.Fields(4) = Mid(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") + 1)
Else
Adodc8.Recordset.Fields(3) = Adodc9.Recordset.Fields(0)
End If
Adodc8.Recordset.Fields(5) = Adodc9.Recordset.Fields(6)
Adodc8.Recordset.Fields(6) = PZH
Adodc8.Recordset.Fields(7) = DTPicker3.Value
Adodc8.Recordset.Fields(8) = Adodc9.Recordset.Fields(2)
Adodc8.Recordset.Fields(9) = ""
Adodc8.Recordset.Fields(10) = ""
Adodc8.Recordset.Fields(11) = adodcCombo3.Text
Adodc8.Recordset.Update
Adodc9.Recordset.Edit    '''''''''''''''''''''''''''''''
Adodc9.Recordset.Fields(12) = "已"
Adodc9.Recordset.Update    ''''''''''''''''''''''''''''''
Adodc9.Recordset.MoveNext
If Adodc9.Recordset.EOF Then
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "银行存款付款凭证")
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "银行付款单"
Adodc2.Recordset.Fields(3) = "付款凭证"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
Exit Sub
End If
Next
Adodc7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLFKPZ WHERE INSTR(凭证号,'4-')>0 AND 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "4-1"
If Not Adodc7.Recordset.EOF Then
PZH = "4-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "银行付款单"
Adodc2.Recordset.Fields(3) = "付款凭证"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "银行存款付款凭证")
End If
End Sub

Public Sub SKHZXJ()    ''''''''收款汇总----现金
On Error Resume Next
If adodcCombo3.Text = "" Then
MsgBox ("请选择复核员")
Exit Sub
End If
Adodc8.RecordSource = "SELECT * FROM CLSKPZ"
Adodc8.Refresh
Adodc7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLSKPZ WHERE INSTR(凭证号,'1-')>0 AND 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "1-1"
If Not Adodc7.Recordset.EOF Then
PZH = "1-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
Adodc9.RecordSource = "SELECT * FROM TZJZMX WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) AND (审核确认<>'已' OR 审核确认=NULL) AND 借方金额<>'0' AND 贷方金额='0' AND INSTR(类别,'现金')>0 AND (审核<>'已' OR 审核=NULL)"
Adodc9.Refresh
If Adodc9.Recordset.EOF Then
Exit Sub
Else
Adodc9.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc9.Recordset.EOF
For i = 1 To 5
Adodc8.Recordset.AddNew
Adodc8.Recordset.Fields(0) = Adodc9.Recordset.Fields(3)
If InStr(Adodc9.Recordset.Fields(0), "-") > 0 Then
Adodc8.Recordset.Fields(1) = Left(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") - 1)
Adodc8.Recordset.Fields(2) = Mid(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") + 1)
Else
Adodc8.Recordset.Fields(1) = Adodc9.Recordset.Fields(0)
End If
If InStr(Adodc9.Recordset.Fields(4), "-") > 0 Then
Adodc8.Recordset.Fields(3) = Left(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") - 1)
Adodc8.Recordset.Fields(4) = Mid(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") + 1)
Else
Adodc8.Recordset.Fields(3) = Adodc9.Recordset.Fields(4)
End If
Adodc8.Recordset.Fields(5) = Adodc9.Recordset.Fields(5)
Adodc8.Recordset.Fields(6) = PZH
Adodc8.Recordset.Fields(7) = DTPicker3.Value
Adodc8.Recordset.Fields(8) = Adodc9.Recordset.Fields(2)
Adodc8.Recordset.Fields(9) = ""
Adodc8.Recordset.Fields(10) = ""
Adodc8.Recordset.Fields(11) = adodcCombo3.Text
Adodc8.Recordset.Update
Adodc9.Recordset.Edit    '''''''''''''''''''''''''''''''
Adodc9.Recordset.Fields(12) = "已"
Adodc9.Recordset.Update    ''''''''''''''''''''''''''''''
Adodc9.Recordset.MoveNext
If Adodc9.Recordset.EOF Then
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "现金收款凭证")
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "现金收款单"
Adodc2.Recordset.Fields(3) = "收款凭证"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
Exit Sub
End If
Next
Adodc7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLSKPZ WHERE INSTR(凭证号,'1-')>0 AND 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "1-1"
If Not Adodc7.Recordset.EOF Then
PZH = "1-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "现金收款单"
Adodc2.Recordset.Fields(3) = "收款凭证"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "现金收款凭证")
End If
End Sub

Public Sub SKHZYH()    ''''''''收款汇总----银行存款
On Error Resume Next
Adodc7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLSKPZ WHERE INSTR(凭证号,'3-')>0 AND 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "3-1"
If Not Adodc7.Recordset.EOF Then
PZH = "3-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
Adodc9.RecordSource = "SELECT * FROM TZJZMX WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) AND (审核确认<>'已' OR 审核确认=NULL) AND 借方金额<>'0' AND 贷方金额='0' AND INSTR(类别,'银行存款')>0 AND (审核<>'已' OR 审核=NULL)"
Adodc9.Refresh
If Adodc9.Recordset.EOF Then
Exit Sub
Else
Adodc9.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc9.Recordset.EOF
For i = 1 To 5
Adodc8.Recordset.AddNew
Adodc8.Recordset.Fields(0) = Adodc9.Recordset.Fields(3)
If InStr(Adodc9.Recordset.Fields(0), "-") > 0 Then
Adodc8.Recordset.Fields(1) = Left(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") - 1)
Adodc8.Recordset.Fields(2) = Mid(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") + 1)
Else
Adodc8.Recordset.Fields(1) = Adodc9.Recordset.Fields(0)
End If
If InStr(Adodc9.Recordset.Fields(4), "-") > 0 Then
Adodc8.Recordset.Fields(3) = Left(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") - 1)
Adodc8.Recordset.Fields(4) = Mid(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") + 1)
Else
Adodc8.Recordset.Fields(3) = Adodc9.Recordset.Fields(4)
End If
Adodc8.Recordset.Fields(5) = Adodc9.Recordset.Fields(5)
Adodc8.Recordset.Fields(6) = PZH
Adodc8.Recordset.Fields(7) = DTPicker3.Value
Adodc8.Recordset.Fields(8) = Adodc9.Recordset.Fields(2)
Adodc8.Recordset.Fields(9) = ""
Adodc8.Recordset.Fields(10) = ""
Adodc8.Recordset.Fields(11) = adodcCombo3.Text
Adodc8.Recordset.Update
Adodc9.Recordset.Edit    '''''''''''''''''''''''''''''''
Adodc9.Recordset.Fields(12) = "已"
Adodc9.Recordset.Update    ''''''''''''''''''''''''''''''
Adodc9.Recordset.MoveNext
If Adodc9.Recordset.EOF Then
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "银行存款收款凭证")
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "银行收款单"
Adodc2.Recordset.Fields(3) = "收款凭证"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
Exit Sub
End If
Next
Adodc7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLSKPZ WHERE INSTR(凭证号,'3-')>0 AND 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "3-1"
If Not Adodc7.Recordset.EOF Then
PZH = "3-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "银行收款单"
Adodc2.Recordset.Fields(3) = "收款凭证"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "银行存款收款凭证")
End If
End Sub




Public Sub CLCKpz()
On Error Resume Next
Adodc8.adodcbase.Execute "DELETE * FROM CLSCHZ"
lo = "d:\数据库\bfrz\" + ljb + "\cw.mdb"       '''''''''''''''''''''''经典
Adodc3.adodcbase.Execute "INSERT INTO CLSCHZ(金额) IN'" & lo & "' SELECT FORMAT(SUM(合计金额),'#0.00') AS 金额 FROM KPD WHERE (审核<>'已' OR 审核=NULL) AND 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc8.adodcbase.Execute "DELETE * FROM CLSCHZ WHERE 金额=NULL"
Adodc9.RecordSource = "SELECT * FROM CLSCHZ"
Adodc9.Refresh
Adodc8.RecordSource = "SELECT * FROM CLSCCB"
Adodc8.Refresh
Adodc7.RecordSource = "SELECT MAX(VAL(MID(CLSCCB.凭证号,3))) FROM CLSCCB WHERE CLSCCB.日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "S-1"
If Not Adodc7.Recordset.EOF Then
PZH = "S-" + Trim(Adodc7.Recordset.Fields(0) + 1)
Else
PZH = "S-1"
End If
If Adodc9.Recordset.EOF Then
Exit Sub
Else
Adodc9.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc9.Recordset.EOF
For i = 1 To 5
Adodc8.Recordset.AddNew
Adodc8.Recordset.Fields(0) = "耗用原材料"
Adodc8.Recordset.Fields(1) = "生产成本"
Adodc8.Recordset.Fields(2) = "直接生产成本"
Adodc8.Recordset.Fields(3) = "库存物资"
Adodc8.Recordset.Fields(4) = ""
Adodc8.Recordset.Fields(5) = Adodc9.Recordset.Fields(4)
Adodc8.Recordset.Fields(6) = PZH
Adodc8.Recordset.Fields(7) = DTPicker3.Value
Adodc8.Recordset.Fields(8) = ""
Adodc8.Recordset.Fields(9) = ""
Adodc8.Recordset.Fields(10) = ""
Adodc8.Recordset.Fields(11) = adodcCombo3.Text
Adodc8.Recordset.Update
Adodc9.Recordset.MoveNext
If Adodc9.Recordset.EOF Then
Adodc4.adodcbase.Execute "UPDATE KPD SET 审核='已' WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "成本凭证")
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "材料出库单"
Adodc2.Recordset.Fields(3) = "成本凭证"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
Exit Sub
End If
Next
KLLLL = KLLLL + 1
Adodc7.RecordSource = "SELECT MAX(VAL(MID(CLSCCB.凭证号,3))) FROM CLSCCB WHERE CLSCCB.日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "S-1"
If Adodc7.Recordset.EOF Then
PZH = "S-1"
Else
PZH = "S-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
Loop
Adodc4.adodcbase.Execute "UPDATE KPD SET 审核='已' WHERE 日期 BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "成本凭证")
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "材料出库单"
Adodc2.Recordset.Fields(3) = "成本凭证"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
End If
End Sub

