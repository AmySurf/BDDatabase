# BDDatabase
将大部门内多个小组产出的数据一起导入到Access数据库中，通过Excel操作取用Access数据。

Private SourceFile, BackupFile, RestoreFile As String

Private Sub CB_Backup_Click()
'备份文件
    SourceFile = DBPath & mydb & ".accdb"
    BackupFile = BackUpPath & "0 ChinaRealtyDB\Daily Backup\" & GetUserID() & "_" & mydb & "_" & Format(Date, "yyyy-mm-dd") & ".accdb"
  
        If Dir$(BackupFile) <> "" Then
            Kill BackupFile
            FileCopy SourceFile, BackupFile
        Else
            FileCopy SourceFile, BackupFile
        End If
        
        Set myConn = New ADODB.Connection
        With myConn
            .Mode = adModeShareExclusive
            .Provider = "Microsoft.ACE.OLEDB.12.0"
            .Properties("Jet OLEDB:Database password") = TB_newpass.Text
            .Open "Data Source=" & BackupFile & ";"
            .Execute "ALTER DATABASE PASSWORD NULL " & "[" & TB_newpass.Text & "];"
        End With
        myConn.Close: Set myConn = Nothing
        Unload Form_Admin

End Sub

Private Sub CB_Compact_Click()
'压缩备份文件
    SourceFile = DBPath & mydb & ".accdb"
    BackupFile = BackUpPath & "0 ChinaRealtyDB\Daily Backup\" & GetUserID() & "_" & mydb & "_" & Format(Date, "yyyy-mm-dd") & ".accdb"
  
        If Dir$(BackupFile) <> "" Then
            Kill BackupFile
            DBEngine.CompactDatabase SourceFile, BackupFile, , , ";pwd=" & TB_newpass.Text
        Else
            DBEngine.CompactDatabase SourceFile, BackupFile, , , ";pwd=" & TB_newpass.Text
        End If

        Unload Form_Admin
End Sub


Private Sub CB_MorningReport_Click()

  Call RetriveAuthority(mydb)

End Sub

Private Sub CB_update_Click()
'修改密码
    Dim myOldPassword As String
        If TB_oldpass.Text = "" Then
          myOldPassword = "NULL"
        Else
          myOldPassword = "[" & TB_oldpass.Text & "]"
        End If
    Dim myNewPassword As String
        If TB_newpass.Text = "" Then
        myNewPassword = "NULL"
        Else
        myNewPassword = "[" & TB_newpass.Text & "]"
        End If
    Dim StrCommand As String
    StrCommand = "ALTER DATABASE PASSWORD " & myNewPassword & " " & myOldPassword & ";"
    
    Set myConn = New ADODB.Connection
    With myConn
        .Mode = adModeShareExclusive
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        If myOldPassword <> "NULL" Then .Properties("Jet OLEDB:Database password") = myOldPassword
        .Open "Data Source=" & DBPath & mydb & ".accdb;"
        .Execute StrCommand
    End With
    myConn.Close: Set myConn = Nothing
    MsgBox "修改成功"
        
        Unload Form_Admin
End Sub



Private Sub OBM_CR_Click()
    mydb = "Database"
End Sub

Private Sub OBM_DS_Click()
    mydb = "DailySales"
End Sub


Private Sub UserForm_Load()

    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
        
End Sub

Private Sub LB_Tablelist_click()
    If LB_Tablelist.Text <> "0User_Information" Then
        CB_MorningReport.Enabled = False
    Else
        CB_MorningReport.Enabled = True
    End If
    CBYes.Enabled = True
    CBExport.Enabled = True
End Sub

Private Sub CBNo_Click()
    End
End Sub

Private Sub CBExport_Click()
    Form_Admin.Hide
    Call Export_table(mydb, "[" & LB_Tablelist.Text & "]")
    Call UserMonitor(mydb, LB_Tablelist.Text, "View")
'    Unload Form_Admin
End Sub


Private Sub OBM1_Click()
    TB_StoreNo.Visible = False
    TB_FiledName.Visible = False
    CB_FiledType.Visible = False
End Sub

Private Sub OBM2_Click()
    TB_StoreNo.Visible = False
    TB_FiledName.Visible = False
    CB_FiledType.Visible = False
End Sub


Private Sub OBM4_Click()
    TB_StoreNo.Visible = True
    TB_FiledName.Visible = False
    CB_FiledType.Visible = False
End Sub
Private Sub OBM5_Click()
    TB_StoreNo.Visible = False
    TB_FiledName.Visible = True
    CB_FiledType.Visible = True
End Sub

Private Sub CBYes_Click()

    Form_Admin.Hide
    If OBM1.Value = True Then Call Maintenance_table(mydb, "[" & LB_Tablelist.Text & "]", Left(OBM1.Caption, 4))
    If OBM2.Value = True Then Call Maintenance_table(mydb, "[" & LB_Tablelist.Text & "]", Left(OBM2.Caption, 4))
    If OBM4.Value = True And TB_StoreNo.Text <> Empty Then
      Call Maintenance_Excute(mydb, "[" & LB_Tablelist.Text & "]", Left(OBM4.Caption, 4), " WHERE [User_ID]='" & TB_StoreNo.Text & "'")
    End If
    If OBM5.Value = True And Len(TB_FiledName.Text) > 5 And CB_FiledType.Text <> Empty Then
      Call Maintenance_Excute(mydb, "[" & LB_Tablelist.Text & "]", Left(OBM5.Caption, 4), "[" & TB_FiledName.Text & "] " & CB_FiledType.Text & ";")
    End If
'    Unload Form_Admin

End Sub


Private Sub UserForm_Load()

    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
        
End Sub
Private Sub OBM_DS_Click()

    OBM_CR.Enabled = False
    mydb = "DailySales"
    Call getTableList(mydb, "Read/Write")

End Sub

Private Sub CB_MorningReport_Click()
  
  Call RetriveExcelData(mydb, "[" & LB_Tablelist.Text & "]")
  Call UserMonitor(mydb, LB_Tablelist.Text, "Edit")
  
End Sub

Private Sub LB_Tablelist_click()
    CBYes.Enabled = True
    CBExport.Enabled = True
    If LB_Tablelist.Text = "Working Timeline" Or LB_Tablelist.Text = "O Operating Store Performance_Monthly" Then
        OBM6.Visible = True
        TB_strDelete.Visible = True
        OBM3.Enabled = True
        OBM2.Enabled = True
    End If
    If LB_Tablelist.Text = "I Budget_RMB" Or LB_Tablelist.Text = "G Homepage" Or LB_Tablelist.Text = "E Operating Store_Area List" _
             Or LB_Tablelist.Text = "D Operating Store_Contract Information" Or LB_Tablelist.Text = "P Daily Sales" Then
        CB_MorningReport.Enabled = True
    Else
        CB_MorningReport.Enabled = False
    End If
End Sub

Private Sub CBNo_Click()
    End
End Sub

Private Sub OBM1_Click()
    TB_StoreNo.Visible = False
    TB_FiledName.Visible = False
    CB_FiledType.Visible = False
End Sub

Private Sub OBM2_Click()
    TB_StoreNo.Visible = False
    TB_FiledName.Visible = False
    CB_FiledType.Visible = False
End Sub

Private Sub OBM3_Click()
    TB_StoreNo.Visible = False
    TB_FiledName.Visible = False
    CB_FiledType.Visible = False
End Sub

Private Sub OBM4_Click()
    TB_StoreNo.Visible = True
    TB_FiledName.Visible = False
    CB_FiledType.Visible = False
End Sub
Private Sub OBM5_Click()
    TB_StoreNo.Visible = False
    TB_FiledName.Visible = True
    CB_FiledType.Visible = True
End Sub
Private Sub CBExport_Click()

    Form_DB.Hide
    Call Export_table(mydb, "[" & LB_Tablelist.Text & "]")
    Call UserMonitor(mydb, LB_Tablelist.Text, "View")
    Unload Form_DB
    
End Sub

Private Sub CBYes_Click()

    Form_DB.Hide
    If OBM1.Value = True Then Call Maintenance_table(mydb, "[" & LB_Tablelist.Text & "]", Left(OBM1.Caption, 4)): Call UserMonitor(mydb, LB_Tablelist.Text, "Edit")
    If OBM2.Value = True Then Call Maintenance_table(mydb, "[" & LB_Tablelist.Text & "]", Left(OBM2.Caption, 4)): Call UserMonitor(mydb, LB_Tablelist.Text, "Edit")
    If OBM3.Value = True Then Call Maintenance_table(mydb, "[" & LB_Tablelist.Text & "]", Left(OBM3.Caption, 4)): Call UserMonitor(mydb, LB_Tablelist.Text, "Edit")
    If OBM4.Value = True And Val(TB_StoreNo.Text) <> Empty Then
        If Val(InputBox$("请确认删除店号", "输入店号")) = Val(TB_StoreNo.Text) Then
            Call Maintenance_Excute(mydb, "[" & LB_Tablelist.Text & "]", Left(OBM4.Caption, 4), " WHERE [Store No]=" & Val(TB_StoreNo.Text) & ";")
            Call UserMonitor(mydb, LB_Tablelist.Text, "Edit")
        Else
            MsgBox "店号错误，操作失败"
            End
        End If
    End If
    If OBM5.Value = True And Len(TB_FiledName.Text) > 5 And CB_FiledType.Text <> Empty Then
        If InputBox$("请确认字段名称", "输入字段名") = TB_FiledName.Text Then
            Call Maintenance_Excute(mydb, "[" & LB_Tablelist.Text & "]", Left(OBM5.Caption, 4), "[" & TB_FiledName.Text & "] " & CB_FiledType.Text & ";")
            Call UserMonitor(mydb, LB_Tablelist.Text, "Edit")
        Else
            MsgBox "字段名错误，操作失败"
            End
        End If
    End If
    If OBM6.Value = True Then
            Call Maintenance_Excute(mydb, "[" & LB_Tablelist.Text & "]", Left(OBM6.Caption, 4), TB_strDelete.Text & ";")
            Call UserMonitor(mydb, LB_Tablelist.Text, "Edit")
    End If
'    Unload Form_DB

End Sub


Option Base 1

Private Sub UserForm_Load()

    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
        
End Sub





Private Sub OBM_DS_Click()

    OBM_CR.Enabled = False
    mydb = "DailySales"
    Call getTableList(mydb, "Read")

End Sub

'
'Comment Report界面
'
Private Sub LB_comment_Click()
                             
    If LB_comment.Text = "Provincial Store Pipeline & Performance" Then
        Label_province.Visible = True
        LB_province.Visible = True
    Else
        Label_province.Visible = False
        LB_province.Visible = False
    End If
    If Left(LB_comment.Text, 9) = "Approved " Or Left(LB_comment.Text, 9) = "Import Fo" Then
        Label_SN.Visible = True:        TB_StoreNo.Visible = True
        Label_EX.Visible = False:        TB_EXR.Visible = False
    Else
        If Left(LB_comment.Text, 27) = "Operating Store Performance" Then
            Label_SN.Visible = True:        TB_StoreNo.Visible = True
            Label_EX.Visible = True:        TB_EXR.Visible = True
        Else
            Label_SN.Visible = False:        TB_StoreNo.Visible = False
            Label_EX.Visible = True:        TB_EXR.Visible = True
        End If
    End If
    CBgetC.Enabled = True
End Sub

Private Sub CBgetC_Click()
'人机交互，获取报告所需信息

    Form_RP.Hide
    Dim myDate As String
    If mydb = "ChinaRealtyDB" Then
        myDate = "O Operating Store Comment Performance"
     Else
        myDate = "C Operating Store Performance"
     End If
        myDate = GetYearMonth(mydb, LB_comment.Text, "View", "Edit_Time_" & myDate) '错误待解决
        
    Call Comment_Report(mydb, Val(TB_EXR.Value), LB_comment.Text, myDate, LB_province.Text, Val(TB_StoreNo.Value))
    Unload Form_RP
End Sub

Private Sub CBexitC_Click()
    End
End Sub


'
'Personal Report界面
'
Private Sub OBM1_Click()
    TB_EX.Visible = False
End Sub
Private Sub OBM2_Click()
    TB_EX.Visible = True
End Sub

Private Sub LB_table_DblClick(ByVal Cancel As MSForms.ReturnBoolean)    '获取表内所有字段名

LB_field.Clear
Set myConn = New ADODB.Connection
If ConnectDatabase(mydb, 1) Then
   Set myRs = New ADODB.Recordset
   myRs.Open "[" & LB_table.Text & "]", myConn, adUseClient, adLockReadOnly, adCmdTable
'   myRs.Open "[" & LB_table.Text & "]", myConn, adUseClient, adLockReadOnly, adCmdUnknown
   
   For i = 0 To myRs.Fields.Count - 1
     LB_field.AddItem
     LB_field.List(i, 0) = myRs.Fields(i).Name
     LB_field.List(i, 1) = LB_table.Text
   Next i
End If
myRs.Close: Set myRs = Nothing
myConn.Close: Set myConn = Nothing

End Sub

Private Sub LB_field_DblClick(ByVal Cancel As MSForms.ReturnBoolean)    '添加字段
  CBgetP.Enabled = True
  With LB_personal
    .AddItem
    .List(.ListCount - 1, 0) = LB_field.Text
    .List(.ListCount - 1, 1) = LB_field.Column(1)
  End With
  With CB_criteria
    .AddItem
    .List(.ListCount - 1, 0) = LB_field.Text
    .List(.ListCount - 1, 1) = LB_field.Column(1)
  End With
  
End Sub

Private Sub LB_personal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)    '移除字段
    LB_personal.RemoveItem LB_personal.ListIndex
End Sub

Private Sub CBgetP_Click()

Form_RP.Hide
Dim myStr As String: myStr = "SELECT DISTINCT "      '开始
Dim myObjS:    Set myObjS = CreateObject("Scripting.Dictionary")

With LB_personal
   If OBM2.Value = True And Val(TB_EX.Text) > 0 Then
        For i = 0 To .ListCount - 1
                If Right(.List(i, 0), 3) = "RMB" Then
                myStr = myStr & Left(.List(i, 1), 1) & ".[" & .List(i, 0) & "]/" & Val(TB_EX.Text) & " AS [" & Left(.List(i, 0), Len(.List(i, 0)) - 3) & "US], "
                Else
                    If Right(.List(i, 0), 3) = "SqM" Then
                    myStr = myStr & Left(.List(i, 1), 1) & ".[" & .List(i, 0) & "]*10.7639 AS [" & Left(.List(i, 0), Len(.List(i, 0)) - 3) & "Sqft], "
                    Else
                myStr = myStr & Left(.List(i, 1), 1) & ".[" & .List(i, 0) & "], "
                    End If
                End If
        myObjS(.List(i, 1)) = ""
        Next
   Else
        For i = 0 To .ListCount - 1
         myStr = myStr & Left(.List(i, 1), 1) & ".[" & .List(i, 0) & "], "    '取字段
         myObjS(.List(i, 1)) = ""
        Next
   End If
End With
  
myK = myObjS.Keys     '准备取表名
    myStr = Left(myStr, Len(myStr) - 2) & " FROM "
    Dim myStr1, myStr2  As String:    myStr1 = "[" & myK(0) & "] AS " & Left(myK(0), 1)
    Select Case UBound(myK)
     Case Is = 0
        myStr = myStr & myStr1
     Case Is < 5
        For i = 1 To UBound(myK)
           myStr2 = myStr2 & "("
        Next
        myStr = myStr & myStr2 & myStr1
        For i = 1 To UBound(myK)
           myStr = myStr & " INNER JOIN [" & myK(i) & "] AS " & Left(myK(i), 1) & " ON " _
                & Left(myK(0), 1) & ".[Store No] = " & Left(myK(i), 1) & ".[Store No])"
        Next
     Case Else
        MsgBox "选择表格过多，请重设"
        Exit Sub
    End Select
If Left(Trim(TB_criteria.Text), 1) = "#" Then 'Where条件
myStr = myStr & " WHERE (((" & Left(CB_criteria.List(CB_criteria.ListIndex, 1), 1) & ".[" & CB_criteria.Text & "])=" & Trim(TB_criteria.Text) & "))"
Else
    If Val(Trim(TB_criteria.Text)) Then
    myStr = myStr & " WHERE (((" & Left(CB_criteria.List(CB_criteria.ListIndex, 1), 1) & ".[" & CB_criteria.Text & "])=" & Val(Trim(TB_criteria.Text)) & "))"
    Else
    If Trim(TB_criteria.Text) <> "" Then myStr = myStr & " WHERE (((" & Left(CB_criteria.List(CB_criteria.ListIndex, 1), 1) & ".[" & CB_criteria.Text & "])='" & Trim(TB_criteria.Text) & "'))"
    End If
End If
myStr = myStr & ";"                      '结束
        
'########监控程序，关闭
'        myUserID = GetUserID()
'        Set myConn = New ADODB.Connection
'        If ConnectDatabase(mydb, 3) Then
'           Set myRs = New ADODB.Recordset
'           myRs.Open "0User_Monitor", myConn, adOpenKeyset, adLockOptimistic, adCmdTable
'
'              myRs.Find "[User_ID] Like '" & myUserID & "'", , adSearchForward, 1       '搜索用户名
'              On Error Resume Next
'              If myRs.EOF Then
'                myRs.AddNew
'                  myRs("User_ID") = myUserID
'                  For i = 0 To UBound(myK)
'                    If IsNull(myRs.Fields("View_Count_" & myK(i)).Value) Then
'                    myRs.Fields("View_Count_" & myK(i)).Value = 1
'                    Else
'                    myRs.Fields("View_Count_" & myK(i)) = myRs.Fields("View_Count_" & myK(i)).Value + 1
'                    End If
'                    myRs.Fields("View_Time_" & myK(i)) = Now
'                  Next
'                myRs.Update
'              Else
'                  For i = 0 To UBound(myK)
'                    If IsNull(myRs.Fields("View_Count_" & myK(i)).Value) Then
'                    myRs.Fields("View_Count_" & myK(i)).Value = 1
'                    Else
'                    myRs.Fields("View_Count_" & myK(i)) = myRs.Fields("View_Count_" & myK(i)).Value + 1
'                    End If
'                    myRs.Fields("View_Time_" & myK(i)) = Now
'                  Next
'                myRs.Update
'              End If
'           myRs.Close
'        End If
'        Set myRs = Nothing
'        myConn.Close: Set myConn = Nothing
'
Set myObjS = Nothing
Call Presonal_Report(mydb, myStr)
Unload Form_RP

End Sub

Private Sub CBreset_Click()
    LB_personal.Clear
End Sub


Private Sub CBexitP_Click()
    End
End Sub


'
'Package界面
'
Private Sub CBget_Click()

Form_RP.Hide
If Val(TB_StoreNo3.Text) <> 0 Then
    If OBM31.Value = True Then Call GetPackage("1. New Stores\" & OBM31.Caption & "\", TB_StoreNo3.Text)
    If OBM32.Value = True Then Call GetPackage("1. New Stores\" & OBM32.Caption & "\", TB_StoreNo3.Text, "Exce1")
    If OBM33.Value = True Then Call GetPackage("1. New Stores\" & OBM33.Caption & "\", TB_StoreNo3.Text, "Exce1")
    If OBM34.Value = True Then Call GetPackage("1. New Stores\" & OBM34.Caption & "\", TB_StoreNo3.Text)
    If OBM35.Value = True Then Call GetPackage("1. New Stores\" & OBM35.Caption & "\", TB_StoreNo3.Text)
    If OBM36.Value = True Then Call GetPackage("1. New Stores\" & OBM36.Caption & "\", TB_StoreNo3.Text, "Exce1")
Else
    MsgBox "请输入4位数店号"
End If

End Sub

Private Sub CBexit_Click()
    End
End Sub

Private Sub CB_AccessApply_Click()

    Form_RP.Hide
    Workbooks.Open FileName:=DBPath & "Version\Application Form.xlsm", ReadOnly:=True
    Unload Form_RP

End Sub





Option Base 1

Public myConn As ADODB.Connection
'Public Const DBPath As String = "\\cnnts8005fs\Public\BD\0 ChinaRealtyDB\"
'Public Const BackUpPath As String = "\\cnnts8005fs\Private\BD\PNSP\"
Public Const DBPath As String = "C:\0 ChinaRealtyDB\"
Public Const BackUpPath As String = "C:\"

Public myRs As ADODB.Recordset
Public mydb As String
Public myUserID As String



Sub F_Open()
'Excel操作菜单加载程序

    On Error Resume Next
    Application.CommandBars("ACS").Delete
    Application.CommandBars.Add(Name:="ACS", Position:=msoBarTop).Visible = True
                 
Dim NewButton
               
                 Set NewButton = CommandBars("BDMenu").Controls.Add(Type:=msoControlButton, Before:=1)
                  Set NewButton = CommandBars("ACS").Controls.Add(Type:=msoControlButton, Before:=1)
                 With NewButton
                     .FaceId = 51
                     .Visible = True
                     .Style = msoButtonIconAndCaption
                     .BeginGroup = True
                     .Caption = "数据查询"
                     .Width = 20
                     .OnAction = "Load_From_RP"
                 End With
                 
                 Set NewButton = CommandBars("ACS").Controls.Add(Type:=msoControlButton, Before:=2)
                 With NewButton
                     .FaceId = 26
                     .Style = msoButtonIconAndCaption
                     .BeginGroup = True
                     .Caption = "数据维护"
                     .Width = 20
'                     .Width = AutoFit
                     .OnAction = "Load_From_DB"
                 End With
                  
                 Set NewButton = CommandBars("ACS").Controls.Add(Type:=msoControlButton, Before:=3)
                 With NewButton
                     .FaceId = 35
                     .Style = msoButtonIconAndCaption
                     .BeginGroup = True
                     .Caption = "Admin权限"
                      .Width = 20
'                    .Width = AutoFit
                     .OnAction = "Load_From_Admin"
                 End With
'
'                 Set NEWBUTTON = CommandBars("ACS").Controls.Add(Type:=msoControlButton, Before:=4)
'                 With NEWBUTTON
'                     .FaceId = 41
'                     .Style = msoButtonIconAndCaption
'                     .BeginGroup = True
'                     .Caption = "DailySales"
'                     .Width = AutoFit
'                     .OnAction = "Load_UserForm_DailySales"
'                 End With
                 
End Sub



Private Sub Load_From_RP()
'Excel操作菜单第一按钮: 调用Form数据查询 数据初始化

Dim arrP(25, 2), myEP, myCP
    myEP = Array("Anhui", "Beijing", "Chongqing", "Fujian", "Guangdong", "Guangxi", "Guizhou", "Hebei", "Henan", "Heilongjiang", "Hubei", "Hunan", _
                         "Jilin", "Jiangsu", "Jiangxi", "Liaoning", "Shandong", "Shanxi", "Shaanxi", "Shanghai", "Sichuan", "Tianjin", "Yunnan", "Zhejiang", "Inner Mongolia")
    myCP = Array("安徽", "北京市", "重庆市", "福建", "广东", "广西", "贵州", "河北", "河南", "黑龙江", "湖北", "湖南", _
                         "吉林", "江苏", "江西", "辽宁", "山东", "山西", "陕西", "上海市", "四川", "天津市", "云南", "浙江", "内蒙古")
    For i = 1 To 25
    arrP(i, 1) = myEP(i)
    Next
    For i = 1 To 25
    arrP(i, 2) = myCP(i)
    Next
    
    Load Form_RP
    Form_RP.CBgetC.Enabled = False
    Form_RP.LB_province.List = arrP
    Form_RP.Label_province.Visible = False
    Form_RP.LB_province.Visible = False
    Form_RP.Label_SN.Visible = False
    Form_RP.TB_StoreNo.Visible = False
    
    Form_RP.CBgetP.Enabled = False
    Form_RP.TB_EX.Visible = False

mydb = "ChinaRealtyDB"
Call getTableList(mydb, "Read")     '调子程序确定用户权限范围信息
Form_RP.show

End Sub


Private Sub Load_From_DB()
'Excel操作菜单第二按钮：调用Form数据维护 数据初始化

    Load Form_DB
    Form_DB.CBYes.Enabled = False
    Form_DB.CBExport.Enabled = False
    Form_DB.OBM3.Enabled = False
    Form_DB.TB_StoreNo.Visible = False
    Form_DB.TB_FiledName.Visible = False
    Form_DB.CB_FiledType.List = Array("TEXT", "NUMERIC", "INT", "DATETIME")
    Form_DB.CB_FiledType.Visible = False
    Form_DB.OBM2.Enabled = False
    Form_DB.OBM6.Visible = False
    Form_DB.TB_strDelete.Visible = False
    Form_DB.CB_MorningReport.Enabled = False

mydb = "ChinaRealtyDB"
Call getTableList(mydb, "Read/Write")     '调子程序确定用户权限范围信息
Form_DB.show

End Sub


Private Sub Load_From_Admin()
'Excel操作菜单第三按钮：调用Form 管理权限 数据初始化

myUserID = GetUserID()

mydb = "ChinaRealtyDB"
Set myConn = New ADODB.Connection
If ConnectDatabase(mydb, 1) Then
   Set myRs = New ADODB.Recordset
   
    myRs.Open "0User_Information", myConn, adLockReadOnly, adCmdTable
    myRs.Find "[User_ID] Like '" & myUserID & "'", , adSearchForward, 1       '搜索用户名
     If Not myRs.EOF And myRs("Comment").Value = "Admin" Then
        Load Form_Admin
        Form_Admin.CBYes.Enabled = False
        Form_Admin.CBExport.Enabled = False
        Form_Admin.TB_StoreNo.Visible = False
        Form_Admin.TB_FiledName.Visible = False
        Form_Admin.CB_MorningReport.Enabled = False
        Form_Admin.CB_FiledType.List = Array("TEXT", "NUMERIC", "INT", "DATETIME")
        Form_Admin.CB_FiledType.Visible = False
        Form_Admin.LB_Tablelist.List = Array("0User_authority", "0Comments_Report", "0User_Information", "0User_Monitor")
     Else
        MsgBox "没有权限": End
     End If
    myRs.Close

End If
Set myRs = Nothing
myConn.Close: Set myConn = Nothing
Form_Admin.show

End Sub



'********************************************************************************************************************************************************

'子程序

'********************************************************************************************************************************************************


Public Function ConnectDatabase(ByVal mydb As String, ByVal myMode As Long) As Boolean
'子程序，连接Databass

    ConnectDatabase = False
    On Error GoTo amyhong
    With myConn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Jet OLEDB:Database password") = Form_Admin.TB_newpass.Text
        .Mode = myMode
        .Open "Data Source=" & DBPath & mydb & ".accdb;"
    End With
    Do
        DoEvents
    Loop Until myConn.State = adStateOpen
  
    ConnectDatabase = True
    On Error GoTo 0: Exit Function
    
amyhong:
    MsgBox "Connect Database Server Error! Please Contact Admin!" & Chr(13) & Err.Description
    Err.Clear: On Error GoTo 0: Exit Function
    
End Function


Function GetUserID() As String
'子程序，获取用户ID
    Dim myID As Object
    Set myID = CreateObject("Wscript.Network")                                                  'Object Wscipt.Network待查
     GetUserID = myID.UserName
End Function


Sub getTableList(ByVal mydb As String, ByVal myAccess As String)
'子程序，只读提取用户权限 并显示对应信息


myUserID = GetUserID()
Set myConn = New ADODB.Connection
If ConnectDatabase(mydb, 1) Then
   Set myRs = New ADODB.Recordset
       
    Select Case myAccess
    Case "Read/Write"
            If Form_DB.LB_Tablelist.ListCount > 0 Then Form_DB.LB_Tablelist.Clear
            myRs.Open "0User_authority", myConn, adLockReadOnly, adCmdTable
            myRs.Find "[User_ID] Like '" & myUserID & "'", , adSearchForward, 1       '搜索用户名
             If Not myRs.EOF Then
                 For i = 1 To myRs.Fields.Count - 1
                   If myRs.Fields(i).Value = myAccess Then Form_DB.LB_Tablelist.AddItem myRs.Fields(i).Name
                 Next i
             End If
             myRs.Close
    Case "Read"
            myRs.Open "1ExchangeRate", myConn, adLockReadOnly, adCmdTable
                 Form_RP.TB_EX.Text = myRs(0).Value
                 Form_RP.TB_EXR.Text = myRs(0).Value
            myRs.Close
            
            If Form_RP.LB_table.ListCount > 0 Then Form_RP.LB_table.Clear
            myRs.Open "0User_authority", myConn, adLockReadOnly, adCmdTable
            myRs.Find "[User_ID] Like '" & myUserID & "'", , adSearchForward, 1       '搜索用户名
             If Not myRs.EOF Then
                 For i = 1 To myRs.Fields.Count - 1
                  If Left(myRs.Fields(i).Value, 4) = myAccess Then Form_RP.LB_table.AddItem myRs.Fields(i).Name
                 Next i
             End If
            myRs.Close
            
            If Form_RP.LB_comment.ListCount > 0 Then Form_RP.LB_comment.Clear
            myRs.Open "0Comments_Report", myConn, adLockReadOnly, adCmdTable
            myRs.Find "[User_ID] Like '" & myUserID & "'", , adSearchForward, 1       '搜索用户名
             If Not myRs.EOF Then
                 For i = 1 To myRs.Fields.Count - 1
                  If Left(myRs.Fields(i).Value, 4) = "Read" Then Form_RP.LB_comment.AddItem myRs.Fields(i).Name
                 Next i
             End If
            myRs.Close
    End Select

End If
Set myRs = Nothing
myConn.Close: Set myConn = Nothing

End Sub

Public Function GetYearMonth(ByVal mydb As String, ByVal TableName As String, ByVal myWay As String, _
                            Optional myDate As String) As String
'监控程序，记录用户登录编辑轨迹，后台运行； mydb,TableName,myWay信息记录来源于Form

  Set myConn = New ADODB.Connection
  If ConnectDatabase(mydb, 3) Then
     Set myRs = New ADODB.Recordset
     
     '只读提取/或确定报告时间
     myRs.Open "0User_Monitor", myConn, adOpenKeyset, adLockOptimistic, adCmdTable
        myRs.MoveFirst
        Dim i As Long
        For i = 0 To myRs.RecordCount - 1
            If Not IsNull(myRs.Fields(myDate).Value) And CDate(myRs(myDate).Value) > DateAdd("m", -2, Now) Then  'DateAdd函数,基于目前时间前推2个月
                GetYearMonth = Format(DateAdd("m", -1, CDate(myRs(myDate).Value)), "mmm yyyy")                    'Format函数,日期格式设置
                Exit For
            End If
            myRs.MoveNext
        Next
        If GetYearMonth = Empty Then GetYearMonth = Format(DateAdd("m", -2, Now), "mmm yyyy")
     
     '开始监控
     myRs.Find "[User_ID] Like '" & myUserID & "'", , adSearchForward, 1       '搜索用户名
        On Error Resume Next
        If myRs.EOF Then
          myRs.AddNew
            myRs("User_ID") = myUserID
            If IsNull(myRs.Fields(myWay & "_Count_" & TableName).Value) Then
            myRs.Fields(myWay & "_Count_" & TableName).Value = 1
            Else
              myRs.Fields(myWay & "_Count_" & TableName) = myRs.Fields(myWay & "_Count_" & TableName).Value + 1
            End If
            myRs.Fields(myWay & "_Time_" & TableName) = Now
          myRs.Update
        Else
            If IsNull(myRs.Fields(myWay & "_Count_" & TableName).Value) Then
              myRs.Fields(myWay & "_Count_" & TableName).Value = 1
            Else
              myRs.Fields(myWay & "_Count_" & TableName) = myRs.Fields(myWay & "_Count_" & TableName).Value + 1
            End If
            myRs.Fields(myWay & "_Time_" & TableName) = Now
          myRs.Update
       End If
     myRs.Close
     
  End If
  Set myRs = Nothing
  myConn.Close: Set myConn = Nothing

End Function

Sub UserMonitor(ByVal mydb As String, ByVal TableName As String, ByVal myWay As String)
'监控程序，记录用户登录编辑轨迹，后台运行； mydb,TableName,myWay信息记录来源于Form

  myUserID = GetUserID()
  Set myConn = New ADODB.Connection
  If ConnectDatabase(mydb, 3) Then
     Set myRs = New ADODB.Recordset
     
     myRs.Open "0User_Monitor", myConn, adOpenKeyset, adLockOptimistic, adCmdTable
     myRs.Find "[User_ID] Like '" & myUserID & "'", , adSearchForward, 1       '搜索用户名
        On Error Resume Next
        If myRs.EOF Then
          myRs.AddNew
            myRs("User_ID") = myUserID
            If IsNull(myRs.Fields(myWay & "_Count_" & TableName).Value) Then            'mydb,TableName,myWay信息记录来源于Form，确定子段名
            myRs.Fields(myWay & "_Count_" & TableName).Value = 1
            Else
              myRs.Fields(myWay & "_Count_" & TableName) = myRs.Fields(myWay & "_Count_" & TableName).Value + 1
            End If
            myRs.Fields(myWay & "_Time_" & TableName) = Now
          myRs.Update
        Else
            If IsNull(myRs.Fields(myWay & "_Count_" & TableName).Value) Then
              myRs.Fields(myWay & "_Count_" & TableName).Value = 1
            Else
              myRs.Fields(myWay & "_Count_" & TableName) = myRs.Fields(myWay & "_Count_" & TableName).Value + 1
            End If
            myRs.Fields(myWay & "_Time_" & TableName) = Now
          myRs.Update
       End If
     myRs.Close
  End If
  Set myRs = Nothing
  myConn.Close: Set myConn = Nothing

End Sub


Private i As Long

Sub Comment_Report(ByVal mydb As String, ByVal EX As Single, ByVal myComment As String, _
                Optional myDate As String, Optional ByVal myProv As String, Optional ByVal myStoreNo As Long)
'通用报告生成，From获取mydb，EX，报告名，报告日期及其他报告所需内容

Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error GoTo Handler


'Excel报告模板内容获取
Dim TitelDate, TitelEX As String
    TitelDate = "Ending as of " & myDate
    TitelEX = "In USD @ " & EX
Dim boLO As Boolean: boLO = False
Dim boSqm As Boolean: boSqm = False
Dim boVertical As Boolean: boVertical = False
Dim ExName, myStr, myStr2 As String
    ExName = myComment & ".xlsx"
    If Dir$(DBPath & "Report Format\" & ExName) = "" Then MsgBox myComment & "改报告不存在，请联系Admin": Exit Sub
    Workbooks.Open FileName:=DBPath & "Report Format\" & ExName, ReadOnly:=True
    Application.Calculation = xlManual
    Calculate
    
    Select Case myComment
'        Case "Monthly store performance"
'            boLO = True
'            myStr = " WHERE (((A.Status)='Open')) ORDER BY A.[Format Name] DESC , A.[OPS Region], A.Province, A.[Great City];"
'            myStr2 = " WHERE (((A.Status)='Open')) ORDER BY A.[Format Name] DESC , A.[OPS Region], A.Province, A.[Great City];"
        Case "Provincial Store Pipeline & Performance"
            myStr = " WHERE (((A.Province)='" & myProv & "') AND ((A.Status)='open')) ORDER BY A.[Store Name];"
'            myStr2 = " WHERE (((A.Province)='" & myProv & "') AND ((A.Status)='In Progress'));"
        Case "Membership Pipeline & Performance"
            myStr = " WHERE (((A.[Format Name])='Membership') AND ((A.Status)='open')) ORDER BY A.Province;"
'            myStr2 = " WHERE (((A.[Format Name])='Membership') AND ((A.Status)='In Progress'));"
       Case Else
            With Sheets("Summary")
                If .Range("TitelEX").Value = "Empty" Then TitelEX = Empty
                If .Range("boLO").Value = "True" Then boLO = True
                If .Range("boSqm").Value = "True" Then boSqm = True
                If NameExists(ExName, "boVertical") Then
                    If .Range("boVertical").Value = "True" Then boVertical = True
                End If
                If myStoreNo = 0 Then
                    myStr = .Range("myStr").Value
                Else
                    myStr = IIf(boVertical, myStoreNo & "));", " WHERE ((([Store No]) = " & myStoreNo & "));")
                End If
            End With
    End Select
    
Set myConn = New ADODB.Connection
If ConnectDatabase(mydb, 1) Then               '连接数据库SRPDatabase
Set myRs = New ADODB.Recordset

Sheets("Summary").Select
With Sheets("Summary")

  .Cells(2, 2).Value = TitelDate
  If TitelEX <> "" Then .Cells(3, 2).Value = TitelEX
  If myProv <> "" Then .Cells(7, 2).Value = myProv: Calculate
        
        myRs.Open GetCommentSql(EX, boSqm, boVertical) & myStr, myConn, adUseClient, adOpenStatic, adLockReadOnly '运行mySql
        If Not myRs.EOF Then
          If boVertical = False Then
                If boLO Then .ListObjects("TSummary").ShowTotals = False
                .Range("B10").CopyFromRecordset myRs                        '将记录集贴到报告中
                If boLO Then .ListObjects("TSummary").ShowTotals = True
          Else
                For i = 0 To myRs.Fields.Count - 1
                .Cells(i + 10, 10).Value = myRs(i).Value
                Next
                .Visible = 0
          End If
        Else
            MsgBox "没有记录"
        End If
        myRs.Close
End With
                          
If SheetExists(ExName, "Report") Then
Sheets("Report").Select
With Sheets("Report")
  If NameExists(ExName, "myStr") Then
  If myStoreNo = 0 Then
    myStr2 = .Range("myStr").Value
  Else
    myStr2 = IIf(boVertical, myStoreNo & "));", " WHERE ((([Store No]) = " & myStoreNo & "));")
  End If
  End If
  .Cells(2, 2).Value = TitelDate
  If TitelEX <> "" Then .Cells(3, 2).Value = TitelEX
        
        myRs.Open GetCommentSql(EX, boSqm, boVertical) & myStr2, myConn, adUseClient, adOpenStatic, adLockReadOnly '运行mySql2
        If Not myRs.EOF Then
          If boVertical = False Then
                If boLO Then .ListObjects("TReport").ShowTotals = False
                .Range("B10").CopyFromRecordset myRs
                If boLO Then .ListObjects("TReport").ShowTotals = True
          Else
                For i = 0 To myRs.Fields.Count - 1
                .Cells(i + 10, 10).Value = myRs(i).Value
                Next
                .Visible = 0
          End If
         End If
        If boLO = False Then Call Province_RP
        myRs.Close
End With
End If

End If
Set myRs = Nothing
myConn.Close: Set myConn = Nothing
Application.Calculation = xlAutomatic
Exit Sub

Handler:
    MsgBox vbNewLine & "An Error occurred!" & vbNewLine & vbNewLine & _
        "Error Number: " & Err.Number & vbNewLine & vbNewLine & _
        "Error Message: " & Err.Description & vbNewLine & vbNewLine & _
        "Error Source: " & Err.Source, vbokoly, "Error"
Application.Calculation = xlAutomatic
End

End Sub

                    Private Sub Province_RP()
                    
                    Dim i As Integer, myRange As Range
                    Calculate
                    
                      Sheets("Existing Store").Select
                        ActiveSheet.Outline.ShowLevels RowLevels:=2
                        Cells.Select
                        Selection.EntireRow.Hidden = False
                      Selection.Copy
                      Selection.PasteSpecial Paste:=xlPasteValues
                      On Error GoTo amyhong
                        Dim arraya(5) As Integer, arrayb(5) As Range
                            arraya(1) = Range("R_HMKT").Value
                            arraya(2) = Range("R_SC").Value
                            arraya(3) = Range("R_NBH").Value
                            arraya(4) = Range("R_DCH").Value
                                Set arrayb(1) = Range("C_HMKT")
                                Set arrayb(2) = Range("C_SC")
                                Set arrayb(3) = Range("C_NBH")
                                Set arrayb(4) = Range("C_DCH")
                            For i = 1 To 4
                              If arraya(i) = 0 Then
                                Set myRange = arrayb(i).Offset(-1).Resize(arrayb(i).Rows.Count + 4)
                                myRange.Select
                                Selection.EntireRow.Delete    '             Selection.EntireRow.Hidden = True
                              Else
                              If arraya(i) <> arrayb(i).Rows.Count Then
                                Set myRange = arrayb(i).Offset(arraya(i)).Resize(arrayb(i).Rows.Count - arraya(i))
                                myRange.Select
                                Selection.EntireRow.Delete    '             Selection.EntireRow.Hidden = True
                              End If
                              End If
                            Next i
amyhong:
                      Range("E24").Select
                        
                      Sheets("Approved Store").Select
                        ActiveSheet.Outline.ShowLevels RowLevels:=2
                        Cells.Select
                        Selection.EntireRow.Hidden = False
                      Selection.Copy
                      Selection.PasteSpecial Paste:=xlPasteValues
                        Dim arrayx(5) As Integer, arrayy(5) As Range
                            arrayx(1) = Range("R_2014").Value
                            arrayx(2) = Range("R_2015").Value
                            arrayx(3) = Range("R_2016").Value
                            arrayx(4) = Range("R_2017").Value
                            arrayx(5) = Range("R_2018").Value
                                Set arrayy(1) = Range("C_2014")
                                Set arrayy(2) = Range("C_2015")
                                Set arrayy(3) = Range("C_2016")
                                Set arrayy(4) = Range("C_2017")
                                Set arrayy(5) = Range("C_2018")
                            For i = 1 To 5
                              If arrayx(i) = 0 Then
                                Set myRange = arrayy(i).Offset(-1).Resize(arrayy(i).Rows.Count + 4)
                                myRange.Select
                                Selection.EntireRow.Delete    '             Selection.EntireRow.Hidden = True
                              Else
                              If arrayx(i) <> arrayy(i).Rows.Count Then
                                Set myRange = arrayy(i).Offset(arrayx(i)).Resize(arrayy(i).Rows.Count - arrayx(i))
                                myRange.Select
                                Selection.EntireRow.Delete   '             Selection.EntireRow.Hidden = True
                              End If
                              End If
                            Next i
                      Range("F10").Select
                      
                    Sheets("Summary").Select
                    ActiveWindow.SelectedSheets.Delete
                    Sheets("Report").Select
                    ActiveWindow.SelectedSheets.Delete
                    
                    End Sub


Private Function GetCommentSql(ByVal EX As Single, ByVal SQM As Boolean, ByVal boVertical As Boolean) As String
'子程序，SQL语句内容：按照Excel模板信息获取mydb中对应的信息

Dim mySql, myStr1, myStr2  As String, ExcelCol As Long
Dim myObjS:    Set myObjS = CreateObject("Scripting.Dictionary")                 '****Objiect Scripting.Dictionary待查
With ActiveSheet
  If boVertical = False Then
  '水平或垂直=水平
            ExcelCol = .Cells(4, 2).CurrentRegion.Columns.Count
         If Left(.Cells(4, 2), 1) = "0" Then
         '公共报告ABC或管理报告01=01，User ID为联系词
            mySql = "SELECT "      '开始
                    For i = 1 To ExcelCol
                        mySql = mySql & "[" & .Cells(4, i + 1) & "].[" & .Cells(5, i + 1) & "], "
                        myObjS(.Cells(4, i + 1).Value) = ""                      '****myObjS(TableName)="",收集不重复的TableName的信息
                    Next
             mySql = Left(mySql, Len(mySql) - 2) & " FROM "          '准备取表名
             myK = myObjS.Keys                                                   '****读取TableName， .keys属性
                    myStr1 = "[" & myK(0) & "]"
                    For i = 1 To UBound(myK)
                       myStr2 = myStr2 & "("
                    Next
                    mySql = mySql & myStr2 & myStr1
                    For i = 1 To UBound(myK)
                       mySql = mySql & " LEFT JOIN [" & myK(i) & "] ON [" & myK(0) & "].[User_ID] = [" & myK(i) & "].[User_ID])"
                    Next
        
         Else
         '公共报告ABC或管理报告01=ABC，Store No为联系词
            mySql = "SELECT "      '开始
                    For i = 1 To ExcelCol
                        If Right(.Cells(5, i + 1), 3) = "RMB" Then                                 'RMB转化为当前汇率， 平方米转化为其他计量单位
                        mySql = mySql & Left(.Cells(4, i + 1), 1) & ".[" & .Cells(5, i + 1) & "]/" & EX & " AS [" & .Cells(5, i + 1) & "_US], "
                        Else
                            If SQM = False And Right(.Cells(5, i + 1), 3) = "SqM" Then
                            mySql = mySql & Left(.Cells(4, i + 1), 1) & ".[" & .Cells(5, i + 1) & "]*10.7639 AS [" & .Cells(5, i + 1) & "_Sqft], "
                            Else
                        mySql = mySql & Left(.Cells(4, i + 1), 1) & ".[" & .Cells(5, i + 1) & "], "
                            End If
                        End If
                    myObjS(.Cells(4, i + 1).Value) = ""
                    Next
             mySql = Left(mySql, Len(mySql) - 2) & " FROM "          '准备取表名
             myK = myObjS.Keys
                    myStr1 = "[" & myK(0) & "] AS " & Left(myK(0), 1)
                    For i = 1 To UBound(myK)
                       myStr2 = myStr2 & "("
                    Next
                    mySql = mySql & myStr2 & myStr1
                    For i = 1 To UBound(myK)
                       mySql = mySql & " LEFT JOIN [" & myK(i) & "] AS " & Left(myK(i), 1) & " ON " _
                            & Left(myK(0), 1) & ".[Store No] = " & Left(myK(i), 1) & ".[Store No])"
                    Next
          End If
  Else
  '水平或垂直=垂直
            ExcelCol = .Cells(10, 4).CurrentRegion.Rows.Count
            mySql = "SELECT "      '开始
                    For i = 1 To ExcelCol
                        If Right(.Cells(i + 9, 5), 3) = "RMB" Then
                        mySql = mySql & Left(.Cells(i + 9, 4), 1) & ".[" & .Cells(i + 9, 5) & "]/" & EX & " AS [" & .Cells(i + 9, 5) & "_US], "
                        Else
                            If SQM = False And Right(.Cells(i + 9, 5), 3) = "SqM" Then
                            mySql = mySql & Left(.Cells(i + 9, 4), 1) & ".[" & .Cells(i + 9, 5) & "]*10.7639 AS [" & .Cells(i + 9, 5) & "_Sqft], "
                            Else
                        mySql = mySql & Left(.Cells(i + 9, 4), 1) & ".[" & .Cells(i + 9, 5) & "], "
                            End If
                        End If
                    myObjS(.Cells(i + 9, 4).Value) = ""
                    Next
             mySql = Left(mySql, Len(mySql) - 2) & " FROM "          '准备取表名
             myK = myObjS.Keys
                    myStr1 = "[" & myK(0) & "] AS " & Left(myK(0), 1)
                    For i = 1 To UBound(myK)
                       myStr2 = myStr2 & "("
                    Next
                    mySql = mySql & myStr2 & myStr1
                    For i = 1 To UBound(myK)
                       mySql = mySql & " LEFT JOIN [" & myK(i) & "] AS " & Left(myK(i), 1) & " ON " _
                            & Left(myK(0), 1) & ".[Store No] = " & Left(myK(i), 1) & ".[Store No])"
                    Next
                    mySql = mySql & " WHERE (((" & Left(myK(0), 1) & ".[Store No]) = "
                    '条件语句，连接Table.[Store No]=
  End If
End With

Set myObjS = Nothing
GetCommentSql = mySql

End Function

Private Function NameExists(ByVal myBook As String, ByVal FindName As String) As Boolean
'子程序，判断Excel内名称定义是否存在
  NameExists = False
  Dim Nm As Name
  For Each Nm In Workbooks(myBook).Names
    If Nm.Name = FindName Then
      NameExists = True
      Exit Function
    End If
  Next Nm
End Function

Private Function SheetExists(ByVal myBook As String, ByVal FindSheet As String) As Boolean
'子程序，判断Excel内Sheet是否存在
  SheetExists = False
  Dim Nm As Worksheet
  For Each Nm In Workbooks(myBook).Sheets
    If Nm.Name = FindSheet Then
      SheetExists = True
      Exit Function
    End If
  Next Nm
End Function


'打开文件
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'******未使用，待查

Sub Presonal_Report(ByVal mydb As String, ByVal mySql As String)
'根据Form记录查询记录，mySql来源于Form

Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error GoTo Handler

Set myConn = New ADODB.Connection
If ConnectDatabase(mydb, 1) Then               '连接数据库SRPDatabase
    Set myRs = New ADODB.Recordset
 myRs.Open mySql, myConn, adOpenStatic, adLockReadOnly

'Excel呈现数据
        If Not myRs.EOF Then
'            Dim iRow, iRo As Long
            Dim iFied, iFd As Integer
'            iRow = myRs.RecordCount - 1
            iFied = myRs.Fields.Count - 1
            myRs.MoveFirst
                
'                Set ObjExcel = New Excel.Application
'                ObjExcel.Visible = True
                Dim objworkbook As Object
                Dim objsheet As Object
                Set objworkbook = Workbooks.Add
                Set objsheet = objworkbook.ActiveSheet
                
                With objsheet
                    For iFd = 0 To iFied
                    .Cells(1, iFd + 1).Value = myRs.Fields(iFd).Name
                    Next
                    .Range(.Cells(1, 1), .Cells(1, iFied + 1)).Font.Bold = True
                    .Range("a2").CopyFromRecordset myRs
                End With
         Else
          MsgBox "没有记录，请核查是否 搜索条件错误"
         End If
 End If
 myRs.Close: Set myRs = Nothing
 myConn.Close: Set myConn = Nothing
Exit Sub

Handler:
    MsgBox vbNewLine & "An Error occurred!" & vbNewLine & vbNewLine & _
        "Error Number: " & Err.Number & vbNewLine & vbNewLine & _
        "Error Message: " & Err.Description & vbNewLine & vbNewLine & _
        "Error Source: " & Err.Source, vbokoly, "Error"
End

End Sub


Sub GetPackage(myPackage As String, myStore As String, Optional myMethod As String)
'附加功能，直指路径下文件信息

Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error GoTo Handler

   Dim i As Integer: i = 0
   Dim myOpen As Long
   Dim myFile As String: myFile = Dir(BackUpPath & myPackage & Val(myStore) & "*")
                    '     myFile = BackUpPath & myPackage & "*" & myStore & "*"
    
    Select Case myMethod
    Case "Exce1"
        Do Until myFile = ""
            i = i + 1
            Workbooks.Open FileName:=BackUpPath & myPackage & myFile, PassWord:="111...", UpdateLinks:=0, ReadOnly:=True
            myFile = Dir
        Loop
    Case Else
        Do Until myFile = ""
            i = i + 1
            myOpen = ShellExecute(0, "Open", BackUpPath & myPackage & myFile, "", "", vbNormalFocus)
            myFile = Dir
        Loop
    End Select
    If i < 1 Then MsgBox "没有找到该项目的资料，请与RE Planning复核."
Exit Sub

Handler:
    MsgBox vbNewLine & "An Error occurred!" & vbNewLine & vbNewLine & _
        "Error Number: " & Err.Number & vbNewLine & vbNewLine & _
        "Error Message: " & Err.Description & vbNewLine & vbNewLine & _
        "Error Source: " & Err.Source, vbokoly, "Error"
End


End Sub




Option Base 1
  Private arrFid() As String, arrRcs()               '用于存放Excel表格信息
  Private TRow, TCol, myR, myC, SN As Long, myKey As String       'TRow记录行数,TCol字段列数，myR, myC为对应计数器
  Private arrF(), myF As Integer, arrTF() As String  '用于存放匹配子段

Sub Maintenance_table(ByVal mydb As String, ByVal myTab As String, ByVal myMet As String)
On Error GoTo Handler

  '确定更新范围及内容
  Dim myRange As Range, myBook As String
  myBook = ActiveWorkbook.Name
  Set myRange = Application.InputBox("请选择更新范围", "选择范围", Type:=8)
         TRow = myRange.Rows.Count:  TCol = myRange.Columns.Count
         ReDim arrRcs(2 To TRow, TCol):  ReDim arrFid(TCol)
         For myC = 1 To TCol
           Select Case myRange(1, myC).Value
            Case "RPUID"
            SN = myC:  myKey = "[RPUID]"
            Exit For
           Case Is = "Store No"
            SN = myC:  myKey = "[Store No]"
            Exit For
           Case "User_ID"
            SN = myC:  myKey = "User_ID"
            Exit For
           Case "CREC Month"
            SN = myC:  myKey = "[CREC Month]"
            Exit For
           Case Is = "ExchangeRate"
            SN = myC:  myKey = "ExchangeRate"
          End Select
       Next myC
       If SN = Empty Then MsgBox "没有店号/UserID，无效表格": Exit Sub
       For myC = 1 To TCol
           arrFid(myC) = Trim(myRange(1, myC).Value)
       Next myC
         For myR = 2 To TRow
           For myC = 1 To TCol
              arrRcs(myR, myC) = Trim(myRange(myR, myC).Value)
           Next myC
         Next myR

'连接操作Database
    Set myConn = New ADODB.Connection
    If ConnectDatabase(mydb, 3) Then
       Set myRs = New ADODB.Recordset
       myRs.CursorLocation = adUseClient
      
            '匹配更新子段
            myRs.Open myTab, myConn, adOpenKeyset, adLockOptimistic, adCmdTable
            myF = 0
            ReDim arrF(TCol), arrTF(TCol)
            For i = 0 To myRs.Fields.Count - 1
              For myC = 1 To TCol
                 If arrFid(myC) = myRs.Fields(i).Name Then
                  myF = myF + 1
                  arrF(myF) = myC
                  arrTF(myF) = myRs.Fields(i).Name
                  Exit For
                 End If
              Next
            Next
            myRs.Close
            
            '开始事物处理
            myConn.BeginTrans
                  
                Select Case myMet
                Case "更新记录"
                     Call Update_Record(mydb, myTab)
                Case "整表替换"
                     Call Delete_Record(mydb, myTab, Empty)
                     Call Addional_Record(mydb, myTab)
                Case "追加记录"
                     Call Addional_Record(mydb, myTab)
                Case Else
                    MsgBox "无效操作，将退出"
                    Exit Sub
                End Select
            
            If MsgBox(myTab & "确认更新表格?", vbQuestion + vbYesNo) = vbYes Then
              myConn.CommitTrans
            Else
              myConn.RollbackTrans
            End If
        
    End If
    myRs.Close: Set myRs = Nothing
    myConn.Close: Set myConn = Nothing

'标出更新子段
     Windows(myBook).Activate
     For i = 1 To myF
       Range(myRange(1, arrF(i)), myRange(TRow, arrF(i))).Interior.Color = RGB(244, 115, 33)
     Next
     MsgBox "黄色部分已经更新如Databasee，请复核"
     Exit Sub

Handler:
    MsgBox vbNewLine & "An Error occurred!" & vbNewLine & vbNewLine & _
        "Error Number: " & Err.Number & vbNewLine & vbNewLine & _
        "Error Message: " & Err.Description & vbNewLine & vbNewLine & _
        "Error Source: " & Err.Source, vbokoly, "Error"
    End
    
End Sub
Sub Maintenance_Excute(ByVal mydb As String, ByVal myTab As String, ByVal myMet As String, ByVal myStr As String)
On Error GoTo Handler

'连接操作Database
    Set myConn = New ADODB.Connection
    If ConnectDatabase(mydb, 3) Then
       Set myRs = New ADODB.Recordset
       myRs.CursorLocation = adUseClient
      
            '开始事物处理
            myConn.BeginTrans
                  
                Select Case myMet
                Case "删除记录"
                     Call Delete_Record(mydb, myTab, myStr)
                Case "添加字段"
                     Call Append_Filed(mydb, myTab, myStr)
                Case "特别删除"
                     Call Delete_Record(mydb, myTab, myStr)
                End Select
            
            If MsgBox(myTab & "确认" & myMet & "?", vbQuestion + vbYesNo) = vbYes Then
              myConn.CommitTrans
            Else
              myConn.RollbackTrans
            End If
        
    End If
    Set myRs = Nothing
    myConn.Close: Set myConn = Nothing
Exit Sub

Handler:
    MsgBox vbNewLine & "An Error occurred!" & vbNewLine & vbNewLine & _
        "Error Number: " & Err.Number & vbNewLine & vbNewLine & _
        "Error Message: " & Err.Description & vbNewLine & vbNewLine & _
        "Error Source: " & Err.Source, vbokoly, "Error"
    End
    
End Sub

Private Sub Delete_Record(ByVal mydb As String, ByVal myTab As String, ByVal myStr As String)
 '删除记录
    mySql = "DELETE * FROM " & myTab & myStr
    myRs.Open mySql, myConn, adOpenStatic, adLockOptimistic

End Sub

Private Sub Append_Filed(ByVal mydb As String, ByVal myTab As String, ByVal myStr As String)
 '添加字段
    mySql = "ALTER TABLE " & myTab & " ADD " & myStr
    myRs.Open mySql, myConn, adOpenStatic, adLockOptimistic

End Sub
Private Sub Update_Record(ByVal mydb As String, ByVal myTab As String)
 '更新记录
    myRs.Open myTab, myConn, , , adCmdTable
    If myRs.RecordCount < 1 Then
       For myR = 2 To TRow
         myRs.AddNew
         For i = 1 To myF
          myRs.Fields(arrTF(i)) = arrRcs(myR, arrF(i))
         Next
         myRs.Update
       Next
    Else
      '搜索店号，如果店号已存在 编辑记录, 否则新增记录
      For myR = 2 To TRow
        myRs.MoveFirst
        myRs.Find myKey & " Like '" & arrRcs(myR, SN) & "'", , adSearchForward, 1
            
        If myRs.EOF Then
          myRs.AddNew
          For i = 1 To myF
            If arrRcs(myR, arrF(i)) <> "" Then
               If myRs.Fields(arrTF(i)).Type = 7 Then
                 myRs.Fields(arrTF(i)) = CDate(Replace(arrRcs(myR, arrF(i)), ".", "/"))
               Else
                 myRs.Fields(arrTF(i)) = arrRcs(myR, arrF(i))
               End If
            End If
          Next
          myRs.Update
        Else
          For i = 1 To myF
            If arrRcs(myR, arrF(i)) <> "" Then
               If myRs.Fields(arrTF(i)).Type = 7 Then
                 myRs.Fields(arrTF(i)).Value = CDate(Replace(arrRcs(myR, arrF(i)), ".", "/"))
               Else
                 myRs.Fields(arrTF(i)).Value = arrRcs(myR, arrF(i))
               End If
            End If
          Next
          myRs.Update
        End If
      Next
    End If
End Sub

Private Sub Addional_Record(ByVal mydb As String, ByVal myTab As String)
  
  myRs.Open myTab, myConn, , , adCmdTable
  '追加记录
   For myR = 2 To TRow
     myRs.AddNew
     For i = 1 To myF
            If arrRcs(myR, arrF(i)) <> "" Then
               If myRs.Fields(arrTF(i)).Type = 7 Then
                 myRs.Fields(arrTF(i)) = CDate(Replace(arrRcs(myR, arrF(i)), ".", "/"))
               Else
                 myRs.Fields(arrTF(i)) = arrRcs(myR, arrF(i))
               End If
            End If
     Next
     myRs.Update
   Next

End Sub

Sub Export_table(ByVal mydb As String, myTable As String)

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Set myConn = New ADODB.Connection
If ConnectDatabase(mydb, 1) Then
    Set myRs = New ADODB.Recordset
 myRs.Open myTable, myConn, adLockReadOnly, adCmdTable

        If Not myRs.EOF Then
            Dim iFied, iFd As Integer
            iFied = myRs.Fields.Count - 1
            myRs.MoveFirst
                
                Dim objworkbook As Object
                Dim objsheet As Object
                Set objworkbook = Workbooks.Add
                Set objsheet = objworkbook.ActiveSheet
                
                With objsheet
                    For iFd = 0 To iFied
                    .Cells(1, iFd + 1).Value = myRs.Fields(iFd).Name
                    Next
                    .Range(.Cells(1, 1), .Cells(1, iFied + 1)).Font.Bold = True
                    .Range("a2").CopyFromRecordset myRs
                End With
         End If
 End If
 myRs.Close: Set myRs = Nothing
 myConn.Close: Set myConn = Nothing
Exit Sub

End Sub





Sub RetriveAuthority(ByVal mydb As String)
'将Excel模板内容导入到Access中

Application.ScreenUpdating = False
Application.DisplayAlerts = False
 On Error GoTo Handler

Dim i As Long, myFile As String
Dim wroC, wroR, TotC As Integer
Dim myTable(0 To 2), myUser(100) As String
    myTable(0) = "User_Information"
    myTable(1) = "Comments_Report"
    myTable(2) = "User_authority"
Dim myTab As String, myRange As Range

Workbooks.Open FileName:=GetExcelName(), UpdateLinks:=0, ReadOnly:=True
    Application.Calculation = xlManual
    Calculate
    myFile = ActiveWorkbook.Name
    With Sheets(mydb)
       wroC = 2
       Do Until Range("User_ID")(1, wroC).Value = ""
        TotC = wroC - 2
        myUser(TotC) = LCase(Range("User_ID")(1, wroC).Value)
        wroC = wroC + 1
       Loop
       
     '连接操作Database
     Set myConn = New ADODB.Connection
     If ConnectDatabase(mydb, 3) Then         '4
       Set myRs = New ADODB.Recordset
       myRs.CursorLocation = adUseClient
       
         For i = 0 To 2                     '3区域
         myTab = "0" & myTable(i)
         Set myRange = .Range(myTable(i))
         myConn.BeginTrans
            myRs.Open myTab, myConn, adOpenKeyset, adLockOptimistic, adCmdTable
            For wroC = 2 To TotC + 2             '2列
                myRs.Find "[User_ID] Like '" & myUser(wroC - 2) & "'", , adSearchForward, 1
                If Not myRs.EOF Then
                    myRs.Close
                    myRs.Open "DELETE * FROM [" & myTab & "] WHERE [User_ID]= '" & myUser(wroC - 2) & "'", myConn, adOpenStatic, adLockOptimistic
                    myRs.Open myTab, myConn, adOpenStatic, adLockOptimistic, adCmdTable
                End If
                myRs.AddNew
                myRs.Fields("User_ID").Value = myUser(wroC - 2)
                If i = 0 Then
                    myRs.Fields("Authority date").Value = Date
                    myRs.Fields("Comment").Value = ""
                End If
                For wroR = 1 To myRange.Rows.Count        '1行
                    If myRange(wroR, wroC).Value <> "" Then
                    myRs.Fields(myRange(wroR, 1).Value).Value = myRange(wroR, wroC).Value
                    End If
                Next                     '1
                myRs.Update
            Next           '2
        myConn.CommitTrans
        myRs.Close
        Next               '3
    End If                 '4
    Set myRs = Nothing
    myConn.Close: Set myConn = Nothing

    End With
Workbooks(myFile).Close
'Application.Calculation = xlAutomatic

MsgBox "Done"
Exit Sub

Handler:
    MsgBox vbNewLine & "An Error occurred!" & vbNewLine & vbNewLine & _
        "Error Number: " & Err.Number & vbNewLine & vbNewLine & _
        "Error Message: " & Err.Description & vbNewLine & vbNewLine & _
        "Error Source: " & Err.Source, vbokoly, "Error"
'Application.Calculation = xlAutomatic
End

End Sub




Sub RetriveExcelData(ByVal mydb As String, ByVal myTab As String)
'将Excel模板内容导入到Access中

Application.ScreenUpdating = False
Application.DisplayAlerts = False
 On Error GoTo Handler

Dim i As Long, myFile As String
Dim myFilling As String
If mydb = "ChinaRealtyDB" Then
'导入Database_Import
        Workbooks.Open FileName:=GetExcelName(), UpdateLinks:=0, ReadOnly:=True
            Application.Calculation = xlManual
            Calculate
            myFile = ActiveWorkbook.Name
        Dim FileName(), FiledVal(), FiledOrder()
        Dim myRange As Range, Trecords As Long
            Set myRange = Sheets("DB tools").Range("Database_Import")
            Trecords = myRange.Rows.Count
        ReDim FileName(Trecords), FiledVal(Trecords)
        Dim wro As Integer
        
        'Dim SN As Long, myKey As String
            For wro = 1 To Trecords
                Select Case myRange(wro, 1).Value
                 Case "RPUID"
                 SN = wro:  myKey = "[RPUID]"
                 Exit For
                Case Is = "Store No"
                 SN = wro:  myKey = "[Store No]"
                 Exit For
                Case "User_ID"
                 SN = wro:  myKey = "User_ID"
                 Exit For
                Case "CREC Month"
                 SN = wro:  myKey = "[CREC Month]"
                 Exit For
               End Select
            Next wro
            If myRange(SN, 2).Value = 0 Or myRange(SN, 2).Value = "" Then MsgBox "没有店号/RPUID，无效表格": Exit Sub
         
            For wro = 1 To Trecords
            If myRange.Cells(wro, 1) <> "" Then
               FileName(wro) = myRange.Cells(wro, 1).Value
               FiledVal(wro) = Trim(myRange.Cells(wro, 2).Value)
            End If
            Next wro
    
     Workbooks(myFile).Close
    
     '连接操作Database
     Set myConn = New ADODB.Connection
     If ConnectDatabase(mydb, 3) Then
       Set myRs = New ADODB.Recordset
       myRs.CursorLocation = adUseClient
       
                    '匹配更新子段
                    myRs.Open myTab, myConn, adOpenKeyset, adLockOptimistic, adCmdTable
                    Dim TTRecords As Long: TTRecords = 0
                    ReDim FiledOrder(Trecords)
                    For i = 0 To myRs.Fields.Count - 1
                    For wro = 1 To Trecords
                       If FileName(wro) = myRs.Fields(i).Name Then
                        TTRecords = TTRecords + 1
                        FiledOrder(TTRecords) = wro
                        Exit For
                       End If
                    Next
                    Next
                    myRs.Close
       
         myConn.BeginTrans
            myRs.Open myTab, myConn, adOpenStatic, adLockOptimistic, adCmdTable
            myRs.Find myKey & " Like '" & FiledVal(SN) & "'", , adSearchForward, 1
            If Not myRs.EOF Then
                myRs.Close
                myRs.Open "DELETE * FROM " & myTab & " WHERE " & myKey & "= " & FiledVal(SN), myConn, adOpenStatic, adLockOptimistic
                myRs.Open myTab, myConn, adOpenStatic, adLockOptimistic, adCmdTable
            End If
            myRs.AddNew
              For i = 1 To TTRecords
              If FiledVal(FiledOrder(i)) <> "" Then
               If myRs.Fields(FileName(FiledOrder(i))).Type = 7 Then
                 myRs.Fields(FileName(FiledOrder(i))).Value = CDate(Replace(FiledVal(FiledOrder(i)), ".", "/"))
               Else
                 myRs.Fields(FileName(FiledOrder(i))).Value = FiledVal(FiledOrder(i))
               End If
              End If
              Next
            myRs.Update
        myConn.CommitTrans
        myRs.Close
     
        If Mid(myTab, 2, 2) = "D " Or Mid(myTab, 2, 2) = "E " Then
        myFilling = "[" & Mid(myTab, 2, 2) & "Filling Records]"
        myConn.BeginTrans
            myRs.Open myFilling, myConn, adOpenStatic, adLockOptimistic, adCmdTable
            myRs.AddNew
              myRs("Modify User ID").Value = GetUserID
              myRs("Modify Date").Value = Now
              For i = 1 To TTRecords
              If FiledVal(FiledOrder(i)) <> "" Then
               If myRs.Fields(FileName(FiledOrder(i))).Type = 7 Then
                 myRs.Fields(FileName(FiledOrder(i))).Value = CDate(Replace(FiledVal(FiledOrder(i)), ".", "/"))     'Replace函数
               Else
                 myRs.Fields(FileName(FiledOrder(i))).Value = FiledVal(FiledOrder(i))
               End If
              End If
              Next
            myRs.Update
        myConn.CommitTrans
        myRs.Close
        End If
        
     End If

Else
'导入Morning Report
        Dim xlColumn, xlRow As Integer
        Dim xlSheet As Worksheet
        Dim myField, myColumn
        '    myField = Array("Store No", "Period", "DATE", "DAY", "Sales", "Transaction", "Ticket")
            myColumn = Array(0, 0, 2, 3, 7, 21, 32)
        Dim myRecords(31, 6)
        '    ChDir BackUpPath & "Morning Reprot\"
        '    myFile = Dir(BackUpPath & "Morning Reprot\*")
            
        Set myConn = New ADODB.Connection
        If ConnectDatabase(mydb, 3) Then
           Set myRs = New ADODB.Recordset
           myRs.Open myTab, myConn, adOpenForwardOnly, adLockOptimistic, adCmdTable
         
            Workbooks.Open FileName:=GetExcelName(), UpdateLinks:=0, ReadOnly:=True
            Application.Calculation = xlManual
            ' Do Until myFile = ""
            ' Workbooks.Open FileName:=BackUpPath & "Morning Reprot\" & myFile, UpdateLinks:=0, ReadOnly:=True
            Calculate
            myFile = ActiveWorkbook.Name
            Sheets("Sheet3").Select
            For Each xlSheet In Worksheets
             If Val(xlSheet.Name) <> 0 And Len(xlSheet.Name) = 3 Or Len(xlSheet.Name) = 4 Then
                i = 6
                With xlSheet
                   Do Until Val(.Cells(i, 2).Value) = 0
                       myRecords(i - 6, 0) = Val(xlSheet.Name)
                       myRecords(i - 6, 1) = CDate(Mid(myFile, 2, 10))
                       For xlColumn = 2 To 6
                       myRecords(i - 6, xlColumn) = .Cells(i, myColumn(xlColumn)).Value
                       Next
                   i = i + 1
                   Loop
                End With
                  myConn.BeginTrans
                    For xlRow = 0 To i - 7
                      myRs.AddNew
                      For xlColumn = 0 To 6
                      myRs.Fields(xlColumn) = myRecords(xlRow, xlColumn)
                      Next
                      myRs.Update
                    Next
                  myConn.CommitTrans
             End If
            Next
           myRs.Close
           Workbooks(myFile).Close
        ' myFile = Dir
        ' Loop
       End If
End If

Set myRs = Nothing
myConn.Close: Set myConn = Nothing
Application.Calculation = xlAutomatic
MsgBox "Done"
Exit Sub

Handler:
    MsgBox vbNewLine & "An Error occurred!" & vbNewLine & vbNewLine & _
        "Error Number: " & Err.Number & vbNewLine & vbNewLine & _
        "Error Message: " & Err.Description & vbNewLine & vbNewLine & _
        "Error Source: " & Err.Source, vbokoly, "Error"
Application.Calculation = xlAutomatic
End

End Sub

'********************************************************************************************************************************************************

'子程序

'********************************************************************************************************************************************************

Private Function GetExcelName() As String
'子程序，确定目标Excel表格
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    If fDialog.show = -1 Then
        GetExcelName = fDialog.SelectedItems(1)
        Else
        MsgBox "no file"
        End
    End If
End Function
