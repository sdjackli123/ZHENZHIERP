VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} GPXS 
   ClientHeight    =   9795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14880
   _ExtentX        =   26247
   _ExtentY        =   17277
   FolderFlags     =   1
   TypeLibGuid     =   "{E619BB91-EEBD-4CC0-8922-57C05142595E}"
   TypeInfoGuid    =   "{E346201F-F8F0-4936-A90F-94CC460B99BB}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "khjg"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=sc01\sql2008"
      IsSQL           =   -1  'True
      QuoteChar       =   34
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   12
   BeginProperty Recordset1 
      CommandName     =   "hzbb"
      CommDispId      =   1002
      RsDispId        =   1019
      CommandText     =   $"GPXS.dsx":0000
      ActiveConnectionName=   "khjg"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "客户"
         Caption         =   "客户"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "hzbbf"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "select * from v_cjbb"
      ActiveConnectionName=   "khjg"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "hzbb"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "客户"
         Caption         =   "客户"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "加工类别"
         Caption         =   "加工类别"
      EndProperty
      BeginProperty Field3 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "总数量"
         Caption         =   "总数量"
      EndProperty
      BeginProperty Field4 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "总金额"
         Caption         =   "总金额"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "客户"
         ChildField      =   "客户"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "jgmx"
      CommDispId      =   1020
      RsDispId        =   1035
      CommandText     =   "select 加工单位,MIN(日期) AS xa,MAX(日期) AS da  from jgmx where 加工单位=? and 日期 between ? and ?  GROUP BY 加工单位"
      ActiveConnectionName=   "khjg"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "加工单位"
         Caption         =   "加工单位"
      EndProperty
      BeginProperty Field2 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "xa"
         Caption         =   "xa"
      EndProperty
      BeginProperty Field3 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "da"
         Caption         =   "da"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   50
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "jgmxf"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "select *  from jgmx where 加工单位=? and 日期 between ?  and ?  order by 单号,IP  "
      ActiveConnectionName=   "khjg"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "jgmx"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   25
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "加工单位"
         Caption         =   "加工单位"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "品名"
         Caption         =   "品名"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "颜色"
         Caption         =   "颜色"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "锅号"
         Caption         =   "锅号"
      EndProperty
      BeginProperty Field5 
         Precision       =   7
         Size            =   4
         Scale           =   0
         Type            =   4
         Name            =   "数量"
         Caption         =   "数量"
      EndProperty
      BeginProperty Field6 
         Precision       =   7
         Size            =   4
         Scale           =   0
         Type            =   4
         Name            =   "单价"
         Caption         =   "单价"
      EndProperty
      BeginProperty Field7 
         Precision       =   7
         Size            =   4
         Scale           =   0
         Type            =   4
         Name            =   "金额"
         Caption         =   "金额"
      EndProperty
      BeginProperty Field8 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "日期"
         Caption         =   "日期"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "IP"
         Caption         =   "IP"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "发票已开"
         Caption         =   "发票已开"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "和约号"
         Caption         =   "和约号"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "备注"
         Caption         =   "备注"
      EndProperty
      BeginProperty Field13 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "顺序号 "
         Caption         =   "顺序号 "
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "单号"
         Caption         =   "单号"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "加工类别"
         Caption         =   "加工类别"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "库类"
         Caption         =   "库类"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "dy"
         Caption         =   "dy"
      EndProperty
      BeginProperty Field18 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "匹数"
         Caption         =   "匹数"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "计划号"
         Caption         =   "计划号"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "提取"
         Caption         =   "提取"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "支付标记"
         Caption         =   "支付标记"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "审核"
         Caption         =   "审核"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   202
         Name            =   "跟单"
         Caption         =   "跟单"
      EndProperty
      BeginProperty Field24 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "开票日期"
         Caption         =   "开票日期"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   4
         Scale           =   0
         Type            =   202
         Name            =   "ZL"
         Caption         =   "ZL"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   50
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   3
      BeginProperty Relation1 
         ParentField     =   "加工单位"
         ChildField      =   "Param1"
         ParentType      =   0
         ChildType       =   1
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "xa"
         ChildField      =   "Param2"
         ParentType      =   0
         ChildType       =   1
      EndProperty
      BeginProperty Relation3 
         ParentField     =   "da"
         ChildField      =   "Param3"
         ParentType      =   0
         ChildType       =   1
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset5 
      CommandName     =   "sczy_z"
      CommDispId      =   1036
      RsDispId        =   1061
      CommandText     =   "select distinct * from v_sczy_z where 单号=?"
      ActiveConnectionName=   "khjg"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "客户"
         Caption         =   "客户"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "单号"
         Caption         =   "单号"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "负责"
         Caption         =   "负责"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   202
         Name            =   "总备注"
         Caption         =   "总备注"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   120
         Scale           =   0
         Type            =   200
         Name            =   "计划"
         Caption         =   "计划"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "投染类别"
         Caption         =   "投染类别"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "面料用途"
         Caption         =   "面料用途"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "合同部门"
         Caption         =   "合同部门"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "合同负责"
         Caption         =   "合同负责"
      EndProperty
      BeginProperty Field10 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "下单日期"
         Caption         =   "下单日期"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "合同交期"
         Caption         =   "合同交期"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   50
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset6 
      CommandName     =   "sczy_x"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "select * from sczy_x order by 序号"
      ActiveConnectionName=   "khjg"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "sczy_z"
      IsRSReturning   =   -1  'True
      NumFields       =   26
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "客户"
         Caption         =   "客户"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "单号"
         Caption         =   "单号"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "款号"
         Caption         =   "款号"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "品名"
         Caption         =   "品名"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "幅宽"
         Caption         =   "幅宽"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "克重"
         Caption         =   "克重"
      EndProperty
      BeginProperty Field7 
         Precision       =   7
         Size            =   4
         Scale           =   0
         Type            =   4
         Name            =   "计划"
         Caption         =   "计划"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "色别"
         Caption         =   "色别"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   202
         Name            =   "备注"
         Caption         =   "备注"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "技要"
         Caption         =   "技要"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "日期"
         Caption         =   "日期"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   20
         Name            =   "序号"
         Caption         =   "序号"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "花型"
         Caption         =   "花型"
      EndProperty
      BeginProperty Field14 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "交期"
         Caption         =   "交期"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "负责"
         Caption         =   "负责"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   4
         Scale           =   0
         Type            =   202
         Name            =   "排布"
         Caption         =   "排布"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "发货"
         Caption         =   "发货"
      EndProperty
      BeginProperty Field18 
         Precision       =   7
         Size            =   4
         Scale           =   0
         Type            =   4
         Name            =   "匹数"
         Caption         =   "匹数"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "色名"
         Caption         =   "色名"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   202
         Name            =   "流程"
         Caption         =   "流程"
      EndProperty
      BeginProperty Field21 
         Precision       =   18
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "单价"
         Caption         =   "单价"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "色牢度"
         Caption         =   "色牢度"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "缩水率"
         Caption         =   "缩水率"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "扭度"
         Caption         =   "扭度"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "布纹"
         Caption         =   "布纹"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "成分"
         Caption         =   "成分"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "单号"
         ChildField      =   "单号"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset7 
      CommandName     =   "sczy_xhz"
      CommDispId      =   1040
      RsDispId        =   1053
      CommandText     =   "select distinct 单号,客户 from sczy_x where 单号=? "
      ActiveConnectionName=   "khjg"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "单号"
         Caption         =   "单号"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "客户"
         Caption         =   "客户"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   50
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset8 
      CommandName     =   "hzf"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "select 单号,客户,款号,品名,报价,色别,round(sum(计划),2) as 合计 from sczy_x where 单号=? group by 单号,客户,款号,品名,报价,色别"
      ActiveConnectionName=   "khjg"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "sczy_xhz"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "单号"
         Caption         =   "单号"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "客户"
         Caption         =   "客户"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "款号"
         Caption         =   "款号"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "品名"
         Caption         =   "品名"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "报价"
         Caption         =   "报价"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "色别"
         Caption         =   "色别"
      EndProperty
      BeginProperty Field7 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "合计"
         Caption         =   "合计"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   50
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "单号"
         ChildField      =   "Param1"
         ParentType      =   0
         ChildType       =   1
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset9 
      CommandName     =   "khhz"
      CommDispId      =   1054
      RsDispId        =   1058
      CommandText     =   $"GPXS.dsx":002A
      ActiveConnectionName=   "khjg"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "加工单位"
         Caption         =   "加工单位"
      EndProperty
      BeginProperty Field2 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "XI"
         Caption         =   "XI"
      EndProperty
      BeginProperty Field3 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "DA"
         Caption         =   "DA"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset10 
      CommandName     =   "khhzf"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "SELECT 加工单位,加工类别,SUM(数量) AS S,SUM(金额) AS E FROM JGMX WHERE 日期 BETWEEN ? AND ? GROUP BY 加工单位,加工类别"
      ActiveConnectionName=   "khjg"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "khhz"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "加工单位"
         Caption         =   "加工单位"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "加工类别"
         Caption         =   "加工类别"
      EndProperty
      BeginProperty Field3 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "S"
         Caption         =   "S"
      EndProperty
      BeginProperty Field4 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "E"
         Caption         =   "E"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   3
      BeginProperty Relation1 
         ParentField     =   "加工单位"
         ChildField      =   "加工单位"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "XI"
         ChildField      =   "Param1"
         ParentType      =   0
         ChildType       =   1
      EndProperty
      BeginProperty Relation3 
         ParentField     =   "DA"
         ChildField      =   "Param2"
         ParentType      =   0
         ChildType       =   1
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset11 
      CommandName     =   "mprkdy"
      CommDispId      =   1062
      RsDispId        =   1070
      CommandText     =   "select distinct 客户名称,日期,单据号 from ckgl  where 单据号=? "
      ActiveConnectionName=   "khjg"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "客户名称"
         Caption         =   "客户名称"
      EndProperty
      BeginProperty Field2 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "日期"
         Caption         =   "日期"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "单据号"
         Caption         =   "单据号"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   50
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset12 
      CommandName     =   "mprkfl"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "select * from ckgl "
      ActiveConnectionName=   "khjg"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "mprkdy"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   16
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "客户名称"
         Caption         =   "客户名称"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "布类"
         Caption         =   "布类"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "毛胚幅宽"
         Caption         =   "毛胚幅宽"
      EndProperty
      BeginProperty Field4 
         Precision       =   18
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "毛胚重量"
         Caption         =   "毛胚重量"
      EndProperty
      BeginProperty Field5 
         Precision       =   18
         Size            =   19
         Scale           =   1
         Type            =   131
         Name            =   "毛胚匹数"
         Caption         =   "毛胚匹数"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "和约号"
         Caption         =   "和约号"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "备注"
         Caption         =   "备注"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "IP"
         Caption         =   "IP"
      EndProperty
      BeginProperty Field9 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "日期"
         Caption         =   "日期"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "存放位置"
         Caption         =   "存放位置"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "负责人"
         Caption         =   "负责人"
      EndProperty
      BeginProperty Field12 
         Precision       =   7
         Size            =   4
         Scale           =   0
         Type            =   4
         Name            =   "实际投放量"
         Caption         =   "实际投放量"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "ny"
         Caption         =   "ny"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "dh"
         Caption         =   "dh"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "单据号"
         Caption         =   "单据号"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "订单品名"
         Caption         =   "订单品名"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   3
      BeginProperty Relation1 
         ParentField     =   "客户名称"
         ChildField      =   "客户名称"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "单据号"
         ChildField      =   "单据号"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation3 
         ParentField     =   "日期"
         ChildField      =   "日期"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "GPXS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
