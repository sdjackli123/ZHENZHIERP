Attribute VB_Name = "yhxx"
Public yhm As String        ''''''用户名
Public yhxx As String        ''''''用户信息
Public yhdm As String           '''''用户代码
Public qxsz(150) As String      ''' 权限数量
Public yhmk As String    '''用户模块
Public ljbl As Integer
Public cxlj, ljb As String
Public beizhu As Integer
Public gh As String
Public hg As String
Public zhci As Integer
Public TT(32) As String
Public asd As Integer ''''''锅数
Public bh As Integer ''''''''编号变量
Public g As Integer  ''''''''编号公量

Public GXBL As Integer       '工序变量
Public YGBL As Integer      '员工变量

Public cxtjsz(30) As String   ''''查询条件数组
Public KMBL, KMMC As Integer

Public Const sh As String = "10E7663411AE6DF3"  '''''加密狗id号
Public clbl As Integer '''''''''材料变量
Public rhlbl As Integer
Public ysbl As Integer  ''''''''编号公量

Public pxbl As String    '''''''排序变量
Public khyj As Integer   ''''''客户预警变量
Public tsgy As Integer   ''''''特殊工艺变量

Public bzdm As Integer   ''''''班组代码变量

Public gyhys As Integer '''''''''工艺化验室
Public xjbl As Integer '''''''询价变量
Public rjgxts As String  '''''软件更新提示

Public ghcx As Integer  '''''锅号补打变量
Public pmbl As Integer '''''窗体变量
Public wwdm As Integer   ''''委外代码
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal _
        Flags As Long) As Long  '''''大小写转换

Public jmg As String   '''加密狗变量

