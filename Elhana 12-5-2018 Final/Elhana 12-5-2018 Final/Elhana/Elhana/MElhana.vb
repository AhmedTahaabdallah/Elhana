Imports System.Data.OleDb
Module MElhana
    Public cn1 As New ADODB.Connection
    Public cn12 As New ADODB.Connection
    Public rs_U As New ADODB.Recordset
    Public rs_Cars As New ADODB.Recordset
    Public rs_Citys As New ADODB.Recordset
    Public rs_Customers As New ADODB.Recordset
    Public rs_Vendors As New ADODB.Recordset
    Public rs_Tender As New ADODB.Recordset
    Public rs_Store As New ADODB.Recordset
    Public rs_StoreCust As New ADODB.Recordset
    Public rs_StoreVend As New ADODB.Recordset
    Public rs_Emp As New ADODB.Recordset
    Public rs_EmpAttand As New ADODB.Recordset
    Public rs_SofUserName As New ADODB.Recordset
    Public rs As New ADODB.Recordset
    Public u_Id As Integer
    Public monthscloseduser_Id As Integer
    Public user_name As String
    Public gcon As New OleDbConnection
    Public sql_str As String
    Public con1 As String
    Public custid_cdid As Integer
    Public sid_Cust_Cdid As Integer
    Public vendid_cdid As Integer
    Public sid_Vend_Cdid As Integer

    Public rs_auth As New ADODB.Recordset
    Public auth_stor As Boolean
    Public auth_blance As Boolean
    Public auth_item As Boolean
    Public auth_addmonth As Boolean
    Public auth_useradd As Boolean
    Public auth_usershow As Boolean
    Public auth_usersearch As Boolean
    Public auth_useredit As Boolean
    Public auth_add As Boolean
    Public auth_show As Boolean
    Public auth_search As Boolean
    Public auth_edit As Boolean
    Public auth_delete As Boolean
    Public auth_report As Boolean
    Public Fetora_id As Integer
    Public Fetora_custid As Integer
    Public Fetora_monthid As Integer
    Public fetore_type As String

    'con1 = "Provider=SQLOLEDB.1;Password=123;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana;Data Source=AHMED1"
End Module
