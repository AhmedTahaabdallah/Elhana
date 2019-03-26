Public Class Frm_Money_Blance

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Me.Dispose()
    End Sub

    Private Sub Frm_Money_Blance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'Dim s As Boolean
            's = False
            'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            'rs_StoreVend.Open("SELECT Sum(Vend_Blance_Dein) as c_No FROM Vendores where Vend_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'If IsDBNull(rs_StoreVend("c_No").Value) Then
            '    Label2.Text = 0
            'Else
            '    Label2.Text = rs_StoreVend("c_No").Value
            'End If

            'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            'rs_StoreVend.Open("SELECT Sum(Cust_Blance_Medin) as c_No FROM Customers where Cust_DelFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'If IsDBNull(rs_StoreVend("c_No").Value) Then
            '    Label3.Text = 0
            'Else
            '    Label3.Text = rs_StoreVend("c_No").Value
            'End If
            Dim s As Boolean
            s = False
            Dim li As Integer
            Dim al As Integer
            Dim fg As String
            fg = "defulit"
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT Sum(Vend_Blance_Dein) as c_No FROM Vendores where Vend_DelFlag='" & s & "' And Vend_visable='" & s & "' And Vend_Blance_d='" & fg & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                Label2.Text = 0
                al = 0
            Else
                Label2.Text = rs_StoreVend("c_No").Value
                al = rs_StoreVend("c_No").Value
            End If
            fg = "notdefulit"
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT Sum(Vend_Blance_Dein) as c_No FROM Vendores where Vend_DelFlag='" & s & "' And Vend_visable='" & s & "' And Vend_Blance_d='" & fg & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                Label3.Text = 0
                li = 0
            Else
                Label3.Text = rs_StoreVend("c_No").Value
                li = rs_StoreVend("c_No").Value
            End If
            Label9.Text = al - li

            fg = "defulit"
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT Sum(Cust_Blance_Medin) as c_No FROM Customers where Cust_DelFlag='" & s & "' And Cust_visable='" & s & "' And Cust_Blance_d='" & fg & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                Label19.Text = 0
                li = 0
            Else
                Label19.Text = rs_StoreVend("c_No").Value
                li = rs_StoreVend("c_No").Value
            End If
            fg = "notdefulit"
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("SELECT Sum(Cust_Blance_Medin) as c_No FROM Customers where Cust_DelFlag='" & s & "' And Cust_visable='" & s & "' And Cust_Blance_d='" & fg & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If IsDBNull(rs_StoreVend("c_No").Value) Then
                Label16.Text = 0
                al = 0
            Else
                Label16.Text = rs_StoreVend("c_No").Value
                al = rs_StoreVend("c_No").Value
            End If
            Label14.Text = li - al

            Dim sMoney_Blance As Boolean
            sMoney_Blance = False
            If rs_Vendors.State = 1 Then rs_Vendors.Close()
            rs_Vendors.Open("Select * From Money_Blance Where DelFlag='" & sMoney_Blance & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_Vendors.EOF Or rs_Vendors.BOF Then
                Exit Sub
            Else
                If rs_Store.State = 1 Then rs_Store.Close()
                'rs_Store.Open("SELECT TOP 1 * FROM Store_Quentity Order By Store_ID DESC", cn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                'rs_Store.Open("Select MAX(Store_ID) AS la,Store_Quntity as rss From Store_Quentity Where Store_DelFlag=" & s & "", cn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Store.Open("SELECT * FROM Money_Blance Where DelFlag='" & sMoney_Blance & "' And  ID = (SELECT MAx(ID)  FROM Money_Blance)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                Label6.Text = rs_Store("Blance_Total").Value
                'Label8.Text = Label2.Text - Label7.Text

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
End Class