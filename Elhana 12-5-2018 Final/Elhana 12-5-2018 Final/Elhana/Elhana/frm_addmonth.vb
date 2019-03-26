Imports ADODB
Public Class frm_addmonth

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        txt_Name.Text = ""
        btn_add.Enabled = False
        btn_save.Enabled = True
        txt_Name.Select()
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        If txt_Name.Text = "" Then
            MsgBox("من فضلك أدخل أسم الشهر")
            Exit Sub
        End If
        Try
            Dim ss As Boolean
            ss = False
            If rs_Citys.State = 1 Then rs_Citys.Close()
            rs_Citys.Open("Select * From year_monthes Where Month_Name='" & txt_Name.Text & "' and Month_DeleFlag='" & ss & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_Citys.EOF Or rs_Citys.BOF Then
                'Dim u_name As Integer
                'Dim u_name1 As Integer
                'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                'rs_StoreVend.CursorLocation = CursorLocationEnum.adUseClient
                'rs_StoreVend.Open("Select * From year_monthes Where Month_ID= (SELECT MAX(Month_ID)  FROM year_monthes)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                'u_name = rs_StoreVend("Month_ID").Value
                'u_name1 = u_name + 1

                If rs_Citys.State = 1 Then rs_Citys.Close()
                rs_Citys.Open("year_monthes", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_Citys.AddNew()
                'rs_Citys("Month_ID").Value = u_name1
                rs_Citys("Month_Name").Value = txt_Name.Text
                rs_Citys("User_ID").Value = u_Id
                rs_Citys("Month_DeleFlag").Value = ss
                rs_Citys.Update()
                MsgBox("تم حفظ الشهر بنجاح", MsgBoxStyle.Information, "حفظ بيانات الشهر")
                txt_Name.Text = ""
                btn_add.Enabled = True
                btn_save.Enabled = False
                rs_Citys.Close()
            Else
                MsgBox("أسم الشهر موجود من قبل أدخل أسم أخر", MsgBoxStyle.Information, "تحذير")
                txt_Name.Text = ""
                txt_Name.Select()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Dispose()
    End Sub

    Private Sub frm_addmonth_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class