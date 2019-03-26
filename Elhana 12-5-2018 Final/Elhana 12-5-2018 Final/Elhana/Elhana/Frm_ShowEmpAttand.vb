Imports ADODB
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Public Class Frm_ShowEmpAttand
    Public E_id As Integer
    Public month_id As Integer
    Public m_name As String
    Public gcon As New OleDbConnection
    Public sarly_day As Integer
    Public mainsalry As Integer
    Public yu As Integer
    Public rcon As SqlConnection
    Public rcomand As SqlCommand
    Public rada As SqlDataAdapter
    Public rdat As New DataSet
    'Public sql_str As String
    Public r_Ename As String
    Public r_monthname As String
    Public r_count_days As String
    Public r_Count_Attanteddays As String
    Public r_Count_Absentdays As String
    Public r_Total_Ancestor As String
    Public r_Total_Punash As String
    Public r_Total_sarly As String
    Public r_Total_Adational As String
    Public r_Salry As String
    Public r_NetSarly As String
    Public r_m_MonthName As String
    Public r_M_netSalry22 As String
    Public r_m_E_Count As String




    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            If ComboBox1.Text = "كل الحضور" Then
                ComboBox2.Visible = False
                ComboBox3.Visible = False
                ComboBox5.Visible = False
            End If
            If ComboBox1.Text = "أسم الموظف" Then
                ComboBox2.Visible = True
                ComboBox3.Visible = False
                ComboBox5.Visible = False
                ComboBox2.Text = ""
            End If
            If ComboBox1.Text = "الشهر بالتفصيل" Then
                ComboBox2.Visible = False
                ComboBox3.Visible = False
                ComboBox5.Visible = True
                ComboBox5.Text = ""
            End If
            If ComboBox1.Text = "أسم الموظف والشهر" Then
                ComboBox2.Visible = True
                ComboBox3.Visible = True
                ComboBox5.Visible = False
                ComboBox2.Text = ""
                ComboBox3.Text = ""
            End If
            If ComboBox1.Text = "الشهر النهائى" Then
                ComboBox2.Visible = False
                ComboBox3.Visible = False
                ComboBox5.Visible = True
                ComboBox5.Text = ""
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox2_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.DropDown
        Try
            If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
            Dim s As Boolean
            s = False
            Dim s1 As Boolean
            s1 = True
            rs_SofUserName.CursorLocation = CursorLocationEnum.adUseClient
            rs_SofUserName.Open("Select * From Employyes Where Emp_Flag='" & s & "' And Emp_Salry_State='" & s1 & "' ORDER BY Emp_Sorted ASC", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox2.Items.Clear()
            Do While Not rs_SofUserName.EOF
                ComboBox2.Items.Add(rs_SofUserName("Emp_Name").Value)
                rs_SofUserName.MoveNext()
            Loop
            rs_SofUserName.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            Dim s As Boolean
            s = False
            Dim s1 As Boolean
            s1 = True
            Dim e_name As String
            e_name = ComboBox2.SelectedItem.ToString()
            If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
            rs_SofUserName.CursorLocation = CursorLocationEnum.adUseClient
            rs_SofUserName.Open("Select * From Employyes Where Emp_Name='" & e_name & "' And Emp_Salry_State='" & s1 & "' And Emp_Flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            E_id = rs_SofUserName("Emp_Id").Value

            rs_SofUserName.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox5_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.DropDown
        Try
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.CursorLocation = CursorLocationEnum.adUseClient
            rs_Cars.Open("Select * From year_monthes Where Month_DeleFlag='" & s & "' ORDER BY Month_ID DESC", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox5.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox5.Items.Add(rs_Cars("Month_Name").Value)
                rs_Cars.MoveNext()
            Loop
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged
        Try
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = ComboBox5.SelectedItem.ToString()
            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.CursorLocation = CursorLocationEnum.adUseClient
            rs_Cars.Open("Select * From year_monthes Where Month_Name='" & u_name & "' And Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            month_id = rs_Cars("Month_ID").Value
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox3_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.DropDown
        Try
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.CursorLocation = CursorLocationEnum.adUseClient
            rs_Cars.Open("Select * From year_monthes Where Month_DeleFlag='" & s & "' ORDER BY Month_ID DESC", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox3.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox3.Items.Add(rs_Cars("Month_Name").Value)
                rs_Cars.MoveNext()
            Loop
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Try
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = ComboBox3.SelectedItem.ToString()
            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.CursorLocation = CursorLocationEnum.adUseClient
            rs_Cars.Open("Select * From year_monthes Where Month_Name='" & u_name & "' And Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            month_id = rs_Cars("Month_ID").Value
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If ComboBox1.Text = "كل الحضور" Then
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_all()
                Label15.Text = "عدد الأيام  :"
                Label17.Text = "عدد أيام الغياب :"
                Label2.Text = "عدد أيام الحضور :"
                Label1.Visible = False
                Label2.Visible = True
                Label14.Visible = True
                Label4.Visible = False
                Label5.Visible = False
                Label6.Visible = False
                Label9.Visible = False
                Label8.Visible = False
                Label11.Visible = False
                Label10.Visible = False
                Label12.Visible = False
                Label13.Visible = False
                Label18.Visible = False
                Label19.Visible = False
                Label20.Visible = False
                Label21.Visible = False
                Label22.Visible = False
                Label23.Visible = False
                Label24.Visible = False
                Label25.Visible = False
                Label28.Visible = False
                Label29.Visible = False
                Label30.Visible = False
                Label31.Visible = False
                Label32.Visible = False
                Label26.Visible = True
                Label27.Visible = True
                Button2.Visible = False
            End If
            If ComboBox1.Text = "أسم الموظف" Then
                If ComboBox2.Text = "" Then
                    MsgBox("من فضلك أختر أسم الموظف", MsgBoxStyle.Information, "تنبيه")
                    ComboBox2.Select()
                    Exit Sub
                End If
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_empname()
                Label15.Text = "عدد الأيام  :"
                Label2.Text = "عدد أيام الحضور :"
                Label17.Text = "عدد أيام الغياب :"
                Label9.Text = "أجمالى دقائق الخصم :"
                Label11.Text = "أجمالى دقائق الأضافى :"
                Label1.Visible = False
                Label2.Visible = True
                Label4.Visible = False
                Label14.Visible = True
                Label5.Visible = False
                Label6.Visible = False
                Label9.Visible = True
                Label8.Visible = True
                Label11.Visible = True
                Label10.Visible = True
                Label12.Visible = False
                Label13.Visible = False
                Label18.Visible = False
                Label19.Visible = False
                Label20.Visible = False
                Label21.Visible = False
                Label22.Visible = False
                Label23.Visible = False
                Label24.Visible = False
                Label25.Visible = False
                Label28.Visible = False
                Label29.Visible = False
                Label30.Visible = False
                Label31.Visible = False
                Label32.Visible = False
                Label26.Visible = True
                Label27.Visible = True
                Label17.Visible = True
                Label16.Visible = True
                Button2.Visible = False

            End If
            If ComboBox1.Text = "الشهر بالتفصيل" Then
                If ComboBox5.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox5.Select()
                    Exit Sub
                End If
                Me.Cursor = Cursors.WaitCursor
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_mname()
                Label14.Visible = True
                Label1.Visible = True
                Label2.Visible = True
                Label4.Visible = True
                Label5.Visible = True
                Label6.Visible = True
                Label9.Visible = True
                Label8.Visible = True
                Label11.Visible = True
                Label10.Visible = True
                Label12.Visible = True
                Label13.Visible = True
                Label18.Visible = True
                Label19.Visible = True
                Label20.Visible = True
                Label21.Visible = True
                Label22.Visible = True
                Label23.Visible = True
                Label24.Visible = True
                Label25.Visible = True
                Label28.Visible = True
                Label29.Visible = True
                Label30.Visible = True
                Label31.Visible = True
                Label32.Visible = True
                Label26.Visible = True
                Label27.Visible = True
                Label17.Visible = True
                Label16.Visible = True
                Button2.Visible = False
                Label4.Text = "إجمالى المرتبات :"
                Label13.Text = "صافى المرتبات :"
                Label19.Text = "إجمالى المرتبات الأساسية :"
                Label15.Text = "عدد الأيام  :"
                Label2.Text = ""
                Label2.Text = "عدد أيام الحضور :"
                Label17.Text = "عدد أيام الغياب :"
                Label11.Text = "إجمالى الخصم :"
                Label9.Text = "إجمالى السلفيات :"
                Label11.Text = "إجمالى الخصم :"
                Me.Cursor = Cursors.Default
            End If
            If ComboBox1.Text = "أسم الموظف والشهر" Then
                If ComboBox2.Text = "" Then
                    MsgBox("من فضلك أختر أسم الموظف", MsgBoxStyle.Information, "تنبيه")
                    ComboBox2.Select()
                    Exit Sub
                End If
                If ComboBox3.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox3.Select()
                    Exit Sub
                End If
                Me.Cursor = Cursors.WaitCursor
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_mname_e()
                Label2.Text = "عدد أيام الحضور :"
                Label17.Text = "عدد أيام الغياب :"
                Label4.Text = "إجمالى الراتب :"
                Label13.Text = "صافى المرتب :"
                Label19.Text = "الراتب الأساسى :"
                Label15.Text = "عدد الأيام  :"
                Label11.Text = "إجمالى الخصم :"
                Label9.Text = "إجمالى السلفيات :"
                Label11.Text = "إجمالى الخصم :"
                Label1.Visible = True
                Label2.Visible = True
                Label14.Visible = True
                Label4.Visible = True
                Label5.Visible = True
                Label6.Visible = True
                Label9.Visible = True
                Label8.Visible = True
                Label11.Visible = True
                Label10.Visible = True
                Label12.Visible = True
                Label13.Visible = True
                Label18.Visible = True
                Label19.Visible = True
                Label20.Visible = True
                Label21.Visible = True
                Label22.Visible = True
                Label23.Visible = True
                Label24.Visible = True
                Label25.Visible = True
                Label28.Visible = True
                Label29.Visible = True
                Label30.Visible = True
                Label31.Visible = True
                Label32.Visible = True
                Label26.Visible = True
                Label27.Visible = True
                Label17.Visible = True
                Label16.Visible = True


                If auth_report = True Then
                    yu = 1
                    r_Ename = ComboBox2.Text
                    r_monthname = ComboBox3.Text
                    r_count_days = Label3.Text
                    r_Count_Attanteddays = Label14.Text
                    r_Count_Absentdays = Label16.Text
                    r_Total_Ancestor = Label8.Text
                    r_Total_Punash = Label10.Text
                    r_Salry = Label18.Text
                    r_Total_sarly = Label1.Text
                    r_Total_Adational = Label5.Text
                    r_NetSarly = Label12.Text
                    Button2.Visible = True
                End If

                If u_Id = 1 Then
                    yu = 1
                    r_Ename = ComboBox2.Text
                    r_monthname = ComboBox3.Text
                    r_count_days = Label3.Text
                    r_Count_Attanteddays = Label14.Text
                    r_Count_Absentdays = Label16.Text
                    r_Total_Ancestor = Label8.Text
                    r_Total_Punash = Label10.Text
                    r_Salry = Label18.Text
                    r_Total_sarly = Label1.Text
                    r_Total_Adational = Label5.Text
                    r_NetSarly = Label12.Text
                    Button2.Visible = True
                End If
                Me.Cursor = Cursors.Default
               
            End If

            If ComboBox1.Text = "الشهر النهائى" Then
                If ComboBox5.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox5.Select()
                    Exit Sub
                End If
                Me.Cursor = Cursors.WaitCursor
                If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
                Dim s As Boolean
                s = False
                Dim s1 As Boolean
                s1 = True
                rs_SofUserName.CursorLocation = CursorLocationEnum.adUseClient
                rs_SofUserName.Open("Select * From Employyes Where Emp_Salry_State='" & s1 & "' And Emp_Flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Dim d1 As Integer

                Do While Not rs_SofUserName.EOF
                    d1 = rs_SofUserName("Emp_Id").Value
                    sarly_day = rs_SofUserName("Emp_Salary_Day").Value
                    mainsalry = rs_SofUserName("Emp_Salary").Value
                    get_Final_Month(d1, month_id, "عرض الشهر النهائى")

                    rs_SofUserName.MoveNext()
                Loop
                DataGridView1.DataSource = ""
                DataGridView1.Refresh()
                get_Finalmname()

                Label11.Text = "أجمالى صافى المرتبات :"
                Label17.Text = "إجمالى المرتبات الأساسية :"

                Label15.Text = "عدد الموظفين :"
                Label1.Visible = False
                Label2.Visible = False
                Label4.Visible = False
                Label14.Visible = False
                Label5.Visible = False
                Label6.Visible = False
                Label9.Visible = False
                Label8.Visible = False
                Label11.Visible = True
                Label10.Visible = True
                Label12.Visible = False
                Label13.Visible = False
                Label18.Visible = False
                Label19.Visible = False
                Label20.Visible = True
                Label21.Visible = False
                Label22.Visible = False
                Label23.Visible = False
                Label24.Visible = False
                Label25.Visible = False
                Label28.Visible = False
                Label29.Visible = False
                Label30.Visible = False
                Label31.Visible = False
                Label32.Visible = False
                Label26.Visible = False
                Label27.Visible = False
                Label17.Visible = True
                Label16.Visible = True
                If auth_report = True Then
                    yu = 2
                    r_m_MonthName = ComboBox5.Text
                    r_m_E_Count = Label3.Text
                    r_M_netSalry22 = Label10.Text
                    Button2.Visible = True
                End If

                If u_Id = 1 Then
                    yu = 2
                    r_m_MonthName = ComboBox5.Text
                    r_m_E_Count = Label3.Text
                    r_M_netSalry22 = Label10.Text
                    Button2.Visible = True
                End If
                Me.Cursor = Cursors.Default
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub dgv()
        Try
            DataGridView1.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(0).Width = 85
            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).Width = 140
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).Width = 70
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).Width = 70
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 100
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
           
            DataGridView1.Columns(5).Width = 60
            DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).Width = 60
            DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).Width = 90
            DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(8).Width = 60
            DataGridView1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
           
            DataGridView1.Columns(9).Width = 160
            DataGridView1.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            If u_Id = 1 Then
                DataGridView1.Columns(10).Width = 150
                DataGridView1.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(11).Width = 160
                DataGridView1.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub dgv_Final_Month()
        Try
            DataGridView1.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(0).Width = 200
            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).Width = 100
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).Width = 120
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).Width = 120
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).Width = 100
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).Width = 100
            DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).Width = 120
            DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).Width = 120
            DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(8).Width = 100
            DataGridView1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(9).Width = 100
            DataGridView1.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub get_all()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim ssd As Boolean
            ssd = False
            Dim ssd1 As Boolean
            ssd1 = True
            Dim stt As String

            If u_Id = 1 Then
                stt = "Select att.EmpP_Id as [رقم الحضور], emm.Emp_Name as [أسم الموظف], att.Emp_Day as [اليوم], mnn.Month_Name as [شهر], att.Emp_Date as [التاريخ], att.Emp_Punash as [خصم], att.Emp_Late as [تأخير], att.Emp_Additional as [وقت أضافى], att.Emp_Ancestor as [سلف], att.Emp_Note as [ملاحظات], (Select uss.User_Name From Users as uss Where uss.User_ID=att.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=att.User_ID_edit) as [مستخدم التعديل] FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = emm.Emp_Id And att.Month_ID=mnn.Month_ID And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "' ORDER BY att.Emp_Date DESC"
            Else
                stt = "Select att.EmpP_Id as [رقم الحضور], emm.Emp_Name as [أسم الموظف], att.Emp_Day as [اليوم], mnn.Month_Name as [شهر], att.Emp_Date as [التاريخ], att.Emp_Punash as [خصم], att.Emp_Late as [تأخير], att.Emp_Additional as [وقت أضافى], att.Emp_Ancestor as [سلف], att.Emp_Note as [ملاحظات] FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = emm.Emp_Id And att.Month_ID=mnn.Month_ID And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "' ORDER BY att.Emp_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(att.EmpP_Id) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = emm.Emp_Id And att.Month_ID=mnn.Month_ID And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(att.EmpP_Id) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = emm.Emp_Id And att.Month_ID=mnn.Month_ID And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label14.Text = 0
            Else
                Label14.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(att.EmpP_Id) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = emm.Emp_Id And att.Month_ID=mnn.Month_ID And att.hodor ='" & ssd & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label16.Text = 0
            Else
                Label16.Text = rs_StoreCust("c_No").Value
            End If
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub get_empname()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim s As Boolean
            s = False
            Dim ssd As Boolean
            ssd = False
            Dim ssd1 As Boolean
            ssd1 = True
            Dim stt As String
            'stt = "SELECT st.EmpP_Id as [رقم الحضور], cu.Emp_Name as [أسم الموظف], st.Emp_Day as [اليوم], st.m_name as [شهر], st.Emp_Date as [التاريخ], st.Emp_In1 as [دخول 1], st.Emp_Out1 as [خروج 1], st.Emp_In2 as [دخول 2], st.Emp_Out2 as [خروج 2], st.Emp_Punash as [عقاب], st.Emp_Late as [تأخير], st.Emp_Additional as [وقت أضافى], st.Emp_Ancestor as [سلف], st.Emp_Total as [اجمالى عدد ساعات العمل] FROM Employyes_Presence as st, Employyes as cu where st.Emp_Id = " & E_id & " And cu.Emp_Id = " & E_id & " And st.Emp_flag =" & s & " And cu.Emp_Flag =" & s & " ORDER BY st.Emp_Date DESC"
            If u_Id = 1 Then
                stt = "Select att.EmpP_Id as [رقم الحضور], emm.Emp_Name as [أسم الموظف], att.Emp_Day as [اليوم], mnn.Month_Name as [شهر], att.Emp_Date as [التاريخ], att.Emp_Punashmenets as [خصم د], att.Emp_Additionalmenets as [أضافى د], att.Emp_Late as [تأخير], att.Emp_Punash as [خصم], att.Emp_Additional as [وقت أضافى], att.Emp_Ancestor as [سلف], att.Emp_Note as [ملاحظات], (Select uss.User_Name From Users as uss Where uss.User_ID=att.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=att.User_ID_edit) as [مستخدم التعديل] FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And att.Month_ID=mnn.Month_ID And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "' ORDER BY att.Emp_Date DESC"
            Else
                stt = "Select att.EmpP_Id as [رقم الحضور], emm.Emp_Name as [أسم الموظف], att.Emp_Day as [اليوم], mnn.Month_Name as [شهر], att.Emp_Date as [التاريخ], att.Emp_Punashmenets as [خصم د], att.Emp_Additionalmenets as [أضافى د], att.Emp_Late as [تأخير], att.Emp_Punash as [خصم], att.Emp_Additional as [وقت أضافى], att.Emp_Ancestor as [سلف], att.Emp_Note as [ملاحظات] FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And att.Month_ID=mnn.Month_ID And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "' ORDER BY att.Emp_Date DESC"
            End If
            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(att.EmpP_Id) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And att.Month_ID=mnn.Month_ID And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(att.EmpP_Id) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And att.Month_ID=mnn.Month_ID And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label14.Text = 0
            Else
                Label14.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(att.EmpP_Id) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And att.Month_ID=mnn.Month_ID And att.hodor ='" & ssd & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label16.Text = 0
            Else
                Label16.Text = rs_StoreCust("c_No").Value
            End If
            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(att.Emp_Punashmenets) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And att.Month_ID=mnn.Month_ID And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label8.Text = 0
            Else
                Label8.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(att.Emp_Additionalmenets) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And att.Month_ID=mnn.Month_ID And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label10.Text = 0
            Else
                Label10.Text = rs_StoreCust("c_No").Value
            End If
            rs_StoreCust.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub get_mname()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim ssd As Boolean
            ssd = False
            Dim ssd1 As Boolean
            ssd1 = True
            Dim stt As String
            'stt = "SELECT st.EmpP_Id as [رقم الحضور], cu.Emp_Name as [أسم الموظف], st.Emp_Day as [اليوم], st.m_name as [شهر], st.Emp_Date as [التاريخ], st.Emp_In1 as [دخول 1], st.Emp_Out1 as [خروج 1], st.Emp_In2 as [دخول 2], st.Emp_Out2 as [خروج 2], st.Emp_Punash as [عقاب], st.Emp_Late as [تأخير], st.Emp_Additional as [وقت أضافى], st.Emp_Ancestor as [سلف], st.Emp_Total as [اجمالى عدد ساعات العمل] FROM Employyes_Presence as st, Employyes as cu where st.Emp_Id = cu.Emp_Id And st.m_name='" & m_name & "' And st.Emp_flag =" & s & " And cu.Emp_Flag =" & s & " ORDER BY st.Emp_Date DESC"
            If u_Id = 1 Then
                stt = "Select att.EmpP_Id as [رقم الحضور], emm.Emp_Name as [أسم الموظف], att.Emp_Day as [اليوم], mnn.Month_Name as [شهر], att.Emp_Date as [التاريخ], att.Emp_Punash as [خصم], att.Emp_Late as [تأخير], att.Emp_Additional as [وقت أضافى], att.Emp_Ancestor as [سلف], att.Emp_Note as [ملاحظات], (Select uss.User_Name From Users as uss Where uss.User_ID=att.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=att.User_ID_edit) as [مستخدم التعديل] FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = emm.Emp_Id And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "' ORDER BY att.Emp_Date DESC"
            Else
                stt = "Select att.EmpP_Id as [رقم الحضور], emm.Emp_Name as [أسم الموظف], att.Emp_Day as [اليوم], mnn.Month_Name as [شهر], att.Emp_Date as [التاريخ], att.Emp_Punash as [خصم], att.Emp_Late as [تأخير], att.Emp_Additional as [وقت أضافى], att.Emp_Ancestor as [سلف], att.Emp_Note as [ملاحظات] FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = emm.Emp_Id And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "' ORDER BY att.Emp_Date DESC"
            End If

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(att.EmpP_Id) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = emm.Emp_Id And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(att.EmpP_Id) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = emm.Emp_Id And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label14.Text = 0
            Else
                Label14.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(att.EmpP_Id) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = emm.Emp_Id And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And att.hodor ='" & ssd & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label16.Text = 0
            Else
                Label16.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(emm.Emp_Salary_Day) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = emm.Emp_Id And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label1.Text = 0
            Else
                Dim yy As Integer
                yy = rs_StoreCust("c_No").Value
                Label1.Text = yy
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(att.Emp_Additional) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = emm.Emp_Id And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label5.Text = 0
            Else
                Label5.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(att.Emp_Ancestor) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = emm.Emp_Id And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label8.Text = 0
            Else
                Label8.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(att.Emp_Punash) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = emm.Emp_Id And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label10.Text = 0
            Else
                Label10.Text = rs_StoreCust("c_No").Value
            End If
            rs_StoreCust.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(Emp_Salary) as c_No FROM Employyes where Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label18.Text = 0
            Else
                Label18.Text = rs_StoreCust("c_No").Value
            End If
            rs_StoreCust.Close()

            Dim tt As Integer
            tt = Val(Label10.Text) + Val(Label8.Text)
            Dim kk As Double
            kk = Val(Label1.Text) + Val(Label5.Text)
            Label12.Text = kk - tt
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub get_mname_e()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim ssd As Boolean
            ssd = False
            Dim ssd1 As Boolean
            ssd1 = True
            Dim stt As String
            If u_Id = 1 Then
                stt = "Select att.EmpP_Id as [رقم الحضور], emm.Emp_Name as [أسم الموظف], att.Emp_Day as [اليوم], mnn.Month_Name as [شهر], att.Emp_Date as [التاريخ], att.Emp_Punash as [خصم], att.Emp_Late as [تأخير], att.Emp_Additional as [وقت أضافى], att.Emp_Ancestor as [سلف], att.Emp_Note as [ملاحظات], (Select uss.User_Name From Users as uss Where uss.User_ID=att.User_ID) as [مستخدم الأضافة], (Select uss.User_Name From Users as uss Where uss.User_ID=att.User_ID_edit) as [مستخدم التعديل] FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "' ORDER BY att.Emp_Date DESC"
            Else
                stt = "Select att.EmpP_Id as [رقم الحضور], emm.Emp_Name as [أسم الموظف], att.Emp_Day as [اليوم], mnn.Month_Name as [شهر], att.Emp_Date as [التاريخ], att.Emp_Punash as [خصم], att.Emp_Late as [تأخير], att.Emp_Additional as [وقت أضافى], att.Emp_Ancestor as [سلف], att.Emp_Note as [ملاحظات] FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "' ORDER BY att.Emp_Date DESC"
            End If
            'stt = "SELECT st.EmpP_Id as [رقم الحضور], cu.Emp_Name as [أسم الموظف], st.Emp_Day as [اليوم], st.m_name as [شهر], st.Emp_Date as [التاريخ], st.Emp_In1 as [دخول 1], st.Emp_Out1 as [خروج 1], st.Emp_In2 as [دخول 2], st.Emp_Out2 as [خروج 2], st.Emp_Punash as [عقاب], st.Emp_Late as [تأخير], st.Emp_Additional as [وقت أضافى], st.Emp_Ancestor as [سلف], st.Emp_Total as [اجمالى عدد ساعات العمل] FROM Employyes_Presence as st, Employyes as cu where st.Emp_Id = " & E_id & " And cu.Emp_Id = " & E_id & " And st.m_name='" & m_name & "' And st.Emp_flag =" & s & " And cu.Emp_Flag =" & s & " ORDER BY st.Emp_Date DESC"

            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv()
            Dim nonotatt As Integer
            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(att.EmpP_Id) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(att.EmpP_Id) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label14.Text = 0
            Else
                nonotatt = rs_StoreCust("c_No").Value
                Label14.Text = nonotatt
            End If
            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(att.EmpP_Id) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And att.hodor ='" & ssd & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label16.Text = 0
            Else
                Label16.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT emm.Emp_Salary_Day as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label1.Text = 0
            Else
                Dim ff As Double
                ff = rs_StoreCust("c_No").Value
                Dim rt As Integer
                rt = Val(ff) * Val(nonotatt)
                Label1.Text = rt
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(att.Emp_Additional) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label5.Text = 0
            Else

                Label5.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(att.Emp_Ancestor) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label8.Text = 0
            Else

                Label8.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(att.Emp_Punash) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & E_id & " And emm.Emp_Id = " & E_id & " And mnn.Month_ID=" & month_id & " And att.Month_ID=" & month_id & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label10.Text = 0
            Else

                Label10.Text = rs_StoreCust("c_No").Value
            End If
            rs_StoreCust.Close()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(Emp_Salary) as c_No FROM Employyes where Emp_Id = " & E_id & " And Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label18.Text = 0
            Else
                Label18.Text = rs_StoreCust("c_No").Value
            End If
            rs_StoreCust.Close()

            Dim tt As Integer
            tt = Val(Label10.Text) + Val(Label8.Text)
            Dim kk As Integer
            kk = Val(Label1.Text) + Val(Label5.Text)
            Label12.Text = kk - tt
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub get_Finalmname()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter
            Dim ssd As Boolean
            ssd = False
            Dim ssd1 As Boolean
            ssd1 = True
            Dim stt As String
            'stt = "SELECT st.EmpP_Id as [رقم الحضور], cu.Emp_Name as [أسم الموظف], st.Emp_Day as [اليوم], st.m_name as [شهر], st.Emp_Date as [التاريخ], st.Emp_In1 as [دخول 1], st.Emp_Out1 as [خروج 1], st.Emp_In2 as [دخول 2], st.Emp_Out2 as [خروج 2], st.Emp_Punash as [عقاب], st.Emp_Late as [تأخير], st.Emp_Additional as [وقت أضافى], st.Emp_Ancestor as [سلف], st.Emp_Total as [اجمالى عدد ساعات العمل] FROM Employyes_Presence as st, Employyes as cu where st.Emp_Id = cu.Emp_Id And st.m_name='" & m_name & "' And st.Emp_flag =" & s & " And cu.Emp_Flag =" & s & " ORDER BY st.Emp_Date DESC"
            stt = "Select emm.Emp_Name as [أسم الموظف], mnn.Month_Name as [شهر], sm.Main_Salry as [الراتب الأساسى], sm.count_NotAbsent as [عدد أيام الحضور], sm.count_Absent as [عدد أيام الغياب], sm.total_Salry as [إجمالى الراتب], sm.total_Additional as [إجمالى الأضافى], sm.total_Ancestor as [إجمالى السلفيات], sm.total_Punash as [إجمالى الخصم], sm.net_Salry as [صافى الراتب] FROM Employyes as emm, year_monthes as mnn, Emp_SalralesWithMonth as sm where sm.Emp_Id = emm.Emp_Id And mnn.Month_ID=" & month_id & " And sm.Month_ID=" & month_id & " And mnn.Month_DeleFlag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "' And emm.Emp_Salry_State ='" & ssd1 & "' ORDER BY emm.Emp_Sorted ASC"
            da = New OleDbDataAdapter(stt, gcon)

            da.Fill(dt)
            DataGridView1.DataSource = dt.DefaultView
            gcon.Close()
            DataGridView1.Refresh()
            dgv_Final_Month()

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(sm.ID) as c_No FROM Employyes as emm, year_monthes as mnn, Emp_SalralesWithMonth as sm where sm.Emp_Id = emm.Emp_Id And mnn.Month_ID=" & month_id & " And sm.Month_ID=" & month_id & " And mnn.Month_DeleFlag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "' And emm.Emp_Salry_State ='" & ssd1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label3.Text = 0
            Else
                Label3.Text = rs_StoreCust("c_No").Value
            End If


            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(sm.Main_Salry) as c_No FROM Employyes as emm, year_monthes as mnn, Emp_SalralesWithMonth as sm where sm.Emp_Id = emm.Emp_Id And mnn.Month_ID=" & month_id & " And sm.Month_ID=" & month_id & " And mnn.Month_DeleFlag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "' And emm.Emp_Salry_State ='" & ssd1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label16.Text = 0
            Else
                Label16.Text = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(sm.net_Salry) as c_No FROM Employyes as emm, year_monthes as mnn, Emp_SalralesWithMonth as sm where sm.Emp_Id = emm.Emp_Id And mnn.Month_ID=" & month_id & " And sm.Month_ID=" & month_id & " And mnn.Month_DeleFlag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "' And emm.Emp_Salry_State ='" & ssd1 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                Label10.Text = 0
            Else
                Label10.Text = rs_StoreCust("c_No").Value
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Frm_ShowEmpAttand_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Try
            Me.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub Frm_ShowEmpAttand_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'gcon.ConnectionString = "Provider=SQLOLEDB.1;Password=141751523;Persist Security Info=True;User ID=sa;Initial Catalog=Tiba;Data Source=ADMIN-PC\SQLEXPRESS"
        Try
            auth_report = rs_auth("User_empatt_report").Value
            'sql_str = "Data Source=ADMIN-PC\AHMED1;Password=141751523;Persist Security Info=True;User ID=sa;Initial Catalog=Tiba"
            If gcon.State = 1 Then gcon.Close()
            gcon.ConnectionString = "Provider=SQLOLEDB.1;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana;Data Source=PC-PC\SQLEXPRESS"
            gcon.Open()
            'gcon.Open()
            get_all()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    Private Sub get_Final_Month(ByVal emp As Integer, ByVal mont As Integer, ByVal st As String)
        Try
            Dim ssd As Boolean
            ssd = False
            Dim ssd1 As Boolean
            ssd1 = True
            Dim emp_absent As Integer
            Dim emp_notabsent As Integer
            Dim total_sarly As Double
            Dim total_aditonal As Double
            Dim total_Ancestor As Double
            Dim total_Punash As Double

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(att.EmpP_Id) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & emp & " And emm.Emp_Id = " & emp & " And mnn.Month_ID=" & mont & " And att.Month_ID=" & mont & " And att.hodor ='" & ssd & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                emp_absent = 0
            Else
                emp_absent = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT COUNT(att.EmpP_Id) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & emp & " And emm.Emp_Id = " & emp & " And mnn.Month_ID=" & mont & " And att.Month_ID=" & mont & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                emp_notabsent = 0
            Else
                emp_notabsent = rs_StoreCust("c_No").Value
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.CursorLocation = CursorLocationEnum.adUseClient
                rs_StoreVend.Open("SELECT * From Employyes Where Emp_Id = " & emp & " And Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                Dim s_a As Integer
                Dim m_c As Integer
                Dim dy As Double
                s_a = rs_StoreVend("Emp_Salary").Value
                m_c = rs_StoreVend("Emp_MonthCount").Value
                dy = s_a / m_c
                total_sarly = emp_notabsent * dy
            End If

            'If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            'rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            'rs_StoreCust.Open("SELECT SUM(emm.Emp_Salary_Day) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & emp & " And emm.Emp_Id = " & emp & " And mnn.Month_ID=" & mont & " And att.Month_ID=" & mont & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            'If IsDBNull(rs_StoreCust("c_No").Value) Then
            '    total_sarly = 0
            'Else
            '    total_sarly = rs_StoreCust("c_No").Value
            'End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(att.Emp_Additional) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & emp & " And emm.Emp_Id = " & emp & " And mnn.Month_ID=" & mont & " And att.Month_ID=" & mont & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                total_aditonal = 0
            Else
                total_aditonal = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(att.Emp_Ancestor) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & emp & " And emm.Emp_Id = " & emp & " And mnn.Month_ID=" & mont & " And att.Month_ID=" & mont & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                total_Ancestor = 0
            Else
                total_Ancestor = rs_StoreCust("c_No").Value
            End If

            If rs_StoreCust.State = 1 Then rs_StoreCust.Close()
            rs_StoreCust.CursorLocation = CursorLocationEnum.adUseClient
            rs_StoreCust.Open("SELECT SUM(att.Emp_Punash) as c_No FROM Employyes_Presence as att, Employyes as emm, year_monthes as mnn where att.Emp_Id = " & emp & " And emm.Emp_Id = " & emp & " And mnn.Month_ID=" & mont & " And att.Month_ID=" & mont & " And att.hodor ='" & ssd1 & "' And mnn.Month_DeleFlag ='" & ssd & "' And att.Emp_flag ='" & ssd & "' And emm.Emp_Flag ='" & ssd & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If IsDBNull(rs_StoreCust("c_No").Value) Then
                total_Punash = 0
            Else
                total_Punash = rs_StoreCust("c_No").Value
            End If
            rs_StoreCust.Close()

            Dim tt As Integer
            tt = total_sarly + total_aditonal
            Dim kk As Integer
            kk = total_Ancestor + total_Punash
            Dim net_Salry As Integer
            net_Salry = tt - kk
            Dim total_sarly1 As Integer
            Dim total_aditonal1 As Integer
            Dim total_Ancestor1 As Integer
            Dim total_Punash1 As Integer
            total_sarly1 = total_sarly
            total_aditonal1 = total_aditonal
            total_Ancestor1 = total_Ancestor
            total_Punash1 = total_Punash

            Dim s As Boolean
            s = False
            If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
            rs_EmpAttand.CursorLocation = CursorLocationEnum.adUseClient
            rs_EmpAttand.Open("Select * From Emp_SalralesWithMonth Where Emp_Id=" & emp & " And Month_ID=" & mont & "", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If rs_EmpAttand.EOF Or rs_EmpAttand.BOF Then

                'Dim u_name As Integer
                'Dim u_name1 As Integer
                'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                'rs_StoreVend.CursorLocation = CursorLocationEnum.adUseClient
                'rs_StoreVend.Open("Select * From Emp_SalralesWithMonth Where ID= (SELECT MAX(ID)  FROM Emp_SalralesWithMonth)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                'u_name = rs_StoreVend("ID").Value
                'u_name1 = u_name + 1

                If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
                rs_EmpAttand.Open("Emp_SalralesWithMonth", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_EmpAttand.AddNew()
                'rs_EmpAttand("ID").Value = u_name1
                rs_EmpAttand("Month_ID").Value = mont
                rs_EmpAttand("User_ID").Value = u_Id
                rs_EmpAttand("Emp_Id").Value = emp
                rs_EmpAttand("e_State").Value = st
                rs_EmpAttand("count_Absent").Value = emp_absent
                rs_EmpAttand("count_NotAbsent").Value = emp_notabsent
                rs_EmpAttand("Main_Salry").Value = mainsalry
                rs_EmpAttand("total_Salry").Value = total_sarly1
                rs_EmpAttand("total_Additional").Value = total_aditonal1
                rs_EmpAttand("total_Ancestor").Value = total_Ancestor1
                rs_EmpAttand("total_Punash").Value = total_Punash1
                rs_EmpAttand("net_Salry").Value = net_Salry
                rs_EmpAttand("Emp_Salary_Day").Value = sarly_day
                rs_EmpAttand.Update()
                rs_EmpAttand.Close()

            Else
                rs_EmpAttand("Month_ID").Value = mont
                rs_EmpAttand("User_ID").Value = u_Id
                rs_EmpAttand("Emp_Id").Value = emp
                rs_EmpAttand("e_State").Value = st
                rs_EmpAttand("count_Absent").Value = emp_absent
                rs_EmpAttand("count_NotAbsent").Value = emp_notabsent
                rs_EmpAttand("Main_Salry").Value = mainsalry
                rs_EmpAttand("total_Salry").Value = total_sarly1
                rs_EmpAttand("total_Additional").Value = total_aditonal1
                rs_EmpAttand("total_Ancestor").Value = total_Ancestor1
                rs_EmpAttand("total_Punash").Value = total_Punash1
                rs_EmpAttand("net_Salry").Value = net_Salry
                rs_EmpAttand("Emp_Salary_Day").Value = sarly_day
                rs_EmpAttand.Update()
                rs_EmpAttand.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Try
            If yu = 1 Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New rpt_EmpAttand
                Dim frm_rpt As New Frm_rpt_EmpAttand
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                
                rcomand = New SqlCommand("Select EmpP_Id,Emp_Date,Emp_Total,Emp_Punash,Emp_Ancestor,Emp_Additional,Emp_Note FROM Employyes_Presence where Emp_Id = " & E_id & " And Month_ID=" & month_id & "  And Emp_flag ='" & ssd & "' ORDER BY Emp_Date DESC", rcon)
                rada = New SqlDataAdapter(rcomand)
                rada.Fill(rdat, "Employyes_Presence")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("EName", r_Ename)
                rpt.SetParameterValue("MonthName", r_monthname)
                rpt.SetParameterValue("MonthNameun", user_name)
                rpt.SetParameterValue("count_days", r_count_days)
                rpt.SetParameterValue("Count_Attanteddays", r_Count_Attanteddays)
                rpt.SetParameterValue("Count_Absentdays", r_Count_Absentdays)
                rpt.SetParameterValue("Total_Ancestor", r_Total_Ancestor)
                rpt.SetParameterValue("Total_Punash", r_Total_Punash)
                rpt.SetParameterValue("Total_sarly", r_Total_sarly)
                rpt.SetParameterValue("Total_Adational", r_Total_Adational)
                rpt.SetParameterValue("Salry", r_Salry)
                rpt.SetParameterValue("NetSarly", r_NetSarly)

                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة حضور موظف"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If
            If yu = 2 Then
                Me.Cursor = Cursors.WaitCursor
                Dim rpt As New rpt_Final_Month
                Dim frm_rpt As New Frm_Rpt_FinalMonth
                rdat.Clear()
                rcon = New SqlConnection("Data Source=PC-PC\SQLEXPRESS;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana")
                rcon.Open()
                Dim ssd As Boolean
                ssd = False
                Dim ssd1 As Boolean
                ssd1 = True
                rcomand = New SqlCommand("Select (Select em.Emp_Name From Employyes as em Where em.Emp_Id = sm.Emp_Id) as Emp_id, sm.Main_Salry, sm.count_NotAbsent, sm.count_Absent, sm.total_Salry, sm.total_Additional, sm.total_Ancestor, sm.total_Punash, sm.net_Salry From Emp_SalralesWithMonth as sm , Employyes as emm Where sm.Month_ID=" & month_id & " And emm.Emp_Id = sm.Emp_Id And emm.Emp_Salry_State ='" & ssd1 & "'", rcon)
                rada = New SqlDataAdapter(rcomand)
                rada.Fill(rdat, "Emp_SalralesWithMonth")
                rpt.SetDataSource(rdat)
                rpt.SetParameterValue("MonthName", r_m_MonthName)
                rpt.SetParameterValue("MonthNamee", r_m_E_Count)
                rpt.SetParameterValue("MonthNamen", r_M_netSalry22)
                rpt.SetParameterValue("MonthNameu", user_name)
                'rpt.SetParameterValue("My Parameter1", r_m_E_Count)
                'rpt.SetParameterValue("My Parameter2", r_M_netSalry22)

                'rpt.SetParameterValue("M_final", r_M_netSalry)

                frm_rpt.CrystalReportViewer1.ReportSource = rpt
                frm_rpt.CrystalReportViewer1.Refresh()
                frm_rpt.CrystalReportViewer1.Dock = DockStyle.Fill
                Dim frm As New Form
                With frm
                    .Controls.Add(frm_rpt.CrystalReportViewer1)
                    .Text = "طباعة الشهر النهائى"
                    .WindowState = FormWindowState.Maximized
                    .ShowDialog()
                End With

                Me.Cursor = Cursors.Default
                Button2.Visible = False
            End If
            'Frm_rpt_EmpAttand.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub
    
End Class