Imports ADODB
Imports System.Data.OleDb
Public Class Frm_EmpAttand
    Public E_id As Integer
    Public Ep_id As Integer
    Public m_name As Integer
    Public d_name As String
    Public month_id As Integer
    Public gcon As New OleDbConnection
    Public mainsalry As Integer
    Public sarly_day As Double
    Public sarly_hour As Double
    Public sarly_minute As Double
    Public dy As String
    Public dm As String
    Public dd As String
    Public dstri As String

    Private Sub get_mname(ByVal emp As Integer, ByVal mont As Integer, ByVal st As String)
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
                Dim dy As Double
                Dim s_a As Integer
                Dim m_c As Integer
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

            Dim tt As Double
            tt = total_sarly + total_aditonal
            Dim kk As Double
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
                rs_EmpAttand("EmpP_Id").Value = Ep_id
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
                rs_EmpAttand("EmpP_Id").Value = Ep_id
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
    Private Sub ComboBox2_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.DropDown
        Try
            If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
            Dim s As Boolean
            s = False
            Dim s1 As Boolean
            s1 = True
            rs_SofUserName.CursorLocation = CursorLocationEnum.adUseClient
            rs_SofUserName.Open("Select * From Employyes Where Emp_Flag='" & s & "' And Emp_State='" & s1 & "' ORDER BY Emp_Sorted ASC", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

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
        TextBox2.Text = 0
        TextBox4.Text = 0
        TextBox5.Text = 0
        Try
            Dim s As Boolean
            s = False
            Dim s1 As Boolean
            s1 = True
            Dim shiftstate As Boolean
            Dim shiftno As Integer
            Dim km As Integer
            km = 60
            Dim e_name As String
            e_name = ComboBox2.SelectedItem.ToString()
            If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
            rs_SofUserName.CursorLocation = CursorLocationEnum.adUseClient
            rs_SofUserName.Open("Select * From Employyes Where Emp_Name='" & e_name & "' And Emp_State='" & s1 & "' And Emp_Flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            E_id = rs_SofUserName("Emp_Id").Value
            Label18.Text = rs_SofUserName("Emp_Salary_Day").Value
            sarly_day = rs_SofUserName("Emp_Salary_Day").Value
            mainsalry = rs_SofUserName("Emp_Salary").Value
            shiftstate = rs_SofUserName("Emp_Shift_State").Value

            If shiftstate = True Then
                shiftno = rs_SofUserName("Emp_Shift_No").Value
                sarly_hour = Val(sarly_day) / Val(shiftno)
                sarly_minute = Val(sarly_hour) / Val(km)
                Label21.Text = sarly_hour
                Label23.Text = sarly_minute
                TextBox2.Enabled = True
                TextBox4.Enabled = True
            End If
            If shiftstate = False Then

                sarly_hour = 0
                sarly_minute = 0
                Label21.Text = sarly_hour
                Label23.Text = sarly_minute
                TextBox2.Enabled = False
                TextBox4.Enabled = False
            End If

            rs_SofUserName.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        TextBox2.Text = 0
        TextBox4.Text = 0
        TextBox5.Text = 0
        ComboBox2.Enabled = True
        ComboBox1.Enabled = True
        DateTimePicker1.Enabled = True
        RadioButton1.Checked = True
        RadioButton2.Checked = False
        TextBox2.Enabled = False
        TextBox4.Enabled = False
        Label18.Text = 0
        Label21.Text = 0
        Label23.Text = 0
        Label11.Text = ""
        TextBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox1.Text = ""
        ComboBox3.Text = ""
        TextBox3.Text = ""
        Txt_at1.Text = ""
        Txt_at2.Text = ""
        Txt_go1.Text = ""
        Txt_go2.Text = ""
        Txt_late.Text = 0
        Txt_parttimemonets.Text = 0
        Txt_punchm.Text = 0
        Txt_parttime.Text = 0
        Txt_parttime.Text = 0
        txt_Ancestor.Text = 0
        Txt_punch.Text = 0
        'Txt_punch.Text = 0
        'Txt_punchm.Text = 0
        Txt_total.Text = 0
        'Txt_late.Text = 0
        btn_All.Enabled = False
        btn_etite.Enabled = False
        btn_delete.Enabled = False
        btn_add.Enabled = False
        GroupBox3.Enabled = False
        btn_save.Enabled = True
        ComboBox2.Select()
        Try
            'Dim s As Boolean
            's = False
            'Dim u_name As Integer
            'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            'rs_StoreVend.CursorLocation = CursorLocationEnum.adUseClient
            'rs_StoreVend.Open("Select * From Employyes_Presence Where EmpP_Id= (SELECT MAX(EmpP_Id)  FROM Employyes_Presence)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)


            'u_name = rs_StoreVend("EmpP_Id").Value
            'Label11.Text = u_name + 1
            'Ep_id = u_name + 1
            If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
            rs_StoreVend.Open("Select Count(EmpP_Id) as df From Employyes_Presence", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)


            If rs_StoreVend("df").Value = 0 Then
                Label11.Text = 1
            Else
                Dim u_name As Integer
                If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                rs_StoreVend.Open("Select * From Employyes_Presence Where EmpP_Id= (SELECT MAX(EmpP_Id)  FROM Employyes_Presence)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                u_name = rs_StoreVend("EmpP_Id").Value
                Label11.Text = u_name + 1
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        TextBox2.Text = 0
        TextBox4.Text = 0
        TextBox5.Text = 0



        Try
            dy = DateTimePicker1.Value.Year
            dm = DateTimePicker1.Value.Month
            dd = DateTimePicker1.Value.Day
            dstri = dy + "-" + dm + "-" + dd
            Dim uk As Integer
            uk = 2
            If RadioButton2.Checked = True Then

                If ComboBox2.Text = "" Then
                    MsgBox("من فضلك أختر أسم الموظف")
                    ComboBox2.Select()
                    Exit Sub
                End If
                If ComboBox3.Text = "" Then
                    MsgBox("من فضلك أختر اليوم", MsgBoxStyle.Information, "تنبيه")
                    ComboBox3.Select()
                    Exit Sub
                End If
                If ComboBox1.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox1.Select()
                    Exit Sub
                End If

                Dim s As Boolean
                s = False
                If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
                rs_EmpAttand.CursorLocation = CursorLocationEnum.adUseClient
                rs_EmpAttand.Open("Select * From Employyes_Presence Where Emp_Id=" & E_id & " And Emp_Datestri='" & dstri & "' And Emp_flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                If rs_EmpAttand.EOF Or rs_EmpAttand.BOF Then


                    If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
                    rs_EmpAttand.Open("Employyes_Presence", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    rs_EmpAttand.AddNew()
                    'rs_EmpAttand("EmpP_Id").Value = Ep_id
                    rs_EmpAttand("Emp_Day").Value = d_name
                    rs_EmpAttand("Emp_Date").Value = DateTimePicker1.Value.ToShortDateString()
                    rs_EmpAttand("Emp_Datestri").Value = dstri
                    rs_EmpAttand("Emp_flag").Value = s
                    rs_EmpAttand("User_ID").Value = u_Id
                    rs_EmpAttand("User_ID_edit").Value = uk
                    rs_EmpAttand("User_ID_delete").Value = uk
                    rs_EmpAttand("Emp_Id").Value = E_id
                    rs_EmpAttand("Emp_Salary_Day").Value = 0
                    rs_EmpAttand("hodor").Value = s
                    rs_EmpAttand("Month_ID").Value = month_id
                    Dim sss As String
                    sss = "غياب"
                    rs_EmpAttand("Emp_In1").Value = sss
                    rs_EmpAttand("Emp_Out1").Value = sss
                    rs_EmpAttand("Emp_In2").Value = sss
                    rs_EmpAttand("Emp_Out2").Value = sss
                    rs_EmpAttand("Emp_Total").Value = 0
                    rs_EmpAttand("Emp_Note").Value = "غياب"
                    rs_EmpAttand.Update()
                    get_mname(E_id, month_id, "حفظ غياب")
                    MsgBox("تم حفظ الحضور بنجاح", MsgBoxStyle.Information, "حفظ بيانات الحضور")
                    Label11.Text = 0
                    TextBox2.Enabled = False
                    TextBox4.Enabled = False
                    Label18.Text = 0
                    Label21.Text = 0
                    Label23.Text = 0
                    ComboBox1.Text = ""
                    ComboBox2.Text = ""
                    ComboBox3.Text = ""
                    TextBox3.Text = ""
                    Txt_at1.Text = ""
                    Txt_at2.Text = ""
                    Txt_go1.Text = ""
                    Txt_go2.Text = ""
                    TextBox1.Text = ""
                    Txt_late.Text = 0
                    txt_Ancestor.Text = 0
                    Txt_parttime.Text = 0
                    Txt_punch.Text = 0
                    Txt_parttimemonets.Text = 0
                    Txt_punchm.Text = 0
                    Txt_total.Text = ""
                    btn_All.Enabled = True
                    btn_etite.Enabled = True
                    btn_delete.Enabled = True
                    btn_add.Enabled = True
                    GroupBox3.Enabled = True
                    btn_save.Enabled = False


                Else
                    MsgBox("تم تسجيل هذالموظف لهذا اليوم من قبل ", MsgBoxStyle.Information, "تحذير")
                    ComboBox2.Text = ""
                    ComboBox2.Select()
                End If
            End If


            If RadioButton1.Checked = True Then

                If ComboBox2.Text = "" Then
                    MsgBox("من فضلك أختر أسم الموظف")
                    ComboBox2.Select()
                    Exit Sub
                End If

                If ComboBox3.Text = "" Then
                    MsgBox("من فضلك أختر اليوم", MsgBoxStyle.Information, "تنبيه")
                    ComboBox3.Select()
                    Exit Sub
                End If
                'If Txt_at1.Text = "" Then
                '    MsgBox("من فضلك أدخل وقت الحضور", MsgBoxStyle.Information, "تنبيه")
                '    Txt_at1.Select()
                '    Exit Sub
                'End If
                'If Txt_go2.Text = "" Then
                '    MsgBox("من فضلك أدخل وقت الأنصراف", MsgBoxStyle.Information, "تنبيه")
                '    Txt_go2.Select()
                '    Exit Sub
                'End If
                'If Txt_total.Text = "" Then
                '    MsgBox("من فضلك أدخل أجمالى عدد ساعات العمل لهذا اليوم ", MsgBoxStyle.Information, "تنبيه")
                '    Txt_total.Select()
                '    Exit Sub
                'End If
                If ComboBox1.Text = "" Then
                    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                    ComboBox1.Select()
                    Exit Sub
                End If
                If TextBox1.Text = "" Then
                    TextBox1.Text = "لا يوجد ملاحظات"
                End If
                If Txt_parttime.Text = "" Then
                    Txt_parttime.Text = 0
                End If
                If Txt_parttimemonets.Text = "" Then
                    Txt_parttimemonets.Text = 0
                End If
                If Txt_punch.Text = "" Then
                    Txt_punch.Text = 0
                End If
                If Txt_punchm.Text = "" Then
                    Txt_punchm.Text = 0
                End If
                Dim s As Boolean
                s = False
                If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
                rs_EmpAttand.CursorLocation = CursorLocationEnum.adUseClient
                rs_EmpAttand.Open("Select * From Employyes_Presence Where Emp_Id=" & E_id & " And Emp_Datestri='" & dstri & "' And Emp_flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                If rs_EmpAttand.EOF Or rs_EmpAttand.BOF Then


                    If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
                    rs_EmpAttand.Open("Employyes_Presence", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    rs_EmpAttand.AddNew()
                    'rs_EmpAttand("EmpP_Id").Value = Ep_id
                    rs_EmpAttand("Emp_Day").Value = d_name
                    rs_EmpAttand("Emp_Date").Value = DateTimePicker1.Value.ToShortDateString()
                    rs_EmpAttand("Emp_Datestri").Value = dstri
                    rs_EmpAttand("Emp_In1").Value = 0
                    rs_EmpAttand("Emp_Out1").Value = 0
                    rs_EmpAttand("Emp_In2").Value = 0
                    rs_EmpAttand("Emp_Out2").Value = 0
                    rs_EmpAttand("Emp_Punash").Value = Txt_punch.Text
                    rs_EmpAttand("Emp_Late").Value = Txt_late.Text
                    rs_EmpAttand("Emp_Additional").Value = Txt_parttime.Text
                    rs_EmpAttand("Emp_Ancestor").Value = txt_Ancestor.Text
                    rs_EmpAttand("Emp_flag").Value = s
                    rs_EmpAttand("User_ID").Value = u_Id
                    rs_EmpAttand("User_ID_edit").Value = uk
                    rs_EmpAttand("User_ID_delete").Value = uk
                    rs_EmpAttand("Emp_Id").Value = E_id
                    rs_EmpAttand("Emp_Total").Value = 0
                    rs_EmpAttand("Month_ID").Value = month_id
                    rs_EmpAttand("Emp_Salary_Day").Value = Label18.Text
                    rs_EmpAttand("Emp_Note").Value = TextBox1.Text
                    rs_EmpAttand("Emp_Additionalmenets").Value = Txt_parttimemonets.Text
                    rs_EmpAttand("Emp_Punashmenets").Value = Txt_punchm.Text
                    s = True
                    rs_EmpAttand("hodor").Value = s
                    rs_EmpAttand.Update()
                    get_mname(E_id, month_id, "حفظ حضور")
                    MsgBox("تم حفظ الحضور بنجاح", MsgBoxStyle.Information, "حفظ بيانات الحضور")
                    Label11.Text = 0
                    TextBox2.Enabled = False
                    TextBox4.Enabled = False
                    Label18.Text = 0
                    Label21.Text = 0
                    Label23.Text = 0
                    ComboBox1.Text = ""
                    ComboBox2.Text = ""
                    TextBox1.Text = ""
                    ComboBox3.Text = ""
                    TextBox3.Text = ""
                    Txt_at1.Text = ""
                    Txt_at2.Text = ""
                    Txt_go1.Text = ""
                    Txt_go2.Text = ""
                    Txt_late.Text = ""
                    Txt_parttime.Text = ""
                    Txt_punch.Text = ""
                    Txt_total.Text = ""
                    Txt_parttimemonets.Text = ""
                    Txt_punchm.Text = ""
                    btn_All.Enabled = True
                    btn_etite.Enabled = True
                    btn_delete.Enabled = True
                    btn_add.Enabled = True
                    GroupBox3.Enabled = True
                    btn_save.Enabled = False


                Else
                    MsgBox("تم تسجيل هذالموظف لهذا اليوم من قبل ", MsgBoxStyle.Information, "تحذير")
                    ComboBox2.Text = ""
                    ComboBox2.Select()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox1_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.DropDown
        Try
            If rs_Cars.State = 1 Then rs_Cars.Close()
            Dim s As Boolean
            s = False
            rs_Cars.CursorLocation = CursorLocationEnum.adUseClient
            rs_Cars.Open("Select * From year_monthes Where Month_DeleFlag='" & s & "' ORDER BY Month_ID DESC", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            ComboBox1.Items.Clear()
            Do While Not rs_Cars.EOF
                ComboBox1.Items.Add(rs_Cars("Month_Name").Value)
                rs_Cars.MoveNext()
            Loop
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        Try
            Dim s As Boolean
            s = False
            Dim u_name As String
            u_name = ComboBox1.SelectedItem.ToString()
            If rs_Cars.State = 1 Then rs_Cars.Close()
            rs_Cars.CursorLocation = CursorLocationEnum.adUseClient
            rs_Cars.Open("Select * From year_monthes Where Month_Name='" & u_name & "' And Month_DeleFlag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            month_id = rs_Cars("Month_ID").Value
            rs_Cars.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        d_name = ComboBox3.Text
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox2.Text = 0
        TextBox4.Text = 0
        TextBox5.Text = 0
        frm_addmonth.ShowDialog()
    End Sub

    Private Sub btn_etite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_etite.Click
        TextBox2.Text = 0
        TextBox4.Text = 0
        TextBox5.Text = 0
        If Label11.Text = "0" Or Label11.Text = "" Then
            MsgBox("من فضلك أدخل رقم الحضور", MsgBoxStyle.Information, "تنبيه")
            TextBox3.Select()
            Exit Sub
        End If

        Try
            dy = DateTimePicker1.Value.Year
            dm = DateTimePicker1.Value.Month
            dd = DateTimePicker1.Value.Day
            dstri = dy + "-" + dm + "-" + dd
            Dim uk As Integer
            uk = 2
            If RadioButton2.Checked = True Then
                Dim g As String
                g = MsgBox("هل تريد تعديل هذا الغياب ؟", MsgBoxStyle.YesNo, "تأكيد الحذف")

                If g = vbYes Then
                    If ComboBox3.Text = "" Then
                        MsgBox("من فضلك أختر اليوم", MsgBoxStyle.Information, "تنبيه")
                        ComboBox3.Select()
                        Exit Sub
                    End If
                    Dim s As Boolean
                    s = False
                    If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
                    rs_EmpAttand.CursorLocation = CursorLocationEnum.adUseClient
                    rs_EmpAttand.Open("Select * From Employyes_Presence Where EmpP_Id=" & Ep_id & " And Emp_flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    rs_EmpAttand("Emp_Day").Value = d_name
                    rs_EmpAttand("hodor").Value = s
                    rs_EmpAttand("Emp_Date").Value = DateTimePicker1.Value.ToShortDateString()
                    rs_EmpAttand("Emp_Datestri").Value = dstri
                    Dim sss As String
                    sss = "غياب"
                    rs_EmpAttand("Emp_In1").Value = sss
                    rs_EmpAttand("Emp_Out1").Value = sss
                    rs_EmpAttand("Emp_In2").Value = sss
                    rs_EmpAttand("Emp_Out2").Value = sss
                    rs_EmpAttand("User_ID_edit").Value = u_Id
                    'rs_EmpAttand("Emp_Id").Value = E_id
                    rs_EmpAttand("Month_ID").Value = month_id
                    rs_EmpAttand("Emp_Salary_Day").Value = 0
                    rs_EmpAttand("Emp_Total").Value = 0
                    rs_EmpAttand("Emp_Note").Value = "غياب"
                    rs_EmpAttand.Update()
                    get_mname(E_id, month_id, "تعديل غياب")
                    MsgBox("تم تعديل الحضور بنجاح", MsgBoxStyle.Information, "تعديل بيانات الحضور")
                    Label11.Text = 0
                    ComboBox2.Text = ""
                    ComboBox1.Text = ""
                    ComboBox3.Text = ""
                    TextBox3.Text = ""
                    Txt_at1.Text = ""
                    Txt_at2.Text = ""
                    Txt_go1.Text = ""
                    Txt_parttimemonets.Text = 0
                    Txt_punchm.Text = 0
                    Txt_go2.Text = ""
                    Txt_late.Text = 0
                    Txt_parttime.Text = 0
                    Txt_punch.Text = 0
                    txt_Ancestor.Text = 0
                    Txt_total.Text = 0
                    TextBox2.Enabled = False
                    TextBox4.Enabled = False
                    Label18.Text = 0
                    Label21.Text = 0
                    Label23.Text = 0
                    DateTimePicker1.Enabled = True
                    ComboBox2.Enabled = True
                    ComboBox1.Enabled = True
                    ComboBox3.Enabled = True


                End If
            End If


            If RadioButton1.Checked = True Then



                If Label11.Text = "0" Or Label11.Text = "" Then
                    MsgBox("من فضلك أدخل رقم الاذن", MsgBoxStyle.Information, "تنبيه")
                    TextBox3.Select()
                    Exit Sub
                End If
                'If ComboBox2.Text = "" Then
                '    MsgBox("من فضلك أختر أسم الموظف")
                '    ComboBox2.Select()
                '    Exit Sub
                'End If

                If ComboBox3.Text = "" Then
                    MsgBox("من فضلك أختر اليوم", MsgBoxStyle.Information, "تنبيه")
                    ComboBox3.Select()
                    Exit Sub
                End If
                'If Txt_at1.Text = "" Then
                '    MsgBox("من فضلك أدخل وقت الحضور", MsgBoxStyle.Information, "تنبيه")
                '    Txt_at1.Select()
                '    Exit Sub
                'End If
                'If Txt_go2.Text = "" Then
                '    MsgBox("من فضلك أدخل وقت الأنصراف", MsgBoxStyle.Information, "تنبيه")
                '    Txt_go2.Select()
                '    Exit Sub
                'End If
                'If Txt_total.Text = "" Then
                '    MsgBox("من فضلك أدخل أجمالى عدد ساعات العمل لهذا اليوم ", MsgBoxStyle.Information, "تنبيه")
                '    Txt_total.Select()
                '    Exit Sub
                'End If
                'If ComboBox1.Text = "" Then
                '    MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
                '    ComboBox1.Select()
                '    Exit Sub
                'End If
                If TextBox1.Text = "" Then
                    TextBox1.Text = "لا يوجد ملاحظات"
                End If
                If Txt_parttime.Text = "" Then
                    Txt_parttime.Text = 0
                End If
                If Txt_parttimemonets.Text = "" Then
                    Txt_parttimemonets.Text = 0
                End If
                If Txt_punch.Text = "" Then
                    Txt_punch.Text = 0
                End If
                If Txt_punchm.Text = "" Then
                    Txt_punchm.Text = 0
                End If
                If Txt_late.Text = "" Then
                    Txt_late.Text = 0
                End If
                Dim g As String
                g = MsgBox("هل تريد تعديل هذا الحضور ؟", MsgBoxStyle.YesNo, "تأكيد الحذف")

                If g = vbYes Then
                    Dim s As Boolean
                    s = False
                    If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
                    rs_EmpAttand.CursorLocation = CursorLocationEnum.adUseClient
                    rs_EmpAttand.Open("Select * From Employyes_Presence Where EmpP_Id=" & Ep_id & " And Emp_flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    rs_EmpAttand("Emp_Day").Value = d_name
                    rs_EmpAttand("Emp_Date").Value = DateTimePicker1.Value.ToShortDateString()
                    rs_EmpAttand("Emp_Datestri").Value = dstri
                    rs_EmpAttand("Emp_In1").Value = 0
                    rs_EmpAttand("Emp_Out1").Value = 0
                    rs_EmpAttand("Emp_In2").Value = 0
                    rs_EmpAttand("Emp_Out2").Value = 0
                    rs_EmpAttand("Emp_Punash").Value = Txt_punch.Text
                    rs_EmpAttand("Emp_Late").Value = Txt_late.Text
                    rs_EmpAttand("Emp_Additional").Value = Txt_parttime.Text
                    rs_EmpAttand("Emp_Ancestor").Value = txt_Ancestor.Text
                    rs_EmpAttand("User_ID_edit").Value = u_Id
                    'rs_EmpAttand("Emp_Id").Value = E_id
                    rs_EmpAttand("Emp_Total").Value = 0
                    rs_EmpAttand("Month_ID").Value = month_id
                    rs_EmpAttand("Emp_Note").Value = TextBox1.Text
                    rs_EmpAttand("Emp_Additionalmenets").Value = Txt_parttimemonets.Text
                    rs_EmpAttand("Emp_Punashmenets").Value = Txt_punchm.Text
                    s = True
                    rs_EmpAttand("hodor").Value = s
                    rs_EmpAttand("Emp_Salary_Day").Value = sarly_day
                    rs_EmpAttand.Update()
                    get_mname(E_id, month_id, "تعديل حضور")
                    MsgBox("تم تعديل الحضور بنجاح", MsgBoxStyle.Information, "تعديل بيانات الحضور")
                    Label11.Text = 0
                    TextBox2.Enabled = False
                    TextBox4.Enabled = False
                    Label18.Text = 0
                    Label21.Text = 0
                    Label23.Text = 0
                    ComboBox2.Text = ""
                    ComboBox1.Text = ""
                    ComboBox3.Text = ""
                    TextBox3.Text = ""
                    Txt_at1.Text = ""
                    Txt_at2.Text = ""
                    Txt_go1.Text = ""
                    Txt_go2.Text = ""
                    Txt_late.Text = ""
                    Txt_parttime.Text = ""
                    Txt_punch.Text = ""
                    Txt_parttimemonets.Text = ""
                    Txt_punchm.Text = ""
                    Txt_total.Text = ""
                    TextBox1.Text = ""
                    DateTimePicker1.Enabled = True
                    ComboBox2.Enabled = True
                    ComboBox1.Enabled = True
                    ComboBox3.Enabled = True

                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox2.Text = 0
        TextBox4.Text = 0
        TextBox5.Text = 0
        Label18.Text = ""
        Label11.Text = ""
        ComboBox2.Text = ""
        ComboBox1.Text = ""
        ComboBox3.Text = ""
        Txt_at1.Text = ""
        Txt_at2.Text = ""
        Txt_go1.Text = ""
        Txt_go2.Text = ""
        Txt_late.Text = ""
        Txt_parttime.Text = ""
        Txt_parttime.Text = 0
        Txt_punch.Text = ""
        Txt_punch.Text = 0
        Txt_total.Text = ""
        If TextBox3.Text = "0" Or TextBox3.Text = "" Then
            MsgBox("من فضلك أدخل رقم الاذن", MsgBoxStyle.Information, "تنبيه")
            TextBox3.Select()
            Exit Sub
        End If

        Try
            Dim shiftstate As Boolean
            Dim shiftno As Integer
            Dim km As Integer
            km = 60
            Dim s As Boolean
            s = False
            Dim trs As Integer
            If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
            rs_EmpAttand.CursorLocation = CursorLocationEnum.adUseClient
            rs_EmpAttand.Open("Select * From Employyes_Presence Where EmpP_Id=" & TextBox3.Text & " And Emp_flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If rs_EmpAttand.EOF Or rs_EmpAttand.BOF Then
                MsgBox("رقم الأذن غير موجود", MsgBoxStyle.Information, "تنبيه")
                TextBox3.Select()
                Exit Sub
            End If
            Dim s15 As Boolean
            s15 = False
            Dim gh5 As Integer
            gh5 = rs_EmpAttand("Month_ID").Value
            If rs_U.State = 1 Then rs_U.Close()
            rs_U.CursorLocation = CursorLocationEnum.adUseClient
            rs_U.Open("Select * From year_monthes Where Month_ID=" & gh5 & " And Month_DeleFlag='" & s15 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ComboBox1.Text = rs_U("Month_Name").Value
            month_id = gh5

            Dim s155 As Boolean
            s155 = False
            Dim gh55 As Integer
            gh55 = rs_EmpAttand("Emp_Id").Value
            If rs_U.State = 1 Then rs_U.Close()
            rs_U.CursorLocation = CursorLocationEnum.adUseClient
            rs_U.Open("Select * From Employyes Where Emp_Id=" & gh55 & " And Emp_Flag='" & s155 & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ComboBox2.Text = rs_U("Emp_Name").Value
            mainsalry = rs_U("Emp_Salary").Value
            sarly_day = rs_U("Emp_Salary_Day").Value
            E_id = rs_U("Emp_Id").Value
            trs = rs_EmpAttand("Emp_Id").Value

            shiftstate = rs_U("Emp_Shift_State").Value

            If shiftstate = True Then
                shiftno = rs_U("Emp_Shift_No").Value
                sarly_hour = Val(sarly_day) / Val(shiftno)
                sarly_minute = Val(sarly_hour) / Val(km)
                Label21.Text = sarly_hour
                Label23.Text = sarly_minute
                TextBox2.Enabled = True
                TextBox4.Enabled = True
            End If
            If shiftstate = False Then

                sarly_hour = 0
                sarly_minute = 0
                Label21.Text = sarly_hour
                Label23.Text = sarly_minute
                TextBox2.Enabled = False
                TextBox4.Enabled = False
            End If

            Ep_id = rs_EmpAttand("EmpP_Id").Value
            Label11.Text = rs_EmpAttand("EmpP_Id").Value
            Label18.Text = sarly_day
            ComboBox3.Text = rs_EmpAttand("Emp_Day").Value
            DateTimePicker1.Value = rs_EmpAttand("Emp_Date").Value

            Dim yu As Boolean
            yu = rs_EmpAttand("hodor").Value
            If yu = True Then
                RadioButton1.Checked = True
                Txt_at1.Text = rs_EmpAttand("Emp_In1").Value
                Txt_go1.Text = rs_EmpAttand("Emp_Out1").Value
                Txt_at2.Text = rs_EmpAttand("Emp_In2").Value
                Txt_go2.Text = rs_EmpAttand("Emp_Out2").Value
                Txt_punch.Text = rs_EmpAttand("Emp_Punash").Value
                Txt_late.Text = rs_EmpAttand("Emp_Late").Value
                Txt_parttime.Text = rs_EmpAttand("Emp_Additional").Value
                txt_Ancestor.Text = rs_EmpAttand("Emp_Ancestor").Value
                TextBox1.Text = rs_EmpAttand("Emp_Note").Value
                Txt_total.Text = rs_EmpAttand("Emp_Total").Value
                Txt_parttimemonets.Text = rs_EmpAttand("Emp_Additionalmenets").Value
                Txt_punchm.Text = rs_EmpAttand("Emp_Punashmenets").Value

            End If

            If yu = False Then
                RadioButton2.Checked = True
            End If
            rs_U.Close()
            rs_EmpAttand.Close()

            'DateTimePicker1.Enabled = False
            ComboBox1.Enabled = False
            ComboBox2.Enabled = False
            'ComboBox3.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        TextBox2.Text = 0
        TextBox4.Text = 0
        TextBox5.Text = 0
        If Label11.Text = "0" Or Label11.Text = "" Then
            MsgBox("من فضلك أدخل رقم الحضور", MsgBoxStyle.Information, "تنبيه")
            TextBox3.Select()
            Exit Sub
        End If
        Try
            Dim g As String
            g = MsgBox("هل تريد حذف هذا الحضور ؟", MsgBoxStyle.YesNo, "تأكيد الحذف")

            If g = vbYes Then
                Dim u As Boolean
                u = True
                Dim s As Boolean
                s = False
                If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
                rs_EmpAttand.CursorLocation = CursorLocationEnum.adUseClient
                rs_EmpAttand.Open("Select * From Employyes_Presence Where EmpP_Id=" & Ep_id & " And Emp_flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_EmpAttand("Emp_flag").Value = u
                rs_EmpAttand("User_ID_delete").Value = u_Id
                rs_EmpAttand.Update()
                get_mname(E_id, month_id, "حذف")
                MsgBox("تم حذف هذا الحضور بنجاح", MsgBoxStyle.Information, "حذف بيانات الحضور")
                Label11.Text = 0
                TextBox2.Enabled = False
                TextBox4.Enabled = False
                Label18.Text = 0
                Label21.Text = 0
                Label23.Text = 0
                ComboBox1.Text = ""
                ComboBox2.Text = ""
                ComboBox3.Text = ""
                TextBox3.Text = ""
                Txt_at1.Text = ""
                Txt_at2.Text = ""
                Txt_go1.Text = ""
                TextBox1.Text = ""
                Txt_go2.Text = ""
                Txt_late.Text = 0
                txt_Ancestor.Text = 0
                Txt_parttime.Text = 0
                Txt_punch.Text = 0
                Txt_total.Text = 0
                Txt_parttimemonets.Text = 0
                Txt_punchm.Text = 0
                DateTimePicker1.Enabled = True
                ComboBox2.Enabled = True
                ComboBox1.Enabled = True
                ComboBox3.Enabled = True

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub btn_All_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_All.Click
        TextBox2.Text = 0
        TextBox4.Text = 0
        TextBox5.Text = 0
        Frm_ShowEmpAttand.Show()
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        TextBox2.Text = 0
        TextBox4.Text = 0
        TextBox5.Text = 0
        Me.Dispose()
    End Sub

    Private Sub Frm_EmpAttand_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim ut As Boolean
            ut = False
            auth_add = rs_auth("User_empatt_add").Value
            auth_show = rs_auth("User_empatt_show").Value
            auth_search = rs_auth("User_empatt_search").Value
            auth_edit = rs_auth("User_empatt_edit").Value
            auth_delete = rs_auth("User_empatt_delete").Value
            auth_addmonth = rs_auth("User_addmonth").Value
            If auth_add = ut Then
                btn_add.Visible = False
                btn_save.Visible = False
            Else
                btn_add.Visible = True
                btn_save.Visible = True
            End If
            If auth_show = ut Then
                btn_All.Visible = False
            Else
                btn_All.Visible = True
            End If
            If auth_search = ut Then
                GroupBox3.Visible = False
            Else
                GroupBox3.Visible = True
            End If
            If auth_edit = ut Then
                btn_etite.Visible = False
            Else
                btn_etite.Visible = True
            End If
            If auth_delete = ut Then
                btn_delete.Visible = False
            Else
                btn_delete.Visible = True
            End If
            If auth_addmonth = ut Then
                Button2.Visible = False
            Else
                Button2.Visible = True
            End If
            If u_Id = 1 Then
                btn_add.Visible = True
                btn_save.Visible = True
                btn_All.Visible = True
                GroupBox3.Visible = True
                btn_etite.Visible = True
                btn_delete.Visible = True
                Button2.Visible = True
            End If
            TextBox2.Text = 0
            TextBox4.Text = 0
            TextBox5.Text = 0
            TextBox2.Enabled = False
            TextBox4.Enabled = False
            Label18.Text = 0
            Label21.Text = 0
            Label23.Text = 0
            TextBox1.Text = ""
            Txt_parttimemonets.Text = ""
            Txt_punchm.Text = ""
            Label11.Text = ""
            ComboBox2.Text = ""
            ComboBox1.Text = ""
            ComboBox3.Text = ""
            TextBox3.Text = ""
            Txt_at1.Text = ""
            Txt_at2.Text = ""
            Txt_go1.Text = ""
            Txt_go2.Text = ""
            Txt_late.Text = ""
            Txt_parttime.Text = ""
            Txt_parttime.Text = 0
            Txt_punch.Text = ""
            Txt_punch.Text = 0
            Txt_total.Text = ""
            Txt_late.Text = 0

            If gcon.State = 1 Then gcon.Close()
            'gcon.ConnectionString = "Provider=SQLOLEDB.1;Password=141751523;Persist Security Info=True;User ID=sa;Initial Catalog=Tiba;Data Source=ADMIN-PC\AHMED1"
            gcon.ConnectionString = "Provider=SQLOLEDB.1;Password=141751527;Persist Security Info=True;User ID=sa;Initial Catalog=Elhana;Data Source=PC-PC\SQLEXPRESS"
            gcon.Open()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        GroupBox4.Enabled = False
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        GroupBox4.Enabled = True
        TextBox1.Text = ""
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            TextBox1.Text = "لا يوجد ملاحظات"
        End If
        Try
            For kk As Integer = 1 To 17
                Dim s As Boolean
                s = False

                If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
                rs_EmpAttand.Open("Employyes_Presence", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rs_EmpAttand.AddNew()
                rs_EmpAttand("Emp_Day").Value = d_name
                rs_EmpAttand("Emp_Date").Value = DateTimePicker1.Value.ToShortDateString()
                rs_EmpAttand("Emp_In1").Value = Txt_at1.Text
                rs_EmpAttand("Emp_Out1").Value = Txt_go1.Text
                rs_EmpAttand("Emp_In2").Value = Txt_at2.Text
                rs_EmpAttand("Emp_Out2").Value = Txt_go2.Text
                rs_EmpAttand("Emp_Punash").Value = Txt_punch.Text
                rs_EmpAttand("Emp_Late").Value = Txt_late.Text
                rs_EmpAttand("Emp_Additional").Value = Txt_parttime.Text
                rs_EmpAttand("Emp_Ancestor").Value = txt_Ancestor.Text
                rs_EmpAttand("Emp_flag").Value = s
                rs_EmpAttand("User_ID").Value = u_Id
                rs_EmpAttand("Emp_Id").Value = E_id
                rs_EmpAttand("Emp_Total").Value = Txt_total.Text
                rs_EmpAttand("Month_ID").Value = month_id
                rs_EmpAttand("Emp_Salary_Day").Value = Label18.Text
                rs_EmpAttand("Emp_Note").Value = TextBox1.Text
                s = True
                rs_EmpAttand("hodor").Value = s
                rs_EmpAttand.Update()

                rs_EmpAttand.Close()
                get_mname(E_id, month_id, "حفظ حضور")

            Next
            MsgBox("تم حفظ الحضور بنجاح", MsgBoxStyle.Information, "حفظ بيانات الحضور")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox2.Text = 0
        TextBox4.Text = 0
        TextBox5.Text = 0
        If TextBox1.Text = "" Then
            TextBox1.Text = "لا يوجد ملاحظات"
        End If
        If ComboBox3.Text = "" Then
            MsgBox("من فضلك أختر اليوم", MsgBoxStyle.Information, "تنبيه")
            ComboBox3.Select()
            Exit Sub
        End If
        If ComboBox1.Text = "" Then
            MsgBox("من فضلك أختر الشهر", MsgBoxStyle.Information, "تنبيه")
            ComboBox1.Select()
            Exit Sub
        End If
        If Txt_at1.Text = "" Then
            MsgBox("من فضلك أدخل وقت الحضور", MsgBoxStyle.Information, "تنبيه")
            Txt_at1.Select()
            Exit Sub
        End If
        If Txt_go2.Text = "" Then
            MsgBox("من فضلك أدخل وقت الأنصراف", MsgBoxStyle.Information, "تنبيه")
            Txt_go2.Select()
            Exit Sub
        End If
        If Txt_total.Text = "" Then
            MsgBox("من فضلك أدخل أجمالى عدد ساعات العمل لهذا اليوم ", MsgBoxStyle.Information, "تنبيه")
            Txt_total.Select()
            Exit Sub
        End If

        Try
            Me.Cursor = Cursors.WaitCursor
            dy = DateTimePicker1.Value.Year
            dm = DateTimePicker1.Value.Month
            dd = DateTimePicker1.Value.Day
            dstri = dy + "-" + dm + "-" + dd
            If rs_SofUserName.State = 1 Then rs_SofUserName.Close()
            Dim s As Boolean
            s = False
            Dim s1 As Boolean
            s1 = True
            rs_SofUserName.CursorLocation = CursorLocationEnum.adUseClient
            rs_SofUserName.Open("Select * From Employyes Where Emp_State='" & s1 & "' And Emp_Flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            Dim d1 As Integer

            Do While Not rs_SofUserName.EOF
                d1 = rs_SofUserName("Emp_Id").Value
                sarly_day = rs_SofUserName("Emp_Salary_Day").Value
                mainsalry = rs_SofUserName("Emp_Salary").Value
                If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
                rs_EmpAttand.CursorLocation = CursorLocationEnum.adUseClient
                rs_EmpAttand.Open("Select * From Employyes_Presence Where Emp_Id=" & d1 & " And Emp_Datestri='" & dstri & "' And Emp_flag='" & s & "'", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                If rs_EmpAttand.EOF Or rs_EmpAttand.BOF Then


                    'Dim u_name As Integer
                    'If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    'rs_StoreVend.CursorLocation = CursorLocationEnum.adUseClient
                    'rs_StoreVend.Open("Select * From Employyes_Presence Where EmpP_Id= (SELECT MAX(EmpP_Id)  FROM Employyes_Presence)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    'u_name = rs_StoreVend("EmpP_Id").Value
                    'Ep_id = u_name + 1

                    If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
                    rs_EmpAttand.Open("Employyes_Presence", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    rs_EmpAttand.AddNew()
                    'rs_EmpAttand("EmpP_Id").Value = Ep_id
                    rs_EmpAttand("Emp_Day").Value = d_name
                    rs_EmpAttand("Emp_Date").Value = DateTimePicker1.Value.ToShortDateString()
                    rs_EmpAttand("Emp_Datestri").Value = dstri
                    rs_EmpAttand("Emp_In1").Value = Txt_at1.Text
                    rs_EmpAttand("Emp_Out1").Value = Txt_go1.Text
                    rs_EmpAttand("Emp_In2").Value = Txt_at2.Text
                    rs_EmpAttand("Emp_Out2").Value = Txt_go2.Text
                    rs_EmpAttand("Emp_Punash").Value = Val(Txt_punch.Text)
                    rs_EmpAttand("Emp_Late").Value = Txt_late.Text
                    rs_EmpAttand("Emp_Additional").Value = Val(Txt_parttime.Text)
                    rs_EmpAttand("Emp_Ancestor").Value = txt_Ancestor.Text
                    rs_EmpAttand("Emp_flag").Value = s
                    rs_EmpAttand("User_ID").Value = u_Id
                    rs_EmpAttand("Emp_Id").Value = d1
                    rs_EmpAttand("Emp_Total").Value = Txt_total.Text
                    rs_EmpAttand("Month_ID").Value = month_id
                    rs_EmpAttand("Emp_Salary_Day").Value = sarly_day
                    rs_EmpAttand("Emp_Note").Value = TextBox1.Text
                    rs_EmpAttand("hodor").Value = s1
                    rs_EmpAttand("Emp_Additionalmenets").Value = 0
                    rs_EmpAttand("Emp_Punashmenets").Value = 0
                    rs_EmpAttand.Update()




                    If rs_StoreVend.State = 1 Then rs_StoreVend.Close()
                    rs_StoreVend.CursorLocation = CursorLocationEnum.adUseClient
                    rs_StoreVend.Open("Select * From Employyes_Presence Where EmpP_Id= (SELECT MAX(EmpP_Id)  FROM Employyes_Presence)", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)


                    Ep_id = rs_StoreVend("EmpP_Id").Value

                    get_mname(d1, month_id, "حفظ حضور")

                End If


                rs_SofUserName.MoveNext()
            Loop
            Me.Cursor = Cursors.Default
            'rs_EmpAttand.Close()
            'rs_SofUserName.Close()
            MsgBox("تم حفظ الحضور بنجاح", MsgBoxStyle.Information, "حفظ بيانات الحضور")
            Label11.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            TextBox1.Text = ""
            ComboBox3.Text = ""
            TextBox3.Text = ""
            Txt_at1.Text = ""
            Txt_at2.Text = ""
            Txt_go1.Text = ""
            Txt_go2.Text = ""
            txt_Ancestor.Text = 0
            Txt_late.Text = 0
            Txt_parttime.Text = 0
            Txt_punch.Text = 0
            Txt_parttimemonets.Text = 0
            Txt_punchm.Text = 0
            Txt_total.Text = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        TextBox4.Text = Val(TextBox2.Text) * Val(sarly_minute)
        Dim i As Integer
        i = TextBox4.Text
        TextBox5.Text = i
    End Sub


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            If rs_EmpAttand.State = 1 Then rs_EmpAttand.Close()
            rs_EmpAttand.CursorLocation = CursorLocationEnum.adUseClient
            rs_EmpAttand.Open("Select * From Employyes_Presence", cn1, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)



            Do While Not rs_EmpAttand.EOF

                Dim yu As Boolean
                yu = rs_EmpAttand("hodor").Value
                If yu = False Then
                    rs_EmpAttand("Emp_Note").Value = "غائب"

                End If
                rs_EmpAttand.Update()


                rs_EmpAttand.MoveNext()
            Loop
            MsgBox("done", MsgBoxStyle.Information, "تنبيه")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "خطأ فى البرنامج")
        End Try
    End Sub

   
End Class