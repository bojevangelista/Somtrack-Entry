Imports System.Data.SqlClient
Public Class Form3
    Dim user = Environment.UserName
    Dim TableSet = My.Computer.FileSystem.ReadAllText("C:\Users\" + user + "\Documents\OPF-TableNo.txt")
    'SERVER'
    Dim con As SqlConnection = New SqlConnection("Data Source=10.130.15.40;Initial Catalog=somtrackdbprod;User ID=somtrack2;Password=sompass12345")
    Dim con2 As SqlConnection = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
    Dim con3 As SqlConnection = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
    Dim BOMStatus = 0
    Dim gCategoryID = 0
    Dim gSubCategoryID = 0
    Dim gSplintID = 0
    Dim Entry = 0
    Dim LabPan = 0

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            Label3.ForeColor = Color.AntiqueWhite
            Label15.Text = TextBox1.Text
            TextBox1.Text = ""
            'LOCALHOST' 
            'con = New SqlConnection("Data Source=localhost;Initial Catalog=SomnoMed;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
            'LOCALHOST'


            '''''''CHECK ACTIVE CASE'''''''''''
            Dim Activequery As String = "Select [SomtrackID], [StationID],Status From [SMProduction].[dbo].[ProductionHead] Where [TableNo] = @TS "
            Dim Activecmd As SqlCommand = New SqlCommand(Activequery, con2)
            Activecmd.Parameters.AddWithValue("@TS", TableSet)
            con2.Open()
            Using reader As SqlDataReader = Activecmd.ExecuteReader()
                If reader.HasRows Then
                    While reader.Read()
                        If Label15.Text = reader.Item("SomtrackID") And reader.Item("StationID") = 1 Then
                            con2.Close()
                            GoTo Accept
                        ElseIf Label15.Text = reader.Item("SomtrackID") And reader.Item("Status") = 4 Then
                            con2.Close()
                            GoTo Accept

                        ElseIf Label15.Text = reader.Item("SomtrackID") Then
                            Label15.Text = reader.Item("SomtrackID")
                            Label3.ForeColor = Color.Red
                            Label3.Text = "already scanned"
                            con2.Close()
                            GoTo Denied
                        ElseIf reader.Item("StationID") = 1 Then
                            Label15.Text = reader.Item("SomtrackID")
                            Label3.ForeColor = Color.Red
                            Label3.Text = "is still active"
                            con2.Close()
                            GoTo Denied
                        End If
                    End While

                End If
            End Using
            con2.Close()
Accept:
            If (Entry = 0) Then

                '''''''CHECK OPEN CASE'''''''''''
                Dim Openquery As String = "SELECT * FROM ProductionDetails as PD LEFT JOIN StationProcess as SP ON PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead as PH ON PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers as TM ON TM.StationID = SP.StationID LEFT JOIN TableSet as TS ON TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS AND SP.StationID = 1 AND TS.TableSetStatus = 1 AND TM.TableMemberStatus = 1 AND PD.Status = 1 AND PH.SomtrackID = @Som"
                Dim Opencmd As SqlCommand = New SqlCommand(Openquery, con2)
                Opencmd.Parameters.AddWithValue("@Som", Label15.Text)
                Opencmd.Parameters.AddWithValue("@TS", TableSet)
                con2.Open()
                Using reader As SqlDataReader = Opencmd.ExecuteReader()
                    If reader.HasRows Then
                        con2.Close()
                        GoTo Update
                    End If
                End Using
                con2.Close()


                '''''''QUERY FOR SELECTING PROCESS OF THE PRODUCT'''''''''''
                Dim prodquery As String = "Select LDT.DeviceTypeId, LDT.DeviceTypeName as DeviceName From LstDevice as LD Left Join LstDeviceType as LDT ON LD.ProductTypeID = LDT.DeviceTypeId WHERE LD.DeviceID = @prod"
                Dim prodcmd As SqlCommand = New SqlCommand(prodquery, con)
                prodcmd.Parameters.AddWithValue("@prod", Label15.Text)
                con.Open()
                Using reader As SqlDataReader = prodcmd.ExecuteReader()
                    If reader.HasRows Then
                        While reader.Read()
                            Label3.Text = reader.Item("DeviceName")
                            Dim DTID = reader.Item("DeviceTypeId")
                            Converter(DTID)
                        End While
                        CheckBox1.Enabled = True
                        Button1.Enabled = True
                        Panel5.Enabled = True
                    Else
                        Label3.ForeColor = Color.Red
                        Label3.Text = "Invalid Somtrack"
                        CheckBox1.Enabled = False
                        Button1.Enabled = False
                        Panel5.Enabled = False
                    End If
                End Using
                con.Close()
            Else
Update:
                ''''UPDATE HEAD''''
                con2.Open()
                Dim UpdateHead As String = "UPDATE ProductionHead SET StationID = 2 WHERE SomtrackID = @Som"
                Dim UpdateHeadQuery As SqlCommand = New SqlCommand(UpdateHead, con2)
                UpdateHeadQuery.Parameters.AddWithValue("@Som", Label15.Text)
                UpdateHeadQuery.ExecuteNonQuery()
                con2.Close()

                ''''UPDATE DETAILS''''
                con2.Open()
                Dim UpdateDetails As String = "UPDATE PD SET PD.DateEnded = GETDATE(), PD.Status = 5 FROM ProductionDetails as PD LEFT JOIN StationProcess as SP ON PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead as PH ON PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers as TM ON TM.StationID = SP.StationID LEFT JOIN TableSet as TS ON TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS AND SP.StationID = 1 AND TS.TableSetStatus = 1 AND TM.TableMemberStatus = 1 AND PD.Status = 1 AND PH.SomtrackID = @Som"
                Dim UpdateDetailsQuery As SqlCommand = New SqlCommand(UpdateDetails, con2)
                UpdateDetailsQuery.Parameters.AddWithValue("@Som", Label15.Text)
                UpdateDetailsQuery.Parameters.AddWithValue("@TS", TableSet)
                UpdateDetailsQuery.ExecuteNonQuery()
                con2.Close()


                ''''UPDATE NEXT DETAILS''''
                con2.Open()
                Dim UpdateNextDetails As String = "Update PD SET PD.Status = 2 FROM [SMProduction].[dbo].[ProductionHead] as PH LEFT JOIN StationProcess as SP ON SP.StationID = PH.StationID LEFT JOIN ProductionDetails as PD ON PD.ProductionHeadID = PH.ProductionHeadID And PD.BOMDID = SP.BOMDID WHERE PH.SomtrackID = @Som"
                Dim UpdateNextDetailsQuery As SqlCommand = New SqlCommand(UpdateNextDetails, con2)
                UpdateNextDetailsQuery.Parameters.AddWithValue("@Som", Label15.Text)
                UpdateNextDetailsQuery.ExecuteNonQuery()
                con2.Close()


                MessageBox.Show("Case is now on Queue", "Complete")
                Me.Controls.Clear()
                Me.InitializeComponent()
                TextBox1.Enabled = True
                TextBox1.Select()


                LoadCategory()
                Entry = 0
            End If

        End If
Denied:
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If (CheckBox1.Checked = True) Then
            Panel1.Enabled = True
            Label7.Enabled = True
            CheckCategory()
            CheckSubCategory()
            Label10.Enabled = True
            Panel4.Enabled = True
        Else
            Panel1.Enabled = False
            Label7.Enabled = False
            CheckCategory()
            CheckSubCategory()
            Label10.Enabled = False
            Panel4.Enabled = False
            gCategoryID = 0
            gSubCategoryID = 0
            gSplintID = 0
        End If
    End Sub
    Private Sub Converter(x)
        '''''''QUERY FOR SELECTING BOM Converter'''''''''''
        Dim convquery As String = "SELECT C.[ProductSubID] , PT.ProductTypeID , PF.ProductFamilyID FROM [Converter] as C LEFT JOIN ProductSubType as PST ON PST.ProductSubID = C.ProductSubID LEFT JOIN ProductType as PT ON PT.ProductTypeID = PST.ProductTypeID LEFT JOIN ProductFamily as PF ON PF.ProductFamilyID = PT.ProductFamilyID WHERE C.Status = 1 AND C.ProductTypeID = @PID"
        Dim convcmd As SqlCommand = New SqlCommand(convquery, con2)
        convcmd.Parameters.AddWithValue("@PID", x)
        con2.Open()
        Using reader As SqlDataReader = convcmd.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    ComboBox1.SelectedValue = reader.Item("ProductFamilyID")
                    ComboBox2.SelectedValue = reader.Item("ProductTypeID")
                    ComboBox3.SelectedValue = reader.Item("ProductSubID")
                End While
                ComboBox1.Enabled = False
                ComboBox2.Enabled = False
                ComboBox3.Enabled = False
                TextBox1.Enabled = False
            Else
                ComboBox1.Enabled = True
                ComboBox2.Enabled = True
                ComboBox3.Enabled = True
                MessageBox.Show("Please supply the correct information needed for this case", "Device info is not available", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                TextBox1.Enabled = False
                LoadCategory()
            End If
        End Using
        con2.Close()



    End Sub

    Private Sub CheckCategory()
        If CheckBox1.Checked = False Then
            Panel2.Enabled = False
            Label8.Enabled = False
            RadioButton1.BackColor = Color.Transparent
            RadioButton2.BackColor = Color.Transparent
        ElseIf RadioButton1.Checked = True Then

            Panel2.Enabled = True
            Label8.Enabled = True
            RadioButton1.BackColor = Color.DarkGreen
            RadioButton2.BackColor = Color.Transparent
            gCategoryID = 1
        ElseIf RadioButton2.Checked = True Then

            Panel2.Enabled = True
            Label8.Enabled = True
            RadioButton2.BackColor = Color.DarkGreen
            RadioButton1.BackColor = Color.Transparent
            gCategoryID = 2
        End If
    End Sub
    Private Sub CheckSubCategory()
        If CheckBox1.Checked = False Then
            Label9.Enabled = False
            Panel3.Enabled = False
        ElseIf RadioButton3.Checked = True Then
            Label9.Enabled = True
            Panel3.Enabled = True
            RadioButton7.Enabled = True
            RadioButton8.Enabled = True
            RadioButton9.Enabled = True
            RadioButton10.Enabled = True
            gSubCategoryID = 1

            RadioButton3.BackColor = Color.DarkGreen
            RadioButton4.BackColor = Color.Transparent

        ElseIf RadioButton4.Checked = True Then
            Label9.Enabled = True
            Panel3.Enabled = True
            RadioButton7.Enabled = False
            RadioButton8.Enabled = False
            RadioButton9.Enabled = False
            RadioButton10.Enabled = False
            RadioButton6.Select()
            gSubCategoryID = 2

            RadioButton4.BackColor = Color.DarkGreen
            RadioButton3.BackColor = Color.Transparent
        End If
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadCategory()
        TextBox1.Select()
        LoadOption()


    End Sub
    Private Sub LoadCategory()
        '''''''QUERY FOR SELECTING LEADMAN'''''''''''
        Dim familyQuery As String = "SELECT ProductFamilyID, ProductFamilyName FROM ProductFamily"
        Dim familyCmd As SqlCommand = New SqlCommand(familyQuery, con2)
        con2.Close()
        con2.Open()
        Using reader As SqlDataReader = familyCmd.ExecuteReader()
            If reader.HasRows Then
                Dim dt As DataTable = New DataTable
                dt.Load(reader)

                ComboBox1.DataSource = dt
                ComboBox1.ValueMember = "ProductFamilyID"
                ComboBox1.DisplayMember = "ProductFamilyName"

            End If
        End Using
        con2.Close()
    End Sub
    Private Sub LoadOption()
        '''''''QUERY FOR SELECTING LEADMAN'''''''''''

        Dim dt As DataTable = New DataTable
        Dim familyQuery As String = "SELECT 0 as OptionHeadID,'---Add on---' as OptionName UNION SELECT [OptionHeadID] ,[OptionName] FROM [SMProduction].[dbo].[OptionHead] WHERE OptionType = 1 ORDER BY OptionName ASC"
        Dim familyCmd As SqlCommand = New SqlCommand(familyQuery, con2)
        con2.Close()
        con2.Open()

        Using reader As SqlDataReader = familyCmd.ExecuteReader()
            If reader.HasRows Then
                dt.Load(reader)
                ComboBox4.DataSource = dt
                ComboBox4.ValueMember = "OptionHeadID"
                ComboBox4.DisplayMember = "OptionName"

            End If
        End Using
        con2.Close()

        familyQuery = "SELECT 0 as OptionHeadID,'---Variants---' as OptionName UNION SELECT [OptionHeadID] ,[OptionName] FROM [SMProduction].[dbo].[OptionHead] WHERE OptionType = 2 ORDER BY OptionName ASC"
        familyCmd = New SqlCommand(familyQuery, con2)
        con2.Close()
        con2.Open()

        Using reader As SqlDataReader = familyCmd.ExecuteReader()
            If reader.HasRows Then
                dt.Load(reader)
                ComboBox4.DataSource = dt
                ComboBox4.ValueMember = "OptionHeadID"
                ComboBox4.DisplayMember = "OptionName"

            End If
        End Using
        con2.Close()

        familyQuery = "SELECT 0 as OptionHeadID,'---Instructions---' as OptionName UNION SELECT [OptionHeadID] ,[OptionName] FROM [SMProduction].[dbo].[OptionHead] WHERE OptionType = 3 ORDER BY OptionName ASC"
        familyCmd = New SqlCommand(familyQuery, con2)
        con2.Close()
        con2.Open()

        Using reader As SqlDataReader = familyCmd.ExecuteReader()
            If reader.HasRows Then
                dt.Load(reader)
                ComboBox4.DataSource = dt
                ComboBox4.ValueMember = "OptionHeadID"
                ComboBox4.DisplayMember = "OptionName"

            End If
        End Using
        con2.Close()


    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        '''''''VARIABLES'''''''
        Dim con As SqlConnection
        Dim CB1 As String = ComboBox1.SelectedValue.ToString
        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
        '''''''QUERY FOR SELECTING LEADMAN'''''''''''
        Dim typeQuery As String = "SELECT [ProductTypeID], [ProductTypeName] FROM [ProductType] WHERE [ProductFamilyID] = @PF"
        Dim typeCmd As SqlCommand = New SqlCommand(typeQuery, con)
        typeCmd.Parameters.AddWithValue("@PF", CB1)
        con.Open()
        Using reader As SqlDataReader = typeCmd.ExecuteReader()
            If reader.HasRows Then
                Dim dt As DataTable = New DataTable
                dt.Load(reader)

                ComboBox2.DataSource = dt
                ComboBox2.ValueMember = "ProductTypeID"
                ComboBox2.DisplayMember = "ProductTypeName"

            End If
        End Using
        con.Close()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        '''''''VARIABLES'''''''
        Dim con As SqlConnection
        Dim CB2 As String = ComboBox2.SelectedValue.ToString
        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
        '''''''QUERY FOR SELECTING LEADMAN'''''''''''
        Dim subQuery As String = "SELECT [ProductSubID], CONCAT([ProductSubTypeName], ' ',[ProductSubClass]) as Subname  FROM [ProductSubType] WHERE [ProductTypeID] = @PT"
        Dim subCmd As SqlCommand = New SqlCommand(subQuery, con)
        subCmd.Parameters.AddWithValue("@PT", CB2)
        con.Open()
        Using reader As SqlDataReader = subCmd.ExecuteReader()
            If reader.HasRows Then
                Dim dt As DataTable = New DataTable
                dt.Load(reader)

                ComboBox3.DataSource = dt
                ComboBox3.ValueMember = "ProductSubID"
                ComboBox3.DisplayMember = "Subname"

            End If
        End Using
        con.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        CheckCategory()
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        CheckCategory()
    End Sub
    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        CheckSubCategory()
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        CheckSubCategory()
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If MessageBox.Show("Are you sure?", "Current inputs will be cleared", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Yes Then
            Me.Controls.Clear()
            Me.InitializeComponent()
            TextBox1.Select()
            LoadCategory()
        End If
    End Sub

    Private Sub SplintSelect(sid)
        For Each ctrl As RadioButton In Panel3.Controls
            If ctrl.Checked Then
                ctrl.BackColor = Color.DarkGreen
                gSplintID = sid
            Else
                ctrl.BackColor = Color.Transparent
            End If
        Next
    End Sub
    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged
        SplintSelect(1)
    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged
        SplintSelect(2)
    End Sub
    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        SplintSelect(3)
    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        SplintSelect(4)
    End Sub

    Private Sub RadioButton10_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton10.CheckedChanged
        SplintSelect(5)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If MessageBox.Show("Do you want to proceed with this case", "Please double check details before proceeding", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Yes Then
            Dim s1 As String = "0"
            Dim TableSetID = 0
            '''''''QUERY FOR SELECTING ACTIVE TABLE'''''''''''
            Dim TableSetIDquery As String = "SELECT TableSetID FROM [SMProduction].[dbo].[TableSet] WHERE TableID = @TS AND TableSetStatus = 1"
            Dim TableSetIDquerycmd As SqlCommand = New SqlCommand(TableSetIDquery, con2)
            TableSetIDquerycmd.Parameters.AddWithValue("@TS", TableSet)
            con2.Open()
            Using reader As SqlDataReader = TableSetIDquerycmd.ExecuteReader()
                If reader.HasRows Then
                    While reader.Read()
                        TableSetID = reader.Item("TableSetID").ToString
                    End While
                Else
                    MessageBox.Show("There is no active roster for your table. Contact your Leadman to activate roster", "Failed to register Case")
                    GoTo StopEntry
                End If
            End Using
            con2.Close()

            ''''INSERT HEAD''''
            con2.Open()
            Dim InsertHead As String = "INSERT INTO ProductionHead (SomtrackID,DateStarted, Status, ProductSubID, CategoryID, SubCategoryID, SplintID, LabPan, TableSetID, TableNo, StationID) Values (@Som, GETDATE(), 1, @PS, @CID, @SCID, @SID, @LP, @TSID, @TS, 1)"
            Dim InsertHeadQuery As SqlCommand = New SqlCommand(InsertHead, con2)
            InsertHeadQuery.Parameters.AddWithValue("@Som", Label15.Text)
            InsertHeadQuery.Parameters.AddWithValue("@PS", ComboBox3.SelectedValue)
            InsertHeadQuery.Parameters.AddWithValue("@CID", gCategoryID)
            InsertHeadQuery.Parameters.AddWithValue("@SCID", gSubCategoryID)
            InsertHeadQuery.Parameters.AddWithValue("@SID", gSplintID)
            InsertHeadQuery.Parameters.AddWithValue("@TS", TableSet)
            InsertHeadQuery.Parameters.AddWithValue("@LP", LabPan)
            InsertHeadQuery.Parameters.AddWithValue("@TSID", TableSetID)
            InsertHeadQuery.ExecuteNonQuery()
            con2.Close()


            ''''INSERT DETAILS''''
            con2.Open()
            Dim InsertDetails As String = "INSERT INTO [ProductionDetails] ([ProductionHeadID], [BOMDID], [Status]) SELECT [ProductionHeadID], BOMDID, 3 as Status FROM [ProductionHead] as PH LEFT JOIN Converter AS C ON C.ProductSubID = PH.ProductSubID LEFT JOIN BillOfMaterials as BM ON BM.BOMID = C.BOMID LEFT JOIN BillOfMaterialsDetails as BD ON BD.BOMID = BM.BOMID WHERE BM.BOMStatus = 1 AND BD.BOMDStatus = 1 AND PH.SomtrackID = @Som"
            Dim InsertDetailsQuery As SqlCommand = New SqlCommand(InsertDetails, con2)
            InsertDetailsQuery.Parameters.AddWithValue("@Som", Label15.Text)
            InsertDetailsQuery.ExecuteNonQuery()
            con2.Close()


            ''''UPDATE DETAILS''''
            con2.Open()
            Dim UpdateDetails As String = "UPDATE PD SET PD.EmployeeID = TM.EmployeeID , PD.DateStarted = GETDATE(), PD.Status = 1 FROM ProductionDetails as PD LEFT JOIN StationProcess as SP ON PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead as PH ON PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers as TM ON TM.StationID = SP.StationID LEFT JOIN TableSet as TS ON TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS AND SP.StationID = 1 AND TS.TableSetStatus = 1 AND TM.TableMemberStatus = 1 AND PH.SomtrackID = @Som"
            Dim UpdateDetailsQuery As SqlCommand = New SqlCommand(UpdateDetails, con2)
            UpdateDetailsQuery.Parameters.AddWithValue("@Som", Label15.Text)
            UpdateDetailsQuery.Parameters.AddWithValue("@TS", TableSet)
            UpdateDetailsQuery.ExecuteNonQuery()
            con2.Close()


            ''''INSERT SUB DETAILS''''
            con2.Open()
            Dim InsertSubDetails As String = "INSERT INTO ProductionSubDetail (ProductionDetailID, ODetailsID, Points) SELECT [ProductionDetailID] ,OD.ODetailsID ,OD.Points FROM ProductionHead as PH LEFT JOIN [SMProduction].[dbo].[ProductionDetails] as PD ON PD.ProductionHeadID = PH.ProductionHeadID LEFT JOIN BillOfMaterialsDetails as BD ON BD.BOMDID = PD.BOMDID LEFT JOIN BillOfMaterials as BM ON BM.BOMID = BD.BOMID LEFT JOIN Operations as O ON O.OperationID = BD.OperationID LEFT JOIN OperationsDetail as OD ON OD.OperationID = O.OperationID WHERE BM.BOMStatus = 1 AND BD.BOMDStatus = 1 AND PH.SomtrackID = @Som"
            Dim InsertSubDetailsQuery As SqlCommand = New SqlCommand(InsertSubDetails, con2)
            InsertSubDetailsQuery.Parameters.AddWithValue("@Som", Label15.Text)
            InsertSubDetailsQuery.ExecuteNonQuery()
            con2.Close()

            Entry = 1
            Button1.Enabled = False
            TextBox1.Enabled = True
            TextBox1.Select()

        End If
StopEntry:
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label1.Text = Format(Now, "hh:mm:ss")
        Label2.Text = Format(Now, "MMMM dd, yyyy")

        '''''''QUERY FOR SELECTING ACTIVE TABLE'''''''''''
        Dim TableSetIDquery As String = "SELECT TableSetName FROM [SMProduction].[dbo].[TableSet] WHERE TableID = @TS AND TableSetStatus = 1"
        Dim TableSetIDquerycmd As SqlCommand = New SqlCommand(TableSetIDquery, con3)
        TableSetIDquerycmd.Parameters.AddWithValue("@TS", TableSet)
        con3.Open()
        Using reader As SqlDataReader = TableSetIDquerycmd.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    Label21.Text = reader.Item("TableSetName").ToString
                End While
                TextBox1.Enabled = True
                If Label15.Text = "" Then
                    TextBox1.Select()
                End If
                If Label13.Text = "No Active Roster" Then
                    Label3.Text = ""
                End If
            Else
                TextBox1.Enabled = False
                Label3.Text = "No Active Roster"
                Label3.ForeColor = Color.Red

            End If
        End Using
        con3.Close()
    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click
        Label12.BackColor = Color.DarkGreen
        Label13.BackColor = Color.Transparent
        Label14.BackColor = Color.Transparent
        Label16.BackColor = Color.Transparent
        Label17.BackColor = Color.Transparent
        Label18.BackColor = Color.Transparent
        LabPan = 1
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        Label12.BackColor = Color.Transparent
        Label13.BackColor = Color.DarkGreen
        Label14.BackColor = Color.Transparent
        Label16.BackColor = Color.Transparent
        Label17.BackColor = Color.Transparent
        Label18.BackColor = Color.Transparent
        LabPan = 2
    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        Label12.BackColor = Color.Transparent
        Label13.BackColor = Color.Transparent
        Label14.BackColor = Color.DarkGreen
        Label16.BackColor = Color.Transparent
        Label17.BackColor = Color.Transparent
        Label18.BackColor = Color.Transparent
        LabPan = 3
    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click
        Label12.BackColor = Color.Transparent
        Label13.BackColor = Color.Transparent
        Label14.BackColor = Color.Transparent
        Label16.BackColor = Color.DarkGreen
        Label17.BackColor = Color.Transparent
        Label18.BackColor = Color.Transparent
        LabPan = 4
    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles Label17.Click
        Label12.BackColor = Color.Transparent
        Label13.BackColor = Color.Transparent
        Label14.BackColor = Color.Transparent
        Label16.BackColor = Color.Transparent
        Label17.BackColor = Color.DarkGreen
        Label18.BackColor = Color.Transparent
        LabPan = 5
    End Sub

    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click
        Label12.BackColor = Color.Transparent
        Label13.BackColor = Color.Transparent
        Label14.BackColor = Color.Transparent
        Label16.BackColor = Color.Transparent
        Label17.BackColor = Color.Transparent
        Label18.BackColor = Color.DarkGreen
        LabPan = 6
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        '''''''VARIABLES'''''''
        Dim con As SqlConnection
        Dim CB4 As String = ComboBox4.SelectedValue.ToString
        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
        '''''''QUERY FOR SELECTING LEADMAN'''''''''''
        Dim typeQuery As String = "SELECT [OptionDetail] ,[OptionDetailName] FROM [SMProduction].[dbo].[OptionDetail] WHERE [OptionHeadID] = @OH"
        Dim typeCmd As SqlCommand = New SqlCommand(typeQuery, con)
        typeCmd.Parameters.AddWithValue("@OH", CB4)
        con.Open()
        Using reader As SqlDataReader = typeCmd.ExecuteReader()
            If reader.HasRows Then
                Dim dt As DataTable = New DataTable
                dt.Load(reader)

                ComboBox5.DataSource = dt
                ComboBox5.ValueMember = "OptionDetail"
                ComboBox5.DisplayMember = "OptionDetailName"

            End If
        End Using
        con.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ListBox1.Items.Add(ComboBox4.Text + " - " + ComboBox5.Text + "                                                                                                                        :" + ComboBox5.SelectedValue.ToString)


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ListBox1.Items.Remove(ListBox1.SelectedItem)
    End Sub
End Class
