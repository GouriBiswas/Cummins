Imports System
Imports System.Data
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types
Imports System.Configuration
Imports System.Threading
Imports System.Threading.Tasks
Public Class D2

    Dim RES As String
    Dim squery As String
    Dim rcount As Integer
    Dim rhp As String
    Dim SCREW1 As Double
    Dim SCREW2 As Double
    Dim SCREW3 As Double
    Dim SCREW4 As Double
    Dim SCREW5 As Double
    Dim SCREW6 As Double
    Dim SCREW7 As Double
    Dim SCREW8 As Double
    Dim SCREW9 As Double
    Dim SCREW10 As Double
    Dim SCREW11 As Double
    Dim SCREW12 As Double
    Dim IMIN As Double
    Dim IMAX As Double
    Dim EMIN As Double
    Dim EMAX As Double


    Dim s As Integer = 0
    Dim client As OLWEBSERVICE.NLWEBSERVICESoapClient
    Dim client1 As NLWEBSERVICE.NLWEBSERVICESoapClient
    Dim cn As OracleConnection
    Dim cmd As OracleCommand
    Dim ESNO As String = ""

    Dim esuccess As Boolean
    Dim colorchng As Integer = 0
    Dim dr As OracleDataReader
    Dim da As OracleDataAdapter
    Dim cnstr As String
    Public btnnop As Integer = 0
    Dim CPRESENT As Boolean
    Dim CVN As Boolean
    Dim ready As Boolean
    Dim ready1 As Boolean
    Dim eid As String
    Dim RKHBYPASS As Boolean
    Dim ROCKER_SUCCESS As Boolean
    ' create a new dataset 

    Private myTimer As System.Threading.Timer
    Private myTimer1 As System.Threading.Timer

    Private Sub D2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
        cnstr = ConfigurationManager.ConnectionStrings("MES1").ConnectionString
        'REQEXOUTLET.Text = ""
        lblmodelno.Text = ""
        bedit.Enabled = False
        bedit.Visible = False
        LBLAUTO.Visible = True
        LBLMANUAL.Visible = False
        lbsubtype.Text = ""
        ready = False
        ready1 = False
        'RefreshDataGrid()

        Dim myCallback As New System.Threading.TimerCallback(AddressOf Task1)
        myTimer = New System.Threading.Timer(myCallback, Nothing, 0, 1000)

    End Sub

    Sub AS001()
        Dim enbl As Boolean
        Dim REL As Boolean

        Try
            client = New OLWEBSERVICE.NLWEBSERVICESoapClient

            CPRESENT = client.VARINTVALUE("AS033:CARRIER_PRESENT")
            RKHBYPASS = client.VARINTVALUE("ROCKERHBYPASS")
            enbl = client.VARINTVALUE("AS033:DATA_ENTRY_ENABLE")
            REL = client.VARINTVALUE("AS033:RELEASE_REQUEST")

            If enbl Then
                ENTRYENABLE.BackColor = Color.DarkGreen
            Else
                ENTRYENABLE.BackColor = Color.FromArgb(156, 35, 27)
            End If

            If REL Then
                LBLRELEASE.BackColor = Color.DarkGreen
            Else
                LBLRELEASE.BackColor = Color.FromArgb(156, 35, 27)
            End If

            If esuccess = True Then
                lbldesuccess.BackColor = Color.DarkGreen
                LBLDENTRY.BackColor = Color.DarkGreen
                'LBLFSAFE.BackColor = Color.DarkGreen
                LBLDENTRY.Text = "Success !"
                TXTESNO.Enabled = False

                BTNSAVE.Enabled = False
                esuccess = True
            Else
                lbldesuccess.BackColor = Color.FromArgb(156, 35, 27)
                LBLDENTRY.BackColor = Color.FromArgb(156, 35, 27)
                'LBLFSAFE.BackColor = Color.FromArgb(156, 35, 27)
                'esuccess = False
            End If

            If CPRESENT And esuccess = False And ready = False Then
                LBCPRESENT.BackColor = Color.DarkGreen
                LBLDENTRY.Text = "Ready For data Entry !"
                'TXTESNO.BackColor = Color.LightYellow
                TXTESNO.Enabled = True

                BTNSAVE.Enabled = True
                btnrefresh.Enabled = True

                If client.VARSTRVALUE("AS033:ESNO") = "" Then
                    LBLDENTRY.BackColor = Color.FromArgb(156, 35, 27)
                Else
                    TXTESNO.Text = client.VARSTRVALUE("AS033:ESNO")

                    ready = True
                    LBLDENTRY.BackColor = Color.FromArgb(156, 35, 27)
                End If


            ElseIf CPRESENT = False Then

                LBLDENTRY.BackColor = Color.FromArgb(156, 35, 27)
                LBCPRESENT.BackColor = Color.FromArgb(156, 35, 27)

                LBLDENTRY.Text = "Waiting for Next Engine"
                TXTESNO.Text = ""
                lblmodelno.Text = ""
                lbsubtype.Text = ""

                esuccess = False
                ready = False
                ready1 = False


                'client.VARSETVALUE("AS033:DATA_ENTRY_SUCESS", False)

                'btnrefresh.Enabled = False
                TXTESNO.Enabled = False
                BTNSAVE.Enabled = False

            End If

        Catch ex As Exception

        End Try


    End Sub
    Private Sub Task1(ByVal state As Object)
        client = New OLWEBSERVICE.NLWEBSERVICESoapClient

        Try

            Parallel.Invoke(New Action(AddressOf AS001))

        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        TSL1.Text = Format(Now, "dd-MM-yyyy HH:mm:ss")
        If ready = True And ready1 = False Then
            ready1 = True
            TXTESNO.Focus()
            If TXTESNO.Text <> "" Then
                SendKeys.Send("{ENTER}")
            End If
        End If

    End Sub

    Private Sub savef()
        Dim cnstr As String
        Dim query As String
        Dim RCOUNT As Integer
        Dim ESNO As String
        Dim pno1 As String = ""
        Dim pno2 As String = ""
        Dim pno3 As String = ""
        Dim sno3 As String = ""
        Dim CAL As String = ""
        Dim CVNO As String = ""
        Dim REMARKS As String = ""
        Dim dt As New DataTable()

        ESNO = UCase(TXTESNO.Text)

        If Len(ESNO) < 11 Then
            Exit Sub
        End If

        Try
            cnstr = ConfigurationManager.ConnectionStrings("MES1").ConnectionString '"User Id=MOVICONL1; Password=Kr0n0s#18; Data Source=172.19.223.251:1521/MESDBP01; Pooling =false;"
            cn = New OracleConnection(cnstr)
            cn.Open()

            query = "SELECT COUNT(1) FROM BLOCKDETAILS WHERE ENGINESERIALNO = '" & ESNO & "'"
            cmd = New OracleCommand(query, cn)
            dr = cmd.ExecuteReader()
            dr.Read()
            RCOUNT = dr.GetInt32(0)

            If RCOUNT = 0 Then
                LBLDENTRY.Text = "Please scan correct esno !"
                LBLDENTRY.BackColor = Color.Black

                cn.Close()
                Exit Sub
            Else

                query = "SELECT RECORDNO,BLOCKNO,MODELNO FROM RECORDNO WHERE ESNO = '" & ESNO & "'"
                cmd = New OracleCommand(query, cn)
                dr = cmd.ExecuteReader()
                dr.Read()
                client.VARSETVALUE("AS033_RECORDNO", CStr(dr.Item(0)))
                client.VARSETVALUE("AS033_BLOCKNO", CStr(dr.Item(1)))
                client.VARSETVALUE("AS033_MODELNO", CStr(dr.Item(2)))

                client.VARSETVALUE("AS033_ESNO", ESNO)


                ' query = "UPDATE ENGINE_PARTNO SET CC_OPTION1 = '" & pno1 & "',CC_OPTION2 = '" & pno2 & "',CCOPTION_DTIME = SYSDATE WHERE ENGINESERIALNO = '" & ESNO & "'"

                'query = "INSERT INTO TCL_T_PARTSCANLOG(LOGTIMESTAMP,ESNO,PARTNO1,PARTNO2,REMARKS,STNO) VALUES(SYSDATE,'" & ESNO & "','" & pno1 & "','" & pno2 & "','CCOPTION','AS032')"
                'cmd = New OracleCommand(query, cn)
                'cmd.ExecuteNonQuery()

                client = New OLWEBSERVICE.NLWEBSERVICESoapClient
                'client.VARSETVALUE("AS033:DATA_ENTRY_SUCESS", True)

                lbldesuccess.BackColor = Color.DarkGreen
                LBLDENTRY.BackColor = Color.DarkGreen

                LBLDENTRY.Text = "Success !"
                TXTESNO.Enabled = False

                BTNSAVE.Enabled = False

                esuccess = True

                cn.Close()
            End If


        Catch ex As Exception
            If cn.State = ConnectionState.Open Then
                cn.Close()
            End If
            LBLDENTRY.Text = ex.Message
        End Try
    End Sub

    Private Sub BTNSAVE_Click(sender As Object, e As EventArgs) Handles BTNSAVE.Click
        Dim query As String
        Dim cn As OracleConnection
        Dim cmd As OracleCommand

        cnstr = ConfigurationManager.ConnectionStrings("MES1").ConnectionString '"User Id=MOVICONL1; Password=Kr0n0s#18; Data Source=172.19.223.251:1521/MESDBP01; Pooling =false;"
        cn = New OracleConnection(cnstr)
        cn.Open()

        query = "INSERT INTO TCL_T_ROCKERHIGHT2 (ESNO , SCREWHIGH1, SCREWHIGH2, SCREWHIGH3, SCREWHIGH4, SCREWHIGH5, SCREWHIGH6, SCREWHIGH7, SCREWHIGH8, SCREWHIGH9, SCREWHIGH10, SCREWHIGH11, SCREWHIGH12) VALUES ('" & ESNO & "', '" & TXTEX6.Text & "', '" & TXTIN6.Text & "', '" & TXTEX5.Text & "',  '" & TXTIN5.Text & "', '" & TXTEX4.Text & "',  '" & TXTIN4.Text & "', '" & TXTEX3.Text & "',  '" & TXTIN3.Text & "', '" & TXTEX2.Text & "',  '" & TXTIN2.Text & "', '" & TXTEX1.Text & "',  '" & TXTIN1.Text & "')"
        cmd = New OracleCommand(query, cn)
        cmd.ExecuteNonQuery()

        cn.Close()
        'MessageBox.Show(" Are you sure that you want To save , If yes Then click On OK")

    End Sub

    Private Sub btnrefresh_Click(sender As Object, e As EventArgs) Handles btnrefresh.Click

        Try
            client = New OLWEBSERVICE.NLWEBSERVICESoapClient
            'client.VARSETVALUE("AS033:DATA_ENTRY_SUCESS", False)
            TXTESNO.Text = ""
            TXTESNO.Enabled = True
            lblmodelno.Text = ""

            lbsubtype.Text = ""

            BTNSAVE.Enabled = True
            TXTESNO.Enabled = True

            esuccess = False

            lbldesuccess.BackColor = Color.FromArgb(156, 35, 27)
            LBLDENTRY.Text = "Ready For data Entry !"
            LBLDENTRY.BackColor = Color.FromArgb(156, 35, 27)

            TXTESNO.Enabled = True

            BTNSAVE.Enabled = True

            TXTESNO.BackColor = Color.LightYellow
            TXTESNO.Focus()



            'TXTIN1.Text = ""
            'TXTIN2.Text = ""
            'TXTIN3.Text = ""
            'TXTIN4.Text = ""
            'TXTIN5.Text = ""
            'TXTIN6.Text = ""

            'TXTEX1.Text = ""
            'TXTEX2.Text = ""
            'TXTEX3.Text = ""
            'TXTEX4.Text = ""
            'TXTEX5.Text = ""
            'TXTEX6.Text = ""

            'Lables are declared False
            Label5.Visible = False
            Label7.Visible = False
            Label20.Visible = False
            Label16.Visible = False
            Label24.Visible = False
            Label22.Visible = False
            Label28.Visible = False
            Label26.Visible = False
            Label32.Visible = False
            Label30.Visible = False
            Label36.Visible = False
            Label34.Visible = False

            ' background color to white initially and then updates the color to red or green based on the screw values' conditions.
            LBLIN1.Visible = False
            LBLIN3.Visible = False
            LBLIN5.Visible = False
            LBLIN7.Visible = False
            LBLIN9.Visible = False
            LBLIN11.Visible = False

            LBLIN2.Visible = False
            LBLIN4.Visible = False
            LBLIN6.Visible = False
            LBLIN8.Visible = False
            LBLIN10.Visible = False
            LBLIN12.Visible = False

            ' Make the visibility of all TXTIN and TXTEX true after setting the values
            TXTIN1.Visible = False
            TXTIN2.Visible = False
            TXTIN3.Visible = False
            TXTIN4.Visible = False
            TXTIN5.Visible = False
            TXTIN6.Visible = False

            TXTEX1.Visible = False
            TXTEX2.Visible = False
            TXTEX3.Visible = False
            TXTEX4.Visible = False
            TXTEX5.Visible = False
            TXTEX6.Visible = False

            'For Disabling The Edit Button
            bedit.Visible = False


        Catch ex As Exception

        End Try


    End Sub
    Private Sub TXTESNO_KeyDown(sender As Object, e As KeyEventArgs) Handles TXTESNO.KeyDown
        Dim cnstr As String
        Dim query As String
        Dim RCOUNT As Integer
        Dim TCODE As String = ""
        'Dim eid As String = ""
        Dim ESNO As String = ""
        Dim RHPNO As String = ""
        Dim dt As New DataTable()

        ESNO = UCase(Trim(TXTESNO.Text))
        ' client = New OLWEBSERVICE.NLWEBSERVICESoapClient
        If ESNO = "REJECTED" Then

            'client.VARSETVALUE("AS033:DATA_ENTRY_SUCESS", True)

            lbldesuccess.BackColor = Color.DarkGreen
            LBLDENTRY.BackColor = Color.DarkGreen

            LBLDENTRY.Text = "Success !"
            TXTESNO.Enabled = False

            BTNSAVE.Enabled = False

            esuccess = True
            Exit Sub
        End If

        If Len(ESNO) < 11 Then
            Exit Sub
        End If

        'client.VARSETVALUE("AS033:ESNO", ESNO)

        Try
            cnstr = ConfigurationManager.ConnectionStrings("MES1").ConnectionString '"User Id=MOVICONL1; Password=Kr0n0s#18; Data Source=172.19.223.251:1521/MESDBP01; Pooling =false;"
            cn = New OracleConnection(cnstr)
            cn.Open()

            query = "SELECT COUNT(1) FROM BLOCKDETAILS WHERE ENGINESERIALNO = '" & ESNO & "'"
            cmd = New OracleCommand(query, cn)
            dr = cmd.ExecuteReader()
            dr.Read()
            RCOUNT = dr.Item(0)

            If RCOUNT = 0 Then
                LBLDENTRY.Text = "Please scan correct esno !"
                cn.Close()
            Else

                query = "SELECT MODELNO FROM BLOCKDETAILS WHERE  ENGINESERIALNO = '" & ESNO & "'"
                cmd = New OracleCommand(query, cn)
                dr = cmd.ExecuteReader()
                dr.Read()
                lblmodelno.Text = dr.Item(0)

                query = "SELECT ESUBTYPE_EID FROM PART_COMPARISON WHERE MODELNO = '" & lblmodelno.Text & "'"
                cmd = New OracleCommand(query, cn)
                dr = cmd.ExecuteReader()
                dr.Read()

                If IsDBNull(dr.Item(0)) Then eid = "99" Else eid = dr.Item(0)

                query = "SELECT ENGINE_SUBTYPE FROM TCL_T_ESUBTYPE_MST WHERE ESUBTYPE_EID = '" & eid & "'"
                cmd = New OracleCommand(query, cn)
                dr = cmd.ExecuteReader()
                dr.Read()

                If IsDBNull(dr.Item(0)) Then lbsubtype.Text = "Unknown" Else lbsubtype.Text = dr.Item(0)
                'RHPNO = dr.Item(1)

                cn.Close()
            End If


        Catch ex As Exception
            If cn.State = ConnectionState.Open Then
                cn.Close()
            End If
            LBLDENTRY.Text = ex.Message
        End Try
    End Sub

    Private Sub txtpartno1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Public Sub LTCHECK()


        Dim RES As String
        Dim squery As String
        Dim rcount As Integer
        Dim rhp As String
        Dim ESNO As String
        Dim SCREW1 As Double
        Dim SCREW2 As Double
        Dim SCREW3 As Double
        Dim SCREW4 As Double
        Dim SCREW5 As Double
        Dim SCREW6 As Double
        Dim SCREW7 As Double
        Dim SCREW8 As Double
        Dim SCREW9 As Double
        Dim SCREW10 As Double
        Dim SCREW11 As Double
        Dim SCREW12 As Double
        Dim IMIN As Double
        Dim IMAX As Double
        Dim EMIN As Double
        Dim EMAX As Double

        'Dim screws As New List(Of Tuple(Of TextBox, Label)) From {
        '  New Tuple(Of TextBox, Label)(TXTIN1, LBLIN1),
        '  New Tuple(Of TextBox, Label)(TXTIN2, LBLIN2)}
        ' Add other textboxes and labels as needed}


        ESNO = UCase(TXTESNO.Text)

        If RKHBYPASS = True Then
            ROCKER_SUCCESS = True
            Exit Sub
        End If

        Try
            cnstr = ConfigurationManager.ConnectionStrings("MES1").ConnectionString '"User Id=MOVICONL1; Password=Kr0n0s#18; Data Source=172.19.223.251:1521/MESDBP01; Pooling =false;"
            cn = New OracleConnection(cnstr)
            cn.Open()

            squery = "SELECT COUNT(1) FROM TCL_T_ROCKERHIGHT WHERE ESNO = '" & ESNO & "'"
            cmd = New OracleCommand(squery, cn)
            dr = cmd.ExecuteReader()
            dr.Read()
            rcount = CInt(dr.Item(0))


            If rcount > 0 Then
                squery = "SELECT ROCKERHIGHTP FROM TCL_T_ESUBTYPE_MST WHERE ESUBTYPE_EID = (SELECT ESUBTYPE_EID FROM PART_COMPARISON WHERE MODELNO = '" & lblmodelno.Text & "')"
                cmd = New OracleCommand(squery, cn)
                dr = cmd.ExecuteReader()
                dr.Read()
                rhp = dr.Item(0)

                If rhp = "1" Then
                    IMIN = 0.01
                    IMAX = 3.12
                    EMIN = 1.2
                    EMAX = 3.8
                Else
                    IMIN = 0.01
                    IMAX = 3.12
                    EMIN = 0.5
                    EMAX = 2.26
                End If

                squery = "SELECT RESULTX,SCREWHIGH1,SCREWHIGH2,SCREWHIGH3,SCREWHIGH4,SCREWHIGH5,SCREWHIGH6,SCREWHIGH7,SCREWHIGH8,SCREWHIGH9,SCREWHIGH10,SCREWHIGH11,SCREWHIGH12 FROM TCL_T_ROCKERHIGHT WHERE LOGTIMESTAMP = (select MAX(LOGTIMESTAMP) FROM TCL_T_ROCKERHIGHT WHERE ESNO = '" & ESNO & "') AND ESNO = '" & ESNO & "'"
                cmd = New OracleCommand(squery, cn)
                dr = cmd.ExecuteReader()
                dr.Read()
                RES = CStr(dr.Item(0))
                SCREW1 = dr.Item(1)
                SCREW2 = dr.Item(2)
                SCREW3 = dr.Item(3)
                SCREW4 = dr.Item(4)
                SCREW5 = dr.Item(5)
                SCREW6 = dr.Item(6)
                SCREW7 = dr.Item(7)
                SCREW8 = dr.Item(8)
                SCREW9 = dr.Item(9)
                SCREW10 = dr.Item(10)
                SCREW11 = dr.Item(11)
                SCREW12 = dr.Item(12)

                TXTIN1.Text = SCREW1
                TXTEX1.Text = SCREW2
                TXTIN2.Text = SCREW3
                TXTEX2.Text = SCREW4
                TXTIN3.Text = SCREW5
                TXTEX3.Text = SCREW6
                TXTIN4.Text = SCREW7
                TXTEX4.Text = SCREW8
                TXTIN5.Text = SCREW9
                TXTEX5.Text = SCREW10
                TXTIN6.Text = SCREW11
                TXTEX6.Text = SCREW12

                ' If SCREW1 > IMAX Or SCREW1 < IMIN Then LBLIN1.BackColor = Color.Red
                'If SCREW3 > IMAX Or SCREW1 < IMIN Then LBLIN3.BackColor = Color.Red
                'If SCREW5 > IMAX Or SCREW1 < IMIN Then LBLIN5.BackColor = Color.Red
                'If SCREW7 > IMAX Or SCREW1 < IMIN Then LBLIN7.BackColor = Color.Red
                ' If SCREW9 > IMAX Or SCREW1 < IMIN Then LBLIN9.BackColor = Color.Red
                'If SCREW11 > IMAX Or SCREW1 < IMIN Then LBLIN11.BackColor = Color.Red


                'Check And set color based on conditions
                If SCREW1 > IMAX Or SCREW1 < IMIN Then
                    LBLIN1.BackColor = Color.Red
                    LBLIN1.Text = "NOK"
                    TXTIN1.BackColor = Color.LightSeaGreen


                Else
                    LBLIN1.BackColor = Color.Green
                    LBLIN1.Text = " OK "
                    TXTIN1.BackColor = Color.White
                    TXTIN1.ReadOnly = True
                End If

                If SCREW3 > IMAX Or SCREW3 < IMIN Then
                    LBLIN3.BackColor = Color.Red
                    LBLIN3.Text = "NOK"
                    TXTIN2.BackColor = Color.LightSeaGreen

                Else
                    LBLIN3.BackColor = Color.Green
                    LBLIN3.Text = " OK "
                    TXTIN2.BackColor = Color.White
                    TXTIN2.ReadOnly = True
                End If

                If SCREW5 > IMAX Or SCREW5 < IMIN Then
                    LBLIN5.BackColor = Color.Red
                    LBLIN5.Text = "NOK"
                    TXTIN3.BackColor = Color.LightSeaGreen
                Else
                    LBLIN5.BackColor = Color.Green
                    LBLIN5.Text = "OK"
                    TXTIN3.BackColor = Color.White
                    TXTIN3.ReadOnly = True

                End If

                If SCREW7 > IMAX Or SCREW7 < IMIN Then
                    LBLIN7.BackColor = Color.Red
                    LBLIN7.Text = "NOK"
                    TXTIN4.BackColor = Color.LightSeaGreen
                Else
                    LBLIN7.BackColor = Color.Green
                    LBLIN7.Text = "OK"
                    TXTIN4.BackColor = Color.White
                    TXTIN4.ReadOnly = True
                End If

                If SCREW9 > IMAX Or SCREW9 < IMIN Then
                    LBLIN9.BackColor = Color.Red
                    LBLIN9.Text = "NOK"
                    TXTIN5.BackColor = Color.LightSeaGreen
                Else
                    LBLIN9.BackColor = Color.Green
                    LBLIN9.Text = "OK"
                    TXTIN5.BackColor = Color.White
                    TXTIN5.ReadOnly = True
                End If

                If SCREW11 > IMAX Or SCREW11 < IMIN Then
                    LBLIN11.BackColor = Color.Red
                    LBLIN11.Text = "NOK"
                    TXTIN6.BackColor = Color.LightSeaGreen
                Else
                    LBLIN11.BackColor = Color.Green
                    LBLIN11.Text = "OK"
                    TXTIN6.BackColor = Color.White
                    TXTIN6.ReadOnly = True
                End If


                If SCREW2 > EMAX Or SCREW2 < EMIN Then
                    LBLIN2.BackColor = Color.Red
                    LBLIN2.Text = "NOK"
                    TXTEX1.BackColor = Color.LightSeaGreen


                Else
                    LBLIN2.BackColor = Color.Green
                    LBLIN2.Text = "OK"
                    TXTEX1.BackColor = Color.White
                    TXTEX1.ReadOnly = True
                End If

                If SCREW4 > EMAX Or SCREW4 < EMIN Then
                    LBLIN4.BackColor = Color.Red
                    LBLIN4.Text = "NOK"
                    TXTEX2.BackColor = Color.LightSeaGreen
                Else
                    LBLIN4.BackColor = Color.Green
                    LBLIN4.Text = "OK"
                    TXTEX2.BackColor = Color.White
                    TXTEX2.ReadOnly = True
                End If

                If SCREW6 > EMAX Or SCREW6 < EMIN Then
                    LBLIN6.BackColor = Color.Red
                    LBLIN6.Text = "NOK"
                    TXTEX3.BackColor = Color.LightSeaGreen
                Else
                    LBLIN6.BackColor = Color.Green
                    LBLIN6.Text = "OK"
                    TXTEX3.BackColor = Color.White
                    TXTEX3.ReadOnly = True
                End If

                If SCREW8 > EMAX Or SCREW8 < EMIN Then
                    LBLIN8.BackColor = Color.Red
                    LBLIN8.Text = "NOK"
                    TXTEX4.BackColor = Color.LightSeaGreen
                Else
                    LBLIN8.BackColor = Color.Green
                    LBLIN8.Text = "OK"
                    TXTEX4.BackColor = Color.White
                    TXTEX4.ReadOnly = True
                End If

                If SCREW10 > EMAX Or SCREW10 < EMIN Then
                    LBLIN10.BackColor = Color.Red
                    LBLIN10.Text = "NOK"
                    TXTEX5.BackColor = Color.LightSeaGreen
                Else
                    LBLIN10.BackColor = Color.Green
                    LBLIN10.Text = "OK"
                    TXTEX5.BackColor = Color.White
                    TXTEX5.ReadOnly = True
                End If

                If SCREW12 > EMAX Or SCREW12 < EMIN Then
                    LBLIN12.BackColor = Color.Red
                    LBLIN12.Text = "NOK"
                    TXTEX6.BackColor = Color.LightSeaGreen
                Else
                    LBLIN12.BackColor = Color.Green
                    LBLIN12.Text = "OK"
                    TXTEX6.BackColor = Color.White
                    TXTEX6.ReadOnly = True
                End If

                Label5.Visible = True
                Label7.Visible = True
                Label20.Visible = True
                Label16.Visible = True
                Label24.Visible = True
                Label22.Visible = True
                Label28.Visible = True
                Label26.Visible = True
                Label32.Visible = True
                Label30.Visible = True
                Label36.Visible = True

                Label34.Visible = True

                ' background color to white initially and then updates the color to red or green based on the screw values' conditions.
                LBLIN1.Visible = True
                LBLIN3.Visible = True
                LBLIN5.Visible = True
                LBLIN7.Visible = True
                LBLIN9.Visible = True
                LBLIN11.Visible = True

                LBLIN2.Visible = True
                LBLIN4.Visible = True
                LBLIN6.Visible = True
                LBLIN8.Visible = True
                LBLIN10.Visible = True
                LBLIN12.Visible = True

                ' Make the visibility of all TXTIN and TXTEX true after setting the values
                TXTIN1.Visible = True
                TXTIN2.Visible = True
                TXTIN3.Visible = True
                TXTIN4.Visible = True
                TXTIN5.Visible = True
                TXTIN6.Visible = True

                TXTEX1.Visible = True
                TXTEX2.Visible = True
                TXTEX3.Visible = True
                TXTEX4.Visible = True
                TXTEX5.Visible = True
                TXTEX6.Visible = True

                'For Enabling The Edit Button
                bedit.Visible = True

                ROCKER_SUCCESS = True
            End If

        Catch ex As Exception
            If cn.State = ConnectionState.Open Then
                cn.Close()
            End If
            LBLDENTRY.Text = ex.Message
        End Try
        'If LBLIN1.BackColor = Color.Red Then

        'End If
    End Sub


    'Sub CheckInputs(EMIN As Double, EMAX As Double, IMIN As Double, IMAX As Double, ByRef labels() As Label, ByVal ParamArray screws() As Double)
    'For i As Integer = 0 To labels.Length - 1
    'Dim lbl As Label = labels(i)
    'Dim screwValue As Double = screws(i)

    ' Check the screw value against the specified ranges
    'If screwValue > EMAX Or screwValue < EMIN Or screwValue > IMAX Or screwValue < IMIN Then
    '        lbl.BackColor = Color.Red
    '      lbl.Text = "NOK"
    'Else
    '      lbl.BackColor = Color.Green
    '      lbl.Text = "OK"
    ' End If
    'Next
    'End Sub
    'Private Sub UpdateScrewValue(screwId As String, newValue As Integer)
    'Dim query As String = "UPDATE screws SET value = @newValue WHERE screw_id = @screwId"

    'End Sub


    Declare Function Wow64DisableWow64FsRedirection Lib "kernel32" (ByRef oldvalue As Long) As Boolean
    Declare Function Wow64EnableWow64FsRedirection Lib "kernel32" (ByRef oldvalue As Long) As Boolean
    Private osk As String = "C:\Windows\System32\osk.exe"
    Private Sub VKEY_Click(sender As Object, e As EventArgs) Handles VKEY.Click
        Dim old As Long
        If Environment.Is64BitOperatingSystem Then
            If Wow64DisableWow64FsRedirection(old) Then
                Process.Start(osk)
                Wow64EnableWow64FsRedirection(old)
            End If
        Else
            Process.Start(osk)
        End If
    End Sub


    Public Sub checkcol()

        Dim ESNO As String
        Dim SCREW1 As Double
        Dim SCREW2 As Double
        Dim SCREW3 As Double
        Dim SCREW4 As Double
        Dim SCREW5 As Double
        Dim SCREW6 As Double
        Dim SCREW7 As Double
        Dim SCREW8 As Double
        Dim SCREW9 As Double
        Dim SCREW10 As Double
        Dim SCREW11 As Double
        Dim SCREW12 As Double
        Dim IMIN As Double
        Dim IMAX As Double
        Dim EMIN As Double
        Dim EMAX As Double

        If SCREW1 > IMAX Or SCREW1 < IMIN Then
            LBLIN1.BackColor = Color.Red
            LBLIN1.Text = "NOK"
            TXTIN1.BackColor = Color.LightSeaGreen


        Else
            LBLIN1.BackColor = Color.Green
            LBLIN1.Text = " OK "
            TXTIN1.ReadOnly = True
        End If



        If SCREW3 > IMAX Or SCREW3 < IMIN Then
            LBLIN3.BackColor = Color.Red
            LBLIN3.Text = "NOK"
            TXTIN2.BackColor = Color.LightSeaGreen

        Else
            LBLIN3.BackColor = Color.Green
            LBLIN3.Text = " OK "
            TXTIN2.ReadOnly = True
        End If

        If SCREW5 > IMAX Or SCREW5 < IMIN Then
            LBLIN5.BackColor = Color.Red
            LBLIN5.Text = "NOK"
            TXTIN3.BackColor = Color.LightSeaGreen
        Else
            LBLIN5.BackColor = Color.Green
            LBLIN5.Text = "OK"
            TXTIN3.ReadOnly = True

        End If

        If SCREW7 > IMAX Or SCREW7 < IMIN Then
            LBLIN7.BackColor = Color.Red
            LBLIN7.Text = "NOK"
            TXTIN4.BackColor = Color.LightSeaGreen
        Else
            LBLIN7.BackColor = Color.Green
            LBLIN7.Text = "OK"
            TXTIN4.ReadOnly = True
        End If

        If SCREW9 > IMAX Or SCREW9 < IMIN Then
            LBLIN9.BackColor = Color.Red
            LBLIN9.Text = "NOK"
            TXTIN5.BackColor = Color.LightSeaGreen
        Else
            LBLIN9.BackColor = Color.Green
            LBLIN9.Text = "OK"
            TXTIN5.ReadOnly = True
        End If

        If SCREW11 > IMAX Or SCREW11 < IMIN Then
            LBLIN11.BackColor = Color.Red
            LBLIN11.Text = "NOK"
            TXTIN6.BackColor = Color.LightSeaGreen
        Else
            LBLIN11.BackColor = Color.Green
            LBLIN11.Text = "OK"
            TXTIN6.ReadOnly = True
        End If


        If SCREW2 > EMAX Or SCREW2 < EMIN Then
            LBLIN2.BackColor = Color.Red
            LBLIN2.Text = "NOK"
            TXTEX1.BackColor = Color.LightSeaGreen


        Else
            LBLIN2.BackColor = Color.Green
            LBLIN2.Text = "OK"
            TXTEX1.ReadOnly = True
        End If

        If SCREW4 > EMAX Or SCREW4 < EMIN Then
            LBLIN4.BackColor = Color.Red
            LBLIN4.Text = "NOK"
            TXTEX2.BackColor = Color.LightSeaGreen
        Else
            LBLIN4.BackColor = Color.Green
            LBLIN4.Text = "OK"
            TXTEX2.ReadOnly = True
        End If

        If SCREW6 > EMAX Or SCREW6 < EMIN Then
            LBLIN6.BackColor = Color.Red
            LBLIN6.Text = "NOK"
            TXTEX3.BackColor = Color.LightSeaGreen
        Else
            LBLIN6.BackColor = Color.Green
            LBLIN6.Text = "OK"
            TXTEX3.ReadOnly = True
        End If

        If SCREW8 > EMAX Or SCREW8 < EMIN Then
            LBLIN8.BackColor = Color.Red
            LBLIN8.Text = "NOK"
            TXTEX4.BackColor = Color.LightSeaGreen
        Else
            LBLIN8.BackColor = Color.Green
            LBLIN8.Text = "OK"
            TXTEX4.ReadOnly = True
        End If

        If SCREW10 > EMAX Or SCREW10 < EMIN Then
            LBLIN10.BackColor = Color.Red
            LBLIN10.Text = "NOK"
            TXTEX5.BackColor = Color.LightSeaGreen
        Else
            LBLIN10.BackColor = Color.Green
            LBLIN10.Text = "OK"
            TXTEX5.ReadOnly = True
        End If

        If SCREW12 > EMAX Or SCREW12 < EMIN Then
            LBLIN12.BackColor = Color.Red
            LBLIN12.Text = "NOK"
            TXTEX6.BackColor = Color.LightSeaGreen
        Else
            LBLIN12.BackColor = Color.Green
            LBLIN12.Text = "OK"
            TXTEX6.ReadOnly = True
        End If
    End Sub

    Public Sub colorchngfn(ESNO As String, SCREW1 As Double, SCREW2 As Double, SCREW3 As Double, SCREW4 As Double, SCREW5 As Double, SCREW6 As Double, SCREW7 As Double, SCREW8 As Double, SCREW9 As Double, SCREW10 As Double, SCREW11 As Double, SCREW12 As Double, IMIN As Double, IMAX As Double, EMIN As Double, EMAX As Double)

        If SCREW1 > IMAX Or SCREW1 < IMIN Then
            LBLIN1.BackColor = Color.Red
            'LBLIN1.Text = "NOK"
            'LBLIN1.Visible = False
            TXTIN1.BackColor = Color.LightSeaGreen


        Else
            LBLIN1.BackColor = Color.Green
            'LBLIN1.Text = " OK "
            'LBLIN1.Visible = False
            TXTIN1.ReadOnly = True
        End If



        If SCREW3 > IMAX Or SCREW3 < IMIN Then
            LBLIN3.BackColor = Color.Red
            'LBLIN3.Text = "NOK"
            'LBLIN3.Visible = False
            TXTIN2.BackColor = Color.LightSeaGreen

        Else
            LBLIN3.BackColor = Color.Green
            'LBLIN3.Text = " OK "
            'LBLIN3.Visible = False
            TXTIN2.ReadOnly = True
        End If

        If SCREW5 > IMAX Or SCREW5 < IMIN Then
            LBLIN5.BackColor = Color.Red
            'LBLIN5.Text = "NOK"
            'LBLIN5.Visible = False
            TXTIN3.BackColor = Color.LightSeaGreen
        Else
            LBLIN5.BackColor = Color.Green
            'LBLIN5.Text = "OK"
            'LBLIN5.Visible = False
            TXTIN3.ReadOnly = True

        End If

        If SCREW7 > IMAX Or SCREW7 < IMIN Then
            LBLIN7.BackColor = Color.Red
            'LBLIN7.Text = "NOK"
            'LBLIN7.Visible = False
            TXTIN4.BackColor = Color.LightSeaGreen
        Else
            LBLIN7.BackColor = Color.Green
            'LBLIN7.Text = "OK"
            'LBLIN7.Visible = False
            TXTIN4.ReadOnly = True
        End If

        If SCREW9 > IMAX Or SCREW9 < IMIN Then
            LBLIN9.BackColor = Color.Red
            'LBLIN9.Text = "NOK"
            'LBLIN9.Visible = False
            TXTIN5.BackColor = Color.LightSeaGreen
        Else
            LBLIN9.BackColor = Color.Green
            'LBLIN9.Text = "OK"
            'LBLIN9.Visible = False
            TXTIN5.ReadOnly = True
        End If

        If SCREW11 > IMAX Or SCREW11 < IMIN Then
            LBLIN11.BackColor = Color.Red
            'LBLIN11.Text = "NOK"
            'LBLIN11.Visible = False
            TXTIN6.BackColor = Color.LightSeaGreen
        Else
            LBLIN11.BackColor = Color.Green
            'LBLIN11.Text = "OK"
            'LBLIN11.Visible = False
            TXTIN6.ReadOnly = True
        End If

        If SCREW2 > EMAX Or SCREW2 < EMIN Then
            LBLIN2.BackColor = Color.Red
            'LBLIN2.Text = "NOK"
            'LBLIN2.Visible = False
            TXTEX1.BackColor = Color.LightSeaGreen

        Else
            LBLIN2.BackColor = Color.Green
            'LBLIN2.Text = "OK"
            'LBLIN2.Visible = False
            TXTEX1.ReadOnly = True
        End If

        If SCREW4 > EMAX Or SCREW4 < EMIN Then
            LBLIN4.BackColor = Color.Red
            'LBLIN4.Text = "NOK"
            'LBLIN4.Visible = False
            TXTEX2.BackColor = Color.LightSeaGreen
        Else
            LBLIN4.BackColor = Color.Green
            'LBLIN4.Text = "OK"
            'LBLIN4.Visible = False
            TXTEX2.ReadOnly = True
        End If

        If SCREW6 > EMAX Or SCREW6 < EMIN Then
            LBLIN6.BackColor = Color.Red
            LBLIN6.BackColor = Color.Red
            'LBLIN6.Text = "NOK"
            'LBLIN6.Visible = False
            TXTEX3.BackColor = Color.LightSeaGreen
        Else
            LBLIN6.BackColor = Color.Green
            'LBLIN6.Text = "OK"
            'LBLIN6.Visible = False
            TXTEX3.ReadOnly = True
        End If

        If SCREW8 > EMAX Or SCREW8 < EMIN Then
            LBLIN8.BackColor = Color.Red
            'LBLIN8.Text = "NOK"
            'LBLIN8.Visible = False
            TXTEX4.BackColor = Color.LightSeaGreen
        Else
            LBLIN8.BackColor = Color.Green
            'LBLIN8.Text = "OK"
            'LBLIN8.Visible = False
            TXTEX4.ReadOnly = True
        End If

        If SCREW10 > EMAX Or SCREW10 < EMIN Then
            LBLIN10.BackColor = Color.Red
            'LBLIN10.Text = "NOK"
            'LBLIN10.Visible = False
            TXTEX5.BackColor = Color.LightSeaGreen
        Else
            LBLIN10.BackColor = Color.Green
            'LBLIN10.Text = "OK"
            'LBLIN10.Visible = False
            TXTEX5.ReadOnly = True
        End If

        If SCREW12 > EMAX Or SCREW12 < EMIN Then
            LBLIN12.BackColor = Color.Red
            'LBLIN12.Text = "NOK"
            'LBLIN12.Visible = False
            TXTEX6.BackColor = Color.LightSeaGreen
        Else
            LBLIN12.BackColor = Color.Green
            'LBLIN12.Text = "OK"
            'LBLIN12.Visible = False
            TXTEX6.ReadOnly = True
        End If
    End Sub

    Private Sub LBLAUTO_Click(sender As Object, e As EventArgs) Handles LBLAUTO.Click

        LBLAUTO.Visible = False
        LBLMANUAL.Visible = True
        colorchng = colorchng + 1
        LBLIN12.BackColor = Color.Silver
        LBLIN11.BackColor = Color.Silver
        LBLIN10.BackColor = Color.Silver
        LBLIN9.BackColor = Color.Silver
        LBLIN8.BackColor = Color.Silver
        LBLIN7.BackColor = Color.Silver
        LBLIN6.BackColor = Color.Silver
        LBLIN5.BackColor = Color.Silver
        LBLIN4.BackColor = Color.Silver
        LBLIN3.BackColor = Color.Silver
        LBLIN2.BackColor = Color.Silver
        LBLIN1.BackColor = Color.Silver


        LBLIN1.Text = ""
        LBLIN2.Text = ""
        LBLIN3.Text = ""
        LBLIN4.Text = ""
        LBLIN5.Text = ""
        LBLIN6.Text = ""
        LBLIN7.Text = ""
        LBLIN8.Text = ""
        LBLIN9.Text = ""
        LBLIN10.Text = ""
        LBLIN11.Text = ""
        LBLIN12.Text = ""


        TXTEX6.Text = ""
        TXTIN6.Text = ""
        TXTEX5.Text = ""
        TXTIN5.Text = ""
        TXTEX4.Text = ""
        TXTIN4.Text = ""
        TXTEX3.Text = ""
        TXTIN3.Text = ""
        TXTEX2.Text = ""
        TXTIN2.Text = ""
        TXTEX1.Text = ""
        TXTIN1.Text = ""

        'If (colorchng Mod 2 = 0) Then
        '    '  Auto Function()
        'End If

        'If (colorchng Mod 2.0! = 0) Then
        '    ' manual function()

        'End If
    End Sub

    Private Sub LBLMANUAL_Click(sender As Object, e As EventArgs) Handles LBLMANUAL.Click
        colorchng = colorchng + 1
        LBLAUTO.Visible = True
        LBLMANUAL.Visible = False

        LTCHECK()
        'colorchngfn(ESNO, SCREW1, SCREW2, SCREW3, SCREW4, SCREW5, SCREW6, SCREW7, SCREW8, SCREW9, SCREW10, SCREW11, SCREW12, IMIN, IMAX, EMIN, EMAX)
        'If (colorchng Mod 2 = 0) Then
        'End If

        'If (colorchng Mod 2.0! = 0) Then
        '    ' manual function()

        '    TXTEX6.Text = ""
        '    TXTIN6.Text = ""
        '    TXTEX5.Text = ""
        '    TXTIN5.Text = ""
        '    TXTEX4.Text = ""
        '    TXTIN4.Text = ""
        '    TXTEX3.Text = ""
        '    TXTIN3.Text = ""
        '    TXTEX2.Text = ""
        '    TXTIN2.Text = ""
        '    TXTEX1.Text = ""
        '    TXTIN1.Text = ""
        'End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        'LTCHECK()
    End Sub



    Private Sub bedit_Click(sender As Object, e As EventArgs) Handles bedit.Click

        'If LBLIN2.BackColor = Color.Red Then
        'Dim a As Double
        'TXTEX1.ReadOnly = False


        'a = TXTEX1.Text
        ' Dim s As String
        'Dim z1 As String
        'z1 = TXTESNO.Text
        's = "update TCL_T_ROCKERHIGHT set screwhigh2=a where WHERE ESNO = '" & z1 & "'"
        'cmd = New OracleCommand(s, cn)
        'dr = cmd.ExecuteReader()
        'dr.Read()

        ' End If

        'Dim rhp As String
        'Dim IMIN As Double
        'Dim IMAX As Double
        'Dim EMIN As Double
        'Dim EMAX As Double

        'If rhp = "1" Then
        'IMIN = 0.01
        ' IMAX = 3.12
        '  EMIN = 1.2
        '   EMAX = 3.8
        'Else
        'IMIN = 0.01
        ' IMAX = 3.12
        '  EMIN = 0.5
        '   EMAX = 2.26
        'End If

        'Check And set color based on conditions
        'If LBLIN1.BackColor = Color.Red Then
        'TXTIN1.Text = InputBox("Please enter the new value for TXTIN1 textbox")
        'If TXTIN1.Text > IMAX Or TXTIN1.Text < IMIN Then
        ' LBLIN1.BackColor = Color.Red
        '  LBLIN1.Text = "NOK"
        '   TXTIN1.BackColor = Color.LightSeaGreen
        'Else
        '    LBLIN1.BackColor = Color.Green
        '     LBLIN1.Text = " OK "
        '      TXTIN1.ReadOnly = True
        '   End If
        'End If

        'If LBLIN3.BackColor = Color.Red Then
        'TXTIN2.Text = InputBox("Please enter the new value for TXTIN2 textbox")
        'If TXTIN2.Text > IMAX Or TXTIN2.Text < IMIN Then
        ' LBLIN3.BackColor = Color.Red
        '  LBLIN3.Text = "NOK"
        '   TXTIN2.BackColor = Color.LightSeaGreen
        'Else
        '    LBLIN3.BackColor = Color.Green
        '     LBLIN3.Text = " OK "
        '      TXTIN2.ReadOnly = True
        '   End If
        'End If

        'If LBLIN5.BackColor = Color.Red Then
        'TXTIN3.Text = InputBox("Please enter the new value for TXTIN3 textbox")
        'If TXTIN3.Text > IMAX Or TXTIN3.Text < IMIN Then
        '  LBLIN5.BackColor = Color.Red
        '   LBLIN5.Text = "NOK"
        '    TXTIN3.BackColor = Color.LightSeaGreen
        ' Else
        '    LBLIN5.BackColor = Color.Green
        '     LBLIN5.Text = " OK "
        '      TXTIN3.ReadOnly = True
        '   End If
        'End If

        'If LBLIN7.BackColor = Color.Red Then
        'TXTIN4.Text = InputBox("Please enter the new value for TXTIN4 textbox")
        'If TXTIN4.Text > IMAX Or TXTIN4.Text < IMIN Then
        ' LBLIN7.BackColor = Color.Red
        '  LBLIN7.Text = "NOK"
        '   TXTIN4.BackColor = Color.LightSeaGreen
        'Else
        '    LBLIN7.BackColor = Color.Green
        '     LBLIN7.Text = " OK "
        '      TXTIN4.ReadOnly = True
        '   End If
        'End If

        'If LBLIN9.BackColor = Color.Red Then
        'TXTIN5.Text = InputBox("Please enter the new value for TXTIN5 textbox")
        'If TXTIN5.Text > IMAX Or TXTIN5.Text < IMIN Then
        '  LBLIN9.BackColor = Color.Red
        '   LBLIN9.Text = "NOK"
        '   TXTIN5.BackColor = Color.LightSeaGreen
        'Else
        '    LBLIN9.BackColor = Color.Green
        '     LBLIN9.Text = " OK "
        '      TXTIN5.ReadOnly = True
        '   End If
        'End If

        'If LBLIN11.BackColor = Color.Red Then
        'TXTIN6.Text = InputBox("Please enter the new value for TXTIN6 textbox")
        'If TXTIN6.Text > IMAX Or TXTIN6.Text < IMIN Then
        ' LBLIN11.BackColor = Color.Red
        '  LBLIN11.Text = "NOK"
        '   TXTIN6.BackColor = Color.LightSeaGreen
        'Else
        '    LBLIN11.BackColor = Color.Green
        '     LBLIN11.Text = " OK "
        '      TXTIN6.ReadOnly = True
        '   End If
        'End If

        'If LBLIN2.BackColor = Color.Red Then
        'TXTEX1.Text = InputBox("Please enter the new value for TXTEX1 textbox")
        ' If TXTEX1.Text > IMAX Or TXTEX1.Text < IMIN Then
        ' LBLIN2.BackColor = Color.Red
        '  LBLIN2.Text = "NOK"
        '   TXTEX1.BackColor = Color.LightSeaGreen
        'Else
        '    LBLIN2.BackColor = Color.Green
        '     LBLIN2.Text = " OK "
        '      TXTEX1.ReadOnly = True
        '   End If
        'End If

        'If LBLIN4.BackColor = Color.Red Then
        'TXTEX2.Text = InputBox("Please enter the new value for TXTEX2 textbox")
        'If TXTEX2.Text > IMAX Or TXTEX2.Text < IMIN Then
        ' LBLIN4.BackColor = Color.Red
        '  LBLIN4.Text = "NOK"
        '   TXTEX2.BackColor = Color.LightSeaGreen
        'Else
        '    LBLIN4.BackColor = Color.Green
        '     LBLIN4.Text = " OK "
        '      TXTEX2.ReadOnly = True
        '   End If
        'End If

        'If LBLIN6.BackColor = Color.Red Then
        'TXTEX3.Text = InputBox("Please enter the new value for TXTEX3 textbox")
        'If TXTEX3.Text > IMAX Or TXTEX3.Text < IMIN Then
        ' LBLIN6.BackColor = Color.Red
        '  LBLIN6.Text = "NOK"
        '   TXTEX3.BackColor = Color.LightSeaGreen
        'Else
        '    LBLIN6.BackColor = Color.Green
        '     LBLIN6.Text = " OK "
        '      TXTEX3.ReadOnly = True
        '   End If
        'End If

        'If LBLIN8.BackColor = Color.Red Then
        'TXTEX4.Text = InputBox("Please enter the new value for TXTEX4 textbox")
        'If TXTEX4.Text > IMAX Or TXTEX4.Text < IMIN Then
        ' LBLIN8.BackColor = Color.Red
        '  LBLIN8.Text = "NOK"
        '   TXTEX4.BackColor = Color.LightSeaGreen
        'Else
        '    LBLIN8.BackColor = Color.Green
        '     LBLIN8.Text = " OK "
        '      TXTEX4.ReadOnly = True
        '   End If
        'End If

        'If LBLIN10.BackColor = Color.Red Then
        'TXTEX5.Text = InputBox("Please enter the new value for TXTEX5 textbox")
        ' If TXTEX5.Text > IMAX Or TXTEX5.Text < IMIN Then
        ' LBLIN10.BackColor = Color.Red
        '  LBLIN10.Text = "NOK"
        '   TXTEX5.BackColor = Color.LightSeaGreen
        'Else
        '    LBLIN10.BackColor = Color.Green
        '     LBLIN10.Text = " OK "
        '      TXTEX5.ReadOnly = True
        '   End If
        'End If

        ' If LBLIN12.BackColor = Color.Red Then
        'TXTEX6.Text = InputBox("Please enter the new value for TXTEX6 textbox")
        'If TXTEX6.Text > IMAX Or TXTEX6.Text < IMIN Then
        'LBLIN12.BackColor = Color.Red
        'LBLIN12.Text = "NOK"
        'TXTEX6.BackColor = Color.LightSeaGreen
        'Else
        'LBLIN12.BackColor = Color.Green
        'LBLIN12.Text = " OK "
        'TXTEX6.ReadOnly = True
        'End If
        'End If


        'For i As Integer = 1 To 12
        'If LBLIN(i).BackColor = Color.Red Then
        'TXTEX(i).Text =
        'If TXTEX(i).Text > EMAX Or TXTEX1.Text < EMIN Then
        'LBLIN(i).BackColor = Color.Red
        'LBLIN(i).Text = "NOK"
        'TXTEX(i).BackColor = Color.LightSeaGreen
        'Else
        'LBLIN(i).BackColor = Color.Green
        ' LBLIN(i).Text = "OK"
        'TXTEX(i).ReadOnly = True
        'End If

        'ElseIf LBLIN(i).BackColor = Color.Red Then
        'IN(i).Text = "NOK"
        'TXTIN(i).BackColor = Color.LightSeaGreen

        'Else
        'LBLIN(i).BackColor = Color.Green
        'LBLIN(i).Text = " OK "
        'TXTIN(i).ReadOnly = True
        ' End If
        'End If
        ' Next
    End Sub

    Private Sub TXTIN1_TextChanged(sender As Object, e As EventArgs) Handles TXTIN1.TextChanged
        If LBLIN1.BackColor = Color.Red Then
            btnnop = 1
            Form1.Show()
        End If

    End Sub

    Private Sub TXTEX1_TextChanged(sender As Object, e As EventArgs) Handles TXTEX1.TextChanged
        If LBLIN2.BackColor = Color.Red Then
            btnnop = 2
            Form1.Show()
        End If

    End Sub

    Private Sub TXTIN2_TextChanged(sender As Object, e As EventArgs) Handles TXTIN2.TextChanged
        If LBLIN3.BackColor = Color.Red Then
            btnnop = 3
            Form1.Show()
        End If

    End Sub

    Private Sub TXTEX2_TextChanged(sender As Object, e As EventArgs) Handles TXTEX2.TextChanged
        If LBLIN4.BackColor = Color.Red Then
            btnnop = 4
            Form1.Show()
        End If

    End Sub

    Private Sub TXTIN3_TextChanged(sender As Object, e As EventArgs) Handles TXTIN3.TextChanged
        If LBLIN5.BackColor = Color.Red Then
            btnnop = 5
            Form1.Show()
        End If

    End Sub

    Private Sub TXTEX3_TextChanged(sender As Object, e As EventArgs) Handles TXTEX3.TextChanged
        If LBLIN6.BackColor = Color.Red Then
            btnnop = 6
            Form1.Show()
        End If

    End Sub

    Private Sub TXTIN4_TextChanged(sender As Object, e As EventArgs) Handles TXTIN4.TextChanged
        If LBLIN7.BackColor = Color.Red Then
            btnnop = 7
            Form1.Show()
        End If

    End Sub

    Private Sub TXTEX4_TextChanged(sender As Object, e As EventArgs) Handles TXTEX4.TextChanged
        If LBLIN8.BackColor = Color.Red Then
            btnnop = 8
            Form1.Show()
        End If

    End Sub

    Private Sub TXTIN5_TextChanged(sender As Object, e As EventArgs) Handles TXTIN5.TextChanged
        If LBLIN9.BackColor = Color.Red Then
            btnnop = 9
            Form1.Show()
        End If

    End Sub

    Private Sub TXTEX5_TextChanged(sender As Object, e As EventArgs) Handles TXTEX5.TextChanged
        If LBLIN10.BackColor = Color.Red Then
            btnnop = 10
            Form1.Show()
        End If

    End Sub

    Private Sub TXTIN6_TextChanged(sender As Object, e As EventArgs) Handles TXTIN6.TextChanged
        If LBLIN11.BackColor = Color.Red Then
            btnnop = 11
            Form1.Show()
        End If

    End Sub

    Private Sub TXTEX6_TextChanged(sender As Object, e As EventArgs) Handles TXTEX6.TextChanged
        If LBLIN12.BackColor = Color.Red Then
            btnnop = 12
            Form1.Show()
        End If

    End Sub

    Private Sub LBCPRESENT_Click(sender As Object, e As EventArgs) Handles LBCPRESENT.Click
        colorchngfn(ESNO, SCREW1, SCREW2, SCREW3, SCREW4, SCREW5, SCREW6, SCREW7, SCREW8, SCREW9, SCREW10, SCREW11, SCREW12, IMIN, IMAX, EMIN, EMAX)
    End Sub

End Class

