Imports System.Xml

Public Class Form1
    Private objFits As clsFits
    Private objAutoTest As clsAutoTest
    Dim cn As New ADODB.Connection()
    Dim cnAutoTest As New ADODB.Connection()
    Dim objInI As clsINI
    Dim vWorkingDir As String
    Dim vServiceURL As String
    Dim vOutPutFolder As String
    Dim vLastID As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnDatabase.Click
        'Dim objFits As New clsFits
        Try
            If btnDatabase.Text = "&Disconnect Database." Then
                cn.Close()
                cnAutoTest.Close()
                initialControl()
            Else
                tssDatabase.Text = "Connecting Database..." : Application.DoEvents()

                objFits = New clsFits
                With objFits
                    .user = objInI.GetString("database", "user", "")
                    .password = objInI.GetString("database", "password", "")
                    .server = objInI.GetString("database", "server", "")
                    .database = objInI.GetString("database", "database", "")
                    cn = .connect()
                    tssDatabase.Text = "(" & .server & "/" & .database & ")Database connected." : Application.DoEvents()
                End With

                'Open AutoTest database
                objAutoTest = New clsAutoTest
                With objAutoTest
                    .user = objInI.GetString("test database", "user", "")
                    .password = objInI.GetString("test database", "password", "")
                    .server = objInI.GetString("test database", "server", "")
                    .database = objInI.GetString("test database", "database", "")
                    cnAutoTest = .connect()
                End With

                btnDatabase.Text = "&Disconnect Database."
                btnImport.Enabled = True
            End If

        Catch ex As Exception
            MsgBox("Unable to connect database!!!" & vbCrLf & _
                "Because " & ex.Message, MsgBoxStyle.Critical, "Unable to connect database")
            tssDatabase.Text = "Database error!" : Application.DoEvents()
            initialControl()
        End Try




    End Sub


   


    '##Look at vw_SMTUnitHistoryTracking First.
    Sub ExportData()
        Dim vBullEyesObj As New clsBullEyes
        With vBullEyesObj
            'Query Data
            Dim rs As New ADODB.Recordset
            Dim vNewFromDate As Date = CDate(lblFrom.Text).AddSeconds(1)
            Dim vDateFrom As String = vNewFromDate.ToString
            Dim vDateTo As String = lblTo.Text


            rs = objAutoTest.getUUTResult(vLastID)
          

            If rs.RecordCount = 0 Then
                lblLastDate.Text = Now() : Application.DoEvents() ' lblNextRun.Text 
                GoTo NoSN
            End If

            'Initial Object
            Dim objFITSDLL As New FITSDLL.clsDB
            Dim vInitResult As Boolean

            With objFITSDLL
                vInitResult = .fn_InitDB("*", "", "2.9", "dbAcacia")
            End With
            If Not vInitResult Then
                MsgBox("Unable to initial FITSDLL", MsgBoxStyle.Critical, "Unable to initial FITDLL")
                Exit Sub
            End If
            '--------------
            Dim vTempTimeOut As String
            Do While Not rs.EOF

                tssStatus.Text = "Importing......." & rs.AbsolutePosition & "/" & rs.RecordCount : Application.DoEvents()
                vTempTimeOut = rs.Fields("start_date_time").Value
                .datetimeout = vTempTimeOut
                '1)=========HandShake============
                'Get Model by Serial number
                Dim vModel As String
                Dim vKittingStation As String = "100"
                Dim vExeStation As String = ""
                Dim vHandShake As String
                Dim vSnParamName As String = "Serial No."
                Dim vSn As String = rs.Fields("uut_serial_number").Value
                Dim vHWPartFIT As String = ""
                Dim vUutID As String = rs.Fields("id").Value
                Dim vProcess As String = rs.Fields("process").Value
                Dim vDeviceTypeFit As String
                Dim vDeviceTypeATS As String
                lblCurrentID.Text = vUutID
                vLastID = vUutID

                vModel = objFits.getParameters(vSn, "1003")
                vHWPartFIT = objFits.getParameters(vSn, "10112")


                vDeviceTypeFit = objFits.getParameters(vSn, "1204")
                vDeviceTypeATS = objAutoTest.getDeviceType(vSn, "DCP")
                Dim vDeviceTypeCheck As Boolean = False
                vDeviceTypeCheck = IIf(vDeviceTypeATS = vDeviceTypeFit, True, False)

                

                'vModel = objFITS.fn_Query(txtModel.Text, vKittingStation, "1", rs.Fields(""), "Model")
                Select Case vModel
                    Case "Acadia"
                        Select Case vProcess
                            Case "DCP" : vExeStation = "180"
                            Case "DBI" : vExeStation = "190"
                            Case "FTU" : vExeStation = "230"
                            Case "FPT" : vExeStation = "290"
                            Case "FAT" : vExeStation = "292"
                        End Select

                    Case "GLACIER"
                        Select Case vProcess
                            Case "DCP" : vExeStation = "118"
                            Case "DBI" : vExeStation = "120"
                            Case "FTU" : vExeStation = "126"
                            Case "FPT" : vExeStation = "144"
                            Case "FAT" : vExeStation = "146"
                        End Select

                    Case "SFF[ORION]"
                        Select Case vProcess
                            Case "DCP" : vExeStation = "116"
                            Case "DBI" : vExeStation = "118"
                            Case "FTU" : vExeStation = "128"
                            Case "FAT" : vExeStation = "147"
                        End Select


                        'Module Level
                    Case "CFP"
                        Select Case vProcess
                            Case "OBS" : vExeStation = "1400"
                            Case "FVT" : vExeStation = "1600"
                            Case "EPT" : vExeStation = "1650"
                            Case "ESS" : vExeStation = "1700"
                            Case "EXS" : vExeStation = "1900"
                            Case "CFG" : vExeStation = "1950"
                        End Select
                    Case "CFP GLACIER"
                        Select Case vProcess
                            Case "OBS" : vExeStation = "1400"
                            Case "FVT" : vExeStation = "1600"
                            Case "EPT" : vExeStation = "1650"
                            Case "ESS" : vExeStation = "1700"
                            Case "EXS" : vExeStation = "1900"
                            Case "CFG" : vExeStation = "1950"
                        End Select

                    Case "CFP2"
                        Select Case vProcess
                            Case "OBS" : vExeStation = "101"
                            Case "FVT" : vExeStation = "150"
                            Case "EPT" : vExeStation = "160"
                            Case "ESS" : vExeStation = "170"
                            Case "FST" : vExeStation = "180"
                            Case "CFG" : vExeStation = "200"
                        End Select

                    Case "AC400"
                        Select Case vProcess
                            Case "OBS" : vExeStation = "190"
                            Case "FVT" : vExeStation = "250"
                            Case "EPT" : vExeStation = "260"
                            Case "ESS" : vExeStation = "270"
                            Case "FST" : vExeStation = "280"
                            Case "EXS" : vExeStation = "290"
                            Case "CFG" : vExeStation = "325"
                        End Select
                End Select


                vHandShake = objFITSDLL.fn_Handshake(vModel, vExeStation, "2.9", vSn)
                If vHandShake <> "True" Then
                    GoTo nextSN
                End If




                Dim vuutStatus As String = rs.Fields("uut_status").Value
                Dim vLoginName As String = rs.Fields("user_login_name").Value
                Dim vProductCode As String = rs.Fields("product_code").Value
                Dim vFixtureID As String = rs.Fields("Fixture_ID").Value
                Dim vStationID As String = rs.Fields("Station_ID").Value
                Dim vDateTime As String = rs.Fields("start_date_time").Value
                Dim vExeTime As String = rs.Fields("execution_time").Value
                Dim vMode As String = rs.Fields("Mode").Value
                Dim vTestCount As String = rs.Fields("test_count").Value
                Dim vTestSocketIndex As String = rs.Fields("TEST_SOCKET_INDEX").Value
                Dim vTpsRev As String = rs.Fields("TPS_REV").Value
                Dim vHWRev As String = rs.Fields("HW_REV").Value
                Dim vFWRev As String = rs.Fields("FW_REV").Value
                Dim vResult As String = rs.Fields("uut_status").Value
                Dim vHWPart As String = rs.Fields("HW_PART_NUMBER").Value
                Dim vRemark As String = ""
                Dim vTopBomRev As String = ""
                Dim vDisposCode As String = ""


                vTopBomRev = objFITSDLL.fn_Query(vModel, vExeStation, "2.9", vSn, "PART_REV", ",")

                If vResult = "Terminated" Then
                    'Keep Terminated ID
                    'vRemark = getTerminatedText(vSn, vProcess, vUutID)
                    'vDisposCode = "Terminated"
                    TerminatedLog(Now() & "--" & vSn & "," & vModel & "," & vProcess & "," & vUutID)
                    GoTo nextSN
                End If

                If vResult = "Failed" Then
                    vRemark = getFailedText(vSn, vProcess, vUutID)
                    If vRemark.Length > 0 Then
                        vDisposCode = vRemark.Split("=")(0).ToUpper
                    Else
                        'Terminated.
                        TerminatedLog(Now() & "--" & vSn & "," & vModel & "," & vProcess & "," & vUutID)
                        GoTo nextSN
                        'vRemark = "Not Found failed record" 'getTerminatedText(vSn, vProcess, vUutID)
                        'vDisposCode = "Terminated"
                    End If
                End If




                Dim vTest1 As String = "FBN Serial No|Login Name|Product Code|" & _
                                   "Fixture ID|Station ID|Date/Time|Execute Time|Mode|TEST_COUNT|" & _
                                   "TEST_SOCKET_INDEX|TPS_REV|HW_REV|FW_REV|Result|Remark|TOP BOM REV.|" & _
                                   "HW_PART_NUMBER  (FITS)|HW_PART_NUMBER|Device Type (FITS)|Device Type (ATS)"

                Dim vTest2 As String = vSn & "|" & vLoginName & "|" & vProductCode & "|" & _
                                vFixtureID & "|" & vStationID & "|" & vDateTime & "|" & vExeTime & "|" & vMode & "|" & vTestCount & "|" & _
                                vTestSocketIndex & "|" & vTpsRev & "|" & vHWRev & "|" & vFWRev & "|" & IIf(vResult = "Passed", "PASS", vDisposCode) & "|" & _
                                Mid(vRemark, 1, 200) & "|" & vTopBomRev & "|" & _
                                vHWPartFIT & "|" & vHWPart & "|" & vDeviceTypeFit & "|" & vDeviceTypeATS
                Dim vCheckIn As String
                Dim vCheckOut As String
                vLastID = vUutID
                Select Case vHandShake
                    Case "True"
                        'Need both check in and Out
                        vCheckIn = objFITSDLL.fn_Log(vModel, vExeStation, "0", "FBN Serial No", vSn)
                        vCheckOut = objFITSDLL.fn_Log(vModel, vExeStation, "1", vTest1, vTest2, "|")
                        Log(Now() & "--" & vSn & "--" & vModel & "--" & vProcess & "--" & vExeStation & "--" & vLastID & "--" & vCheckOut & "--" & IIf(vResult = "Passed", "PASS", vDisposCode))
                    Case vHandShake.Contains("in-processing in " & vExeStation)
                        'already check-in,Need only Check-out
                        vCheckOut = objFITSDLL.fn_Log(vModel, vExeStation, "1", vTest1, vTest2, "|")
                        Log(Now() & "--" & vSn & "--" & vModel & "--" & vProcess & "--" & vExeStation & "--" & vLastID & "--" & vCheckOut & "--" & IIf(vResult = "Passed", "PASS", vDisposCode))
                    Case Else
                        GoTo nextSN
                End Select
                '---------
                '2)=========Check In =========================





                'Dim vTest1 As String = "FBN Serial No|Status|Login Name|Product Code|" & _
                '    "Fixture ID|Station ID|Date/Time|Execute Time|Mode|TEST_COUNT|" & _
                '    "TEST_SOCKET_INDEX|TPS_REV|HW_REV|FW_REV|Result|Remark|TOP BOM REV."
                'Dim vTest2 As String = "" & "|PASS|nyotpanya|SIPH100-Q01-001|143210059|ATE_SIPH_057|2016-11-29 09:43:56.000|51.6991105|Production|1|1|1.1.1|27|-"



                ''.operation = rs.Fields("operation").Value

                ''.datetimeout = IIf(IsDBNull(rs.Fields("date_time").Value), "", rs.Fields("date_time").Value)

                ''If Not (.operation = "3261") Then
                ''    GoTo nextSN
                ''End If

                ''.serialnumber = rs.Fields("serial_no").Value
                ''.sn_attr_code = rs.Fields("sn_attr_code").Value
                ''.trans_seq = rs.Fields("trans_seq").Value



                ''-----Get PCB Info -----
                ' ''History (only EBT)
                ''Dim vPCBHistory As ADODB.Recordset
                'Dim vPCBSN As String = rs.Fields("serial_no").Value

                ''vPCBHistory = objFits.getPCBAlist(vPCBSN)
                ' ''Comp.onent tracking (Only value is not 0)
                ''If vPCBHistory.RecordCount = 0 Then
                ''    GoTo nextSN
                ''End If
                ''vPCBHistory.MoveLast()

                ''Dim vEBTstation As String = vPCBHistory.Fields("operation").Value
                ''.operation = vEBTstation

                ''Dim vDateOutPCB As String = vPCBHistory.Fields("date_time").Value
                'Dim vTesterPCB As String = "ATE_AC100M_062"
                ''-----------------------
                ''Modify all Parameter
                'Dim vTempTimeOut As String = .datetimeout

                '.serialnumber = vPCBSN
                '.workorder = rs.Fields("workorder").Value
                '.model = "PCBA" 'set family

                '.partnumber = rs.Fields("part_no").Value
                '.operation = rs.Fields("operation").Value
                '.operationname = IIf(IsDBNull(rs.Fields("description").Value), "", rs.Fields("description").Value)
                '.buildtype = "PROD" 'rs.Fields("buildtype").Value
                '.runtype = "100" 'rs.Fields("runtype").Value
                '.employee = rs.Fields("emp_no").Value
                '.sn_attr_code = "1001" 'rs.Fields("sn_attr_code").Value
                '.trans_seq = rs.Fields("trans_seq").Value 'vPCBHistory.Fields("date_time").Value
                '.datetimein = rs.Fields("date_time").Value
                '.datetimeout = rs.Fields("date_time").Value
                '.shift = "DAY" 'rs.Fields("shift").Value
                '.tester = vTesterPCB 'rs.Fields("equip_id").Value
                '.outputPath = vWorkingDir
                '.result = IIf(IsDBNull(rs.Fields("result").Value), "PASS", rs.Fields("result").Value)
                '.disposecode = IIf(IsDBNull(rs.Fields("result").Value), "", rs.Fields("result").Value)
                '.next_operation = "382" ' objFits.getNextStation(.serialnumber, .operation, .trans_seq, .model)
                ''End Modify




                ''Get Testing Data
                ''1)get Process from BullsEye -- by Station.
                'Dim vProcess As String = requestData(vServiceURL & "production/station/" & .operation & "/" & .model & "/")
                ''2)get Measurement data.
                'Dim vTestDataRst As New ADODB.Recordset


                'If vProcess <> "" And vProcess <> "None" Then
                '    vTestDataRst = objAutoTest.getTestData(.serialnumber, vProcess, .tester, .datetimeout)
                '    '.tester = vTestDataRst.Fields("STATION_ID").Value
                'Else
                '    vTestDataRst = Nothing
                'End If

                ''Get Component--
                ''Dim vRstComponent As ADODB.Recordset
                ''vRstComponent = objFits.getComponentData(.serialnumber)
                ''If vRstComponent.RecordCount > 0 Then
                ''    tssStatus.Text = "Uploading " & .serialnumber & " " & rs.AbsolutePosition & "/" & rs.RecordCount : Application.DoEvents()
                ''End If
                ''---------------



                ''.makeXML(Nothing, vTestDataRst)
                ''.makeXML(objFits.getParameters(.serialnumber, .sn_attr_code, .trans_seq), vTestDataRst, vRstComponent)



                '' uploadData(.outputfile)
                .datetimeout = vTempTimeOut

nextSN:

                lblLastDate.Text = .datetimeout : Application.DoEvents()
                '---save last date to INI file---
                objInI.WriteString("Last execution", "id", vLastID)
                objInI.WriteString("Last execution", "date", .datetimeout)
                '--------------------------------
                rs.MoveNext()
            Loop
        End With
NoSN:
        '---Update From/To date
        lblFrom.Text = lblLastDate.Text
        '---save last date to INI file---
        objInI.WriteString("Last execution", "date", lblLastDate.Text)
        '--------------------------------
        lblTo.Text = getDateTo(lblLastDate.Text)
        lblNextRun.Text = Now.AddMinutes(Val(objInI.GetString("import", "interval", "")))

    End Sub

    Function getFailedText(vSn As String, vProcess As String, vID As String) As String
        getFailedText = objAutoTest.getTestDataString(vSn, vProcess, vID)
        'Dim vRst As ADODB.Recordset
        'vRst = objAutoTest.getTestData(vSn, vProcess, vID)
        'vRst.Filter = "status <> 'Passed'"
        'Dim vStrResult As String

        'If vRst.RecordCount > 0 Then
        '    vStrResult = vRst.Fields("step_name").Value & "=" & _
        '        IIf(IsDBNull(vRst.Fields("data").Value), "", vRst.Fields("data").Value) & " " & _
        '        IIf(IsDBNull(vRst.Fields("units").Value), "", vRst.Fields("units").Value) & _
        '        "(" & _
        '        IIf(IsDBNull(vRst.Fields("low_limit").Value), "", vRst.Fields("low_limit").Value) & _
        '        "/" & _
        '        IIf(IsDBNull(vRst.Fields("high_limit").Value), "", vRst.Fields("high_limit").Value) & _
        '        ")"
        '    getFailedText = vStrResult
        'Else
        '    vStrResult = ""
        '    getFailedText = vStrResult
        'End If
    End Function


    Function getTerminatedText(vSn As String, vProcess As String, vID As String) As String
        Dim vRst As ADODB.Recordset
        vRst = objAutoTest.getStepTestData(vSn, vProcess, vID)
        vRst.Filter = "status = 'Failed'"
        Dim vStrResult As String
        If vRst.RecordCount = 0 Then
            vRst.Filter = ""
            vRst.Filter = "report_text <> ''"
            If vRst.RecordCount = 0 Then
                vRst.Filter = ""
                vRst.MoveLast()
                vStrResult = "Terminated at " & vRst.Fields("step_name").Value & vbCrLf & _
                        "Error Message : " & IIf(IsDBNull(vRst.Fields("report_text").Value), "", vRst.Fields("report_text").Value)
            Else
                vStrResult = "Terminated at " & vRst.Fields("step_name").Value & vbCrLf & _
                        "Error Message : " & IIf(IsDBNull(vRst.Fields("report_text").Value), "", vRst.Fields("report_text").Value)
            End If
            
        Else
            vRst.MoveLast()
            vStrResult = "Terminated at " & vRst.Fields("step_name").Value & vbCrLf & _
                        "Error Message : " & IIf(IsDBNull(vRst.Fields("report_text").Value), "", vRst.Fields("report_text").Value)
        End If
        Return vStrResult
    End Function


    'Function getDeviceTypeATS(vSn As String) As String
    '    Dim vRst As ADODB.Recordset
    '    vRst = objAutoTest.getTestData(vSn, 'DCP', vID)
    '    vRst.Filter = "status <> 'Passed'"
    '    If vRst.RecordCount > 0 Then
    '        Return vRst.Fields("step_name").Value & "=" & vRst.Fields("data").Value & " " & vRst.Fields("units").Value & _
    '            "(" & vRst.Fields("low_limit").Value & "-" & vRst.Fields("high_limit").Value & ")"
    '    Else
    '        Return ""
    '    End If
    'End Function

    Sub Log(vMessage As String)
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(vOutPutFolder & Now().ToString("yyyy-M-dd") & ".txt", True) '"2016-11-30.txt"
        file.WriteLine(vMessage)
        file.Close()
    End Sub

    Sub TerminatedLog(vMessage As String)
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(vOutPutFolder & Now().ToString("yyyy-M-dd") & "_Terminated.txt", True) '"2016-11-30.txt"
        file.WriteLine(vMessage)
        file.Close()
    End Sub

    Function uploadData(vFile As String) As Boolean
        'HTTP variable
        Dim myHTTP As New MSXML.XMLHTTPRequest
        Dim doc As XmlDocument = New XmlDocument()
        Dim strReturn As String
        uploadData = False
        Try
            doc.Load(vFile)
            With myHTTP
                .open("Post", vServiceURL & "production/fits/upload/", False)
                .setRequestHeader("Content-Type", "text/xml")
                .send(doc.InnerXml) '+64
                strReturn = .responseText
                If strReturn <> """Successful""" Then
                    MsgBox(strReturn)
                End If
            End With
            Return True
        Catch ex As Exception

            MsgBox("Unable to upload XML!!!" & vbCrLf & _
                "Because " & ex.Message, MsgBoxStyle.Critical, "Unable to upload XML")
        End Try

    End Function



    Function requestData(vURL As String) As String
        'HTTP variable
        Dim myHTTP As New MSXML.XMLHTTPRequest
        'Dim doc As XmlDocument = New XmlDocument()
        Dim strReturn As String
        'uploadData = False
        Try
            'doc.Load(vFile)
            With myHTTP
                .open("Post", vURL, False)
                .setRequestHeader("Content-Type", "text/xml")
                .send("") '+64
                strReturn = .responseText
                'If strReturn <> """Successful""" Then
                '    MsgBox(strReturn)
                'End If
            End With
            Return strReturn
        Catch ex As Exception

            MsgBox("Unable to upload XML!!!" & vbCrLf & _
                "Because " & ex.Message, MsgBoxStyle.Critical, "Unable to upload XML")
            Return "Error"
        End Try

    End Function

    Sub initialControl()
        btnDatabase.Text = "&Connect Database" : btnDatabase.Enabled = True
        btnImport.Text = "&Start Import data" : btnImport.Enabled = False
        tssDatabase.Text = "Program ready"
        lblLastDate.Text = objInI.GetString("Last execution", "date", "")
        lblFrom.Text = lblLastDate.Text
        lblTo.Text = getDateTo(lblLastDate.Text)
        gbImport.Text = "Import Details"
        lblPeriod.Text = objInI.GetString("import", "range", "")
        lblLoop.Text = objInI.GetString("import", "interval", "")
        lblNextRun.Text = Now
        vServiceURL = objInI.GetString("service", "url", "")
        vOutPutFolder = objInI.GetString("path", "working dir", "")
        vLastID = objInI.GetString("Last execution", "id", "")
    End Sub

    Function getDateTo(vDateFrom As String) As String
        Dim vRange As String = objInI.GetString("import", "range", "")
        Dim date2 As Date = vDateFrom
        getDateTo = date2.AddHours(Val(vRange))
    End Function


    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        If MsgBox("Are you sure to close program", vbQuestion + vbYesNo, "Confrim close program") = vbYes Then
            On Error Resume Next
            objFits.disconnect()
            Me.Close()
        End If

    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        objInI = New clsINI(Application.StartupPath & "\import.ini")
        Timer1.Interval = (Val(objInI.GetString("import", "interval", ""))) * 1000 * 60
        Timer1.Enabled = False

        vWorkingDir = objInI.GetString("path", "working dir", "")
        If vWorkingDir = "" Then
            vWorkingDir = Application.StartupPath & "\"
        End If



        initialControl()
        lblCurrentID.Text = vLastID
        Me.Text = Me.Text + " version : " + Application.ProductVersion.Trim()
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        'First import then using Timer.
        If btnImport.Text = "&Start Import data" Then
            'Added by Chutchai on Dec 7,2016
            'To verify all database connection
            'If objFits.database .

            ExportData()
            Timer1.Enabled = True
            btnImport.Text = "&Stop Import data"
            Timer2.Enabled = True
            '--Update From/To date

        Else
            Timer1.Enabled = False
            Timer2.Enabled = False
            btnImport.Text = "&Start Import data"
        End If

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        Timer2.Enabled = False
        ExportData()
        lblNextRun.Text = Now.AddMinutes(Val(objInI.GetString("import", "interval", "")))
        Timer1.Enabled = True
        Timer2.Enabled = True
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        tssStatus.Text = "Next run (remaining time) -->" & Print_Remaining_Time(CDate(lblNextRun.Text))
    End Sub

    Public Function Print_Remaining_Time(EndTime As DateTime)
        'If EndTime.ToString = "01/01/0001 0:00:00" Then
        '    EndTime = Now
        '    EndTime = EndTime.AddMilliseconds(Time_Out - 1000)
        'End If
        Dim RemainingTime As TimeSpan
        RemainingTime = Now().Subtract(EndTime)
        Return String.Format("{0:00}:{1:00}:{2:00}", CInt(Math.Floor(RemainingTime.TotalHours)) Mod 60, CInt(Math.Floor(RemainingTime.TotalMinutes)) Mod 60, CInt(Math.Floor(RemainingTime.TotalSeconds)) Mod 60).Replace("-", "")
    End Function

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        uploadData("C:\Users\chutchais\Documents\Visual Studio 2013\Projects\test.xml")
    End Sub
End Class
