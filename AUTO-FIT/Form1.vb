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
    Dim vAutoRun As Boolean

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnDatabase.Click
        'Dim objFits As New clsFits

        Try
            If btnDatabase.Text = "&Disconnect Database." Then
                cn.Close()
                cnAutoTest.Close()
                initialControl()
            Else
                tssDatabase.Text = "Connecting Database..." : Application.DoEvents()
                cn = objFits.connect()
                cnAutoTest = objAutoTest.connect()
                'objFits = New clsFits
                'With objFits
                '    .user = objInI.GetString("database", "user", "")
                '    .password = objInI.GetString("database", "password", "")
                '    .server = objInI.GetString("database", "server", "")
                '    .database = objInI.GetString("database", "database", "")
                '    cn = .connect()
                '    tssDatabase.Text = "(" & .server & "/" & .database & ")Database connected." : Application.DoEvents()
                'End With

                ''Open AutoTest database
                'objAutoTest = New clsAutoTest
                'With objAutoTest
                '    .user = objInI.GetString("test database", "user", "")
                '    .password = objInI.GetString("test database", "password", "")
                '    .server = objInI.GetString("test database", "server", "")
                '    .database = objInI.GetString("test database", "database", "")
                '    cnAutoTest = .connect()
                'End With

                btnDatabase.Text = "&Disconnect Database."
                btnImport.Enabled = True
            End If

        Catch ex As Exception
            Log(Now() & "--Unable to connect database!!!")
            tssDatabase.Text = "Database error!" : Application.DoEvents()
            initialControl()
        End Try




    End Sub


   


    '##Look at vw_SMTUnitHistoryTracking First.
    Sub ExportData()
        'Add error handling to fix Error on March 8,2017
        Try

        'Add by Chutchai S on Dec 9,2016
        'To verify /Reconnect connections (cn,cnAutoTest)
        If cn.State = 0 Then
            objFits.reconnect()
        End If
        If cnAutoTest.State = 0 Then
            objAutoTest.reconnect()
        End If

        If cn.State = 0 Or cnAutoTest.State = 0 Then
            Log(Now() & "-- Database is closed (FIT=" & cn.State & " and ATS=" & cnAutoTest.State)
            GoTo NoSN
        End If

        'End Verify Database connection



        Dim vBullEyesObj As New clsBullEyes
        With vBullEyesObj
            'Query Data
            Dim rs As New ADODB.Recordset
            Dim vNewFromDate As Date = CDate(lblFrom.Text).AddSeconds(1)
            Dim vDateFrom As String = vNewFromDate.ToString
            Dim vDateTo As String = lblTo.Text


            rs = objAutoTest.getUUTResult(vLastID)

            If rs Is Nothing Then
                lblLastDate.Text = Now() : Application.DoEvents() ' lblNextRun.Text 
                'GoTo NoSN
                Me.Close()
            End If


            If rs.RecordCount = 0 Then
                lblLastDate.Text = Now() : Application.DoEvents() ' lblNextRun.Text 
                GoTo NoSN
            End If

            'Initial Object
            Dim objFITSDLL As New FITSDLL.clsDB
                Dim vInitResult As String

                'Comment by Chutchai on Feb 16,2017 -- To support EBT station (PCB level)


                'With objFITSDLL
                '    'vInitResult = .fn_InitDB("*", "", "2.9", "dbAcacia")
                '    vInitResult = .fn_InitDB("*", "", "2.9", "dbSMT_BU")
                'End With
                'If Not vInitResult Then
                '    Log(Now() & "--Unable to initial FITSDLL")
                '    Exit Sub
                'End If
                '--------------
                Dim vTempTimeOut As String
                Do While Not rs.EOF

                    tssStatus.Text = "Importing......." & rs.AbsolutePosition & "/" & rs.RecordCount : Application.DoEvents()
                    vTempTimeOut = rs.Fields("start_date_time").Value
                    .datetimeout = vTempTimeOut
                    '1)=========HandShake============
                    'Get Model by Serial number
                    Dim vModel As String
                    Dim vModelType As String
                    Dim vKittingStation As String = "100"
                    Dim vExeStation As String = ""
                    Dim vHandShake As String
                    Dim vSnParamName As String = "Serial No."
                    Dim vSn As String = rs.Fields("uut_serial_number").Value
                    Dim vHWPartFIT As String = ""
                    Dim vUutID As String = rs.Fields("id").Value
                    Dim vProcess As String = rs.Fields("process").Value
                    Dim vDeviceTypeFit As String = ""
                    Dim vDeviceTypeATS As String = ""
                    lblCurrentID.Text = vUutID
                    vLastID = vUutID

                    Dim vDeviceTypeCheck As Boolean = False

                    If vProcess = "EBT" Then
                        With objFITSDLL
                            vInitResult = .fn_InitDB("*", "", "2.9", "dbSMT_BU")
                        End With
                        vExeStation = "3261"
                        vModel = objFITSDLL.fn_Query("", vExeStation, "2.9", vSn, "Part_Number")
                    Else
                        With objFITSDLL
                            vInitResult = .fn_InitDB("*", "", "2.9", "dbAcacia")
                        End With
                        vModelType = objFits.getEventMaster(vSn, "model_type")
                        vModel = objFits.getEventMaster(vSn, "model")

                        vHWPartFIT = objFits.getParameters(vSn, "10112")
                        vModel = vModelType

                        vDeviceTypeFit = objFits.getParameters(vSn, "1204")
                        vDeviceTypeATS = objAutoTest.getDeviceType(vSn, "DCP")

                        vDeviceTypeCheck = IIf(vDeviceTypeATS = vDeviceTypeFit, True, False)
                    End If


                    'vModel = objFITS.fn_Query(txtModel.Text, vKittingStation, "1", rs.Fields(""), "Model")
                    Select Case vModel.ToUpper
                        Case "ACADIA"
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

                        Case "AC400 [NON ETOF]"
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
                    If vHandShake <> "True" And Not vHandShake.Contains("in-processing in " & vExeStation) Then
                        WrongRoutingLog(Now() & "--" & vSn & "--" & vModel & "--" & vProcess & "--" & vUutID & "--" & vHandShake)
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
                        TerminatedLog(Now() & "--" & vSn & "--" & vModel & "--" & vProcess & "--" & vUutID)
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

                    'Add by Chutchai on Dec 13,2016 
                    'To put EN number into FIT by mapping from setting.
                    Dim vEn As String
                    If vLoginName <> "" Then
                        vEn = objInI.GetString("operator", vLoginName, "000000")
                    Else
                        vEn = "000000"
                    End If
                    Dim vCustomer As String = ""
                    If vProcess = "CFG" Then
                        vCustomer = objAutoTest.getCustomerString(vSn, vProcess, vUutID)
                    End If
                    '------------------------------------------------------------
                    Dim vTest1 As String
                    Dim vTest2 As String
                    If vProcess = "EBT" Then
                        vTest1 = "PCBA S/N|Login Name|Product Code|Fixture ID|" & _
                            "Station ID|Date/Time|Execute Time|Mode|TEST_COUNT|" & _
                            "TEST_SOCKET_INDEX|TPS_REV|HW_REV|EBT Result|FAIL MODE"

                        vTest2 = vSn & "|" & vLoginName & "|" & vProductCode & "|" & _
                                        vFixtureID & "|" & vStationID & "|" & vDateTime & "|" & vExeTime & "|" & vMode & "|" & vTestCount & "|" & _
                                        vTestSocketIndex & "|" & vTpsRev & "|" & vHWRev & "|" & IIf(vResult = "Passed", "PASS", "FAIL") & "|" & _
                                        vDisposCode

                    Else
                        vTest1 = "FBN Serial No|Login Name|Product Code|" & _
                                           "Fixture ID|Station ID|Date/Time|Execute Time|Mode|TEST_COUNT|" & _
                                           "TEST_SOCKET_INDEX|TPS_REV|HW_REV|FW_REV|Result|Remark|TOP BOM REV.|" & _
                                           "HW_PART_NUMBER  (FITS)|HW_PART_NUMBER|Device Type (FITS)|Device Type (ATS)|EN|Customer"

                        vTest2 = vSn & "|" & vLoginName & "|" & vProductCode & "|" & _
                                        vFixtureID & "|" & vStationID & "|" & vDateTime & "|" & vExeTime & "|" & vMode & "|" & vTestCount & "|" & _
                                        vTestSocketIndex & "|" & vTpsRev & "|" & vHWRev & "|" & vFWRev & "|" & IIf(vResult = "Passed", "PASS", vDisposCode) & "|" & _
                                        Mid(vRemark, 1, 200) & "|" & vTopBomRev & "|" & _
                                        vHWPartFIT & "|" & vHWPart & "|" & vDeviceTypeFit & "|" & vDeviceTypeATS & "|" & vEn & "|" & vCustomer
                    End If

                    Dim vCheckIn As String
                    Dim vCheckOut As String
                    vLastID = vUutID
                    Select Case vHandShake
                        Case "True"
                            'Need both check in and Out
                            vCheckIn = objFITSDLL.fn_Log(vModel, vExeStation, "0", "FBN Serial No", vSn)
                            vCheckOut = objFITSDLL.fn_Log(vModel, vExeStation, "1", vTest1, vTest2, "|")

                            'vCheckIn = objFITSDLL.fn_Log(vModel, vExeStation, "5", "FBN Serial No", vSn) 'Checkout Delete
                            'vCheckIn = objFITSDLL.fn_Log(vModel, vExeStation, "6", "FBN Serial No", vSn) 'Checkin Delete 


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

            If Now.Minute < 5 Then
                TerminatedLog(Now() & " -- Terminated Program...")
                Close()
                Exit Sub
            End If

        Catch ex As Exception
            TerminatedLog(Now() & " -- Error exception : " & ex.Message)
            Close()
            Exit Sub
        End Try
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
        Dim vNewMessage As String
        vNewMessage = vMessage.Replace(vbCr, "").Replace(vbLf, "")
        file = My.Computer.FileSystem.OpenTextFileWriter(vOutPutFolder & Now().ToString("yyyy-MM-dd") & ".txt", True) '"2016-11-30.txt"
        file.WriteLine(vNewMessage)
        file.Close()
    End Sub

    Sub TerminatedLog(vMessage As String)
        Dim file As System.IO.StreamWriter
        Dim vNewMessage As String
        vNewMessage = vMessage.Replace(vbCr, "").Replace(vbLf, "")
        file = My.Computer.FileSystem.OpenTextFileWriter(vOutPutFolder & Now().ToString("yyyy-MM-dd") & "_Terminated.txt", True) '"2016-11-30.txt"
        file.WriteLine(vNewMessage)
        file.Close()
    End Sub

    Sub WrongRoutingLog(vMessage As String)
        Dim file As System.IO.StreamWriter
        Dim vNewMessage As String
        vNewMessage = vMessage.Replace(vbCr, "").Replace(vbLf, "")
        file = My.Computer.FileSystem.OpenTextFileWriter(vOutPutFolder & Now().ToString("yyyy-MM-dd") & "_WrongRoute.txt", True) '"2016-11-30.txt"
        file.WriteLine(vNewMessage)
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
            objAutoTest.disconnect()
            Me.Close()
        End If

    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
       

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            cn.Close()
            cnAutoTest.Close()
        Catch ex As Exception

        End Try
        
    End Sub



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Dim ss As String = Now().ToString("yyyy-MM-dd")

        'Me.IsHandleCreated


        objInI = New clsINI(Application.StartupPath & "\import.ini")

        'Initial all objects
        objFits = New clsFits
        With objFits
            .user = objInI.GetString("database", "user", "")
            .password = objInI.GetString("database", "password", "")
            .server = objInI.GetString("database", "server", "")
            .database = objInI.GetString("database", "database", "")
            'cn = .connect()
            tssDatabase.Text = "(" & .server & "/" & .database & ")Database connected." : Application.DoEvents()
        End With

        'Open AutoTest database
        objAutoTest = New clsAutoTest
        With objAutoTest
            .user = objInI.GetString("test database", "user", "")
            .password = objInI.GetString("test database", "password", "")
            .server = objInI.GetString("test database", "server", "")
            .database = objInI.GetString("test database", "database", "")
            'cnAutoTest = .connect()
        End With
        '--------------------


        Timer1.Interval = (Val(objInI.GetString("import", "interval", ""))) * 1000 * 60
        Timer1.Enabled = False

        vWorkingDir = objInI.GetString("path", "working dir", "")
        If vWorkingDir = "" Then
            vWorkingDir = Application.StartupPath & "\"
        End If



        initialControl()
        lblCurrentID.Text = vLastID
        Me.Text = Me.Text + " version : " + Application.ProductVersion.Trim()

        'Auto Mode Check
        vAutoRun = IIf(objInI.GetString("running mode", "auto", "") = "True", True, False)
        If vAutoRun Then
            btnAuto.BackColor = Color.Green

            Button1_Click(sender, e) : Application.DoEvents()
            btnImport_Click(sender, e) : Application.DoEvents()
        Else
            btnAuto.BackColor = Color.Red
        End If
       
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        'First import then using Timer.
        If btnImport.Text = "&Start Import data" Then

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

NoConnection:


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

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles btnAuto.Click
        'Me.Close()
    End Sub
End Class
