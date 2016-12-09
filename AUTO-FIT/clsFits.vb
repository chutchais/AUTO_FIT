Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports ADODB


Public Class clsFits
    Private vServer As String
    Private vPassword As String
    Private vUser As String
    Private vDatabase As String

    Dim cn As New ADODB.Connection()

    Public Property server As String
        Get
            server = vServer
        End Get
        Set(value As String)
            vServer = value
        End Set
    End Property

    Public Property user As String
        Get
            user = vUser
        End Get
        Set(value As String)
            vUser = value
        End Set
    End Property

    Public Property password As String
        Get
            password = vPassword
        End Get
        Set(value As String)
            vPassword = value
        End Set
    End Property

    Public Property database As String
        Get
            database = vDatabase
        End Get
        Set(value As String)
            vDatabase = value
        End Set
    End Property

    Public Function connect() As ADODB.Connection

        Dim connectionString As String = "Provider=SQLOLEDB;Initial Catalog=" & vDatabase & ";" & _
            "Data Source=" & vServer & ";User Id=" & vUser & ";" & _
            "Password=" & vPassword & ";Database=" & vDatabase & ";"

        cn.ConnectionString = connectionString
        cn.ConnectionTimeout = 300
        cn.Open()
        cn.CursorLocation = CursorLocationEnum.adUseClient
        connect = cn

    End Function

    Public Sub disconnect()

    End Sub

    Public Function getEvents(vDateFrom As String, vDateTo As String) As ADODB.Recordset
        'Date format : 2016-07-01
        'Dim vSql As String = "select event.*,operation_map.description as operation_name " & _
        '                "from event INNER JOIN operation_map ON event.operation = operation_map.operation " & _
        '                "where event.buildtype in ('RMA','PRODUCTION','QUALIFICATION') and " & _
        '                "event.model=operation_map.model_type and event.timestamp between ? and ?  " & _
        '                "order by event.date_time"
        'Edit by Chutchai S on Sep 28,2016
        'To use Datetime_checkout -- Completed process.
        'Dim vSql As String = "select event.*,operation_map.description as operation_name " & _
        '        "from event INNER JOIN operation_map ON event.operation = operation_map.operation " & _
        '        "where event.buildtype in ('RMA','PRODUCTION','QUALIFICATION') and " & _
        '        "event.model=operation_map.model_type and event.date_time_checkout between ? and ?  " & _
        '        "order by event.date_time_checkout"

        Dim vSql As String = "select event.*,operation_map.description as operation_name,event_master.model as model2 " &
                "from event INNER JOIN " &
                "operation_map ON event.operation = operation_map.operation INNER JOIN " &
                "event_master ON event.serial_no = event_master.serial_no " &
                "where event.buildtype in ('RMA','PRODUCTION','QUALIFICATION','NORMAL','SERVICE UPGRADE') and " &
                "event_master.workorder <> 'N/A'  " &
                "and event_master.model_type=operation_map.model_type " &
                "and event.date_time_checkout between ? and ? " &
                "order by event.date_time_checkout"

        Dim cmd As New ADODB.Command()
        Dim sDateFromParam As ADODB.Parameter
        Dim sDateToParam As ADODB.Parameter

        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sDateFromParam = .CreateParameter("vDateFrom", DataTypeEnum.adVarChar,
                                                 ParameterDirectionEnum.adParamInput, 50, vDateFrom)
            sDateToParam = .CreateParameter("vDateTo", DataTypeEnum.adVarChar,
                                                 ParameterDirectionEnum.adParamInput, 50, vDateTo)
            .Parameters.Append(sDateFromParam)
            .Parameters.Append(sDateToParam)
            getEvents = .Execute
        End With

    End Function


    'Public Function getPCBAlist(vDateFrom As String, vDateTo As String) As ADODB.Recordset

    '    Dim vSql As String = "select * " &
    '                        "from vw_SMTUnitHistoryTracking " &
    '                        "where date_time between ? and ? "

    '    Dim cmd As New ADODB.Command()
    '    Dim sDateFromParam As ADODB.Parameter
    '    Dim sDateToParam As ADODB.Parameter

    '    With cmd
    '        .ActiveConnection = cn
    '        .CommandText = vSql
    '        .CommandType = CommandTypeEnum.adCmdText
    '        sDateFromParam = .CreateParameter("vDateFrom", DataTypeEnum.adVarChar,
    '                                             ParameterDirectionEnum.adParamInput, 50, vDateFrom)
    '        sDateToParam = .CreateParameter("vDateTo", DataTypeEnum.adVarChar,
    '                                             ParameterDirectionEnum.adParamInput, 50, vDateTo)
    '        .Parameters.Append(sDateFromParam)
    '        .Parameters.Append(sDateToParam)
    '        getPCBAlist = .Execute
    '    End With

    'End Function
    Public Function getPCBAlist(vSerialNumber As String) As ADODB.Recordset

        Dim vSql As String = "select * " &
                            "from vw_SMTUnitHistoryTracking " &
                            "where serial_no = ?"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter

        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vDateFrom", DataTypeEnum.adVarChar,
                                                 ParameterDirectionEnum.adParamInput, 50, vSerialNumber)
            .Parameters.Append(sSnParam)
            getPCBAlist = .Execute
        End With

    End Function

    Public Function getPCBAlist(vDateFrom As String, vDateTo As String) As ADODB.Recordset

        Dim vSql As String = "select * " &
                            "from vw_SMTUnitHistoryTracking " &
                            "where operation = '3261' and date_time between ? and ?"

        Dim cmd As New ADODB.Command()
        Dim sDateFromParam As ADODB.Parameter
        Dim sDateToParam As ADODB.Parameter

        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sDateFromParam = .CreateParameter("vDateFrom", DataTypeEnum.adVarChar,
                                                 ParameterDirectionEnum.adParamInput, 50, vDateFrom)
            sDateToParam = .CreateParameter("vDateTo", DataTypeEnum.adVarChar,
                                                 ParameterDirectionEnum.adParamInput, 50, vDateTo)
            .Parameters.Append(sDateFromParam)
            .Parameters.Append(sDateToParam)
            .CommandTimeout = 300
            getPCBAlist = .Execute
        End With

    End Function



    Public Function getComponentData(vSerialnumber As String) As ADODB.Recordset

        Dim vSql As String = "select * " & _
                            "from vw_SMTUnitComponentTracking " & _
                            "where serial_no=?"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter
        'sStationParam,
        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
            .Parameters.Append(sSnParam)

            getComponentData = .Execute
        End With

    End Function

    Public Function getEvents(vSerialnumber As String) As ADODB.Recordset
        'Date format : 2016-07-01
        'Dim vSql As String = "select event.*,operation_map.description as operation_name " & _
        '                "from event INNER JOIN operation_map ON event.operation = operation_map.operation " & _
        '                "where serial_no = ?  and event.model=operation_map.model_type " & _
        '                "order by date_time"

        Dim vSql As String = "select event.*,operation_map.description as operation_name,event_master.model as model2 " & _
                "from event INNER JOIN " & _
                "operation_map ON event.operation = operation_map.operation INNER JOIN " & _
                "event_master ON event.serial_no = event_master.serial_no " & _
                "where serial_no = ? and event.buildtype in ('RMA','PRODUCTION','QUALIFICATION','NORMAL','SERVICE UPGRADE') and " & _
                "event_master.workorder <> 'N/A'  " & _
                "order by event.date_time_checkout"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter
        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
            .Parameters.Append(sSnParam)
            getEvents = .Execute
        End With
    End Function


    Public Function getNextStation(vSerialnumber As String,
                                   vStation As String,
                                   vTrans_seq As String, vModel As String) As String
        'Date format : 2016-07-01
        Dim vSql As String = "select * " & _
                            "from route_table " & _
                            "where model_type=? and prev_valid_operation=?"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter
        Dim sStation As ADODB.Parameter
        Dim vRouteRst As ADODB.Recordset
        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vModelType", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vModel)
            sStation = .CreateParameter("vModelType", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vStation)
            .Parameters.Append(sSnParam)
            .Parameters.Append(sStation)
            vRouteRst = .Execute
        End With
        getNextStation = vStation 'Default value
        Dim vCriteria As String
        Dim cmdCriteria As New ADODB.Command()
        With cmdCriteria
            .ActiveConnection = cn
            .CommandType = CommandTypeEnum.adCmdText
        End With
        Try
            With vRouteRst
                Do While Not .EOF
                    'Query on criteria field
                    If IsDBNull(.Fields("criteria").Value) Then
                        getNextStation = .Fields("cur_operation").Value
                        Exit Do
                    End If
                    vCriteria = .Fields("criteria").Value
                    If vCriteria = "where m.buildetype ='RMA' c" Then
                        GoTo NextLoop
                    End If
                    vCriteria = Replace(vCriteria, "m.max_trans_seq", "m.trans_seq")

                    vSql = "select * " & _
                            "from event as m " & _
                            " " & vCriteria & " " & _
                            "and m.serial_no='" & vSerialnumber & "' " & _
                            "and m.operation='" & vStation & "' " & _
                            "and m.trans_seq=" & vTrans_seq
                    '"and m.date_time =(select max(date_time) from event where serial_no='163256019' and operation=100 )"
                    cmdCriteria.CommandText = vSql
                    If cmdCriteria.Execute.RecordCount > 0 Then
                        getNextStation = .Fields("cur_operation").Value
                        Exit Do
                    End If
NextLoop:
                    .MoveNext()
                Loop
            End With
        Catch ex As Exception

        End Try
        

    End Function

    Public Function getEventMaster(vSerialnumber As String) As ADODB.Recordset
        'Date format : 2016-07-01
        Dim vSql As String = "select * " & _
                        "from event_master " & _
                        "where serial_no = ? " & _
                        "order by date_time"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter
        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
            .Parameters.Append(sSnParam)
            getEventMaster = .Execute
        End With
    End Function

    Public Function getParameters(vSerialnumber As String, vAtrrCode As Integer, _
                                  vTransSeq As Integer) As ADODB.Recordset
        'Date format : 2016-07-01
        Dim vSql As String = "SELECT attribute.* ,attribute_map.* " & _
                        "FROM attribute_map INNER JOIN " & _
                         "attribute ON attribute_map.attribute_code = attribute.attribute_code " & _
                        "WHERE (attribute.serial_no =? " & _
                        "and attribute.trans_seq=? " & _
                        "and attribute.sn_attr_code=?)"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter
        Dim sTranSeqParam As ADODB.Parameter
        Dim sAttCodeParam As ADODB.Parameter

        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
            sTranSeqParam = .CreateParameter("vSerialnumber", DataTypeEnum.adInteger, _
                                                 ParameterDirectionEnum.adParamInput, 50, vTransSeq)
            sAttCodeParam = .CreateParameter("vSerialnumber", DataTypeEnum.adInteger, _
                                                 ParameterDirectionEnum.adParamInput, 50, vAtrrCode)
            .Parameters.Append(sSnParam)
            .Parameters.Append(sTranSeqParam)
            .Parameters.Append(sAttCodeParam)
            getParameters = .Execute
        End With
    End Function

    Public Function getParameters(vSerialnumber As String, vAtrrCode As Integer) As String
        Try
            'Date format : 2016-07-01
            Dim vSql As String = "SELECT attribute.* ,attribute_map.* " & _
                            "FROM attribute_map INNER JOIN " & _
                             "attribute ON attribute_map.attribute_code = attribute.attribute_code " & _
                            "WHERE (attribute.serial_no =? " & _
                            "and attribute.attribute_code=?)"

            Dim cmd As New ADODB.Command()
            Dim sSnParam As ADODB.Parameter
            Dim sAttCodeParam As ADODB.Parameter
            Dim vRstTmp As New ADODB.Recordset

            With cmd
                .ActiveConnection = cn
                .CommandText = vSql
                .CommandType = CommandTypeEnum.adCmdText
                sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                     ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
                sAttCodeParam = .CreateParameter("vSerialnumber", DataTypeEnum.adInteger, _
                                                     ParameterDirectionEnum.adParamInput, 50, vAtrrCode)
                .Parameters.Append(sSnParam)
                .Parameters.Append(sAttCodeParam)
                vRstTmp = .Execute
                If vRstTmp.RecordCount > 0 Then
                    Return vRstTmp.Fields("attribute_value").Value
                Else
                    Return ""
                End If
            End With
        Catch ex As Exception
            Return ""
        End Try
        
    End Function














End Class


Public Class clsAutoTest
    Private vServer As String
    Private vPassword As String
    Private vUser As String
    Private vDatabase As String

    Dim cn As New ADODB.Connection()

    Public Property server As String
        Get
            server = vServer
        End Get
        Set(value As String)
            vServer = value
        End Set
    End Property

    Public Property user As String
        Get
            user = vUser
        End Get
        Set(value As String)
            vUser = value
        End Set
    End Property

    Public Property password As String
        Get
            password = vPassword
        End Get
        Set(value As String)
            vPassword = value
        End Set
    End Property

    Public Property database As String
        Get
            database = vDatabase
        End Get
        Set(value As String)
            vDatabase = value
        End Set
    End Property

    Public Function connect() As ADODB.Connection

        Dim connectionString As String = "Provider=SQLOLEDB;Initial Catalog=" & vDatabase & ";" & _
            "Data Source=" & vServer & ";User Id=" & vUser & ";" & _
            "Password=" & vPassword & ";Database=" & vDatabase & ";"

        cn.ConnectionString = connectionString
        cn.Open()
        cn.CursorLocation = CursorLocationEnum.adUseClient
        connect = cn

    End Function

    Public Sub disconnect()

    End Sub


    Public Function getUUTResult(vDateFrom As String, vDateTo As String) As ADODB.Recordset
        'Date format : 2016-07-01
        Dim vSql As String = "select * " & _
                            "from uut_result " & _
                            "where process='DCP' " & _
                            "and start_date_time between ? and ?  " & _
                            "order by start_date_time"

        Dim cmd As New ADODB.Command()
        Dim sDateFromParam As ADODB.Parameter
        Dim sDateToParam As ADODB.Parameter

        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sDateFromParam = .CreateParameter("vDateFrom", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vDateFrom)
            sDateToParam = .CreateParameter("vDateTo", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vDateTo)
            .Parameters.Append(sDateFromParam)
            .Parameters.Append(sDateToParam)
            getUUTResult = .Execute
        End With

    End Function


    Public Function getUUTResult(vUUTID As String) As ADODB.Recordset
        'Date format : 2016-07-01
        Dim vSql As String = "select * " & _
                            "from uut_result " & _
                            "where process in ('DCP','DBI','FTU','FPT','FAT','OBS','FVT','EPT','ESS','FST','EXS','CFG') " & _
                            "and id > ?  and uut_status <>'Error'" & _
                            "order by id"
        'order by start_date_time
        'PIC level : 'DCP','DBI','FTU','FPT','FAT'
        'Module level : 'OBS','FVT','EPT','ESS','FST','EXS','CFG'
        Dim cmd As New ADODB.Command()
        Dim sDateFromParam As ADODB.Parameter

        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sDateFromParam = .CreateParameter("vDateFrom", DataTypeEnum.adBigInt, _
                                                 ParameterDirectionEnum.adParamInput, 50, vUUTID)
            .Parameters.Append(sDateFromParam)
            getUUTResult = .Execute
        End With

    End Function

    Public Function getEvents(vDateFrom As String, vDateTo As String) As ADODB.Recordset
        'Date format : 2016-07-01
        Dim vSql As String = "select event.*,operation_map.description as operation_name " & _
                        "from event INNER JOIN operation_map ON event.operation = operation_map.operation " & _
                        "where event.buildtype in ('RMA','PRODUCTION','QUALIFICATION') and " & _
                        "event.model=operation_map.model_type and event.timestamp between ? and ?  " & _
                        "order by event.date_time"

        Dim cmd As New ADODB.Command()
        Dim sDateFromParam As ADODB.Parameter
        Dim sDateToParam As ADODB.Parameter

        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sDateFromParam = .CreateParameter("vDateFrom", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vDateFrom)
            sDateToParam = .CreateParameter("vDateTo", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vDateTo)
            .Parameters.Append(sDateFromParam)
            .Parameters.Append(sDateToParam)
            getEvents = .Execute
        End With

    End Function

    Public Function getEvents(vSerialnumber As String) As ADODB.Recordset
        'Date format : 2016-07-01
        Dim vSql As String = "select event.*,operation_map.description as operation_name " & _
                        "from event INNER JOIN operation_map ON event.operation = operation_map.operation " & _
                        "where serial_no = ?  and event.model=operation_map.model_type " & _
                        "order by date_time"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter
        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
            .Parameters.Append(sSnParam)
            getEvents = .Execute
        End With
    End Function


    Public Function getNextStation(vSerialnumber As String,
                                   vStation As String,
                                   vTrans_seq As String, vModel As String) As String
        'Date format : 2016-07-01
        Dim vSql As String = "select * " & _
                            "from route_table " & _
                            "where model_type=? and prev_valid_operation=?"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter
        Dim sStation As ADODB.Parameter
        Dim vRouteRst As ADODB.Recordset
        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vModelType", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vModel)
            sStation = .CreateParameter("vModelType", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vStation)
            .Parameters.Append(sSnParam)
            .Parameters.Append(sStation)
            vRouteRst = .Execute
        End With
        getNextStation = vStation 'Default value
        Dim vCriteria As String
        Dim cmdCriteria As New ADODB.Command()
        With cmdCriteria
            .ActiveConnection = cn
            .CommandType = CommandTypeEnum.adCmdText
        End With
        Try
            With vRouteRst
                Do While Not .EOF
                    'Query on criteria field
                    If IsDBNull(.Fields("criteria").Value) Then
                        getNextStation = .Fields("cur_operation").Value
                        Exit Do
                    End If
                    vCriteria = .Fields("criteria").Value
                    If vCriteria = "where m.buildetype ='RMA' c" Then
                        GoTo NextLoop
                    End If
                    vCriteria = Replace(vCriteria, "m.max_trans_seq", "m.trans_seq")

                    vSql = "select * " & _
                            "from event as m " & _
                            " " & vCriteria & " " & _
                            "and m.serial_no='" & vSerialnumber & "' " & _
                            "and m.operation='" & vStation & "' " & _
                            "and m.trans_seq=" & vTrans_seq
                    '"and m.date_time =(select max(date_time) from event where serial_no='163256019' and operation=100 )"
                    cmdCriteria.CommandText = vSql
                    If cmdCriteria.Execute.RecordCount > 0 Then
                        getNextStation = .Fields("cur_operation").Value
                        Exit Do
                    End If
NextLoop:
                    .MoveNext()
                Loop
            End With
        Catch ex As Exception

        End Try


    End Function


    Public Function getTestData(vSerialnumber As String, vProcess As String, _
                                ByRef vTester As String, vDatetime As String) As ADODB.Recordset
        'Date format : 2016-07-01
        'Dim vSql As String = "select id,station_id,start_date_time,uut_status,process,tps_name " & _
        '                    "from uut_result " & _
        '                    "where uut_serial_number=?  " & _
        '                    "and station_id=? " & _
        '                    "and process=? and start_date_time <= ? " & _
        '                    "order by id"
        Dim vSql As String = "select id,station_id,start_date_time,uut_status,process,tps_name " & _
                    "from uut_result " & _
                    "where uut_serial_number=?  " & _
                    "and process=? and start_date_time <= ? " & _
                    "order by id"

        Dim cmd As New ADODB.Command()
        Dim sSnParam, sProcessParam, sDdatetimeParam As ADODB.Parameter
        'sStationParam,
        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
            'sStationParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
            '                                     ParameterDirectionEnum.adParamInput, 50, vTester)
            sProcessParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vProcess)
            sDdatetimeParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vDatetime)
            .Parameters.Append(sSnParam)
            '.Parameters.Append(sStationParam)
            .Parameters.Append(sProcessParam)
            .Parameters.Append(sDdatetimeParam)
            getTestData = .Execute
        End With

        If getTestData.RecordCount > 0 Then
            getTestData.MoveLast() 'Get last record.
            vTester = getTestData.Fields("station_id").Value
            Dim vUutResult As String
            vUutResult = getTestData.Fields("id").Value
            vSql = "SELECT STEP_RESULT.UUT_RESULT, STEP_RESULT.STEP_NAME, " & _
                "STEP_RESULT.STEP_GROUP, STEP_RESULT.STATUS, PROP_RESULT.DATA, " & _
                "PROP_NUMERICLIMIT.HIGH_LIMIT, PROP_NUMERICLIMIT.LOW_LIMIT, " & _
                "PROP_NUMERICLIMIT.UNITS, PROP_NUMERICLIMIT.COMP_OPERATOR,STEP_RESULT.ID,UUT_RESULT.STATION_ID " & _
                "FROM STEP_RESULT INNER JOIN " & _
                         "PROP_RESULT ON STEP_RESULT.ID = PROP_RESULT.STEP_RESULT INNER JOIN " & _
                         "PROP_NUMERICLIMIT ON PROP_RESULT.ID = PROP_NUMERICLIMIT.PROP_RESULT INNER JOIN " & _
                         "UUT_RESULT ON STEP_RESULT.UUT_RESULT = UUT_RESULT.ID " & _
                "WHERE(STEP_RESULT.UUT_RESULT = ? and PROP_RESULT.DATA<>'0' )"
            Dim cmd2 As New ADODB.Command()
            Dim sUUTparam As ADODB.Parameter
            With cmd2
                .ActiveConnection = cn
                .CommandText = vSql
                .CommandType = CommandTypeEnum.adCmdText
                sUUTparam = .CreateParameter("vUUTserial", DataTypeEnum.adVarChar, _
                                                     ParameterDirectionEnum.adParamInput, 50, vUutResult)
                .Parameters.Append(sUUTparam)
                getTestData = .Execute
            End With

        End If


    End Function


    Public Function getDeviceType(vSerialnumber As String, vProcess As String) As String
        Try
            Dim vSql As String = "SELECT STEP_RESULT.STEP_NAME, PROP_RESULT.DATA, PROP_RESULT.DISPLAY_FORMAT, UUT_RESULT.UUT_SERIAL_NUMBER " & _
                    "FROM     STEP_RESULT INNER JOIN " & _
                                      "PROP_RESULT ON STEP_RESULT.ID = PROP_RESULT.STEP_RESULT INNER JOIN " & _
                                      "UUT_RESULT ON STEP_RESULT.UUT_RESULT = UUT_RESULT.ID " & _
                    "WHERE  (STEP_RESULT.STEP_NAME LIKE 'Device Type%') AND (UUT_RESULT.UUT_SERIAL_NUMBER = ?) "

            Dim cmd As New ADODB.Command()
            Dim vRst As New ADODB.Recordset
            Dim sSN As ADODB.Parameter
            'sStationParam,
            With cmd
                .ActiveConnection = cn
                .CommandText = vSql
                .CommandType = CommandTypeEnum.adCmdText
                sSN = .CreateParameter("vSerialnumber", DataTypeEnum.adChar, _
                                                     ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
                .Parameters.Append(sSN)

                vRst = .Execute
            End With
            If vRst.RecordCount > 0 Then
                getDeviceType = vRst.Fields("DATA").Value
            Else
                getDeviceType = ""
            End If
        Catch ex As Exception
            getDeviceType = ""
        End Try





    End Function


    Public Function getTestData(vSerialnumber As String, vProcess As String, _
                                ByRef vID As String) As ADODB.Recordset
        Dim vSql As String
            Dim vUutResult As String
        vUutResult = vID

        vSql = "select * " & _
                "from step_result " & _
                "where UUT_RESULT = ? " & _
                "and status='Failed'"
        'Midify by Chutchai on Dec 8,2016
        'To remove and PROP_RESULT.DATA<>'0' 
        vSql = "SELECT STEP_RESULT.UUT_RESULT, STEP_RESULT.STEP_NAME, " & _
            "STEP_RESULT.STEP_GROUP, STEP_RESULT.STATUS, PROP_RESULT.DATA, " & _
            "PROP_NUMERICLIMIT.HIGH_LIMIT, PROP_NUMERICLIMIT.LOW_LIMIT, " & _
            "PROP_NUMERICLIMIT.UNITS, PROP_NUMERICLIMIT.COMP_OPERATOR,STEP_RESULT.ID,UUT_RESULT.STATION_ID " & _
            "FROM STEP_RESULT INNER JOIN " & _
                     "PROP_RESULT ON STEP_RESULT.ID = PROP_RESULT.STEP_RESULT INNER JOIN " & _
                     "PROP_NUMERICLIMIT ON PROP_RESULT.ID = PROP_NUMERICLIMIT.PROP_RESULT INNER JOIN " & _
                     "UUT_RESULT ON STEP_RESULT.UUT_RESULT = UUT_RESULT.ID " & _
            "WHERE(STEP_RESULT.UUT_RESULT = ? )"
        'and PROP_RESULT.DATA<>'0' 
        Dim cmd2 As New ADODB.Command()
        Dim sUUTparam As ADODB.Parameter
        With cmd2
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sUUTparam = .CreateParameter("vUUTserial", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vUutResult)
            .Parameters.Append(sUUTparam)
            getTestData = .Execute
        End With
    End Function


    Public Function getTestDataString(vSerialnumber As String, vProcess As String, _
                                ByRef vID As String) As String
        Dim vSql As String
        Dim vUutResult As String
        vUutResult = vID

        vSql = "select * " & _
                "from step_result " & _
                "where UUT_RESULT = ? " & _
                "and status='Failed'"
        'and PROP_RESULT.DATA<>'0' 
        Dim vRst As New ADODB.Recordset
        Dim cmdUUTResult As New ADODB.Command()
        Dim sUUTparam As ADODB.Parameter
        With cmdUUTResult
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sUUTparam = .CreateParameter("vUUTResult", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vUutResult)
            .Parameters.Append(sUUTparam)
            vRst = .Execute
        End With

        If vRst.RecordCount = 0 Then
            vUutResult = "Terminated = Not found failed test record"
        End If

        vRst.MoveLast()
        Dim vSympTom As String
        Dim vFailDetail As String

        If vRst.Fields("step_type").Value = "PassFailTest" Then
            vSympTom = Mid(vRst.Fields("step_name").Value, 1, 50)
            vSympTom = vSympTom.Replace("|", " ")
            vFailDetail = IIf(IsDBNull(vRst.Fields("report_text").Value), "No error message text", vRst.Fields("report_text").Value)
            vFailDetail = vFailDetail.Replace("|", " ")
            vUutResult = vSympTom & "=" & Mid(vFailDetail, 1, 190)
            getTestDataString = vUutResult
        ElseIf vRst.Fields("step_type").Value = "NumericLimitTest" Then
            'Midify by Chutchai on Dec 8,2016
            'To remove and PROP_RESULT.DATA<>'0' 
            vSql = "SELECT STEP_RESULT.UUT_RESULT, STEP_RESULT.STEP_NAME, " & _
                "STEP_RESULT.STEP_GROUP, STEP_RESULT.STATUS, PROP_RESULT.DATA, " & _
                "PROP_NUMERICLIMIT.HIGH_LIMIT, PROP_NUMERICLIMIT.LOW_LIMIT, " & _
                "PROP_NUMERICLIMIT.UNITS, PROP_NUMERICLIMIT.COMP_OPERATOR,STEP_RESULT.ID,UUT_RESULT.STATION_ID " & _
                "FROM STEP_RESULT INNER JOIN " & _
                         "PROP_RESULT ON STEP_RESULT.ID = PROP_RESULT.STEP_RESULT INNER JOIN " & _
                         "PROP_NUMERICLIMIT ON PROP_RESULT.ID = PROP_NUMERICLIMIT.PROP_RESULT INNER JOIN " & _
                         "UUT_RESULT ON STEP_RESULT.UUT_RESULT = UUT_RESULT.ID " & _
                "WHERE(STEP_RESULT.UUT_RESULT = ? )"
            Dim vRst2 As ADODB.Recordset
            Dim cmdUUTResult2 As New ADODB.Command()
            Dim sUUTparam2 As ADODB.Parameter
            With cmdUUTResult2
                .ActiveConnection = cn
                .CommandText = vSql
                .CommandType = CommandTypeEnum.adCmdText
                sUUTparam2 = .CreateParameter("vUUTResult", DataTypeEnum.adVarChar, _
                                                     ParameterDirectionEnum.adParamInput, 50, vUutResult)
                .Parameters.Append(sUUTparam2)
                vRst2 = .Execute
            End With
            vRst2.Filter = "status <> 'Passed'"
            'Dim vStrResult As String

            If vRst2.RecordCount > 0 Then
                vUutResult = vRst2.Fields("step_name").Value & "=" & _
                    IIf(IsDBNull(vRst2.Fields("data").Value), "", vRst2.Fields("data").Value) & " " & _
                    IIf(IsDBNull(vRst2.Fields("units").Value), "", vRst2.Fields("units").Value) & _
                    "(" & _
                    IIf(IsDBNull(vRst2.Fields("low_limit").Value), "", vRst2.Fields("low_limit").Value) & _
                    "/" & _
                    IIf(IsDBNull(vRst2.Fields("high_limit").Value), "", vRst2.Fields("high_limit").Value) & _
                    ")"
                getTestDataString = vUutResult
            Else
                vUutResult = vUutResult = vRst.Fields("step_name").Value & "=" & IIf(IsDBNull(vRst.Fields("report_text").Value), "No error message text", vRst.Fields("report_text").Value)
                getTestDataString = vUutResult
            End If
        Else
            vUutResult = vUutResult = vRst.Fields("step_name").Value & "=" & IIf(IsDBNull(vRst.Fields("report_text").Value), "No error message text", vRst.Fields("report_text").Value)
            getTestDataString = vUutResult
        End If
        

    End Function


    Public Function getStepTestData(vSerialnumber As String, vProcess As String, _
                            ByRef vID As String) As ADODB.Recordset
        Dim vSql As String
        Dim vUutResult As String
        vUutResult = vID
        'Midify by Chutchai on Dec 8,2016

        vSql = "select * " & _
                "from step_result " & _
                "where UUT_RESULT=?"
        'and PROP_RESULT.DATA<>'0' 
        Dim cmd2 As New ADODB.Command()
        Dim sUUTparam As ADODB.Parameter
        With cmd2
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sUUTparam = .CreateParameter("vUUTserial", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vUutResult)
            .Parameters.Append(sUUTparam)
            getStepTestData = .Execute
        End With
    End Function


    Public Function getComponentData(vSerialnumber As String) As ADODB.Recordset

        Dim vSql As String = "select * " & _
                            "from vw_SMTUnitComponentTracking " & _
                            "where serial_no=?"

        Dim cmd As New ADODB.Command()
        Dim sSnParam As ADODB.Parameter
        'sStationParam,
        With cmd
            .ActiveConnection = cn
            .CommandText = vSql
            .CommandType = CommandTypeEnum.adCmdText
            sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
                                                 ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
            .Parameters.Append(sSnParam)

            getComponentData = .Execute
        End With

    End Function

    'Public Function getParameters(vSerialnumber As String, vAtrrCode As Integer, _
    '                              vTransSeq As Integer) As ADODB.Recordset
    '    'Date format : 2016-07-01
    '    Dim vSql As String = "SELECT attribute.* ,attribute_map.* " & _
    '                    "FROM attribute_map INNER JOIN " & _
    '                     "attribute ON attribute_map.attribute_code = attribute.attribute_code " & _
    '                    "WHERE (attribute.serial_no =? " & _
    '                    "and attribute.trans_seq=? " & _
    '                    "and attribute.sn_attr_code=?)"

    '    Dim cmd As New ADODB.Command()
    '    Dim sSnParam As ADODB.Parameter
    '    Dim sTranSeqParam As ADODB.Parameter
    '    Dim sAttCodeParam As ADODB.Parameter

    '    With cmd
    '        .ActiveConnection = cn
    '        .CommandText = vSql
    '        .CommandType = CommandTypeEnum.adCmdText
    '        sSnParam = .CreateParameter("vSerialnumber", DataTypeEnum.adVarChar, _
    '                                             ParameterDirectionEnum.adParamInput, 50, vSerialnumber)
    '        sTranSeqParam = .CreateParameter("vSerialnumber", DataTypeEnum.adInteger, _
    '                                             ParameterDirectionEnum.adParamInput, 50, vTransSeq)
    '        sAttCodeParam = .CreateParameter("vSerialnumber", DataTypeEnum.adInteger, _
    '                                             ParameterDirectionEnum.adParamInput, 50, vAtrrCode)
    '        .Parameters.Append(sSnParam)
    '        .Parameters.Append(sTranSeqParam)
    '        .Parameters.Append(sAttCodeParam)
    '        getParameters = .Execute
    '    End With
    'End Function














End Class
