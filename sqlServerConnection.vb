Function connect(ByVal dbname As String, ByVal lastRow As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim CKCounter As Integer
    Dim BuildString As String
    Dim BuildCKString As String
    Dim Cn As ADODB.Connection
    Dim Server_Name As String
    Dim Database_Name As String
    Dim User_ID As String
    Dim Password As String
    Dim SQLStr As String
    Dim rs As ADODB.Recordset
    Dim str As String
    Dim begin As Integer
    rs = New ADODB.Recordset

    begin = lastRow + 1
    i = 2
    Server_Name = "10.23.16.209" ' Enter your server name here
    Database_Name = dbname ' Enter your database name here
    User_ID = "DBREAD" ' Enter your user ID here
    Password = "db@2017" ' Enter your password here

    If Sheet1.Range("A2").Value = "" Then
        MsgBox("No data input!")
        End
    End If
    Do While Worksheets("Sheet1").Cells(i, 1).Value <> ""
        If i > 30 Then
            MsgBox("Maximum number of partno is 30!")
            End
        End If
        str = Worksheets("Sheet1").Cells(i, 1)
        If i <> 2 Then
            BuildString = BuildString & ",'" & str & "'"
            i = i + 1
        Else
            BuildString = BuildString & "'" & str & "'"
            i = i + 1
        End If
    Loop

    begin = Worksheets("Event scan history").Cells(Rows.Count, 1).End(xlUp).Row + 1
    SQLStr = "select  sysserialno,eventname,scandatetime,productionline,scanby from mfsysevent (nolock) where sysserialno in ( " & BuildString & " ) order by sysserialno, scandatetime "

    Cn = New ADODB.Connection
    Cn.Open("Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
    ";Uid=" & User_ID & ";Pwd=" & Password & ";")

    rs.Open(SQLStr, Cn, adOpenStatic)
    ' Dump to spreadsheet
    With Worksheets("Event scan history").Range("a" & begin & ":az5000") ' Enter your sheet name and range here
        .ClearContents()
        .CopyFromRecordset(rs)
        lastRow = Worksheets("Event scan history").Cells(Rows.Count, 1).End(xlUp).Row
        connect = lastRow
    End With
    '            Tidy up
    rs.Close()
    rs = Nothing
    Cn.Close()
    Cn = Nothing

    begin = Worksheets("System CT# history").Cells(Rows.Count, 1).End(xlUp).Row + 1
    rs = New ADODB.Recordset
    SQLStr = "select sysserialno,cserialno,eeecode,partno,categoryname,formtype,eventpoint,lasteditby,lasteditdt from mfsyscserial (nolock) where sysserialno in ( " & BuildString & " )"

    Cn = New ADODB.Connection
    Cn.Open("Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
    ";Uid=" & User_ID & ";Pwd=" & Password & ";")

    rs.Open(SQLStr, Cn, adOpenStatic)
    ' Dump to spreadsheet
    With Worksheets("System CT# history").Range("a" & begin & ":az5000") ' Enter your sheet name and range here
        .ClearContents()
        .CopyFromRecordset(rs)
        lastRow = Worksheets("System CT# history").Cells(Rows.Count, 1).End(xlUp).Row
        connect = lastRow
    End With
    '            Tidy up
    rs.Close()
    rs = Nothing
    Cn.Close()
    Cn = Nothing

    If Worksheets("System CT# history").Cells(2, 2).Value <> "" Then

        j = 2
        CKCounter = 1
        Do While Worksheets("System CT# history").Cells(j, 2).Value <> ""
            If Worksheets("System CT# history").Cells(j, 2).Value Like "CK*" Then
                If CKCounter <> 1 Then
                    BuildCKString = BuildCKString & ",'" & Worksheets("System CT# history").Cells(j, 2).Value & "'"
                    CKCounter = CKCounter + 1
                Else
                    BuildCKString = BuildCKString & "'" & Worksheets("System CT# history").Cells(j, 2).Value & "'"
                    CKCounter = CKCounter + 1
                End If
            End If
            j = j + 1
        Loop


        begin = Worksheets("CK Part History installed").Cells(Rows.Count, 1).End(xlUp).Row + 1
        rs = New ADODB.Recordset
        SQLStr = "SELECT ctserialno, partno, qty, keypart, installed, categoryname, scantype, lasteditby, lasteditdt FROM mfctcomponent (NOLOCK) WHERE ctserialno in ( " & BuildCKString & " ) ORDER BY lasteditdt"

        Cn = New ADODB.Connection
        Cn.Open("Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
        ";Uid=" & User_ID & ";Pwd=" & Password & ";")

        rs.Open(SQLStr, Cn, adOpenStatic)
        ' Dump to spreadsheet
        With Worksheets("CK Part History installed").Range("a2:az5000") ' Enter your sheet name and range here
            .CopyFromRecordset(rs)
            lastRow = Worksheets("CK Part History installed").Cells(Rows.Count, 1).End(xlUp).Row
            connect = lastRow
        End With
        '            Tidy up
        rs.Close()
        rs = Nothing
        Cn.Close()
        Cn = Nothing
    End If


    begin = Worksheets("Component PN# history").Cells(Rows.Count, 1).End(xlUp).Row + 1
    rs = New ADODB.Recordset
    SQLStr = "SELECT sysserialno, partno, qty, keypart, installed, eeecode, categoryname, lasteditby, lasteditdt  FROM mfsyscomponent (NOLOCK) WHERE sysserialno IN ( " & BuildString & " ) ORDER BY lasteditdt"

    Cn = New ADODB.Connection
    Cn.Open("Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
    ";Uid=" & User_ID & ";Pwd=" & Password & ";")

    rs.Open(SQLStr, Cn, adOpenStatic)
    ' Dump to spreadsheet
    With Worksheets("Component PN# history").Range("a" & begin & ":az5000") ' Enter your sheet name and range here
        .ClearContents()
        .CopyFromRecordset(rs)
        lastRow = Worksheets("Component PN# history").Cells(Rows.Count, 1).End(xlUp).Row
        connect = lastRow
    End With
    '            Tidy up
    rs.Close()
    rs = Nothing
    Cn.Close()
    Cn = Nothing

End Function