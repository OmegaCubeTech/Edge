Imports System
Imports System.IO
Imports System.Collections
Imports System.Configuration
Imports System.Data
Imports System.Linq
Imports System.Web
Imports System.Web.Security
Imports System.Data.OleDb
Imports System.Web.UI
Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Xml.Linq
Imports System.Drawing
Imports System.Reflection
Imports System.Diagnostics
Imports DevExpress.XtraPrinting
Imports DevExpress.XtraReports.UI
Imports System.Net
Imports System.Text.RegularExpressions

Partial Class Email_All_Vendors
    Inherits System.Web.UI.Page

    Public sid As String
    Dim rptDs_PKFieldNames As DataSet() = New DataSet(1) {}
    Dim pkFieldNamesDS As DataSet

    Public conStr As String = CType(ConfigurationManager.AppSettings("dbconn.ConnectionString"), String)

    Public ds1, ds2 As New Data.DataSet()
    Protected dbad As New OleDbDataAdapter

    Dim dstTempData As New DataSet

    Public conn As New Dbconn
    Dim statusText As String = ""

    Dim str, str1, str2, str3, str4, str5, str6, str7, strTempId, wClause, sf, viewName, toemail, contactsQuery As String
    Dim strQuery As String = String.Empty
    Dim sp1, sp2, sp3
    Dim fName, fValue
    Dim i As Integer
    Dim count As Integer = 0
    Dim rpt As New XtraReport
    Dim rpt1 As New XtraReport

    Public pdfGenerator As New EventControl()
    Public openPDFStr As String
    Dim pdfFileNameSuffix As String = String.Empty
    Public pid As String

    Dim ScriptOpenModalDialog As String = "javascript:OpenModalDialog('{0}','{1}');"


    Dim strRFQNo As String = String.Empty
    Dim rname As String = "PO_RFQ_Report_Qty_Breaks.repx"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            strRFQNo = String.Empty

            str3 = "PO_RFQ_VENDORS_RFQ_NO[_]PO_RFQ_VENDORS_VENDOR_NO[_]"
            str2 = "RFQ_NO[_]VENDOR_NO[_]"

            dbad.SelectCommand = New OleDbCommand
            dbad.InsertCommand = New OleDbCommand
            dbad.UpdateCommand = New OleDbCommand

            If Not Request.QueryString("RFQ_NO") Is Nothing Then
                strRFQNo = Request.QueryString.Get("RFQ_NO")
            Else
                strRFQNo = String.Empty
            End If

            If Not String.IsNullOrEmpty(strRFQNo) Then
                If Not String.IsNullOrEmpty(rname) Then
                    strQuery = String.Empty
                    strQuery = "SELECT RFQ_NO,VENDOR_NO,EMAIL FROM PO_RFQ_VENDORS WHERE EMAIL IS NOT NULL AND RFQ_NO='" & strRFQNo & "'"
                    Dim dsRFQ As New DataSet
                    dsRFQ.Clear()
                    dsRFQ = Execute_Query(strQuery)

                    If Not dsRFQ.Tables(0) Is Nothing Then
                        If dsRFQ.Tables(0).Rows.Count > 0 Then
                            For i As Integer = 0 To dsRFQ.Tables(0).Rows.Count - 1
                                If Not (Equals(dsRFQ.Tables(0).Rows(i)("VENDOR_NO"), System.DBNull.Value)) Then
                                    If dsRFQ.Tables(0).Rows(i)("VENDOR_NO").ToString() <> "" Then
                                        ''str4=900000[_]1000018[_]
                                        str4 = String.Empty
                                        str4 = dsRFQ.Tables(0).Rows(i)("RFQ_NO") & "[_]" & dsRFQ.Tables(0).Rows(i)("VENDOR_NO") & "[_]"

                                        Call Send_Email_To_Vendor(dsRFQ.Tables(0).Rows(i)("RFQ_NO").ToString(), dsRFQ.Tables(0).Rows(i)("VENDOR_NO").ToString(), dsRFQ.Tables(0).Rows(i)("EMAIL").ToString())

                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            End If

            Dim strFile As String = "More"
            Dim strCmd As String = String.Format("alert('Emails sent Successfully to all Vendors.');window.close();", strFile)
            ScriptManager.RegisterStartupScript(Me, Me.GetType(), "nothing", strCmd, True)

            'Dim strFile1 As String = "TEST"
            'Dim strCmd1 As String = String.Format("", strFile1)
            'ScriptManager.RegisterStartupScript(Me, Me.GetType(), "nothing", strCmd1, True)

        Catch ex As Exception
            Call insert_sys_log(RID.Value & "  Page_Load", ex.Message)
        End Try
    End Sub


    Public Function Execute_Query(ByVal strQuery As String) As DataSet
        Dim dsData As New DataSet
        Try
            dbad.SelectCommand.Connection = conn.getconnection()
            dbad.SelectCommand.CommandText = strQuery
            dsData.Clear()
            If dbad.SelectCommand.Connection.State = ConnectionState.Closed Then
                dbad.SelectCommand.Connection.Open()
            End If
            dbad.Fill(dsData)
            If dbad.SelectCommand.Connection.State = ConnectionState.Open Then
                dbad.SelectCommand.Connection.Close()
            End If

        Catch ex As Exception
            dsData.Clear()
            '            Call insert_sys_log(RID.Value & " - " & "Execute_Query", ex.Message.ToString())
        End Try
        Return dsData
    End Function


    Public Sub Send_Email_To_Vendor(ByVal rfqNo As String, ByVal vendorNo As String, ByVal strEmailId As String)
        Try
            If Not String.IsNullOrEmpty(rname) Then
                Dim rnames = Split(rname, ",")
                If rnames.Length > 1 Then
                    Dim rn = Split(rname, ",")
                    Dim rptNew1 As New XtraReport
                    For j As Integer = 0 To rn.Length - 1
                        If j > 0 Then
                            rptNew1 = GenerateReport(str1, rn(j), wClause)
                            rpt1.Pages.AddRange(rptNew1.Pages)
                        Else
                            rpt1 = GenerateReport(str1, rn(j), wClause)
                        End If
                    Next
                    rname = rname.Replace(",", "_")
                Else
                    rpt1 = GenerateReport(str1, rname, wClause)
                End If

                Dim reportPath As String = Server.MapPath(System.Configuration.ConfigurationManager.AppSettings("defaultEmailAttachmentsFolderPath") & "\Vendors")
                If Not Directory.Exists(reportPath) Then
                    Directory.CreateDirectory(reportPath)
                End If
                ' SAVE THE CURRENT FOLDER PATH
                'sp3 = Split(str4, "[_]")

                Dim dsPageName As New DataSet
                Dim attachRepName As String = String.Empty

                dsPageName.Clear()
                dsPageName = Execute_Query("SELECT REPLACE(ENTITY_LABEL,' ','_') AS PAGE_ID  FROM SYS_PAGE_ENTITY  WHERE  PARENT_ENTITY_NAME IS NULL AND ENTITY_LABEL IS NOT NULL AND UPPER(PAGE_ID) = UPPER('PO_RFQ')")

                If Not dsPageName.Tables(0) Is Nothing Then
                    If dsPageName.Tables(0).Rows.Count > 0 Then
                        If Not Equals(dsPageName.Tables(0).Rows(0)(0), System.DBNull.Value) Then
                            If Not Equals(dsPageName.Tables(0).Rows(0)(0), "_") Then
                                attachRepName = dsPageName.Tables(0).Rows(0)(0).ToString()
                            Else
                                attachRepName = pid
                            End If
                        Else
                            attachRepName = pid
                        End If
                    Else
                        attachRepName = pid
                    End If
                Else
                    attachRepName = pid
                End If


                'Dim currentFileName As String = (rname.ToUpper.Replace(".REPX", "")).Trim & "_" & sp3(0).ToString.Trim & ".pdf"
                Dim currentFileName As String = attachRepName.Trim & "_" & sp3(0).ToString.Trim & ".pdf"
                attachmentFileName.Value = currentFileName

                ''reportPath = String.Concat(reportPath, "\\", currentFileName) '''' (reportPath & "\") + currentFileName

                reportPath = Server.MapPath(System.Configuration.ConfigurationManager.AppSettings("defaultEmailAttachmentsFolderPath") & "\Vendors\" & currentFileName)
                attachmentFolderName.Value = reportPath
                rpt1.ExportToPdf(reportPath)

                Call Send_Mail(rfqNo, vendorNo, strEmailId)

            End If
        Catch ex As Exception
            Call insert_sys_log(RID.Value & " - " & "Send_Email_To_Vendor", ex.Message.ToString())
        End Try
    End Sub

    Public Sub Send_Mail(ByVal rfqNo As String, ByVal vendorNo As String, ByVal Email_ID As String)
        Try
            Dim ds As New DataSet()
            Dim cmd As New OleDbCommand()

            'Dim sqlStr As String = "select max(to_number(sno)) + 1 as newSNo from sys_generic_email"
            Dim sqlStr As String = "select NVL(max(to_number(sno)),999) + 1 as newSNo from sys_generic_email"
            Dim defaultAttachmentsFolder As String = ""
            Dim _newSNo As String = Nothing

            Using conn1 As New OleDbConnection(conStr)
                cmd.Connection = conn1
                cmd.CommandText = sqlStr
                cmd.CommandType = CommandType.Text

                Dim oDa As New OleDbDataAdapter(cmd)

                oDa.Fill(ds)

                If IsDBNull(ds.Tables(0).Rows(0)("NEWSNO").ToString) Then
                    _newSNo = "1"
                Else
                    _newSNo = ds.Tables(0).Rows(0)("NEWSNO").ToString()
                End If


                Dim sqlNewStr As String = String.Empty


                sqlNewStr = "insert into sys_generic_email (sno, to_address,mail_subject,created_by, created_date,SOURCE_NAME, SOURCE_TYPE,SEND_CONFIRMATION )" _
                               + " values ('" + _newSNo + "','" + Replace(Email_ID, "'", "''") + "','Request for Quotation : " & rfqNo & "','" + Replace(HttpContext.Current.Session("userid"), "'", "''") + "',sysdate,'RFQ','EMAIL','P')"

                cmd.CommandText = sqlNewStr
                cmd.CommandType = CommandType.Text
                cmd.Connection.Open()
                Dim count As Integer = cmd.ExecuteNonQuery()


                ' B: CREATE A FOLDER UNDER "emailAttachments"
                Dim currentFolderPath As String = Server.MapPath(System.Configuration.ConfigurationManager.AppSettings("folderPath"))
                Dim directoryString As String = (currentFolderPath & "\") + _newSNo

                If Not Directory.Exists(directoryString) Then
                    Directory.CreateDirectory(directoryString)
                    newDirectoryPath.Value = directoryString
                End If



                ' SAVE THE CURRENT FOLDER PATH
                Dim sourceFile As String = attachmentFolderName.Value
                Dim destinationFile As String = (directoryString & "\") + attachmentFileName.Value

                If Not String.IsNullOrEmpty(attachmentFolderName.Value) AndAlso File.Exists(sourceFile) Then
                    If File.Exists(destinationFile) Then
                        File.Delete(destinationFile)
                    End If
                    File.Move(sourceFile, destinationFile)
                End If




                Dim _directoryString As String = (currentFolderPath & "\") + _newSNo

                Dim di As New DirectoryInfo(_directoryString & "\")
                Dim rgFiles As FileInfo() = di.GetFiles("*.*")
                Dim strFiles As String = ""
                For Each fi As FileInfo In rgFiles
                    If strFiles <> "" Then
                        If strFiles.Length >= 3900 Then
                            Exit For
                        Else
                            strFiles = (strFiles & ",") + fi.FullName
                        End If
                    Else
                        strFiles = strFiles + fi.FullName
                    End If
                Next

                Dim _emailBody As String = "Please find attached request for quote# " & rfqNo & ". Please review and return your quote to me as noted." '"Request for Quotation :  " & rfqNo & " Attached"
                'Dim _emailBody As String = Editor1.Text



                ' F: DELETE THE DEFAULT REPORT FROM "TEMP_EMAIL_DEFAULT_ATTACHMENTS"
                If File.Exists(sourceFile) Then
                    File.Delete(sourceFile)
                End If

                Try
                    '634798  Niharika (change in update stmt)
                    Dim ds_list As New Data.DataSet
                    ds_list.Clear()
                    ds_list = Return_record_set("select 'C:\Omegacube_ERP\PORTALS\PELHAM_DEV\'||FILE_PATH||FILE_NAME FROM sys_linked_docs WHERE page_id='PO_RFQ' AND PRIMAY_KEY1='" & rfqNo & "'")
                    If ds_list.Tables(0).Rows.Count > 0 Then
                        For Each dr In ds_list.Tables(0).Rows
                            If strFiles = "" Then
                                strFiles = strFiles & dr(0) & ","
                            Else
                                strFiles = strFiles & "," & dr(0)
                            End If
                        Next
                    End If
                Catch ex As Exception

                End Try
                conn.update_clob(_emailBody, "UPDATE SYS_GENERIC_EMAIL SET MAIL_BODY=:TEXT_DATA WHERE SNO='" + _newSNo + "'")
                cmd.Connection.Close()
                cmd.CommandText = "UPDATE SYS_GENERIC_EMAIL SET SEND_CONFIRMATION='N',attachment='" + strFiles.Replace("'", "''") + "' WHERE SNO='" + _newSNo + "'"
                cmd.CommandType = CommandType.Text
                cmd.Connection.Open()
                Dim count1 As Integer = cmd.ExecuteNonQuery()
                cmd.Connection.Close()

                'Dim strFile1 As String = "TEST"
                'Dim strCmd1 As String
                'strCmd1 = String.Format("window.close();", strFile1)
                'ScriptManager.RegisterStartupScript(Me, Me.GetType(), "nothing", strCmd1, True)
            End Using
        Catch ex As Exception
            'lblErr.Text = ""
            'lblErr.Text = ex.Message.ToString
            Call insert_sys_log("Email Page btnSend1", ex.Message)
        End Try
    End Sub

    Public Function Return_record_set(ByVal str_select As String) As System.Data.DataSet
        Try
            Dim ds_new As New System.Data.DataSet
            dbad.SelectCommand = New OleDbCommand
            dbad.SelectCommand.Connection = conn.getconnection()
            dbad.SelectCommand.CommandText = str_select
            dbad.Fill(ds_new)
            dbad.SelectCommand.Connection.Close()
            Return ds_new
        Catch ex As Exception
            Call insert_sys_log("Return_record_set", ex.Message)
        End Try
    End Function
    Public Function GenerateReport(ByVal str1 As String, ByVal rname As String, ByVal wCluase As String) As XtraReport
        Try
            If rname.ToLower.IndexOf("repx") <> -1 Then
                Dim strQuery As String
                Dim tName As String = String.Empty
                Dim rpt As New XtraReport
                str = Server.MapPath("SourceReports") & "\" & rname
                If Not (System.IO.File.Exists(str)) Then
                    Response.Write("Report file name doesn't exist n sourcereports folder")
                    Exit Function
                End If
                rpt.LoadLayout(str)

                '' Fo Loading CSS
                If File.Exists(Server.MapPath("SourceReports\CSS\OCT_REP_CSS.repss")) Then
                    rpt.StyleSheetPath = Server.MapPath("SourceReports\CSS\OCT_REP_CSS.repss")
                End If
                str5 = rpt.DataMember.ToString
                If str5 = "" Then
                    Response.Write("Issue in the report dataset.")
                    Exit Function
                End If

                ''Dim paramsList1 As New DevExpress.XtraReports.Native.Parameters.ReportParameterCollection(rpt)

                Dim paraMetersList As New DevExpress.XtraReports.Parameters.ParameterCollection

                paraMetersList = rpt.Parameters

                If paraMetersList.Count > 0 Then
                    For Each ParameterInfo As DevExpress.XtraReports.Parameters.Parameter In paraMetersList
                        If ParameterInfo.Name = "USER_ID" Then
                            rpt.Parameters("USER_ID").Value = Session("user_id").ToString()
                        End If
                    Next
                End If

                sp1 = Split(str2, "[_]")
                sp2 = Split(str3, "[_]")
                sp3 = Split(str4, "[_]")
                str7 = rpt.FilterString

                dbad.SelectCommand = New OleDbCommand
                dbad.SelectCommand.Connection = conn.getconnection()
                Dim dsp, dsp1 As New Data.DataSet

                If (str7 <> "") Then
                    str7 = Replace(str7, "[", "")
                    str7 = Replace(str7, "]", "")
                    str6 = str5 & " Where "
                    For i = 0 To UBound(sp1)
                        If (sp1(i) <> "") Then
                            'str7 = Replace(str7, "Parameters." & sp1(i), "'" & sp3(i) & "'")
                            str7 = Replace(str7, "?" & sp1(i), "'" & sp3(i) & "'")
                        End If
                    Next
                Else
                    If (UBound(sp1) > 0) Then
                        str6 = str5 & " Where "
                        For i = 0 To UBound(sp1)
                            If (sp1(i) <> "") Then
                                str6 = str6 & sp1(i) & "='" & sp3(i) & "' and "
                            End If
                        Next
                        str6 = Mid(str6, 1, Len(str6) - 4)
                    Else
                        str6 = str5
                    End If
                End If
                str6 = str6 & str7

                If Not String.IsNullOrEmpty(rpt.DataSourceSchema.ToString()) Then
                    Dim xmlSchema As String = rpt.DataSourceSchema.ToString()
                    Dim MemStream As New IO.MemoryStream(System.Text.Encoding.ASCII.GetBytes(xmlSchema))
                    MemStream.Position = 0
                    dsp1.ReadXmlSchema(MemStream)


                    Dim str8 As String
                    str8 = str6 & str7
                    str6 = str5 & str6 & str7

                    Dim i1 As Integer = dsp1.Tables.Count

                    Dim i2, i3 As Integer
                    If (i1 > 1) Then

                        For i2 = 0 To i1 - 1
                            dbad.SelectCommand.CommandText = "select * from " & dsp1.Tables(i2).ToString() & " " & str8
                            dbad.Fill(dsp, dsp1.Tables(i2).ToString())
                        Next
                    Else
                        If str6.ToLower.IndexOf("where") <> -1 AndAlso Not String.IsNullOrEmpty(wCluase) Then
                            strQuery = ""
                            strQuery = wClause
                        Else
                            strQuery = String.Concat(" where ", wClause)
                        End If

                        If Not String.IsNullOrEmpty(wCluase) Then
                            If Not strQuery.ToString.ToUpper.Contains("WHERE") Then
                                dbad.SelectCommand.CommandText = "select * from " & str5 & " where " & strQuery
                            Else
                                dbad.SelectCommand.CommandText = "select * from " & str5 & " " & strQuery
                            End If
                            ''dbad.SelectCommand.CommandText = "select * from " & str5 & " where " & strQuery '' Old
                        Else
                            dbad.SelectCommand.CommandText = "select * from " & str6
                        End If
                        'dbad.SelectCommand.CommandText = "select * from " & str6
                        dbad.Fill(dsp)
                    End If

                    i3 = dsp1.Relations.Count
                    Dim sss1, sss2, sss3, sss4 As String
                    Dim i4, i5, i6 As Integer

                    Dim drps(), drcs() As Data.DataColumn

                    For i2 = 0 To i3 - 1
                        drps = dsp1.Relations(i2).ParentColumns
                        drcs = dsp1.Relations(i2).ChildColumns
                        i4 = drcs.Length
                        Dim dr1(i4 - 1), dr2(i4 - 1) As Data.DataColumn

                        For i5 = 0 To i4 - 1
                            Dim drps1, drcs1 As New Data.DataColumn
                            sss1 = dsp1.Relations(i2).ParentTable.ToString
                            sss2 = dsp1.Relations(i2).ChildTable.ToString
                            sss3 = drps(i5).ColumnName
                            sss4 = drcs(i5).ColumnName
                            drps1 = dsp.Tables(sss1).Columns(sss3)
                            drcs1 = dsp.Tables(sss2).Columns(sss4)
                            dr1(i5) = drps1
                            dr2(i5) = drcs1
                        Next
                        Dim drst As New Data.DataRelation(dsp1.Relations(i2).RelationName, dr1, dr2, True)
                        dsp.Relations.Add(drst)
                    Next

                    If (dsp.Tables(0).Rows.Count > 0) Then
                        'Session("data_adop") = dbad
                        ' Session("dataset") = dsp
                        rpt.DataAdapter = dbad
                        rpt.DataSource = dsp
                        rpt.DataMember = dsp.Tables(0).TableName
                        If (str7 <> "") Then
                            For i = 0 To UBound(sp1)
                                If (sp1(i) <> "") Then
                                    rpt.Parameters(sp1(i)).Value = sp3(i)
                                End If
                            Next
                        End If

                        'ReportViewer1.Report = rpt
                        'Session("rpt_class") = String.Empty
                        'Session("rpt_class") = rpt
                    End If

                Else  '' For Single Table or View 

                    If str6.ToLower.IndexOf("where") <> -1 AndAlso Not String.IsNullOrEmpty(wCluase) Then
                        strQuery = ""
                        'strQuery = String.Concat(" and ", wClause)
                        strQuery = wClause
                    Else
                        strQuery = String.Concat(" where ", wClause)
                    End If

                    If Not String.IsNullOrEmpty(wCluase) Then
                        If Not strQuery.ToString.ToUpper.Contains("WHERE") Then
                            dbad.SelectCommand.CommandText = "select * from " & str5 & " where " & strQuery
                        Else
                            dbad.SelectCommand.CommandText = "select * from " & str5 & " " & strQuery
                        End If
                        ''dbad.SelectCommand.CommandText = "select * from " & str5 & " where " & strQuery '' Old
                        tName = String.Empty
                        tName = str5
                    Else
                        dbad.SelectCommand.CommandText = "select * from " & str6
                        tName = str6
                    End If

                    dsp.Clear()
                    dbad.Fill(dsp, tName)

                    'rpt.DataAdapter = dbad
                    'rpt.DataSource = dsp
                    'rpt.DataMember = dsp.Tables(0).TableName

                    rpt.DataSource = dsp
                    rpt.DataMember = dsp.Tables(0).TableName
                    rpt.DataAdapter = dbad

                    If (str7 <> "") Then
                        For i = 0 To UBound(sp1)
                            If (sp1(i) <> "") Then
                                rpt.Parameters(sp1(i)).Value = sp3(i)
                            End If
                        Next
                    End If
                End If

                rpt.CreateDocument()
                Return rpt
                dbad.SelectCommand.Connection.Close()
                Exit Function

            Else
                If (UCase(str1) = "CS-REPORT" And str2 <> "" And str3 <> "" And str4 <> "" And rname <> "") Then
                    If (rname <> "") Then
                        Dim ass, curr_ass As Assembly
                        For Each ass In AppDomain.CurrentDomain.GetAssemblies
                            If (ass.FullName.Contains("App_Code")) Then
                                curr_ass = ass
                                Dim curr_t, t As Type
                                For Each t In curr_ass.GetExportedTypes
                                    If (t.Name = rname) Then
                                        curr_t = t
                                        Dim rpt As New XtraReport

                                        rpt = Activator.CreateInstance(curr_t)

                                        '' Fo Loading CSS
                                        If File.Exists(Server.MapPath("SourceReports\CSS\OCT_REP_CSS.repss")) Then
                                            rpt.StyleSheetPath = Server.MapPath("SourceReports\CSS\OCT_REP_CSS.repss")
                                        End If

                                        str5 = rpt.DataMember.ToString

                                        Dim s As String = rpt.DataAdapter.ToString

                                        If str5 = "" Then
                                            Response.Write("Issue in the report dataset.")
                                            Exit Function
                                        End If
                                        sp1 = Split(str2, "[_]")
                                        sp2 = Split(str3, "[_]")
                                        sp3 = Split(str4, "[_]")
                                        str7 = rpt.FilterString
                                        If (str7 <> "") Then
                                            str7 = Replace(str7, "[", "")
                                            str7 = Replace(str7, "]", "")
                                            str6 = " Where "
                                            For i = 0 To UBound(sp1)
                                                If (sp1(i) <> "") Then
                                                    'str7 = Replace(str7, "Parameters." & sp1(i), "'" & sp3(i) & "'")
                                                    str7 = Replace(str7, "?" & sp1(i), "'" & sp3(i) & "'")
                                                End If
                                            Next
                                        Else
                                            If (UBound(sp1) > 0) Then
                                                str6 = " Where "
                                                For i = 0 To UBound(sp1)
                                                    If (sp1(i) <> "") Then
                                                        str6 = str6 & sp1(i) & "='" & sp3(i) & "' and "
                                                    End If
                                                Next
                                                str6 = Mid(str6, 1, Len(str6) - 4)
                                            End If
                                        End If
                                        'str6 = str6 & str7
                                        dbad.SelectCommand = New OleDbCommand
                                        dbad.SelectCommand.Connection = conn.getconnection()
                                        Dim dsp, dsp1 As New Data.DataSet

                                        dsp1 = rpt.DataSource

                                        Dim str8 As String
                                        str8 = str6 & str7
                                        str6 = str5 & str6 & str7

                                        Dim i1 As Integer = dsp1.Tables.Count
                                        Dim i2, i3 As Integer
                                        If (i1 > 1) Then

                                            For i2 = 0 To i1 - 1

                                                dbad.SelectCommand.CommandText = "select * from " & dsp1.Tables(i2).ToString() & " " & str8
                                                dbad.Fill(dsp, dsp1.Tables(i2).ToString())

                                            Next
                                        Else
                                            dbad.SelectCommand.CommandText = "select * from " & str6
                                            dbad.Fill(dsp)

                                        End If
                                        i3 = dsp1.Relations.Count
                                        Dim sss1, sss2, sss3, sss4 As String
                                        Dim i4, i5, i6 As Integer

                                        Dim drps(), drcs() As Data.DataColumn

                                        For i2 = 0 To i3 - 1
                                            drps = dsp1.Relations(i2).ParentColumns
                                            drcs = dsp1.Relations(i2).ChildColumns
                                            i4 = drcs.Length
                                            Dim dr1(i4 - 1), dr2(i4 - 1) As Data.DataColumn

                                            For i5 = 0 To i4 - 1
                                                Dim drps1, drcs1 As New Data.DataColumn
                                                sss1 = dsp1.Relations(i2).ParentTable.ToString
                                                sss2 = dsp1.Relations(i2).ChildTable.ToString
                                                sss3 = drps(i5).ColumnName
                                                sss4 = drcs(i5).ColumnName
                                                drps1 = dsp.Tables(sss1).Columns(sss3)
                                                drcs1 = dsp.Tables(sss2).Columns(sss4)
                                                dr1(i5) = drps1
                                                dr2(i5) = drcs1
                                            Next
                                            Dim drst As New Data.DataRelation(dsp1.Relations(i2).RelationName, dr1, dr2, True)
                                            ''Dim drst1 As New Data.DataRelation(,,,
                                            dsp.Relations.Add(drst)
                                        Next

                                        If (dsp.Tables(0).Rows.Count > 0) Then
                                            Session("data_adop") = dbad
                                            Session("dataset") = dsp
                                            'rpt.DataAdapter = dbad
                                            'rpt.DataSource = dsp

                                            rpt.DataAdapter = dbad
                                            rpt.DataSource = dsp
                                            rpt.DataMember = dsp.Tables(0).TableName
                                            If (str7 <> "") Then
                                                For i = 0 To UBound(sp1)
                                                    If (sp1(i) <> "") Then
                                                        rpt.Parameters(sp1(i)).Value = sp3(i)
                                                    End If
                                                Next
                                            End If

                                        End If
                                        dbad.SelectCommand.Connection.Close()
                                        Return rpt
                                        Exit Function
                                    End If
                                Next
                            End If
                        Next
                    End If
                ElseIf (UCase(str1) = "CS-REPORT-DS PARAM" And str2 <> "" And str3 <> "" And str4 <> "" And rname <> "") Then
                    If (rname <> "") Then
                        Dim ass, curr_ass As Assembly
                        For Each ass In AppDomain.CurrentDomain.GetAssemblies
                            If (ass.FullName.Contains("App_Code")) Then
                                curr_ass = ass
                                Dim curr_t, t As Type
                                For Each t In curr_ass.GetExportedTypes
                                    If (t.Name = rname) Then
                                        curr_t = t
                                        Dim rpt As New XtraReport
                                        Dim spp() As String
                                        sp1 = Split(str2, "[_]")
                                        sp3 = Split(str4, "[_]")
                                        ReDim Preserve spp(UBound(sp3) - 1)
                                        For i = 0 To UBound(sp3) - 1
                                            spp(i) = sp3(i)
                                        Next

                                        rpt = Activator.CreateInstance(curr_t, spp)
                                        '' Fo Loading CSS
                                        If File.Exists(Server.MapPath("SourceReports\CSS\OCT_REP_CSS.repss")) Then
                                            rpt.StyleSheetPath = Server.MapPath("SourceReports\CSS\OCT_REP_CSS.repss")
                                        End If
                                        str7 = rpt.FilterString
                                        If (str7 <> "") Then
                                            For i = 0 To UBound(sp1)
                                                If (sp1(i) <> "") Then
                                                    rpt.Parameters(sp1(i)).Value = sp3(i)
                                                End If
                                            Next
                                        End If
                                        Return rpt
                                        Exit Function
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If
            End If
            count = count + 1
            Return rpt
        Catch ex As Exception
            Call insert_sys_log(RID.Value & " -  GenerateReport", ex.Message)
        End Try
    End Function


    Public Sub insert_sys_log(ByVal str1 As String, ByVal message As String)

        Dim sterr1, sterr2, sterr3, sterr4, sterr As String
        sterr = Replace(message, "'", "''")
        If (Len(sterr) > 4000) Then
            sterr1 = Mid(sterr, 1, 4000)
            If (Len(sterr) > 8000) Then
                sterr2 = Mid(sterr, 4000, 8000)
                If (Len(sterr) > 12000) Then
                    sterr3 = Mid(sterr, 8000, 12000)
                    If (Len(sterr) > 16000) Then
                        sterr4 = Mid(sterr, 12000, 16000)
                    Else
                        sterr4 = Mid(sterr, 12000, Len(sterr))
                    End If
                Else
                    sterr3 = Mid(sterr, 8000, Len(sterr))
                    sterr4 = ""
                End If
            Else
                sterr2 = Mid(sterr, 4000, Len(sterr))
                sterr3 = ""
                sterr3 = ""
                sterr4 = ""
            End If
        Else
            sterr1 = sterr
            sterr2 = ""
            sterr3 = ""
            sterr4 = ""
        End If
        dbad.InsertCommand = New OleDbCommand
        dbad.InsertCommand.Connection = conn.getconnection()
        dbad.InsertCommand.CommandText = "Insert into SYS_ACTIVATE_STATUS_LOG (LINE_NO, CHANGE_REQUEST_NO,  OBJECT_TYPE, OBJECT_NAME, ERROR_TEXT, STATUS,LOG_DATE,ERROR_TEXT1, ERROR_TEXT2, ERROR_TEXT3) values ((select nvl(max(to_number(line_no)),0)+1 from SYS_ACTIVATE_STATUS_LOG),'','EmailPage','" & str1 & "','" & sterr1 & "','N',sysdate,'" & sterr2 & "','" & sterr3 & "','" & sterr4 & "')"
        dbad.InsertCommand.ExecuteNonQuery()
        dbad.InsertCommand.Connection.Close()
    End Sub
End Class
