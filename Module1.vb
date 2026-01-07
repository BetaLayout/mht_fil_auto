'VER 24 04/11/2025

Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports RestSharp
Imports System.Net
Imports System.Text
Imports System.Data.OleDb
Imports RestSharp.Contrib

Module Module1

    '////////////////

    Private Const FOLDER_CONFIG_CACHE_FILE As String = "S:\OMS_DATA\folder_config_cache.json"

    'Set as defaults which are those used by Foxpro database in Shannon
    Private inputDirectory As String = "C:\Source"
    Private precamFinishedFolder As String = "S:\Ordersforcam\ForCAM"
    Private precamZipFilesSource As String = "S:\Job\CAM-India\Work_Finished"
    Private precamUploadFilesSource As String = "S:\Job\CAM-India\Upload"
    Private archiveUploadFilesTemp As String = "s:\uploaded_files_temp"
    Private statusUpdates As String = "S:\Job\CAM-India\Work\new_data_for_work\Status_Updates"
    Private commentUpdates As String = "S:\Job\CAM-India\Work\new_data_for_work\comments_updates"
    Private errorLogLocation As String = "T:\IN_HOUSE_SOFTWARE\ALL_SOFTWARE\Visual_Studio_2010\OMS-Download\TEMP"
    Private notesLocation As String = "T:\Database3\NOTES"

    '///

    Public foxproConnection As Boolean = False
    Public sqlConnection As Boolean = False
    Public useFoxproConnection As Boolean = False
    Public useSqlConnection As Boolean = False

    'Private Const CACHE_FILE As String = "S:\OMS_DATA\dev_token_cache.json" ''Devlopment
    Private Const CACHE_FILE As String = "S:\OMS_DATA\token_cache.json"

    Dim grantType As String = "client_credentials"
    Dim clientID As String = "9"
    Dim clientSecret As String = "PlJHMRDPpvxiQrSP4FYhbS69hWT2YtWsahmyu5DE"
    ' Dim clientScope As String = "create-tracking-card"
    Dim clientScope As String = "beta-full-access-scope create-tracking-card"

    Dim bearer As String = "0"

    '////

    Dim times_run As Integer
    Dim myarraylist As Object
    Dim filepartlist As Object
    Dim uploadlist As Object
    Dim new_orders As Integer
    Dim on_hold_message As Boolean

    ' ----------

    Dim AppendFileContent1 As String
    Dim quantity As Int32
    Dim quantity_str As String
    Dim x_size As Double
    Dim x_size_str As String
    Dim y_size As Double
    Dim y_size_str As String
    Dim area As Decimal
    Dim area_str As String
    Dim num_small_pcb As Integer
    Dim all_x As String()
    Dim all_y As String()
    Dim all_Qty As String()
    Dim num_of_mht As Integer
    Dim num_of_mht_ori As Integer
    Dim precam_inst As Boolean
    Dim small_filled As Boolean
    Dim contains_small As Boolean
    Dim bigger As Boolean
    Dim rout As Boolean
    Dim conrad_orders As Boolean
    Dim MHTArray(30, 1) As String
    Dim d1 As DateTime = DateTime.Now
    Dim pool_array(200, 1) As String
    Dim series_flag As Boolean
    Dim data_update_done As Boolean = False

    Dim donelist As Object

    Sub get_uploads(ByVal source As String)

        Dim au As String
        Dim dot As Integer

        Dim di3 As DirectoryInfo = New DirectoryInfo(source) ' Original Source directory

        Dim files3 As Object

        files3 = di3.GetFileSystemInfos()
        For Each file In files3
            If InStr(UCase(file.name), "GERBERS") > 0 Then
                dot = InStr(UCase(file.name), "GERBERS")
                au = Mid(file.name, 1, dot - 2)
                uploadlist.add(au)
                new_orders = True
            End If
        Next

    End Sub


    Sub check_for_uploads(ByVal source As String)
        Dim au As String
        Dim dot As Integer

        Dim pdf, zip As Boolean

        Dim di3 As DirectoryInfo = New DirectoryInfo(source) ' Original Source directory

        Dim files3 As Object

        files3 = di3.GetFileSystemInfos()
        For Each order In files3
            If InStr(UCase(order.name), "GERBERS") > 0 Then
                dot = InStr(UCase(order.name), "GERBERS")
                au = Mid(order.name, 1, dot - 2)

                pdf = False
                zip = False

                For Each order1 In files3
                    If InStr(UCase(order1.name), au) > 0 Then
                        If InStr(UCase(order1.name), "PDF") > 0 Then pdf = True
                        If (InStr(UCase(order1.name), ".ZIP") > 0 And InStr(UCase(order1.name), "GERBER") = 0) Then zip = True
                    End If
                    If (zip = True And pdf = True) Then Exit For
                Next


                If zip = False Or pdf = False Then GoTo not_this_time

                If Not (uploadlist.contains(au)) Then
                    uploadlist.add(au)
                    GoTo not_this_time
                End If


                '------------------------------------

                '//TODO Sort out these s drive locations 
                For Each order1 In files3
                    If InStr(UCase(order1.name), au) > 0 Then
                        Try
                            File.Copy(order1.fullname, "s:\uploaded_files_done\" & LCase(order1.name))
                        Catch
                        End Try

                        wait(100)

                        Try
                            File.Move(order1.fullname, "s:\uploaded_files_temp\" & LCase(order1.name))
                        Catch
                        End Try
                    End If
                Next


                '************************************

            End If


not_this_time:

        Next

    End Sub

    Public Class DataItem
        Public Property rows As List(Of Object)
        Public Property extras As Object
    End Class


    Private Sub wait(ByVal interval As Integer)
        'Dim i As Integer
        ' Do While i < interval
        ' i = i + 1
        ' Loop
        Threading.Thread.Sleep(interval)
    End Sub

    Private Sub SendErrorLog(ByVal fileName As String, ByVal message As String)
        Directory.CreateDirectory(errorLogLocation)
        Using errorLogFile As New StreamWriter(errorLogLocation + "\" + fileName + ".txt", True)
            errorLogFile.WriteLine(message)
        End Using
    End Sub

    Public Class FolderConfigCache
        Public Property path As String
        Public Property name As String
    End Class

    '//////

    Sub Main()

        getBearer()

        If bearer = "0" Or bearer.Length < 500 Then
            MsgBox("Issue getting bearer token, Cannot cnnect to database!. Try dain later", MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        '//////////////////////////////////////////////////////////////////


        '//Is SQL Available
        Dim response As DataItem
        Try
            sqlConnection = True
            Dim jsonString1 As String = JsonConvert.SerializeObject(New With {
       .fields = {"thickness"},
       .table = "OMS.orders_info",
       .conditions = {
                        New Condition With {
                               .field = "id",
                               .op = "=",
                               .value = "94851"
                                }
                             },
       .all_data = True
   })

            Dim encodedJson As String = HttpUtility.UrlEncode(jsonString1)
            'Dim uri As String = $"http://localhost:8009/api/omsgetdata?req={encodedJson}"  'Development
            Dim uri As String = $"https://internal.beta-layout.com/oms/api/omsgetdata?req={encodedJson}"
            response = SendGETRequest(uri)
        Catch
            response = Nothing
            sqlConnection = False
        End Try

        Try
            Dim items = response.rows.ToArray
            For Each item In items
                Dim jsonObject As JObject = DirectCast(item, JObject)
                If jsonObject("thickness") = "1_60MM_MATERIAL" Then
                    sqlConnection = True
                Else
                    sqlConnection = False
                End If
                Exit For
            Next
        Catch
            sqlConnection = False
        End Try

        '//Is Foxpro available
        If File.Exists("T:\Database3\names.dbf") Then
            foxproConnection = True
        Else
            foxproConnection = False
            If sqlConnection = False Then
                Exit Sub
            End If
        End If

        If sqlConnection = True And File.Exists("S:\OMS_DATA\Applications\mht_fill_auto\Release\UseSQL.txt") Then
            useFoxproConnection = False
            useSqlConnection = True
        End If
        If foxproConnection = True And Not File.Exists("S:\OMS_DATA\Applications\mht_fill_auto\Release\UseSQL.txt") Then
            useFoxproConnection = True
            useSqlConnection = False
        End If

        If useFoxproConnection = False And useSqlConnection = False Then
            MsgBox("No suitable database connection available ", vbExclamation)
            Exit Sub
        End If

        If useFoxproConnection = True Then
            If sqlConnection = True Then
                MsgBox("Using FoxPro database AND will also update SQL database :-)")
            Else
                MsgBox("Using FoxPro database BUT UNABLE TO update SQL database ", vbExclamation)
            End If
        End If

        If useSqlConnection = True Then
            If foxproConnection = True Then
                MsgBox("Using SQL database AND will also update FoxPro database :-)")
            Else
                MsgBox("Using SQL database BUT UNABLE TO update FoxPro database ", vbExclamation)
            End If
        End If

        '///////////////

        If useSqlConnection = True Then
            'Update Folders
            If File.Exists(FOLDER_CONFIG_CACHE_FILE) Then
                Dim path As String
                Dim name As String

                Dim inputFolder As Boolean = False
                Dim finishedPrecamFolder As Boolean = False
                Dim precamZip As Boolean = False
                Dim precamUpload As Boolean = False
                Dim archiveUpload As Boolean = False
                Dim status As Boolean = False
                Dim comments As Boolean = False
                Dim errorLogs As Boolean = False
                Dim notes As Boolean = False


                Dim jsonString As String = File.ReadAllText(FOLDER_CONFIG_CACHE_FILE)

                ' Deserialize into a List of FolderConfigCache objects'
                Dim cacheDatas As List(Of FolderConfigCache) = JsonConvert.DeserializeObject(Of List(Of FolderConfigCache))(jsonString)

                For Each item As FolderConfigCache In cacheDatas
                    If item IsNot Nothing Then
                        path = item.path
                        name = item.name
                        Select Case name
                            Case "input_directory"
                                inputFolder = True
                                inputDirectory = path
                            Case "finished_precam_folder"
                                finishedPrecamFolder = True
                                precamFinishedFolder = path
                            Case "precam_zip"
                                precamZip = True
                                precamZipFilesSource = path
                            Case "precam_upload"
                                precamUpload = True
                                precamUploadFilesSource = path
                            Case "archive_upload"
                                archiveUpload = True
                                archiveUploadFilesTemp = path
                            Case "status"
                                status = True
                                statusUpdates = path
                            Case "comments"
                                comments = True
                                commentUpdates = path
                            Case "error_message_folder"
                                errorLogs = True
                                errorLogLocation = path
                            Case "notes_folder"
                                notes = True
                                notesLocation = path
                        End Select
                    End If
                Next

                If notes = False Or inputFolder = False Or finishedPrecamFolder = False Or precamZip = False Or precamUpload = False Or archiveUpload = False Or status = False Or comments = False Or errorLogs = False Then
                    MsgBox("Please Complete Folder Configerations using the OMS application")
                    Exit Sub
                End If
            Else
                MsgBox("Cannot find folder configurations file" + FOLDER_CONFIG_CACHE_FILE)
                Exit Sub
            End If
        End If

        Dim folderwatch As Integer

        ' folderwatch = MsgBox("1) Are FolderWatch Programs running  AND   2) Is all ok to continue ?", vbQuestion + vbYesNo, "Continue?")
        folderwatch = MsgBox("Are FolderWatch Programs running ?", vbQuestion + vbYesNo, "Continue?")

        If folderwatch = 7 Then GoTo leave_program

        Directory.CreateDirectory("s:\oms_data\applications\mht_fill_auto")
        Using copy_directory_batch_file As New StreamWriter("s:\oms_data\applications\mht_fill_auto\copy_directory.bat", False)
            Dim line As String
            ' Write xcopy command for precamFinishedFolder
            line = $"xcopy C:\Source ""{precamFinishedFolder}"" /s/e/y/c"
            copy_directory_batch_file.WriteLine(line)
            ' Write xcopy command for backups
            line = "xcopy C:\Source C:\Backups /s/e/y/c"
            copy_directory_batch_file.WriteLine(line)
            ' Write delete command
            line = "del C:\Source\*.* /s/q"
            copy_directory_batch_file.WriteLine(line)
        End Using

        Dim p = New Process()
        p.StartInfo.FileName = "S:\OMS_DATA\applications\mht_fill_auto\copy_directory.bat"

        on_hold_message = False
        myarraylist = CreateObject("System.Collections.ArrayList")
        filepartlist = CreateObject("System.Collections.ArrayList")
        uploadlist = CreateObject("System.Collections.ArrayList")

        times_run = 0

        Dim source As String
        source = precamZipFilesSource
        Dim upload_source As String
        upload_source = precamUploadFilesSource

check_again:

        donelist = CreateObject("System.Collections.ArrayList")
        donelist.clear()

        If foxproConnection = True Then
            If Not (File.Exists("S:\i_am_here.txt")) Then
                MsgBox("Appears to be something wromg with network in Shannon")
                GoTo no_network
            End If
            If Not (File.Exists("T:\i_am_here.txt")) Then
                MsgBox("Appears to be something wromg with network in Shannon")
                GoTo no_network
            End If
        End If

        If times_run = 0 Then
            If foxproConnection = True Then get_uploads(upload_source) ' just initialize list of au's in the folder
            get_orders(source) ' just initialize list of au's in the folder
        Else
            new_orders = 0
            check_for_orders(source)
            wait(500)

            If new_orders > 0 Then
                UNZIP()
                wait(15000)
                update_database()
                wait(500)
                Notes()
                wait(500)

                p.Start()
                p.WaitForExit()

            End If


            If foxproConnection = True Then check_for_uploads(upload_source)  'Uploads data 

        End If



        times_run = times_run + 1
        Console.WriteLine(" ********************* ")

no_network:

        wait(450000)  '  about 10 mins '450000

        'wait(5000)  ' For Testing


        GoTo check_again

leave_program:

    End Sub

    Public Class Update
        Public Property field As String
        Public Property value As String
    End Class
    Public Class Condition
        Public Property field As String
        Public Property op As String
        Public Property value As String
        '// Public Property logical_operator As String
    End Class
    Public Class OrderBy
        Public Property field As String
        Public Property direction As String
    End Class

    Sub Tiff_Status_R(ByVal order_num As String)

        Dim dash As Integer
        dash = InStr(UCase(order_num), ".ZI")
        If dash > 0 Then
            order_num = Mid(order_num, 1, dash - 1)
        End If
        dash = InStr(order_num, "-")
        If dash > 0 Then
            order_num = Mid(order_num, 1, dash - 1)
        End If
        dash = InStr(order_num, "_")
        If dash > 0 Then
            order_num = Mid(order_num, 1, dash - 1)
        End If

        'Foxpro
        If foxproConnection = True Then
            Dim status_table As String
            status_table = "t:\database\tiff_status.dbf"
            Dim oConnstring As String = "Provider=VFPOLEDB.1;Data Source= " + status_table

            Dim myconnection As New OleDbConnection(oConnstring)
            Dim CommandText = " UPDATE " & status_table & " SET status = 'R' WHERE order_num = '" & order_num & "'"
            Dim myCommand As New OleDbCommand(CommandText, myconnection)

            Try
                myconnection.Open()
                myCommand.ExecuteNonQuery()
            Catch
                myconnection.Close()
            End Try
            myconnection.Close()
        End If

        'SQL
        If sqlConnection = True Then
            Dim tableValue As String = "OMS.tiff_status"
            Dim updatesValue(1) As Update
            Dim conditionsValue() As Condition = Nothing

            updatesValue(0) = New Update With {
                .field = "status",
                .value = "R"
            }
            conditionsValue = {
                         New Condition With {
                             .field = "order_num",
                            .op = "=",
                            .value = order_num
                                }
                            }


            Try
                Dim jsonString As String = JsonConvert.SerializeObject(New With {
           .updates = updatesValue,
           .table = tableValue,
           .conditions = conditionsValue
       })

                'Dim uri As String = "http://localhost:8009/api/omsupdate?"  'Development
                Dim uri As String = "https://internal.beta-layout.com/oms/api/omsupdate?"
                Dim rawResponse As String

                rawResponse = SendOMSRequest(uri, jsonString)

            Catch
                SendErrorLog("MHT_FILL_AND_PRINT_ERROR", "Tiff_Status_R -SQL REQUEST FAILED")
            End Try
        End If


    End Sub
    Public Function SendRequest(ByVal uri As Uri, ByVal jsonDataBytes As Byte(), ByVal contentType As String, ByVal method As String, ByVal header As WebHeaderCollection) As String
        Try
            Dim req As WebRequest = WebRequest.Create(uri)
            req.Headers = header
            req.ContentType = contentType
            req.Method = method
            If Not (jsonDataBytes Is Nothing) Then req.ContentLength = jsonDataBytes.Length


            Dim stream = req.GetRequestStream()
            If Not (jsonDataBytes Is Nothing) Then stream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
            stream.Close()

            '//
            'Dim response = req.GetResponse().GetResponseStream()

            Dim webResponse As WebResponse = req.GetResponse()
            Dim response = webResponse.GetResponseStream()

            Try
                Dim responseURI As String = UCase(webResponse.ResponseUri.ToString)
                If responseURI.Contains("OMS/API/") Then
                    Dim remaining As String = webResponse.Headers.GetValues("X-RateLimit-Remaining").FirstOrDefault()
                    If remaining < 50 Then
                        wait(10000)
                    End If
                End If
            Catch
            End Try
            '//

            Dim reader As New StreamReader(response)
            Dim res = reader.ReadToEnd()
            reader.Close()
            response.Close()
            Return res
        Catch ex As Exception
            'Throw New Exception(ex.Message)
            Return Nothing
        End Try

    End Function

    Public Function SendOMSRequest(ByVal route As String, ByVal omsData As String)
        If sqlConnection = False Then GoTo noconnection
        Dim result As String

        Try
            Dim jsonString As String = omsData

            Dim Uri As New Uri(String.Format(route))
            Dim data = Encoding.UTF8.GetBytes(jsonString)

            Dim header As New WebHeaderCollection
            header.Add("Authorization", bearer)
            result = SendRequest(Uri, data, "application/json", "POST", header)
            Return result
        Catch ex As Exception
            'MsgBox(ex.Message, "Error")
            Return Nothing
        End Try

noconnection:
        Return Nothing
    End Function
    Public Function SendGETRequest(ByVal uri As String) As DataItem
        If sqlConnection = False Then GoTo noconnection

        Dim request As WebRequest = WebRequest.Create(uri)

        Dim header As New WebHeaderCollection
        header.Add("Authorization", bearer)
        request.Headers = header
        request.Method = "GET"
        request.ContentType = "application/json"

        ' Get the response
        Try
            Using response As WebResponse = request.GetResponse()

                Dim remaining As String = response.Headers.GetValues("X-RateLimit-Remaining").FirstOrDefault()
                If remaining < 50 Then
                    wait(10000)
                End If

                Using reader As New StreamReader(response.GetResponseStream())
                    Dim jsonString As String = reader.ReadToEnd()

                    ' Deserialize JSON string into array of ResponseItem objects
                    Dim dataItems As DataItem = JsonConvert.DeserializeObject(Of DataItem)(jsonString)

                    Return dataItems

                End Using
            End Using
        Catch
            Return Nothing
        End Try

noconnection:
        Return Nothing

    End Function
    Private Class RawTokenCache
        Public Property Token As String
        Public Property ExpiresAt As DateTime
    End Class

    Sub getBearer()
        Dim jwtToken As String
        If Not File.Exists("S:\OMS_DATA\Applications\Access\encrypted_token.bin") Then

            '//If want a new key then run this  and put breakpoint in the GenerateKey() sub
            ' TokenCache.GenerateKey()

            jwtToken = getBearerToken()
            TokenCache.EncryptToken(jwtToken)
        Else

            jwtToken = TokenCache.DecryptToken()
            bearer = "Bearer " + jwtToken

        End If

    End Sub

    Function getBearerToken() As String
        Dim bearerToken As String
        Dim content As JObject
        Try
            ' Dim client As New RestClient("http://localhost:8009/oauth/token?")  ' For development NB Disable saving new bearer key when developing
            Dim client As New RestClient("https://internal.beta-layout.com/oms/oauth/token?")
            Dim request As New RestRequest(Method.POST)
            request.AddHeader("content-type", "application/json")
            request.AddHeader("accept", "application/json")

            request.AddParameter("grant_type", grantType)
            request.AddParameter("client_id", clientID)
            request.AddParameter("client_secret", clientSecret)
            request.AddParameter("scope", clientScope)

            Dim response As IRestResponse = client.Execute(request)

            If response.StatusCode = HttpStatusCode.OK Then

                content = JsonConvert.DeserializeObject(response.Content)


                bearerToken = CStr(content("access_token"))
                bearer = "Bearer " + CStr(content("access_token"))

                Return bearerToken

            Else
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
        End Try

    End Function



    Sub database_note(ByVal order_num As String)

        Dim dash As Integer
        dash = InStr(order_num, "-")
        If dash > 0 Then
            order_num = Mid(order_num, 1, dash - 1)
        End If
        dash = InStr(order_num, "_")
        If dash > 0 Then
            order_num = Mid(order_num, 1, dash - 1)
        End If

        '// Foxpro
        If foxproConnection = True Then
            Dim status_table As String
            status_table = "T:\Database3\tiff_status.dbf"
            Dim oConnstring As String = "Provider=VFPOLEDB.1;Data Source= " + status_table

            Dim myconnection As New OleDbConnection(oConnstring)
            Dim CommandText = " UPDATE " & status_table & " SET note = '1' WHERE order_num = '" & order_num & "'"
            Dim myCommand As New OleDbCommand(CommandText, myconnection)

            Try
                myconnection.Open()
                myCommand.ExecuteNonQuery()
            Catch
                myconnection.Close()
            End Try
            myconnection.Close()
        End If

        '// SQL
        If sqlConnection = True Then
            Dim tableValue As String = "OMS.tiff_status"
            Dim updatesValue(1) As Update
            Dim conditionsValue() As Condition = Nothing

            updatesValue(0) = New Update With {
                .field = "note",
                .value = "1"
            }
            conditionsValue = {
                         New Condition With {
                             .field = "order_num",
                            .op = "=",
                            .value = order_num
                                }
                            }


            Try
                Dim jsonString As String = JsonConvert.SerializeObject(New With {
           .updates = updatesValue,
           .table = tableValue,
           .conditions = conditionsValue
       })


                ' Dim uri As String = "http://localhost:8009/api/omsupdate?"  'Development
                Dim uri As String = "https://internal.beta-layout.com/oms/api/omsupdate?"
                Dim rawResponse As String

                rawResponse = SendOMSRequest(uri, jsonString)
            Catch
                SendErrorLog("MHT_FILL_AND_PRINT_ERROR", "database_note -SQL REQUEST FAILED")
            End Try
        End If

    End Sub


    Sub Create_Note(ByVal au As String, ByVal file1 As String)
        Dim dash As Integer
        Dim commentStr, ref As String

        If foxproConnection = True Then
            If File.Exists("c:\in_house_files\note_file.txt") Then File.Delete("c:\in_house_files\note_file.txt")
            Dim note_file As New FileStream("c:\in_house_files\note_file.txt", FileMode.Append, FileAccess.Write)
            Dim file_out As New StreamWriter(note_file)

            file_out.WriteLine(au)
            file_out.WriteLine("           ")

            dash = InStr(au, "-")
            If dash > 0 Then
                au = Mid(au, 1, dash - 1)
            End If

            FileOpen(57, file1, OpenMode.Input, OpenAccess.Read)
            Dim note_line As String

            Do Until EOF(57)
                note_line = LineInput(57)
                file_out.WriteLine(note_line)
            Loop
            FileClose(57)
            file_out.Close()


            Dim old_note As String
            old_note = au & "_note.txt"

            old_note = notesLocation + "\" & old_note

            If File.Exists(old_note) Then
                Dim linesFromFile1() As String
                Dim linesFromFile2() As String
                Dim combinedLines As New List(Of String)

                linesFromFile1 = System.IO.File.ReadAllLines(old_note)
                linesFromFile2 = System.IO.File.ReadAllLines("c:\in_house_files\note_file.txt")

                For linePos As Integer = 0 To System.Math.Max(linesFromFile1.Length, linesFromFile2.Length) - 1
                    If linePos < linesFromFile1.Length Then combinedLines.Add(linesFromFile1(linePos))
                    If linePos < linesFromFile2.Length Then combinedLines.Add(linesFromFile2(linePos))
                Next

                System.IO.File.WriteAllLines("c:\in_house_files\note_final.txt", combinedLines.ToArray())

            Else

                FileCopy("c:\in_house_files\note_file.txt", "c:\in_house_files\note_final.txt")

            End If

        End If

        ' /////////////////////

        If sqlConnection = True Then
            '// SQL
            Dim jsonnote, action, oper As String
            action = "PRECAM"
            oper = "ARTIFEX"

            Try
                commentStr = File.ReadAllText(file1)
                commentStr = Trim(commentStr.Replace(vbLf, " | "))
                commentStr = Trim(commentStr.Replace(vbCr, " | "))
                commentStr = Trim(commentStr.Replace(vbTab, "  "))
                commentStr = Trim(commentStr.Replace("""", " "))
                If commentStr.Length < 4 Then commentStr = "No additional information given"
            Catch
                commentStr = "Unable to read information given"
            End Try

            '  ref = "143550" ' Testing
            ref = au
            If InStr(ref, "P") > 0 Then
                ref = GetRefFromAu(au)
            End If

            Try
                jsonnote = "{"
                jsonnote = jsonnote & """" & "ref_number" & """" & ":" & """" & ref & """" & ","
                jsonnote = jsonnote & """" & "action" & """" & ":" & """" & action & """" & ","
                jsonnote = jsonnote & """" & "created_by" & """" & ":" & """" & oper & """" & ","
                jsonnote = jsonnote & """" & "comment" & """" & ":" & """" & commentStr & """" & ","
                jsonnote = Trim(jsonnote)
                jsonnote = Mid(jsonnote, 1, jsonnote.Length - 1)
                jsonnote = jsonnote + "}"

                Dim uri As String = "https://internal.beta-layout.com/oms/api/updateorderstatus?"
                'Dim uri As String = "http://localhost:8009/api/updateorderstatus?" 'TESTING

                Dim rawResponse As String

                rawResponse = SendOMSRequest(uri, jsonnote)
                Try
                    Dim jsonObject As JObject = JObject.Parse(rawResponse)
                    If jsonObject.HasValues Then
                        Dim status As Boolean = jsonObject.Property("status")
                        If status = False Then
                            Dim responseError As StreamWriter
                            responseError = File.AppendText("T:\IN_HOUSE_SOFTWARE\ALL_SOFTWARE\Visual_Studio_2010\OMS-Download\TEMP\oms_api_response_failed.txt")
                            responseError.WriteLine("SendOMSRequest Fail for precam note create " + ref)
                            responseError.WriteLine("----------")
                            responseError.Close()
                        End If

                    End If
                Catch ex As Exception
                    Dim responseError As StreamWriter
                    responseError = File.AppendText("T:\IN_HOUSE_SOFTWARE\ALL_SOFTWARE\Visual_Studio_2010\OMS-Download\TEMP\oms_api_response_failed.txt")
                    responseError.WriteLine("SendOMSRequest Fail for precam note create" + ref)
                    responseError.WriteLine("----------")
                    responseError.Close()
                End Try

            Catch
                Dim responseError As StreamWriter
                responseError = File.AppendText("T:\IN_HOUSE_SOFTWARE\ALL_SOFTWARE\Visual_Studio_2010\OMS-Download\TEMP\oms_api_response_failed.txt")
                responseError.WriteLine("SendOMSRequest Fail for precam note create " + ref)
                responseError.WriteLine("----------")
                responseError.Close()
            End Try

        End If

    End Sub


    Public Function GetRefFromAu(ByVal au As String)
        Dim refNumber As String = "NotFound"

        au = au & ".mht"

        Dim response As DataItem
        Try
            Dim jsonString As String = JsonConvert.SerializeObject(New With {
           .fields = {"ref_number"},
           .table = "OMS.tiff_status",
           .conditions = {
                            New Condition With {
                                   .field = "au_num",
                                   .op = "=",
                                   .value = au
                                    }
                                 },
            .order_by = {
                                New OrderBy With {
                                       .field = "id",
                                       .direction = "desc"
                                        }
                                     },
           .all_data = False
       })

            Dim encodedJson As String = HttpUtility.UrlEncode(jsonString)
            ' Dim uri As String = $"http://localhost:8009/api/omsgetdata?req={encodedJson}" 'Development
            Dim uri As String = $"https://internal.beta-layout.com/oms/api/omsgetdata?req={encodedJson}"
            response = SendGETRequest(uri)
        Catch
            response = Nothing
        End Try

        Try
            Dim items = response.rows.ToArray
            For Each item In items
                Dim jsonObject As JObject = DirectCast(item, JObject)
                refNumber = jsonObject("ref_number")
                Exit For
            Next
        Catch
            refNumber = "NotFound"
        End Try

        Return refNumber
    End Function

    Sub Notes()
        Dim Input_Dir As String
        Input_Dir = inputDirectory ' "C:\Source"

        ' ***************  changes to be made here ( group all individual txt files into one ) **

        Dim path1 As String = Input_Dir
        Dim files1 As String() = Directory.GetFiles(path1)
        Dim file1 As String

        Dim strFileName As String = ""
        Dim au As String
        Dim dash As Integer

        For Each file1 In files1
            If (InStr(UCase(file1), "_NOTE") > 0 And InStr(UCase(file1), ".TXT") > 0) Then
                strFileName = file1

                dash = InStr(UCase(file1), "_NOTE")
                au = Mid(file1, 1, dash - 1)
                dash = au.LastIndexOf("\")
                au = Mid(au, dash + 2)

                Create_Note(au, file1)

                dash = InStr(au, "-")
                If dash > 0 Then
                    au = Mid(au, 1, dash - 1)
                End If

                If foxproConnection = True Then
                    Dim old_note As String
                    old_note = au & "_note.txt"
                    old_note = notesLocation + "\" & old_note

                    If File.Exists(old_note) Then Kill(old_note)
                    File.Move("c:\in_house_files\note_final.txt", old_note)
                End If

                database_note(au)

                End If
        Next
    End Sub

    Sub UNZIP()
        Dim Input_Dir As String
        Input_Dir = inputDirectory ' "C:\Source"

        ' ***************  changes to be made here ( group all individual txt files into one ) **

        Dim path1 As String = Input_Dir
        Dim files1 As String() = Directory.GetFiles(path1)
        Dim file1 As String

        Dim strFileName As String = ""

        For Each file1 In files1
            If (InStr(UCase(file1), ".ZIP") > 0 And InStr(UCase(file1), "GERB") = 0) Then
                strFileName = file1
                Try
                    Dim p As New Process
                    With p
                        .StartInfo.FileName = "C:\7Zip\7z.exe"
                        .StartInfo.Arguments = " e " & strFileName & " -o" & path1 & " -aoa"
                        .Start()
                    End With
                Catch
                    SendErrorLog("MHT_FILL_AND_PRINT_ERROR", "UNZIP - Problem with " + file1)
                    MsgBox("Problem with " & file1 & " :-(")
                End Try
            End If
        Next
    End Sub

    Sub get_orders(ByVal source As String)

        Dim au As String
        Dim dot As Integer

        Dim di3 As DirectoryInfo = New DirectoryInfo(source) ' Original Source directory

        Dim files3 As Object

        files3 = di3.GetFileSystemInfos()
        For Each file In files3
            If InStr(UCase(file.name), ".ZIP") > 0 Then
                dot = InStr(UCase(file.name), ".ZIP")
                au = Mid(file.name, 1, dot - 1)
                myarraylist.add(au)
                new_orders = True
            End If
        Next

    End Sub

    Sub check_for_orders(ByVal source As String)
        Dim au As String
        Dim dot As Integer
        Dim onhold As Boolean
        Dim anAssembly As Boolean

        Dim onholds As Boolean
        onholds = False

        Dim di3 As DirectoryInfo = New DirectoryInfo(source) ' Original Source directory

        Dim files3 As Object
        Dim c As Integer
        c = 0
        files3 = di3.GetFileSystemInfos()
        For Each order In files3

            If c < 10 Then
                If InStr(UCase(order.name), ".ZIP") > 0 Then
                    ' c = c + 1
                    dot = InStr(UCase(order.name), ".ZIP")
                    au = Mid(order.name, 1, dot - 1)


                    If Not (myarraylist.contains(au)) Then
                        myarraylist.add(au)
                        GoTo not_this_time
                    End If

                    If InStr(UCase(order.name), ".FILEPART") > 0 Then
                        If (filepartlist.contains(au)) Then
                            Try
                                SendErrorLog("MHT_FILL_AND_PRINT_ERROR", "check_for_orders - PrecamMoveFail on order " + au)
                            Catch
                            End Try
                        Else
                            filepartlist.add(au)
                        End If
                        GoTo not_this_time
                    End If

                    new_orders = new_orders + 1
                    c = c + 1

                    '************************************

                    'seperate Assembly Orders
                    'anAssembly = False
                    'anAssembly = AmIAssembly(au)
                    'If anAssembly = True Then
                    ' File.Copy(order.fullname, "s: \job\Dinesh_Assembly\" & order.name, True)
                    'End If

                    'seperate_on_holds
                    onhold = False
                    onhold = AmIOnHold(au)

                    'If onhold = False Then
                    ' File.Move(order.fullname, "c:\source\" & order.name)
                    ' Console.WriteLine(au & "   " & CStr(Now.Hour) & ":" & CStr(Now.Minute))
                    ' Else
                    '     File.Move(order.fullname, "t:\off_holds\" & order.name)
                    '     onholds = True
                    '     Tiff_Status_R(order.name)
                    '     Console.WriteLine(au & "   " & CStr(Now.Hour) & ":" & CStr(Now.Minute) & "--- OFF_HOLD")
                    '     new_orders = new_orders - 1
                    ' End If

                    'File.Move(order.fullname, "c:\source\" & order.name)
                    File.Move(order.fullname, inputDirectory + "\" & order.name)
                    If onhold = True Then
                        Console.WriteLine(au & "   " & CStr(Now.Hour) & ":" & CStr(Now.Minute) & "--- OFF_HOLD")
                    Else
                        Console.WriteLine(au & "   " & CStr(Now.Hour) & ":" & CStr(Now.Minute))
                    End If
skip_me:

                End If
            End If

not_this_time:

            If (InStr(UCase(order.name), ".TXT") > 0) Then
                File.Move(order.fullname, statusUpdates + "\" & order.name)  ' They are then handles in here by FolderWatcher 
                Console.WriteLine(order.name & " -- Moved")
            End If

        Next

    End Sub


    Function AmIOnHold(ByVal order_num As String)
        Dim onhold As Boolean = False
        Dim dash As Integer
        Dim au As String

        dash = order_num.IndexOf("-")
        If dash > 0 Then
            au = Mid(order_num, 1, dash)
        Else
            au = order_num
        End If

        ''''''''''''''''''
        'FoxPro
        If useFoxproConnection = True Then
            Dim status_table As String
            status_table = "T:\Database3\tiff_status.dbf"
            Dim oConnstring As String = "Provider=VFPOLEDB.1;Data Source= " + status_table
            Dim myconnection As New OleDbConnection(oConnstring)
            Dim commandtext As String = "Select job_num From " & status_table & " where order_num = '" & au & "' Order by time_stamp desc"
            Dim myCommand As New OleDbCommand(commandtext, myconnection)
            Dim auhold As String

            Try
                myconnection.Open()
                Dim myReader As OleDbDataReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    auhold = Trim(myReader.GetString(0))
                    If InStr(auhold, "HOLD") Then onhold = True
                End While
            Catch
                myconnection.Close()
            End Try
            myconnection.Close()
        End If


        If useSqlConnection = True Then
            'SQL
            Dim auhold As String
            Dim response As DataItem
            Try
                Dim jsonString As String = JsonConvert.SerializeObject(New With {
           .fields = {"job_num"},
           .table = "OMS.tiff_status",
           .conditions = {
                            New Condition With {
                                   .field = "order_num",
                                   .op = "=",
                                   .value = au
                                    }
                                 },
            .order_by = {
                                New OrderBy With {
                                       .field = "time_stamp",
                                       .direction = "desc"
                                        }
                                     },
           .all_data = False
       })

                Dim encodedJson As String = HttpUtility.UrlEncode(jsonString)
                ' Dim uri As String = $"http://localhost:8009/api/omsgetdata?req={encodedJson}" 'Development
                Dim uri As String = $"https://internal.beta-layout.com/oms/api/omsgetdata?req={encodedJson}"
                response = SendGETRequest(uri)
            Catch
                response = Nothing
            End Try

            Try
                Dim items = response.rows.ToArray
                For Each item In items
                    Dim jsonObject As JObject = DirectCast(item, JObject)
                    auhold = jsonObject("job_num")
                    If InStr(auhold, "HOLD") Then onhold = True
                Next
            Catch

            End Try

        End If

        Return onhold
    End Function


    Function AmIAssembly(ByVal order_num As String)
        Dim anAssembly As Boolean = False
        Dim dash As Integer
        Dim au As String

        dash = order_num.IndexOf("-")
        If dash > 0 Then
            au = Mid(order_num, 1, dash)
        Else
            au = order_num
        End If

        ''''''''''''''''''

        '//Foxpro
        If useFoxproConnection = True Then
            Dim status_table As String
            status_table = "T:\Database3\orders_info.dbf"
            Dim oConnstring As String = "Provider=VFPOLEDB.1;Data Source= " + status_table
            Dim myconnection As New OleDbConnection(oConnstring)
            Dim commandtext As String = "Select order_ass From " & status_table & " where order_num = '" & au & "' Order by time_stamp desc"
            Dim myCommand As New OleDbCommand(commandtext, myconnection)
            Dim auAss As String

            Try
                myconnection.Open()
                Dim myReader As OleDbDataReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    auAss = Trim(myReader.GetString(0))
                    If InStr(auAss, "YE") Then anAssembly = True
                End While
            Catch
                myconnection.Close()
            End Try
            myconnection.Close()
        End If


        If useSqlConnection = True Then
            'SQL
            Dim auAss As String
            Dim response As DataItem
            Try
                Dim jsonString As String = JsonConvert.SerializeObject(New With {
           .fields = {"order_ass"},
           .table = "OMS.orders_info",
           .conditions = {
                            New Condition With {
                                   .field = "order_num",
                                   .op = "=",
                                   .value = au
                                    }
                                 },
            .order_by = {
                                New OrderBy With {
                                       .field = "time_stamp",
                                       .direction = "desc"
                                        }
                                     },
           .all_data = False
       })

                Dim encodedJson As String = HttpUtility.UrlEncode(jsonString)
                ' Dim uri As String = $"http://localhost:8009/api/omsgetdata?req={encodedJson}" ' Development
                Dim uri As String = $"https://internal.beta-layout.com/oms/api/omsgetdata?req={encodedJson}"
                response = SendGETRequest(uri)
            Catch
                response = Nothing
            End Try

            Try
                Dim items = response.rows.ToArray
                For Each item In items
                    Dim jsonObject As JObject = DirectCast(item, JObject)
                    auAss = jsonObject("order_ass")
                    If InStr(auAss, "YE") Then anAssembly = True
                Next
            Catch

            End Try

        End If
noSQL:

        Return anAssembly
    End Function

    '**************************

    Sub update_database()

        Dim Input_Dir As String
        Input_Dir = inputDirectory ' "c:\source"


        'UPLOADING
        If File.Exists("S:\OMS_DATA\applications\mht_fil_auto\upload_also.txt") Then
            Dim name As String
            Dim dash As Integer
            Dim files3 As String() = Directory.GetFiles(Input_Dir)
            Dim file3 As String
            For Each file3 In files3
                If InStr(UCase(file3), "ERR") = 0 Then
                    If InStr(UCase(file3), ".ZIP") > 0 Or InStr(UCase(file3), ".PDF") > 0 Then
                        dash = file3.LastIndexOf("\")
                        name = Mid(file3, dash + 2)
                        name = archiveUploadFilesTemp + "\" & name
                        Try
                            FileCopy(file3, name)
                        Catch
                        End Try
                    End If
                End If
            Next
        End If


        ' ******************************************

        Dim tot As Integer = 0
        Dim order_counter As Integer
        Dim cam_name As String

        order_counter = 0

        wait(2000)
        num_small_pcb = 0
        conrad_orders = 0



        If File.Exists("c:\in_house_files\temp.txt") Then File.Delete("c:\in_house_files\temp.txt") 'For DAN  
        If File.Exists("c:\in_house_files\temp2.txt") Then File.Delete("c:\in_house_files\temp2.txt") 'For Faxback
        If File.Exists("c:\in_house_files\temp3.txt") Then File.Delete("c:\in_house_files\temp3.txt") 'For NAMES.TXT


        Dim AppendFileContent As String
        Dim AppendFileContent2 As String
        Dim AppendFileContent3 As String
        Dim AppendFileContent4 As String
        Dim AppendFileContent5 As String
        Dim AppendFileContent1 As String

        ' ***************  changes to be made here ( group all individual txt files into one ) **

        Dim path1 As String = Input_Dir
        Dim files1 As String() = Directory.GetFiles(path1)
        Dim file1 As String
        Dim file2 As String
        Dim dot_posit As Integer
        Dim slash_posit As Integer
        Dim order_num As String
        Dim order_danned As Boolean

        For Each file1 In files1

            If InStr(file1, "onhold") > 0 And InStr(file1, ".xl") = 0 Then
                'Try
                put_on_hold(file1)
                ' Catch
                ' End Try
            End If
        Next

        ' ------------------------------------------

        For Each file1 In files1
            If InStr(file1, ".gwk") Or InStr(file1, ".GWK") Then
                ' -----  Get Order Number
                order_num = LCase(file1)
                dot_posit = order_num.IndexOf(".")
                order_num = Mid(order_num, 1, dot_posit)
slash_again:
                slash_posit = order_num.IndexOf("\")
                If slash_posit < 0 Then
                    order_num = Mid(order_num, slash_posit + 2)
                Else
                    order_num = Mid(order_num, slash_posit + 2)
                    GoTo slash_again
                End If
                order_num = Trim(order_num)
                order_num = UCase(order_num)

                ' -------  Does text file for this order exist  ****************
                ' -------  If yes add it to a temp txt file     ****************
                file2 = path1 + "\" + order_num + ".txt"
                If File.Exists(file2) Then

                    order_danned = False

                    Dim in_file As New FileStream(file2, FileMode.Open, FileAccess.Read)
                    Dim file_in As New StreamReader(in_file)
                    AppendFileContent = file_in.ReadLine()
                    AppendFileContent = RTrim(AppendFileContent)
                    AppendFileContent2 = AppendFileContent

                    AppendFileContent3 = " "

                    '  ***** ADD ORDERS WITHOUT "-" to names.txt file *******

                    If Not InStr(file2, "-") Then
                        d1 = Trim(d1.ToString)
                        Dim dd, mm, yyyy, tttt, date_edited, date_printed As String
                        Dim cam_name_divide As String
                        dd = Mid(d1, 1, 2) & "-"
                        mm = Mid(d1, 4, 2) & "-"
                        yyyy = Mid(d1, 7, 4)
                        tttt = "   00:00:00"
                        date_printed = mm & dd & yyyy & tttt

                        AppendFileContent4 = file_in.ReadLine()  ' CAM OPERATORS NAME & DATE_EDITED

                        If Len(AppendFileContent4) < 20 Then
                            AppendFileContent4 = AppendFileContent4 & "      " & date_printed
                        End If

                        AppendFileContent4 = RTrim(AppendFileContent4)
                        cam_name_divide = InStr(AppendFileContent4, "  ")
                        cam_name = Mid(AppendFileContent4, 1, cam_name_divide + 1)
                        date_edited = Mid(AppendFileContent4, cam_name_divide + 1)
                        cam_name = Trim(cam_name)
                        date_edited = Trim(date_edited)
                        AppendFileContent4 = cam_name
                        AppendFileContent5 = AppendFileContent2

                        REM ****************************************

                        Dim spacer As Integer
                        spacer = InStr(AppendFileContent5, "  ")
                        AppendFileContent5 = Mid(AppendFileContent5, 1, spacer + 1)

                        REM ***********************************************

                        AppendFileContent5 = Trim(AppendFileContent5)

                        AppendFileContent5 = AppendFileContent5 & ".MHT"

                        AppendFileContent1 = AppendFileContent2
                        Do While AppendFileContent1 <> AppendFileContent3
                            AppendFileContent3 = AppendFileContent1
                            AppendFileContent1 = AppendFileContent1.Replace("    ", "   ")
                        Loop
                        AppendFileContent1 = AppendFileContent1.Replace("   ", " * ")
                        AppendFileContent1 = AppendFileContent1 + " * "

                        Dim star_loc As Integer
                        Dim counter As Integer

                        counter = 0
                        Do While counter <> 16
                            star_loc = InStr(AppendFileContent1, "*")
                            AppendFileContent1 = Mid(AppendFileContent1, star_loc + 1)
                            counter = counter + 1
                        Loop
                        star_loc = InStr(AppendFileContent1, "*")
                        AppendFileContent1 = Mid(AppendFileContent1, 1, star_loc - 1)
                        AppendFileContent1 = UCase(Trim(AppendFileContent1))
                        If AppendFileContent1 = "1" Then AppendFileContent1 = "True"
                        If AppendFileContent1 = "0" Then AppendFileContent1 = "False"

                        AppendFileContent4 = AppendFileContent4 & "   " & AppendFileContent5 & "    " & date_edited & "      " & AppendFileContent1


                        ' ********** put pool types of each order into an array **********


                        Dim txt_pool As String = ""

                        series_flag = False

                        ' the order number in already got
                        ' Get the pool type

                        txt_pool = AppendFileContent3
                        txt_pool = AppendFileContent3.Replace("   ", " * ")
                        txt_pool = txt_pool + " * "
                        counter = 0
                        Do While counter <> 11
                            star_loc = InStr(txt_pool, "*")
                            txt_pool = Mid(txt_pool, star_loc + 1)
                            counter = counter + 1
                        Loop
                        star_loc = InStr(txt_pool, "*")
                        txt_pool = Mid(txt_pool, 1, star_loc - 1)
                        txt_pool = Trim(txt_pool)

                        If CInt(txt_pool) > 809 And CInt(txt_pool) < 851 Then series_flag = True ' SS Series
                        If CInt(txt_pool) > 9 And CInt(txt_pool) < 51 Then series_flag = True ' DS Series
                        If CInt(txt_pool) = 7 Then series_flag = True ' 4 layer series10

                        pool_array(order_counter, 0) = order_num
                        pool_array(order_counter, 1) = series_flag
                        order_counter = order_counter + 1

                        ' ******************************************************************

                    End If


                    ' *****************************************************************************************************************

                    Do While AppendFileContent2 <> AppendFileContent3
                        AppendFileContent3 = AppendFileContent2
                        AppendFileContent2 = AppendFileContent2.Replace("    ", "   ")
                    Loop
                    AppendFileContent2 = AppendFileContent2.Replace("   ", " * ")
                    AppendFileContent2 = AppendFileContent2 + " * "
                    file_in.Close()

                    Dim out_file As New FileStream("c:\in_house_files\temp.txt", FileMode.Append, FileAccess.Write)  ' DAN
                    Dim file_out As New StreamWriter(out_file)
                    Dim out_file2 As New FileStream("c:\in_house_files\temp2.txt", FileMode.Append, FileAccess.Write)  ' FAXBACK
                    Dim file_out2 As New StreamWriter(out_file2)
                    Dim out_file3 As New FileStream("c:\in_house_files\temp3.txt", FileMode.Append, FileAccess.Write)  ' NAMES
                    Dim file_out3 As New StreamWriter(out_file3)


                    REM *************************  Formating Output To Faxback_info.txt *******************************

                    Dim star_position As Integer
                    Dim num_stars As Integer
                    Dim temp101 As String

                    star_position = 1
                    num_stars = 0
                    temp101 = AppendFileContent2
                    Do While star_position > 0
                        star_position = InStr(temp101, "*")
                        temp101 = Mid(temp101, star_position + 1, Len(temp101))
                        num_stars = num_stars + 1
                    Loop

                    If num_stars = 22 Then
                        AppendFileContent2 = AppendFileContent2 & " 0 *"
                    End If

                    If num_stars = 21 Then
                        AppendFileContent2 = AppendFileContent2 & " U_NKNOWN * 0 *"
                    End If

                    REM *************************************************************************************

                    file_out.WriteLine(AppendFileContent)
                    file_out.Close()
                    file_out2.WriteLine(AppendFileContent2)
                    file_out2.Close()

                    ' **** Update faxback database ****************************
                    Try
                        wait(500)
                        update_faxback(AppendFileContent2, cam_name)
                    Catch
                    End Try

                    ' *************************************************

                    If Not InStr(file2, "-") Then file_out3.WriteLine(AppendFileContent4)
                    'file_out3.WriteLine(AppendFileContent4)
                    file_out3.Close()

                    ' **** Update names database ****************************
                    Try
                        wait(500)

                        update_names(AppendFileContent4)

                    Catch
                    End Try
                    ' *************************************************

                End If

                If order_danned = False Then

                    order_danned = True

                    Dim file31 As String
                    Dim file32 As String
                    Dim file33 As String
                    Dim ext31 As Integer
                    'ext31 = 1
                    ext31 = 97
                    Do While ext31 < 123
                        file31 = path1 & "\" & order_num & "_" & Chr(ext31) & ".txt"
                        'file32 = path1 & "\" & order_num & "-" & ext31 & ".gwk"
                        'file33 = UCase(file32)
                        ' If File.Exists(file31) Then
                        If File.Exists(file31) And (Not (donelist.contains(file31))) Then
                            donelist.add(file31)
                            Dim in_file As New FileStream(file31, FileMode.Open, FileAccess.Read)
                            Dim file_in As New StreamReader(in_file)
                            'AppendFileContent = file_in.ReadToEnd()
                            AppendFileContent = file_in.ReadLine()
                            AppendFileContent = RTrim(AppendFileContent)
                            AppendFileContent2 = AppendFileContent
                            AppendFileContent3 = " "

                            Do While AppendFileContent2 <> AppendFileContent3
                                AppendFileContent3 = AppendFileContent2
                                AppendFileContent2 = AppendFileContent2.Replace("    ", "   ")
                            Loop
                            AppendFileContent2 = AppendFileContent2.Replace("   ", " * ")
                            AppendFileContent2 = AppendFileContent2 + " * "
                            file_in.Close()

                            Dim out_file As New FileStream("c:\in_house_files\temp.txt", FileMode.Append, FileAccess.Write)
                            Dim file_out As New StreamWriter(out_file)
                            Dim out_file2 As New FileStream("c:\in_house_files\temp2.txt", FileMode.Append, FileAccess.Write)
                            Dim file_out2 As New StreamWriter(out_file2)
                            'file_out.WriteLine("")
                            file_out.WriteLine(AppendFileContent)
                            file_out.Close()
                            file_out2.WriteLine(AppendFileContent2)
                            file_out2.Close()


                            ' **** Update faxback database ****************************
                            Try
                                wait(500)
                                update_faxback(AppendFileContent2, cam_name)
                            Catch
                            End Try

                            ' *************************************************

                            If File.Exists(file32) Or File.Exists(file33) Then
                                Dim out_file3 As New FileStream("c:\in_house_files\temp3.txt", FileMode.Append, FileAccess.Write)
                                Dim file_out3 As New StreamWriter(out_file3)
                                file_out3.WriteLine(AppendFileContent4)
                                file_out3.Close()
                                ' **** Update names database ****************************
                                Try
                                    wait(500)

                                    update_names(AppendFileContent4)

                                Catch
                                End Try
                                ' *************************************************
                            End If
                        End If

                        ext31 = ext31 + 1
                    Loop


                    ext31 = 1
                    Do While ext31 < 100
                        file31 = path1 & "\" & order_num & "-" & ext31 & ".txt"
                        file32 = path1 & "\" & order_num & "-" & ext31 & ".gwk"
                        file33 = UCase(file32)
                        If File.Exists(file31) And (Not (donelist.contains(file31))) Then
                            donelist.add(file31)
                            Dim in_file As New FileStream(file31, FileMode.Open, FileAccess.Read)
                            Dim file_in As New StreamReader(in_file)
                            'AppendFileContent = file_in.ReadToEnd()
                            AppendFileContent = file_in.ReadLine()
                            AppendFileContent = RTrim(AppendFileContent)
                            AppendFileContent2 = AppendFileContent
                            AppendFileContent3 = " "

                            Do While AppendFileContent2 <> AppendFileContent3
                                AppendFileContent3 = AppendFileContent2
                                AppendFileContent2 = AppendFileContent2.Replace("    ", "   ")
                            Loop
                            AppendFileContent2 = AppendFileContent2.Replace("   ", " * ")
                            AppendFileContent2 = AppendFileContent2 + " * "
                            file_in.Close()

                            Dim out_file As New FileStream("c:\in_house_files\temp.txt", FileMode.Append, FileAccess.Write)
                            Dim file_out As New StreamWriter(out_file)
                            Dim out_file2 As New FileStream("c:\in_house_files\temp2.txt", FileMode.Append, FileAccess.Write)
                            Dim file_out2 As New StreamWriter(out_file2)
                            'file_out.WriteLine("")
                            file_out.WriteLine(AppendFileContent)
                            file_out.Close()
                            file_out2.WriteLine(AppendFileContent2)
                            file_out2.Close()


                            ' **** Update faxback database ****************************
                            Try
                                wait(500)
                                update_faxback(AppendFileContent2, cam_name)
                            Catch
                            End Try

                            ' *************************************************

                            If File.Exists(file32) Or File.Exists(file33) Then
                                Dim out_file3 As New FileStream("c:\in_house_files\temp3.txt", FileMode.Append, FileAccess.Write)
                                Dim file_out3 As New StreamWriter(out_file3)
                                file_out3.WriteLine(AppendFileContent4)
                                file_out3.Close()
                                ' **** Update names database ****************************
                                Try
                                    wait(500)

                                    update_names(AppendFileContent4)

                                Catch
                                End Try
                                ' *************************************************
                            End If
                        End If

                        ext31 = ext31 + 1
                    Loop

                End If



                ' ***************************************************************


            End If
        Next

        ' *****  Append temp.txt file above to DAN.TXT


        '/////////////////////////////////////////////////

        'Directory.CreateDirectory("s:\oms_data\david") ' TODO Remove if all 3 below not needed 

        'If Not (File.Exists("c:\in_house_files\temp.txt")) Then GoTo leave

        'Dim fr As New FileStream("c:\in_house_files\temp.txt", FileMode.Open, FileAccess.Read)
        'Dim r As New StreamReader(fr)
        'AppendFileContent = r.ReadToEnd()
        'r.Close()

        'Dim fs As New FileStream("s:\OMS_data\david\dan.txt", FileMode.Append, FileAccess.Write)
        'Dim s As New StreamWriter(fs)
        's.WriteLine(AppendFileContent)
        's.Close()

        'If Not (File.Exists("c:\in_house_files\temp2.txt")) Then GoTo leave

        'Dim fr2 As New FileStream("c:\in_house_files\temp2.txt", FileMode.Open, FileAccess.Read)
        'Dim r2 As New StreamReader(fr2)
        'AppendFileContent2 = r2.ReadToEnd()
        'r2.Close()

        'Dim fs2 As New FileStream("s:\oms_data\david\faxback_info.txt", FileMode.Append, FileAccess.Write)
        'Dim s2 As New StreamWriter(fs2)
        's2.WriteLine(AppendFileContent2)
        's2.Close()


        'If Not (File.Exists("c:\in_house_files\temp3.txt")) Then GoTo leave

        'Dim fr3 As New FileStream("c:\in_house_files\temp3.txt", FileMode.Open, FileAccess.Read)
        'Dim r3 As New StreamReader(fr3)
        'AppendFileContent4 = r3.ReadToEnd()
        'r3.Close()

        'Dim fs3 As New FileStream("s:\oms_data\david\names.txt", FileMode.Append, FileAccess.Write)
        'Dim s3 As New StreamWriter(fs3)
        's3.WriteLine(AppendFileContent4)
        's3.Close()

        '///////////////////////////////////////


already_danned:

        Try
            File.Delete("c:\in_house_files\temp.txt")
        Catch
        End Try

        GoTo leave1


leave:
        SendErrorLog("MHT_FILL_AND_PRINT_ERROR", "update_databae - No text files in the input directory !!")
        MsgBox("No text files in the input directory !!")
leave1:


    End Sub

    ' ---------------------

    Sub put_on_hold(ByVal file1 As String)
        Dim id1 As Integer
        Dim names_database As String
        Dim names_database2 As String
        Dim ord As String
        Dim dash As Integer

        ord = UCase(file1)
        dash = InStr(ord, ".TX")
        ord = Mid(ord, 1, dash - 1)
        dash = InStr(ord, "ONHOLD")
        ord = Mid(ord, dash - 10)
        dash = InStr(ord, "\")
        ord = Mid(ord, dash + 1)

        'FoxPro
        If foxproConnection = True Then
            names_database = "s: \job\in_house_software\on_hold_home.exe " & ord
            names_database2 = "s:\job\in_house_software\on_hold_home2.exe " & ord

            id1 = Shell(names_database, 1)
            wait(100)

            id1 = Shell(names_database2, 1)
            wait(100)
        End If

        'SQL
        If sqlConnection = True Then

            Try
                Dim star As Integer = InStr(ord, "_")
                Dim order_num As String = Mid(ord, 1, star - 1)
                ord = Mid(ord, star + 8)
                Dim name As String = Trim(ord)

                Dim tableValue As String = "OMS.tiff_status"
                Dim updatesValue(4) As Update
                Dim conditionsValue() As Condition = Nothing

                updatesValue(0) = New Update With {
                    .field = "was_on_hold",
                    .value = "1"
                }
                updatesValue(1) = New Update With {
                    .field = "job_num",
                    .value = "ON_HOLD"
                }
                updatesValue(2) = New Update With {
                    .field = "precammed",
                    .value = "0"
                }
                updatesValue(3) = New Update With {
                    .field = "pre_cam_op",
                    .value = name
                }
                conditionsValue = {
                             New Condition With {
                                 .field = "order_num",
                                .op = "=",
                                .value = order_num
                                    }
                                }

                Dim jsonString As String = JsonConvert.SerializeObject(New With {
           .updates = updatesValue,
           .table = tableValue,
           .conditions = conditionsValue
       })


                'Dim uri As String = "http://localhost:8009/api/omsupdate?"  'Development
                Dim uri As String = "https://internal.beta-layout.com/oms/api/omsupdate?"
                Dim rawResponse As String

                rawResponse = SendOMSRequest(uri, jsonString)
            Catch
                SendErrorLog("MHT_FILL_AND_PRINT_ERROR", "put_on_hold -SQL REQUEST FAILED")
            End Try

        End If

    End Sub



    Private Sub update_faxback(ByVal f_info As String, ByVal c_info As String)
        wait(500)
        Dim id1 As Integer
        Dim fax_database As String
        Dim ord As String
        ord = f_info
        Dim data As String
        Dim tiff_num As String

        id1 = InStr(ord, "*")
        fax_database = Trim(Mid(ord, 1, id1 - 1))
        tiff_num = fax_database
        fax_database = "'" & fax_database & "',"
        ord = Mid(ord, id1 + 1)

        Do While ord.Length > 1
            id1 = InStr(ord, "*")
            data = Trim(Mid(ord, 1, id1 - 1))
            data = "'" & data & "',"
            fax_database = fax_database & data
            ord = Mid(ord, id1 + 1)
        Loop

        fax_database = Mid(fax_database, 1, Len(fax_database) - 1)

        Dim to_dan As String
        to_dan = fax_database


        If foxproConnection = True Then
            'FoxPro
            Dim fox_fax_database As String = "s:\job\in_house_software\faxback_insert_finish.exe " & fax_database
            id1 = Shell(fox_fax_database, 1)
            wait(200)
        End If
        If sqlConnection = True Then
            'SQL
            Try
                Dim dash As Integer
                Dim stamp As Date = Now
                Dim timeStamp As Long = CStr(stamp.Ticks)
                Dim order As String = fax_database

                Dim cnt As Integer = 0
                For Each c As Char In fax_database
                    If c = "," Then cnt += 1
                Next

check_again:
                If cnt < 24 Then
                    fax_database = fax_database & ",'0'"
                    cnt = cnt + 1
                End If
                If cnt < 24 Then GoTo check_again

                dash = InStr(order, ",")
                order = Mid(order, 2, dash - 3)

                dash = InStr(order, "-")
                If dash > 0 Then
                    order = Mid(order, 1, dash - 1)
                End If
                dash = InStr(order, "_")
                If dash > 0 Then
                    order = Mid(order, 1, dash - 1)
                End If
                order = "'" & order & "',"

                fax_database = order & fax_database & "," & timeStamp

                fax_database = fax_database.Replace(".T.", "1")
                fax_database = fax_database.Replace(".F.", "0")
                fax_database = fax_database.Replace(",_", "_")
                fax_database = fax_database.Replace("'", "")
                Dim faxDatabase As String() = fax_database.Split(",")

                Dim jsonFax As String = "{"
                jsonFax = "{"
                jsonFax = jsonFax & """" & "time_stamp" & """" & ":" & CLng(timeStamp) & ","
                jsonFax = jsonFax & """" & "order_num" & """" & ":" & """" & UCase(faxDatabase(0)) & """" & ","
                jsonFax = jsonFax & """" & "au_num" & """" & ":" & """" & UCase(faxDatabase(1)) & """" & ","
                jsonFax = jsonFax & """" & "x_size" & """" & ":" & """" & faxDatabase(2) & """" & ","
                jsonFax = jsonFax & """" & "y_size" & """" & ":" & """" & faxDatabase(3) & """" & ","
                jsonFax = jsonFax & """" & "qty" & """" & ":" & """" & faxDatabase(4) & """" & ","
                jsonFax = jsonFax & """" & "mask" & """" & ":" & """" & faxDatabase(5) & """" & ","
                jsonFax = jsonFax & """" & "silk" & """" & ":" & """" & faxDatabase(6) & """" & ","
                jsonFax = jsonFax & """" & "ger" & """" & ":" & """" & faxDatabase(7) & """" & ","
                jsonFax = jsonFax & """" & "cam" & """" & ":" & """" & faxDatabase(8) & """" & ","
                jsonFax = jsonFax & """" & "rout1" & """" & ":" & """" & faxDatabase(9) & """" & ","
                jsonFax = jsonFax & """" & "odel" & """" & ":" & """" & faxDatabase(10) & """" & ","
                jsonFax = jsonFax & """" & "rout2" & """" & ":" & """" & faxDatabase(11) & """" & ","
                jsonFax = jsonFax & """" & "pool" & """" & ":" & """" & faxDatabase(12) & """" & ","
                jsonFax = jsonFax & """" & "nn" & """" & ":" & """" & faxDatabase(13) & """" & ","
                jsonFax = jsonFax & """" & "score" & """" & ":" & """" & faxDatabase(14) & """" & ","
                jsonFax = jsonFax & """" & "t_gaps_spec" & """" & ":" & """" & faxDatabase(15) & """" & ","
                jsonFax = jsonFax & """" & "holes_spec" & """" & ":" & """" & faxDatabase(16) & """" & ","
                jsonFax = jsonFax & """" & "ul_logo" & """" & ":" & """" & faxDatabase(17) & """" & ","
                jsonFax = jsonFax & """" & "stencil" & """" & ":" & """" & faxDatabase(18) & """" & ","
                jsonFax = jsonFax & """" & "top_cu" & """" & ":" & """" & faxDatabase(19) & """" & ","
                jsonFax = jsonFax & """" & "bot_cu" & """" & ":" & """" & faxDatabase(20) & """" & ","
                jsonFax = jsonFax & """" & "precam" & """" & ":" & """" & faxDatabase(21) & """" & ","
                jsonFax = jsonFax & """" & "rfid_code" & """" & ":" & """" & faxDatabase(22) & """" & ","
                jsonFax = jsonFax & """" & "holes_cct" & """" & ":" & """" & faxDatabase(23) & """" & ","
                jsonFax = jsonFax & """" & "counter" & """" & ":" & """" & faxDatabase(24) & """" & ","
                jsonFax = jsonFax & """" & "depth" & """" & ":" & """" & fax_database(25) & """" & ","

                jsonFax = Trim(jsonFax)
                jsonFax = Mid(jsonFax, 1, jsonFax.Length - 1)
                jsonFax = jsonFax + "}"

                '  Dim uri As String = "http://localhost:8009/api/omscreatefaxback?"  'Development
                Dim uri As String = "https://internal.beta-layout.com/oms/api/omscreatefaxback?"
                Dim rawResponse As String

                rawResponse = SendOMSRequest(uri, jsonFax)
                Try
                    Dim jsonObject As JObject = JObject.Parse(rawResponse)
                    If jsonObject.HasValues AndAlso jsonObject.Property("msg") IsNot Nothing Then
                        Dim msg As String = jsonObject("msg").ToString()
                        If InStr(msg, "SUCCESS") = 0 Then
                            SendErrorLog("MHT_FILL_AND_PRINT_ERROR", "update_faxback - " + fax_database(1))
                        End If
                    End If
                Catch ex As Exception
                    SendErrorLog("MHT_FILL_AND_PRINT_ERROR", "update_faxback - SQL REQUEST FAILED " + fax_database(1))
                End Try

            Catch
                SendErrorLog("MHT_FILL_AND_PRINT_ERROR", "update_faxback - SQL REQUEST FAILED " + fax_database(1))
            End Try

            wait(200)

        End If

        ' --------------------------

        Dim au As String
        Dim to_database As String
        Dim id As Integer
        au = tiff_num & "*" & c_info

        If foxproConnection = True Then
            'FoxPro
            to_database = "s:\job\in_house_software\status_insert_finish3.exe " & au
            id = Shell(to_database, 1)
            wait(200)
        End If
        If sqlConnection = True Then
            'SQL
            Try
                Dim tableValue As String = "OMS.tiff_status"
                Dim updatesValue(2) As Update
                Dim conditionsValue() As Condition = Nothing

                updatesValue(0) = New Update With {
                    .field = "precammed",
                    .value = "1"
                }
                updatesValue(1) = New Update With {
                    .field = "pre_cam_op",
                    .value = Trim(c_info)
                }
                conditionsValue = {
                             New Condition With {
                                 .field = "au_num",
                                .op = "=",
                                .value = Trim(tiff_num) & ".mht"
                                    }
                                }

                Dim jsonString As String = JsonConvert.SerializeObject(New With {
           .updates = updatesValue,
           .table = tableValue,
           .conditions = conditionsValue
       })


                '  Dim uri As String = "http://localhost:8009/api/omsupdate?"  'Development
                Dim uri As String = "https://internal.beta-layout.com/oms/api/omsupdate?"
                Dim rawResponse As String

                rawResponse = SendOMSRequest(uri, jsonString)
            Catch
                SendErrorLog("MHT_FILL_AND_PRINT_ERROR", "update_faxback -SQL REQUEST FAILED " + Trim(tiff_num))
            End Try

        End If

        ' -------------------------------------

        Dim faxArr(), orderArr() As String
        Dim iAmHighSpec As Boolean = False
        faxArr = fax_database.Split(",")
        If faxArr(14) = "'1'" Then iAmHighSpec = True


    End Sub


    Private Sub update_names(ByVal f_info As String)
        Dim id1 As Integer
        Dim names_database As String
        Dim names_database2 As String
        Dim ord As String
        ord = f_info

        ord = ord.Replace(" ", "*")

        Do While InStr(ord, "**") > 0
            ord = ord.Replace("**", "*")
        Loop

        If foxproConnection = True Then 'And useFoxproConnection = True Then
            names_database = "s:\job\in_house_software\home_names_insert_finish.exe " & ord
            names_database2 = "s:\job\in_house_software\home_names_insert_finish2.exe " & ord
            id1 = Shell(names_database, 1)
            wait(200)
        End If
        If sqlConnection = True Then ' And useFoxproConnection = True Then
            Dim au_num As String = ord
            Try
                Dim stamp As Date
                stamp = Now
                Dim timeStamp As Long
                timeStamp = CStr(stamp.Ticks)
                Dim star As Integer
                Dim au As String
                Dim pre As String
                Dim dat As String
                Dim ul As String
                Dim ord_base As String
                Dim tim As String = "0"

                star = InStr(au_num, "*")
                pre = Trim(Mid(au_num, 1, star - 1))
                pre = UCase(pre)
                au_num = Mid(au_num, star + 1)

                star = InStr(au_num, "*")
                au = Trim(Mid(au_num, 1, star - 1))
                au_num = Mid(au_num, star + 1)
                star = InStr(au, ".MH")
                au = Trim(Mid(au, 1, star - 1))

                star = InStr(au_num, "*")
                dat = Trim(Mid(au_num, 1, star - 1))
                au_num = Mid(au_num, star + 1)
                dat = Trim(Mid(dat, 1, 10))
                dat = dat.Replace("-", "/")

                star = InStr(au_num, "*")
                tim = Trim(Mid(au_num, 1, star - 4))
                au_num = Mid(au_num, star + 1)


                ul = Trim(au_num)
                ' If ul = "True" Then
                ' ul = "1"
                ' Else
                ' ul = "0"
                ' End If

                ord_base = au
                star = InStr(ord_base, "_")
                If star > 0 Then
                    ord_base = Mid(ord_base, 1, star - 1)
                End If
                star = InStr(ord_base, "-")
                If star > 0 Then
                    ord_base = Mid(ord_base, 1, star - 1)
                End If

                Dim jsonNames As String = "{"
                jsonNames = "{"
                jsonNames = jsonNames & """" & "time_stamp" & """" & ":" & CLng(timeStamp) & ","
                jsonNames = jsonNames & """" & "order_base" & """" & ":" & """" & UCase(ord_base) & """" & ","
                jsonNames = jsonNames & """" & "order_num" & """" & ":" & """" & UCase(au) & """" & ","
                jsonNames = jsonNames & """" & "operator" & """" & ":" & """" & UCase(pre) & """" & ","
                jsonNames = jsonNames & """" & "done_date" & """" & ":" & """" & dat & """" & ","
                jsonNames = jsonNames & """" & "ul_logo" & """" & ":" & """" & ul & """" & ","
                jsonNames = jsonNames & """" & "time" & """" & ":" & """" & tim & """" & ","
                jsonNames = jsonNames & """" & "invoiced" & """" & ":" & """" & CInt(0) & """" & ","


                jsonNames = Trim(jsonNames)
                jsonNames = Mid(jsonNames, 1, jsonNames.Length - 1)
                jsonNames = jsonNames + "}"

                ' Dim uri As String = "http://localhost:8009/api/omscreatenames?"  'Development
                Dim uri As String = "https://internal.beta-layout.com/oms/api/omscreatenames?"
                Dim rawResponse As String

                rawResponse = SendOMSRequest(uri, jsonNames)
                Try
                    Dim jsonObject As JObject = JObject.Parse(rawResponse)
                    If jsonObject.HasValues AndAlso jsonObject.Property("msg") IsNot Nothing Then
                        Dim msg As String = jsonObject("msg").ToString()
                        If InStr(msg, "SUCCESS") = 0 Then
                            SendErrorLog("MHT_FILL_AND_PRINT_ERROR", "update_faxback - " + au_num)
                        End If
                    End If
                Catch ex As Exception
                    SendErrorLog("MHT_FILL_AND_PRINT_ERROR", "update_faxback - SQL REQUEST FAILED " + au_num)
                End Try

            Catch
                SendErrorLog("MHT_FILL_AND_PRINT_ERROR", "update_faxback - SQL REQUEST FAILED " + au_num)
            End Try

            wait(200)

        End If
    End Sub


End Module
