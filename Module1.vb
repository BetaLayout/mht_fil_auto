'VER 17

Imports System.IO
Imports System
Imports System.Decimal
Imports System.IO.File
Imports Microsoft.VisualBasic


Imports System.Data.OleDb
Module Module1

    'Declare Sub Sleep Lib "Kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

    Dim times_run As Integer
    Dim myarraylist As Object
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



    ' ---------------------------

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


    ' ------------------

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

        ' If onholds = True Then
        ' 'MsgBox("THERE ARE SOME OFF HOLDS BACK FROM INDIA :-)")
        ' Console.WriteLine("THERE ARE SOME OFF HOLDS BACK FROM INDIA ! .. IN C_OFF_HOLDS FOLDER")
        ' on_hold_message = True
        ' End If


    End Sub




    Private Sub wait(ByVal interval As Integer)
        'Dim i As Integer
        ' Do While i < interval
        ' i = i + 1
        ' Loop
        Threading.Thread.Sleep(interval)
    End Sub


    Sub Main()

        Dim folderwatch As Integer

        folderwatch = MsgBox("Is FolderWatch Program running?", vbQuestion + vbYesNo, "Continue?")

        If folderwatch = 7 Then GoTo leave_program

        Dim p = New Process()
        p.StartInfo.FileName = "c:\david\vbasic\mht_fil_auto\copy_directory.bat"

        on_hold_message = False
        myarraylist = CreateObject("System.Collections.ArrayList")
        uploadlist = CreateObject("System.Collections.ArrayList")

        times_run = 0

        Dim source As String
        source = "S:\Job\CAM-India\Work_Finished"
        Dim upload_source As String
        upload_source = "S:\Job\CAM-India\Upload"



        'source = "c:\david\wip\artifex"

check_again:

        donelist = CreateObject("System.Collections.ArrayList")
        donelist.clear()

        If Not (File.Exists("T:\i_am_here.txt")) Then GoTo no_network
        If Not (File.Exists("S:\i_am_here.txt")) Then GoTo no_network
        'Notes()


        If times_run = 0 Then
            get_uploads(upload_source) ' just initialize list of au's in the folder
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


            check_for_uploads(upload_source)  'Uploads data 

        End If



        times_run = times_run + 1
        Console.WriteLine(" ********************* ")

no_network:

        wait(450000)  ' about 10 mins '450000

        'wait(5000)  ' For Testing
        'times_run = times_run

        GoTo check_again

leave_program:

    End Sub

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


    End Sub

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


    End Sub


    Sub Create_Note(ByVal au As String, ByVal file1 As String)
        If File.Exists("c:\in_house_files\note_file.txt") Then File.Delete("c:\in_house_files\note_file.txt")
        Dim note_file As New FileStream("c:\in_house_files\note_file.txt", FileMode.Append, FileAccess.Write)
        Dim file_out As New StreamWriter(note_file)
        Dim dash As Integer

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
        old_note = "T:\Database3\notes\" & old_note

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

    End Sub

    Sub Notes()
        Dim Input_Dir As String
        Input_Dir = "C:\Source"

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


                Dim old_note As String
                old_note = au & "_note.txt"
                old_note = "T:\Database3\notes\" & old_note
                If File.Exists(old_note) Then Kill(old_note)
                File.Move("c:\in_house_files\note_final.txt", old_note)

                database_note(au)

            End If
        Next
    End Sub





    Sub UNZIP()
        Dim Input_Dir As String
        Input_Dir = "C:\Source"

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
                    c = c + 1
                    dot = InStr(UCase(order.name), ".ZIP")
                    au = Mid(order.name, 1, dot - 1)


                    If Not (myarraylist.contains(au)) Then
                        myarraylist.add(au)
                        GoTo not_this_time
                    End If

                    new_orders = new_orders + 1

                    '************************************
                    'seperate_on_holds
                    onhold = False
                    onhold = AmIOnHold(au)
                    'seperate Assembly Orders
                    'anAssembly = False
                    'anAssembly = AmIAssembly(au)

                    'If InStr(order.name, "20140P01") > 0 Then GoTo skip_me
                    ' If InStr(order.name, "24932P01") > 0 Then GoTo skip_me

                    'If anAssembly = True Then
                    ' File.Copy(order.fullname, "s:\job\Dinesh_Assembly\" & order.name, True)
                    'End If
                    If onhold = False Then
                        File.Move(order.fullname, "c:\source\" & order.name)
                        Console.WriteLine(au & "   " & CStr(Now.Hour) & ":" & CStr(Now.Minute))
                    Else
                        File.Move(order.fullname, "t:\off_holds\" & order.name)
                        onholds = True
                        Tiff_Status_R(order.name)
                        Console.WriteLine(au & "   " & CStr(Now.Hour) & ":" & CStr(Now.Minute) & "--- OFF_HOLD")
                        new_orders = new_orders - 1
                    End If
skip_me:

                End If
                'c = c + 1
            End If

not_this_time:

            ' If (InStr(UCase(order.name), "_ONHOLD_") > 0) Then ' And new_orders > 0) Then
            ' File.Move(order.fullname, "c:\source\" & order.name)
            ' End If

            If (InStr(UCase(order.name), ".TXT") > 0) Then ' And new_orders > 0) Then
                File.Move(order.fullname, "S:\Job\CAM-India\Work\new_data_for_work\Status_Updates\" & order.name)  ' They are then handles in here by FolderWatcher
                Console.WriteLine(order.name & " -- Moved")
            End If


        Next

        ' If onholds = True Then
        ' 'MsgBox("THERE ARE SOME OFF HOLDS BACK FROM INDIA :-)")
        ' Console.WriteLine("THERE ARE SOME OFF HOLDS BACK FROM INDIA ! .. IN C_OFF_HOLDS FOLDER")
        ' on_hold_message = True
        ' End If


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


        Return onhold
    End Function


    '*************************

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


        Return anAssembly
    End Function

    '**************************

    Sub update_database()

        Dim Input_Dir As String
        Input_Dir = "c:\source"


        'UPLOADING
        If File.Exists("C:\David\VBASIC\mht_fil_auto\upload_also.txt") Then
            Dim name As String
            Dim dash As Integer
            Dim files3 As String() = Directory.GetFiles(Input_Dir)
            Dim file3 As String
            For Each file3 In files3
                If InStr(UCase(file3), "ERR") = 0 Then
                    If InStr(UCase(file3), ".ZIP") > 0 Or InStr(UCase(file3), ".PDF") > 0 Then
                        dash = file3.LastIndexOf("\")
                        name = Mid(file3, dash + 2)
                        name = "s:\uploaded_files_temp\" & name
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

                        '******  Update tiff_status database here ********************************

                        '  Dim au As String
                        '  Dim to_database As String
                        '  Dim id As Integer
                        '  au = AppendFileContent5 & "*" & cam_name
                        '  to_database = "s:\job\in_house_software\status_insert_finish3.exe " & au
                        '  id = Shell(to_database)

                        ' ****************************************************************************

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

        If Not (File.Exists("c:\in_house_files\temp.txt")) Then GoTo leave

        Dim fr As New FileStream("c:\in_house_files\temp.txt", FileMode.Open, FileAccess.Read)
        Dim r As New StreamReader(fr)
        AppendFileContent = r.ReadToEnd()
        r.Close()
        Dim fs As New FileStream("s:\job\david\dan.txt", FileMode.Append, FileAccess.Write)
        Dim s As New StreamWriter(fs)
        s.WriteLine(AppendFileContent)
        s.Close()


        If Not (File.Exists("c:\in_house_files\temp2.txt")) Then GoTo leave

        Dim fr2 As New FileStream("c:\in_house_files\temp2.txt", FileMode.Open, FileAccess.Read)
        Dim r2 As New StreamReader(fr2)
        AppendFileContent2 = r2.ReadToEnd()
        r2.Close()
        Dim fs2 As New FileStream("s:\job\david\faxback_info.txt", FileMode.Append, FileAccess.Write)
        Dim s2 As New StreamWriter(fs2)
        s2.WriteLine(AppendFileContent2)
        s2.Close()


        If Not (File.Exists("c:\in_house_files\temp3.txt")) Then GoTo leave

        Dim fr3 As New FileStream("c:\in_house_files\temp3.txt", FileMode.Open, FileAccess.Read)
        Dim r3 As New StreamReader(fr3)
        AppendFileContent4 = r3.ReadToEnd()
        r3.Close()

        'If RadioButton2.Checked = False Then
        Dim fs3 As New FileStream("s:\job\david\names.txt", FileMode.Append, FileAccess.Write)
        Dim s3 As New StreamWriter(fs3)
        s3.WriteLine(AppendFileContent4)
        s3.Close()
        'End If



already_danned:



        File.Delete("c:\in_house_files\temp.txt")
        GoTo leave1
        'End If

leave:
        MsgBox("No text files in the input directory !!")
leave1:

        If precam_inst = True Then
            Process.Start("c:\in_house_files\precam_inst.xls")
        End If

        If conrad_orders = True Then
            MsgBox("You have to fill in and print CONRAD orders manually")
        End If

        ' MsgBox("FINISHED !! - Information passed to Database ")




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


        names_database = "s:\job\in_house_software\on_hold_home.exe " & ord
        names_database2 = "s:\job\in_house_software\on_hold_home2.exe " & ord


        id1 = Shell(names_database, 1)
        wait(100)

        id1 = Shell(names_database2, 1)
        wait(100)

    End Sub



    Private Sub update_faxback(ByVal f_info As String, ByVal c_info As String)
        wait(500)
        Dim id1 As Integer
        Dim fax_database As String
        ' Dim fax_database2 As String
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
        'fax_database2 = fax_database

        Dim to_dan As String
        to_dan = fax_database

        fax_database = "s:\job\in_house_software\faxback_insert_finish.exe " & fax_database
        id1 = Shell(fax_database, 1)
        wait(200)

        'fax_database2 = "s:\job\in_house_software\faxback_insert_finish2.exe " & fax_database2
        'id1 = Shell(fax_database2, 1)
        'wait(200)


        ' --------------------------

        Dim au As String
        Dim to_database As String
        'Dim to_database2 As String
        Dim id As Integer
        au = tiff_num & "*" & c_info

        to_database = "s:\job\in_house_software\status_insert_finish3.exe " & au
        id = Shell(to_database, 1)
        wait(200)

        'to_database2 = "s:\job\in_house_software\status_insert_finish3_2.exe " & au
        'id = Shell(to_database2, 1)
        'wait(200)

        ' -------------------------------------

        to_dan = "s:\job\in_house_software\dan_insert_finish.exe " & to_dan
        ' id1 = Shell(to_dan, 1)
        'wait(500)
        data_update_done = True

    End Sub



    Private Sub update_names(ByVal f_info As String)
        Dim id1 As Integer
        Dim names_database As String
        Dim names_database2 As String
        Dim ord As String
        ord = f_info
        'Dim dots As Integer
        'Dim ord2 As String
        'Dim ord3 As String

        'dots = InStr(ord, ":")
        'ord2 = Trim(Mid(ord, dots + 7))
        'ord3 = Trim(Mid(ord, dots - 2, 8))
        'ord = Trim(Mid(ord, 1, dots - 3))



        'ord = ord & "  " & ord2 & "  " & ord3



        ord = ord.Replace(" ", "*")

        Do While InStr(ord, "**") > 0
            ord = ord.Replace("**", "*")
        Loop

        names_database = "s:\job\in_house_software\home_names_insert_finish.exe " & ord
        names_database2 = "s:\job\in_house_software\home_names_insert_finish2.exe " & ord

        id1 = Shell(names_database, 1)
        wait(200)

        'id1 = Shell(names_database2, 1)
        'wait(200)


    End Sub


End Module
