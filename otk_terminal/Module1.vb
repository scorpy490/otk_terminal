'Imports System.Data.SqlClient
Imports System.Collections.Specialized

Module Module1
    Dim CnStr = "Provider=SQLOLEDB;Server=srv-otk;Database=otk;Trusted_Connection=yes;Integrated Security=SSPI;Persist Security Info=False"
    Dim CnStr2 = "Provider=SQLOLEDB;Server=srv15;Database=pech;Trusted_Connection=yes;Integrated Security=SSPI;Persist Security Info=False"
    Dim ConnSQL, cnnPech, conn_fl, dk, d, kpr, knsp, kdef, kup, goreem
    ''' <summary>
    ''' 
    ''' </summary>
    Sub Main()
        Dim Cipher As Object
        Dim Cnins, logfl, ts1, ps, folder, sqlstr, fl, path, buf, arr, k, rez
        Dim fso, i
        Dim dbins As New StringCollection
        goreem = 0
        conn_fl = True
        fso = CreateObject("Scripting.FileSystemObject")
        'fsot = CreateObject("Scripting.FileSystemObject")
        Dim innkpp = "0"
        i = 0
        path = ""
        fl = ""
        dk = 0 ' кол-во дублей
        kpr = 0 ' кол-во принятых
        knsp = 0 ' кол-во несопоставленных
        'kup = 0 'кол-во упаковок
        'kdef = 0 ' кол дефектов упаковки
        d = 0 ' не найденых ШК22
        ps = False
        folder = fso.GetFolder("d:\Terminal\out")
        k = 0
        For Each file In folder.Files
            If Left(file.Name, 3) = "emo" And Right(file.name, 4) = ".txt" Then
                path = file
                fl = file.name
                Exit For
            End If
        Next
        If Len(fl) < 5 Then
            Console.WriteLine("Файл данных не найден!")
            System.Threading.Thread.Sleep(7000)
            Exit Sub
        End If
        'repfl = fso.OpenTextFile("d:\Terminal\tmp.txt", 2, True)
        ConnSQL = CreateObject("ADODB.Connection")
        cnnPech = CreateObject("ADODB.Connection")
        ConnSQL.ConnectionString = CnStr
        cnnPech.ConnectionString = CnStr2
        ConnSQL.Open
        Try
            cnnPech.Open
        Catch ex As Exception
            conn_fl = False
            Console.WriteLine("Не удается получить данные с печей!")
        End Try

        ts1 = fso.OpenTextFile(path, 1, False)
        Do While Not ts1.AtEndOfStream
            buf = ts1.ReadLine
            arr = Split(buf, ";")
            If UBound(arr) > 3 Then
                rez = parse_pr(arr)
            ElseIf UBound(arr) = 3 Then
                rez = parse_vozvr(arr)
            Else
                rez = parse_goreem(arr)
            End If
            If rez Is Nothing Then Continue Do
            For Each i In rez
                dbins.Add(i)
                'repfl.WriteLine(i)
            Next

        Loop
        ConnSQL.Close
        If conn_fl = True Then cnnPech.Close
        Cnins = CreateObject("ADODB.Connection")
        Cnins.ConnectionString = CnStr
        Cnins.Open
        Cnins.BeginTrans
        For Each i In dbins
            'Console.WriteLine(i)
            sqlstr = i
            Cnins.execute(sqlstr)
            'repfl.WriteLine(i)
        Next
        Cnins.CommitTrans
        Cnins.Close
        ts1.Close
        sqlstr = "D:\Terminal\Arhiv\" & fl
        ts1 = fso.GetFile(path)
        ts1.Move(sqlstr)
        logfl = fso.OpenTextFile("d:\Terminal\logs\logdate.csv", 8, True)
        sqlstr = Now.ToShortDateString + " " + Now.ToShortTimeString & vbTab & k & vbTab & kdef & vbTab & dk
        logfl.WriteLine(sqlstr)
        logfl.Close
        If kpr > 0 Then
            Console.WriteLine("Готово! ")
            Console.WriteLine("Всего принято:" & vbTab & vbTab & kpr)
            Console.WriteLine("Не сопоставлено:" & vbTab & knsp)
            Console.WriteLine("Дублей:" & vbTab & vbTab & vbTab & dk)
        End If
        If kup > 0 Or d > 0 Then

            Console.WriteLine("=============================================================================================")
            Console.WriteLine("Возврат:")
            Console.WriteLine("Всего:      " & vbTab & vbTab & kup + d)
            Console.WriteLine("Не найдено: " & vbTab & vbTab & d)
        End If
        If goreem > 0 Then
            Console.WriteLine("=============================================================================================")
            Console.WriteLine("Отправлено на реэмалирование:")
            Console.WriteLine("Всего:      " & vbTab & vbTab & goreem)
        End If
        'repfl.Close
        System.Threading.Thread.Sleep(5000)
        'Process.Start("d:\Terminal\tmp.txt", "Notepad.exe")

        'Console.ReadLine()


    End Sub

    Function parse_pr(arr As Array)
        Dim rez As New StringCollection
        Dim reem, ps, dt, smena, yestoday, dtsmena, kodObj, famobj, q1
        Dim def(2)
        Dim sqlstr = "Select [TYPE], [razm], [ruchky],[pechid] from dbo.typeizd where [shtr]=" & arr(3)
        'MsgBox(sqlstr)
        Dim rs0 = ConnSQL.execute(sqlstr)
        Dim typestr = rs0(0).value.ToString.Trim
        Dim razm = rs0(1).value.ToString
        Dim ruchky = rs0(2).value.ToString
        Dim pechid = rs0(3).value.ToString
        If arr(6) = "" Then arr(6) = 1
        If CInt(arr(6)) > 10 And CInt(arr(6)) < 20 Then
            arr(6) = arr(6) - 10
            reem = True
            ps = False
        ElseIf CInt(arr(6)) > 20 Then
            arr(6) = arr(6) - 20
            ps = True
            reem = False

        Else
            reem = False
            ps = False
        End If

        If arr(7) = "" Then arr(7) = "0"
        'sqlstr = "SELECT [Фамилия] From dbo.[Обжигальщики] WHERE "
        sqlstr = "Select [Data], [Контролер1], [Контролер2], [Смена] From dbo.smena_def Where id = 1"
        Dim rs1 = ConnSQL.execute(sqlstr)
        'dtsmena = CDate(rs1(0).value.ToString).ToString("yyyyMMdd")
        'yestoday = DateAdd("d", -1, CDate(rs1(0).value.ToString)).ToString("yyyyMMdd")
        Dim Contr1 = rs1(1).value.ToString.Trim
        Dim Contr2 = rs1(2).value.ToString.Trim

        If Now.Hour < 7 Then
            dt = DateAdd("d", -1, Now)
            smena = "2"
        ElseIf Now.Hour >= 19 Then
            dt = Now
            smena = "2"
        Else
            dt = Now
            smena = "1"
        End If
        yestoday = DateAdd("d", -1, dt).ToString("yyyyMMdd")
        dtsmena = CDate(dt.ToString).ToString("yyyyMMdd")

        sqlstr = "Select [nom_pech], [Pomochnik], [Емкость_верх], [Емкость_борт], [Емкость_низ], [Мастер], [Бригада], [Смена], [Дата] from dbo.[Сопоставление] WHERE [nom_obj]=" & arr(4) & " And ([Дата]='" & dtsmena & "' OR ([Дата]='" & yestoday & "' AND [Смена]=2 AND CONVERT (time, getdate())<'12:00:00' AND CONVERT (time, getdate())>'07:00:00' ))"
        'MsgBox(sqlstr)
        Dim nom_pechi = "Null"
        Dim pom = "Null"
        Dim em_up = "Null"
        Dim em_bort = "Null"
        Dim em_down = "Null"
        Dim mas = "Null"
        Dim brig = "Null"
        Dim rs2 = ConnSQL.execute(sqlstr)
        If rs2.EOF = False Then
            nom_pechi = rs2(0).value.ToString
            pom = rs2(1).value.ToString.Trim
            em_up = rs2(2).value.ToString
            em_bort = rs2(3).value.ToString
            em_down = rs2(4).value.ToString
            mas = rs2(5).value.ToString.Trim
            brig = rs2(6).value.ToString
            smena = rs2(7).value.ToString
            dtsmena = CDate(rs2(8).value.ToString).ToString("yyyyMMdd")

            'MsgBox(nom_pechi.value.ToString)
        Else

            knsp = knsp + 1
        End If

        If nom_pechi = "" Then nom_pechi = "NULL"
        'If pom = "" Then pom = "NULL"
        If em_up = "" Then em_up = "NULL"
        If em_bort = "" Then em_bort = "NULL"
        If em_down = "" Then em_down = "NULL"
        'If mas = "" Then mas = "NULL"
        If brig = "" Then brig = "NULL"


        sqlstr = "SELECT [ОбжКод], [Фамилия] From dbo.[Обжигальщики] WHere [Номер]=" & arr(4)
        Dim rs3 = ConnSQL.Execute(sqlstr)
        If rs3.EOF = False Then
            kodObj = rs3(0).value.ToString
            famobj = rs3(1).value.ToString.Trim
        Else
            kodObj = "116"
            pom = "Обж:" & arr(4)
            famobj = "Не существует"

        End If
        If InStr(arr(7), ".") > 0 Then
            def(0) = Left(arr(7), InStr(arr(7), ".") - 1)
            def(1) = Right(arr(7), Len(arr(7)) - InStr(arr(7), "."))
            If def(1) = "" Then def(1) = "null"
        Else
            def(0) = arr(7)
            def(1) = "null"
        End If


        sqlstr = "SELECT [shtr_kod] From dbo.[Изделия] Where [shtr_kod]=" & arr(2)
        Dim rs4 = ConnSQL.Execute(sqlstr)
        If rs4.EOF = False Then
            Console.WriteLine(arr(2) & vbTab & "Дубль!")
            dk = dk + 1
            'Exit Function
        End If
        If arr(5) = "" Then arr(5) = "0"

        dt = CDate(arr(0) & " " & arr(1))

        kpr = kpr + 1
        'Данные печи
        If conn_fl = True Then
            'dt = DateAdd(DateInterval.Hour, -12, dt)
            sqlstr = "SELECT [DATA_TIME],[ID_OBJIG],[TIME_OBJIG],[ID_PECH],[ID_OBJIGALSHIC],[TIP_VANNA],[COL_VANNA],[TEMP],[TEMP_MIN],[TEMP_AVG],[TEMP_MAX],[TIME_ITERATION] FROM [pech].[dbo].[WORK_PECH] WHERE [ID_OBJIGALSHIC]=" _
            & arr(4) & " AND [ID_PECH]=" & nom_pechi & " AND [DATA_TIME] >'" & DateAdd(DateInterval.Hour, -12, dt) & "' AND [COL_VANNA]=" & arr(5)
            q1 = cnnPech.Execute(sqlstr)
            Do While Not q1.EOF
                Dim a7 = Replace(q1(7).value.ToString, ",", ".")
                Dim a8 = Replace(q1(8).value.ToString, ",", ".")
                Dim a9 = Replace(q1(9).value.ToString, ",", ".")
                Dim a10 = Replace(q1(10).value.ToString, ",", ".")
                'Console.WriteLine(q1(0).value.ToString & vbTab & q1(1).value.ToString & vbTab & q1(2).value.ToString & vbTab & q1(3).value.ToString & vbTab & q1(4).value.ToString & vbTab & q1(5).value.ToString & vbTab & q1(6).value.ToString & vbTab & q1(7).value.ToString)
                sqlstr = "INSERT INTO [dbo].[WORK_PECH] ([DATA_TIME],[ID_OBJIG],[TIME_OBJIG],[ID_PECH],[ID_OBJIGALSHIC],[TIP_VANNA],[COL_VANNA],[TEMP],[TEMP_MIN],[TEMP_AVG],[TEMP_MAX],[TIME_ITERATION],[shtr]) VALUES ('" & q1(0).value & "'," & q1(1).value & "," & q1(2).value & "," & q1(3).value & "," & q1(4).value & "," & q1(5).value & "," & q1(6).value & "," & a7 & "," & a8 & "," & a9 & "," & a10 & "," & q1(11).value & "," & arr(2) & ")"
                rez.Add(sqlstr)
                'Console.WriteLine(sqlstr)
                'ConnSQL.Execute = sqlstr
                q1.MoveNext
            Loop
        End If
        sqlstr = "Insert Into dbo.Изделия ([Номер_бригады],[КодОбж],[odj_str][Помощник],[Дата_период], [Дата],  [Контролер ОТК], [Контроллер ОТК2], [Мастер смены], [Номер_печи], [Реэмаоирование], [Сорт], [ID_Brak], [shtr_kod], [Смена], [Емкость],[Емкость_верх], [Емкость_борт], [Порядк_номер_изд], [term_pr], [dop_param], [pskstr], [kod_izd]) SELECT " _
            & brig + "," + kodObj + ",'" + famobj + "' , '" + pom + "' ,'" + dt & "' ,'" & dtsmena.ToString & "','" & Contr1 & "' ,'" & Contr2 & "' ,'" & mas & "' ," & nom_pechi + " ,'" + reem.ToString + "' ," + arr(6) + " ," + def(0) + " ," + arr(2) + ", " + smena + ", " + em_down + "," + em_up + "," + em_bort + "," + arr(5) & ", 'True'," & def(1) & ",'" & ps.ToString & "'," + arr(3)
        rez.Add(sqlstr)
        Return rez

    End Function

    Function parse_vozvr(arr As Array)
        Dim rez As New StringCollection
        Dim innkpp
        Dim sqlstr = "SELECT [shtr_kod] FROM dbo.[Изделия] WHERE [shtr_kod]=" & arr(2)
        If ConnSQL.Execute(sqlstr).EOF = True Then
            'Console.WriteLine(arr(2) & " не существует")
            'errfl.WriteLine(CDate(arr(0) & " " & arr(1)) & vbTab & Now.ToShortTimeString & vbTab & arr(2) & " не существует")
            d = d + 1
            Return rez
        End If
        sqlstr = "SELECT [innkpp] FROM dbo.[pretenz_kontr] WHERE [id]=" & arr(3)
        Dim rs5 = ConnSQL.Execute(sqlstr)
        If rs5.EOF = True Then innkpp = "0" Else innkpp = rs5(0).Value.ToString
        sqlstr = "Update dbo.pretenz_van SET [vzvr]='true'  WHERE [shtr]=" & arr(2)
        rez.Add(sqlstr)
        'ConnSQL.Execute(sqlstr)
        sqlstr = "Update dbo.Изделия SET [Data_vozvr] ='" & CDate(arr(0) & " " & arr(1)) & "', [vozvr_inn] ='" & innkpp & "' WHERE [shtr_kod]=" & arr(2)
        rez.Add(sqlstr)
        Return rez
    End Function

    Function parse_goreem(arr As Array)
        Dim sqlstr = "Update dbo.Изделия SET goreem =1 WHERE [shtr_kod]=" & arr(2)
        Dim rez As New StringCollection
        rez.Add(sqlstr)
        goreem = goreem + 1
        Return rez
    End Function


    'Public Declare Function GetError Lib "stdCipherLab.dll" Alias "stdGetError" (ByRef szData As String) As Integer

End Module
