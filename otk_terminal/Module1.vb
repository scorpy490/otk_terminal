'Imports System.Data.SqlClient
Imports System.Collections.Specialized

Module Module1
    Dim CnStr = "Provider=SQLOLEDB;Server=srv-otk;Database=otktmp;Trusted_Connection=yes;Integrated Security=SSPI;Persist Security Info=False"
    Dim CnStr2 = "Provider=SQLOLEDB;Server=srv15;Database=pech;Trusted_Connection=yes;Integrated Security=SSPI;Persist Security Info=False"
    Dim ConnSQL, cnnPech, conn_fl, dk, d, kpr, knsp, kup, goreem, contrl, contr1, erup, kup13
    Dim Contrl_fl = 0
    ''' <summary>
    ''' 
    ''' </summary>
    Sub Main()
        Dim Cnins, ts1, folder, sqlstr, fl, path, buf, arr, rez
        Dim fso, i
        Dim dbins As New StringCollection
        goreem = 0
        conn_fl = True
        fso = CreateObject("Scripting.FileSystemObject")
        'fsot = CreateObject("Scripting.FileSystemObject")
        path = ""
        fl = ""
        dk = 0 ' кол-во дублей
        kpr = 0 ' кол-во принятых
        knsp = 0 ' кол-во несопоставленных
        'kup = 0 'кол-во упаковок
        'kdef = 0 ' кол дефектов упаковки
        d = 0 ' не найденых ШК22
        folder = fso.GetFolder("d:\Terminal\out")
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

        Console.WriteLine("Введите номер контролера:")
        contrl = Console.ReadLine




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
        rez = ""
        Do While Not ts1.AtEndOfStream
            buf = ts1.ReadLine
            arr = Split(buf, ";")
            If arr(0) = "Приемка" Then rez = parse_pr(arr)
            If arr(0) = "Возврат" Then rez = parse_vozvr(arr)
            If arr(0) = "НаРеэмалир" Then rez = parse_goreem(arr)
            If arr(0) = "Упаковка" Then rez = parse_up13(arr)
            If arr(0) = "Отгрузка" Then rez = parse_otgruzka(arr)

            'If UBound(arr) > 3 Then
            '    rez = parse_pr(arr)
            'ElseIf UBound(arr) = 3 Then
            '    rez = parse_vozvr(arr)
            'Else
            '    rez = parse_goreem(arr)
            'End If
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

        If kpr > 0 Then
            Console.WriteLine("Готово! ")
            Console.WriteLine("Всего принято:" & vbTab & vbTab & kpr)
            Console.WriteLine("Не сопоставлено:" & vbTab & knsp)
            Console.WriteLine("Дублей:" & vbTab & vbTab & vbTab & dk)
        End If
        If kup > 0 Or d > 0 Then

            Console.WriteLine("=", 95)
            Console.WriteLine("Возврат:")
            Console.WriteLine("Всего:      " & vbTab & vbTab & kup + d)
            Console.WriteLine("Не найдено: " & vbTab & vbTab & d)
        End If
        If goreem > 0 Then
            Console.WriteLine("=", 95)
            Console.WriteLine("Отправлено на реэмалирование:")
            Console.WriteLine("Всего:      " & vbTab & vbTab & goreem)
        End If
        If kup13 > 0 Then
            Console.WriteLine("=", 95)
            Console.WriteLine("=", 95)
            Console.WriteLine("Упаковано " & kup13 & " изделий.")
            System.Threading.Thread.Sleep(10000)

        End If
        'repfl.Close
        'System.Threading.Thread.Sleep(5000)
        'Process.Start("d:\Terminal\tmp.txt", "Notepad.exe")

        Console.ReadLine()


    End Sub

    Function Parse_pr(arr As Array)
        Dim rez As New StringCollection
        Dim reem, ps, dt, smena, yestoday, dtsmena, kodObj, famobj, q1
        Dim def(2)
        Dim sqlstr = "Select [TYPE], [razm], [ruchky],[pechid] from dbo.typeizd where [shtr]=" & arr(3 + 1)
        'MsgBox(sqlstr)
        Dim rs0 = ConnSQL.execute(sqlstr)
        Dim typestr = rs0(0).value.ToString.Trim
        Dim razm = rs0(1).value.ToString
        Dim ruchky = rs0(2).value.ToString
        Dim pechid = rs0(3).value.ToString
        If arr(6 + 1) = "" Then arr(6 + 1) = 1
        If arr(contrl) = "" Then
            arr(contrl) = "0"
            conn_fl = Contrl_fl + 1
        End If
        If CInt(arr(6 + 1)) > 10 And CInt(arr(6 + 1)) < 20 Then
            arr(6 + 1) = arr(6 + 1) - 10
            reem = True
            ps = False
        ElseIf CInt(arr(6 + 1)) > 20 Then
            arr(6) = arr(6 + 1) - 20
            ps = True
            reem = False

        Else
            reem = False
            ps = False
        End If

        If arr(7 + 1) = "" Then arr(7 + 1) = "0"

        sqlstr = "Select [Фамилия] From dbo.[Мастера] Where nom =" & contrl
        'MsgBox(sqlstr)
        Dim rs1 = ConnSQL.execute(sqlstr)
        If rs1.EOF = False Then
            contr1 = rs1(0).value.ToString
        Else
            Console.WriteLine("Для изделия " + arr(2 + 1) + " неверно указан контролер")
        End If



        'Dim Contr2 = rs1(2).value.ToString.Trim

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


        sqlstr = "SELECT [ОбжКод], [Фамилия] From dbo.[Обжигальщики] WHere [Номер]=" & arr(4 + 1)
        Dim rs3 = ConnSQL.Execute(sqlstr)
        If rs3.EOF = False Then
            kodObj = rs3(0).value.ToString
            famobj = rs3(1).value.ToString.Trim
        Else
            kodObj = "116"
            pom = "Обж:" & arr(4 + 1)
            famobj = "Не существует"

        End If
        If InStr(arr(7 + 1), ".") > 0 Then
            def(0) = Left(arr(7 + 1), InStr(arr(7 + 1), ".") - 1)
            def(1) = Right(arr(7 + 1), Len(arr(7)) - InStr(arr(7 + 1), "."))
            If def(1) = "" Then def(1) = "null"
        Else
            def(0) = arr(7 + 1)
            def(1) = "null"
        End If


        sqlstr = "SELECT [shtr_kod] From dbo.[Изделия] Where [shtr_kod]=" & arr(2 + 1)
        Dim rs4 = ConnSQL.Execute(sqlstr)
        If rs4.EOF = False Then
            Console.WriteLine(arr(2) & vbTab & "Дубль!")
            dk = dk + 1
            'Exit Function
        End If
        If arr(5 + 1) = "" Then arr(5 + 1) = "0"

        dt = CDate(arr(0 + 1) & " " & arr(1 + 1))

        kpr = kpr + 1
        'Данные печи
        If conn_fl = True Then
            'dt = DateAdd(DateInterval.Hour, -12, dt)
            sqlstr = "SELECT [DATA_TIME],[ID_OBJIG],[TIME_OBJIG],[ID_PECH],[ID_OBJIGALSHIC],[TIP_VANNA],[COL_VANNA],[TEMP],[TEMP_MIN],[TEMP_AVG],[TEMP_MAX],[TIME_ITERATION] FROM [pech].[dbo].[WORK_PECH] WHERE [ID_OBJIGALSHIC]=" _
            & arr(4 + 1) & " AND [ID_PECH]=" & nom_pechi & " AND [DATA_TIME] >'" & DateAdd(DateInterval.Hour, -12, dt) & "' AND [COL_VANNA]=" & arr(5 + 1)
            q1 = cnnPech.Execute(sqlstr)
            Do While Not q1.EOF
                Dim a7 = Replace(q1(7).value.ToString, ",", ".")
                Dim a8 = Replace(q1(8).value.ToString, ",", ".")
                Dim a9 = Replace(q1(9).value.ToString, ",", ".")
                Dim a10 = Replace(q1(10).value.ToString, ",", ".")
                'Console.WriteLine(q1(0).value.ToString & vbTab & q1(1).value.ToString & vbTab & q1(2).value.ToString & vbTab & q1(3).value.ToString & vbTab & q1(4).value.ToString & vbTab & q1(5).value.ToString & vbTab & q1(6).value.ToString & vbTab & q1(7).value.ToString)
                sqlstr = "INSERT INTO [dbo].[WORK_PECH] ([DATA_TIME],[ID_OBJIG],[TIME_OBJIG],[ID_PECH],[ID_OBJIGALSHIC],[TIP_VANNA],[COL_VANNA],[TEMP],[TEMP_MIN],[TEMP_AVG],[TEMP_MAX],[TIME_ITERATION],[shtr]) VALUES ('" & q1(0).value & "'," & q1(1).value & "," & q1(2).value & "," & q1(3).value & "," & q1(4).value & "," & q1(5).value & "," & q1(6).value & "," & a7 & "," & a8 & "," & a9 & "," & a10 & "," & q1(11).value & "," & arr(2 + 1) & ")"
                rez.Add(sqlstr)
                'Console.WriteLine(sqlstr)
                'ConnSQL.Execute = sqlstr
                q1.MoveNext
            Loop
        End If
        sqlstr = "Insert Into dbo.Изделия ([Номер_бригады],[КодОбж],[obj_str],[Помощник],[Дата_период], [Дата],  [Контролер ОТК], [Мастер смены], [Номер_печи], [Реэмаоирование], [Сорт], [ID_Brak], [shtr_kod], [Смена], [Емкость],[Емкость_верх], [Емкость_борт], [Порядк_номер_изд], [term_pr], [dop_param], [pskstr], [kod_izd], [ContrEMO_ID]) SELECT " _
            & brig + "," + kodObj + ",'" + famobj + "' , '" + pom + "' ,'" + dt & "' ,'" & dtsmena.ToString & "','" & contr1 & "' ,'" & mas & "' ," & nom_pechi + " ,'" + reem.ToString + "' ," + arr(6 + 1) + " ," + def(0) + " ," + arr(2 + 1) + ", " + smena + ", " + em_down + "," + em_up + "," + em_bort + "," + arr(5 + 1) & ", 'True'," & def(1) & ",'" & ps.ToString & "'," + arr(3 + 1) & "," + contrl
        rez.Add(sqlstr)
        Return rez

    End Function

    Function Parse_vozvr(arr As Array)
        Dim rez As New StringCollection
        Dim innkpp
        Dim sqlstr = "SELECT [shtr_kod] FROM dbo.[Изделия] WHERE [shtr_kod]=" & arr(2 + 1)
        If ConnSQL.Execute(sqlstr).EOF = True Then
            'Console.WriteLine(arr(2) & " не существует")
            'errfl.WriteLine(CDate(arr(0) & " " & arr(1)) & vbTab & Now.ToShortTimeString & vbTab & arr(2) & " не существует")
            d = d + 1
            Return rez
        End If
        sqlstr = "SELECT [innkpp] FROM dbo.[pretenz_kontr] WHERE [id]=" & arr(3 + 1)
        Dim rs5 = ConnSQL.Execute(sqlstr)
        If rs5.EOF = True Then innkpp = "0" Else innkpp = rs5(0).Value.ToString
        sqlstr = "Update dbo.pretenz_van SET [vzvr]='true'  WHERE [shtr]=" & arr(2 + 1)
        rez.Add(sqlstr)
        'ConnSQL.Execute(sqlstr)
        sqlstr = "Update dbo.Изделия SET [Data_vozvr] ='" & CDate(arr(0) & " " & arr(1 + 1)) & "', [vozvr_inn] ='" & innkpp & "' WHERE [shtr_kod]=" & arr(2 + 1)
        rez.Add(sqlstr)
        Return rez
    End Function

    Function Parse_goreem(arr As Array)
        Dim sqlstr = "Update dbo.Изделия SET goreem =3, goreamal='true' WHERE [shtr_kod]=" & arr(0 + 1)
        goreem = goreem + 1
        Return sqlstr
    End Function

    Function Parse_up13(arr As Array)
        Dim sqlstr = "SELECT [shtr_kod], [Сорт],[sort13] FROM dbo.[Изделия] WHERE [shtr_kod]=" & arr(2 + 1)
        Dim rs1 = ConnSQL.Execute(sqlstr)
        If rs1.EOF = True Then
            Console.WriteLine(arr(2 + 1) & " не существует")
            erup = erup + 1
            Return ""
            Exit Function
        End If
        Dim countrs = rs1.RecordCount
        If countrs > 1 Then Console.WriteLine(arr(2 + 1) + "     Дубль     " + countrs)
        Dim sort = rs1(1).Value.ToString
        If sort = "2" Or sort = "6" Or sort = "7" Then sort = "1"
        Dim sort2 = rs1(2).Value.ToString
        If sort2 = "" Then sort2 = sort

        If Left(arr(3 + 1), 1) = 2 Then
            sqlstr = "Update dbo.Изделия SET [DataUp] ='" & CDate(arr(0 + 1) & " " & arr(1 + 1)) & "', [NomUp] =" & Mid(arr(3 + 1), 2, 11) & ", [Sort13]=" & sort2 & ", [Control_13_nom]=" & contrl & " WHERE [shtr_kod]=" & arr(2 + 1)
            kup13 = kup13 + 1
            ConnSQL.execute("Update dbo.sklad SET [13skl]='1' WHERE [shtr]=" & arr(2 + 1))
            Return sqlstr
        Else
            Console.WriteLine("Некорректный номер упаковки: " + arr(3 + 1))
            Return ""
            Exit Function
        End If


    End Function


    Function Parse_otgruzka(arr As Array)

        If arr(2 + 1).Substring(0, 1 = "2") Then
            arr(2 + 1) = arr(2 + 1).Substring(1, 11).TrimStart(0)
        Else
            Return ""
            Exit Function
        End If

        Dim inn = arr(3 + 1).Substring(0, 12).ToString()
        Dim nar = arr(3 + 1).Substring(12, 5).TrimStart(0).ToString()
        Dim year = arr(3 + 1).Substring(17, 4).TrimStart(0).ToString()
        Dim kpp = "000000000"
        If arr(3 + 1).Length > 21 Then kpp = arr(3 + 1).Substring(21, 9)
        If (nar = "") Then nar = "0"
        Dim dt = arr(0 + 1) + " " + arr(1 + 1)
        Dim dt1 = Convert.ToDateTime(dt)
        Dim sqlstr = "Update dbo.Изделия SET [otgr_data]='" + dt1.ToString() + "', [otgr_inn]='" + inn + kpp + "', [otgr_nar]='" + nar + "/" + year + "' WHERE [NomUp]=" + arr(2 + 1)
        'Console.WriteLine(inn + kpp + "    " + nar + "      " + year + "     " + arr[2]);
        Return sqlstr

    End Function

    'Public Declare Function GetError Lib "stdCipherLab.dll" Alias "stdGetError" (ByRef szData As String) As Integer

End Module
