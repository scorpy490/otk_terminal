Module Module1

    ''' <summary>
    ''' 
    ''' </summary>
    Sub Main()
        Dim Cipher As Object
        Dim ConnSQL, Cnins, logfl, famObj, dt, ts1, dk, mas, brig, smena, dtsmena, yestoday, Contr1, Contr2, folder, sqlstr, fl, ruchky, path, buf, arr, k, kdef, rs0, rs1, rs2, rs3, rs4, typestr, razm, reem, nom_pechi, pom, em_up, em_bort, em_down, kodObj
        Dim fso, i, knsp, kpr, kup, d
        Dim dbins(1000) As String
        Dim def(2) As String
        fso = CreateObject("Scripting.FileSystemObject")
        i = 0
        path = ""
        fl = ""
        dk = 0 ' кол-во дублей
        kpr = 0 ' кол-во принятых
        knsp = 0 ' кол-во несопоставленных
        kup = 0 'кол-во упаковок
        kdef = 0 ' кол дефектов упаковки
        d = 0 ' не найденых ШК22
        Dim CnStr = "Provider=SQLOLEDB;Server=srv-otk;Database=otk;Trusted_Connection=yes;Integrated Security=SSPI;Persist Security Info=False"
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
            Console.WriteLine("Идет зарядка...")
            System.Threading.Thread.Sleep(7000)
            Exit Sub
        End If

        ConnSQL = CreateObject("ADODB.Connection")
        ConnSQL.ConnectionString = CnStr
        ConnSQL.Open
        ts1 = fso.OpenTextFile(path, 1, False)
        Do While Not ts1.AtEndOfStream
            buf = ts1.ReadLine
            arr = Split(buf, ";")
            If UBound(arr) > 3 Then
                sqlstr = "Select [TYPE], [razm], [ruchky] from dbo.typeizd where [shtr]=" & arr(3)
                'MsgBox(sqlstr)
                rs0 = ConnSQL.execute(sqlstr)
                typestr = rs0(0).value.ToString
                razm = rs0(1).value.ToString
                ruchky = rs0(2).value.ToString
                If arr(6) = "" Then arr(6) = 1
                If CInt(arr(6)) > 10 Then
                    arr(6) = arr(6) - 10
                    reem = True
                Else reem = False
                End If

                If arr(7) = "" Then arr(7) = "0"

                sqlstr = "Select [Data], [Контролер1], [Контролер2], [Смена] From dbo.smena_def Where id = 1"
                rs1 = ConnSQL.execute(sqlstr)
                'dtsmena = CDate(rs1(0).value.ToString).ToString("yyyyMMdd")
                'yestoday = DateAdd("d", -1, CDate(rs1(0).value.ToString)).ToString("yyyyMMdd")
                Contr1 = rs1(1).value.ToString
                Contr2 = rs1(2).value.ToString

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
                nom_pechi = "Null"
                pom = DBNull.Value
                em_up = "Null"
                em_bort = "Null"
                em_down = "Null"
                mas = DBNull.Value
                brig = "Null"
                rs2 = ConnSQL.execute(sqlstr)
                If rs2.EOF = False Then
                    nom_pechi = rs2(0).value.ToString
                    pom = rs2(1).value
                    em_up = rs2(2).value.ToString
                    em_bort = rs2(3).value.ToString
                    em_down = rs2(4).value.ToString
                    mas = rs2(5).value
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
                rs3 = ConnSQL.Execute(sqlstr)
                If rs3.EOF = False Then
                    kodObj = rs3(0).value.ToString
                    famObj = rs3(1).value.ToString
                Else
                    kodObj = "116"
                    pom = "Обж:" & arr(4)
                    famObj = "Не существует"

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
                rs4 = ConnSQL.Execute(sqlstr)
                If rs4.EOF = False Then
                    Console.WriteLine(arr(2) & vbTab & "Дубль!")
                    dk = dk + 1
                End If
                If arr(5) = "" Then arr(5) = "0"

                dt = CDate(arr(0) & " " & arr(1))
                dbins(k) = "Insert Into dbo.Изделия ([Номер_бригады],[КодОбж],[Помощник],[Дата_период], [Дата],  [Контролер ОТК], [Контроллер ОТК2], [Мастер смены], [Номер_печи], [Объем], [Тип_ванны], [Ручки], [Реэмаоирование], [Сорт], [ID_Brak], [shtr_kod], [Смена], [Емкость],[Емкость_верх], [Емкость_борт], [Порядк_номер_изд], [term_pr], [dop_param])  SELECT " & brig + "," + kodObj + " , '" + pom + "' ,'" + dt & "' ,'" & dtsmena.ToString & "','" & Contr1 & "' ,'" & Contr2 & "' ,'" & mas & "' ," & nom_pechi + " ," + razm + " ,'" + typestr + "' ,'" + ruchky.ToString + "' ,'" + reem.ToString + "' ," + arr(6) + " ," + def(0) + " ," + arr(2) + ", " + smena + ", " + em_down + "," + em_up + "," + em_bort + "," + arr(5) & ", 'True'," & def(1)
                kpr = kpr + 1
            Else ' предъявление
                sqlstr = "SELECT [shtr_kod] FROM dbo.[Изделия] WHERE [shtr_kod]=" & arr(2)
                If ConnSQL.Execute(sqlstr).EOF = True Then

                    'Console.WriteLine(arr(2) & " не существует")
                    'errfl.WriteLine(CDate(arr(0) & " " & arr(1)) & vbTab & Now.ToShortTimeString & vbTab & arr(2) & " не существует")
                    d = d + 1
                    Continue Do
                End If

                If Left(arr(3), 1) = 2 Then
                    sqlstr = "Update dbo.Изделия SET [predjvl]='true', [DataUp] ='" & CDate(arr(0) & " " & arr(1)) & "', [NomUp] =" & Mid(arr(3), 2, 11) & " WHERE [shtr_kod]=" & arr(2)
                    kup = kup + 1

                Else
                    sqlstr = "Update dbo.Изделия SET [predjvl]='false', [DataUp] ='" & CDate(arr(0) & " " & arr(1)) & "', DefUp =" & arr(3) & ", [NomUp]=null WHERE [shtr_kod]=" & arr(2)
                    kdef = kdef + 1

                End If
                dbins(k) = sqlstr

            End If
            k = k + 1

        Loop
        ConnSQL.Close
        Cnins = CreateObject("ADODB.Connection")
        Cnins.ConnectionString = CnStr
        Cnins.Open
        Cnins.BeginTrans
        For i = 0 To k - 1
            Cnins.execute = dbins(i)
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
        Console.WriteLine("Готово! ")
        Console.WriteLine("Всего принято:" & vbTab & vbTab & kpr)
        Console.WriteLine("Не сопоставлено:" & vbTab & knsp)
        Console.WriteLine("Дублей:" & vbTab & vbTab & vbTab & dk)
        If kup > 0 Or kdef > 0 Or d > 0 Then

            Console.WriteLine("=============================================================================================")
            Console.WriteLine("Упаковка:")
            Console.WriteLine("Всего: " & vbTab & vbTab & kup + kdef)
            Console.WriteLine("Упаковано: " & vbTab & vbTab & kup)
            Console.WriteLine("Браков: " & vbTab & vbTab & kdef)
            Console.WriteLine("Не существует: " & vbTab & vbTab & d)
        End If
        System.Threading.Thread.Sleep(7000)


    End Sub


    'Public Declare Function GetError Lib "stdCipherLab.dll" Alias "stdGetError" (ByRef szData As String) As Integer

End Module
