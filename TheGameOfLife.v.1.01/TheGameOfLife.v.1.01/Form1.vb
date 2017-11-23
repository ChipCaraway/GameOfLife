Public Class Form1

    Structure Person
        Dim PersonType As String
        Dim FName As String
        Dim MName As String
        Dim Lname As String
        Dim Gender As String
        Dim Age As String
        Dim ParentMale As String
        Dim ParentFemale As String
        Dim Reproductive As String
        Dim MoveFirst As String
        Dim MoveSecond As String
        Dim MoveThird As String
        Dim MoveForth As String
        Dim MateFirst As Boolean
        Dim CurrentLeft As String
        Dim CurrentTop As String
    End Structure

    Dim BMP As New Drawing.Bitmap(1000, 800)
    Dim GFX As Graphics = Graphics.FromImage(BMP)
    Public LastTop As Integer = 0
    Public LastLeft As Integer = 0
    Public Universe(1000, 800) As Person
    Public Counts(1000, 1) As String
    Public FFname(100000) As String
    Public MFname(100000) As String
    Public Lname(100000) As String
    Public FFCount As Single = 0
    Public MFCount As Single = 0
    Public LCount As Single = 0

    Public RunStamp = Format(Now, "yyyyMMdd_HHmmss")


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim Rec As String = ""
        Dim N As Double = 0

        Counts(0, 0) = "LIFES"
        Counts(1, 0) = "MALES"
        Counts(2, 0) = "FEMALES"
        Counts(3, 0) = "CHILDREN"
        Counts(4, 0) = "FERTILE"
        Counts(5, 0) = "POST-FERTILE"
        Counts(6, 0) = "Current Year"
        Counts(7, 0) = "Male Child"
        Counts(8, 0) = "Male Fertile"
        Counts(9, 0) = "Male Post-Fertile"
        Counts(10, 0) = "Female Child"
        Counts(11, 0) = "Female Fertile"
        Counts(12, 0) = "Female Post-Fertile"
        Counts(13, 0) = "Moves"
        Counts(14, 0) = "Mated Pairs"
        Counts(15, 0) = "Births"
        Counts(16, 0) = "Deaths"

        Counts(0, 1) = "0"
        Counts(1, 1) = "0"
        Counts(2, 1) = "0"
        Counts(3, 1) = "0"
        Counts(4, 1) = "0"
        Counts(5, 1) = "0"
        Counts(6, 1) = "2100"
        Counts(7, 1) = "0"
        Counts(8, 1) = "0"
        Counts(9, 1) = "0"
        Counts(10, 1) = "0"
        Counts(11, 1) = "0"
        Counts(12, 1) = "0"
        Counts(13, 1) = "0"
        Counts(14, 1) = "0"
        Counts(15, 1) = "0"
        Counts(16, 1) = "0"

        Label1.Text = Counts(0, 0) + StrDup(30 - Len(Counts(0, 0)), ".")
        Label2.Text = Counts(1, 0) + StrDup(30 - Len(Counts(1, 0)), ".")
        Label3.Text = Counts(2, 0) + StrDup(30 - Len(Counts(2, 0)), ".")
        Label4.Text = Counts(3, 0) + StrDup(30 - Len(Counts(3, 0)), ".")
        Label5.Text = Counts(4, 0) + StrDup(30 - Len(Counts(4, 0)), ".")
        Label6.Text = Counts(5, 0) + StrDup(30 - Len(Counts(5, 0)), ".")
        Label7.Text = Counts(6, 0) + StrDup(30 - Len(Counts(6, 0)), ".")
        Label8.Text = Counts(7, 0) + StrDup(30 - Len(Counts(7, 0)), ".")
        Label9.Text = Counts(8, 0) + StrDup(30 - Len(Counts(8, 0)), ".")
        Label10.Text = Counts(9, 0) + StrDup(30 - Len(Counts(9, 0)), ".")
        Label11.Text = Counts(10, 0) + StrDup(30 - Len(Counts(10, 0)), ".")
        Label12.Text = Counts(11, 0) + StrDup(30 - Len(Counts(11, 0)), ".")
        Label13.Text = Counts(12, 0) + StrDup(30 - Len(Counts(12, 0)), ".")
        Label14.Text = Counts(13, 0) + StrDup(30 - Len(Counts(13, 0)), ".")
        Label15.Text = Counts(14, 0) + StrDup(30 - Len(Counts(14, 0)), ".")
        Label16.Text = Counts(15, 0) + StrDup(30 - Len(Counts(15, 0)), ".")
        Label17.Text = Counts(16, 0) + StrDup(30 - Len(Counts(16, 0)), ".")
        Label18.Text = "" 'Counts(17, 0) + StrDup(30 - Len(Counts(17, 0)), ".")
        Label19.Text = "" 'Counts(18, 0) + StrDup(30 - Len(Counts(18, 0)), ".")
        Label20.Text = "" 'Counts(19, 0) + StrDup(30 - Len(Counts(19, 0)), ".")

        Label21.Text = StrDup(7 - Len(Counts(0, 1)), " ") + Counts(0, 1)
        Label22.Text = StrDup(7 - Len(Counts(1, 1)), " ") + Counts(1, 1)
        Label23.Text = StrDup(7 - Len(Counts(2, 1)), " ") + Counts(2, 1)
        Label24.Text = StrDup(7 - Len(Counts(3, 1)), " ") + Counts(3, 1)
        Label25.Text = StrDup(7 - Len(Counts(4, 1)), " ") + Counts(4, 1)
        Label26.Text = StrDup(7 - Len(Counts(5, 1)), " ") + Counts(5, 1)
        Label27.Text = StrDup(5 - Len(Counts(6, 1)), " ") + Counts(6, 1)
        Label28.Text = StrDup(7 - Len(Counts(7, 1)), " ") + Counts(7, 1)
        Label29.Text = StrDup(7 - Len(Counts(8, 1)), " ") + Counts(8, 1)
        Label30.Text = StrDup(7 - Len(Counts(9, 1)), " ") + Counts(9, 1)
        Label31.Text = StrDup(7 - Len(Counts(10, 1)), " ") + Counts(10, 1)
        Label32.Text = StrDup(7 - Len(Counts(11, 1)), " ") + Counts(11, 1)
        Label33.Text = StrDup(7 - Len(Counts(12, 1)), " ") + Counts(12, 1)
        Label34.Text = StrDup(7 - Len(Counts(13, 1)), " ") + Counts(13, 1)
        Label35.Text = StrDup(7 - Len(Counts(14, 1)), " ") + Counts(14, 1)
        Label36.Text = StrDup(7 - Len(Counts(15, 1)), " ") + Counts(15, 1)
        Label37.Text = StrDup(7 - Len(Counts(16, 1)), " ") + Counts(16, 1)
        Label38.Text = "" 'StrDup(7 - Len(Counts(17, 1)), " ") + Counts(17, 1)
        Label39.Text = "" 'StrDup(7 - Len(Counts(18, 1)), " ") + Counts(18, 1)
        Label40.Text = "" 'StrDup(7 - Len(Counts(19, 1)), " ") + Counts(19, 1)

        FileOpen(1, "Male.First.txt", OpenMode.Input)
        While Not EOF(1)
            Rec = LineInput(1)
            MFname(N) = Trim(Mid(Rec, 1, InStr(Rec, " ") - 1))
            N = N + 1
            Application.DoEvents()
        End While
        FileClose(1)

        N = 0
        FileOpen(1, "Female.First.txt", OpenMode.Input)
        While Not EOF(1)
            Rec = LineInput(1)
            FFname(N) = Trim(Mid(Rec, 1, InStr(Rec, " ") - 1))
            N = N + 1
            Application.DoEvents()
        End While
        FileClose(1)

        N = 0
        FileOpen(1, "Last.txt", OpenMode.Input)
        While Not EOF(1)
            Rec = LineInput(1)
            Lname(N) = Trim(Mid(Rec, 1, InStr(Rec, " ") - 1))
            N = N + 1
            Application.DoEvents()
        End While
        FileClose(1)


        GFX.FillRectangle(Brushes.Black, 0, 0, 1000, 800)
        PictureBox1.Image = BMP

    End Sub

    Public Function CreateLife(LastLeft As Single, LastTop As Single) As String

        CreateLife = 0
        Dim PercentMales As Double = 40
        Dim PercentFemales As Double = 60
        Dim PercentChild As Double = 33.2
        Dim PercentFertile As Double = 33.4
        Dim PercentMenopausal As Double = 33.4
        Dim ImgFile As String = ""

        Dim Seed1 As Single = Val(Mid(Format(Now, "yyyyMMddHHmmsss"), 15, 1))
        Dim Seed2 As Single = Val(Mid(Format(Now, "yyyyMMddHHmmsss"), 14, 1))
        Dim Seed3 As Single = Val(Mid(Format(Now, "yyyyMMddHHmmsss"), 14, 2))
        Randomize()
        Dim Rand1 As Integer = CInt(Math.Floor((3 - 1 + 1) * Rnd())) + 1
        Randomize()
        Dim Rand2 As Integer = CInt(Math.Floor((100 - 1 + 1) * Rnd())) + 1
        Randomize()
        Dim Rand3 As Integer = CInt(Math.Floor((100 - 1 + 1) * Rnd())) + 1

        'MsgBox("(" + Str(Rand1) + "," + Str(Rand2) + "," + Str(Rand3) + ")")

        If Rand1 = 1 Then
            CreateLife = "EMPTY"
            ImgFile = "BLANK.bmp"
            Universe(LastLeft, LastTop).PersonType = "BLANK"
            Universe(LastLeft, LastTop).CurrentLeft = Trim(Str(LastLeft))
            Universe(LastLeft, LastTop).CurrentTop = Trim(Str(LastTop))
            GoTo DRAW_IT
        End If

        If Rand2 > 0 And Rand2 <= PercentMales Then
            If Rand3 > 0 And Rand3 <= PercentChild Then
                CreateLife = "MCW"
                ImgFile = "M_C_W.bmp"
                Counts(0, 1) = Val(Counts(0, 1)) + 1
                Counts(1, 1) = Val(Counts(1, 1)) + 1
                Counts(3, 1) = Val(Counts(3, 1)) + 1
                Counts(7, 1) = Val(Counts(7, 1)) + 1
                Universe(LastLeft, LastTop).PersonType = "MCW"
                Universe(LastLeft, LastTop).FName = GetMaleFirstName()
                Universe(LastLeft, LastTop).MName = "A"
                Universe(LastLeft, LastTop).Lname = GetLastName()
                Universe(LastLeft, LastTop).Gender = "MALE"
                Universe(LastLeft, LastTop).Age = Trim(Str(Rand3))
                Universe(LastLeft, LastTop).ParentMale = "FOUNDER MALE"
                Universe(LastLeft, LastTop).ParentFemale = "FOUNDER FEMALE"
                Universe(LastLeft, LastTop).Reproductive = "CHILD"
                Universe(LastLeft, LastTop).MoveFirst = "R"
                Universe(LastLeft, LastTop).MoveSecond = "D"
                Universe(LastLeft, LastTop).MoveThird = "U"
                Universe(LastLeft, LastTop).MoveForth = "L"
                Universe(LastLeft, LastTop).MateFirst = True
                Universe(LastLeft, LastTop).CurrentLeft = Trim(Str(LastLeft))
                Universe(LastLeft, LastTop).CurrentTop = Trim(Str(LastTop))
                GoTo DRAW_IT
            End If
            If Rand3 > PercentChild And Rand3 <= PercentChild + PercentFertile Then
                CreateLife = "MFW"
                ImgFile = "M_F_W.bmp"
                Counts(0, 1) = Val(Counts(0, 1)) + 1
                Counts(1, 1) = Val(Counts(1, 1)) + 1
                Counts(4, 1) = Val(Counts(4, 1)) + 1
                Counts(8, 1) = Val(Counts(8, 1)) + 1
                Universe(LastLeft, LastTop).PersonType = "MFW"
                Universe(LastLeft, LastTop).FName = GetMaleFirstName()
                Universe(LastLeft, LastTop).MName = "E"
                Universe(LastLeft, LastTop).Lname = GetLastName()
                Universe(LastLeft, LastTop).Gender = "MALE"
                Universe(LastLeft, LastTop).Age = Trim(Str(Rand3))
                Universe(LastLeft, LastTop).ParentMale = "FOUNDER MALE"
                Universe(LastLeft, LastTop).ParentFemale = "FOUNDER FEMALE"
                Universe(LastLeft, LastTop).Reproductive = "FERTILE"
                Universe(LastLeft, LastTop).MoveFirst = "D"
                Universe(LastLeft, LastTop).MoveSecond = "U"
                Universe(LastLeft, LastTop).MoveThird = "L"
                Universe(LastLeft, LastTop).MoveForth = "R"
                Universe(LastLeft, LastTop).MateFirst = True
                Universe(LastLeft, LastTop).CurrentLeft = Trim(Str(LastLeft))
                Universe(LastLeft, LastTop).CurrentTop = Trim(Str(LastTop))
                GoTo DRAW_IT
            End If
            If Rand3 > PercentChild + PercentFertile And Rand3 <= PercentChild + PercentFertile + PercentMenopausal Then
                CreateLife = "MMW"
                ImgFile = "M_M_W.bmp"
                Counts(0, 1) = Val(Counts(0, 1)) + 1
                Counts(1, 1) = Val(Counts(1, 1)) + 1
                Counts(5, 1) = Val(Counts(5, 1)) + 1
                Counts(9, 1) = Val(Counts(9, 1)) + 1
                Universe(LastLeft, LastTop).PersonType = "MMW"
                Universe(LastLeft, LastTop).FName = GetMaleFirstName()
                Universe(LastLeft, LastTop).MName = "I"
                Universe(LastLeft, LastTop).Lname = GetLastName()
                Universe(LastLeft, LastTop).Gender = "MALE"
                Universe(LastLeft, LastTop).Age = Trim(Str(Rand3))
                Universe(LastLeft, LastTop).ParentMale = "FOUNDER MALE"
                Universe(LastLeft, LastTop).ParentFemale = "FOUNDER FEMALE"
                Universe(LastLeft, LastTop).Reproductive = "POST-FERTILE"
                Universe(LastLeft, LastTop).MoveFirst = "U"
                Universe(LastLeft, LastTop).MoveSecond = "L"
                Universe(LastLeft, LastTop).MoveThird = "R"
                Universe(LastLeft, LastTop).MoveForth = "D"
                Universe(LastLeft, LastTop).MateFirst = True
                Universe(LastLeft, LastTop).CurrentLeft = Trim(Str(LastLeft))
                Universe(LastLeft, LastTop).CurrentTop = Trim(Str(LastTop))
                GoTo DRAW_IT
            End If
        End If
        If Rand2 > PercentMales And Rand2 <= PercentMales + PercentFemales Then
            If Rand3 > 0 And Rand3 <= PercentChild Then
                CreateLife = "FCW"
                ImgFile = "F_C_W.bmp"
                Counts(0, 1) = Val(Counts(0, 1)) + 1
                Counts(2, 1) = Val(Counts(2, 1)) + 1
                Counts(3, 1) = Val(Counts(3, 1)) + 1
                Counts(10, 1) = Val(Counts(10, 1)) + 1
                Universe(LastLeft, LastTop).PersonType = "FCW"
                Universe(LastLeft, LastTop).FName = GetFemaleFirstName()
                Universe(LastLeft, LastTop).MName = "O"
                Universe(LastLeft, LastTop).Lname = GetLastName()
                Universe(LastLeft, LastTop).Gender = "FEMALE"
                Universe(LastLeft, LastTop).Age = Trim(Str(Rand3))
                Universe(LastLeft, LastTop).ParentMale = "FOUNDER MALE"
                Universe(LastLeft, LastTop).ParentFemale = "FOUNDER FEMALE"
                Universe(LastLeft, LastTop).Reproductive = "CHILD"
                Universe(LastLeft, LastTop).MoveFirst = "R"
                Universe(LastLeft, LastTop).MoveSecond = "D"
                Universe(LastLeft, LastTop).MoveThird = "U"
                Universe(LastLeft, LastTop).MoveForth = "L"
                Universe(LastLeft, LastTop).MateFirst = True
                Universe(LastLeft, LastTop).CurrentLeft = Trim(Str(LastLeft))
                Universe(LastLeft, LastTop).CurrentTop = Trim(Str(LastTop))
                GoTo DRAW_IT
            End If
            If Rand3 > PercentChild And Rand3 <= PercentChild + PercentFertile Then
                CreateLife = "FFW"
                ImgFile = "F_F_W.bmp"
                Counts(0, 1) = Val(Counts(0, 1)) + 1
                Counts(2, 1) = Val(Counts(2, 1)) + 1
                Counts(4, 1) = Val(Counts(4, 1)) + 1
                Counts(11, 1) = Val(Counts(11, 1)) + 1
                Universe(LastLeft, LastTop).PersonType = "FFW"
                Universe(LastLeft, LastTop).FName = GetFemaleFirstName()
                Universe(LastLeft, LastTop).MName = "U"
                Universe(LastLeft, LastTop).Lname = GetLastName()
                Universe(LastLeft, LastTop).Gender = "FEMALE"
                Universe(LastLeft, LastTop).Age = Trim(Str(Rand3))
                Universe(LastLeft, LastTop).ParentMale = "FOUNDER MALE"
                Universe(LastLeft, LastTop).ParentFemale = "FOUNDER FEMALE"
                Universe(LastLeft, LastTop).Reproductive = "FERTILE"
                Universe(LastLeft, LastTop).MoveFirst = "L"
                Universe(LastLeft, LastTop).MoveSecond = "U"
                Universe(LastLeft, LastTop).MoveThird = "R"
                Universe(LastLeft, LastTop).MoveForth = "D"
                Universe(LastLeft, LastTop).MateFirst = True
                Universe(LastLeft, LastTop).CurrentLeft = Trim(Str(LastLeft))
                Universe(LastLeft, LastTop).CurrentTop = Trim(Str(LastTop))
                GoTo DRAW_IT
            End If
            If Rand3 > PercentChild + PercentFertile And Rand3 <= PercentChild + PercentFertile + PercentMenopausal Then
                CreateLife = "FMW"
                ImgFile = "F_M_W.bmp"
                Counts(0, 1) = Val(Counts(0, 1)) + 1
                Counts(2, 1) = Val(Counts(2, 1)) + 1
                Counts(5, 1) = Val(Counts(5, 1)) + 1
                Counts(12, 1) = Val(Counts(12, 1)) + 1
                Universe(LastLeft, LastTop).PersonType = "FMW"
                Universe(LastLeft, LastTop).FName = GetFemaleFirstName()
                Universe(LastLeft, LastTop).MName = "Y"
                Universe(LastLeft, LastTop).Lname = GetLastName()
                Universe(LastLeft, LastTop).Gender = "FEMALE"
                Universe(LastLeft, LastTop).Age = Trim(Str(Rand3))
                Universe(LastLeft, LastTop).ParentMale = "FOUNDER MALE"
                Universe(LastLeft, LastTop).ParentFemale = "FOUNDER FEMALE"
                Universe(LastLeft, LastTop).Reproductive = "POST-FERTILE"
                Universe(LastLeft, LastTop).MoveFirst = "D"
                Universe(LastLeft, LastTop).MoveSecond = "L"
                Universe(LastLeft, LastTop).MoveThird = "U"
                Universe(LastLeft, LastTop).MoveForth = "R"
                Universe(LastLeft, LastTop).MateFirst = True
                Universe(LastLeft, LastTop).CurrentLeft = Trim(Str(LastLeft))
                Universe(LastLeft, LastTop).CurrentTop = Trim(Str(LastTop))
                GoTo DRAW_IT
            End If
        End If

DRAW_IT:
        If Universe(LastLeft, LastTop).PersonType <> "BLANK" Then
            ToLog("CREATE WORLD", "(" + Universe(LastLeft, LastTop).CurrentLeft + "," +
                  Universe(LastLeft, LastTop).CurrentTop + ") Created a " + Universe(LastLeft, LastTop).Reproductive +
                  " " + Universe(LastLeft, LastTop).Gender + " with age of " + Universe(LastLeft, LastTop).Age +
                  " named " + Universe(LastLeft, LastTop).FName + " " + Universe(LastLeft, LastTop).MName + " " +
                  Universe(LastLeft, LastTop).Lname + ". . .")
        End If
        GFX.DrawImage(Image.FromFile(ImgFile), LastLeft, LastTop)
        PictureBox1.Image = BMP
        UpdateCounts()

    End Function

    Public Function UpdateCounts()

        UpdateCounts = 0

        Label21.Text = StrDup(7 - Len(Format(Val(Counts(0, 1)), "#,##0")), " ") + Format(Val(Counts(0, 1)), "#,##0")
        Label22.Text = StrDup(7 - Len(Format(Val(Counts(1, 1)), "#,##0")), " ") + Format(Val(Counts(1, 1)), "#,##0")
        Label23.Text = StrDup(7 - Len(Format(Val(Counts(2, 1)), "#,##0")), " ") + Format(Val(Counts(2, 1)), "#,##0")
        Label24.Text = StrDup(7 - Len(Format(Val(Counts(3, 1)), "#,##0")), " ") + Format(Val(Counts(3, 1)), "#,##0")
        Label25.Text = StrDup(7 - Len(Format(Val(Counts(4, 1)), "#,##0")), " ") + Format(Val(Counts(4, 1)), "#,##0")
        Label26.Text = StrDup(7 - Len(Format(Val(Counts(5, 1)), "#,##0")), " ") + Format(Val(Counts(5, 1)), "#,##0")
        Label27.Text = StrDup(5 - Len(Counts(6, 1)), " ") + Counts(6, 1)
        Label28.Text = StrDup(7 - Len(Format(Val(Counts(7, 1)), "#,##0")), " ") + Format(Val(Counts(7, 1)), "#,##0")
        Label29.Text = StrDup(7 - Len(Format(Val(Counts(8, 1)), "#,##0")), " ") + Format(Val(Counts(8, 1)), "#,##0")
        Label30.Text = StrDup(7 - Len(Format(Val(Counts(9, 1)), "#,##0")), " ") + Format(Val(Counts(9, 1)), "#,##0")
        Label31.Text = StrDup(7 - Len(Format(Val(Counts(10, 1)), "#,##0")), " ") + Format(Val(Counts(10, 1)), "#,##0")
        Label32.Text = StrDup(7 - Len(Format(Val(Counts(11, 1)), "#,##0")), " ") + Format(Val(Counts(11, 1)), "#,##0")
        Label33.Text = StrDup(7 - Len(Format(Val(Counts(12, 1)), "#,##0")), " ") + Format(Val(Counts(12, 1)), "#,##0")
        Label34.Text = StrDup(7 - Len(Format(Val(Counts(13, 1)), "#,##0")), " ") + Format(Val(Counts(13, 1)), "#,##0")
        Label35.Text = StrDup(7 - Len(Format(Val(Counts(14, 1)), "#,##0")), " ") + Format(Val(Counts(14, 1)), "#,##0")
        Label36.Text = StrDup(7 - Len(Format(Val(Counts(15, 1)), "#,##0")), " ") + Format(Val(Counts(15, 1)), "#,##0")
        Label37.Text = StrDup(7 - Len(Format(Val(Counts(16, 1)), "#,##0")), " ") + Format(Val(Counts(16, 1)), "#,##0")
        Label38.Text = "" 'StrDup(7 - Len(Counts(17, 1)), " ") + Counts(17, 1)
        Label39.Text = "" 'StrDup(7 - Len(Counts(18, 1)), " ") + Counts(18, 1)
        Label40.Text = "" 'StrDup(7 - Len(Counts(19, 1)), " ") + Counts(19, 1)

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'GFX.DrawImage(Image.FromFile("F_C_W.bmp"), LastLeft, LastTop)
        While LastTop < 800
            CreateLife(LastLeft, LastTop)
            LastLeft = LastLeft + 10
            If LastLeft > 1000 Then
                LastLeft = 0
                LastTop = LastTop + 10
                If LastTop > 800 Then
                    GoTo BAIL_OUT
                End If
            End If
            Application.DoEvents()
        End While

BAIL_OUT:
        Button1.Enabled = False

    End Sub

    Public Function ToLog(Type As String, Rec As String)

        ToLog = 0
        Dim FileName As String = "LOG_" + RunStamp + ".txt"
        FileOpen(1, FileName, OpenMode.Append)
        Print(1, Format(Now, "yyyy-MM-dd HH:mm:sss") + "|" + Type + "|" + Rec + vbCrLf)
        ListBox1.Items.Add(Type + "|" + Rec)
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        FileClose(1)
        Application.DoEvents()

    End Function

    Public Function GetMaleFirstName() As String

        GetMaleFirstName = 0
        If Trim(MFname(MFCount)) <> "" Then
            GetMaleFirstName = MFname(MFCount)
            MFCount = MFCount + 1
            If Trim(MFname(MFCount)) = "" Then
                MFCount = 0
            End If
        Else
            MFCount = 0
            GetMaleFirstName = MFname(MFCount)
            MFCount = MFCount + 1
        End If

    End Function

    Public Function GetFemaleFirstName() As String

        GetFemaleFirstName = 0
        If Trim(FFname(FFCount)) <> "" Then
            GetFemaleFirstName = FFname(FFCount)
            FFCount = FFCount + 1
            If Trim(FFname(FFCount)) = "" Then
                FFCount = 0
            End If
        Else
            FFCount = 0
            GetFemaleFirstName = FFname(FFCount)
            FFCount = FFCount + 1
        End If

    End Function

    Public Function GetLastName() As String

        GetLastName = 0
        If Trim(Lname(LCount)) <> "" Then
            GetLastName = Lname(LCount)
            LCount = LCount + 1
            If Trim(Lname(LCount)) = "" Then
                LCount = 0
            End If
        Else
            LCount = 0
            GetLastName = Lname(LCount)
            LCount = LCount + 1
        End If

    End Function


End Class
