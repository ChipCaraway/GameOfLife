Public Class Form1

    Structure Person
        Dim PersonNumber As String
        Dim PersonType As String
        Dim FName As String
        Dim MName As String
        Dim Lname As String
        Dim Gender As String
        Dim Age As String
        Dim MateNumber As String
        Dim ParentMale As String
        Dim ParentFemale As String
        Dim Reproductive As String
        Dim MoveFirst As String
        Dim MoveSecond As String
        Dim MoveThird As String
        Dim MoveForth As String
        Dim MateFirst As Boolean
        Dim Sick As Boolean
        Dim SickIncubation As Single
        Dim SickContagiousness As Single
        Dim CurrentLeft As String
        Dim CurrentTop As String
    End Structure

    Structure GlobalAge
        Dim MaxAge As Single
        Dim MaleChild As Single
        Dim MaleFertile As Single
        Dim MalePostFertile As Single
        Dim FemaleChild As Single
        Dim FemaleFertile As Single
        Dim FemalePostFertile As Single
    End Structure

    Structure GlobalConfig
        Dim PairsForLife As Boolean
        Dim OffspringTouchesParent As Boolean
        Dim PreFertileDoesNotMove As Boolean
        Dim PostFertileDoesNotMove As Boolean
        Dim NonPairedMovesToFertileSpot As Boolean
        Dim NonPairedMovesToNextSpot As Boolean
        Dim AccidentalDeath As Boolean
    End Structure

    Structure GlobalDisease
        Dim IntroduceDisease As Boolean
        Dim AffectsMales As Boolean
        Dim AffectsFemales As Boolean
        Dim AffectsChildren As Boolean
        Dim AffectsFertile As Boolean
        Dim AffectsElderly As Boolean
        Dim LivesInitiallyInfected As Single
        Dim IncubationPeriod As Single
        Dim Contagiosness As Double
        Dim ModSeed As Integer
    End Structure

    Dim BMP As New Drawing.Bitmap(1000, 800)
    Dim GFX As Graphics = Graphics.FromImage(BMP)
    Public LastTop As Integer = 0
    Public LastLeft As Integer = 0
    Public Age(1) As GlobalAge
    Public Cfg As GlobalConfig
    Public Sick As GlobalDisease
    Public Universe(1000, 800) As Person
    Public Counts(1000, 1) As String
    Public FFname(100000) As String
    Public MFname(100000) As String
    Public Lname(100000) As String
    Public FFCount As Single = 0
    Public MFCount As Single = 0
    Public LCount As Single = 0
    Public PersonNumberSeq As Double = 0
    Public InitialPopulate As Boolean = False
    Public RunStamp = Format(Now, "yyyyMMdd_HHmmss")
    Public FileName As String = "LOG_" + RunStamp + ".txt"


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim Rec As String = ""
        Dim N As Double = 0

        FileOpen(13, FileName, OpenMode.Append)

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
        Counts(17, 0) = "Terminally Infected"
        Counts(18, 0) = "Infected Males"
        Counts(19, 0) = "Infected Females"

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
        Counts(17, 1) = "0"
        Counts(18, 1) = "0"
        Counts(19, 1) = "0"

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
        Label18.Text = Counts(17, 0) + StrDup(30 - Len(Counts(17, 0)), ".")
        Label19.Text = Counts(18, 0) + StrDup(25 - Len(Counts(18, 0)), ".")
        Label20.Text = Counts(19, 0) + StrDup(25 - Len(Counts(19, 0)), ".")

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
        Label38.Text = StrDup(7 - Len(Counts(17, 1)), " ") + Counts(17, 1)
        Label39.Text = StrDup(7 - Len(Counts(18, 1)), " ") + Counts(18, 1)
        Label40.Text = StrDup(7 - Len(Counts(19, 1)), " ") + Counts(19, 1)

        FileOpen(1, "Male.First.txt", OpenMode.Input)
        While Not EOF(1)
            Rec = LineInput(1)
            MFname(N) = Trim(Mid(Rec, 1, InStr(Rec, " ") - 1))
            N = N + 1
            Application.DoEvents()
        End While
        FileClose(1)

        N = 0
        FileOpen(2, "Female.First.txt", OpenMode.Input)
        While Not EOF(2)
            Rec = LineInput(2)
            FFname(N) = Trim(Mid(Rec, 1, InStr(Rec, " ") - 1))
            N = N + 1
            Application.DoEvents()
        End While
        FileClose(2)

        N = 0
        FileOpen(3, "Last.txt", OpenMode.Input)
        While Not EOF(3)
            Rec = LineInput(3)
            Lname(N) = Trim(Mid(Rec, 1, InStr(Rec, " ") - 1))
            N = N + 1
            Application.DoEvents()
        End While
        FileClose(3)


        GFX.FillRectangle(Brushes.Black, 0, 0, 1000, 800)
        PictureBox1.Image = BMP
        lblModSeed.Text = Format(HScrollBar1.Value, "#,##0")
        Label59.Text = Format(HScrollBar2.Value, "#,##0") + " Maximum Units of Age"
        HScrollBar3.Maximum = HScrollBar2.Value
        HScrollBar4.Maximum = HScrollBar2.Value
        HScrollBar5.Maximum = HScrollBar2.Value
        HScrollBar6.Maximum = HScrollBar2.Value
        lblMaleChild.Text = "From 0 to " + Format(HScrollBar3.Value, "#,##0")
        lblMaleFertile.Text = "From " + Format(HScrollBar3.Value + 1, "#,##0") + " to " + Format(HScrollBar4.Value - 1, "#,##0")
        lblMalePostFertile.Text = "From " + Format(HScrollBar4.Value, "#,##0") + " to " + Format(HScrollBar2.Value, "#,##0")
        lblFemaleChild.Text = "From 0 to " + Format(HScrollBar5.Value, "#,##0")
        lblFemaleFertile.Text = "From " + Format(HScrollBar5.Value + 1, "#,##0") + " to " + Format(HScrollBar6.Value - 1, "#,##0")
        lblFemalePostFertile.Text = "From " + Format(HScrollBar6.Value, "#,##0") + " to " + Format(HScrollBar2.Value, "#,##0")

        Age(0).MaleChild = 0
        Age(1).MaleChild = HScrollBar3.Value
        Age(0).MaleFertile = HScrollBar3.Value + 1
        Age(1).MaleFertile = HScrollBar4.Value - 1
        Age(0).MalePostFertile = HScrollBar4.Value
        Age(1).MalePostFertile = HScrollBar2.Value
        Age(0).FemaleChild = 0
        Age(1).FemaleChild = HScrollBar5.Value
        Age(0).FemaleFertile = HScrollBar5.Value + 1
        Age(1).FemaleFertile = HScrollBar6.Value - 1
        Age(0).FemalePostFertile = HScrollBar6.Value
        Age(1).FemalePostFertile = HScrollBar2.Value

    End Sub

    Public Function CreateUniversalLife(LastLeft As Single, LastTop As Single) As String

        CreateUniversalLife = 0
        Dim PercentMales As Double = 40
        Dim PercentFemales As Double = 60
        Dim PercentChild As Double = 33.2
        Dim PercentFertile As Double = 33.4
        Dim PercentMenopausal As Double = 33.4
        Dim ImgFile As String = ""


        Dim Seed1 As Single = Val(Mid(Format(Now, "yyyyMMddHHmmsss"), 15, 1))
        Dim Seed2 As Single = Val(Mid(Format(Now, "yyyyMMddHHmmsss"), 14, 1))
        Dim Seed3 As Single = Val(Mid(Format(Now, "yyyyMMddHHmmsss"), 14, 2))

        'RAND1 = Life or No Life
        Randomize()
        Dim Rand1 As Integer = CInt(Math.Floor((2 - 1 + 1) * Rnd())) + 1

        'RAND2 = Male or Female
        Randomize()
        Dim Rand2 As Integer = CInt(Math.Floor((2 - 1 + 1) * Rnd())) + 1

        'RAND3 = Age of Life
        Randomize()
        Dim Rand3 As Integer = CInt(Math.Floor((Age(1).MaxAge - 1 + 1) * Rnd())) + 1

        'MsgBox("(" + Str(Rand1) + "," + Str(Rand2) + "," + Str(Rand3) + ")")

        If Rand1 = 1 Then
            CreateUniversalLife = "EMPTY"
            ImgFile = "BLANK.bmp"
            Universe(LastLeft, LastTop).PersonType = "BLANK"
            Universe(LastLeft, LastTop).CurrentLeft = Trim(Str(LastLeft))
            Universe(LastLeft, LastTop).CurrentTop = Trim(Str(LastTop))
            GoTo DRAW_IT
        End If

        PersonNumberSeq = PersonNumberSeq + 1

        If Rand2 = 1 Then
            If Rand3 >= Age(0).MaleChild And Rand3 <= Age(1).MaleChild Then
                CreateUniversalLife = "MCW"
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
                Universe(LastLeft, LastTop).Sick = False
                Universe(LastLeft, LastTop).SickContagiousness = 0
                Universe(LastLeft, LastTop).SickIncubation = 0
                Universe(LastLeft, LastTop).PersonNumber = "P-" + Format(PersonNumberSeq, "00000")
                Universe(LastLeft, LastTop).CurrentLeft = Trim(Str(LastLeft))
                Universe(LastLeft, LastTop).CurrentTop = Trim(Str(LastTop))
                GoTo DRAW_IT
            End If
            If Rand3 >= Age(0).MaleFertile And Rand3 <= Age(1).MaleFertile Then
                CreateUniversalLife = "MFW"
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
                Universe(LastLeft, LastTop).Sick = False
                Universe(LastLeft, LastTop).SickContagiousness = 0
                Universe(LastLeft, LastTop).SickIncubation = 0
                Universe(LastLeft, LastTop).PersonNumber = "P-" + Format(PersonNumberSeq, "00000")
                Universe(LastLeft, LastTop).CurrentLeft = Trim(Str(LastLeft))
                Universe(LastLeft, LastTop).CurrentTop = Trim(Str(LastTop))
                GoTo DRAW_IT
            End If
            If Rand3 >= Age(0).MalePostFertile And Rand3 <= Age(1).MalePostFertile Then
                CreateUniversalLife = "MMW"
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
                Universe(LastLeft, LastTop).Sick = False
                Universe(LastLeft, LastTop).SickContagiousness = 0
                Universe(LastLeft, LastTop).SickIncubation = 0
                Universe(LastLeft, LastTop).PersonNumber = "P-" + Format(PersonNumberSeq, "00000")
                Universe(LastLeft, LastTop).CurrentLeft = Trim(Str(LastLeft))
                Universe(LastLeft, LastTop).CurrentTop = Trim(Str(LastTop))
                GoTo DRAW_IT
            End If
        End If
        If Rand2 = 2 Then
            If Rand3 >= Age(0).FemaleChild And Rand3 <= Age(1).FemaleChild Then
                CreateUniversalLife = "FCW"
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
                Universe(LastLeft, LastTop).Sick = False
                Universe(LastLeft, LastTop).SickContagiousness = 0
                Universe(LastLeft, LastTop).SickIncubation = 0
                Universe(LastLeft, LastTop).PersonNumber = "P-" + Format(PersonNumberSeq, "00000")
                Universe(LastLeft, LastTop).CurrentLeft = Trim(Str(LastLeft))
                Universe(LastLeft, LastTop).CurrentTop = Trim(Str(LastTop))
                GoTo DRAW_IT
            End If
            If Rand3 >= Age(0).FemaleFertile And Rand3 <= Age(1).FemaleFertile Then
                CreateUniversalLife = "FFW"
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
                Universe(LastLeft, LastTop).Sick = False
                Universe(LastLeft, LastTop).SickContagiousness = 0
                Universe(LastLeft, LastTop).SickIncubation = 0
                Universe(LastLeft, LastTop).PersonNumber = "P-" + Format(PersonNumberSeq, "00000")
                Universe(LastLeft, LastTop).CurrentLeft = Trim(Str(LastLeft))
                Universe(LastLeft, LastTop).CurrentTop = Trim(Str(LastTop))
                GoTo DRAW_IT
            End If
            If Rand3 >= Age(0).FemalePostFertile And Rand3 <= Age(1).FemalePostFertile Then
                CreateUniversalLife = "FMW"
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
                Universe(LastLeft, LastTop).Sick = False
                Universe(LastLeft, LastTop).SickContagiousness = 0
                Universe(LastLeft, LastTop).SickIncubation = 0
                Universe(LastLeft, LastTop).PersonNumber = "P-" + Format(PersonNumberSeq, "00000")
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
                  Universe(LastLeft, LastTop).Lname + " (Person#: " + Universe(LastLeft, LastTop).PersonNumber + ") . . .")
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
            CreateUniversalLife(LastLeft, LastTop)
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
        Button3.Enabled = True

    End Sub

    Public Function ToLog(Type As String, Rec As String)

        ToLog = 0
        Print(13, Format(Now, "yyyy-MM-dd HH:mm:sss") + "|" + Type + "|" + Rec + vbCrLf)
        ListBox1.Items.Add(Type + "|" + Rec)
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
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

    Private Sub HScrollBar1_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar1.Scroll

        lblModSeed.Text = HScrollBar1.Value

    End Sub

    Private Sub HScrollBar2_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar2.Scroll

        Label59.Text = Format(HScrollBar2.Value, "#,##0") + " Maximum Units of Age"
        HScrollBar3.Maximum = HScrollBar2.Value
        HScrollBar4.Maximum = HScrollBar2.Value
        lblMaleChild.Text = "From 0 to " + Format(HScrollBar3.Value, "#,##0")
        lblMaleFertile.Text = "From " + Format(HScrollBar3.Value + 1, "#,##0") + " to " + Format(HScrollBar4.Value - 1, "#,##0")
        lblMalePostFertile.Text = "From " + Format(HScrollBar4.Value, "#,##0") + " to " + Format(HScrollBar2.Value, "#,##0")
        HScrollBar5.Maximum = HScrollBar2.Value
        HScrollBar6.Maximum = HScrollBar2.Value
        lblFemaleChild.Text = "From 0 to " + Format(HScrollBar5.Value, "#,##0")
        lblFemaleFertile.Text = "From " + Format(HScrollBar5.Value + 1, "#,##0") + " to " + Format(HScrollBar6.Value - 1, "#,##0")
        lblFemalePostFertile.Text = "From " + Format(HScrollBar6.Value, "#,##0") + " to " + Format(HScrollBar2.Value, "#,##0")


    End Sub

    Private Sub HScrollBar3_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar3.Scroll

Check_Again:
        If HScrollBar4.Value - HScrollBar3.Value < 3 Then
            If HScrollBar4.Value < HScrollBar4.Maximum Then
                HScrollBar4.Value = HScrollBar4.Value + 1
                GoTo Check_Again
            End If
        End If
        lblMaleChild.Text = "From 0 to " + Format(HScrollBar3.Value, "#,##0")
        lblMaleFertile.Text = "From " + Format(HScrollBar3.Value + 1, "#,##0") + " to " + Format(HScrollBar4.Value - 1, "#,##0")
        lblMalePostFertile.Text = "From " + Format(HScrollBar4.Value, "#,##0") + " to " + Format(HScrollBar2.Value, "#,##0")

    End Sub

    Private Sub HScrollBar4_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar4.Scroll

Check_Again:
        If HScrollBar4.Value - HScrollBar3.Value < 3 Then
            HScrollBar3.Value = HScrollBar3.Value - 1
            GoTo Check_Again
        End If
        lblMaleChild.Text = "From 0 to " + Format(HScrollBar3.Value, "#,##0")
        lblMaleFertile.Text = "From " + Format(HScrollBar3.Value + 1, "#,##0") + " to " + Format(HScrollBar4.Value - 1, "#,##0")
        lblMalePostFertile.Text = "From " + Format(HScrollBar4.Value, "#,##0") + " to " + Format(HScrollBar2.Value, "#,##0")

    End Sub

    Private Sub HScrollBar5_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar5.Scroll

Check_Again:
        If HScrollBar6.Value - HScrollBar5.Value < 3 Then
            If HScrollBar6.Value < HScrollBar6.Maximum Then
                HScrollBar6.Value = HScrollBar6.Value + 1
                GoTo Check_Again
            End If
        End If
        lblFemaleChild.Text = "From 0 to " + Format(HScrollBar5.Value, "#,##0")
        lblFemaleFertile.Text = "From " + Format(HScrollBar5.Value + 1, "#,##0") + " to " + Format(HScrollBar6.Value - 1, "#,##0")
        lblFemalePostFertile.Text = "From " + Format(HScrollBar6.Value, "#,##0") + " to " + Format(HScrollBar2.Value, "#,##0")

    End Sub

    Private Sub HScrollBar6_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar6.Scroll

Check_Again:
        If HScrollBar6.Value - HScrollBar5.Value < 3 Then
            HScrollBar5.Value = HScrollBar5.Value - 1
            GoTo Check_Again
        End If
        lblFemaleChild.Text = "From 0 to " + Format(HScrollBar5.Value, "#,##0")
        lblFemaleFertile.Text = "From " + Format(HScrollBar5.Value + 1, "#,##0") + " to " + Format(HScrollBar6.Value - 1, "#,##0")
        lblFemalePostFertile.Text = "From " + Format(HScrollBar6.Value, "#,##0") + " to " + Format(HScrollBar2.Value, "#,##0")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        LoadConfig()
        Button2.Enabled = False
        If InitialPopulate = False Then
            Button1.Enabled = True
            InitialPopulate = True
        End If
        If CheckBox1.Checked = True Then
            Label48.Text = "The Saved Configuration will be used for year " + Counts(6, 1) + "."
        Else
            Label48.Text = "The Saved Configuration will be used from year " + Counts(6, 1) + " indefinately..."
        End If

    End Sub

    Public Function LoadConfig()

        LoadConfig = 0
        Age(0).MaxAge = 0
        Age(1).MaxAge = HScrollBar2.Value
        Age(0).MaleChild = 0
        Age(1).MaleChild = HScrollBar3.Value
        Age(0).MaleFertile = HScrollBar3.Value + 1
        Age(1).MaleFertile = HScrollBar4.Value - 1
        Age(0).MalePostFertile = HScrollBar4.Value
        Age(1).MalePostFertile = HScrollBar2.Value
        Age(0).FemaleChild = 0
        Age(1).FemaleChild = HScrollBar5.Value
        Age(0).FemaleFertile = HScrollBar5.Value + 1
        Age(1).FemaleFertile = HScrollBar6.Value - 1
        Age(0).FemalePostFertile = HScrollBar6.Value
        Age(1).FemalePostFertile = HScrollBar2.Value

        If lblPairsForLife.Checked = True Then
            Cfg.PairsForLife = True
        Else
            Cfg.PairsForLife = False
        End If
        If lblOffspringTouchParent.Checked = True Then
            Cfg.OffspringTouchesParent = True
        Else
            Cfg.OffspringTouchesParent = False
        End If
        If lblPreFertileNoMove.Checked = True Then
            Cfg.PreFertileDoesNotMove = True
        Else
            Cfg.PreFertileDoesNotMove = False
        End If
        If lblPostFertileNoMove.Checked = True Then
            Cfg.PostFertileDoesNotMove = True
        Else
            Cfg.PostFertileDoesNotMove = False
        End If
        If lblFertileMovesToFertileSpot.Checked = True Then
            Cfg.NonPairedMovesToFertileSpot = True
        Else
            Cfg.NonPairedMovesToFertileSpot = False
        End If
        If lblFertileMovesToBlankSpot.Checked = True Then
            Cfg.NonPairedMovesToNextSpot = True
        Else
            Cfg.NonPairedMovesToNextSpot = False
        End If
        If lblAccidentalDeath.Checked = True Then
            Cfg.AccidentalDeath = True
        Else
            Cfg.AccidentalDeath = False
        End If

        Sick.IntroduceDisease = lblIntroduceDisease.Checked
        Sick.AffectsMales = lblAffectsMales.Checked
        Sick.AffectsFemales = lblAffectsFemales.Checked
        Sick.AffectsChildren = lblAffectsChildren.Checked
        Sick.AffectsFertile = lblAffectsFertile.Checked
        Sick.AffectsElderly = lblAffectsElderly.Checked
        Sick.ModSeed = Val(lblModSeed.Text)
        Sick.LivesInitiallyInfected = cmbLivesAffected.SelectedIndex + 1
        Sick.IncubationPeriod = cmbIncubation.SelectedIndex + 1

        Select Case cmbContagiousness.SelectedIndex
            Case 0
                Sick.Contagiosness = 0.25
            Case 1
                Sick.Contagiosness = 0.5
            Case 2
                Sick.Contagiosness = 0.75
            Case 3
                Sick.Contagiosness = 0.9
            Case Else
                Sick.Contagiosness = 0.25
        End Select

    End Function

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked = True Then
            Button3.Text = "STEP 3:  Execute a Year's Worth of Life"
        Else
            Button3.Text = "STEP 3:  Execute Life Indefinately"
        End If
        If CheckBox1.Checked = True Then
            If Label48.Text <> "Press Button to Save Config" Then
                Label48.Text = "The Saved Configuration will be used for year " + Counts(6, 1) + "."
            End If
        Else
            If Label48.Text <> "Press Button to Save Config" Then
                Label48.Text = "The Saved Configuration will be used from year " + Counts(6, 1) + " indefinately..."
            End If
        End If

    End Sub

End Class
