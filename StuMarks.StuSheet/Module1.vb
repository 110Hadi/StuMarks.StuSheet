Module Module1

    Dim grade, Subj As String

    Sub heading()

        Console.WriteLine("/////////////////////////////////////////////////MARKSHEET\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\")
        Console.WriteLine("_________________________________________________________________________________________________________")
        Console.WriteLine()
        Console.WriteLine()

    End Sub

    Sub ColumnC(ByVal Subj As String)

        Console.WriteLine("________________________________________________________________________________________________________________")
        Console.WriteLine("Roll No.   Name      " & Subj & "Total     Average   Grade     Status    Group")

    End Sub

    Sub gradeC(ByVal FormattedAverage As Integer)

        Select Case FormattedAverage

            Case Is > 90
                grade = "A*"

            Case Is > 80
                grade = "A"

            Case Is > 70
                grade = "B"

            Case Is > 60
                grade = "C"

            Case Is > 50
                grade = "D"

            Case Is > 45
                grade = "E"

            Case Else
                grade = "U"

        End Select

    End Sub



    Sub Main()

        ''DECLARING VARIABLES
        Dim NumStu, NumSub, total, marks, output, counter, index As Integer
        Dim name, status, group, choice, Search As String
        Dim average As Decimal
        Const Columns As Integer = 6
        Subj = ""

        Call heading()

        ''INITIAL INPUTS + VALIDATION

        Do
            Console.Write("Enter The Number Of Students in Your Class: ")
            NumStu = Console.ReadLine

            If NumStu < 3 Then
                Console.WriteLine("Invalid Response. Too Few Students.")
            End If
            Console.WriteLine()
        Loop Until NumStu >= 3

        Do
            Console.Write("Enter The Number Of Subjects for Your Class: ")
            NumSub = Console.ReadLine

            If NumSub < 3 Then
                Console.WriteLine("Invalid Response. Too Few Subjects.")
            End If
            Console.WriteLine()
        Loop Until NumSub >= 3
        Console.WriteLine()

        ''DECLARING ARRAYS
        Dim Subjects(NumSub + 1), Marksheet(NumStu, NumSub + Columns) As String

        ''NAMES OF SUBJECTS

        For counter = 2 To NumSub + 1

            Do
                Console.Write("Enter Subject No." & counter - 1 & " Name: ")
                Subjects(counter) = Console.ReadLine

                If IsNumeric(Subjects(counter)) Then
                    Console.WriteLine("Invalid Response. Subject Name Must Be String.")
                End If
                Console.WriteLine()
            Loop Until Not IsNumeric(Subjects(counter))
            Subj += Subjects(counter) & vbTab & vbTab

        Next
        Console.Clear()


        ''USER INPUT & VALIDATION
        Call heading()

        For row = 1 To NumStu
            total = 0
            Console.Write("Enter Name of Roll No." & row & " : ")
            name = Console.ReadLine

            Marksheet(row, 1) = name & vbTab & "    "

            ''ENTERING SUBJET MARKS
            For counter = 2 To NumSub + 1

                Do
                    Console.Write("Enter Roll No." & row & " " & Subjects(counter) & " Marks: ")
                    marks = Console.ReadLine

                    If marks < 0 Or marks > 100 Then
                        Console.WriteLine("Invalid Response. Enter Marks Between 0 and 100 Inclusive.")
                    End If
                    Console.WriteLine()
                Loop Until marks >= 0 And marks <= 100

                Marksheet(row, counter) = marks & vbTab & vbTab
                total += marks
                Marksheet(row, NumSub + 2) = total & vbTab & "   "
            Next


            ''FURTHER PROCESSING
            average = (CInt(Marksheet(row, NumSub + 2)) / (NumSub * 100)) * 100
            average = Math.Round(average, 2)
            Dim FormattedAverage As String = average.ToString("0.00")


            Call gradeC(FormattedAverage)

            If grade = "U" Then
                status = "Fail"
            Else
                status = "Pass"
            End If

            If FormattedAverage >= 75 Then
                group = "Science"
            Else
                group = "Commerce"
            End If

            Marksheet(row, 0) = row & vbTab & "     "

            Marksheet(row, NumSub + 3) = FormattedAverage & "    "
            Marksheet(row, NumSub + 4) = grade & vbTab & "      "
            Marksheet(row, NumSub + 5) = status & vbTab
            Marksheet(row, NumSub + 6) = group & vbTab

        Next

        ''OUTPUT
        Console.Clear()
        Call heading()

        Console.WriteLine(":::::OUTPUT:::::")
        Console.WriteLine("________________")
        Console.WriteLine()

        ''Asking For Result
        Do
            Console.Write("Do You Want Any Type of Output? ")
            choice = UCase(Console.ReadLine)

            If choice <> "Y" And choice <> "N" Then
                Console.WriteLine("Invalid Response. Enter a 'Y' or 'N'.")
            End If
            Console.WriteLine()
        Loop Until choice = "Y" Or choice = "N"

        ''OUTPUTS
        While choice = "Y"
            Console.WriteLine("What Type Of Output You Want? ")
            Console.WriteLine("1. Full Result")
            Console.WriteLine("2. Subject Result")
            Console.WriteLine("3. Student Result")

            Do
                Console.Write("Enter Your Choice: ")
                output = Console.ReadLine

                If output < 1 Or output > 3 Then
                    Console.WriteLine("Invalid Response. Chose From The Given Options.")
                End If
                Console.WriteLine()
            Loop Until output >= 1 And output <= 3
            Console.Clear()
            Call heading()

            ''Choice for Full Result Display
            Select Case output

                Case 1
                    Console.WriteLine("Full Result: ")
                    Console.WriteLine("_____________")
                    Console.WriteLine()

                    Call ColumnC(Subj)

                    For row = 1 To NumStu

                        For column = 0 To NumSub + Columns
                            Console.Write(Marksheet(row, column))
                        Next
                        Console.WriteLine()

                    Next

                    ''Subjects
                Case 2
                    Console.WriteLine("Subject Result: ")
                    Console.WriteLine("________________")
                    Console.WriteLine()

                    index = 0
                    Do
                        Console.Write("Which Subject You Want The Result Of? ")
                        Search = Console.ReadLine

                        For counter = 1 To NumSub
                            If LCase(Search) = LCase(Subjects(counter)) Then
                                index = counter
                            End If
                        Next

                        If index = 0 Then
                            Console.WriteLine("Invalid Response. Subject Not Found.")
                        End If
                        Console.WriteLine()
                    Loop Until index <> 0
                    Console.Clear()

                    Console.WriteLine("Roll No.    Name     " & Subjects(index))
                    For row = 1 To NumStu

                        For column = 0 To 1
                            Console.Write(Marksheet(row, column))
                        Next
                        Console.Write(Marksheet(row, index + 1))
                        Console.WriteLine()

                    Next

                Case Else
                    ''Student Result
                    Console.WriteLine("Student Result: ")
                    Console.WriteLine("________________")
                    Console.WriteLine()

                    index = 0
                    Do
                        Console.Write("Which Student(Roll No.) You Want The Result Of? ")
                        Search = CInt(Console.ReadLine)

                        If Search > NumStu Then
                            Console.WriteLine("Invalid Response. Student Not Found.")
                        End If
                        Console.WriteLine()

                    Loop Until Search < NumStu
                    Console.Clear()

                    Call ColumnC(Subj)

                    For column = 0 To NumSub + Columns
                        Console.Write(Marksheet(Search, column))
                    Next
                    Console.WriteLine()

            End Select

            Do
                Console.Write("Do You Want Any Type of Output? ")
                choice = UCase(Console.ReadLine)

                If choice <> "Y" And choice <> "N" Then
                    Console.WriteLine("Invalid Response. Enter a 'Y' or 'N'.")
                End If
                Console.WriteLine()
            Loop Until choice = "Y" Or choice = "N"

        End While

        Console.ReadLine()
    End Sub

End Module
