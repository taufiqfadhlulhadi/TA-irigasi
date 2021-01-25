Imports System.IO
Imports System.Data.OleDb
Imports System.Math
Imports System.Drawing.KnownColor

Public Class Fungsii
    Dim connect As OleDbConnection
    Dim comand As OleDbCommand
    Dim read As OleDbDataReader
    Dim kebutuhan(2) As Decimal
    Dim Qblock(2) As Decimal
    Dim Qin(2) As Decimal
    Dim Qout(2) As Decimal
    Dim Qeff(2) As Decimal

    Dim kebutuhanHilir(2) As Decimal
    Dim QblockHilir(2) As Decimal
    Dim QinHilir(2) As Decimal
    Dim QoutHilir(2) As Decimal
    Dim QeffHilir(2) As Decimal

    Dim kebutuhanTengah(2) As Decimal
    Dim QblockTengah(2) As Decimal
    Dim QinTengah(2) As Decimal
    Dim QoutTengah(2) As Decimal
    Dim QeffTengah(2) As Decimal

    Public Sub Open_Excel(open_dialog As OpenFileDialog)
        open_dialog.Filter = "Excel File|*.xlsx;*.xls"
        Dim dir_destiny As String = Application.StartupPath

        If open_dialog.ShowDialog() = DialogResult.OK Then
            Dim dir_source As String = Path.GetFullPath(open_dialog.FileName)
            Dim fileName = Path.GetFileName(open_dialog.FileName)
            'MsgBox(dir_source)

            If File.Exists(dir_destiny & "\" & fileName) Then
                'MsgBox("file sudah ada")
                File.Delete(dir_destiny & "\" & fileName)
                File.Copy(dir_source, dir_destiny & "\" & fileName)
            Else
                'MsgBox("file belum ada")
                File.Copy(dir_source, dir_destiny & "\" & fileName)
            End If
        End If

        If File.Exists(dir_destiny & "\" & Path.GetFileName(open_dialog.FileName)) Then
            connect = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dir_destiny & "\" & Path.GetFileName(open_dialog.FileName) & ";Extended Properties=Excel 12.0;")
            comand = New OleDbCommand()

            Try
                connect.Open()
                With comand
                    .Connection = connect
                    .CommandText = "select * from [HULU$]"
                    .ExecuteNonQuery()
                    read = .ExecuteReader
                End With

                While read.Read
                    Try
                        If Trim(read(0)) = "A" Then
                            Qin(0) = read(1)
                            kebutuhan(0) = read(2)
                            Qblock(0) = read(3)
                            Qout(0) = read(4)
                            Qeff(0) = read(5)
                        End If

                        If Trim(read(0)) = "B" Then
                            Qin(1) = read(1)
                            kebutuhan(1) = read(2)
                            Qblock(1) = read(3)
                            Qout(1) = read(4)
                            Qeff(1) = read(5)
                        End If

                        If Trim(read(0)) = "C" Then
                            kebutuhan(2) = read(2)
                            Qblock(2) = read(3)
                            Qeff(2) = read(5)
                        End If
                    Catch ex As Exception
                    End Try
                End While
                connect.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            connect.Close()
            Try
                connect.Open()
                With comand
                    .Connection = connect
                    .CommandText = "select * from [HILIR$]"
                    .ExecuteNonQuery()
                    read = .ExecuteReader
                End With

                While read.Read
                    Try
                        If Trim(read(0)) = "X" Then
                            QinHilir(0) = read(1)
                            kebutuhanHilir(0) = read(2)
                            QblockHilir(0) = read(3)
                            QoutHilir(0) = read(4)
                            QeffHilir(0) = read(5)
                        End If

                        If Trim(read(0)) = "Y" Then
                            kebutuhanHilir(1) = read(2)
                            QblockHilir(1) = read(3)
                            QeffHilir(1) = read(5)
                        End If

                        If Trim(read(0)) = "Z" Then
                            kebutuhanHilir(2) = read(2)
                            QblockHilir(2) = read(3)
                            QeffHilir(2) = read(5)
                        End If
                    Catch exe As Exception
                    End Try
                End While

                connect.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
                connect.Close()
            End Try

            Try
                connect.Open()
                With comand
                    .Connection = connect
                    .CommandText = "select * from [TENGAH$]"
                    .ExecuteNonQuery()
                    read = .ExecuteReader
                End With

                While read.Read
                    Try
                        If Trim(read(0)) = "M" Then
                            QinTengah(0) = read(1)
                            kebutuhanTengah(0) = read(2)
                            QblockTengah(0) = read(3)
                            QoutTengah(0) = read(4)
                            QeffTengah(0) = read(5)
                        End If

                        If Trim(read(0)) = "N" Then
                            QinTengah(1) = read(1)
                            kebutuhanTengah(1) = read(2)
                            QblockTengah(1) = read(3)
                            QoutTengah(1) = read(4)
                            QeffTengah(1) = read(5)
                        End If

                        If Trim(read(0)) = "O" Then
                            kebutuhanTengah(2) = read(2)
                            QblockTengah(2) = read(3)
                            QeffTengah(2) = read(5)
                        End If
                    Catch ex As Exception

                    End Try
                End While
                connect.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
                connect.Close()
            End Try
        End If
    End Sub

    Public Sub Hulu(listview As ListView, picture1 As PictureBox, picture2 As PictureBox, picture3 As PictureBox, label1 As Label, label2 As Label, label3 As Label, panel1 As Panel, panel2 As Panel, panel3 As Panel)
        picture1.BackColor = Color.Red
        picture2.BackColor = Color.Red
        picture2.BackColor = Color.Red

        label1.Show()
        label2.Show()
        label3.Show()

        Dim hari(3) As Integer
        Dim hari_tambah As Integer = -1
        Dim kebutuhanPerJam(2) As Decimal
        Dim lama(2) As Decimal
        Dim efektifitas(2) As Decimal
        For x As Integer = 0 To kebutuhan.Length - 1
            kebutuhanPerJam(x) = ((11.84083 * kebutuhan(x)) - 341.67) / 24
            'MsgBox(kebutuhanPerJam(x))
        Next
        Dim sisa As Decimal = 0
        Dim j As Integer = 0
        Dim panelS() As Panel = {panel1, panel2, panel3}
        Dim labelS() As Label = {label1, label2, label3}
        Dim pictureboxs() As PictureBox = {picture1, picture2, picture3}
        Dim A, B, C As Decimal
        Dim totalLama As Decimal = 0
        For i As Integer = 0 To Qblock.Length - 1
            pictureboxs(i).BackColor = Color.Turquoise
            lama(i) = kebutuhanPerJam(i) / (Qblock(i) * 3600)
            labelS(i).Text = lama(i).ToString
            panelS(i).BackColor = Color.Lime
            pictureboxs(i).BackColor = Color.Red
            totalLama += lama(i)
            'MsgBox(kebutuhanPerJam(i).ToString + " / ( " + Qblock(i).ToString + " * 3600)")
        Next

        'AB-A
        Dim AB(2) As Decimal
        Dim waktuTerbesar As Decimal
        'AB(0) merupakan efisiensi AB
        AB(0) = ((Qout(0) + ((Qblock(0) * Qeff(0)) + (Qblock(1) * Qeff(1)))) * 100 / Qin(0))
        'MsgBox(AB(0).ToString + "     AB")
        C = (((Qout(1) + (Qblock(2) * Qeff(2))) * 100) / Qin(1))
        'MsgBox(C.ToString + "     C")
        'AB(1) merupakan efisiensi total AB-A
        AB(1) = (AB(0) * C) / 100
        'MsgBox(AB(1).ToString + "     totalAB")
        'AB(2) merupakan total waktu AB-A
        If (lama(0) > lama(1)) Then
            waktuTerbesar = lama(0)
        Else
            waktuTerbesar = lama(1)
        End If
        AB(2) = waktuTerbesar + lama(2)

        'pengosongan variable bersama
        waktuTerbesar = 0
        C = 0

        'A-B-C
        A = ((((Qout(0) + Qblock(1)) + (Qblock(0) * Qeff(0))) * 100) / Qin(0))
        B = ((((Qout(0) + Qblock(0)) + (Qblock(1) * Qeff(1))) * 100) / Qin(0))
        C = (((Qout(1) + (Qblock(2)) * Qeff(2)) * 100) / Qin(1))
        Dim efftotal = (A * B * C) / 10000
        'MsgBox(A.ToString + "     " + B.ToString + "     " + C.ToString)

        'pengosongan variable bersama
        A = 0
        B = 0
        C = 0

        'A-BC
        Dim BC(3) As Decimal
        A = ((Qout(0) + (Qblock(0) * Qeff(0))) * 100 / Qin(0))
        'MsgBox(A.ToString + "     A")
        'BC(0) merupakan effisiensi BC
        BC(0) = (((Qout(1) + Qblock(0)) + ((Qblock(1) * Qeff(1)) + (Qblock(2) * Qeff(2)))) * 100 / Qin(0))
        'MsgBox(BC(0).ToString + "     BC")
        'BC(1) merupakan efisiensi total A-BC
        BC(1) = (A * BC(0)) / 100
        'MsgBox(BC(1).ToString + "     BC")
        'BC(2) merupakan total waktu A-BC
        If (lama(1) > lama(2)) Then
            waktuTerbesar = lama(1)
        Else
            waktuTerbesar = lama(2)
        End If
        BC(2) = lama(0) + waktuTerbesar

        'pengosongan variable bersama
        A = 0
        waktuTerbesar = 0

        'BC-A
        'BC(3) merupakan total BC-A
        A = (((Qout(0) + Qblock(1)) + (Qblock(0) * Qeff(0))) * 100 / Qin(0))
        BC(3) = (BC(0) * A) / 100

        'AC-B
        Dim AC(3) As Decimal
        'AC(0) merupakan effisiensi AC
        AC(0) = (((Qout(1) + Qblock(1)) + ((Qblock(0) * Qeff(0)) + (Qblock(2) * Qeff(2)))) * 100 / Qin(0))
        'MsgBox(AC(0).ToString + "     AC")
        B = (((Qout(0) + Qblock(0)) + (Qblock(1) * Qeff(1))) * 100 / Qin(0))
        'MsgBox(B.ToString + "     B")
        'AC(1) merupakan efisiensi total AC-B
        AC(1) = (AC(0) * B) / 100
        'MsgBox(AC(1).ToString + "     AC")
        'AC(2) merupakan total waktu AC-B
        If (lama(0) > lama(2)) Then
            AC(2) = lama(0) + lama(1)
        Else
            AC(2) = lama(1) + lama(2)
        End If

        'pengosongan variable bersama
        B = 0

        'ABC
        Dim ABC(2) As Decimal
        'ABC(0) merupakan efisiensi total
        ABC(0) = (((Qout(1) + ((Qblock(0) * Qeff(0)) + (Qblock(1) * Qeff(1)) + (Qblock(2) * Qeff(2))))) * 100 / Qin(0))
        'MsgBox(ABC(0).ToString + "     ABC")
        'ABC(1) merupakan waktu total
        If (lama(0) > lama(1) And lama(0) > lama(2)) Then
            ABC(1) = lama(0)
        ElseIf (lama(1) > lama(0) And lama(1) > lama(2)) Then
            ABC(1) = lama(1)
        ElseIf (lama(2) > lama(0) And lama(2) > lama(1)) Then
            ABC(1) = lama(2)
        End If

        'Dim efisiensiAB As Double = ((Qout(1) - ((Qmasuk(0) * 0.923086) + (Qmasuk(1) * 0.7600961))) / tersedia(0)) * 100
        'MsgBox(efisiensiAB.ToString + "     AB")
        'C
        'Dim efisiensiC As Double = (Qout(2) - (Qmasuk(2) * 0.8830205) / tersedia(2)) * 100
        'MsgBox(efisiensiC.ToString + "       C")
        ''A
        'Dim efisiensiA As Double = ((Qout(1)) - (Qmasuk(0) * 0.923086) / tersedia(0)) * 100
        'MsgBox(efisiensiA.ToString + "       A")
        ''B
        'Dim efisiensiB As Double = ((Qout(1) - Qmasuk(1) * 0.7600961) / tersedia(0)) * 100
        'MsgBox(efisiensiB.ToString + "       B")
        ''BC
        'Dim efisiensiBC As Double = (((Qout(2) + Qmasuk(0)) - ((Qmasuk(1) * 0.7600961) + (Qmasuk(2) * 0.8830205))) / tersedia(0)) * 100
        'MsgBox(efisiensiBC.ToString + "       BC")
        ''AC
        'Dim efisiensiAC As Double = ((Qout(1) - ((Qmasuk(0) * 0.923086) + (Qmasuk(1) * 0.7600961))) / tersedia(0)) * 100
        'MsgBox(efisiensiAC.ToString + "        AC")
        'Dim A_B_C As Double = efisiensiA * efisiensiB * efisiensiC / 10000
        'Dim AB_C As Double = efisiensiAB * efisiensiC / 10000
        'Dim A_BC As Double = efisiensiA * efisiensiBC / 10000
        'Dim AC_B As Double = efisiensiAC * efisiensiB / 10000

        Dim input() As String
        input = {"A-B-C", Convert.ToString(Math.Round(totalLama, 4)), Convert.ToString(Math.Round(efftotal, 4))}
        insertListview(input, listview)
        Erase input

        input = {"AB-C", Convert.ToString(Math.Round(AB(2), 4)), Convert.ToString(Math.Round(AB(1), 4))}
        insertListview(input, listview)
        Erase input

        input = {"A-BC", Convert.ToString(Math.Round(BC(2), 4)), Convert.ToString(Math.Round(BC(1), 4))}
        insertListview(input, listview)
        Erase input

        input = {"AC-B", Convert.ToString(Math.Round(AC(2), 4)), Convert.ToString(Math.Round(AC(1), 4))}
        insertListview(input, listview)
        Erase input

        input = {"B-AC", Convert.ToString(Math.Round(AC(2), 4)), Convert.ToString(Math.Round(AC(1), 4))}
        insertListview(input, listview)
        Erase input

        input = {"BC-A", Convert.ToString(Math.Round(BC(2), 4)), Convert.ToString(Math.Round(BC(3), 4))}
        insertListview(input, listview)
        Erase input

        input = {"ABC", Convert.ToString(Math.Round(ABC(1), 4)), Convert.ToString(Math.Round(ABC(0), 4))}
        insertListview(input, listview)
        Erase input
    End Sub

    Public Sub tengah(listview As ListView, picture1 As PictureBox, picture2 As PictureBox, picture3 As PictureBox, label1 As Label, label2 As Label, label3 As Label, panel1 As Panel, panel2 As Panel, panel3 As Panel)
        picture1.BackColor = Color.Red
        picture2.BackColor = Color.Red
        picture2.BackColor = Color.Red

        label1.Show()
        label2.Show()
        label3.Show()

        Dim hari(3) As Integer
        Dim hari_tambah As Integer = -1
        Dim kebutuhanPerJam(2) As Decimal
        Dim lama(2) As Decimal
        Dim efektifitas(2) As Decimal
        For x As Integer = 0 To kebutuhan.Length - 1
            kebutuhanPerJam(x) = ((11.84083 * kebutuhan(x)) - 341.67) / 24
            'MsgBox(kebutuhanPerJam(x))
        Next
        Dim sisa As Decimal = 0
        Dim j As Integer = 0
        Dim panelS() As Panel = {panel1, panel2, panel3}
        Dim labelS() As Label = {label1, label2, label3}
        Dim pictureboxs() As PictureBox = {picture1, picture2, picture3}
        Dim M, N, O As Decimal
        Dim totalLama As Decimal = 0
        For i As Integer = 0 To Qblock.Length - 1
            pictureboxs(i).BackColor = Color.Turquoise
            lama(i) = kebutuhanPerJam(i) / (Qblock(i) * 3600)
            labelS(i).Text = lama(i).ToString
            panelS(i).BackColor = Color.Lime
            pictureboxs(i).BackColor = Color.Red
            totalLama += lama(i)
            'MsgBox(kebutuhanPerJam(i).ToString + " / ( " + Qblock(i).ToString + " * 3600)")
        Next

        'MN-O
        Dim MN(2) As Decimal
        Dim waktuTerbesar As Decimal
        'MN(0) merupakan efisiensi MN
        MN(0) = ((Qout(1) + Qblock(2)) + ((Qblock(0) * Qeff(0)) + (Qblock(1) * Qeff(1)))) * 100 / Qin(0)
        'MsgBox(MN(0).ToString + "     MN")
        O = ((Qout(1) + Qblock(1)) + (Qblock(2) * Qeff(2))) * 100 / Qin(1)
        'MsgBox(O.ToString + "     O")
        'MN(1) merupakan efisiensi total MN-O
        MN(1) = (MN(0) * O) / 100
        'MsgBox(MN(1).ToString + "     totalMN")
        'MN(2) merupakan total waktu MN-O
        If (lama(0) > lama(1)) Then
            waktuTerbesar = lama(0)
        Else
            waktuTerbesar = lama(1)
        End If
        MN(2) = waktuTerbesar + lama(2)

        'pengosongan variable bersama
        waktuTerbesar = 0
        O = 0

        'A-B-C
        M = ((Qout(0) + (Qblock(0) * Qeff(0))) * 100 / Qin(0))
        N = ((Qout(1) + Qblock(2)) - (Qblock(1) * Qeff(1))) * 100 / Qin(1)
        O = ((Qout(1) + Qblock(1)) - (Qblock(2) * Qeff(2))) * 100 / Qin(1)
        Dim efftotal = (M * N * O) / 10000
        'MsgBox(A.ToString + "     " + B.ToString + "     " + C.ToString)

        'pengosongan variable bersama
        M = 0
        N = 0
        O = 0

        'M-NO
        Dim NO(3) As Decimal
        M = ((Qout(0) + (Qblock(0) * Qeff(0))) * 100 / Qin(0))
        'MsgBox(M.ToString + "     M")
        'NO(0) merupakan effisiensi NO
        NO(0) = ((Qout(1)) + ((Qblock(1) * Qeff(1)) + (Qblock(2) * Qeff(2)))) * 100 / Qin(1)
        'MsgBox(NO(0).ToString + "     NO")
        'NO(1) merupakan efisiensi total M-NO
        NO(1) = (M * NO(0)) / 100
        'MsgBox(NO(1).ToString + "     NO")
        'NO(2) merupakan total waktu M-NO
        If (lama(1) > lama(2)) Then
            waktuTerbesar = lama(1)
        Else
            waktuTerbesar = lama(2)
        End If
        NO(2) = lama(0) + waktuTerbesar

        'pengosongan variable bersama
        M = 0
        waktuTerbesar = 0

        'MO-N
        Dim MO(3) As Decimal
        'MO(0) merupakan effisiensi MO
        MO(0) = ((Qout(1) + Qblock(1)) + ((Qblock(0) * Qeff(0)) + (Qblock(2) * Qeff(2)))) * 100 / Qin(0)
        'MsgBox(MO(0).ToString + "     MO")
        N = ((Qout(1) + Qblock(2)) + (Qblock(1) * Qeff(1))) * 100 / Qin(1)
        'MsgBox(N.ToString + "     N")
        'MO(1) merupakan efisiensi total MO-N
        MO(1) = (MO(0) * N) / 100
        'MsgBox(MO(1).ToString + "     MO")
        'MO(2) merupakan total waktu MO-N
        If (lama(0) > lama(2)) Then
            MO(2) = lama(0) + lama(1)
        Else
            MO(2) = lama(1) + lama(2)
        End If

        'pengosongan variable bersama
        N = 0

        'N-MO
        MO(0) = ((Qin(1) + Qblock(1)) + ((Qblock(0) * Qeff(0)) + (Qblock(2) * Qeff(2)))) * 100 / Qin(0)
        N = ((Qout(1) + Qblock(2)) + (Qblock(1) * Qeff(1))) * 100 / Qin(1)
        MO(3) = (MO(0) * N) / 100
        'MO(3) merupakan total MO-N

        'NO-M

        M = ((Qout(0) + (Qblock(0) * Qeff(0))) * 100 / Qin(0))
        NO(3) = (NO(0) * M) / 100
        'MNO
        Dim MNO(2) As Decimal
        'MNO(0) merupakan efisiensi total
        MNO(0) = ((Qout(1) + ((Qblock(0) * Qeff(0)) + (Qblock(0) * Qeff(0)) + (Qblock(2) * Qeff(2))))) * 100 / Qin(0)
        'MsgBox(MNO(0).ToString + "     MNO")
        'MNO(1) merupakan waktu total
        If (lama(0) > lama(1) And lama(0) > lama(2)) Then
            MNO(1) = lama(0)
        ElseIf (lama(1) > lama(0) And lama(1) > lama(2)) Then
            MNO(1) = lama(1)
        ElseIf (lama(2) > lama(0) And lama(2) > lama(1)) Then
            MNO(1) = lama(2)
        End If

        'Dim efisiensiAB As Double = ((Qout(1) - ((Qmasuk(0) * 0.923086) + (Qmasuk(1) * 0.7600961))) / tersedia(0)) * 100
        'MsgBox(efisiensiAB.ToString + "     AB")
        'C
        'Dim efisiensiC As Double = (Qout(2) - (Qmasuk(2) * 0.8830205) / tersedia(2)) * 100
        'MsgBox(efisiensiC.ToString + "       C")
        ''A
        'Dim efisiensiA As Double = ((Qout(1)) - (Qmasuk(0) * 0.923086) / tersedia(0)) * 100
        'MsgBox(efisiensiA.ToString + "       A")
        ''B
        'Dim efisiensiB As Double = ((Qout(1) - Qmasuk(1) * 0.7600961) / tersedia(0)) * 100
        'MsgBox(efisiensiB.ToString + "       B")
        ''BC
        'Dim efisiensiBC As Double = (((Qout(2) + Qmasuk(0)) - ((Qmasuk(1) * 0.7600961) + (Qmasuk(2) * 0.8830205))) / tersedia(0)) * 100
        'MsgBox(efisiensiBC.ToString + "       BC")
        ''AC
        'Dim efisiensiAC As Double = ((Qout(1) - ((Qmasuk(0) * 0.923086) + (Qmasuk(1) * 0.7600961))) / tersedia(0)) * 100
        'MsgBox(efisiensiAC.ToString + "        AC")
        'Dim A_B_C As Double = efisiensiA * efisiensiB * efisiensiC / 10000
        'Dim AB_C As Double = efisiensiAB * efisiensiC / 10000
        'Dim A_BC As Double = efisiensiA * efisiensiBC / 10000
        'Dim AC_B As Double = efisiensiAC * efisiensiB / 10000

        Dim input() As String
        input = {"M-N-O", Convert.ToString(Math.Round(totalLama, 4)), Convert.ToString(Math.Round(efftotal, 4))}
        insertListview(input, listview)
        Erase input

        input = {"MN-O", Convert.ToString(Math.Round(MN(2), 4)), Convert.ToString(Math.Round(MN(1), 4))}
        insertListview(input, listview)
        Erase input

        input = {"M-NO", Convert.ToString(Math.Round(NO(2), 4)), Convert.ToString(Math.Round(NO(1), 4))}
        insertListview(input, listview)
        Erase input

        input = {"MO-N", Convert.ToString(Math.Round(MO(2), 4)), Convert.ToString(Math.Round(MO(1), 4))}
        insertListview(input, listview)
        Erase input

        input = {"N-MO", Convert.ToString(Math.Round(MO(2), 4)), Convert.ToString(Math.Round(MN(1), 4))}
        insertListview(input, listview)
        Erase input

        input = {"NO-M", Convert.ToString(Math.Round(NO(2), 4)), Convert.ToString(Math.Round(NO(3), 4))}
        insertListview(input, listview)
        Erase input

        input = {"MNO", Convert.ToString(Math.Round(MNO(1), 4)), Convert.ToString(Math.Round(MNO(0), 4))}
        insertListview(input, listview)
        Erase input
    End Sub

    Public Sub Hilir(listview As ListView, picture1 As PictureBox, picture2 As PictureBox, picture3 As PictureBox, label1 As Label, label2 As Label, label3 As Label, panel1 As Panel, panel2 As Panel, panel3 As Panel)
        picture1.BackColor = Color.Red
        picture2.BackColor = Color.Red
        picture2.BackColor = Color.Red

        label1.Show()
        label2.Show()
        label3.Show()

        Dim hari(3) As Integer
        Dim hari_tambah As Integer = -1
        Dim kebutuhanPerJam(2) As Decimal
        Dim lama(2) As Decimal
        Dim efektifitas(2) As Decimal
        For b As Integer = 0 To kebutuhan.Length - 1
            kebutuhanPerJam(b) = ((11.84083 * kebutuhan(b)) - 341.67) / 24
            'MsgBox(kebutuhanPerJam(x))
        Next
        Dim sisa As Decimal = 0
        Dim j As Integer = 0
        Dim panelS() As Panel = {panel1, panel2, panel3}
        Dim labelS() As Label = {label1, label2, label3}
        Dim pictureboxs() As PictureBox = {picture1, picture2, picture3}
        Dim X, Y, Z As Decimal
        Dim totalLama As Decimal = 0
        For i As Integer = 0 To Qblock.Length - 1
            pictureboxs(i).BackColor = Color.Turquoise
            lama(i) = kebutuhanPerJam(i) / (Qblock(i) * 3600)
            labelS(i).Text = lama(i).ToString
            panelS(i).BackColor = Color.Lime
            pictureboxs(i).BackColor = Color.Red
            totalLama += lama(i)
            'MsgBox(kebutuhanPerJam(i).ToString + " / ( " + Qblock(i).ToString + " * 3600)")
        Next

        'XY-Z
        Dim XY(2) As Decimal
        Dim waktuTerbesar As Decimal
        'XY(0) merupakan efisiensi XY
        XY(0) = ((Qout(0) + Qblock(2)) + ((Qblock(1) * Qeff(1)) + (Qblock(0) * Qeff(0)))) * 100 / Qin(0)
        'MsgBox(AB(0).ToString + "     AB")
        Z = ((Qout(0) + Qblock(0) + Qblock(1)) + (Qblock(2) * Qeff(2))) * 100 / Qin(0)
        'MsgBox(Z.ToString + "     Z")
        'XY(1) merupakan efisiensi total XY-Z
        XY(1) = (XY(0) * Z) / 100
        'MsgBox(XY(1).ToString + "     totalXY")
        'XY(2) merupakan total waktu XY-Z
        If (lama(0) > lama(1)) Then
            waktuTerbesar = lama(0)
        Else
            waktuTerbesar = lama(1)
        End If
        XY(2) = waktuTerbesar + lama(2)

        'pengosongan variable bersama
        waktuTerbesar = 0
        Z = 0

        'X-Y-Z
        X = ((Qout(0) + Qblock(1) + Qblock(2)) + (Qblock(0) * Qeff(0))) * 100 / Qin(0)
        Y = ((Qout(0) + Qblock(0) + Qblock(2)) + (Qblock(1) * Qeff(1))) * 100 / Qin(0)
        Z = ((Qout(0) + Qblock(0) + Qblock(1)) + (Qblock(2) * Qeff(2))) * 100 / Qin(0)
        Dim efftotal = (X * Y * Z) / 10000
        'MsgBox(X.ToString + "     " + Y.ToString + "     " + Z.ToString)

        'pengosongan variable bersama
        X = 0
        Y = 0
        Z = 0

        'X-YZ
        Dim YZ(3) As Decimal
        X = ((Qout(0) + Qblock(1) + Qblock(2)) * (Qblock(0) * Qeff(0))) * 100 / Qin(0)
        'MsgBox(X.ToString + "     X")
        'YZ(0) merupakan effisiensi YZ
        YZ(0) = ((Qout(0) + Qblock(0)) + ((Qblock(1) * Qeff(1)) + (Qblock(2) * Qeff(2)))) * 100 / Qin(0)
        'MsgBox(YZ(0).ToString + "     YZ")
        'YZ(1) merupakan efisiensi total X-YZ
        YZ(1) = (X * YZ(0)) / 100
        'MsgBox(YZ(1).ToString + "     YZ")
        'YZ(2) merupakan total waktu X-YZ
        If (lama(1) > lama(2)) Then
            waktuTerbesar = lama(1)
        Else
            waktuTerbesar = lama(2)
        End If
        YZ(2) = lama(0) + waktuTerbesar

        'pengosongan variable bersama
        X = 0
        waktuTerbesar = 0

        'YZ-X
        'YZ(3) merupakan total YZ-X
        X = ((Qout(0) + Qblock(1) + Qblock(2)) + (Qblock(0) * Qeff(0))) * 100 / Qin(0)
        YZ(3) = (YZ(0) * X) / 100

        'XZ-Y
        Dim XZ(3) As Decimal
        'XZ(0) merupakan effisiensi XZ
        XZ(0) = ((Qout(0) + Qblock(1)) + ((Qblock(0) * Qeff(0)) + (Qblock(2) * Qeff(2)))) * 100 / Qin(0)
        'MsgBox(XZ(0).ToString + "     XZ")
        Y = ((Qout(0) + Qblock(0) + Qblock(2)) * (Qblock(1) * Qeff(1))) * 100 / Qin(0)
        'MsgBox(Y.ToString + "     Y")
        'XZ(1) merupakan efisiensi total XZ-Y
        XZ(1) = (XZ(0) * Y) / 100
        'MsgBox(XZ(1).ToString + "     XZ")
        'XZ(2) merupakan total waktu XZ-Y
        If (lama(0) > lama(2)) Then
            XZ(2) = lama(0) + lama(1)
        Else
            XZ(2) = lama(1) + lama(2)
        End If

        'pengosongan variable bersama
        Y = 0

        'XYZ
        Dim XYZ(2) As Decimal
        'XYZ(0) merupakan efisiensi total
        XYZ(0) = ((Qout(0)) + ((Qblock(0) * Qeff(0)) + (Qblock(1) * Qeff(1)) + (Qblock(2) * Qeff(2)))) * 100 / Qin(0)
        'MsgBox(XYZ(0).ToString + "     XYZ")
        'XYZ(1) merupakan waktu total
        If (lama(0) > lama(1) And lama(0) > lama(2)) Then
            XYZ(1) = lama(0)
        ElseIf (lama(1) > lama(0) And lama(1) > lama(2)) Then
            XYZ(1) = lama(1)
        ElseIf (lama(2) > lama(0) And lama(2) > lama(1)) Then
            XYZ(1) = lama(2)
        End If

        Dim input() As String
        input = {"X-Y-Z", Convert.ToString(Math.Round(totalLama, 4)), Convert.ToString(Math.Round(efftotal, 4))}
        insertListview(input, listview)
        Erase input

        input = {"XY-Z", Convert.ToString(Math.Round(XY(2), 4)), Convert.ToString(Math.Round(XY(1), 4))}
        insertListview(input, listview)
        Erase input

        input = {"X-YZ", Convert.ToString(Math.Round(YZ(2), 4)), Convert.ToString(Math.Round(YZ(1), 4))}
        insertListview(input, listview)
        Erase input

        input = {"XZ-Y", Convert.ToString(Math.Round(XZ(2), 4)), Convert.ToString(Math.Round(XZ(1), 4))}
        insertListview(input, listview)
        Erase input

        input = {"Y-XZ", Convert.ToString(Math.Round(XZ(2), 4)), Convert.ToString(Math.Round(XZ(1), 4))}
        insertListview(input, listview)
        Erase input

        input = {"YZ-X", Convert.ToString(Math.Round(YZ(2), 4)), Convert.ToString(Math.Round(YZ(1), 4))}
        insertListview(input, listview)
        Erase input

        input = {"XYZ", Convert.ToString(Math.Round(XYZ(1), 4)), Convert.ToString(Math.Round(XYZ(0), 4))}
        insertListview(input, listview)
        Erase input
    End Sub

    Private Sub insertListview(input() As String, listview1 As ListView)
        Dim lvitem As New ListViewItem(input)
        listview1.Items.Add(lvitem)
    End Sub
End Class
