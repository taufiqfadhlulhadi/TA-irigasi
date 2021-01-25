Imports System.IO
Imports System.Data.OleDb
Imports System.Math
Imports System.Drawing.KnownColor

Public Class Tengah
    Dim connect As OleDbConnection
    Dim comand As OleDbCommand
    Dim read As OleDbDataReader
    Dim kebutuhan(2) As Decimal
    Dim Qblock(2) As Decimal
    Dim Qin(2) As Decimal
    Dim Qout(2) As Decimal
    Dim Qeff(2) As Decimal



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'MsgBox(tersedia.Length.ToString)
        OpenFileDialog1.Filter = "Excel File|*.xlsx;*.xls"
        Dim dir_destiny As String = Application.StartupPath

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim dir_source As String = Path.GetFullPath(OpenFileDialog1.FileName)
            Dim fileName = Path.GetFileName(OpenFileDialog1.FileName)
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

        If File.Exists(dir_destiny & "\" & Path.GetFileName(OpenFileDialog1.FileName)) Then
            connect = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dir_destiny & "\" & Path.GetFileName(OpenFileDialog1.FileName) & ";Extended Properties=Excel 12.0;")
            comand = New OleDbCommand()

            Try
                connect.Open()
                'MsgBox("connected")
                With comand
                    .Connection = connect
                    .CommandText = "select * from [TENGAH$]"
                    .ExecuteNonQuery()
                    read = .ExecuteReader
                End With

                While read.Read
                    Try
                        If Trim(read(0)) = "M" Then
                            Qin(0) = read(1)
                            kebutuhan(0) = read(2)
                            Qblock(0) = read(3)
                            Qout(0) = read(4)
                            Qeff(0) = read(5)
                            'MsgBox(Qin(0).ToString + "   " + kebutuhan(0).ToString + "   " + Qblock(0).ToString + "    " + Qout(0).ToString + "   " + Qeff(0).ToString)
                            'tersedia(0) = read(1)
                            'Qmasuk(0) = read(3)
                            'Qout(0) = read(4)
                            'If kebutuhan(0) > tersedia(0) Then
                            '    PanelA.BackColor = Color.Red
                            'ElseIf kebutuhan(0) < tersedia(0) Then
                            '    PanelA.BackColor = Color.Yellow
                            'Else
                            '    PanelA.BackColor = Color.Green
                            'End If
                            'tersedia(1) = tersedia(0) - kebutuhan(0)
                            'If tersedia(1) < 0 Then
                            '    tersedia(1) = 0
                            'End If
                        End If

                        If Trim(read(0)) = "N" Then
                            Qin(1) = read(1)
                            kebutuhan(1) = read(2)
                            Qblock(1) = read(3)
                            Qout(1) = read(4)
                            Qeff(1) = read(5)
                            'MsgBox(Qin(1).ToString + "   " + kebutuhan(1).ToString + "   " + Qblock(1).ToString + "    " + Qout(1).ToString + "   " + Qeff(1).ToString)
                            'kebutuhan(1) = read(2)
                            'tersedia(1) = read(1)
                            'Qmasuk(1) = read(3)
                            'Qout(1) = read(4)
                            'If kebutuhan(1) > tersedia(1) Then
                            '    PanelB.BackColor = Color.Red
                            'ElseIf kebutuhan(1) < tersedia(1) Then
                            '    PanelB.BackColor = Color.Yellow
                            'Else
                            '    PanelB.BackColor = Color.Green
                            'End If
                            'tersedia(2) = tersedia(1) - kebutuhan(1)
                            'If tersedia(2) < 0 Then
                            '    tersedia(2) = 0
                            'End If
                        End If

                        If Trim(read(0)) = "O" Then
                            kebutuhan(2) = read(2)
                            Qblock(2) = read(3)
                            Qeff(2) = read(5)
                            'MsgBox(kebutuhan(2).ToString + "   " + Qblock(2).ToString + "   " + Qeff(2).ToString)
                            'kebutuhan(2) = read(2)
                            'tersedia(2) = read(1)
                            'Qmasuk(2) = read(3)
                            'Qout(2) = read(4)
                            'If kebutuhan(2) > tersedia(2) Then
                            '    PanelC.BackColor = Color.Red
                            'ElseIf kebutuhan(2) < tersedia(2) Then
                            '    PanelC.BackColor = Color.Yellow
                            'Else
                            '    PanelC.BackColor = Color.Green
                            'End If
                            'tersedia(3) = tersedia(2) - kebutuhan(2)
                            'If tersedia(3) < 0 Then
                            '    tersedia(3) = 0
                            'End If
                        End If

                        'If Trim(read(0)) = "D" Then
                        '    kebutuhan(3) = read(2)
                        '    'tersedia(3) = read(1)
                        '    If kebutuhan(3) > tersedia(3) Then
                        '        PanelD.BackColor = Color.Red
                        '    ElseIf kebutuhan(3) < tersedia(3) Then
                        '        PanelD.BackColor = Color.Yellow
                        '    Else
                        '        PanelD.BackColor = Color.Green
                        '    End If
                        'End If
                    Catch ex As Exception

                    End Try
                End While
                connect.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
                connect.Close()
            End Try
        End If

        Button2.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        PictureBox2.BackColor = Color.Red
        PictureBox3.BackColor = Color.Red
        PictureBox4.BackColor = Color.Red

        Label13.Show()
        Label14.Show()
        Label15.Show()

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
        Dim panelS() As Panel = {PanelA, PanelB, PanelC}
        Dim labelS() As Label = {Label13, Label14, Label15}
        Dim pictureboxs() As PictureBox = {PictureBox2, PictureBox3, PictureBox4}
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
        insertListview(input)
        Erase input

        input = {"MN-O", Convert.ToString(Math.Round(MN(2), 4)), Convert.ToString(Math.Round(MN(1), 4))}
        insertListview(input)
        Erase input

        input = {"M-NO", Convert.ToString(Math.Round(NO(2), 4)), Convert.ToString(Math.Round(NO(1), 4))}
        insertListview(input)
        Erase input

        input = {"MO-N", Convert.ToString(Math.Round(MO(2), 4)), Convert.ToString(Math.Round(MO(1), 4))}
        insertListview(input)
        Erase input

        input = {"N-MO", Convert.ToString(Math.Round(MO(2), 4)), Convert.ToString(Math.Round(MN(1), 4))}
        insertListview(input)
        Erase input

        input = {"NO-M", Convert.ToString(Math.Round(NO(2), 4)), Convert.ToString(Math.Round(NO(3), 4))}
        insertListview(input)
        Erase input

        input = {"MNO", Convert.ToString(Math.Round(MNO(1), 4)), Convert.ToString(Math.Round(MNO(0), 4))}
        insertListview(input)
        Erase input
    End Sub
    Private Sub insertListview(input() As String)
        Dim lvitem As New ListViewItem(input)
        ListView1.Items.Add(lvitem)
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button2.Enabled = False
        Label13.Hide()
        Label14.Hide()
        Label15.Hide()
    End Sub

    Private Sub PanelA_Paint(sender As Object, e As PaintEventArgs) Handles PanelA.Paint

    End Sub
End Class
