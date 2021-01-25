Imports System.IO
Imports System.Data.OleDb
Imports System.Math
Imports System.Drawing.KnownColor

Public Class Hilir
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
                    .CommandText = "select * from [HILIR$]"
                    .ExecuteNonQuery()
                    read = .ExecuteReader
                End With

                While read.Read
                    Try
                        If Trim(read(0)) = "X" Then
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

                        If Trim(read(0)) = "Y" Then
                            kebutuhan(1) = read(2)
                            Qblock(1) = read(3)
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

                        If Trim(read(0)) = "Z" Then
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
        For b As Integer = 0 To kebutuhan.Length - 1
            kebutuhanPerJam(b) = ((11.84083 * kebutuhan(b)) - 341.67) / 24
            'MsgBox(kebutuhanPerJam(x))
        Next
        Dim sisa As Decimal = 0
        Dim j As Integer = 0
        Dim panelS() As Panel = {PanelA, PanelB, PanelC}
        Dim labelS() As Label = {Label13, Label14, Label15}
        Dim pictureboxs() As PictureBox = {PictureBox2, PictureBox3, PictureBox4}
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
        insertListview(input)
        Erase input

        input = {"XY-Z", Convert.ToString(Math.Round(XY(2), 4)), Convert.ToString(Math.Round(XY(1), 4))}
        insertListview(input)
        Erase input

        input = {"X-YZ", Convert.ToString(Math.Round(YZ(2), 4)), Convert.ToString(Math.Round(YZ(1), 4))}
        insertListview(input)
        Erase input

        input = {"XZ-Y", Convert.ToString(Math.Round(XZ(2), 4)), Convert.ToString(Math.Round(XZ(1), 4))}
        insertListview(input)
        Erase input

        input = {"Y-XZ", Convert.ToString(Math.Round(XZ(2), 4)), Convert.ToString(Math.Round(XZ(1), 4))}
        insertListview(input)
        Erase input

        input = {"YZ-X", Convert.ToString(Math.Round(YZ(2), 4)), Convert.ToString(Math.Round(YZ(1), 4))}
        insertListview(input)
        Erase input

        input = {"XYZ", Convert.ToString(Math.Round(XYZ(1), 4)), Convert.ToString(Math.Round(XYZ(0), 4))}
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

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

    End Sub
End Class
