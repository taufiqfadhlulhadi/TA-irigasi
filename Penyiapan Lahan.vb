﻿
Imports System.IO
Imports System.Data.OleDb
Imports System.Math
Imports System.Drawing.KnownColor
Public Class Penyiapan_Lahan
    Dim connect As OleDbConnection
    Dim comand As OleDbCommand
    Dim read As OleDbDataReader
    Dim kebutuhan(2) As Decimal
    Dim Qblock(2) As Decimal
    Dim Qin(2) As Decimal
    Dim Qout(2) As Decimal
    Dim Qeff(2) As Decimal



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

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
        insertListview(input)
        Erase input

        input = {"AB-C", Convert.ToString(Math.Round(AB(2), 4)), Convert.ToString(Math.Round(AB(1), 4))}
        insertListview(input)
        Erase input

        input = {"A-BC", Convert.ToString(Math.Round(BC(2), 4)), Convert.ToString(Math.Round(BC(1), 4))}
        insertListview(input)
        Erase input

        input = {"AC-B", Convert.ToString(Math.Round(AC(2), 4)), Convert.ToString(Math.Round(AC(1), 4))}
        insertListview(input)
        Erase input

        input = {"B-AC", Convert.ToString(Math.Round(AC(2), 4)), Convert.ToString(Math.Round(AC(1), 4))}
        insertListview(input)
        Erase input

        input = {"BC-A", Convert.ToString(Math.Round(BC(2), 4)), Convert.ToString(Math.Round(BC(3), 4))}
        insertListview(input)
        Erase input

        input = {"ABC", Convert.ToString(Math.Round(ABC(1), 4)), Convert.ToString(Math.Round(ABC(0), 4))}
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