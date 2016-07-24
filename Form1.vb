﻿Imports System.Text
Imports System
Imports System.Configuration
Imports System.Math
Imports System.Collections.Generic
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Globalization

Public Class Form1
    Dim omega0(4001) As Double           'Excitation equal to waaier Hz
    Dim omega1(4001) As Double           'For calculation on Eigen frequency
    Dim omega2(4001) As Double           'For calculation on Eigen frequency
    Dim omega3(4001) As Double           'For calculation on Eigen frequency
    Dim omega4(4001) As Double           'For calculation on Eigen frequency

    Dim form533(4000, 2) As Double       'Formule 5.33 pagina 330 Machinendynamik


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown9.ValueChanged
        Calc_nr6()
        draw_chart1()
    End Sub
    Private Sub Calc_nr6()
        Dim i As Integer
        Dim L1, L2, L3, massa, speed_rad As Double
        Dim C1, C2 As Double
        Dim d11, d12, d22 As Double
        Dim E_steel, shaft_radius, shaft_overhang_radius, I1_shaft, I2_overhung As Double
        Dim JP_imp, JA_imp As Double
        Dim om1, om2, om3, om4, om10, om20, term1, term2 As Double
        Dim om_krit1, om_krit2 As Double

        E_steel = 210.0 * 10 ^ 3                            'Young N/mm^2
        L1 = NumericUpDown1.Value                           'Length 1 [mm] tussen lagers
        L2 = NumericUpDown2.Value                           'Length 2 [mm] overhung
        L3 = NumericUpDown3.Value                           'Starre Length 3 [m]
        massa = NumericUpDown4.Value                        'Weight waaier [kg]

        C1 = NumericUpDown6.Value * 10 ^ 6                  '[N/mm]
        C2 = NumericUpDown7.Value * 10 ^ 6                  '[N/mm]
        shaft_radius = NumericUpDown8.Value / 2             '[mm] as tussen de lagers radius
        shaft_overhang_radius = NumericUpDown9.Value / 2    '[mm] as tussen de lagers radius
        JP_imp = NumericUpDown10.Value                      '[kg.m2] Massa Traagheid hartlijn (JP=1/b.m.D^2)
        JA_imp = NumericUpDown11.Value                      '[kg.m2] Massa Traagheid haaks op hartlijn (JA= 1/16.m.D^2(1+4/3(h/D)^2))

        'I circel is PI/4 * r^4
        I1_shaft = PI / 4 * shaft_radius ^ 4                 'Traagheidsmoment cirkel
        I2_overhung = PI / 4 * shaft_overhang_radius ^ 4     'Traagheidsmoment cirkel

        '---------------- Tabelle 5.1 Nr 4 ---------------------------
        '--------------- d11= Alfa ------------------------
        d11 = L2 ^ 2 / (C1 * L1 ^ 2)
        d11 += (L1 + L2) ^ 2 / (C2 * L1 ^ 2)
        d11 += L1 * L2 ^ 2 / (3 * E_steel * I1_shaft)
        d11 += (L2 ^ 3 - L3 ^ 3) / (3 * E_steel * I2_overhung)
        d11 /= 1000                                             '[m/N]
        '--------------- d12= Delta en Gamma -------------
        d12 = L2 / (C1 * L1 ^ 2)
        d12 += (L1 + L2) / (C2 * L1 ^ 2)
        d12 += L1 * L2 / (3 * E_steel * I1_shaft)
        d12 += (L2 ^ 2 - L3 ^ 2) / (2 * E_steel * I2_overhung)  '[1/N]

        '--------------- d22= Beta ----------------
        d22 = (1 / C1 + 1 / C2) / L1 ^ 2
        d22 += L1 / (3 * E_steel * I1_shaft)
        d22 += (L2 - L3) / (E_steel * I2_overhung)
        d22 *= 1000                                             '[1/(meter.N)]

        For i = 1 To 4000                        'Waaier hoeksnelheid [rad/s]
            speed_rad = i - 2000                 'run from -200 to +200
            om1 = 1 / (d22 * JP_imp * speed_rad)
            om2 = Sqrt(d22 / (massa * (d11 * d22 - d12 ^ 2)))
            om3 = om2
            om4 = JP_imp / JA_imp * speed_rad

            '--------- store in array for later use -----
            omega0(i) = speed_rad          '[Hz] waaier snelheid in Hz
            omega1(i) = rad_to_hz(om1)           '[Hz]
            omega2(i) = rad_to_hz(om2)           '[Hz]
            omega3(i) = rad_to_hz(om3)           '[Hz]
            omega4(i) = rad_to_hz(om4)           '[Hz]
        Next

        '----------------------- formel 5.33 ------------------------
        TextBox16.Clear()
        For i = 1 To 4000                   'Waaier hoeksnelheid [rad/s]
            speed_rad = i - 2000              'run from -200 to +200
            form533(i, 0) = speed_rad                                         'Waaier hoeksnelheid [rad/s]

            form533(i, 1) = -1 + speed_rad ^ 2 * d11 * massa
            form533(i, 1) /= d22 - ((d11 * d22 - d12 ^ 2) * massa * speed_rad ^ 2) * JP_imp * speed_rad
            form533(i, 1) += JA_imp / JP_imp * speed_rad
        Next

        '----------- Omega kritisch #1 ------------------
        om_krit1 = (d11 * massa + d22 * (JA_imp - JP_imp))
        om_krit1 += Sqrt((d11 * massa + d22 * (JA_imp - JP_imp)) ^ 2 - 4 * (d11 * d22 - d12 ^ 2) * massa * (JA_imp - JP_imp))
        om_krit1 *= 0.5
        om_krit1 = Sqrt(1 / om_krit1)

        '----------- Omega kritisch #2 ------------------
        om_krit2 = (d11 * massa + d22 * (JA_imp - JP_imp))
        om_krit2 -= Sqrt((d11 * massa + d22 * (JA_imp - JP_imp)) ^ 2 - 4 * (d11 * d22 - d12 ^ 2) * massa * (JA_imp - JP_imp))
        om_krit2 *= 0.5
        om_krit2 = Sqrt(1 / om_krit2)


        '------------ om10 en om20 (bij stilstand)---formule 5.32--------
        term1 = (d11 * massa + d22 * JA_imp) / (2 * massa * JA_imp * (d11 * d22 - d12 ^ 2))
        term2 = 4 * massa * JA_imp * (d11 * d22 - d12 ^ 2) / (d11 * massa + d22 * JA_imp) ^ 2
        term2 = 1 - term2

        om10 = Sqrt(term1 * (1 + Sqrt(term2)))
        om20 = Sqrt(term1 * (1 - Sqrt(term2)))

        TextBox2.Text = d11.ToString("0.###e0")                     'alfa
        TextBox3.Text = d12.ToString("0.###e0")                     'gamma en delta
        TextBox4.Text = d22.ToString("0.###e0")                     'beta

        TextBox7.Text = Math.Round(rad_to_hz(om1), 1).ToString            'Omega 1 [Hz]
        TextBox8.Text = Math.Round(rad_to_hz(om2), 1).ToString            'Omega 2 [Hz]
        TextBox9.Text = Math.Round(rad_to_hz(om3), 1).ToString            'Omega 3 [Hz]
        TextBox10.Text = Math.Round(rad_to_hz(om4), 1).ToString           'Omega 4 [Hz]

        TextBox5.Text = Math.Round(rad_to_hz(om_krit1), 1).ToString       'om_krit1 [Hz]
        TextBox6.Text = Math.Round(rad_to_hz(om_krit2), 1).ToString       'om_krit2 [Hz]

        TextBox11.Text = Math.Round(rad_to_hz(om10), 1).ToString          'Omega 10 bij stilstand
        TextBox12.Text = Math.Round(rad_to_hz(om20), 1).ToString          'Omega 20 bij stilstand

        TextBox1.Text = Math.Round(rad_to_hz(om_krit1) * 60, 0).ToString   'om_krit1 [rmp]
        TextBox13.Text = Math.Round(rad_to_hz(om_krit2) * 60, 0).ToString  'om_krit2 [rmp]

        TextBox30.Text = I1_shaft.ToString("0.###e0")                   'Buigtraagheidsmoment [m^4]
        TextBox31.Text = I2_overhung.ToString("0.###e0")                'Buigtraagheidsmoment [m^4]
    End Sub

    Private Sub draw_chart1()

        Dim hh, limit As Integer
        Dim om10, om20 As Double

        Try
            Chart1.Series.Clear()
            Chart1.ChartAreas.Clear()
            Chart1.Titles.Clear()

            For hh = 0 To 7
                Chart1.Series.Add("s" & hh.ToString)
                Chart1.Series(hh).ChartType = SeriesChartType.Line
                Chart1.Series(hh).IsVisibleInLegend = False
            Next

            Chart1.ChartAreas.Add("ChartArea0")
            Chart1.Series(0).ChartArea = "ChartArea0"
            Chart1.Titles.Add("Campbell diagram, anisotropic short bearings, flex shaft")
            Chart1.Titles(0).Font = New Font("Arial", 16, System.Drawing.FontStyle.Bold)

            Chart1.Series(0).Name = "Omg1"
            Chart1.Series(0).Color = Color.LightGreen
            Chart1.Series(0).BorderWidth = 1

            Chart1.ChartAreas("ChartArea0").AxisX.Title = "Hoeksnelheid waaier[rad/s]"
            Chart1.ChartAreas("ChartArea0").AxisY.Title = "Eigenfrequentie [rad/s]"
            ' Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Chart1.ChartAreas("ChartArea0").AlignmentOrientation = DataVisualization.Charting.AreaAlignmentOrientations.Vertical
            Chart1.Series(0).YAxisType = AxisType.Primary

            Double.TryParse(TextBox11.Text, om10)
            Double.TryParse(TextBox12.Text, om20)

            Chart1.Series(6).Points.AddXY(0, om10)            'Omega 10
            Chart1.Series(7).Points.AddXY(0, om20)            'Omega 20

            Chart1.Series(6).Points(0).MarkerStyle = MarkerStyle.Circle
            Chart1.Series(6).Points(0).MarkerSize = 20
            Chart1.Series(7).Points(0).MarkerStyle = MarkerStyle.Circle
            Chart1.Series(7).Points(0).MarkerSize = 20

            limit = 5000          'Limit in Hz
            For hh = 1 To 4000           'Waaier hoeksnelheid [rad/s]

                If omega0(hh) < limit And omega0(hh) > -limit Then Chart1.Series(4).Points.AddXY(omega0(hh), omega0(hh))
                If omega1(hh) < limit And omega1(hh) > -limit Then Chart1.Series(0).Points.AddXY(omega0(hh), omega1(hh))
                If omega2(hh) < limit And omega2(hh) > -limit Then Chart1.Series(1).Points.AddXY(omega0(hh), omega2(hh))
                If omega3(hh) < limit And omega3(hh) > -limit Then Chart1.Series(2).Points.AddXY(omega0(hh), omega3(hh))
                If omega4(hh) < limit And omega4(hh) > -limit Then Chart1.Series(3).Points.AddXY(omega0(hh), omega4(hh))
                If form533(hh, 1) < limit And form533(hh, 1) > -limit Then Chart1.Series(5).Points.AddXY(form533(hh, 0), form533(hh, 1))

                TextBox16.Text &= Environment.NewLine & omega0(hh).ToString & ", In= " & form533(hh, 0).ToString & ", Out= " & form533(hh, 1)
            Next

            Chart1.Series(0).Points(10).Label = "Omega 1 (tegenloop)"       'Add Remark 
            Chart1.Series(1).Points(100).Label = "Omega 2 (meeloop)"        'Add Remark 
            Chart1.Series(2).Points(100).Label = "Omega 3 (tegenloop)"      'Add Remark 
            Chart1.Series(3).Points(80).Label = "Omega 4 (meeloop)"         'Add Remark 
            Chart1.Series(4).Points(100).Label = "Onbalans"                 'Add Remark 
        Catch ex As Exception
            'MessageBox.Show("nnnnnn")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, TabPage4.Enter, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown12.ValueChanged, NumericUpDown18.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown19.ValueChanged
        Dim length_L, length_A, length_B, massa, diam_tussen, diam_overhung, young As Double
        Dim C_tussen, I_as_tussen, I_as_overhung, fr_krit As Double

        length_L = NumericUpDown12.Value
        length_A = NumericUpDown13.Value
        length_B = length_L - length_A
        massa = NumericUpDown15.Value
        diam_tussen = NumericUpDown16.Value
        diam_overhung = NumericUpDown19.Value
        young = NumericUpDown17.Value * 1000                            '[N/mm2]

        '-------------- Tussen de lagers -----------------
        I_as_tussen = PI / 4 * (diam_tussen / 2) ^ 4                    '[mm4]
        C_tussen = 3 * young * I_as_tussen * length_L * 1000 / (length_A ^ 2 * length_B ^ 2)
        fr_krit = 1 / (2 * PI) * Sqrt(C_tussen / massa) * 60        '[rmp]

        TextBox17.Text = Round(length_B, 0).ToString
        TextBox18.Text = Round(I_as_tussen / 10 ^ 6, 0).ToString
        TextBox19.Text = Round(C_tussen / 1000, 0).ToString
        TextBox20.Text = Round(fr_krit, 0).ToString                     '[rpm]
        TextBox27.Text = Round(fr_krit / 60, 0).ToString                '[Hz]

        '-------------- Overhung -----------------
        I_as_overhung = PI / 4 * (diam_overhung / 2) ^ 4                '[mm4]
        Dim Overhung_L, Overhung_A, C_Overhung, fr_krit_overhung As Double
        Overhung_L = NumericUpDown14.Value
        Overhung_A = NumericUpDown18.Value
        C_Overhung = 3 * young * I_as_overhung * 1000 / (Overhung_A ^ 2 * (Overhung_A + Overhung_L))
        fr_krit_overhung = 1 / (2 * PI) * Sqrt(C_Overhung / massa) * 60 '[rmp]

        TextBox26.Text = Round(I_as_overhung / 10 ^ 6, 0).ToString
        TextBox21.Text = Round(C_Overhung / 1000, 0).ToString
        TextBox22.Text = Round(fr_krit_overhung, 0).ToString            '[rpm]   
        TextBox28.Text = Round(fr_krit_overhung / 60, 0).ToString       '[Hz]

        '---------------- Check lengtes --------------------
        If length_A > length_L * 0.95 Then   'Residual torque too big,  problem in choosen bouderies
            NumericUpDown13.BackColor = Color.Red
        Else
            NumericUpDown13.BackColor = SystemColors.Window
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, TabPage5.Enter, NumericUpDown21.ValueChanged, NumericUpDown20.ValueChanged
        Dim Dia, hoog, massa, Iz, Ix As Double

        Dia = NumericUpDown20.Value / 1000          '[m]
        hoog = NumericUpDown21.Value / 1000         '[m]

        massa = PI / 4 * Dia ^ 2 * hoog * 7800      'Staal
        Iz = 0.5 * massa * (Dia / 2) ^ 2
        Ix = massa / 12 * (3 * (Dia / 2) ^ 2 + hoog ^ 2)

        TextBox23.Text = Round(Iz, 0).ToString
        TextBox24.Text = Round(Ix, 0).ToString
        TextBox25.Text = Round(massa, 0).ToString
    End Sub
    'Converts Radial per second to Hz
    Private Function rad_to_hz(rads As Double)
        Return (rads / (2 * PI))
    End Function
End Class
