﻿Imports System.Text
Imports System
Imports System.Configuration
Imports System.Math
Imports System.Collections.Generic
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Globalization
Imports System.Threading
'Imports Word = Microsoft.Office.Interop.Word
Imports System.Runtime.InteropServices

Public Class Form1
    Dim form533(2000, 2) As Double       'Formule 5.33 pagina 330 Machinendynamik

    'Naam
    'Model
    'Tekst
    'L1-Tussen de lagers
    'L2-overhung
    'L3-Star deel in de waaier
    'As dikte tussen de lagers
    'As dikte
    'Weight impeller
    'Jp
    'Ja
    'Stiffness bearing/support buiten
    'Stiffness bearing/support binnen
    'overhungY/N
    Public Shared fan() As String = {
     "Vrije invoer;Q16..;inlet/dia/Type;                    471;66;0;15;15;         3.7;0.0185;0.00925;400;400;Y",
     "Machinendynamik;Test #1, overhung;Aufgabe A5.5;                 472;66;0;15;15;         3.7;0.0185;0.00925;400;400;Y",
     "Dynamics Rotating Machines;Test #2, tussen lagers;Example 3.5.1;     250;250;000;200;111;    122.7;0.6134;2.8625;1.0;1.0;N",
     "Tetrapak;Bedum 3;1800/1825/T33;                       1000;500;100;180;180;   1125;450;230;20;20;Y",
     "Foster Wheeler;Q16.0071;2600/2610/T33;                2380;2380;000;400;0;    1525;864;432;89;89;N"
     }

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim words() As String

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")      'Decimal separator "."
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")    'Decimal separator "."
        ComboBox1.Items.Clear()                    'Note Combobox1 contains"startup" to prevent exceptions

        '-------Fill combobox1, Fan type selection------------------
        For hh = 0 To (fan.Length - 1)            'Fill combobox3 with steel data
            words = fan(hh).Split(";")
            ComboBox1.Items.Add(words(0))
        Next hh

        If ComboBox1.Items.Count > 0 Then
            ComboBox1.SelectedIndex = 2                 'Select Fan data
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown9.ValueChanged, NumericUpDown22.ValueChanged, RadioButton1.CheckedChanged, CheckBox1.CheckedChanged, CheckBox2.CheckedChanged, CheckBox3.CheckedChanged
        Calc_nr()
        draw_chart1()
    End Sub
    Private Sub Calc_nr()
        Dim i As Integer
        Dim L1, L2, L3, massa, speed_rad As Double
        Dim C1, C2 As Double
        Dim d11, d12, d22 As Double
        Dim E_steel, shaft_radius, shaft_overhang_radius, I1_shaft, I2_overhung As Double
        Dim JP_imp, JA_imp As Double
        Dim om10, om20, term1, term2 As Double
        Dim om_krit1, om_krit2, omega_asym As Double

        Try
            E_steel = 210.0 * 10 ^ 3                        'Young N/mm^2

            L1 = NumericUpDown1.Value                           'Length 1 [mm] tussen lagers
            L2 = NumericUpDown2.Value                           'Length 2 [mm] overhung
            L3 = NumericUpDown3.Value                           'Starre Length 3 [m]
            massa = NumericUpDown4.Value                        'Weight waaier [kg]

            C1 = NumericUpDown6.Value * 1000                    '[N/mm]
            C2 = NumericUpDown7.Value * 1000                    '[N/mm]
            shaft_radius = NumericUpDown8.Value / 2             '[mm] as tussen de lagers radius
            shaft_overhang_radius = NumericUpDown9.Value / 2    '[mm] as tussen de lagers radius
            JP_imp = NumericUpDown10.Value                      '[kg.m2] Massa Traagheid hartlijn (JP=1/b.m.D^2)
            JA_imp = NumericUpDown11.Value                      '[kg.m2] Massa Traagheid haaks op hartlijn (JA= 1/16.m.D^2(1+4/3(h/D)^2))

            'I circel is PI/4 * r^4
            I1_shaft = PI / 4 * shaft_radius ^ 4                 'Traagheidsmoment cirkel
            I2_overhung = PI / 4 * shaft_overhang_radius ^ 4     'Traagheidsmoment cirkel

            If JA_imp > JP_imp Then
                GroupBox1.Text = "Massa traagheid waaier (walsvormig !!) "
            Else
                GroupBox1.Text = "Massa traagheid waaier (schijfvormig)"
            End If

            If RadioButton1.Checked Then
                '---------------- Tabelle 5.1 Nr 4 (Overhung) -------------
                Label1.Text = "L1, lengte tussen de lagers [mm]"
                Label2.Text = "L2, Overhang [mm]"
                Label3.Visible = True
                NumericUpDown3.Visible = True
                Label11.Visible = True
                NumericUpDown9.Visible = True
                TextBox31.Visible = True
                '--------------- d11= Alfa --------------------------------
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
            Else
                '---------------- Tabelle 5.1 Nr 3 (Between bearings) -------------
                '--------------- d11= Alfa ----------------------------------------
                Label1.Text = "L1, lengte lager#1-waaier [mm]"
                Label2.Text = "L2, lengte waaier-lager#2 [mm]"
                Label3.Visible = False
                NumericUpDown3.Visible = False
                Label11.Visible = False
                NumericUpDown9.Visible = False
                TextBox31.Visible = False

                '--------------- d11= Alfa --------------------------------
                d11 = L2 ^ 2 / (C1 * (L1 + L2) ^ 2)
                d11 += L1 ^ 2 / (C2 * (L1 + L2) ^ 2)
                d11 += (L1 ^ 2 * L2 ^ 2) / (3 * E_steel * I1_shaft * (L1 + L2))
                d11 /= 1000                                             '[m/N]

                '--------------- d12= Delta en Gamma -------------
                d12 = -L2 / (C1 * (L1 + L2) ^ 2)
                d12 += L1 / (C2 * (L1 + L2) ^ 2)
                d12 += (L1 * L2 * (L2 - L1)) / (3 * E_steel * I1_shaft * (L1 + L2)) '[1/N]

                '--------------- d22= Beta ----------------
                d22 = (1 / C1 + 1 / C2)
                d22 /= (L1 + L2) ^ 2
                d22 += (L1 ^ 3 + L2 ^ 3) / (3 * E_steel * I1_shaft * (L1 + L2) ^ 2)
                d22 *= 1000                                             '[1/(meter.N)]
            End If

            '----------------------- formel 5.33 ------------------------
            speed_rad = -NumericUpDown22.Value
            For i = 1 To 2000                                       'Array size
                speed_rad += NumericUpDown22.Value * 2 / 2000       'increment step [rad/s]
                form533(i, 0) = speed_rad                           'Waaier hoeksnelheid [rad/s]

                form533(i, 1) = -1 + (speed_rad ^ 2 * d11 * massa)
                form533(i, 1) /= (d22 - ((d11 * d22 - d12 ^ 2) * massa * speed_rad ^ 2)) * JP_imp * speed_rad
                form533(i, 1) += JA_imp / JP_imp * speed_rad

                ' TextBox16.Text += form533(i, 0).ToString & ",  " & form533(i, 1).ToString & vbCrLf
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

            '---------- omega _asymptote----------------
            omega_asym = d22 / (massa * (d11 * d22 - d12 ^ 2))
            omega_asym = Sqrt(omega_asym)

            TextBox2.Text = d11.ToString((("0.00 E0")))                     'alfa
            TextBox3.Text = d12.ToString((("0.00 E0")))                     'gamma en delta
            TextBox4.Text = d22.ToString((("0.00 E0")))                     'beta

            TextBox5.Text = Math.Round(rad_to_hz(om_krit1), 1).ToString     'om_krit1 [Hz]
            TextBox6.Text = Math.Round(rad_to_hz(om_krit2), 1).ToString     'om_krit2 [Hz]

            TextBox11.Text = Math.Round(rad_to_hz(om10), 0).ToString        'Omega 10 bij stilstand
            TextBox12.Text = Math.Round(rad_to_hz(om20), 0).ToString        'Omega 20 bij stilstand

            TextBox14.Text = Math.Round(omega_asym, 0).ToString                     'Omega asymptote
            TextBox15.Text = Math.Round(rad_to_hz(omega_asym), 0).ToString          'Omega asymptote
            TextBox39.Text = Math.Round(rad_to_hz(omega_asym) * 60, 0).ToString     'Omega asymptote

            TextBox34.Text = Math.Round(om10, 0).ToString                   'Omega 10 bij stilstand
            TextBox35.Text = Math.Round(om20, 0).ToString                   'Omega 20 bij stilstand

            TextBox1.Text = Math.Round(rad_to_hz(om_krit1) * 60, 0).ToString   'om_krit1 [rmp]
            TextBox13.Text = Math.Round(rad_to_hz(om_krit2) * 60, 0).ToString  'om_krit2 [rmp]

            TextBox32.Text = Math.Round(om_krit1, 0).ToString               'om_krit1 [Rad/s]
            TextBox33.Text = Math.Round(om_krit2, 0).ToString               'om_krit2 [Rad/s]

            TextBox30.Text = I1_shaft.ToString((("0.00 E0")))                   'Buigtraagheidsmoment [m^4]
            TextBox31.Text = I2_overhung.ToString((("0.00 E0")))                'Buigtraagheidsmoment [m^4]

            ' ------- Check sanity -------
            If L3 > L2 Then
                NumericUpDown3.BackColor = Color.Red
            Else
                NumericUpDown3.BackColor = Color.Yellow
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub draw_chart1()
        Dim hh, limit As Integer
        Dim om10, om20, krit1 As Double

        Try
            Chart1.Series.Clear()
            Chart1.ChartAreas.Clear()
            Chart1.Titles.Clear()

            For hh = 0 To 6
                Chart1.Series.Add("s" & hh.ToString)
                Chart1.Series(hh).ChartType = SeriesChartType.Point
                Chart1.Series(hh).IsVisibleInLegend = False
                Chart1.Series(hh).Color = Color.Black
            Next
            Chart1.Series(5).ChartType = SeriesChartType.Line       'Onbalans lijn
            Chart1.Series(6).ChartType = SeriesChartType.Line       'X=0 lijn

            Chart1.ChartAreas.Add("ChartArea0")
            Chart1.Series(0).ChartArea = "ChartArea0"
            If RadioButton1.Checked Then
                Chart1.Titles.Add("Campbell diagram, overhung, isotropic short bearings, flex shaft, no damping")
            Else
                Chart1.Titles.Add("Campbell diagram, between bearing, isotropic short bearings, flex shaft, no damping")
            End If
            Chart1.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)

            '--------------- Legends and titles ---------------
            Chart1.ChartAreas("ChartArea0").AxisX.Title = "Angular speed impeller [rad/s]"
            Chart1.ChartAreas("ChartArea0").AxisY.Title = "Eigenfrequentie [rad/s]"
            Chart1.ChartAreas("ChartArea0").AxisY.RoundAxisValues()
            Chart1.ChartAreas("ChartArea0").AxisX.RoundAxisValues()

            '--------- Chart min size---------------
            If CheckBox1.Checked Then                           'Flip
                Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Else
                Chart1.ChartAreas("ChartArea0").AxisX.Minimum = -NumericUpDown22.Value
            End If

            '--------- Chart max size---------------
            Chart1.ChartAreas("ChartArea0").AxisX.Maximum = NumericUpDown22.Value
            Chart1.ChartAreas("ChartArea0").AxisY.Maximum = NumericUpDown22.Value
            Chart1.ChartAreas("ChartArea0").AlignmentOrientation = DataVisualization.Charting.AreaAlignmentOrientations.Vertical

            '-------- snijpunten -----------
            Double.TryParse(TextBox34.Text, om10)
            Double.TryParse(TextBox35.Text, om20)
            Double.TryParse(TextBox32.Text, krit1)
            Chart1.Series(2).Points.AddXY(0, om10)            'Omega 10 [Rad/sec]
            Chart1.Series(3).Points.AddXY(0, om20)            'Omega 20 [Rad/sec]
            Chart1.Series(4).Points.AddXY(krit1, krit1)       'Kritisch1 [Rad/sec]
            Chart1.Series(2).Points(0).MarkerStyle = MarkerStyle.Circle
            Chart1.Series(2).Points(0).MarkerSize = 10
            Chart1.Series(3).Points(0).MarkerStyle = MarkerStyle.Circle
            Chart1.Series(3).Points(0).MarkerSize = 10
            Chart1.Series(4).Points(0).MarkerStyle = MarkerStyle.Star10
            Chart1.Series(4).Points(0).MarkerSize = 20

            '---------------- draw formule 5.33 -------------------------------
            If CheckBox3.Checked Then
                Chart1.Series(1).ChartType = SeriesChartType.Point
            Else
                Chart1.Series(1).ChartType = SeriesChartType.Line
            End If

            Chart1.Series(1).BorderWidth = 1        'Formule 5.33
            limit = NumericUpDown22.Value                       'Limit in [rad/s]
            For hh = 1 To 2000                                  'Array size
                If form533(hh, 1) < limit And form533(hh, 0) > 0 Then
                    If CheckBox1.Checked Then
                        Chart1.Series(1).Points.AddXY(Abs(form533(hh, 1)), form533(hh, 0))
                    Else
                        Chart1.Series(1).Points.AddXY(form533(hh, 1), form533(hh, 0))
                    End If
                End If
            Next

            '--------draw onbalanslijn----------
            If CheckBox2.Checked Then
                Chart1.Series(5).Points.AddXY(0, 0)
                Chart1.Series(5).Points.AddXY(limit / 2, limit / 2)
                Chart1.Series(5).Points.AddXY(limit, limit)
                Chart1.Series(5).Points(1).Label = "Unbalance"       'Add Remark 
                Chart1.Series(5).BorderWidth = 3
            End If

            '--------X=0 lijn----------
            If Not CheckBox1.Checked Then
                Chart1.Series(6).Points.AddXY(0, 0)
                Chart1.Series(6).Points.AddXY(0, limit)
                Chart1.Series(6).BorderWidth = 3
            End If

        Catch ex As Exception
            'MessageBox.Show("nnnnnn")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, TabPage4.Enter, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown12.ValueChanged, NumericUpDown18.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown19.ValueChanged
        Dim length_L, length_A, length_B, mmassa, diam_tussen, diam_overhung, young As Double
        Dim C_tussen, I_as_tussen, I_as_overhung, fr_krit As Double

        length_L = NumericUpDown12.Value / 1000
        length_A = NumericUpDown13.Value / 1000
        length_B = length_L - length_A
        mmassa = NumericUpDown15.Value                                  '[kg]
        diam_tussen = NumericUpDown16.Value / 1000                      '[mm]

        young = NumericUpDown17.Value * 10 ^ 9                          '[N/m2]

        '-------------- Tussen de lagers -----------------
        I_as_tussen = PI / 4 * (diam_tussen / 2) ^ 4                    '[m4]
        C_tussen = 3 * young * I_as_tussen * length_L
        C_tussen /= (length_A ^ 2 * length_B ^ 2)
        fr_krit = Sqrt(C_tussen / mmassa)                               '[Rad/sec]
        fr_krit /= (2 * PI)                                             '[Hz]

        TextBox17.Text = Round(length_B * 1000, 0).ToString
        TextBox18.Text = I_as_tussen.ToString((("0.00 E0")))
        TextBox19.Text = C_tussen.ToString((("0.00 E0")))                   'Buigstijfheid
        TextBox20.Text = Round(fr_krit, 0).ToString                     '[Hz]
        TextBox27.Text = Round(fr_krit * 60, 0).ToString                '[rpm]


        '-------------- Overhung -----------------
        Dim Overhung_L, Overhung_A, C_Overhung, fr_krit_overhung As Double

        diam_overhung = NumericUpDown19.Value / 1000                     '[m]
        Overhung_L = NumericUpDown14.Value / 1000
        Overhung_A = NumericUpDown18.Value / 1000                        'Overhung

        I_as_overhung = PI / 4 * (diam_overhung / 2) ^ 4                '[m4]
        C_Overhung = 3 * young * I_as_overhung
        C_Overhung /= (Overhung_A ^ 2 * (Overhung_A + Overhung_L))

        fr_krit_overhung = Sqrt(C_Overhung / mmassa)                    '[Rad/sec]
        fr_krit_overhung /= (2 * PI)                                    '[Hz]

        TextBox26.Text = I_as_overhung.ToString((("0.00 E0")))
        TextBox21.Text = C_Overhung.ToString((("0.00 E0")))                 'Buigstijfheid
        TextBox22.Text = Round(fr_krit_overhung, 0).ToString            '[Hz]   
        TextBox28.Text = Round(fr_krit_overhung * 60, 0).ToString       '[rpm]

        '---------------- Check lengtes --------------------
        If length_A > length_L * 0.95 Then   'Residual torque too big,  problem in choosen bouderies
            NumericUpDown13.BackColor = Color.Red
        Else
            NumericUpDown13.BackColor = SystemColors.Window
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, TabPage5.Enter, NumericUpDown21.ValueChanged, NumericUpDown20.ValueChanged, NumericUpDown25.ValueChanged, NumericUpDown24.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown26.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown31.ValueChanged, NumericUpDown30.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown28.ValueChanged
        Dim Dia, hoog, massa, Iz, Ix As Double
        Dim sp1, sp2, spc As Double

        Dia = NumericUpDown20.Value / 1000          '[m]
        hoog = NumericUpDown21.Value / 1000         '[m]

        massa = PI / 4 * Dia ^ 2 * hoog * 7800      'Staal
        Iz = 0.5 * massa * (Dia / 2) ^ 2
        Ix = massa / 12 * (3 * (Dia / 2) ^ 2 + hoog ^ 2)

        TextBox23.Text = Round(Iz, 2).ToString
        TextBox24.Text = Round(Ix, 2).ToString
        TextBox25.Text = Round(massa, 0).ToString

        sp1 = NumericUpDown24.Value
        sp2 = NumericUpDown25.Value

        spc = sp1 * sp2 / (sp1 + sp2)

        TextBox10.Text = Round(spc, 1).ToString
        TextBox36.Text = Round(sp1 + sp2, 1).ToString

        '-------- Berekening veerconstante staal-----------------
        Dim Breed, Dik, Lang, Veer_C, E_carbon As Double

        Breed = NumericUpDown5.Value                    '[mm]
        Dik = NumericUpDown23.Value                     '[mm]
        Lang = NumericUpDown26.Value                    '[mm]
        E_carbon = NumericUpDown27.Value                '[kN/mm2]
        Veer_C = E_carbon * Breed * Dik / Lang          '[kN/mm]

        TextBox40.Text = Round(Veer_C, 0).ToString      '1 plaat
        TextBox41.Text = Round(Veer_C * 4, 0).ToString  'doos 4 platen

        '-------- Eigenfrequen+ gewicht geeft veerconstante staal-----------------
        Dim gewicht, eigenfreq, C_Veer_support As Double

        gewicht = NumericUpDown30.Value                     '[kg]
        eigenfreq = NumericUpDown31.Value * 2 * PI          '[Hz]
        C_Veer_support = eigenfreq ^ 2 * gewicht / 10 ^ 6   '[kN/mm]

        TextBox42.Text = Round(C_Veer_support, 0).ToString  'doos 4 platen

        '-------- Eigenfrequentie door gewicht en veerconstante support-----------------
        Dim gewicht2, eigenfreq2, C_Veer_support2 As Double

        gewicht2 = NumericUpDown28.Value                        '[kg]
        C_Veer_support2 = NumericUpDown29.Value                 '[Hz]
        eigenfreq2 = Sqrt(C_Veer_support2 * 10 ^ 6 / gewicht2)  '[Rad/sec]
        eigenfreq2 /= 2 * PI                                    '[Hz]

        TextBox43.Text = Round(eigenfreq2, 0).ToString       '
    End Sub
    'Converts Radial per second to Hz
    Private Function rad_to_hz(rads As Double)
        Return (rads / (2 * PI))
    End Function
    'Reading data into the comboxes
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim words() As String

        If (ComboBox1.SelectedIndex > 0) Then      'Prevent exceptions
            words = fan(ComboBox1.SelectedIndex).Split(";")
            TextBox7.Text = words(0)                            'Naam
            TextBox8.Text = words(1)                            'Model
            TextBox9.Text = words(2)                            'Tekst
            NumericUpDown1.Value = Convert.ToDouble(words(3))   'L1-Tussen de lagers
            NumericUpDown2.Value = Convert.ToDouble(words(4))   'L2-overhung
            NumericUpDown3.Value = Convert.ToDouble(words(5))   'L3-Star deel in de waaier
            NumericUpDown8.Value = Convert.ToDouble(words(6))   'As dikte tussen de lagers
            NumericUpDown9.Value = Convert.ToDouble(words(7))   'As dikte
            NumericUpDown4.Value = Convert.ToDouble(words(8))   'Weight
            NumericUpDown10.Value = Convert.ToDouble(words(9))  'Jp
            NumericUpDown11.Value = Convert.ToDouble(words(10)) 'Ja
            NumericUpDown6.Value = Convert.ToDouble(words(11))  'Stiffness bearing/support buiten
            NumericUpDown7.Value = Convert.ToDouble(words(12))  'Stiffness bearing/support binnen
            If String.Compare(words(13), "N") Then
                RadioButton1.Checked = True
                RadioButton2.Checked = False
            Else
                RadioButton1.Checked = False
                RadioButton2.Checked = True
            End If
        End If


        '-------- decimal places for presenting data ---------------------
        If (ComboBox1.SelectedIndex = 1) Or (ComboBox1.SelectedIndex = 2) Then  'Test
            NumericUpDown4.DecimalPlaces = 1        'Massa 
            NumericUpDown6.DecimalPlaces = 1        'Stiffness bearing/support buiten
            NumericUpDown7.DecimalPlaces = 1        'Stiffness bearing/support buiten
            NumericUpDown10.DecimalPlaces = 3       'JP
            NumericUpDown11.DecimalPlaces = 3       'JA
        Else
            NumericUpDown6.DecimalPlaces = 0
            NumericUpDown4.DecimalPlaces = 0
            NumericUpDown7.DecimalPlaces = 0
            NumericUpDown10.DecimalPlaces = 0
            NumericUpDown11.DecimalPlaces = 0
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Under construction
    End Sub
End Class
