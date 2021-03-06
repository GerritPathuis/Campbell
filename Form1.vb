﻿Imports System.Text
Imports System.IO
Imports System.Math
Imports System.Threading
Imports System.Management
Imports Word = Microsoft.Office.Interop.Word
Imports System.Globalization
Imports System.Windows.Forms.DataVisualization.Charting

Public Class Form1
    Public form533(2000, 2) As Double       'Formule 5.33 pagina 330 Machinendynamik

    Public Shared based_on() As String = {
    "Based on",
    "Maschinendynamik, 11 Auflage, ISBN 978-3-642-29570-6",
    "Dynamics of Rotary Machines, ISBN 978-0-521-85016-2",
    " ",
    "Example #1, Maschinendynamik page,Aufgabe A5.5, page 357",
    "Overhung, disk weight 3.7 kg, shaft diameter 15 mm,",
    "Length between bearing 470 mm, overhung 66 mm",
    "Disc diameter 200mm, 15mm wide, rigid bearings (1000 kN/mm)",
    "E= 210 kN/mm2, Critical Natural frequency 6127 rpm",
    " ",
    "Example #2, Dynamics Rotating Machines; Example 3.5.1, page 85",
    "Between bearings, disk weight 122.7 kg, shaft diameter 200 mm,",
    "Length between bearings 500 mm",
    "Disc/roll diameter 200 mm, 500 mm wide (Jp=0.613,=Ja=2.859 kg.m2)",
    "Bearing stiffness horz. and vert. 1MN/m (1 kN/mm)",
    "Critical Natural frequencies 1219 and 1996 rpm"}

    Public Shared fundation() As String = {
    "ACI 351.3R-04",
    "A long-established rule-of-thumb for machinery on blocktype",
    "foundations Is to make the weight of the foundation",
    "block at least three times the weight Of a rotating machine",
    "And at least five times the weight of a reciprocating machine.",
    "For pile - supported foundations, these ratios are sometimes",
    "reduced so that the foundation block weight, including pile cap,",
    "Is at least 2-1/2 times the weight of a rotating machine",
    "And at least four times the weight of a reciprocating machine.",
    "These ratios are machine weights inclusive Of moving and",
    "stationary parts as compared With the weight Of the concrete",
    "foundation block",
    "Additionally, many designers require the",
    "foundation to be of such weight that the resultant of lateral",
    "And vertical loads falls within the middle third of the foundation",
    "base. That Is, the net effect of lateral and vertical loads Or the",
    "eccentricity of the vertical load should not cause uplift.",
    " ",
    "Supply rotor machine weight is 2500 kg",
    "Inertia block weight shall be > 3.0 x 2500= 7500 kg"}

    Public Shared bearing_support() As String = {
    "VTK bearing support structures",
    "",
    "BETWEEN THE BEARINGS, BOLTED TO THE FLOOR",
    "The bearing support near the motor > 100 kN/mm",
    "Motor side calculate with the bearing stiffness 100 kN/mm",
    "The opposite support side Is 15 kN/mm",
    "",
    "BETWEEN THE BEARINGS, ON VIBRATION ISOLATORS",
    "The bearing support near the motor > 100 kN/mm",
    "Motor calculate with 100 kN/mm",
    "The opposite support side 2 kN/mm !!",
    "With HEB300 + 15mm welded side plates 9 kN/mm @ NDE",
    "With IPE400 + 20mm welded side plates 10 kN/mm @ NDE",
    "Supezet (P17.1053) between bearings stiffness 12 kN/mm @ NDE",
    "Lummes (P18.1076) between bearings stiffness 20 kN/mm @ NDE",
    " ",
    "OVERHUNG BEARING SUPPORT",
    "The bearing support near the motor Is stiff > 100 kN/mm",
    "Calculate with the bearing stiffness 100 kN/mm",
    ""}

    '----------- directory's-----------
    Public dirpath_Eng As String = "N:\Engineering\VBasic\Campbell_input\"
    Public dirpath_Rap As String = "N:\Engineering\VBasic\Campbell_rapport_copy\"
    Public dirpath_Home As String = "C:\Temp\"
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
     "Machinendynamik;Test #1, overhung;Aufgabe A5.5;       472;66;0;15;15;         3.7;0.0185;0.00925;400;400;Y",
     "Dynamics Rotating Machines;Test #2, tussen lagers;Example 3.5.1;     250;250;000;200;111;    122.7;0.6134;2.8625;1.0;1.0;N",
     "Tetrapak;Bedum 3;1800/1825/T33;                       750;562;263;180;130;    968;450;230;400;400;Y",
     "Foster Wheeler;Q16.0071;2600/2610/T33;                2380;2380;000;400;0;    1525;864;432;89;89;N",
     "Tecnimont;P16.0078;HD2 407/1230/T16B;                 850;850;000;190;0;      346;64;32;428;428;N"
     }

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Pro_user, HD_number As String
        Dim user_list As New List(Of String)
        Dim hard_disk_list As New List(Of String)
        Dim pass_name As Boolean = False
        Dim pass_disc As Boolean = False

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")      'Decimal separator "."
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")    'Decimal separator "."

        '------ allowed users with hard disc id's -----
        user_list.Add("GP")
        hard_disk_list.Add("058F63626371") 'Privee PC, graslaan25

        user_list.Add("GerritP")
        hard_disk_list.Add("S2R6NX0H740154H")  'VTK PC, GP

        user_list.Add("GerritP")
        hard_disk_list.Add("0008_0D02_003E_0FBB.")   'VTK laptop, GP

        user_list.Add("KarelB")
        hard_disk_list.Add("165214800214")   'VTK PC, Karel Bakker

        user_list.Add("GP")                  'Privee laptop GP
        hard_disk_list.Add("S28ZNXAG521979")

        user_list.Add("User")                'Privee PC GP
        hard_disk_list.Add("058F63646471")

        user_list.Add("JeroenA")
        hard_disk_list.Add("170228801578")   'VTK laptop, Jeroen
        hard_disk_list.Add("171095402070")   'VTK desktop, Jeroen

        'user_list.Add("ABI")
        'hard_disk_list.Add("174741803447")         'VTK desktop, Ab van Iterson <a.van.Iterson@vtk.nl>

        user_list.Add("bertk")
        'hard_disk_list.Add("WD-WXB1A14C9942")      'VTK oude desktop. BKo
        hard_disk_list.Add("0025_3886_01E9_11D6.")  'VTK new desktop. BKo (24/11/2020)


        Pro_user = Environment.UserName     'User name on the screen
        HD_number = HardDisc_Id()           'Harddisk identification
        Me.Text &= "  (" & Pro_user & ")"

        ComboBox1.SelectedIndex = 0     '8 balls
        ComboBox2.SelectedIndex = 0     '8 rollers

        'Check user name 
        For i = 0 To user_list.Count - 1
            If StrComp(LCase(Pro_user), LCase(user_list.Item(i))) = 0 Then pass_name = True
        Next

        'Check disc_id
        For i = 0 To hard_disk_list.Count - 1
            If CBool(HD_number = Trim(hard_disk_list(i))) Then pass_disc = True
        Next

        If pass_name = False Or pass_disc = False Then
            Form2.Text = "VTK Campbell diagram program"
            Form2.Label2.Text = "User_name= " & Pro_user & ", Pass name= " & pass_name.ToString
            Form2.Label3.Text = "HD_id= "
            Form2.TextBox1.Text = "*" & HD_number & "*"
            Form2.Label4.Text = "Pass disc= " & pass_disc.ToString
            Me.Hide()
            Me.Opacity = 0
            Form2.Show()
        End If

        For hh = 0 To (based_on.Length - 1)
            TextBox60.Text &= based_on(hh) & vbCrLf
        Next hh

        For hh = 0 To (bearing_support.Length - 1)
            TextBox66.Text &= bearing_support(hh) & vbCrLf
        Next hh

        For hh = 0 To (fundation.Length - 1)
            TextBox72.Text &= fundation(hh) & vbCrLf
        Next hh

        TextBox73.Text =
        "DIN 1940" & vbCrLf &
        "G2.5 (2.5 mm/s [p-p])" & vbCrLf &
        "   -Gas turbines" & vbCrLf &
        "G6.3 (6.3 mm/s [p-p])" & vbCrLf &
        "   -Fans " & vbCrLf &
        "   -Machinery general" & vbCrLf &
        "   -Electric motors <950 rpm" & vbCrLf &
        "G16 (16 mm/s [p-p]) " & vbCrLf &
        "   -Agricultutal machinery" & vbCrLf &
        "G40 (40 mm/s [p-p])" & vbCrLf &
        "   -Car wheels" & vbCrLf &
        "G100 (100 mm/s [p-p]) " & vbCrLf &
        "   -Reciprocating Car engines" & vbCrLf &
        ""

        TextBox29.Text =
        "VTK motor support @ drive side 640 kN/mm" & vbCrLf &
        "VTK motor support @ NON drive side 12 kN/mm" & vbCrLf & vbCrLf &
        "SKF rollager 110-150 kN/mm" & vbCrLf &
        "API 684 Sleeve bearing 89 kN/mm" & vbCrLf &
        "API 684 Tilting pad bearing 125 kN/mm" & vbCrLf & vbCrLf &
        "Maschinendynamik seite 34" & vbCrLf &
        "8x Stahlfeder Machinefundamente  = 5-10 kN/mm" & vbCrLf &
        "8x Industrie vibratie demper 3 kN/mm = 24 kN/mm" & vbCrLf &
        "Walzlager(d50 mm) = 250 - 500 kN/mm" & vbCrLf &
        "Walzlager(d100 mm) = 500 - 1000 kN/mm" & vbCrLf &
        "FS Dynamics used probably 18 kN/mm on Bedum 3"

        TextBox37.Text =
        "Fundation weight > 3x rotor weight (impeller + as)" & vbCrLf &
        "Weight reduces the natural frequency and therefore" & vbCrLf &
        "reduces the swing amplitude" & vbCrLf

        TextBox92.Text =
        "Project P2006.1050, Tata Big fan" & vbCrLf &
        "Diameter 4130 mm, 750 [rpm], impeller weight 7600 [kg]" & vbCrLf &
        "Shaft 690x580 mm (wall 55 [mm]) " & vbCrLf &
        "Between bearing-bearing length 6400 [mm]" & vbCrLf &
        "Bearings houses are sitting on concrete" & vbCrLf &
        "Design temperature 300 [c], Sleeve bearings d= 250 mm" & vbCrLf &
        "Result: Max speed is 983 [rpm] acc. API 673" & vbCrLf &
        "Vibration measured is 1.4 [mm/s] " & vbCrLf

        TextBox93.Text =
        "Reducing impeller weight gives a higher critical speed." & vbCrLf &
        "For overhung fans, bigger distance between bearing up to 1000 mm " & vbCrLf &
        "results is a higher critical speed." & vbCrLf & vbCrLf &
        "Single row ball bearing low C, Double row roller bearing high C " & vbCrLf &
        " "

        TextBox94.Text =
        "Project P20001.1158, Standaard Fasel" & vbCrLf &
        "FD Fan MD 1120/1900/T36" & vbCrLf &
        "Rated 35c, ro= 1.12 kg/m3, 104000 kg/hr, dp= 121.76 mbar, 373 kW" & vbCrLf &
        "Diameter 1900 mm, 1500 [rpm], impeller weight 580 [kg]" & vbCrLf &
        "Between bearing-bearing length 1750x140 [mm]" & vbCrLf &
        "Bearings houses are sitting on concrete (excellent)" & vbCrLf &
        "Design temperature 35 [c], Oil bearings GOF 224BF and GOF220AL" & vbCrLf &
        "Result: Max speed is 1962 [rpm] acc. API 673" & vbCrLf

        TextBox97.Text =
        "Project P16.0051, Biowanze" & vbCrLf &
        "Fan HD 1980/2280/T31A" & vbCrLf &
        "Rated 82c, ro= 0.854 kg/m3, 209000 kg/hr, dp= 107 mbar, 910 kW" & vbCrLf &
        "Diameter 2280 mm, 1500 [rpm], impeller weight 863 [kg]" & vbCrLf &
        "Overhung d=145, L=479 mm (CL bearing-COG) " & vbCrLf &
        "Bearinghousing ZGLO140A d=180, CL-CL= 750 mm" & vbCrLf &
        "LOW Steel Motor-bearing support" & vbCrLf &
        "Result: Max speed is 1508 [rpm] acc. API 673" & vbCrLf

        TextBox98.Text =
        "Project P19.1065, Cargill Krefeld" & vbCrLf &
        "Fan HD 1250/2155/T31A" & vbCrLf &
        "Rated 83c, ro= 0.9832 kg/m3, 128400 kg/hr, dp= 128 mbar, 580 kW" & vbCrLf &
        "Diameter 2155 mm, 1490 [rpm], impeller weight 638 [kg]" & vbCrLf &
        "Overhung d=145, L=410 mm (CL bearing-COG) " & vbCrLf &
        "Bearinghousing ZGLO140A d=180, CL-CL= 750 mm" & vbCrLf &
        "LOW Steel Motor-bearing support" & vbCrLf &
        "Result: Max speed is 2036 [rpm] acc. API 673" & vbCrLf

        TextBox99.Text =
        "Project P17.1053, Supezet" & vbCrLf &
        "Fan MD 1500/1535/T33" & vbCrLf &
        "Rated .c, ro= . kg/m3, . kg/hr, dp= . mbar, .kW" & vbCrLf &
        "Diameter 1490 mm, 1490 [rpm], impeller weight 450 [kg]" & vbCrLf &
        "Between bearing-bearing length 2575x145 [mm]" & vbCrLf &
        "Bearings houses are sitting on top of furnace" & vbCrLf &
        "Design temperature 225 [c], 2x Oil bearings GOF 218BF" & vbCrLf &
        "Result: Max speed is 1637 [rpm] acc. API 673" & vbCrLf

        TextBox100.Text =
        "Project P03.1033, Lummus-CNOOC China" & vbCrLf &
        "Fan 2MD 2100/2370/T33" & vbCrLf &
        "Rated 400c, ro= . kg/m3, . kg/hr, dp= . mbar, . kW" & vbCrLf &
        "Diameter 2370 mm, 750 [rpm], impeller weight 1750 [kg], shaft 400 [mm]" & vbCrLf &
        "Between bearings, bearing-imp-bearing 2998x2998 [mm]" & vbCrLf &
        "Steel Bearings support are sitting on top of furnace" & vbCrLf &
        "Design temperature 400 [c], 2x sleeve bearing" & vbCrLf &
        "Result: Max speed is 901 [rpm] acc. API 673" & vbCrLf

        TextBox101.Text =
        "Project P04.1257, Technip Benelux" & vbCrLf &
        "Fan 2LD 1600/1777/T33" & vbCrLf &
        "Rated 186c, ro= 0.692. kg/m3, 187380 kg/hr, dp= 13.6 mbar, 130 kW" & vbCrLf &
        "Diameter 1770 mm, 900 [rpm], impeller weight 875 [kg], shaft 300 [mm]" & vbCrLf &
        "Between bearings, bearing-imp-bearing 2150x2150 [mm]" & vbCrLf &
        "Steel Bearings support are sitting on top of furnace" & vbCrLf &
        "Design temperature 350 [c], 2x sleeve bearing" & vbCrLf &
        "Result: Max speed is 1244 [rpm] acc. API 673" & vbCrLf

        TextBox102.Text =
        "Concrete slab stiffness" & vbCrLf &
        "See http://homepage.tudelft.nl/p3r3s/MSc_projects/reportBreeveld.pdf " & vbCrLf &
        "page 53  concrete slab stiffness C2= 4.2 kN/mm " & vbCrLf &
        " " & vbCrLf &
        " " & vbCrLf &
        " " & vbCrLf

        TextBox106.Text =
        "Concrete slab foundation" & vbCrLf &
        "Concrete slab adjacent shall be spaced by a minimum of 15 mm" & vbCrLf &
        "Rigid block foundation slab for machinery thickness shall be no less then" & vbCrLf &
        "0.6 + L/30 Where L [m] is the greater of the length or breadth " & vbCrLf &
        "Minimum Concrete steel reinforcement for slabs shall be 30 kg/m3 " & vbCrLf & vbCrLf &
        "Static design" & vbCrLf &
        "Vertical impact 50% of machine dead weight" & vbCrLf &
        "Lateral force 25% dead weight applied between bearings" & vbCrLf &
        "Longitudinal force 25% dead weight applied along axis" & vbCrLf &
        "Lateral and Longitudinal force do not work concurrently" & vbCrLf &
        "Anchor stress not exceed 120 N/mm2" & vbCrLf &
        "Anchor corrosion allowence 3 mm" & vbCrLf &
        "Concrete bearing stress not exceed 5 N/mm2" & vbCrLf &
        "" & vbCrLf

        TextBox112.Text =
        "The eigen frequency must be below the lowest operating speed frequency " &
        "of the fan" & vbCrLf & vbCrLf &
        "For example" & vbCrLf &
        "VSD operate between 600 and 1000 rpm (10 and 16.7Hz)" & vbCrLf &
        "The machine weight and 10Hz give a max spring stiffness" & vbCrLf

        TextBox71.Text =
        "Possible NDE bearing support vibration problem solutions" & vbCrLf & vbCrLf &
        "Use the stiffness of the foundation or floor" & vbCrLf &
        "Increase the stiffness of the NDE bearing support" & vbCrLf &
        "Increase the weight/inertia of NDE bearing support" & vbCrLf &
        "Fill up with approx 750 kg concrete (2500 [kg/m3])" & vbCrLf &
        " "

        'see http://hyperphysics.phy-astr.gsu.edu/hbase/mi2.html
        TextBox88.Text =
        "Linear Calculation basis" & vbCrLf &
        "L_Stifness = Force / displacement [kN/mm]" & vbCrLf &
        "L_period = 2 * PI * Sqrt(fan_weight / L_stiffness) [s]" & vbCrLf &
        "L_Freq = 1 / period [hz]" & vbCrLf & vbCrLf &
        "Torsional calculation basis" & vbCrLf &
        "T_Stifness = Torque / displacement [kNm/rad]" & vbCrLf &
        "T_moment of inertia = (Point mass, radius=R) M*R^2 [kg/m2]" & vbCrLf &
        "T_Moment of inertia = (solid ring, radius=R) 1/2*M*L^2 [kg/m2]" & vbCrLf &
        "T_Moment of inertia = (rod trough end, length=L) 1/3*M*L^2 [kg/m2]" & vbCrLf &
        "T_period = 2 * PI * Sqrt(T_inertia / T_stiffness) [s]" & vbCrLf &
        "T_freq = 1 / T_period [hz]" & vbCrLf &
        "Note: Linear calc is the preferred option due difficulty of determining the " &
        "Moment of Inertia of the NDE bearing support"

        TextBox81.Text =
        "Note, only a part of the weight is taken into account " &
        "depending on the COG position of the support because" & vbCrLf &
        "near the bearing the movement is more intens than at the foot of the support. "

        TextBox7.Text = "P" & DateTime.Now.ToString("yy") & ".10"

        Example_5()
        Bearing_support_stiffnes()
        Calc_sequence()
        Timer1.Enabled = True
    End Sub
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Bearing_support_stiffnes()
        Calc_sequence()
    End Sub
    Private Sub Bearing_support_stiffnes()

        Select Case True
            Case RadioButton1.Checked And RadioButton3.Checked  'Overhung + Steel bearing supports 
                NumericUpDown6.Value = 90 '[kN/mm] Stiffness C1 @ drive
                NumericUpDown7.Value = 90 '[kN/mm] Stiffness C2 @ impeller

            Case RadioButton2.Checked And RadioButton3.Checked 'Between bearings + Steel supports on RUBBER BLOCKS
                NumericUpDown6.Value = 80  '[kN/mm] Stiffness frame C1 @ drive bearing
                NumericUpDown7.Value = 12  '[kN/mm] Stiffness frame C2 @ Not drive bearing

            Case RadioButton2.Checked And RadioButton4.Checked     'Between bearings + concrete supports
                NumericUpDown6.Value = 90   '[kN/mm] Stiffness C1 @ drive
                NumericUpDown7.Value = 90   '[kN/mm] Stiffness C2 @ impeller

            Case RadioButton2.Checked And RadioButton5.Checked     'Between bearings + (steel/Concrete) supports
                NumericUpDown6.Value = 90   '[kN/mm] Stiffness C1 @ drive
                NumericUpDown7.Value = 30   '[kN/mm] Stiffness C2 @ impeller
        End Select

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown9.ValueChanged, NumericUpDown22.ValueChanged, CheckBox1.CheckedChanged, CheckBox2.CheckedChanged, CheckBox3.CheckedChanged, NumericUpDown55.ValueChanged, RadioButton3.CheckedChanged, CheckBox5.CheckedChanged, NumericUpDown68.ValueChanged, RadioButton4.CheckedChanged, RadioButton2.CheckedChanged
        Calc_sequence()
    End Sub
    Private Sub Calc_sequence()
        Bearing_support_stiffnes()  'Check bearing stiffness
        If RadioButton1.Checked Then
            PictureBox15.Image = Campbell_diagram.My.Resources.Resources.Overhung1
            GroupBox5.Visible = CBool(IIf(RadioButton2.Checked, False, True)) 'Overhung
        Else
            PictureBox15.Image = Campbell_diagram.My.Resources.Resources.Between_bearings2
            GroupBox2.Visible = CBool(IIf(RadioButton1.Checked, False, True)) 'Between bearings
        End If


        '------------- Check hollow shaft bore diameter ----
        If NumericUpDown68.Value > NumericUpDown8.Value * 0.6 Then
            NumericUpDown68.Value = CDec(NumericUpDown8.Value * 0.6)
        End If

        NumericUpDown4.DecimalPlaces = CInt(IIf((NumericUpDown4.Value > 100), 0, 1))

        GroupBox12.Text = "Chart settings"
        TextBox54.Text = TextBox23.Text 'Inertia hart line
        TextBox55.Text = TextBox24.Text
        Calc_nr()
        Draw_chart1()
        Calc_simple()
    End Sub
    Private Sub Set_num_digits(num As NumericUpDown)
        Dim tmp As Double
        tmp = num.Value
        num.DecimalPlaces = 1
        Select Case True
            Case tmp < 0.1
                num.DecimalPlaces = 3
                num.Increment = CDec(0.001)
            Case tmp >= 0.1 And tmp < 1
                num.DecimalPlaces = 2
                num.Increment = CDec(0.01)
            Case tmp >= 1 And tmp < 10
                num.DecimalPlaces = 1
                num.Increment = CDec(0.1)
            Case tmp >= 10 And tmp < 100
                num.DecimalPlaces = 0
                num.Increment = 1
            Case Else
                num.DecimalPlaces = 0
                num.Increment = 10
        End Select

    End Sub

    Private Sub Calc_nr()
        Dim i As Integer
        Dim L1, L2, L3, massa, speed_rad As Double
        Dim C1, C2 As Double
        Dim d11, d12, d22 As Double
        Dim E_steel, shaft_overhang_radius As Double
        Dim shaft_r_out, shaft_r_in As Double   '[mm]
        Dim I1_shaft, I2_overhung As Double

        Dim JP_imp, JA_imp, Jr As Double
        Dim discrim As Double
        Dim ω10, ω20, term1, term2 As Double
        Dim ω_krit1, ω_krit2, ω_asym As Double
        Dim max_api_673 As Double

        Set_num_digits(NumericUpDown10)
        Set_num_digits(NumericUpDown11)
        NumericUpDown15.DecimalPlaces = CInt(IIf(NumericUpDown11.Value < 10, 1, 0))

        Try
            E_steel = Calc_young(NumericUpDown55.Value)         'Shaft Young mod. [N/mm^2] 
            TextBox61.Text = (E_steel / 1000).ToString("F0")    'Young's modulus [kN/mm^2]
            L1 = NumericUpDown1.Value                           'Length 1 [mm] tussen lagers
            L2 = NumericUpDown2.Value                           'Length 2 [mm] overhung
            L3 = NumericUpDown3.Value                           'Starre Length 3 [m]
            massa = NumericUpDown4.Value                        'Weight waaier [kg]

            C1 = NumericUpDown6.Value * 1000                    '[N/mm] rigidness support
            C2 = NumericUpDown7.Value * 1000                    '[N/mm] rigidness support

            shaft_r_out = NumericUpDown8.Value / 2         '[mm] radius as tussen de lagers 
            shaft_r_in = NumericUpDown68.Value / 2         '[mm] inside radius hollow shaft 

            shaft_overhang_radius = NumericUpDown9.Value / 2    '[mm] as tussen de lagers radius
            JP_imp = NumericUpDown10.Value                      '[kg.m2] Massa Traagheid hartlijn (JP=1/b.m.D^2)
            JA_imp = NumericUpDown11.Value                      '[kg.m2] Massa Traagheid haaks op hartlijn (JA= 1/16.m.D^2(1+4/3(h/D)^2))

            'I circlw = PI/4 * r^4
            I1_shaft = PI / 4 * (shaft_r_out ^ 4 - shaft_r_in ^ 4)          'Traagheidsmoment cirkel
            I2_overhung = PI / 4 * shaft_overhang_radius ^ 4     'Traagheidsmoment cirkel

            If JA_imp > JP_imp Then
                GroupBox1.Text = "Massa traagheid waaier (walsvormig NOK) "
                GroupBox1.BackColor = Color.Red
            Else
                GroupBox1.Text = "Massa traagheid waaier (schijfvormig OK)"
                GroupBox1.BackColor = Color.White
            End If


            If RadioButton1.Checked Then
                '---------------- Tabelle 5.1 Nr 4 (Overhung) -------------
                Label1.Text = "L1, aslengte tussen de lagers [mm]"
                Label2.Text = "L2, Overhang waaier, incl L3 [mm]"
                Label3.Visible = True
                NumericUpDown3.Visible = True
                Label11.Visible = True
                NumericUpDown9.Visible = True
                TextBox31.Visible = True
                Label114.Visible = True
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
                d22 = 1 / C1 + 1 / C2
                d22 /= L1 ^ 2
                d22 += L1 / (3 * E_steel * I1_shaft)
                d22 += (L2 - L3) / (E_steel * I2_overhung)
                d22 *= 1000                                             '[1/(meter.N)]
            Else
                '---------------- Tabelle 5.1 Nr 3 (Between bearings) -------------
                '--------------- d11= Alfa ----------------------------------------
                Label1.Text = "L1, aslengte fI_diamed lager-waaier [mm] (drive end)"
                Label2.Text = "L2, aslengte waaier-float lager [mm] (non drive end)"
                Label3.Visible = False
                NumericUpDown3.Visible = False
                Label11.Visible = False
                NumericUpDown9.Visible = False
                TextBox31.Visible = False
                Label114.Visible = False

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
            speed_rad = -NumericUpDown22.Value                      'Chart range
            For i = 1 To 2000                                       'Array sI_polare
                speed_rad += Abs(NumericUpDown22.Value * 2 / 2000)  'increment step [rad/s]
                form533(i, 0) = speed_rad                           'Waaier hoeksnelheid [rad/s]

                form533(i, 1) = -1 + (speed_rad ^ 2 * d11 * massa)
                form533(i, 1) /= (d22 - ((d11 * d22 - d12 ^ 2) * massa * speed_rad ^ 2)) * JP_imp * speed_rad
                form533(i, 1) += JA_imp / JP_imp * speed_rad

                ' TextBox16.Text += form533(i, 0).ToString & ",  " & form533(i, 1).ToString & vbCrLf
            Next

            '----------- Omega kritisch(eq. 5.35)------------------
            Jr = JA_imp - JP_imp
            discrim = Sqrt((d11 * massa + d22 * Jr) ^ 2 - 4 * (d11 * d22 - d12 ^ 2) * massa * Jr) 'discriminant

            '----------- Omega kritisch #1 (eq. 5.35)------------------
            ω_krit1 = 0.5 * (d11 * massa + d22 * Jr + discrim)
            ω_krit1 = Sqrt(1 / ω_krit1)

            '----------- Omega kritisch #2 (eq. 5.35)------------------
            ω_krit2 = 0.5 * (d11 * massa + d22 * Jr - discrim)
            ω_krit2 = Sqrt(1 / ω_krit2)

            '------------ ω10 en ω20 (bij stilstand)---(eq. 5.32)--------
            term1 = (d11 * massa + d22 * JA_imp) / (2 * massa * JA_imp * (d11 * d22 - d12 ^ 2))
            term2 = 4 * massa * JA_imp * (d11 * d22 - d12 ^ 2) / (d11 * massa + d22 * JA_imp) ^ 2
            term2 = 1 - term2

            ω10 = Sqrt(term1 * (1 + Sqrt(term2)))
            ω20 = Sqrt(term1 * (1 - Sqrt(term2)))

            '---------- omega _asymptote (eq. 5.34)----------------
            ω_asym = d22 / (massa * (d11 * d22 - d12 ^ 2))
            ω_asym = Sqrt(ω_asym)


            '-------- present results-------------
            TextBox2.Text = d11.ToString((("0.000 E0")))                    'alfa
            TextBox3.Text = d12.ToString((("0.000 E0")))                    'gamma en delta
            TextBox4.Text = d22.ToString((("0.000 E0")))                    'beta

            '--------- krit1 and krit 2----------
            TextBox5.Text = Math.Round(Rad_to_hz(ω_krit1), 1).ToString     'ω_krit1 [Hz]
            TextBox6.Text = Math.Round(Rad_to_hz(ω_krit2), 1).ToString     'ω_krit2 [Hz]

            TextBox1.Text = Math.Round((Rad_to_hz(ω_krit1) * 60), 0).ToString   'ω_krit1 [rmp]
            TextBox13.Text = Math.Round((Rad_to_hz(ω_krit2) * 60), 0).ToString  'ω_krit2 [rmp]

            TextBox32.Text = Math.Round(ω_krit1, 0).ToString               'ω_krit1 [Rad/s]
            TextBox33.Text = Math.Round(ω_krit2, 0).ToString               'ω_krit2 [Rad/s]

            '--------- ω10  and ω20 -----------
            TextBox34.Text = Math.Round(ω10, 0).ToString                   'Omega 10 bij stilstand
            TextBox35.Text = Math.Round(ω20, 0).ToString                   'Omega 20 bij stilstand

            TextBox11.Text = Math.Round(Rad_to_hz(ω10), 0).ToString        'Omega 10 bij stilstand
            TextBox12.Text = Math.Round(Rad_to_hz(ω20), 0).ToString        'Omega 20 bij stilstand

            TextBox59.Text = Math.Round(Rad_to_hz(ω10) * 60, 0).ToString   'Omega 10 bij stilstand
            TextBox58.Text = Math.Round(Rad_to_hz(ω20) * 60, 0).ToString   'Omega 20 bij stilstand

            '---------- asymtote--------------
            TextBox14.Text = Math.Round(ω_asym, 0).ToString                     'Omega asymptote
            TextBox15.Text = Math.Round(Rad_to_hz(ω_asym), 0).ToString          'Omega asymptote
            TextBox39.Text = Math.Round((Rad_to_hz(ω_asym) * 60), 0).ToString   'Omega asymptote

            max_api_673 = Rad_to_hz(ω_krit1) * 60 / 1.2

            TextBox62.Text = Math.Round(max_api_673, 0).ToString   'Max speed API 673 [rpm]

            TextBox30.Text = I1_shaft.ToString((("0.00 E0")))      'Buigtraagheidsmoment as  [m^4]
            TextBox31.Text = I2_overhung.ToString((("0.00 E0")))   'Buigtraagheidsmoment overhung as [m^4]

            ' ------- Check sanity -------
            If L3 > L2 Then
                NumericUpDown3.BackColor = Color.Red
            Else
                NumericUpDown3.BackColor = Color.Yellow
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Draw_chart1()
        Dim hh, limit As Integer
        Dim ω10, ω20, krit1 As Double
        Dim px, py As Double

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
                Chart1.Titles.Add("Campbell diagram, overhung, isotropic Short bearings, flex shaft, no damping")
            Else
                Chart1.Titles.Add("Campbell diagram, between bearing, isotropic Short bearings, flex shaft, no damping")
            End If
            Chart1.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)

            '--------------- Legends and titles ---------------
            If CheckBox5.Checked Then
                Chart1.ChartAreas("ChartArea0").AxisX.Title = "Speed impeller [rpm]"
                Chart1.ChartAreas("ChartArea0").AxisY.Title = "Eigenfrequenty [rpm]"
            Else
                Chart1.ChartAreas("ChartArea0").AxisX.Title = "Angular speed impeller [rad/s]"
                Chart1.ChartAreas("ChartArea0").AxisY.Title = "Eigenfrequenty [rad/s]"
            End If

            Chart1.ChartAreas("ChartArea0").AxisY.RoundAxisValues()
            Chart1.ChartAreas("ChartArea0").AxisX.RoundAxisValues()

            limit = CInt(NumericUpDown22.Value)     'Limit in [rad/s]
            If CheckBox5.Checked Then
                limit = CInt(Convert_to_rpm(limit)) 'Limit in [rpm]
            End If


            '--------- Chart min sI_polare---------------
            If CheckBox1.Checked Then                           'Flip
                Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Else
                Chart1.ChartAreas("ChartArea0").AxisX.Minimum = -limit
            End If

            '--------- Chart max sI_polare---------------
            Chart1.ChartAreas("ChartArea0").AxisX.Maximum = limit
            Chart1.ChartAreas("ChartArea0").AxisY.Maximum = limit
            ' Chart1.ChartAreas("ChartArea0").AlignmentOrientation = DataVisualI_polaration.Charting.AreaAlignmentOrientations.Vertical

            '-------- snijpunten -----------

            Double.TryParse(TextBox34.Text, ω10)    '[Rad/sec] Omega 10 
            Double.TryParse(TextBox35.Text, ω20)    '[Rad/sec] Omega 20 
            Double.TryParse(TextBox32.Text, krit1)  '[Rad/sec] Kritisch1 

            If CheckBox5.Checked Then
                ω10 = Convert_to_rpm(ω10)           '[rpm] Omega 10 
                ω20 = Convert_to_rpm(ω20)           '[rpm] Omega 20 
                krit1 = Convert_to_rpm(krit1)       '[rpm] Kritisch1 
            End If


            Chart1.Series(2).Points.AddXY(0, ω10)
            Chart1.Series(3).Points.AddXY(0, ω20)
            Chart1.Series(4).Points.AddXY(krit1, krit1)

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

            '------- Chart

            For hh = 1 To 2000                      'Array sI_polare
                px = form533(hh, 1)                 '[rad/s] Plot x coordinate
                py = form533(hh, 0)                 '[rad/s] Plot y coordinate

                If CheckBox5.Checked Then
                    px = Convert_to_rpm(px)
                    py = Convert_to_rpm(py)
                End If

                If px < limit And py > 0 Then
                    If CheckBox1.Checked Then
                        Chart1.Series(1).Points.AddXY(Abs(px), py)
                    Else
                        Chart1.Series(1).Points.AddXY(px, py)
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
    Private Function Convert_to_rpm(rps As Double) As Double
        'Convert [rad/s] to [rpm]
        Return (rps * 60 / (2 * PI))
    End Function

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        Sync_data()
        Calc_sequence()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, TabPage4.Enter, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown12.ValueChanged, NumericUpDown18.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown19.ValueChanged
        Sync_data()
        Calc_sequence()
    End Sub
    Private Sub Sync_data()
        Dim yng As Integer
        If CheckBox4.Checked Then
            NumericUpDown15.Value = NumericUpDown4.Value    'Weight
            Int32.TryParse(TextBox61.Text, yng)

            NumericUpDown12.Value = NumericUpDown1.Value + NumericUpDown2.Value
            NumericUpDown13.Value = NumericUpDown2.Value
            NumericUpDown16.Value = NumericUpDown8.Value
            NumericUpDown17.Value = yng                     'Young

            NumericUpDown14.Value = NumericUpDown1.Value
            NumericUpDown18.Value = NumericUpDown2.Value
            NumericUpDown19.Value = NumericUpDown8.Value

            NumericUpDown12.Enabled = False
            NumericUpDown13.Enabled = False
            NumericUpDown14.Enabled = False
            NumericUpDown15.Enabled = False
            NumericUpDown16.Enabled = False
            NumericUpDown17.Enabled = False
            NumericUpDown18.Enabled = False
            NumericUpDown19.Enabled = False
        Else
            NumericUpDown12.Enabled = True
            NumericUpDown13.Enabled = True
            NumericUpDown14.Enabled = True
            NumericUpDown15.Enabled = True
            NumericUpDown16.Enabled = True
            NumericUpDown17.Enabled = True
            NumericUpDown18.Enabled = True
            NumericUpDown19.Enabled = True
        End If
    End Sub

    Private Sub Calc_simple()
        Dim length_L, length_A, length_B, mmassa, diam_tussen, diam_overhung, young As Double
        Dim C_tussen, I_as_tussen, I_as_overhung, fr_krit As Double

        length_L = NumericUpDown12.Value / 1000
        length_A = NumericUpDown13.Value / 1000
        length_B = length_L - length_A
        mmassa = NumericUpDown15.Value                                  '[kg]
        diam_tussen = NumericUpDown16.Value / 1000                      '[m]

        young = NumericUpDown17.Value * 10 ^ 9                          '[N/m2]

        '-------------- Tussen de lagers -----------------
        I_as_tussen = PI / 4 * (diam_tussen / 2) ^ 4                    '[m4]
        C_tussen = 3 * young * I_as_tussen * length_L
        C_tussen /= (length_A ^ 2 * length_B ^ 2)
        fr_krit = Sqrt(C_tussen / mmassa)                               '[Rad/sec]
        fr_krit /= (2 * PI)                                             '[Hz]

        TextBox17.Text = Round(length_B * 1000, 0).ToString
        'TextBox18.Text = I_as_tussen.ToString("0.00 E0")
        TextBox18.Text = (I_as_tussen * 1000 ^ 4).ToString("F0")
        TextBox19.Text = Round(C_tussen / 1000).ToString                'Buigstijfheid [kN/m]
        TextBox20.Text = Round(fr_krit, 0).ToString                     '[Hz]
        TextBox27.Text = Round(fr_krit * 60, 0).ToString                '[rpm]


        '-------------- Overhung -----------------
        Dim Overhung_L, Overhung_A, C_Overhung, fr_krit_overhung As Double

        diam_overhung = NumericUpDown19.Value / 1000                    '[m]
        Overhung_L = NumericUpDown14.Value / 1000
        Overhung_A = NumericUpDown18.Value / 1000                       'Overhung

        I_as_overhung = PI / 4 * (diam_overhung / 2) ^ 4                '[m4]
        C_Overhung = 3 * young * I_as_overhung
        C_Overhung /= (Overhung_A ^ 2 * (Overhung_A + Overhung_L))

        fr_krit_overhung = Sqrt(C_Overhung / mmassa)                    '[Rad/sec]
        fr_krit_overhung /= (2 * PI)                                    '[Hz]

        'TextBox26.Text = I_as_overhung.ToString("0.00 E0")
        TextBox26.Text = (I_as_overhung * 1000 ^ 4).ToString("F0")      '[mm4]
        TextBox21.Text = (C_Overhung / 1000).ToString("F2")             'Buigstijfheid [kN/m]
        TextBox22.Text = (fr_krit_overhung).ToString("F0")              '[Hz]   
        TextBox28.Text = (fr_krit_overhung * 60).ToString("F0")         '[rpm]

        '---------------- Check lengtes --------------------
        If length_A > length_L * 0.95 Then   'Residual torque too big,  problem in choosen bouderies
            NumericUpDown13.BackColor = Color.Red
        Else
            NumericUpDown13.BackColor = SystemColors.Window
        End If
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, TabPage5.Enter, NumericUpDown21.ValueChanged, NumericUpDown20.ValueChanged, NumericUpDown25.ValueChanged, NumericUpDown24.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown26.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown31.ValueChanged, NumericUpDown30.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown28.ValueChanged, NumericUpDown71.ValueChanged, NumericUpDown70.ValueChanged
        Dim Dia_imp, radius, hoog, massa As Double
        Dim I_polar As Double   'Cylinder spin around the center line
        Dim I_diam As Double    'Cylinder spin around the diameter
        Dim sp1, sp2, spc As Double


        Dia_imp = NumericUpDown20.Value / 1000          '[m]
        hoog = NumericUpDown21.Value / 1000             '[m]
        radius = Dia_imp / 2                            '[m]

        massa = PI / 4 * Dia_imp ^ 2 * hoog * 7800      'Staal
        '---- cylinder Polar moment of inertia ------
        I_polar = 0.5 * massa * (radius) ^ 2

        '---- cylinder Diametral moment of inertia -----
        I_diam = massa / 12 * (3 * (radius) ^ 2 + hoog ^ 2)

        TextBox23.Text = IIf(I_polar < 50, Round(I_polar, 3), Round(I_polar, 1)).ToString
        TextBox24.Text = IIf(I_diam < 50, Round(I_diam, 3), Round(I_diam, 1)).ToString
        TextBox25.Text = IIf(massa < 10, Round(massa, 1), Round(massa, 0)).ToString

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


        '--------- Gyroscopic couple -------
        'Dynamics of rotating machines page 79
        'ISBN 978-0-521-85016-2 Cambride University Press
        Dim rpm As Double           '[rpm] speed
        Dim ω As Double             '[rad/s] speed
        Dim ang_moment As Double    '[kg.m2/s] Angular moment
        Dim δψδt As Double          '[rad/s] Angle speed (δψ/δt) 
        Dim δψ As Double            '[rad] tilt angle
        Dim δt As Double            '[s] time one revolution
        Dim tilt_couple As Double   '[Nm]
        Dim runout As Double        '[mm] peak-peak impeller runout

        rpm = NumericUpDown71.Value
        ω = rpm * 2 * PI / 60       '[rad/s]

        '----- Angular moment calculation ----
        ang_moment = ω * I_polar            '[kg.m2/s] Angular moment

        '----- Angle speed (δψ/δt) [rad/s] -----
        δψ = (NumericUpDown70.Value / 360) * 2 * PI '[rad]
        δt = 60 / rpm                       '[s] time one turn
        δψδt = δψ / δt                      '[rad/s]

        '----- Angular moment  ----
        tilt_couple = ang_moment * δψδt     '[Nm]

        '----- bearing reaction forces -----
        '----- only between bearing case----
        Dim Ra, Rb As Double    '[N] bearing reaction force
        Dim L, A As Double      '[m]

        L = (NumericUpDown1.Value + NumericUpDown2.Value) / 1000    '[m]
        A = NumericUpDown2.Value / 1000                             '[m]

        If RadioButton2.Checked Then
            Ra = (tilt_couple / L) * (L - A) / L
            Rb = (tilt_couple / L) * A / L
        Else
            Ra = 0
            Rb = 0
        End If

        '----- runout -----
        runout = 2 * Asin(δψ) * (Dia_imp * 0.5)     '[m]

        '===== this effect proved to be very small
        Dim θs As Double            '[rad] angular deflection impeller due to tilt couple
        θs = Calc_shaft_angular_displacement(tilt_couple)

        '----------- Present data ----------
        TextBox43.Text = eigenfreq2.ToString("F0")
        TextBox104.Text = ω.ToString("F0")
        TextBox103.Text = ang_moment.ToString("F0")
        TextBox105.Text = δψδt.ToString("F2")               '[rad/s]
        TextBox107.Text = tilt_couple.ToString("F0")        '[Nm]
        TextBox108.Text = (runout * 1000).ToString("F1")    '[mm]
        TextBox111.Text = θs.ToString("E1")                 '[rad] 

        TextBox110.Text = Ra.ToString("F0")        '[N] Non Drive End
        TextBox109.Text = Rb.ToString("F0")        '[N] Drive end

    End Sub
    Private Function Calc_shaft_angular_displacement(tilt_couple As Double) As Double
        '---------- Couple induced angular displacement ------------
        '-- Impeller tilt gives couple, couple gives extra tilt 
        'Roark's Formulas, 8th edition page 216, reference 3e
        'Roark's Formulas, 8th edition page 974, 2nd moment of area
        'Tilt couple in [N.m]
        Dim ix As Double            'Area moment of inertia
        Dim r_shaft As Double       '[mm2]
        Dim L, A As Double          '[mm]
        Dim Mo As Double            '[N.mm]
        Dim θs As Double            '[rad] angular deflection
        Dim Elas As Double
        Dim y1 As Double

        If TextBox61.Text.Length > 0 Then   'preventing problems as startup
            Double.TryParse(TextBox61.Text, Elas)
            Elas *= 1000                            '[N/mm]
            L = NumericUpDown1.Value + NumericUpDown2.Value '[mm]
            A = NumericUpDown2.Value                '[mm]
            r_shaft = NumericUpDown8.Value / 2      '[mm]
            ix = PI / 4 * r_shaft ^ 4               '[mm4]
            Mo = tilt_couple * 10 ^ -3              '[N.mm]

            y1 = (2 * L ^ 2 - 6 * A * L + 3 * A ^ 2)
            θs = (Mo * y1) / (6 * Elas * ix * L)

        End If

        Return (θs)
    End Function


    'Converts Radial per second to Hz
    Private Function Rad_to_hz(rads As Double) As Double
        Return (rads / (2 * PI))
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim oWord As Word.Application ' = Nothing

        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2, oPara4 As Word.Paragraph
        Dim row, font_sI_polarze As Integer
        Dim ufilename As String
        ufilename = "Campbell_Calculation_" & TextBox7.Text & "_" & TextBox8.Text & "_" & DateTime.Now.ToString("yyyy_MM_dd") & ".docx"

        Try
            oWord = New Word.Application()

            'Start Word and open the document template. 
            font_sI_polarze = 9
            oWord = CType(CreateObject("Word.Application"), Word.Application)
            oWord.Visible = True
            oDoc = oWord.Documents.Add

            oDoc.PageSetup.TopMargin = 35
            oDoc.PageSetup.BottomMargin = 15

            'Insert a paragraph at the beginning of the document. 
            oPara1 = oDoc.Content.Paragraphs.Add

            oPara1.Range.Text = "VTK Engineering"
            oPara1.Range.Font.Name = "Arial"
            oPara1.Range.Font.Size = Font.Size + 3
            oPara1.Range.Font.Bold = CInt(True)
            oPara1.Format.SpaceAfter = 1                '24 pt spacing after paragraph. 
            oPara1.Range.InsertParagraphAfter()

            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara2.Range.Font.Size = font_sI_polarze + 1
            oPara2.Format.SpaceAfter = 1
            oPara2.Range.Font.Bold = CInt(False)
            oPara2.Range.Text = "Campbell diagram (based On Maschinendynamik, 11 Auflage, ISBN 978-3-642-29570-6)" & vbCrLf
            oPara2.Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 6, 2)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sI_polarze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

            row = 1
            oTable.Cell(row, 1).Range.Text = "Project Name"
            oTable.Cell(row, 2).Range.Text = TextBox7.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Project number "
            oTable.Cell(row, 2).Range.Text = TextBox8.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Fan type"
            oTable.Cell(row, 2).Range.Text = TextBox9.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Author"
            oTable.Cell(row, 2).Range.Text = Environment.UserName
            row += 1
            oTable.Cell(row, 1).Range.Text = "Date"
            oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            row += 1
            oTable.Cell(row, 1).Range.Text = "File name"
            oTable.Cell(row, 2).Range.Text = ufilename

            oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(4)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a 18 (row) x 3 table (column), fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 19, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sI_polarze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sI_polarze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Input Data"
            row += 1
            If RadioButton1.Checked Then
                oTable.Cell(row, 1).Range.Text = "Fan type"
                oTable.Cell(row, 2).Range.Text = "Overhung"
            ElseIf RadioButton2.Checked Then
                oTable.Cell(row, 1).Range.Text = "Fan type"
                oTable.Cell(row, 2).Range.Text = "Between bearings"
            End If

            row += 2
            oTable.Rows.Item(4).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(4).Range.Font.Size = font_sI_polarze
            oTable.Cell(row, 1).Range.Text = "Fan Housing"

            If RadioButton1.Checked Then    'Overhung "
                row += 1
                oTable.Cell(row, 1).Range.Text = "L1, Bearing distance"
                oTable.Cell(row, 2).Range.Text = CType(NumericUpDown1.Value, String)
                oTable.Cell(row, 3).Range.Text = "[mm]"

                row += 1
                oTable.Cell(row, 1).Range.Text = "L2, Overhung length incl. L3"
                oTable.Cell(row, 2).Range.Text = CType(NumericUpDown2.Value, String)
                oTable.Cell(row, 3).Range.Text = "[mm]"
                row += 1
                oTable.Cell(row, 1).Range.Text = "L3, Rigid length in impeller"
                oTable.Cell(row, 2).Range.Text = CType(NumericUpDown3.Value, String)
                oTable.Cell(row, 3).Range.Text = "[mm]"
            Else                            'Between bearings
                row += 1
                oTable.Cell(row, 1).Range.Text = "Shaft length DE Bearing -- impeller (drive side)"
                oTable.Cell(row, 2).Range.Text = CType(NumericUpDown1.Value, String)
                oTable.Cell(row, 3).Range.Text = "[mm]"
                row += 1
                oTable.Cell(row, 1).Range.Text = "Shaft length NDE Bearing -- impeller"
                oTable.Cell(row, 2).Range.Text = CType(NumericUpDown2.Value, String)
                oTable.Cell(row, 3).Range.Text = "[mm]"
            End If

            row += 1
            oTable.Cell(row, 1).Range.Text = "Weight impeller"
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown4.Value, String)
            oTable.Cell(row, 3).Range.Text = "[kg]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Support"
            Select Case True
                Case RadioButton3.Checked
                    oTable.Cell(row, 2).Range.Text = "Steel frame on rubber"
                Case RadioButton4.Checked
                    oTable.Cell(row, 2).Range.Text = "Bearings on concrete"
                Case RadioButton5.Checked
                    oTable.Cell(row, 2).Range.Text = "Steel frame on concrete"
            End Select

            row += 1
            oTable.Cell(row, 1).Range.Text = "C1 Stiffness bearing Drive End"
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown6.Value, String)
            oTable.Cell(row, 3).Range.Text = "[kN/mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "C2 Stiffness bearing NDE side"
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown7.Value, String)
            oTable.Cell(row, 3).Range.Text = "[kN/mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft dia. between bearings"
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown8.Value, String)
            oTable.Cell(row, 3).Range.Text = "[mm]"

            If RadioButton1.Checked Then    'Overhung "
                row += 1
                oTable.Cell(row, 1).Range.Text = "Shaft dia. overhung"
                oTable.Cell(row, 2).Range.Text = CType(NumericUpDown9.Value, String)
                oTable.Cell(row, 3).Range.Text = "[mm]"
            End If

            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft young's modulus"
            oTable.Cell(row, 2).Range.Text = TextBox61.Text
            oTable.Cell(row, 3).Range.Text = "[kN/mm]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft max. operating temp."
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown55.Value, String)
            oTable.Cell(row, 3).Range.Text = "[c]"

            row += 2
            oTable.Rows.Item(row).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(row).Range.Font.Size = font_sI_polarze
            oTable.Cell(row, 1).Range.Text = "Rotation Inertia impeller"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Ja, inertia (radial line)"
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown11.Value, String)
            oTable.Cell(row, 3).Range.Text = "[kg.m2]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Jp, inertia (center line)"
            oTable.Cell(row, 2).Range.Text = CType(NumericUpDown10.Value, String)
            oTable.Cell(row, 3).Range.Text = "[kg.m2]"

            oTable.Columns(1).Width = oWord.InchesToPoints(3.0)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.2)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '----------------------------------------------------------------------
            'Insert a 5 (row) x 7 (column) table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 6, 7)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sI_polarze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sI_polarze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Results"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft area moment of inertia"
            oTable.Cell(row, 2).Range.Text = TextBox30.Text
            oTable.Cell(row, 3).Range.Text = "[mm^4]"

            If RadioButton1.Checked Then    'Overhung
                row += 1
                oTable.Cell(row, 1).Range.Text = "Overhung shaft m.o.i."
                oTable.Cell(row, 2).Range.Text = TextBox31.Text
                oTable.Cell(row, 3).Range.Text = "[mm^4]"
            End If

            row += 1
            oTable.Cell(row, 1).Range.Text = "Omega critical #1"
            oTable.Cell(row, 2).Range.Text = TextBox32.Text
            oTable.Cell(row, 3).Range.Text = "[rad/s]"
            oTable.Cell(row, 4).Range.Text = TextBox5.Text
            oTable.Cell(row, 5).Range.Text = "[Hz]"
            oTable.Cell(row, 6).Range.Text = TextBox1.Text
            oTable.Cell(row, 7).Range.Text = "[rpm]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Omega critical #2"
            oTable.Cell(row, 2).Range.Text = TextBox33.Text
            oTable.Cell(row, 3).Range.Text = "[rad/s]"
            oTable.Cell(row, 4).Range.Text = TextBox6.Text
            oTable.Cell(row, 5).Range.Text = "[Hz]"
            oTable.Cell(row, 6).Range.Text = TextBox13.Text
            oTable.Cell(row, 7).Range.Text = "[rpm]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Maximum speed acc API 673"
            oTable.Cell(row, 2).Range.Text = TextBox62.Text
            oTable.Cell(row, 3).Range.Text = "[rpm]"


            oTable.Columns(1).Width = oWord.InchesToPoints(1.8)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(0.9)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.7)    '"[rad/s]"
            oTable.Columns(4).Width = oWord.InchesToPoints(0.7)
            oTable.Columns(5).Width = oWord.InchesToPoints(0.4)    '"[Hz]"
            oTable.Columns(6).Width = oWord.InchesToPoints(0.5)
            oTable.Columns.Item(7).Width = oWord.InchesToPoints(0.45)   '"[rpm]"
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------save picture ---------------- 
            Chart1.SaveImage("c:\Temp\MainChart.gif", System.Drawing.Imaging.ImageFormat.Gif)
            oPara4 = oDoc.Content.Paragraphs.Add
            oPara4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oPara4.Range.InlineShapes.AddPicture("c:\Temp\MainChart.gif")
            oPara4.Range.InlineShapes.Item(1).LockAspectRatio = CType(True, Microsoft.Office.Core.MsoTriState)
            oPara4.Range.InlineShapes.Item(1).Width = 310
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            '--------------Save file word file------------------
            'See https://msdn.microsoft.com/en-us/library/63w57f4b.aspx

            If (Not System.IO.Directory.Exists(dirpath_Eng)) Then System.IO.Directory.CreateDirectory(dirpath_Eng)
            If (Not System.IO.Directory.Exists(dirpath_Rap)) Then System.IO.Directory.CreateDirectory(dirpath_Rap)
            If (Not System.IO.Directory.Exists(dirpath_Home)) Then System.IO.Directory.CreateDirectory(dirpath_Home)

            If Directory.Exists(dirpath_Rap) Then
                oWord.ActiveDocument.SaveAs(dirpath_Rap & ufilename)
            Else
                oWord.ActiveDocument.SaveAs(dirpath_Home & ufilename)
            End If

        Catch ex As Exception
            MessageBox.Show(ufilename & vbCrLf & ex.Message)  ' Show the exception's message.
        End Try

    End Sub

    Private Sub Calc_rolling_element_bearings()
        'Based on Dynamics of Rotary Machines page 183, ISBN 978-0-521-85016-2

        Dim dia_ball, length_roller, alfa, no_balls, no_rollers As Double
        Dim K_ball, K_roller As Double
        Dim Kvv_ball, Kvv_roller As Double 'vertical stiffness
        Dim Kuu_ball, Kuu_roller As Double 'horI_polarontal stiffness
        Dim force_ball, force_roller, row_ball, row_roller As Double

        K_ball = 13 * 10 ^ 6        '[N^2/3.m^-4/3]
        K_roller = 1.0 * 10 ^ 9     '[N^0,9.m^-1.8]
        alfa = Math.PI * 0 / 180     'Pressure angle [radials]

        NumericUpDown43.Value = CDec(Round(NumericUpDown35.Value * 0.9, 0)) 'Ball diameter is 90% van de lager breedte
        NumericUpDown32.Value = CDec(Round(NumericUpDown36.Value * 0.9, 0)) 'Total rollers length is 90% van de lager breedte
        dia_ball = NumericUpDown43.Value / 1000             '[m]
        length_roller = NumericUpDown32.Value / 1000        '[m]
        force_ball = NumericUpDown40.Value * 1000           '[N]
        force_roller = NumericUpDown45.Value * 1000         '[N]
        Double.TryParse(CType(ComboBox1.SelectedItem, String), no_balls)
        Double.TryParse(CType(ComboBox2.SelectedItem, String), no_rollers)

        'Page 183, Equation 5.89, vertical stiffness, in SI units
        Kvv_ball = K_ball * no_balls ^ (2 / 3) * dia_ball ^ (1 / 3) * force_ball ^ (1 / 3) * Cos(alfa) ^ (5 / 3)   '[N/m]
        Kvv_ball /= 10 ^ 6              '[kN/mm]
        Select Case no_balls
            Case 8
                Kuu_ball = Kvv_ball * 0.46
            Case 12
                Kuu_ball = Kvv_ball * 0.64
            Case 16
                Kuu_ball = Kvv_ball * 0.73
        End Select

        'Page 183, Equation 5.90, vertical stiffness, in SI units
        Kvv_roller = K_roller * no_rollers ^ 0.9 * length_roller ^ 0.8 * force_roller ^ 0.1 * Cos(alfa) ^ 1.9 '[N/m]
        Kvv_roller /= 10 ^ 6            '[kN/mm]   

        Select Case no_rollers
            Case 8
                Kuu_roller = Kvv_roller * 0.49
            Case 12
                Kuu_roller = Kvv_roller * 0.66
            Case 16
                Kuu_roller = Kvv_roller * 0.74
        End Select

        row_ball = NumericUpDown34.Value                        'Single of double row bearing
        row_roller = NumericUpDown41.Value                      'Single of double row bearing
        TextBox46.Text = (Kvv_ball * row_ball).ToString("F0")    'Vertical stiffness x row
        TextBox56.Text = (Kuu_ball * row_ball).ToString("F0")    'horI_polarontal stiffness x row

        TextBox47.Text = (Kvv_roller * row_roller).ToString("F0") 'Vertical stiffness
        TextBox57.Text = (Kuu_roller * row_roller).ToString("F0") 'horI_polarontal stiffness
    End Sub
    Private Sub Calc_dydrodynamic_bearing()
        Dim dia, omega, visco, length, f_load, clearance As Double
        Dim Sommerfeld As Double

        Try
            dia = NumericUpDown49.Value / 1000              '[m]
            omega = NumericUpDown42.Value * 2 * PI / 60     '[rad/s]
            visco = NumericUpDown44.Value / 1000            '[Pa.s]
            length = NumericUpDown48.Value / 1000           '[m]
            f_load = NumericUpDown46.Value * 1000           '[kN]

            '------- Clearance 0.1 - 0.2% of journal diameter--------
            NumericUpDown47.Value = NumericUpDown49.Value * NumericUpDown33.Value / 100
            clearance = NumericUpDown47.Value / 1000        '[m]

            'Formula 5.84
            Sommerfeld = dia * omega * visco * length ^ 3 / (8 * f_load * clearance ^ 2)

            If Not Double.IsNaN(Sommerfeld) Then
                Itterate(Sommerfeld)
                TextBox50.Text = Math.Round(Sommerfeld, 5).ToString
            End If
        Catch ex As Exception
            MessageBox.Show("Line 753 " & ex.Message)  ' Show the exception's message.
        End Try
    End Sub
    Private Sub Itterate(sommerf As Double)
        Dim Ecc1, Ecc2, Ecc3, Exc_fin2 As Double
        Dim Dev1, Dev2, Dev3 As Double
        Dim H_nul, Kvv, Kuu, force, clearance As Double
        Dim Cvv, Cuu, rads As Double
        Dim jjr As Integer

        force = NumericUpDown46.Value * 1000        '[N]
        clearance = NumericUpDown47.Value / 1000    '[m]
        rads = NumericUpDown42.Value * 2 * PI / 60  '[rad/sec]

        Ecc1 = 0        'Start lower limit of eccentricity [-]
        Ecc2 = 1.0      'Start upper limit of eccentricity [-]
        Ecc3 = 0.5      'In the middle of eccentricity [-]

        Dev1 = CDbl(Calc_epsilon(sommerf, Ecc1))
        Dev2 = CDbl(Calc_epsilon(sommerf, Ecc2))
        Dev3 = CDbl(Calc_epsilon(sommerf, Ecc3))

        '-------------Iteratie 30x halveren moet voldoende zijn ---------------
        '---------- Exc= excentricity, looking for Deviation is zero ---------

        For jjr = 0 To 30
            If Dev1 * Dev3 < 0 Then
                Ecc2 = Ecc3
            Else
                Ecc1 = Ecc3
            End If
            Ecc3 = (Ecc1 + Ecc2) / 2
            Dev1 = CDbl(Calc_epsilon(sommerf, Ecc1))
            Dev2 = CDbl(Calc_epsilon(sommerf, Ecc2))
            Dev3 = CDbl(Calc_epsilon(sommerf, Ecc3))
        Next jjr
        TextBox49.Text = Round(Ecc3, 2).ToString

        '-------- Controle nulpunt zoek functie ----------------
        If Dev3 > 0.01 Then
            TextBox49.BackColor = Color.Red
        Else
            TextBox49.BackColor = Color.LightGreen
        End If

        If Ecc3 < 0.6 Or Ecc3 > 0.7 Then
            TextBox49.BackColor = Color.Red
        Else
            TextBox49.BackColor = Color.LightGreen
        End If

        'Dynamics of Rotating Machines ISBN 9780511780509, page 179

        '-------------- Vertical stiffness------------------------

        Exc_fin2 = Ecc3 ^ 2
        H_nul = 1 / ((PI ^ 2 * (1 - Exc_fin2) + 16 * Exc_fin2) ^ 1.5)
        Kvv = H_nul * 4 * (PI ^ 2 * (1 + 2 * Exc_fin2) + 32 * Exc_fin2 * (1 + Exc_fin2) / (1 - Exc_fin2))
        Kvv = Kvv * force / clearance       '[N/m]
        Kvv /= 10 ^ 6                       '[kN/mm]

        '-------------- HorI_polarontal stiffness -------------------------
        Kuu = H_nul * 4 * (PI ^ 2 * (2 - Exc_fin2) + 16 * Exc_fin2)
        Kuu = Kuu * force / clearance       '[N/m]
        Kuu /= 10 ^ 6                       '[kN/mm]

        '-------------- Vertical damping -------------------------
        Cvv = H_nul * (2 * PI * (PI ^ 2 * (1 - Exc_fin2) ^ 2 + 48 * Exc_fin2)) / (Ecc3 * (1 - Exc_fin2) ^ 0.5)
        Cvv = Cvv * force / (clearance * rads)  '[Ns/m]
        Cvv /= 10 ^ 3                           '[kNs/m]

        '-------------- Vertical damping -------------------------
        Cuu = H_nul * (2 * PI * (1 - Exc_fin2) ^ 0.5 * (PI ^ 2 * (1 + 2 * Exc_fin2) - 16 * Exc_fin2)) / Ecc3
        Cuu = Cuu * force / (clearance * rads)  '[Ns/m]
        Cuu /= 10 ^ 3                           '[kNs/m]

        TextBox48.Text = Round(Kvv, 1).ToString     'Vertical stiffness [kN/mm]
        TextBox16.Text = Round(Kuu, 1).ToString     'HorI_polarontal stiffness [kN/mm]

        TextBox44.Text = Round(Cvv, 1).ToString     'Vertical damping [kNs/m]
        TextBox45.Text = Round(Cuu, 1).ToString     'HorI_polarontal damping [kNs/m]

        Draw_Chart2(sommerf)
    End Sub

    Private Function Calc_epsilon(sommerf As Double, eps As Double) As Double
        Dim deviation, som2 As Double

        som2 = sommerf ^ 2

        'Dynamics of Rotating Machines ISBN 9780511780509, equation (5.83) page 178
        'Hydrodunamic journal bearings, bearing eccentric
        deviation = eps ^ 8 - 4 * eps ^ 6 + (6 - som2 * (16 - PI ^ 2)) * eps ^ 4 - (4 + PI ^ 2 * som2) * eps ^ 2 + 1

        Return (deviation)
    End Function

    Private Sub Draw_Chart2(sommerf As Double)
        Dim x, y As Double
        Try
            'Clear all series And chart areas so we can re-add them
            Chart2.Series.Clear()
            Chart2.ChartAreas.Clear()
            Chart2.Titles.Clear()
            Chart2.Series.Add("Series0")
            Chart2.ChartAreas.Add("ChartArea0")
            Chart2.Series(0).ChartArea = "ChartArea0"
            Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart2.Titles.Add("Determine Eccentricity" & vbCrLf & "Formula 5.83= 0.0 ")
            Chart2.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)
            Chart2.Series(0).Name = "Koppel[%]"
            Chart2.Series(0).Color = Color.Blue
            Chart2.Series(0).IsVisibleInLegend = False
            Chart2.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Chart2.ChartAreas("ChartArea0").AxisX.Maximum = 1
            Chart2.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True
            Chart2.ChartAreas("ChartArea0").AxisY.Title = "Formula 5.83 [-]"
            Chart2.ChartAreas("ChartArea0").AxisX.Title = "Eccentricity [-]"

            For x = 0 To 1.01 Step 0.01
                y = CDbl(Calc_epsilon(sommerf, x))
                Chart2.Series(0).Points.AddXY(x, y)
            Next x

        Catch ex As Exception
            'MessageBox.Show(ex.Message &"Error 845")  ' Show the exception's message.
        End Try

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, TabPage3.Enter, NumericUpDown49.ValueChanged, NumericUpDown48.ValueChanged, NumericUpDown47.ValueChanged, NumericUpDown46.ValueChanged, NumericUpDown45.ValueChanged, NumericUpDown44.ValueChanged, NumericUpDown43.ValueChanged, NumericUpDown42.ValueChanged, NumericUpDown40.ValueChanged, NumericUpDown32.ValueChanged, NumericUpDown33.ValueChanged, NumericUpDown36.ValueChanged, NumericUpDown35.ValueChanged, ComboBox1.SelectedIndexChanged, NumericUpDown41.ValueChanged, NumericUpDown34.ValueChanged
        'Dynamics of Rotating Machines , page 178 

        Calc_rolling_element_bearings()
        Calc_dydrodynamic_bearing()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, TabPage6.Enter, NumericUpDown50.ValueChanged, NumericUpDown39.ValueChanged, NumericUpDown54.ValueChanged, NumericUpDown53.ValueChanged, NumericUpDown52.ValueChanged, NumericUpDown51.ValueChanged, NumericUpDown38.ValueChanged, NumericUpDown37.ValueChanged
        Dim p_pump, stiff_dimless, stiffness, area_recess As Double
        Dim clear_procent, h0_clearance, Journal_dia As Double
        Dim no_pockets, b_pocket, L_pocket As Double    'pocket dimensions
        Dim L_rim As Double                           'length next to the pocket
        Dim omtrek As Double                            'journal omtrek

        'Hydrostatic bearing orifice type


        '---------- pocket sI_polaring------------ 
        no_pockets = NumericUpDown38.Value              '[-] No pockets
        b_pocket = NumericUpDown53.Value / 1000         '[m] Pocket width
        Journal_dia = NumericUpDown51.Value / 1000      '[m] Journal diameter
        omtrek = PI * Journal_dia

        If no_pockets < 1 Then no_pockets = 2
        L_pocket = omtrek * 0.5 / no_pockets            '[m] pocket length

        NumericUpDown37.Value = CDec(L_pocket * 1000)   '[mm]

        stiff_dimless = NumericUpDown39.Value           '[-] Stiffness dimensionless
        clear_procent = NumericUpDown54.Value           '[-] Stiffness dimensionless

        L_rim = NumericUpDown52.Value / 1000            '[m] Bearing rim width

        h0_clearance = Journal_dia * clear_procent / 100  '[m]

        p_pump = NumericUpDown50.Value * 10 ^ 5     '[N/m2] lubricating pump pressure
        stiffness = no_pockets * 3 * L_pocket * (L_rim + b_pocket) * p_pump * stiff_dimless / h0_clearance
        area_recess = no_pockets * b_pocket * L_pocket * 10 ^ 6     '[pockets [mm2]

        TextBox51.Text = Round(stiffness / 10 ^ 6, 0).ToString          '[kN/mm]
        TextBox52.Text = CType((h0_clearance * 1000).ToString, String)  '[mm]
        TextBox53.Text = Round(area_recess, 0).ToString                 '[mm2]
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Save_tofile_vtk5()
    End Sub

    'Save control settings and case_x_conditions to file
    Private Sub Save_tofile_vtk5()
        Dim temp_string As String
        Dim filename As String = "Campbell_select_" & TextBox7.Text & "_" & TextBox8.Text & DateTime.Now.ToString("_yyyy_MM_dd") & ".vtk5"
        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
        Dim i As Integer

        If String.IsNullOrEmpty(TextBox8.Text) Then
            TextBox8.Text = "-"
        End If

        temp_string = TextBox7.Text & ";" & TextBox8.Text & ";" & TextBox9.Text & ";"
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all numeric controls -----------------
        FindControlRecursive(all_num, Me, GetType(NumericUpDown))   'Find the control
        all_num = all_num.OrderBy(Function(x) x.Name).ToList()      'Alphabetical order
        For i = 0 To all_num.Count - 1
            Dim numbt As NumericUpDown = CType(all_num(i), NumericUpDown)
            temp_string &= numbt.Name & ";" & numbt.Value.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all combobox controls and save
        FindControlRecursive(all_combo, Me, GetType(ComboBox))      'Find the control
        all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()   'Alphabetical order
        For i = 0 To all_combo.Count - 1
            Dim combt As ComboBox = CType(all_combo(i), ComboBox)
            temp_string &= combt.Name & ";" & combt.SelectedItem.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all checkbox controls and save
        FindControlRecursive(all_check, Me, GetType(CheckBox))      'Find the control
        all_check = all_check.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_check.Count - 1
            Dim chkbt As CheckBox = CType(all_check(i), CheckBox)
            temp_string &= chkbt.Name & ";" & chkbt.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all radio controls and save
        FindControlRecursive(all_radio, Me, GetType(RadioButton))   'Find the control
        all_radio = all_radio.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_radio.Count - 1
            Dim radbt As RadioButton = CType(all_radio(i), RadioButton)
            temp_string &= radbt.Name & ";" & radbt.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '--------- add notes -----
        temp_string &= TextBox63.Text & ";"

        '---- if path not exist then create one----------
        Try
            If (Not System.IO.Directory.Exists(dirpath_Home)) Then System.IO.Directory.CreateDirectory(dirpath_Home)
            If (Not System.IO.Directory.Exists(dirpath_Eng)) Then System.IO.Directory.CreateDirectory(dirpath_Eng)
            If (Not System.IO.Directory.Exists(dirpath_Rap)) Then System.IO.Directory.CreateDirectory(dirpath_Rap)
        Catch ex As Exception
        End Try

        Try
            If CInt(temp_string.Length.ToString) > 100 Then      'String may be empty
                If Directory.Exists(dirpath_Eng) Then
                    File.WriteAllText(dirpath_Eng & filename, temp_string, Encoding.ASCII)      'used at VTK
                Else
                    File.WriteAllText(dirpath_Home & filename, temp_string, Encoding.ASCII)     'used at home
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Line 5062, " & ex.Message)  ' Show the exception's message.
        End Try
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Read_file_vtk5()
        Calc_sequence()
    End Sub
    'Retrieve control settings and case_x_conditions from file
    'Split the file string into 5 separate strings
    'Each string represents a control type (combobox, checkbox,..)
    'Then split up the secton string into part to read into the parameters
    Private Sub Read_file_vtk5()
        Dim control_words(), words() As String

        Dim all_num As New List(Of Control)
        Dim all_combo As New List(Of Control)
        Dim all_check As New List(Of Control)
        Dim all_radio As New List(Of Control)
        Dim separators() As String = {";"}
        Dim separators1() As String = {"BREAK"}

        OpenFileDialog1.FileName = "Campbell*.vtk5"
        If Directory.Exists(dirpath_Eng) Then
            OpenFileDialog1.InitialDirectory = dirpath_Eng  'used at VTK
        Else
            OpenFileDialog1.InitialDirectory = dirpath_Home  'used at home
        End If

        OpenFileDialog1.Title = "Open a Text File"
        OpenFileDialog1.Filter = "VTK5 Files|*.vtk5"

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim readText As String = File.ReadAllText(OpenFileDialog1.FileName, Encoding.ASCII)
            control_words = readText.Split(separators1, StringSplitOptions.None) 'Split the read file content

            '----- retrieve case condition-----
            words = control_words(0).Split(separators, StringSplitOptions.None) 'Split first line the read file content
            TextBox7.Text = words(0)                  'Project number
            TextBox8.Text = words(1)                  'Item name
            TextBox9.Text = words(2)                  'Fan Type

            '---------- terugzetten numeric controls (Updated version) -----------------
            FindControlRecursive(all_num, Me, GetType(NumericUpDown))
            words = control_words(1).Split(separators, StringSplitOptions.None)     'Split the read file content
            Restore_num_controls(words, all_num)

            '---------- terugzetten combobox controls (Updated version) -----------------
            FindControlRecursive(all_combo, Me, GetType(ComboBox))
            words = control_words(2).Split(separators, StringSplitOptions.None)     'Split the read file content
            Restore_combo_controls(words, all_combo)

            '---------- terugzetten checkboxes controls (Updated version) -----------------
            FindControlRecursive(all_check, Me, GetType(CheckBox))
            words = control_words(3).Split(separators, StringSplitOptions.None)    'Split the read file content
            Restore_checkbox_controls(words, all_check)

            '---------- terugzetten Radio button controls (Updated version) -----------------
            FindControlRecursive(all_radio, Me, GetType(RadioButton))
            words = control_words(4).Split(separators, StringSplitOptions.None)    'Split the read file content
            Restore_radiobutton_controls(words, all_radio)

            '---------- terugzetten Notes -- ---------------
            If control_words.Count > 5 Then
                words = control_words(5).Split(separators, StringSplitOptions.None) 'Split the read file content
                TextBox63.Clear()
                TextBox63.AppendText(words(1))
            Else
                MessageBox.Show("Warning Notes not found in file")
            End If
        End If
    End Sub

    Private Sub Restore_num_controls(words As String(), all_num As List(Of Control))
        Dim ttt As Double

        For i = 0 To all_num.Count - 1
            Dim updown As NumericUpDown = CType(all_num(i), NumericUpDown)
            '============ find the stored numeric control list ====

            For j = 0 To all_num.Count - 1
                If (j * 2 + 2) < words.Count Then
                    If updown.Name = words(j * 2 + 1) Then    '==== Found ====
                        'Debug.WriteLine("FOUND !! updown.Name= " & updown.Name & ", words(j *2)= " & words(j * 2) & ", words(j*2+1)= " & words(j * 2 + 1) & ", words(j*2+2)= " & words(j * 2 + 2))
                        If Not (Double.TryParse(words(j * 2 + 2), ttt)) Then MessageBox.Show("Numeric controls conversion problem occured")
                        If ttt <= updown.Maximum And ttt >= updown.Minimum Then
                            updown.Value = CDec(ttt)          'OK
                        Else
                            updown.Value = updown.Minimum       'NOK
                            MessageBox.Show("Numeric controls value out of outside min-max range, Minimum value is used")
                        End If
                        Exit For
                    End If
                Else
                    MessageBox.Show(updown.Name & " was NOT Stored in file and is NOT updated")
                End If
            Next
        Next
    End Sub
    Private Sub Restore_combo_controls(words As String(), all_combo As List(Of Control))
        For i = 0 To all_combo.Count - 1
            Dim combobx As ComboBox = CType(all_combo(i), ComboBox)
            '============ find the stored numeric control list ====

            For j = 0 To all_combo.Count - 1
                If (j * 2 + 2) < words.Count Then
                    If combobx.Name = words(j * 2 + 1) Then    '==== Found ====
                        'Debug.WriteLine("FOUND !! combobx.Name= " & combobx.Name & ", words(j *2)= " & words(j * 2) & ", words(j*2+1)= " & words(j * 2 + 1) & ", words(j*2+2)= " & words(j * 2 + 2))
                        If (i < words.Length - 1) Then
                            combobx.SelectedItem = words(j * 2 + 2)
                        Else
                            MessageBox.Show("Warning last combobox not found in file")
                        End If
                        Exit For
                    End If
                Else
                    MessageBox.Show(combobx.Name & " was NOT Stored in file and is NOT updated")
                End If
            Next
        Next
    End Sub

    Private Sub Restore_checkbox_controls(words As String(), all_check As List(Of Control))
        For i = 0 To all_check.Count - 1
            Dim chbx As CheckBox = CType(all_check(i), CheckBox)
            '============ find the stored numeric control list ====

            For j = 0 To all_check.Count - 1
                If (j * 2 + 2) < words.Count Then
                    If chbx.Name = words(j * 2 + 1) Then    '==== Found ====
                        'Debug.WriteLine("FOUND !! chbx.Name= " & chbx.Name & ", words(j *2)= " & words(j * 2) & ", words(j*2+1)= " & words(j * 2 + 1) & ", words(j*2+2)= " & words(j * 2 + 2))
                        If CBool(words(j * 2 + 2)) = True Then
                            chbx.Checked = True
                        Else
                            chbx.Checked = False
                        End If

                        Exit For
                    End If
                Else
                    MessageBox.Show(chbx.Name & " was NOT Stored in file and is NOT updated")
                End If
            Next
        Next
    End Sub
    Private Sub Restore_radiobutton_controls(words As String(), all_radio As List(Of Control))
        For i = 0 To all_radio.Count - 1
            Dim radiobut As RadioButton = CType(all_radio(i), RadioButton)
            '============ find the stored numeric control list ====
            For j = 0 To all_radio.Count - 1
                If (j * 2 + 2) < words.Count Then
                    If radiobut.Name = words(j * 2 + 1) Then    '==== Found ====
                        'Debug.WriteLine("j= " & j.ToString & ", FOUND !! radiobut.Name= " & radiobt.Name & ", words(j *2)= " & words(j * 2) & ", words(j*2+1)= " & words(j * 2 + 1) & ", words(j*2+2)= " & words(j * 2 + 2))
                        Boolean.TryParse(words(j * 2 + 2), radiobut.Checked)
                        Exit For
                    End If
                Else
                    MessageBox.Show(radiobut.Name & " was NOT Stored in file and is NOT updated")
                End If
            Next
        Next
    End Sub

    '----------- Find all controls on form1------
    'Nota Bene, sequence of found control may be differen, List sort is required
    Public Shared Function FindControlRecursive(ByVal list As List(Of Control), ByVal parent As Control, ByVal ctrlType As System.Type) As List(Of Control)
        If parent Is Nothing Then Return list

        If parent.GetType Is ctrlType Then
            list.Add(parent)
        End If
        For Each child As Control In parent.Controls
            FindControlRecursive(list, child, ctrlType)
        Next
        Return list
    End Function

    Public Function HardDisc_Id() As String
        'Add system.management as reference !!
        Dim tmpStr2 As String = ""
        Dim myScop As New ManagementScope("\\" & Environment.MachineName & "\root\cimv2")
        Dim oQuer As New SelectQuery("SELECT * FROM WIN32_DiskDrive")

        Dim oResult As New ManagementObjectSearcher(myScop, oQuer)
        Dim oIte As ManagementObject
        Dim oPropert As PropertyData
        For Each oIte In oResult.Get()
            For Each oPropert In oIte.Properties
                If Not oPropert.Value Is Nothing AndAlso oPropert.Name = "SerialNumber" Then
                    tmpStr2 = oPropert.Value.ToString
                    Exit For
                End If
            Next
            Exit For
        Next
        Return (Trim(tmpStr2))         'Harddisk identification
    End Function
    'Young for steel S355 temperature dependent
    'https://www.engineeringtoolbox.com/young-modulus-d_773.html

    Public Function Calc_young(t As Double) As Double
        Dim young As Double
        young = -0.000000324 * t ^ 3 + 0.000049951 * t ^ 2 - 0.04930174 * t + 203.386
        Return (young * 1000)
    End Function
    'Çalculate natural frequency bearing support
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click, NumericUpDown57.ValueChanged, NumericUpDown56.ValueChanged, NumericUpDown66.ValueChanged, NumericUpDown67.ValueChanged, NumericUpDown64.ValueChanged, NumericUpDown69.ValueChanged, NumericUpDown65.ValueChanged
        Dim fan_weight As Double
        Dim stiff1 As Double
        Dim stiff2 As Double
        Dim freq, speed As Double
        Dim COG_factor As Double
        Dim Height_COG, Height_CL As Double

        fan_weight = NumericUpDown56.Value      '[kg]
        stiff1 = NumericUpDown57.Value * 10 ^ 6  '[kN/mm]->[N/m]

        '== NOTE the Spring constant is determined at the bearing !!===
        'at a lower position the spring is stiffer !

        Height_CL = NumericUpDown66.Value       'CL bearing house
        Height_COG = NumericUpDown65.Value      'COG bearing support

        COG_factor = Height_COG / Height_CL     'Height Ratio
        stiff2 = stiff1 / COG_factor            'Stiffness ar COG height

        freq = Sqrt(stiff2 / fan_weight)        '[rad/s]
        freq /= (2 * PI)                        '[Hz]
        speed = freq * 60                       '[rpm]

        TextBox68.Text = freq.ToString("F1")
        TextBox69.Text = speed.ToString("F0")               'Bending speed
        TextBox70.Text = (speed / 1.2).ToString("F0")       'Save speed (20% clearance)
        TextBox90.Text = COG_factor.ToString("F2")          'ratio
        TextBox67.Text = (stiff2 * 10 ^ -6).ToString("F1")  'ratio

        '------ check ---
        If COG_factor > 1.0 Then
            TextBox90.BackColor = Color.Red
        Else
            TextBox90.BackColor = Color.LightGreen
        End If

        '============= torsional stiffness =======
        'https://en.wikipedia.org/wiki/Stiffness
        'https://en.wikipedia.org/wiki/List_of_moments_of_inertia
        Dim radius As Double                            '[m]
        Dim R_stiff As Double                           '[N/m]
        Dim L_displac As Double                         '[rad] linear displacement
        Dim R_displac As Double                         '[rad] radial displacement
        Dim R_inertia As Double                         '[kg.m2]
        Dim R_force As Double                           '[N]
        Dim R_period, R_freq, R_speed As Double
        Dim wt As Double

        '----- get data from screen -----------
        R_force = NumericUpDown64.Value * 10 ^ 3        '[N] horI_polarontal force on bearing
        L_displac = NumericUpDown67.Value * 10 ^ -3     '[m] verplaatsing
        radius = NumericUpDown66.Value                  '[m] Centerline height

        '----- Radial dispacement ---------
        R_displac = L_displac / (PI * 2 * radius) * 2 * PI

        '----- Torsional stiffness -----
        R_stiff = R_force / R_displac
        TextBox87.Text = (R_stiff / 1000).ToString("F0")    '[kN/rad]

        'Estimate moment of inertia
        wt = NumericUpDown69.Value                      '[kg] complete NDE bearing support
        R_inertia = 1 / 3 * wt * radius ^ 2             '[kg.m2] 

        'Natural torsional frequency
        R_period = 2 * PI * Sqrt(R_inertia / R_stiff)
        R_freq = 1 / R_period                           '[Hz]
        R_speed = R_freq * 60                           '[Rpm]

        TextBox86.Text = R_inertia.ToString("F1")       '[kg.m2] Moment inertia
        TextBox85.Text = R_period.ToString("F4")        '[s] Period torsional
        TextBox84.Text = R_freq.ToString("F1")          '[Hz] frequency torsional
        TextBox83.Text = R_speed.ToString("F0")         '[rpm] speed
        TextBox82.Text = (R_speed * 2).ToString("F0")   '[rpm] speed first harmonic
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click, TabPage10.Enter, NumericUpDown60.ValueChanged, NumericUpDown59.ValueChanged, NumericUpDown58.ValueChanged, NumericUpDown63.ValueChanged, NumericUpDown62.ValueChanged, NumericUpDown61.ValueChanged
        Dim weight_r As Double
        Dim un_bal_speed_rms As Double  'rms valuemeasured in the field
        Dim un_bal_speed_pp As Double   'peak-peak value
        Dim un_bal_travel_r As Double   'Unbalance travel radius
        Dim un_bal_travel_pp As Double  'Unbalance travel peak-peak
        Dim ang_speed As Double
        Dim un_bal_force As Double
        Dim rpm As Double
        Dim F_dyn_found As Double

        Calc_Rotary_feeder()

        un_bal_speed_rms = NumericUpDown58.Value / 1000             '[mm/s]-->[m/s]
        un_bal_speed_pp = un_bal_speed_rms * 2 * Sqrt(2)            '[m/s] peak-peak
        rpm = NumericUpDown60.Value                                 '[rpm]
        weight_r = NumericUpDown59.Value                            '[kg]
        ang_speed = rpm * 2 * PI / 60                               '[rad/s]

        '------------- Unbalance travel --------------------
        un_bal_travel_r = (un_bal_speed_pp * 0.5) / ang_speed       '[m] unbalance travel radius
        un_bal_travel_pp = un_bal_travel_r * 2                      '[m] unbalance travel peak-peak

        '------------- Calc centrifugal force --------
        un_bal_force = weight_r * ang_speed * un_bal_speed_pp       '[N]

        '------------- Dynamic foundation force ------
        '@ 20 mm/s, rms
        F_dyn_found = 20 * 2 * Sqrt(2) * 10 ^ -3 * weight_r * ang_speed

        Label159.Visible = CBool(IIf(F_dyn_found / 10 < weight_r, vbFalse, vbTrue))
        Label160.Visible = CBool(IIf(un_bal_force / 10 < weight_r, vbFalse, vbTrue))
        TextBox74.Text = (un_bal_speed_pp * 1000).ToString("F2")  '[mm/s, p-p]
        TextBox91.Text = (un_bal_travel_pp * 1000).ToString("F2")  '[mm, p-p]
        TextBox75.Text = F_dyn_found.ToString("F0")               '[N]
        TextBox96.Text = ang_speed.ToString("F1")                 '[rad/s]
        TextBox95.Text = un_bal_force.ToString("F0")              '[N]
        TextBox113.Text = (rpm / 60).ToString("F1")              '[Hz]
    End Sub

    Private Sub Calc_Rotary_feeder()
        Dim rps, dia As Double      'Rotary feeder
        Dim lump, a_time As Double  'Lump data
        Dim acc As Double
        Dim tip_speed As Double
        Dim Force As Double

        '---- get data -------------
        lump = NumericUpDown61.Value        '[kg]
        rps = NumericUpDown62.Value / 60    '[rot_per_sec]
        dia = NumericUpDown63.Value / 1000  '[m]

        '---- calc -----------------
        tip_speed = PI * dia * rps          '[m/s]

        '---- The lump acceleration takes place in 1/4 rotor turn ---
        a_time = 1 / (4 * rps)              'Accelation time [s]

        acc = tip_speed / a_time            'Accelation [m/s]

        Force = lump * acc                  'Accelation force [N]

        TextBox77.Text = tip_speed.ToString("F1")
        TextBox78.Text = acc.ToString("F0")
        TextBox79.Text = Force.ToString("F0")
        TextBox80.Text = a_time.ToString("F2")
    End Sub

    Private Sub Print_torsion()
        Dim oWord As Word.Application ' = Nothing

        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2, oPara3 As Word.Paragraph
        Dim row, font_sI_polarze As Integer
        Dim chart_sI_polare As Integer = 55  '% of original picture sI_polare
        Dim ufilename, filename As String
        ufilename = "Frame_Vibration_Calculation_" & TextBox7.Text & "_" & TextBox8.Text & "_" & DateTime.Now.ToString("yyyy_MM_dd") & ".docx"

        Try
            oWord = New Word.Application()

            'Start Word and open the document template. 
            'oWord.font.size = 9
            oWord = CType(CreateObject("Word.Application"), Word.Application)
            oWord.Visible = True
            oDoc = oWord.Documents.Add

            oDoc.PageSetup.TopMargin = 65
            oDoc.PageSetup.BottomMargin = 15

            'Insert a paragraph at the beginning of the document. 
            oPara1 = oDoc.Content.Paragraphs.Add

            oPara1.Range.Text = "VTK Engineering"
            oPara1.Range.Font.Name = "Arial"
            oPara1.Range.Font.Size = Font.Size + 3
            oPara1.Range.Font.Bold = CInt(True)
            oPara1.Format.SpaceAfter = 1                '24 pt spacing after paragraph. 
            oPara1.Range.InsertParagraphAfter()

            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara2.Range.Font.Size = font_sI_polarze + 1
            oPara2.Format.SpaceAfter = 1
            oPara2.Range.Font.Bold = CInt(False)
            oPara2.Range.Text = "HorI_polarontal Vibration Analyses of the NDE bearing support of a 'between bearings' fan" & vbCrLf
                        oPara2.Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 2)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sI_polarze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

            row = 1
            oTable.Cell(row, 1).Range.Text = "Project Number"
            oTable.Cell(row, 2).Range.Text = TextBox7.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Intem name"
            oTable.Cell(row, 2).Range.Text = TextBox8.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Fan type"
            oTable.Cell(row, 2).Range.Text = TextBox9.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Author"
            oTable.Cell(row, 2).Range.Text = Environment.UserName
            row += 1
            oTable.Cell(row, 1).Range.Text = "Date"
            oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            row += 1
            oTable.Cell(row, 1).Range.Text = "File name"
            oTable.Cell(row, 2).Range.Text = ufilename

            oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(4)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            '----------------------General data-----------------------------------------

            'Insert a 6 (row) x 3 (column) table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 1)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sI_polarze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sI_polarze + 2
            row = 1

            oTable.Cell(row, 1).Range.Text = "General info"
            row += 1
            oTable.Cell(row, 1).Range.Text = "The NDE Bearing support is sensative to horI_polarontal vibration when the fan is not mounted on a conctrete support. "
            row += 1
            oTable.Cell(row, 1).Range.Text = "The concrete support gives weight and stiffnes to the spring mass system thus reducing the vibrations."
            row += 2
            oTable.Cell(row, 1).Range.Text = "In the absence of a concrete support the forces need to be contained by the steel frame. "
            row += 1
            oTable.Cell(row, 1).Range.Text = "The horI_polarontal stiffness is FEA calculated by adding a horI_polarontal force of 10 kN to the side "
            row += 1
            oTable.Cell(row, 1).Range.Text = "of the NDE bearing house and then determining the horI_polarontal deflection (C= Force/deflection [kN/mm])"
            row += 1
            oTable.Cell(row, 1).Range.Text = "The Center Of Gravity (COG of the NDE support) distance to the floor is also FEA calculated."
            row += 1

            oTable.Columns(1).Width = oWord.InchesToPoints(6.5)   'Change width of columns 1 .
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------- NDE bearing support data --------------
            'Insert a 6 (row) x 3 (column) table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sI_polarze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sI_polarze + 2
            row = 1

            oTable.Cell(row, 1).Range.Text = "NDE Bearing support data"
            row += 1

            oTable.Cell(row, 1).Range.Text = "Bearing to floor height"
            oTable.Cell(row, 2).Range.Text = NumericUpDown66.Value.ToString("F2")
            oTable.Cell(row, 3).Range.Text = "[m]"
            row += 1

            oTable.Cell(row, 1).Range.Text = "NDE support COG to floor height"
            oTable.Cell(row, 2).Range.Text = NumericUpDown65.Value.ToString("F2")
            oTable.Cell(row, 3).Range.Text = "[m]"
            row += 1

            oTable.Cell(row, 1).Range.Text = "Distance Ratio "
            oTable.Cell(row, 2).Range.Text = TextBox90.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1

            oTable.Columns(1).Width = oWord.InchesToPoints(3.0)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.2)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)

            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a 8 (row) x 3 table (column), fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sI_polarze
            oTable.Range.Font.Bold = CInt(False)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sI_polarze + 2
            row = 1

            oTable.Cell(row, 1).Range.Text = "Linear Natural frequency NDE bearing support"
            row += 1

            oTable.Cell(row, 1).Range.Text = "Weight bearing support"
            oTable.Cell(row, 2).Range.Text = NumericUpDown56.Value.ToString("F0")
            oTable.Cell(row, 3).Range.Text = "[kg]"
            row += 1

            oTable.Cell(row, 1).Range.Text = "HorI_polarontal Linear stiffness @ bearing"
            oTable.Cell(row, 2).Range.Text = NumericUpDown57.Value.ToString("F1")
            oTable.Cell(row, 3).Range.Text = "[kN/mm]"
            row += 1

            oTable.Cell(row, 1).Range.Text = "HorI_polarontal Linear stiffness @ COG"
            oTable.Cell(row, 2).Range.Text = TextBox67.Text
            oTable.Cell(row, 3).Range.Text = "[kN/mm]"
            row += 1

            oTable.Cell(row, 1).Range.Text = "Natural vibration frequency"
            oTable.Cell(row, 2).Range.Text = TextBox68.Text
            oTable.Cell(row, 3).Range.Text = "[Hz]"
            row += 1

            oTable.Cell(row, 1).Range.Text = "Natural vibration speed"
            oTable.Cell(row, 2).Range.Text = TextBox69.Text
            oTable.Cell(row, 3).Range.Text = "[rpm]"
            row += 1

            oTable.Cell(row, 1).Range.Text = "Maximum allowed speed"
            oTable.Cell(row, 2).Range.Text = TextBox70.Text
            oTable.Cell(row, 3).Range.Text = "[rpm]"
            row += 1

            oTable.Columns(1).Width = oWord.InchesToPoints(3.0)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.2)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '----------- Insert picturebox14 -------
            filename = dirpath_Rap & "Picturebox14.Jpeg"
            PictureBox14.Image.Save(filename, System.Drawing.Imaging.ImageFormat.Jpeg)
            oPara3 = oDoc.Content.Paragraphs.Add
            oPara3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oPara3.Range.InlineShapes.AddPicture(filename)
            oPara3.Range.InlineShapes.Item(1).LockAspectRatio = CType(True, Microsoft.Office.Core.MsoTriState)
            oPara3.Range.InlineShapes.Item(1).ScaleWidth = chart_sI_polare       'SI_polare
            oPara3.Range.InsertParagraphAfter()

            '--------------Save file word file------------------
            'See https://msdn.microsoft.com/en-us/library/63w57f4b.aspx

            If (Not System.IO.Directory.Exists(dirpath_Eng)) Then System.IO.Directory.CreateDirectory(dirpath_Eng)
            If (Not System.IO.Directory.Exists(dirpath_Rap)) Then System.IO.Directory.CreateDirectory(dirpath_Rap)
            If (Not System.IO.Directory.Exists(dirpath_Home)) Then System.IO.Directory.CreateDirectory(dirpath_Home)

            If Directory.Exists(dirpath_Rap) Then
                oWord.ActiveDocument.SaveAs(dirpath_Rap & ufilename)
            Else
                oWord.ActiveDocument.SaveAs(dirpath_Home & ufilename)
            End If

        Catch ex As Exception
            MessageBox.Show(ufilename & vbCrLf & ex.Message)  ' Show the exception's message.
        End Try
    End Sub

    Private Sub Button11_Click_1(sender As Object, e As EventArgs) Handles Button11.Click
        Print_torsion()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Example_6()
    End Sub

    Private Sub Example_6()
        TextBox7.Text = "P06.1050"
        TextBox8.Text = "Chorus"
        TextBox9.Text = "2250/4130/T36A"

        RadioButton1.Checked = False        'Overhung
        RadioButton2.Checked = True         'Between bearing
        RadioButton3.Checked = False        'Steel support
        RadioButton4.Checked = True         'Concrete

        NumericUpDown1.Value = 3200         '[mm] bearing-impeller
        NumericUpDown2.Value = 3200         '[mm] impeller-bearing
        NumericUpDown3.Value = 0            '[mm] rigid length shaft
        NumericUpDown4.Value = 7600         '[kg] impeller weight
        NumericUpDown55.Value = 300         '[c] operating temp
        NumericUpDown8.Value = 690          '[mm] shaft OD (28" schedule 40)
        NumericUpDown68.Value = 580         '[mm] shaft ID

        NumericUpDown20.Value = 4130        '[mm] impeller dia
        NumericUpDown21.Value = CDec(72.8)  '[mm] impeller width

        NumericUpDown10.Value = 16219       '[kg.m2] Jp
        NumericUpDown11.Value = 8113        '[kg.m2] Ja

        TabControl1.SelectedIndex = 0
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Example_5()
    End Sub

    Private Sub Example_5()
        TextBox7.Text = "P01.1158"
        TextBox8.Text = "Std Fasel"
        TextBox9.Text = "1120/1900/T36"

        RadioButton1.Checked = False        'Overhung
        RadioButton2.Checked = True         'Between bearing
        RadioButton3.Checked = False        'Steel support
        RadioButton4.Checked = True         'Concrete

        NumericUpDown1.Value = 600          '[mm] bearing-impeller
        NumericUpDown2.Value = 1150         '[mm] impeller-bearing
        NumericUpDown3.Value = 0            '[mm] rigid length shaft
        NumericUpDown4.Value = 580          '[kg] impeller weight
        NumericUpDown55.Value = 35          '[c] operating temp
        NumericUpDown8.Value = 140          '[mm] shaft OD 
        NumericUpDown68.Value = 0           '[mm] shaft ID

        NumericUpDown20.Value = 1900        '[mm] impeller dia
        NumericUpDown21.Value = CDec(26.3)  '[mm] impeller width

        NumericUpDown10.Value = CDec(262.5) '[kg.m2] Jp
        NumericUpDown11.Value = CDec(131.3) '[kg.m2] Ja

        TabControl1.SelectedIndex = 0

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Example_4()
    End Sub
    Private Sub Example_4()
        TextBox7.Text = "P16.0051"
        TextBox8.Text = "Biowanze SA"
        TextBox9.Text = "1980/2280/T31A"

        RadioButton1.Checked = True         'Overhung
        RadioButton2.Checked = False        'Overhung
        RadioButton3.Checked = True         'Steel support
        RadioButton4.Checked = False        'Concrete support

        NumericUpDown1.Value = 750          '[mm] distance bearing-bearing
        NumericUpDown8.Value = 180          '[mm] shaft OD bearing-bearing 

        NumericUpDown2.Value = 479          '[mm] overhung
        NumericUpDown3.Value = 0            '[mm] rigid length shaft
        NumericUpDown9.Value = 145          '[mm] overhung shaft OD 
        NumericUpDown68.Value = 0           '[mm] overhung shaft ID
        NumericUpDown4.Value = 862          '[kg] impeller weight
        NumericUpDown55.Value = 80          '[c] operating temp

        NumericUpDown20.Value = 2280        '[mm] impeller dia
        NumericUpDown21.Value = CDec(27.1)  '[mm] impeller width

        NumericUpDown10.Value = CDec(508.5) '[kg.m2] Jp
        NumericUpDown11.Value = CDec(280.4) '[kg.m2] Ja

        TabControl1.SelectedIndex = 0
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Example_3()
    End Sub
    Private Sub Example_3()
        TextBox7.Text = "P19.1065"
        TextBox8.Text = "Krefeld"
        TextBox9.Text = "1250/2155/T36"

        RadioButton1.Checked = True         'Overhung
        RadioButton2.Checked = False        'Overhung
        RadioButton3.Checked = True         'Steel support
        RadioButton4.Checked = False        'Concrete support

        NumericUpDown1.Value = 750          '[mm] distance bearing-bearing
        NumericUpDown8.Value = 180          '[mm] shaft OD bearing-bearing 

        NumericUpDown2.Value = 410          '[mm] overhung
        NumericUpDown3.Value = 0            '[mm] rigid length shaft
        NumericUpDown9.Value = 145          '[mm] overhung shaft OD 
        NumericUpDown68.Value = 0           '[mm] overhung shaft ID
        NumericUpDown4.Value = 638          '[kg] impeller + hub weight
        NumericUpDown55.Value = 70          '[c] operating temp

        NumericUpDown20.Value = 2155        '[mm] impeller dia
        NumericUpDown21.Value = CDec(23.4)  '[mm] impeller width

        NumericUpDown10.Value = CDec(386.5) '[kg.m2] Jp
        NumericUpDown11.Value = CDec(193.3) '[kg.m2] Ja

        TabControl1.SelectedIndex = 0
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Example_2()
    End Sub
    Private Sub Example_2()
        TextBox7.Text = "P19.1065"
        TextBox8.Text = "Krefeld"
        TextBox9.Text = "1250/2155/T36"

        RadioButton1.Checked = True         'Overhung
        RadioButton2.Checked = False        'Overhung
        RadioButton3.Checked = True         'Steel support
        RadioButton4.Checked = False        'Concrete support

        NumericUpDown1.Value = 900          '[mm] distance bearing-bearing
        NumericUpDown8.Value = 180          '[mm] shaft OD bearing-bearing 

        NumericUpDown2.Value = 410          '[mm] overhung
        NumericUpDown3.Value = 0            '[mm] rigid length shaft
        NumericUpDown9.Value = 145          '[mm] overhung shaft OD 
        NumericUpDown68.Value = 0           '[mm] overhung shaft ID
        NumericUpDown4.Value = 638          '[kg] impeller + hub weight
        NumericUpDown55.Value = 70          '[c] operating temp

        NumericUpDown20.Value = 2155        '[mm] impeller dia
        NumericUpDown21.Value = CDec(23.4)  '[mm] impeller width

        NumericUpDown10.Value = CDec(386.5) '[kg.m2] Jp
        NumericUpDown11.Value = CDec(193.3) '[kg.m2] Ja

        TabControl1.SelectedIndex = 0
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Example_1()
    End Sub
    Private Sub Example_1()
        TextBox7.Text = "P17.1053"
        TextBox8.Text = "Supezet"
        TextBox9.Text = "1500/1535/T36"

        RadioButton1.Checked = False        'Overhung
        RadioButton2.Checked = True         'Between bearing
        RadioButton3.Checked = True        ' Steel support
        RadioButton4.Checked = False        'Concrete

        NumericUpDown1.Value = CDec(613.5)  '[mm] bearing-impeller
        NumericUpDown2.Value = CDec(2026.5) '[mm] impeller-bearing 
        NumericUpDown3.Value = 0            '[mm] rigid length shaft
        NumericUpDown4.Value = 455          '[kg] modified impeller weight
        NumericUpDown55.Value = 225         '[c] operating temp
        NumericUpDown8.Value = 145          '[mm] shaft OD 
        NumericUpDown68.Value = 0           '[mm] shaft ID

        NumericUpDown20.Value = 1535        '[mm] modified impeller dia
        NumericUpDown21.Value = CDec(31.5)  '[mm] impeller width

        NumericUpDown10.Value = CDec(133.9) '[kg.m2] Jp
        NumericUpDown11.Value = CDec(67.4)  '[kg.m2] Ja

        TabControl1.SelectedIndex = 0
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        TextBox7.Text = "P03.1033"
        TextBox8.Text = "Lummus-China"
        TextBox9.Text = "2MD 2100/2370/T33"

        RadioButton1.Checked = False        'Overhung
        RadioButton2.Checked = True         'Between bearing
        RadioButton3.Checked = True         'Steel support
        RadioButton4.Checked = False        'Concrete

        NumericUpDown1.Value = CDec(2997)   '[mm] bearing-impeller
        NumericUpDown2.Value = CDec(2997)   '[mm] impeller-bearing 
        NumericUpDown3.Value = 0            '[mm] rigid length shaft
        NumericUpDown4.Value = 1750         '[kg] impeller weight

        NumericUpDown6.Value = 80           '[kN/mm2] stiffness
        NumericUpDown7.Value = 12           '[kN/mm2] stiffness

        NumericUpDown55.Value = 400         '[c] design temp
        NumericUpDown8.Value = 400          '[mm] shaft OD 
        NumericUpDown68.Value = 0           '[mm] shaft ID

        NumericUpDown20.Value = 2370        '[mm] impeller dia
        NumericUpDown21.Value = CDec(50.8)  '[mm] impeller width

        NumericUpDown10.Value = CDec(1227.3) '[kg.m2] Jp
        NumericUpDown11.Value = CDec(614)  '[kg.m2] Ja

        TabControl1.SelectedIndex = 0
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        TextBox7.Text = "P04.1257"
        TextBox8.Text = "Technip"
        TextBox9.Text = "2LD 1600/1770/T33"

        RadioButton1.Checked = False        'Overhung
        RadioButton2.Checked = True         'Between bearing
        RadioButton3.Checked = True         'Steel support
        RadioButton4.Checked = False        'Concrete

        NumericUpDown1.Value = CDec(2150)   '[mm] bearing-impeller
        NumericUpDown2.Value = CDec(2150)   '[mm] impeller-bearing 
        NumericUpDown3.Value = 0            '[mm] rigid length shaft
        NumericUpDown4.Value = 875         '[kg] impeller weight

        NumericUpDown6.Value = 80           '[kN/mm2] stiffness
        NumericUpDown7.Value = 12           '[kN/mm2] stiffness

        NumericUpDown55.Value = 350         '[c] design temp
        NumericUpDown8.Value = 300          '[mm] shaft OD 
        NumericUpDown68.Value = 0           '[mm] shaft ID

        NumericUpDown20.Value = 1770        '[mm] impeller dia
        NumericUpDown21.Value = CDec(45.6)  '[mm] impeller width

        NumericUpDown10.Value = CDec(342.7) '[kg.m2] Jp
        NumericUpDown11.Value = CDec(171.5) '[kg.m2] Ja

        TabControl1.SelectedIndex = 0
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Directory.Exists(dirpath_Eng) Then
            Label207.Visible = False
        Else
            Label207.Visible = True
        End If
    End Sub
End Class
