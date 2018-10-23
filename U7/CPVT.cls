VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPVT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=======================================================================================
'Unifloc7.3  Testudines                                           khabibullinra@gmail.com
'���������� ��������� ������� �� ��������� �����������
'2000 - 2018 �
'
'=======================================================================================


Option Explicit
 ' ������� ������� ��� ������� ���������� ������
 Private p_PTcalc As PTtype                         ' ��������������� ������� ��� ������� ��� �������� ������
 ' ������� ��������� ������
 Private p_gamma_o As Double                        ' ��������� ����� ��������
 Private p_gamma_g As Double                        ' ��������� ���� ��������
 Private p_gamma_w As Double                        ' ��������� ���� ��������
 Private p_Rsb_m3m3 As Double                       ' �������������� ��� �������� ���������
 ' ������������� ��������� �����
 Private p_Pb_atma As Double                         ' �������� ���������  (������������� ��������)
 Private p_Muob_cP As Double                       ' �������� ����� ��� �������� ��������� (������������� ��������)
 Private p_Bob_m3m3 As Double                       ' �������� ����������� ��� �������� ���������
 Private p_Tres_C As Double                         ' ��������� ����������� ��� ������� ������� �������� �������� ��������� � ��������� ������������
 ' ��������� ��������� �����
 Private p_Pbcalc_atma As Double                     ' ��������� �������� �������� ��������� �� ����������
 Private p_Rs_m3m3 As Double                        ' ��������� �������� �������������� � ����� ��� ������� ��������
 Private p_Bo_m3m3 As Double                        ' �������� ����������� ����� ��� ������� ��������
 Private p_Muo_cP As Double                         ' �������� ����� ��� ������� ��������
 Private p_Muo_deadoil_cP As Double                 ' �������� �������������� �����
 Private p_copmressibility_o_1atm As Double         ' ����������� �����
' Private p_sigma_o_Nm As Double                     ' ������������� ���������
' Private p_sigma_w_Nm As Double                     ' ������������� ���������
 Private p_ST_oilgas_dyncm As Double                ' ������������� ��������� ����� ���
 Private p_ST_watgas_dyncm As Double                ' ������������� ��������� ���� ���
 Private p_ST_liqgas_dyncm As Double
 ' ��������� ��������� ����
 Private p_Z As Double                              ' ��������� �������� ������������ ����������������
 Private p_Bg_m3m3 As Double                        ' �������� ����������� ����
 Private p_Mug_cP As Double                         ' �������� ���� ��� ������� ��������
 ' ��������� ��������� ����
 Private p_Bw_m3m3 As Double                        ' ��������� �������� ��������� ������������ ����
 Private p_BwSC_m3m3 As Double
 Private p_Muw_cP As Double                         ' �������� ����
 Private p_Salinity_ppm As Double                   ' ��������� ����
 ' ��������� ������
 Private p_Rp_m3m3 As Double                        ' ������� ������ �������� (����������� � ����������� ��������)
 Private p_wc_fr As Double                          ' �������� ���� ���� � ������
 Private p_Qliq_scm3day As Double                   ' ������ ��� ������ ����� � ������, ��� �������� ���������� ������� �������� � ������ ��������
 Private p_Qo_m3day As Double
 Private p_Qw_m3day As Double
 Private p_Qgas_m3day As Double
 Private p_Qliq_m3day As Double
 Private p_GasFraction_d As Double
 Private p_MuMix_cP As Double
 Private p_rho_oil_kgm3 As Double
 Private p_rho_water_kgm3 As Double
 Private p_rho_liq_kgm3 As Double
 Private p_rho_mix_kgm3 As Double
 Private p_Qgfree_scm3day As Double                 ' ����� ����, ������ ������� ��� �������� Qgas, � ����������������� �����
 ' ����� ���������� ��� ������������� ��������
 Private p_Co_JkgC As Double                        ' oil heat capacity   ������������ �����  ��/��/�
 Private p_Cw_JkgC As Double                        ' water heat capacity  ������������ ����
 Private p_Cg_JkgC As Double                        ' ������������ ���� gas heat capacity
 ' ��������� ������
 Private p_PVT_correlation As PVT_CORRELATION       ' PVT correlation
 Public calculated As Boolean
 Private p_ZNLF As Boolean                          ' ���� ���������� ��� ����� � ������ ���������
 Private p_Twh_C As Double                          ' ����������� ��� ������� �������� ������
                                                    '  � ������ ��������� ��� ���������� ����� ����������� ����� �������� (����� )
                                                    '  �������� ������������� ������ ���������� ���� ����� ����� ��������
                                                    '  ��� ���� ��� �� ���������� �� ��������
                                                    '  �������������� ������������� - � ����������� ������ ����� �������� ������ �����
 Private c_RsTres_Curve As New TInterpolation       ' ������ �������������� �� �������� ��� ��������� �����������
 Private c_RsT_Curve As New TInterpolation          ' ������ �������������� �� �������� ��� ������������ �����������
 Private c_BoTres_Curve As New TInterpolation       ' ������ ��������� ����������� ����� �� �������� ��� ��������� �����������
 Private c_BoT_Curve As New TInterpolation          ' ������ ��������� ����������� ������� �������� ��� ��������� �����������
 Private c_MuoTres_Curve As New TInterpolation      ' ������ �������� ����� �� �������� ��� ��������� �����������
 Private c_MuoT_Curve As New TInterpolation         ' ������ �������� ����� �� �������� ��� ��������� �����������
 Private c_MugTres_Curve As New TInterpolation      ' ������ �������� ���� �� �������� ��� ��������� �����������
 Private c_MugT_Curve As New TInterpolation         ' ������ �������� ���� �� �������� ��� ��������� �����������

 Public Property Get ZNLF() As Boolean
    ZNLF = p_ZNLF
 End Property
  
 Public Property Let ZNLF(ByVal val As Boolean)
    p_ZNLF = val
    If p_ZNLF Then
        wc_fr = 0      ' �������� ���������� ����
        Qliq_scm3day = 0 ' �������� ����� �������� - ������ �� ��������. ��� ������� ����� ���� �� ��������
    Else
        ' ������ �� ������
    End If
 End Property
  
 Public Property Get wc_fr() As Double
   wc_fr = p_wc_fr
 End Property
  
 Public Property Let wc_fr(val As Double)
   If val < 0 Then
        addLogMsg "CPVT.wc_fr: ������������� ������ ����. �������� = " & s(val) & " �������� �� 0"
        val = 0
   ElseIf val > 1 Then
        val = 1
        addLogMsg "CPVT.wc_fr: ������������� ������ �������. �������� = " & s(val) & " �������� �� 1"
   End If
   p_wc_fr = val
 End Property

 Public Property Get wc_perc() As Double
   wc_perc = p_wc_fr * 100
 End Property
  
 Public Property Let wc_perc(val As Double)
   wc_fr = val / 100    ' �������� ��������, ����� �������� ��� ��������� ������
 End Property

 Public Property Get Qo_scm3day() As Double
   Qo_scm3day = p_Qliq_scm3day * (1 - p_wc_fr)
   ' ��� ��������� ����� 0
 End Property
  
 Public Property Get Qliq_scm3day() As Double
    Qliq_scm3day = p_Qliq_scm3day
   ' ��� ��������� ����� 0
 End Property
 
 Public Property Let Qliq_scm3day(val As Double)
    p_Qliq_scm3day = val
 End Property
  
 Private Sub calc_wc()
    If (Qliq_scm3day) > 0 Then
        wc_fr = Qw_scm3day / (Qliq_scm3day)
    Else
        wc_fr = 0
    End If
 End Sub
 
 Public Property Get Qw_scm3day() As Double
   Qw_scm3day = p_Qliq_scm3day * p_wc_fr
   ' ��� ��������� ����� 0
 End Property
  
 Public Property Get Qgas_scm3day() As Double
   Qgas_scm3day = Qo_scm3day * Rp_m3m3 + Qgfree_scm3day    ' ��������� ������� ���������� ���� � ������
   ' ��� ��������� ����� ��� ���������
 End Property

 Public Property Get QgasInSitu_scm3day() As Double
 ' ������ ���� � �������� ��������������� �������� ����������� � ����������� ��������
   QgasInSitu_scm3day = (Qgas_scm3day - rs_m3m3 * Qo_scm3day)
   ' ��� ��������� ����� ��� ��� ���� - ��������� ��� ��� ����� � ������ ���� � ����������� ��������
 End Property

 Public Property Get Qgas_m3day() As Double
   Qgas_m3day = QgasInSitu_scm3day * Bg_m3m3
   If Qgas_m3day < 0 Then Qgas_m3day = 0
   ' ��� ��������� ����� ��� ��� ����
 End Property
 
 Public Property Get Qgas_rc_m3day() As Double
   Qgas_rc_m3day = QgasInSitu_scm3day * Bg_m3m3
   If Qgas_m3day < 0 Then Qgas_rc_m3day = 0
   ' ��� ��������� ����� ��� ��� ����
 End Property

 Public Property Get Qo_m3day() As Double
   Qo_m3day = Qo_scm3day * Bo_m3m3
   ' ��� ��������� ����� ��� ��� ����
 End Property
 
 Public Property Get Qo_rc_m3day() As Double
   Qo_rc_m3day = Qo_scm3day * Bo_m3m3
   ' ��� ��������� ����� ��� ��� ����
 End Property

 Public Property Get Qw_m3day() As Double
   Qw_m3day = Qw_scm3day * Bw_m3m3
   ' ��� ��������� ����� ��� ��� ����
 End Property
 
 Public Property Get Qw_rc_m3day() As Double
   Qw_rc_m3day = Qw_scm3day * Bw_m3m3
   ' ��� ��������� ����� ��� ��� ����
 End Property
  
 Public Property Get Qliq_m3day() As Double
   Qliq_m3day = Qw_m3day + Qo_m3day
   ' ��� ��������� ����� ��� ��� ����
 End Property
 
 Public Property Get Qliq_rc_m3day() As Double
   Qliq_rc_m3day = Qw_rc_m3day + Qo_rc_m3day
   ' ��� ��������� ����� ��� ��� ����
 End Property
   
 Public Property Get Qmix_m3day() As Double
   Qmix_m3day = p_Qw_m3day + p_Qo_m3day + p_Qgas_m3day
   ' ��� ��������� ����� ��� ��� ����
 End Property

 Public Property Get Qgfree_scm3day() As Double
   Qgfree_scm3day = p_Qgfree_scm3day
 End Property

 Public Property Let Qgfree_scm3day(ByVal vNewValue As Double)
   p_Qgfree_scm3day = vNewValue
 End Property
  
 Public Property Get Wm_kgsec() As Double
   Wm_kgsec = wg_kgsec + wo_kgsec + wwat_kgsec
   ' ��������� ��� ������ ��������� ������� �� ����������� ����� ������� ��� �������� ��������������� ������� ����� ���� ������ ��� ������� � ������
 End Property
  
 Public Property Get Compressibility_oil_1atm() As Double
    Compressibility_oil_1atm = p_copmressibility_o_1atm
 End Property
 
  
 Public Property Get Co_JkgC() As Double  ' oil heat capacity   ������������ �����  ��/��/�
    Co_JkgC = p_Co_JkgC
 End Property
 
 Public Property Get Cw_JkgC() As Double   ' water heat capacity  ������������ ����
    Cw_JkgC = p_Cw_JkgC
 End Property
 
 Public Property Get Cg_JkgC() As Double  ' ������������ ���� gas heat capacity
    Cg_JkgC = p_Cg_JkgC
 End Property

 ' �������� ������ ���� ��� �������� ��������������� ��������
 Public Property Get wg_kgsec() As Double
    wg_kgsec = rho_gas_kgm3 * Qgas_m3day * const_conver_sec_day
 End Property

 ' �������� ������ ����� ��� �������� ��������������� ��������
 Public Property Get wo_kgsec() As Double
    wo_kgsec = rho_oil_kgm3 * Qo_m3day * const_conver_sec_day
 End Property
 ' �������� ������ ���� ��� �������� ��������������� ��������
 Public Property Get wwat_kgsec() As Double
    wwat_kgsec = rho_water_kgm3 * Qw_m3day * const_conver_sec_day
 End Property
 ' �������� ������ �������� ��� �������� ��������������� ��������
 Public Property Get wliq_kgsec() As Double
    wliq_kgsec = wo_kgsec + wwat_kgsec * const_conver_sec_day
 End Property
 
 Public Property Get Cliq_JkgC() As Double  ' mixture heat capacity   ������������ ��������  ��/��/�
    If Qmix_m3day > 0 Then
         Cliq_JkgC = (Co_JkgC * wo_kgsec + Cw_JkgC * wwat_kgsec) / (wwat_kgsec + wo_kgsec)
    Else
        Cliq_JkgC = Co_JkgC
    End If
 End Property

 Public Property Get Cmix_JkgC() As Double  ' mixture heat capacity   ������������ ��������  ��/��/�
    If Qmix_m3day > 0 Then
         Cmix_JkgC = (Cliq_JkgC * wliq_kgsec + Cg_JkgC * wg_kgsec) / (wliq_kgsec + wg_kgsec)
    Else
        Cmix_JkgC = Co_JkgC
    End If
 End Property
 
 Public Property Get CJT_Katm() As Double
 ' ���������� ������ ������� ��� ����������� �����
    Dim X As Double
    Dim wm As Double
    Dim z1 As Double, z2 As Double
    Dim dTz As Double
    Dim dZdT As Double
    Dim TZdZdT As Double
    wm = Wm_kgsec
    ' ����� ���������� ����������� �� ����������������, ����� �� � �������, ����� �������� ������ �������
    dTz = 5
      ' ������ ����������� ����������� ����� ����� �������� ������ - ���� ��������, ���� ������, ��� ������ ������
    z1 = unf_Calc_Zgas_d(Tcalc_K, Pcalc_MPaa, p_gamma_g)
    z2 = unf_Calc_Zgas_d(Tcalc_K + dTz, Pcalc_MPaa, p_gamma_g)
    p_Z = z1
    dZdT = (z2 - z1) / dTz
    TZdZdT = Tcalc_K / p_Z * dZdT
    If wm > 0 Then
        X = wg_kgsec / Wm_kgsec    ' �������� ���� ���� � ������
    Else
        X = 0
    End If
    CJT_Katm = 1 / Cmix_JkgC * (X / rho_gas_kgm3 * (-TZdZdT) + (1 - X) / rho_liq_kgm3) * const_convert_atma_Pa
 End Property
 
 Public Property Get oil_API() As Double
    oil_API = 141.5 / p_gamma_o - 131.5
 End Property
 
 Public Property Get rho_oil_kgm3() As Double
    Dim msg As String
    If Bo_m3m3 > 0 Then
        rho_oil_kgm3 = 1000 * (p_gamma_o + p_Rs_m3m3 * p_gamma_g * const_rho_air / 1000) / p_Bo_m3m3
    Else
        msg = "CPVT.rho_oil_kgm3: ������ ��������� � ��������������� ��������� Bo_m3m3" & Bo_m3m3 & "�������� Bo ���������������"
        addLogMsg msg
        rho_oil_kgm3 = 1000 * (p_gamma_o + p_Rs_m3m3 * p_gamma_g * const_rho_air / 1000)
        Err.Raise kErrPVTinput, , msg
    End If
 End Property
 
 Public Property Get rho_water_kgm3() As Double
    Dim msg As String
    If Bw_m3m3 > 0 Then
        rho_water_kgm3 = 1000 * (p_gamma_w) / p_Bw_m3m3
    Else
        msg = "CPVT.rho_water_kgm3: ������ ��������� � ��������������� ��������� Bw_m3m3" & Bw_m3m3 & "�������� Bw ���������������"
        addLogMsg msg
        rho_water_kgm3 = 1000 * (gamma_w)
        Err.Raise kErrPVTinput, , msg
    End If
 End Property

 Public Property Get rho_liq_kgm3() As Double
    rho_liq_kgm3 = p_rho_liq_kgm3 '(1 - wc_fr) * rho_oil_kgm3 + wc_fr * rho_water_kgm3
 End Property
 
 Public Property Get rho_gas_kgm3() As Double
    Dim msg As String
    If Bg_m3m3 > 0 Then
        rho_gas_kgm3 = gamma_g * const_rho_air / Bg_m3m3
    Else
        msg = "CPVT.rho_gas_kgm3: ������ ��������� � ��������������� ��������� Bg_m3m3" & Bg_m3m3 & "�������� Bg ���������������"
        addLogMsg msg
        rho_gas_kgm3 = gamma_g * const_rho_air
        Err.Raise kErrPVTinput, , msg
    End If
 End Property

 Public Property Get f_g() As Double
    If Qmix_m3day > 0 Then
       f_g = Qgas_m3day / Qmix_m3day
    Else
       f_g = 0
    End If
   ' ��� ��������� ����� ��� ��� ����
 End Property

 Public Property Get rho_mix_kgm3() As Double
    rho_mix_kgm3 = p_rho_mix_kgm3  ' rho_liq_kgm3 * (1 - f_g) + rho_gas_kgm3 * f_g
 End Property
 
 Public Property Get sigma_liq_Nm() As Double
    'sigma_liq_Nm = p_sigma_o_Nm * (1 - wc_fr) + p_sigma_w_Nm * wc_fr
    sigma_liq_Nm = p_ST_liqgas_dyncm * 0.001
 End Property

 Public Property Get sigma_o_Nm() As Double
    'sigma_o_Nm = p_sigma_o_Nm
    sigma_o_Nm = p_ST_oilgas_dyncm * 0.001
 End Property
 
 Public Property Let sigma_o_Nm(val As Double)
    'p_sigma_o_Nm = val
    p_ST_oilgas_dyncm = val / 0.001
 End Property
 
 Public Property Get sigma_w_Nm() As Double
    'sigma_w_Nm = p_sigma_w_Nm
    sigma_w_Nm = p_ST_watgas_dyncm * 0.001
 End Property
 
 Public Property Let sigma_w_Nm(val As Double)
    'p_sigma_w_Nm = val
    p_ST_watgas_dyncm = val / 0.001
 End Property
 
  
 Public Property Get Tres_C() As Double
    Tres_C = p_Tres_C
 End Property
 
 Public Property Get Tres_K() As Double
    Tres_K = p_Tres_C + const_T_K_min
 End Property
 
 Public Property Let Tres_C(val As Double)
    p_Tres_C = val
 End Property
  
 ' ������ ����� ����   (����������� �������� � �������)
 Public Property Get m_g_kgmol() As Double
    m_g_kgmol = const_m_a_kgmol * p_gamma_g
 End Property

 Public Property Get Sal_ppm() As Double
    Sal_ppm = p_Salinity_ppm
 End Property
 
 Public Property Let Sal_ppm(val As Double)
    p_Salinity_ppm = val
 End Property
 
 Public Property Get PVT_CORRELATION() As Integer
    PVT_CORRELATION = p_PVT_correlation
 End Property
 
 Public Property Let PVT_CORRELATION(val As Integer)
   p_PVT_correlation = val
 End Property
 ' ----- gamma_o ----------------------------------------------------------------------------------------
 Public Property Get gamma_o() As Double
    gamma_o = p_gamma_o
 End Property
 
 Public Property Get RhoOil_sckgm3() As Double
    RhoOil_sckgm3 = p_gamma_o * const_rho_ref
 End Property
 
 Public Property Get rhoGas_sckgm3() As Double
    rhoGas_sckgm3 = p_gamma_g * const_rho_air
 End Property

 Public Property Get rhoWat_sckgm3() As Double
    rhoWat_sckgm3 = p_gamma_w * const_rho_ref
 End Property
 
 Public Property Let gamma_o(val As Double)
    If (val > const_gamma_oil_min) And (val < const_gamma_oil_max) Then
        p_gamma_o = val
        calculated = False
    Else
        If (val < const_gamma_oil_min) Then p_gamma_o = const_gamma_oil_min
        If (val > const_gamma_oil_max) Then p_gamma_o = const_gamma_oil_max
        addLogMsg "CPipe.gamma_o: ������� ������������� ����� gamma_o = " & val & " ����������� �������� = " & p_gamma_o, msgDataQualityReport
    End If
 End Property
 
 ' ----- gamma_g ----------------------------------------------------------------------------------------
 Public Property Get gamma_g() As Double
    gamma_g = p_gamma_g
 End Property
 
 Public Property Let gamma_g(val As Double)
 '��������� �������� ��������� ����, ������ �� ��� ��������� ������
    If (val > const_gamma_gas_min) And (val < const_gamma_gas_max) Then
        p_gamma_g = val
        calculated = False
    Else
        If (val < const_gamma_gas_min) Then p_gamma_g = const_gamma_gas_min
        If (val > const_gamma_gas_max) Then p_gamma_g = const_gamma_gas_max
        addLogMsg "CPipe.gamma_o: ������� ������������� ����� gamma_o = " & val & " ����������� �������� = " & p_gamma_g, msgDataQualityReport
    End If
 End Property
 
 ' ----- gamma_w ----------------------------------------------------------------------------------------
 Public Property Get gamma_w() As Double
    gamma_w = p_gamma_w
 End Property
 
 Public Property Let gamma_w(val As Double)
    If (val > const_gamma_water_min) And (val < const_gamma_water_max) Then
        p_gamma_w = val
        calculated = False
    Else
        If (val < const_gamma_water_min) Then p_gamma_w = const_gamma_water_min
        If (val > const_gamma_water_max) Then p_gamma_w = const_gamma_water_max
        addLogMsg "CPipe.gamma_o: ������� ������������� ����� gamma_o = " & val & " ����������� �������� = " & p_gamma_w, msgDataQualityReport
    End If
 End Property
 
' ----- Rp - GOR  ----------------------------------------------------------------------------------------
Public Property Get Rp_m3m3() As Double
    Rp_m3m3 = p_Rp_m3m3
End Property

Property Let Rp_m3m3(Rpval As Double)
    If (Rpval >= 0) Then
        p_Rp_m3m3 = Rpval
        If p_Rp_m3m3 < p_Rsb_m3m3 Then   ' ��������, ��� ������� ������ ������ ���� ������ ��� ��������������
            addLogMsg "������� ������ ��� ����� ������ �������������� Rp = " & Format(p_Rp_m3m3, "####0.00") & " < Rsb = " & Format(p_Rsb_m3m3, "#0.00") & ". �������������� ����������", msgDataQualityReport
            p_Rsb_m3m3 = p_Rp_m3m3
        End If
        calculated = False
    Else
        addLogMsg "������� ���������� ������������� �������� �������� ������� Rp =" & Format(Rpval, "####0.00"), msgError, kErrPVTinput
    End If
End Property
  
' ----- Rsb -----------------------------------------------------------------------------------------
Public Property Get rsb_m3m3() As Double
    rsb_m3m3 = p_Rsb_m3m3
End Property

Property Let rsb_m3m3(Rsbval As Double)
    If (Rsbval >= 0) Then
        p_Rsb_m3m3 = Rsbval
        If p_Rp_m3m3 < p_Rsb_m3m3 Then   ' ��������, ��� ������� ������ ������ ���� ������ ��� ��������������
            addLogMsg "�������������� ��� ����� ������ �������� �������  Rp = " & Format(p_Rp_m3m3, "#0.00") & " < Rsb = " & Format(p_Rsb_m3m3, "#0.00") & ". �������������� ����������", msgDataQualityReport
            p_Rsb_m3m3 = p_Rp_m3m3
        End If
        calculated = False
    Else
        addLogMsg "������� ���������� ������������� �������� �������������� Rsb =" & Format(Rsbval, "####0.00"), msgError, kErrPVTinput
    End If
End Property
 
Public Function SetRpRsb(Rpval_m3m3 As Double, Rsbval_m3m3 As Double) As Boolean
' ���������� � ����� ������ ���������� ������� ������ ��������� ������������ �������� �������� ������� � �������
    If Rpval_m3m3 > 0 Then
        If Rpval_m3m3 >= Rsbval_m3m3 Then
            p_Rp_m3m3 = Rpval_m3m3
            If Rsbval_m3m3 > 0 Then
                p_Rsb_m3m3 = Rsbval_m3m3
            Else
                p_Rsb_m3m3 = Rpval_m3m3
            End If
            SetRpRsb = True
        Else
            p_Rp_m3m3 = Rpval_m3m3
            p_Rsb_m3m3 = Rpval_m3m3
            SetRpRsb = True
            addLogMsg "�������������� ��� ����� ������ �������� �������  Rp = " & Format(p_Rp_m3m3, "#0.00") & " < Rsb = " & Format(p_Rsb_m3m3, "#0.00") & ". �������������� ����������", msgDataQualityReport
        End If
    Else
        If Rpval_m3m3 <= 0 And Rsbval_m3m3 > 0 Then
            p_Rp_m3m3 = Rsbval_m3m3
            p_Rsb_m3m3 = Rsbval_m3m3
            SetRpRsb = True
        Else
            SetRpRsb = False
        End If
    End If
    ' �������� ��� �������� ��� ��������� �������
    Rp_m3m3 = Rp_m3m3
    rsb_m3m3 = rsb_m3m3
End Function

'----- Pb -----------------------------------------------------------------------------------------
Public Property Get Pb_atma() As Double
    If p_Pb_atma > 0 Then       ' ���� �� ��������, ��� ������ ��� �������� �����������
    ' ���� �������� ������������� �������� ��� ��������� �����������, �� ���������� ���
        Pb_atma = p_Pb_atma
    Else
    ' ����� ������� ��� ��������� �� ������� �� ���������� �� ��������������
        Pb_atma = Calc_Pb_atma(rsb_m3m3, Tres_C)
    End If
End Property
Property Let Pb_atma(Pbval As Double)
    If (Pbval >= 0) Then
        p_Pb_atma = Pbval
        calculated = False
    Else
        p_Pb_atma = -1 ' ����� ����� �������� �������� ��������� ����� �������� ��� ����� ����������� �� ����������
    End If
End Property

'----- Bo -----------------------------------------------------------------------------------------
Public Property Get Bob_m3m3() As Double
    Bob_m3m3 = p_Bob_m3m3
End Property

Property Let Bob_m3m3(Boval As Double)
    If (Boval >= 0) Then
        p_Bob_m3m3 = Boval
        calculated = False
    Else
        ' ����� ����� �������� �������� ��������� ����� �������� ��� ����� ����������� �� ����������
        p_Bob_m3m3 = -1
    End If
End Property

Public Property Get Muob_cP() As Double
    Muob_cP = p_Muob_cP
End Property

Property Let Muob_cP(muoval As Double)
    If (muoval >= 0) Then
        p_Muob_cP = muoval
        calculated = False
    Else
        p_Muob_cP = -1 ' ����� ����� �������� �������� ��������� ����� �������� ��� ����� ����������� �� ����������
    End If
End Property

Public Property Get rs_m3m3() As Double
    rs_m3m3 = p_Rs_m3m3
End Property

Public Property Get Bo_m3m3() As Double
    Bo_m3m3 = p_Bo_m3m3
End Property

Public Property Get Bg_m3m3() As Double
    Bg_m3m3 = p_Bg_m3m3
End Property

Public Property Get Bw_m3m3() As Double
    Bw_m3m3 = p_Bw_m3m3
End Property

Public Property Get Muo_cP() As Double
 Muo_cP = p_Muo_cP
End Property

Public Property Get Muw_cP() As Double
 Muw_cP = p_Muw_cP
End Property

Public Property Get Mug_cP() As Double
 Mug_cP = p_Mug_cP
End Property

Public Property Get MuLiq_rc_cP() As Double
'
' todo ���� �������� ��� ������� �������� ��� ����� - ���� ����� ���� ������ ������������
'
    MuLiq_rc_cP = (Muo_cP * Qo_m3day / Qliq_m3day + _
                Muw_cP * Qw_m3day / Qliq_m3day)
End Property


Public Property Get MuMix_cP() As Double
'
' todo ���� �������� ��� ������� �������� ��� ����� - ���� ����� ���� ������ ������������
'
    MuMix_cP = p_MuMix_cP
'               (p_Muo_cP * p_Qo_m3day / Qliq_m3day + _
'                p_Muw_cP * p_Qw_m3day / Qliq_m3day) * (1 - GasFraction_d) + _
'                p_Mug_cP * (1 - GasFraction_d)
End Property

' �������������� �������� ����� � ������������
Public Property Get MuMix_cSt() As Double
    MuMix_cSt = p_MuMix_cP / (p_rho_mix_kgm3 / 1000)
End Property

' �������� ������ �����
Public Property Get mo_kgsec() As Double
    mo_kgsec = Qo_m3day * rho_oil_kgm3 / const_conver_day_sec
End Property
' �������� ������ ����
Public Property Get mw_kgsec() As Double
    mw_kgsec = Qw_m3day * rho_water_kgm3 / const_conver_day_sec
End Property
' �������� ������ ����
Public Property Get mg_kgsec() As Double
    mg_kgsec = Qgas_m3day * rho_gas_kgm3 / const_conver_day_sec
End Property

Public Property Get z() As Double
    z = p_Z
End Property

Public Property Get Pcalc_atma() As Double
    Pcalc_atma = p_PTcalc.P_atma
End Property

Public Property Get Pcalc_MPaa() As Double
    Pcalc_MPaa = Pcalc_atma * const_convert_atma_MPa
End Property

Public Property Get Tcalc_C() As Double
    Tcalc_C = p_PTcalc.T_C
End Property

Public Property Get Tcalc_K() As Double
    Tcalc_K = Tcalc_C + const_T_K_min
End Property

Public Property Get Tcalc_F() As Double
    Tcalc_F = Tcalc_C * 1.8 + 32
End Property

Public Property Get Muo_deadoil_cP() As Double
    Muo_deadoil_cP = p_Muo_deadoil_cP
End Property

'===================================================================================
' ������� � ���������
'
'===================================================================================
 
 Private Sub Class_Initialize()
        p_PVT_correlation = StandingBased
        gamma_o = 0.86
        gamma_g = 0.6
        gamma_w = 1
        Rp_m3m3 = 100
        rsb_m3m3 = 100
        Pb_atma = 80
        Bob_m3m3 = 1.2
        Tres_C = 90
        sigma_o_Nm = const_sigma_oil_Nm
        sigma_w_Nm = const_sigma_w_Nm
        wc_perc = 0
        p_Qliq_scm3day = 100
        p_Qgfree_scm3day = 0
        ' ��� ������ ��� �������� �������������� ������������ ������� ��� ���������
        ' ����� ����� ����� �������� ������ � ����������� �� �������
        p_Co_JkgC = 2100
        p_Cw_JkgC = 4100
        p_Cg_JkgC = 2200
 End Sub
 
 
 Public Sub Init(Optional gg = 0.6, Optional gamma_oil = 0.86, Optional gamma_wat = 1, _
                 Optional rsb_m3m3 = 100, Optional Pb_atma = -1, _
                 Optional Bob_m3m3 = -1, Optional PVTcorr = 0, _
                 Optional Tres_C = 90, Optional Rp_m3m3 = -1, Optional Muob_cP = -1)
' P ��������, ���
' � �����������,  �
' gg ��������� ����
' gamma_oil ��������� �����
        gamma_g = gg
        gamma_o = gamma_oil
        gamma_w = gamma_wat
        Me.SetRpRsb CDbl(Rp_m3m3), CDbl(rsb_m3m3)
        Me.Pb_atma = Pb_atma
        Me.Tres_C = Tres_C
        Me.Bob_m3m3 = Bob_m3m3
        Me.Muob_cP = Muob_cP
        PVT_CORRELATION = PVTcorr
 End Sub
 
 Public Function Clone()
        Dim fl As New CPVT
        Call fl.Copy(Me)
        Set Clone = fl
 End Function
 
 Public Function Copy(fl As CPVT)
        PVT_CORRELATION = fl.PVT_CORRELATION
        ZNLF = fl.ZNLF
        gamma_o = fl.gamma_o
        gamma_g = fl.gamma_g
        gamma_w = fl.gamma_w
        SetRpRsb fl.Rp_m3m3, fl.rsb_m3m3
        Pb_atma = fl.Pb_atma
        Bob_m3m3 = fl.Bob_m3m3
        Muob_cP = fl.Muob_cP
        Tres_C = fl.Tres_C
        wc_fr = fl.wc_fr
        Qliq_scm3day = fl.Qliq_scm3day
        Qgfree_scm3day = fl.Qgfree_scm3day
        sigma_o_Nm = fl.sigma_o_Nm
        sigma_w_Nm = fl.sigma_w_Nm
 End Function

Public Sub Calc_PVT_PT(PT As PTtype)
        Call Calc_PVT(PT.P_atma, PT.T_C)
End Sub

Public Sub Calc_PVT(P_atma As Double, T_C As Double)
' ������ ������� ���� ����� � ���� ��� �������� �������� � �����������
'    Dim timeStamp
'    timeStamp = Time()
    Dim T_K As Double  ' internal K temp
    'PVT properties
    Dim rho_o As Double
    Dim Bob_m3m3_sat As Double
    'internal buffers used to store output values
    Dim p_bi As Double
    Dim r_si As Double
    Dim rho_o_sat As Double
    Dim p_fact As Double
    Dim p_offs As Double
    Dim b_fact As Double
    Dim mu_fact As Double
    'Oil pressure in MPa
    Dim P_MPa As Double
    Dim Pb_calbr_MPa As Double
    Dim Rsb_calbr_m3m3 As Double
    Dim Bo_calbr_m3m3 As Double
    Dim muo_calibr_cP As Double
    Dim ranges_good As Boolean
    Dim Muo_deadoil_cP As Double
    Dim Muo_saturated_cP As Double
On Error GoTo err1:
    T_K = T_C + const_T_K_min
    p_PTcalc.P_atma = P_atma
    p_PTcalc.T_C = T_C
    Rsb_calbr_m3m3 = rsb_m3m3
    Bo_calbr_m3m3 = Bob_m3m3
    muo_calibr_cP = Muob_cP
    P_MPa = P_atma * const_convert_atma_MPa
    'convert user specified bubblepoint pressure
    Pb_calbr_MPa = p_Pb_atma * const_convert_atma_MPa    ' ������� ����� �������� �������� ��� ������� Pb � ������� ���������
    'for saturated oil calibration is applied by application of factor p_fact to input pressure
    'for undersaturated - by shifting according to p_offs
    'calculate PVT properties
    'calculate water properties at current pressure and temperature
    p_BwSC_m3m3 = gamma_w / 1
    p_Salinity_ppm = unf_Calc_Sal_BwSC_ppm(p_BwSC_m3m3)
    p_Bw_m3m3 = unf_Calc_Bw_m3m3(P_MPa, T_K) '* p_BwSC_m3m3
    p_Muw_cP = unf_Calc_Muw_cP(P_MPa, T_K, p_Salinity_ppm)   ' ������ ��� ��������� ���������?!!!
    'if no calibration gas-oil ratio specified, then set it to some very large value and
    'switch of calibration for bubblepoint and oil formation volume factor
    p_Z = unf_Calc_Zgas_d(T_K, P_MPa, p_gamma_g, 1)
    p_Bg_m3m3 = unf_calc_Bg_z_m3m3(T_K, P_MPa, p_Z)
    p_Mug_cP = unf_calc_Mug_cP(T_K, P_MPa, p_Z, p_gamma_g)
    If PVT_CORRELATION = StandingBased Then
            Muo_deadoil_cP = unf_DeadOilViscosity_Beggs_Robinson(T_K, p_gamma_o) 'dead oil viscosity
            Muo_saturated_cP = unf_SaturatedOilViscosity_Beggs_Robinson(Rsb_calbr_m3m3, Muo_deadoil_cP) 'saturated oil viscosity Beggs & Robinson
            p_bi = unf_Bubblepoint_Standing(Rsb_calbr_m3m3, p_gamma_g, Tres_K, p_gamma_o)    ' ������� �������� ��������� �� ���������� Standing ��� ��������� ����������� ��� ������� ������� ������������� ��������
            ' ������ ���� ������������� ������������
            'Calculate bubble point correction factor
            If (Pb_calbr_MPa > 0) Then 'user specified
                p_fact = p_bi / Pb_calbr_MPa
            Else ' not specified, use from correlations
                p_fact = 1
            End If
            If (Bo_calbr_m3m3 > 0) Then 'Calculate oil formation volume factor correction factor
                Bob_m3m3_sat = unf_FVF_Saturated_Oil_Standing(Rsb_calbr_m3m3, p_gamma_g, Tres_K, p_gamma_o)  ' �������� �� ���������� ������� ����� ��� ��������� �����������
                b_fact = (Bo_calbr_m3m3 - 1) / (Bob_m3m3_sat - 1)
            Else ' not specified, use from correlations
                b_fact = 1
            End If
            If muo_calibr_cP > 0 Then           ' ���������� ������������� ����������� ��� �������� ��� �������� ���������
                mu_fact = muo_calibr_cP / Muo_saturated_cP
            Else
                mu_fact = 1
            End If
            p_bi = unf_Bubblepoint_Standing(Rsb_calbr_m3m3, p_gamma_g, T_K, p_gamma_o)   ' �������� ��������� �� ���������� ��� ������� �����������
            P_MPa = P_MPa * p_fact   ' �������� �������� ����� �������� ��� �� ������������� ��������
            If P_MPa > p_bi Then 'apply correction to undersaturated oil 'undersaturated oil
                 r_si = Rsb_calbr_m3m3   ' �������� �������� ����� �����
                 Bob_m3m3_sat = b_fact * (unf_FVF_Saturated_Oil_Standing(Rsb_calbr_m3m3, p_gamma_g, T_K, p_gamma_o) - 1) + 1 ' it is assumed that at pressure 1 atma bo=1
                 p_copmressibility_o_1atm = unf_Compressibility_Oil_VB(Rsb_calbr_m3m3, p_gamma_g, T_K, p_gamma_o, P_MPa) 'calculate compressibility at bubble point pressure
                 p_Bo_m3m3 = Bob_m3m3_sat * Exp(p_copmressibility_o_1atm * (p_bi - P_MPa))
                 p_Muo_cP = mu_fact * unf_OilViscosity_Vasquez_Beggs(Muo_saturated_cP, P_MPa, p_bi)  'Vesquez&Beggs
            Else 'apply correction to saturated oil
                r_si = unf_GOR_Standing(P_MPa, p_gamma_g, T_K, p_gamma_o)
                p_Bo_m3m3 = b_fact * (unf_FVF_Saturated_Oil_Standing(r_si, p_gamma_g, T_K, gamma_o) - 1) + 1 ''Standing. it is assumed that at pressure 1 atma bo=1
                p_Muo_cP = mu_fact * unf_SaturatedOilViscosity_Beggs_Robinson(r_si, Muo_deadoil_cP)  'Beggs & Robinson
            End If
    End If
    If PVT_CORRELATION = McCainBased Then
            ranges_good = True
            ranges_good = ranges_good And CheckRanges(T_K, "t_K", const_TMcCain_K_min, const_T_K_max, "����������� ������ ��� ��������� ��� ���������� ��������", "calc_PVT (McCain)", True)
            Muo_deadoil_cP = unf_DeadOilViscosity_Standing(T_K, p_gamma_o)  'dead oil viscosity
            Muo_saturated_cP = unf_SaturatedOilViscosity_Beggs_Robinson(Rsb_calbr_m3m3, Muo_deadoil_cP) 'saturated oil viscosity Beggs & Robinson
            p_bi = unf_Bubblepoint_Valko_McCainSI(Rsb_calbr_m3m3, p_gamma_g, Tres_K, p_gamma_o)
            'Calculate bubble point correction factor
            If (Pb_calbr_MPa > 0) Then 'user specifie
                p_fact = p_bi / Pb_calbr_MPa
            Else ' not specified, use from correlations
                p_fact = 1
            End If
            P_MPa = P_MPa * p_fact
            p_copmressibility_o_1atm = unf_Compressibility_Oil_VB(Rsb_calbr_m3m3, p_gamma_g, Tres_K, gamma_o, p_bi) 'calculate compressibility at bubble point pressure
            If (Bo_calbr_m3m3 > 0) Then 'Calculate oil formation volume factor correction factor
                  rho_o_sat = unf_Density_McCainSI(p_bi, p_gamma_g, Tres_K, gamma_o, Rsb_calbr_m3m3, p_bi, p_copmressibility_o_1atm)    ' ��� ��������� ���� ����������� �� ����������� �� ������� ��� �� ������ (��� ��� ������ ���� ��� �������� ���������)
                  Bob_m3m3_sat = unf_FVF_McCainSI(Rsb_calbr_m3m3, p_gamma_g, p_gamma_o * const_rho_ref, rho_o_sat)
                  b_fact = (Bo_calbr_m3m3 - 1) / (Bob_m3m3_sat - 1)
            Else ' not specified, use from correlations
                  b_fact = 1
            End If
            p_bi = unf_Bubblepoint_Valko_McCainSI(Rsb_calbr_m3m3, p_gamma_g, T_K, p_gamma_o)
            r_si = unf_GOR_VelardeSI(P_MPa, p_bi, p_gamma_g, T_K, p_gamma_o, Rsb_calbr_m3m3)
            If P_MPa > p_bi Then 'apply correction to undersaturated oil
                  p_copmressibility_o_1atm = unf_Compressibility_Oil_VB(Rsb_calbr_m3m3, p_gamma_g, T_K, p_gamma_o, P_MPa)  'calculate compressibility at bubble point pressure
                  rho_o_sat = unf_Density_McCainSI(p_bi, p_gamma_g, T_K, p_gamma_o, Rsb_calbr_m3m3, p_bi, p_copmressibility_o_1atm)
                  Bob_m3m3_sat = unf_FVF_McCainSI(Rsb_calbr_m3m3, p_gamma_g, p_gamma_o * const_rho_ref, rho_o_sat)
                  Bob_m3m3_sat = b_fact * (Bob_m3m3_sat - 1) + 1 ' it is assumed that at pressure 1 atma bo=1
                  p_Bo_m3m3 = Bob_m3m3_sat * Exp(p_copmressibility_o_1atm * (p_bi - P_MPa))
            Else 'apply correction to saturated oil
                  rho_o = unf_Density_McCainSI(P_MPa, p_gamma_g, T_K, p_gamma_o, r_si, p_bi, p_copmressibility_o_1atm)
                  p_Bo_m3m3 = b_fact * (unf_FVF_McCainSI(r_si, p_gamma_g, p_gamma_o * const_rho_ref, rho_o) - 1) + 1 ' it is assumed that at pressure 1 atma bo=1
            End If
            If muo_calibr_cP > 0 Then           ' ���������� ������������� ����������� ��� �������� ��� �������� ���������
                  If (Rsb_calbr_m3m3 < 350) Then
                       mu_fact = muo_calibr_cP / unf_Oil_Viscosity_Standing(Rsb_calbr_m3m3, Muo_deadoil_cP, p_bi, p_bi)
                  Else
                        mu_fact = muo_calibr_cP / Muo_saturated_cP
                  End If
            Else
                  mu_fact = 1
            End If
            If (Rsb_calbr_m3m3 < 350) Then 'Calculate oil viscosity acoording to Standing
                  p_Muo_cP = mu_fact * unf_Oil_Viscosity_Standing(r_si, Muo_deadoil_cP, P_MPa, p_bi)
            Else 'Calculate according to Begs&Robinson (saturated) and Vasquez&Begs (undersaturated)
               If P_MPa > p_bi Then 'undersaturated oil
                  p_Muo_cP = mu_fact * unf_OilViscosity_Vasquez_Beggs(Muo_saturated_cP, P_MPa, p_bi)
               Else 'saturated oil
                 'Beggs & Robinson
                 p_Muo_cP = mu_fact * unf_SaturatedOilViscosity_Beggs_Robinson(r_si, Muo_deadoil_cP)
               End If
            End If
    End If
    If PVT_CORRELATION = 2 Then  'Debug mode. Linear Rs and bo vs P, Pb_calbr_atma should be specified.
         'gas properties
         p_Z = 0.95 'ideal gas
         p_Bg_m3m3 = unf_calc_Bg_z_m3m3(T_K, P_MPa, p_Z)
         p_Mug_cP = 0.0000000001
         p_fact = 1         'Set to default. b_rb should be specified by user!
         p_offs = 0
         p_bi = Pb_calbr_MPa
         If P_MPa > (p_bi) Then 'undersaturated oil
            r_si = Rsb_calbr_m3m3
         Else 'saturate
            r_si = P_MPa / Pb_calbr_MPa * Rsb_calbr_m3m3
         End If
         'if Bob_m3m3 is not specified by the user then
         'set Bob_m3m3 so, that oil density, recalculated with Rs_m3m3 would be equal to dead oil density
         If (Bo_calbr_m3m3 < 0) Then
            p_Bo_m3m3 = (1 + r_si * (p_gamma_g * const_rho_air) / (p_gamma_o * const_rho_ref))
         Else
            If P_MPa > (p_bi) Then 'undersaturated oil
                p_Bo_m3m3 = Bo_calbr_m3m3
            Else 'saturate
                p_Bo_m3m3 = 1 + (Bo_calbr_m3m3 - 1) * ((P_MPa - const_convert_atma_MPa) / (p_bi - const_convert_atma_MPa))
            End If
         End If
         If muo_calibr_cP > 0 Then
             p_Muo_cP = muo_calibr_cP
         Else
            p_Muo_cP = 1
         End If
    End If
    'Assign output variables
    p_Pbcalc_atma = p_bi / p_fact / const_convert_atma_MPa
    p_Rs_m3m3 = r_si
    p_Muo_deadoil_cP = Muo_deadoil_cP
    p_Qo_m3day = p_Qliq_scm3day * (1 - p_wc_fr) * p_Bo_m3m3   ' ��� ��������� �������� ����� ��� ��� ����� ���������� ���
    p_Qw_m3day = p_Qliq_scm3day * p_wc_fr * p_Bw_m3m3
    p_Qgas_m3day = (p_Qliq_scm3day * (1 - p_wc_fr) * p_Rp_m3m3 + p_Qgfree_scm3day - p_Rs_m3m3 * p_Qliq_scm3day * (1 - p_wc_fr)) * p_Bg_m3m3
    p_Qliq_m3day = p_Qw_m3day + p_Qo_m3day
    p_GasFraction_d = p_Qgas_m3day / (p_Qw_m3day + p_Qo_m3day + p_Qgas_m3day)
    p_MuMix_cP = (p_Muo_cP * p_Qo_m3day / p_Qliq_m3day + _
                  p_Muw_cP * p_Qw_m3day / p_Qliq_m3day) * (1 - p_GasFraction_d) + _
                  p_Mug_cP * (1 - p_GasFraction_d)
    p_rho_oil_kgm3 = 1000 * (p_gamma_o + p_Rs_m3m3 * p_gamma_g * const_rho_air / 1000) / p_Bo_m3m3
    p_rho_water_kgm3 = 1000 * (p_gamma_w) / p_Bw_m3m3
    p_rho_liq_kgm3 = (1 - p_wc_fr) * rho_oil_kgm3 + wc_fr * rho_water_kgm3
    p_rho_mix_kgm3 = rho_liq_kgm3 * (1 - f_g) + rho_gas_kgm3 * f_g
    
    Call calc_ST(P_atma, T_C)
'    timeStamp = Time() - timeStamp
'    timePVTtotal = timePVTtotal + timeStamp
    Exit Sub
err1:
    addLogMsg ("CPVT.Calc_PVT: ������ ����� ��")
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub calc_ST(ByVal P_atma As Double, ByVal T_C As Double)
' calculate surface tension according Baker Sverdloff correlation

'������ ������������ �������������� ��������� ���-�����
    Dim ST68 As Double, ST100 As Double
    Dim STw74 As Double, STw280 As Double
    Dim Tst As Double, Tstw As Double
    Dim STo As Double, STw As Double, ST As Double
    Dim T_F As Double
    Dim P_psia As Double, P_MPa As Double
        T_F = T_C * 1.8 + 32
        P_psia = P_atma / 0.068046
        P_MPa = P_atma / 10
        ST68 = 39 - 0.2571 * oil_API
        ST100 = 37.5 - 0.2571 * oil_API
        If T_F < 68 Then
            STo = ST68
        Else
            Tst = T_F
            If T_F > 100 Then Tst = 100
            STo = (68 - (((Tst - 68) * (ST68 - ST100)) / 32)) * (1 - (0.024 * (P_psia) ^ 0.45))
        End If
        '������ ������������ �������������� ��������� ���-����  (��� �������)
        STw74 = (75 - (1.108 * (P_psia) ^ 0.349))
        STw280 = (53 - (0.1048 * (P_psia) ^ 0.637))
        If T_F < 74 Then
            STw = STw74
        Else
            Tstw = T_F
            If T_F > 280 Then Tstw = 280
            STw = STw74 - (((Tstw - 74) * (STw74 - STw280)) / 206)
        End If
        ' ����� ������ ������
        STw = 10 ^ (-(1.19 + 0.01 * P_MPa)) * 1000
        ' ������ ������������ �������������� ��������� ���-��������
        ST = (STw * wc_fr) + STo * (1 - wc_fr)

        p_ST_oilgas_dyncm = STo
        p_ST_watgas_dyncm = STw
        p_ST_liqgas_dyncm = ST
End Sub

Public Function calc_Rs_m3m3(ByVal P_atma As Double, ByVal T_C As Double) As Double
'function calculates solution gas oil ratio
    Call Calc_PVT(P_atma, T_C)
    calc_Rs_m3m3 = p_Rs_m3m3
End Function

Public Function Calc_Pb_atma(ByVal rsb_m3m3 As Double, ByVal T_C As Double) As Double
'function calculates oil bubble point pressure
    p_Rsb_m3m3 = rsb_m3m3
    Call Calc_PVT(1, T_C)
    Calc_Pb_atma = p_Pbcalc_atma
End Function

Public Function Calc_Bo_m3m3(ByVal P_atma As Double, ByVal T_C As Double) As Double
'Function calculates oil formation volume factor
    Call Calc_PVT(P_atma, T_C)
    Calc_Bo_m3m3 = p_Bo_m3m3
End Function

Public Function Calc_Muo_cP(ByVal P_atma As Double, ByVal T_C As Double) As Double
'function calculates oil viscosity
    Call Calc_PVT(P_atma, T_C)
    Calc_Muo_cP = p_Muo_cP
End Function

Public Function Qgas_cas_scm3day(Optional ByVal Ksep As Double = -1) As Double
   If Ksep > 0 And Ksep < 1 Then
        Qgas_cas_scm3day = Qgas_m3day * Ksep
   Else
        Qgas_cas_scm3day = 0
   End If
End Function

Public Function GasFraction_d(Optional ByVal Ksep As Double = 0) As Double
' ����� ������� ���� ���� � ������ ��� �������� �������� ��� �������� ��������
' �������������� ��� �������� ����� ���� � ���� ��� ��������� � ������ ��� ����������� ��������
    Dim qmix As Double
    GasFraction_d = 0
    qmix = Qmix_m3day   ' �������� ����� ������� ���������� �� �������� �������� ��������
    If qmix > 0 And Ksep >= 0 And Ksep < 1 Then
        GasFraction_d = Qgas_m3day * (1 - Ksep) / (Qw_m3day + Qo_m3day + Qgas_m3day * (1 - Ksep))
    Else
        addLogMsg "GasFraction_d: ����� ������ ��� (qo + qw + qg) = " & qmix
    End If
End Function

Public Function PGasFraction_atma(FreeGas, T_C As Double, Optional Es As Double = 0, Optional P_init_atma = 300) As Double
    'P_init     - �������� �������������, ���
    'FreeGas    - ���� ��� �� ������ �������
    'Es         - ����������� ��������� ������
    Dim P1 As Double
    Dim P2 As Double
    Dim max_iter As Integer, i As Integer
    Dim E As Double
    Dim p_gas As Double, P As Double
    max_iter = 100
    E = 0.0001
    P1 = P_init_atma
    P2 = 0
On Error GoTo err1:
    For i = 1 To max_iter
        P = (P1 + P2) / 2
        Call Calc_PVT(P, T_C)

       ' Q_g = Ql * (1 - Wc / 100) * (Gf_1 - Calc_Rs(P11, t, Ro_o, Ro_g, PVT_cor, Gf_1)) * GasFVF_gamma(t, P12, Ro_g, PVT_cor) * (1 - Es)
       ' Q_l = Ql * ((1 - Wc / 100) * Calc_Bo(P21, t, Ro_o, Ro_g, PVT_cor, Gf_1) + Wc / 100 * Calc_Bw(P22, t, , PVT_cor))
        p_gas = GasFraction_d(Es)
        If Abs(p_gas - FreeGas) <= E Then Exit For
        If p_gas > FreeGas Then
           P2 = P
        Else
           P1 = P
        End If
    Next
    PGasFraction_atma = P
    Exit Function
err1:
    PGasFraction_atma = 0
End Function


Public Function RpGasFraction_m3m3(FreeGas, P_atma As Double, T_C As Double, Optional Es As Double = 0, Optional Rp_init_m3m3 = 500) As Double
    'P_init     - �������� �������������, ���
    'FreeGas    - ���� ��� �� ������ �������
    'Es         - ����������� ��������� ������
    Dim g1 As Double
    Dim g2 As Double
    Dim max_iter As Integer, i As Integer
    Dim E As Double
    Dim p_gas As Double, G As Double
    max_iter = 100
    E = 0.0001
    g1 = Rp_init_m3m3
    g2 = 0
On Error GoTo err1:
    For i = 1 To max_iter
        G = (g1 + g2) / 2
        Rp_m3m3 = G
        Call Calc_PVT(P_atma, T_C)

       ' Q_g = Ql * (1 - Wc / 100) * (Gf_1 - Calc_Rs(P11, t, Ro_o, Ro_g, PVT_cor, Gf_1)) * GasFVF_gamma(t, P12, Ro_g, PVT_cor) * (1 - Es)
       ' Q_l = Ql * ((1 - Wc / 100) * Calc_Bo(P21, t, Ro_o, Ro_g, PVT_cor, Gf_1) + Wc / 100 * Calc_Bw(P22, t, , PVT_cor))
        p_gas = GasFraction_d(Es)
        If Abs(p_gas - FreeGas) <= E Then Exit For
        If p_gas > FreeGas Then
           g1 = G
        Else
           g2 = G
        End If
    Next
    RpGasFraction_m3m3 = G
    Exit Function
err1:
    RpGasFraction_m3m3 = 0
End Function

Public Function GetCloneModAfterSeparation(P_atma As Double, T_C As Double, Ksep As Double, _
                                         Optional ByVal GasSol As GAS_INTO_SOLUTION = GasnotGoesIntoSolution) As CPVT
    Dim newFluid As CPVT
    Set newFluid = Me.Clone
    Call newFluid.ModAfterSeparation(P_atma, T_C, Ksep, GasSol)
    Set GetCloneModAfterSeparation = newFluid
End Function

Public Sub ModAfterSeparation(ByVal P_atma As Double, ByVal T_C As Double, ByVal Ksep As Double, _
                                         Optional ByVal GasSol As GAS_INTO_SOLUTION = GasnotGoesIntoSolution) 'As CPVT
' ������� ����������� ������� ����� ����� ���������
' �������� ����� ���� ������ �������� ����� - ������ ���������� ���� �������� �� �������
' �� ����� ������� ��� ������� ��������� ���������
    Dim Rs As Double
    Dim Bo As Double
    Dim Pb_Rs_curve As New TInterpolation ' ��������� ������ ����������� �������������� �� �������� ���������
    Dim Bo_Rs_curve As New TInterpolation
    Dim Pb_atma_tab As Double, Rsb_m3m3_tab As Double, Bo_m3m3_tab As Double
    Dim delta As Double
    Dim i As Integer
    Const N = 10
    Dim Rpnew As Double
        ' ������ ������� ���� �������� � �������� ��� �������� ���������
    With Me
        Rs = .calc_Rs_m3m3(P_atma, T_C)
        Bo = .Calc_Bo_m3m3(P_atma, T_C)
        
        If GasSol = GasGoesIntoSolution Then   ' ����� ��� ������ �����������
            Rpnew = .Rp_m3m3 - (.Rp_m3m3 - Rs) * Ksep
        Else                                   ' ��� �� ������������, �� �� �����, ��� ���� = 1
            Rpnew = .Rp_m3m3 - (.Rp_m3m3 - Rs)
        End If
        delta = (.Pb_atma - 1) / N    ' ������� ����� ������ � ��������� ��� ���������� Pb �� ��� ����� ������� ����������������
        ' ������� ����������� �������������� �� �������� ��������� �� ������
        For i = 0 To N
            Pb_atma_tab = 1 + delta * i
            Rsb_m3m3_tab = .calc_Rs_m3m3(Pb_atma_tab, .Tres_C)
            Pb_Rs_curve.AddPoint Rsb_m3m3_tab, Pb_atma_tab
            Bo_m3m3_tab = .Calc_Bo_m3m3(Pb_atma_tab, .Tres_C)
            Bo_Rs_curve.AddPoint Rsb_m3m3_tab, Bo_m3m3_tab
        Next i
        ' ������ ������� ����� ���� �������� � ������
        If Rpnew < .rsb_m3m3 Then
            ' ���� ������� ������ ���������� ������ ��������������, ����� ���� ��������������� �������������� � �������� ���������,
            ' ������� ����� �� ���� ��������
            .Pb_atma = Pb_Rs_curve.GetPoint(Rpnew)
            .Bob_m3m3 = Bo_Rs_curve.GetPoint(Rpnew)
            .rsb_m3m3 = Rpnew
            ' ����� ���� �� �������� �� ������������� - �������� �� �������� ������ ������ �� ����
        End If
    .Rp_m3m3 = Rpnew
    End With
End Sub

Public Sub BuildCurves(T_C As Double)
' ����� ��� ���������� ���������� �������� �� PVT ���������
    Dim i As Integer
    ' ������ ����������� ������������ ��� ��������� �����������
    c_RsTres_Curve.ClearPoints
    c_RsT_Curve.ClearPoints
    c_BoTres_Curve.ClearPoints
    c_BoT_Curve.ClearPoints
    c_MuoTres_Curve.ClearPoints
    c_MuoT_Curve.ClearPoints
    c_MugTres_Curve.ClearPoints
    c_MugT_Curve.ClearPoints
    p_Twh_C = T_C
        Dim NumPointBeforePb As Integer
        NumPointBeforePb = 15
        Dim Pcalc As Double
        Dim Pmin As Double, Pmax As Double
        Pmin = 1
        Pmax = p_Pb_atma
        For i = 0 To NumPointBeforePb + 5
            Pcalc = Pmin + (Pmax - Pmin) * i / NumPointBeforePb
            Calc_PVT Pcalc, Tres_C
            c_RsTres_Curve.AddPoint Pcalc, p_Rs_m3m3
            c_BoTres_Curve.AddPoint Pcalc, p_Bo_m3m3
            c_MuoTres_Curve.AddPoint Pcalc, p_Muo_cP
            c_MugTres_Curve.AddPoint Pcalc, p_Mug_cP
        Next i
        Pmin = 1
        Pmax = Calc_Pb_atma(p_Rsb_m3m3, T_C)
        For i = 0 To NumPointBeforePb + 5
            Pcalc = Pmin + (Pmax - Pmin) * i / NumPointBeforePb
            Calc_PVT Pcalc, T_C
            c_RsT_Curve.AddPoint Pcalc, p_Rs_m3m3
            c_BoT_Curve.AddPoint Pcalc, p_Bo_m3m3
            c_MuoT_Curve.AddPoint Pcalc, p_Muo_cP
            c_MugT_Curve.AddPoint Pcalc, p_Mug_cP
        Next i
End Sub

 Public Function SaveState()
 ' ��������� ��������� ������� � ���������� ������ ��� ����������� ������ (��� �������)
    Dim stor()
    Dim i As Integer
    ReDim stor(const_OutputCurveNumPoints, STOR_SIZE)
    stor(0, 0) = "PVT SaveState"   ' ��� ����� ����� ������ � ����� ���������� �����
    AddS stor, 1, 0, "base props    "
    AddS stor, 2, 0, "p_gamma_o              ", p_gamma_o ' ��������� ����� ��������
    AddS stor, 3, 0, "p_gamma_g                ", p_gamma_g ' ��������� ���� ��������
    AddS stor, 4, 0, "p_gamma_w             ", p_gamma_w ' ��������� ���� ��������
    AddS stor, 5, 0, "p_Rp_m3m3             ", p_Rp_m3m3 ' ������� ������ �������� (����������� � ����������� �������� )
    AddS stor, 6, 0, "������������� ���������                    ", ""
    AddS stor, 7, 0, "p_Rsb_m3m3                    ", p_Rsb_m3m3 ' �������������� ��� �������� ���������
    AddS stor, 8, 0, "p_Pb_atma                    ", p_Pb_atma ' �������� ���������
    AddS stor, 9, 0, "p_Bob_m3m3                    ", p_Bob_m3m3 ' �������� ����������� ��� �������� ���������
    AddS stor, 10, 0, "p_Tres_C         ", p_Tres_C  ' ����������� ��� ������� ������ ������������� ���������
    AddS stor, 11, 0, "p_Twh_C         ", p_Twh_C
    AddS stor, 12, 0, "p_Pcalc_atma         ", Pcalc_atma
    AddS stor, 13, 0, "p_Tcalc_C        ", Tcalc_C
    AddS stor, 14, 0, "p_Tcalc_K               ", Tcalc_C
    AddS stor, 15, 0, "--                       ", ""
    AddS stor, 16, 0, "p_Pbcalc_atma                       ", p_Pbcalc_atma ' ��������� �������� �������� ��������� �� ����������
    AddS stor, 17, 0, "p_Rs_m3m3   ", p_Rs_m3m3 ' ��������� �������� �������������� � ����� ��� ������� ��������
    AddS stor, 18, 0, "p_Bo_m3m3       ", p_Bo_m3m3 ' �������� ����������� ����� ��� ������� ��������
    AddS stor, 19, 0, "p_Muo_cP                     ", p_Muo_cP ' �������� ����� ��� ������� ��������
    AddS stor, 20, 0, "p_Mug_cP            ", p_Mug_cP ' �������� ���� ��� ������� ��������
    AddS stor, 21, 0, "p_Z        ", p_Z
    AddS stor, 22, 0, "p_Bw_m3m3     ", p_Bw_m3m3
    AddS stor, 23, 0, "p_Bg_m3m3                ", p_Bg_m3m3
    AddS stor, 24, 0, "p_Muw_cP                 ", p_Muw_cP
    AddS stor, 25, 0, "p_copmressibility_o_1atm             ", p_copmressibility_o_1atm
   ' AddS stor, 26, 0, "p_Muo_deadoil_cP  ", Muo_deadoil_cP
   ' AddS stor, 27, 0, "p_Muo_saturated_cP          ", p_Muo_saturated_cP
    AddS stor, 28, 0, "--   ", ""
    AddS stor, 29, 0, "p_PVT_correlation           ", p_PVT_correlation
    AddS stor, 30, 0, "params               "
    AddS stor, 31, 0, "p_Salinity_ppm        ", p_Salinity_ppm
    AddS stor, 32, 0, "sigma_o    ", sigma_o_Nm
    AddS stor, 33, 0, "--     ", ""
    AddS stor, 33, 0, "--     ", ""
    Const flStart = 1
    i = 0
    AddS stor, 0, 4, "P, atma", " RsTres"
    AddSCurve stor, 1, 4, c_RsTres_Curve
    AddS stor, 0, 6, "P, atma", " RsT"
    AddSCurve stor, 1, 6, c_RsT_Curve
    AddS stor, 0, 8, "P, atma", " BoTres"
    AddSCurve stor, 1, 8, c_BoTres_Curve
    AddS stor, 0, 10, "P, atma", " BoT"
    AddSCurve stor, 1, 10, c_BoT_Curve
    AddS stor, 0, 12, "P, atma", " MuoTres"
    AddSCurve stor, 1, 12, c_MuoTres_Curve
    AddS stor, 0, 14, "P, atma", " MuoT"
    AddSCurve stor, 1, 14, c_MuoT_Curve
    AddS stor, 0, 16, "P, atma", " MugTres"
    AddSCurve stor, 1, 16, c_MugTres_Curve
    AddS stor, 0, 18, "stage num", " MugT"
    AddSCurve stor, 1, 18, c_MugT_Curve
    SaveState = stor
 End Function



