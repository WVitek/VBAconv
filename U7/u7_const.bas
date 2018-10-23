Attribute VB_Name = "u7_const"
'=======================================================================================
'Unifloc7.3  Testudines                                           khabibullinra@gmail.com
'���������� ��������� ������� �� ��������� �����������
'2000 - 2018 �
'
'=======================================================================================
' constant definition module
'

'
' ��������� ��� ������������ ����������� ���������� ������������ ��� ��������
'

Public Const const_default_wrong_read_num As Double = -999999

Public Const const_name_Qliq_m3day          As String = "Qliq_m3day"
Public Const const_name_wc_perc             As String = "wc_perc"
Public Const const_name_Pcas_atma            As String = "Pcas_atma"
Public Const const_name_Hdyn_m              As String = "Hdyn_m"
Public Const const_name_Pline_atma           As String = "Pline_atma"
Public Const const_name_Pbuf_atma            As String = "Pbuf_atma"
Public Const const_name_Pwf_atma             As String = "Pwf_atma"

Public Const const_name_gamma_oil           As String = "gamma_oil"
Public Const const_name_gamma_water         As String = "gamma_water"
Public Const const_name_gamma_gas           As String = "gamma_gas"
Public Const const_name_Rp_m3m3             As String = "Rp_m3m3"
Public Const const_name_Rsb_m3m3            As String = "Rsb_m3m3"
Public Const const_name_Pb_atma              As String = "Pb_atma"
Public Const const_name_HResMes_m           As String = "HResMes_m"
Public Const const_name_HPumpMes_m          As String = "HPumpMes_m"
Public Const const_name_Dchoke_mm           As String = "Dchoke_mm, mm"
Public Const const_name_�������������_�     As String = "�������������_�, �"
Public Const const_name_ESP_Qliq_m3day      As String = "ESP_Qliq_m3day"
Public Const const_name_ESP_NumStages       As String = "ESP_NumStages"
Public Const const_name_ESP_Freq_Hz         As String = "ESP_Freq_Hz"
Public Const const_name_ESP_Pintake_atma     As String = "ESP_Pintake_atma"
Public Const const_name_Pres_atma            As String = "Pres_atma"
Public Const const_name_PI_m3dayatm         As String = "PI_m3dayatm"
Public Const const_name_ESP_Tintake_C       As String = "ESP_Tintake_C"
Public Const const_name_Tres_C              As String = "Tres_C"


Public Const const_name_HmesCurve       As String = "HmesCurve"
Public Const const_name_DcasCurve       As String = "DcasCurve"
Public Const const_name_DtubCurve       As String = "DtubCurve"
Public Const const_name_TAmbCurve       As String = "TAmbCurve"
'
'
'

'Public Const a = 1

Public Const str_VLPcurve = "VLPcurve"                      ' ������ ������ -  ����������� ��������� �������� �� ������ ��������
Public Const str_HvertCurve = "HvertCurve"                  ' ������ ���������� (��������) ��������

Public Const str_DcasCurve = "DcasCurve"                    ' ������ ��������� �������� ���������������� �������
Public Const str_DtubCurve = "DtubCurve"                    ' ������ ��������� �������� ���

Public Const str_RoughnessCasCurve = "RoughnessCasCurve"    ' ������ ��������� ������������� �� ����� ���������������� �������
Public Const str_RoughnessTubCurve = "RoughnessTubCurve"    ' ������ ��������� ������������� �� ����� ���

Public Const str_Hd_Depend_Pwf = "Hd_Depend_Pwf"            ' ������ - ����������� ������������� ������ �� ��������� ��������, ��� �������� �������� � �������
Public Const str_Pan_Depend_Pwf = "Pan_Depend_Pwf"          ' ������ - ����������� ���������� �������� �� ��������� ��������

' ����������� ��� �������� � ��������� �������� �� ��� ������ ������� ���������� �� ����� �������
Public Const str_Plin_Depend_Pwf = "Plin_Depend_Pwf"        ' ������ - ����������� ��������� �������� �� ��� ������
Public Const str_Pbuf_Pwf_curve = "Pbuf_Pwf_curve"          ' ����������� ��������� �������� �� ��� ������
 
Public Const str_KsepNatQl_curve = "KsepNatQl_curve"         ' ����������� ������������ ��������� �� ������
Public Const str_KsepNatRp_curve = "KsepNatRp_curve"         ' ����������� ������������ ��������� �� �������� �������

Public Const str_KsepTotalQl_curve = "KsepTotalQl_curve"     ' ������ ������ ������������ ��������� �� ������
Public Const str_KsepTotalRp_curve = "KsepTotalRp_curve"     ' ������ ������ ������������ ��������� �� �������� �������
Public Const str_KsepGasSepQl_curve = "KsepGasSepQl_curve"   ' ������ ������������ ��������� �������������� �� ������
Public Const str_KsepGasSepRp_curve = "KsepGasSepQl_curve"   ' ������ ������������ ��������� �������������� �� �������� �������

Public Const str_PdisKdegr_curve = "PdisKdegr_curve"         ' ������ ����������� �������� �� ����� �� ���������� ������ ����


Public Const str_TambHmes_curve = "TambHmes_curve"           ' ������� ����������� ����������� ����������� �� ���������� ����������

Public Const str_PtubHmes_curve = "PtubHmes_curve"           ' ������� �������� �� ������ �������� �� ���� ��� � �� ��� �� �����
Public Const str_TtubHmes_curve = "TtubHmes_curve"           ' ������� ����������� �� ������ �������� �� ���

Public Const str_PcasHmes_curve = "PcasHmes_curve"           ' ������� �������� �� ������ �������� �� ���� ��� � �� ������� �� �����
Public Const str_TcasHmes_curve = "TcasHmes_curve"           ' ������� ����������� �� ������ �������� ���� ������ � ���� ������ �� �������

Public Const str_RstubHmes_curve = "RstubHmes_curve"         ' ������� ����������� ���������� ���� � ����� �� ������ � ���
Public Const str_RscasHmes_curve = "RscasHmes_curve"         ' ������� ����������� ���������� ���� � ����� �� ������ �� �������

Public Const str_GasFracTubHmes_curve = "GasFracTubHmes_curve" ' ��������� ���������� ���� � ������ � ���
Public Const str_GasFracCasHmes_curve = "GasFracCasHmes_curve" ' ��������� ���������� ���� � ������ �� �������

Public Const str_HlHmes_curve = "HlHmes_curve"         ' Liquid holdup (���������� ��������) � ������ ����� ���
Public Const str_HLtubHmes_curve = "HLtubHmes_curve"         ' Liquid holdup (���������� ��������) � ������ ����� ���
Public Const str_HLcasHmes_curve = "HLcasHmes_curve"         ' Liquid holdup (���������� ��������) � ������ �� �������

 Public Const str_muoTubCurve = "muoTubCurve" '
 Public Const str_muwTubCurve = "muwTubCurve" '
 Public Const str_mugTubCurve = "mugTubCurve" '
 Public Const str_mumixTubCurve = "mumixTubCurve" '
 
 Public Const str_rhooTubCurve = "rhooTubCurve" '
 Public Const str_rhowTubCurve = "rhowTubCurve" '
 Public Const str_rholTubCurve = "rholTubCurve" '
 Public Const str_rhogTubCurve = "rhogTubCurve" '
 Public Const str_rhomixTubCurve = "rhomixTubCurve" '
 
 Public Const str_qoTubCurve = "qoTubCurve" '
 Public Const str_qwTubCurve = "qwTubCurve" '
 Public Const str_qgTubCurve = "qgTubCurve" '
 
 Public Const str_moTubCurve = "moTubCurve" '
 Public Const str_mwTubCurve = "mwTubCurve" '
 Public Const str_mgTubCurve = "mgTubCurve" '
 
 Public Const str_vlTubCurve = "vlTubCurve" '
 Public Const str_vgTubCurve = "vgTubCurve" '
 
 Public Const str_muoCasCurve = "muoCasCurve" '
 Public Const str_muwCasCurve = "muwCasCurve" '
 Public Const str_mugCasCurve = "mugCasCurve" '
 Public Const str_mumixCasCurve = "mumixCasCurve" '
 
 Public Const str_rhooCasCurve = "rhooCasCurve" '
 Public Const str_rhowCasCurve = "rhowCasCurve" '
 Public Const str_rholCasCurve = "rholCasCurve" '
 Public Const str_rhogCasCurve = "rhogCasCurve" '
 Public Const str_rhomixCasCurve = "rhomixCasCurve" '
 
 Public Const str_qoCasCurve = "qoCasCurve" '
 Public Const str_qwCasCurve = "qwCasCurve" 'a's
 Public Const str_qgCasCurve = "qgCasCurve" '
 
 Public Const str_moCasCurve = "moCasCurve" '
 Public Const str_mwCasCurve = "mwCasCurve" '
 Public Const str_mgCasCurve = "mgCasCurve" '
 
 Public Const str_vlCasCurve = "vlCasCurve" '
 Public Const str_vgCasCurve = "vgCasCurve" '





'
'
'


Public Const const_T_K_min = 273         ' ���� ���� ������ �� �������?
Public Const const_TMcCain_K_min = 289         ' ���� ���� ������ �� �������?
Public Const const_T_K_max = 573         ' ���� ���� ������ �� �������?
Public Const const_T_K_zero_C = 273
Public Const const_T_C_min = const_T_K_min - const_T_K_zero_C
Public Const const_T_C_max = const_T_K_max - const_T_K_min
Public Const aaaa = 1
Public Const const_Pi As Double = 3.14159265358979
Public Const const_Tsc_C = 20
Public Const const_Tsc_K As Double = const_Tsc_C + const_T_K_zero_C ' ����������� ����������� ��������, �
Public Const const_Psc_atma As Double = 1
'Universal gas constant
Public Const const_r As Double = 8.31
Public Const const_g = 9.81
Public Const const_rho_air = 1.2217
Public Const const_gamma_w = 1
Public Const const_rho_ref = 1000

Public Const const_ZNLF_rate = 0.01

'Air molar mass
Public Const const_m_a_kgmol As Double = 0.029


Public Const STOR_SIZE As Integer = 50

' ������������� ��������� �� ������� � �������� (�����) - ������� �������� ��� ��������� ����������  �/�
Public Const const_sigma_w_Nm = 0.01
Public Const const_sigma_oil_Nm = 0.025

Public Const const_mu_w = 0.36
Public Const const_mu_g = 0.0122
Public Const const_mu_o = 0.7

Public Const const_gamma_gas_default = 0.6
Public Const const_gamma_wat_default = 1
Public Const const_gamma_oil_default = 0.86

Public Const const_Rsb_default = 100
Public Const const_Bob_default = 1.2
Public Const const_Tres_default = 90

Public Const const_Roughness_default = 0.0001


' ����� �������� ��� ����� ����������� �������� ����������
Public Const const_gamma_gas_min = 0.5   ' ��������� ������ 0.59 - ������������ ����� ����� �� �����
Public Const const_gamma_gas_max = 2     ' ��������� �������������� ����� (������) ����� �������� �� 4, �� �� ������� ��� � ����� ����� �� ����� ������ ����
Public Const const_gamma_water_min = 0.9 ' ��������� ���� �� 0.9 �� 1.5
Public Const const_gamma_water_max = 1.5
Public Const const_gamma_oil_min = 0.5   ' ��������� �����
Public Const const_gamma_oil_max = 1.5

Public Const const_P_MPa_min = 0
Public Const const_P_MPa_max = 50
Public Const const_Salinity_ppm_min = 0
Public Const const_Salinity_ppm_max = 265000  ' equal to weigh percent salinity 26.5%.  ����������� �� �������� ������������ ���������� ��������
Public Const const_Rsb_m3m3_min = 0
Public Const const_Rsb_m3m3_max = 100000 ' Rsb more that 100 000 not allowed
Public Const const_Ppr_min = 0.002
Public Const const_Ppr_max = 30
Public Const const_Tpr_min = 0.7
Public Const const_Tpr_max = 3
Public Const const_Z_min = 0.05
Public Const const_Z_max = 5
Public Const const_TGeoGrad_C100m = 3   ' ������������� �������� � �������� �� 100 �
Public Const const_Heps_m = 0.001       ' ������ ��� ������������ ������ �����, - �������� ������������� ����� ���������� ����
Public Const const_ESP_length = 1      ' ����� ���� �� ���������

' ����� �������� ��� �������� ������ ��������� � ��������� ������������
Public Const const_convert_atma_Pa = 101325
Public Const const_convert_Pa_atma = 1 / const_convert_atma_Pa
' ��������� ��� ��������� ������ �������� �� atma � �� (� ����� atma 101325 Pa)
Public Const const_convert_kgfcm2_Pa = 98066.5
' ��������� ��� ��������� ������ �������� �� atma � MPa
Public Const const_convert_m3day_bbl = 6.289810569
Public Const const_convert_gpm_m3day = 5.450992992     ' (US) gallon per minute
Public Const const_convert_m3day_gpm = 1 / const_convert_gpm_m3day
Public Const const_convert_m3m3_scfbbl = 5.614583544
Public Const const_convert_scfbbl_m3m3 = 1 / const_convert_m3m3_scfbbl

Public Const const_convert_bbl_m3day = 1 / const_convert_m3day_bbl
Public Const const_conver_day_sec = 86400   ' updated for test  rnt21
Public Const const_convert_hr_sec = 3600
Public Const const_convert_m3day_m3sec = 1 / const_conver_day_sec

' ��������� ��� ��������� ������ �������� �������� �� �3/��� � �������
Public Const const_conver_sec_day = 1 / const_conver_day_sec
Public Const const_convert_atma_psi = 14.7
Public Const const_convert_ft_m = 0.3048
Public Const const_convert_m_ft = 1 / const_convert_ft_m
Public Const const_convert_m_mm = 1000
Public Const const_convert_mm_m = 1 / const_convert_m_mm
Public Const const_convert_cP_Pasec = 1 / 1000

Public Const const_convert_HP_W = 745.69987  ' 735.49875  ' ����������� ��������� ����. ������� ������, ��� ������ ����� ����������� ����������� ��������� ���� (1.013 �����������)
Public Const const_convert_W_HP = 1 / const_convert_HP_W
Public Const const_convert_Nm_dynescm = 1000
Public Const const_convert_lbmft3_kgm3 = 16.01846
Public Const const_convert_kgm3_lbmft3 = 1 / const_convert_lbmft3_kgm3
' pressure gradient conversion factor
Public Const const_convert_psift_atmm = 1 / const_convert_atma_psi / const_convert_ft_m
Public Const const_convert_MPa_atma = 1000000 / const_convert_atma_Pa  ' 9.8692
' ��������� ��� ��������� ������ �������� �� ��� � atma
Public Const const_convert_atma_MPa = 1 / const_convert_MPa_atma ' 0.101325
' ��������� ��� ��������� ������ �������� �� atma � MPa
Public Const const_P_atma_min = const_P_MPa_min * const_convert_MPa_atma
Public Const const_P_atma_max = const_P_MPa_max * const_convert_MPa_atma
Public Const MAXIT = 100
' ��������� ��� ������� ������������ ������
Public Const const_MaxSegmLen = 100
Public Const const_n_n = 20
Public Const const_MaxdP = 10
Public Const const_minPpipe_atma = 0.9

Public Const const_well_P_tolerance = 0.05     ' ���������� ����������� ��� ������� ��������� �������� � ��������
Public Const const_P_difference = 0.0001       ' ���������� ����������� ��� ��������� (� ��������) ��������

Public Const ang_max = 5
Public Const const_OutputCurveNumPoints = 50
Public Const DEFAULT_PAN_STEP = 15


Public Const kErrWellConstruction = 513 + vbObjectError
Public Const kErrPVTinput = 514 + vbObjectError
Public Const kErrNodalCalc = 515 + vbObjectError
Public Const kErrInitCalc = 516 + vbObjectError
Public Const kErrESPbase = 517 + vbObjectError

Public Const kErrArraySize = 701 + vbObjectError
Public Const kErrBuildCurve = 702 + vbObjectError
Public Const kErrCurveStablePointIndex = 703 + vbObjectError
Public Const kErrCurvePointIndex = 704 + vbObjectError
Public Const kErrReadDataFromWorksheet = 705 + vbObjectError
Public Const kErrWriteDataFromWorksheet = 706 + vbObjectError

Public Const sDELIM As String = vbLf & vbNewLine

Public Const MinCountPoints_Calc_Pwf_PanHd_atma = 5
' hydraulic correlation (0 BB, 1 Ansari, 2 Unified)
' todo ��������, 0, 1, 2 �� BeggsBriilCor, AnsariCor, UnifiedCor
Public Enum H_CORRELATION
    BeggsBriilCor = 0
    AnsariCor = 1
    UnifiedCor = 2
    Gray = 3
    HagedornBrown = 4
    SakharovMokhov = 5
End Enum
' Yukos_standard, ������ PVT_CORRELATION
' todo �������� ���������� � �������������
Public Enum PVT_CORRELATION
    StandingBased = 0 '
    McCainBased = 1 '
    StraigthLine = 2
End Enum
' 0 - ������� �� �������� �������� � �����������, �� ���������.��������,
' �� ��������� �������� ������� ������ ������ �� ��������� ��������
' todo �������� ���������� � �������������
Public Enum ALONG_FLOW
    AgainstStream = 0   ' ������� ������ ������
    AlongStream = 1     ' ������� �� ������
    AlongStreamZNLF = 2 ' ������� �� ������ ����� ����������� ��������
End Enum
Public Enum FLOW_DIRECTION
    FlowAlongCoord = 1
    FlowAgainstCoord = -1
End Enum
Public Enum FLOW_DIRECTION_AT_START
    FlowInPipe = -1
    FlowOutPipe = 1
End Enum
Public Enum TEMP_DIRECTION
    TempDownAlongCoord = -1
    TempRaiseAlongCoord = 1
End Enum

Public Enum TEMP_CALC_METHOD
    StartEndTemp = 0
    GeoGradTemp = 1
    AmbientTemp = 2
'    AmbientTempSimple = 3
End Enum

' ��������� ���������� �������� ��� ����� � ��������
Public Type PARAMCALC
 correlation As H_CORRELATION
 FlowDirection As FLOW_DIRECTION         '   ����������, ��� ����� ��������� ������������ ��������� �� �����
 'FlowAtStart As FLOW_DIRECTION_AT_START  '   ����������, ��� ����� ���������� �������� ����������� ������ ���������� � ����� (�����)
 'TempDirection As TEMP_DIRECTION         '   ����  false  �� ������ � ��������� ����� ����� �������� �� ����� (�����)
 tempMethod As TEMP_CALC_METHOD          ' ����� ������� ����������� � �����
End Type

' ��� ��� �������� ��������������� �������  (��� ��������)
' ������������ ��� �������� ������. ������������ ������ ���� ������� �����������
Public Type PTtype
    P_atma As Double
    T_C As Double
End Type

' ��� ��� �������� ���������� ������ �������������� ���������
Public Type MOTOR_DATA
    U_lin_V As Double       ' ���������� �������� (����� ������)
    I_lin_A As Double       ' ��� �������� (� �����)
    U_phase_V As Double     ' ���������� ������ (����� ����� � �����)
    I_phase_A As Double     ' ��� ������ (� �������)
    f_Hz As Double          ' ������� ���������� (�������� ����)
    Eff_d As Double         ' ���
    Cosphi As Double       ' ����������� ��������
    S_d As Double           ' ���������������
    Pshaft_kW As Double     ' �������� �� ���� ������������
    Pelectr_kW As Double    ' �������� ���������� �������������
    Mshaft_Nm As Double     ' ������ �� ���� - ������������
    load_d As Double        ' �������� ���������
End Type

Public Type PIPE_FLOW_PARAMS
  md_m As Double         ' pipe measured depth (from start - top)
  vd_m As Double         ' pipe vertical depth from start - top
  P_atma As Double        ' pipe pressure at measured depth
  T_C As Double          ' pipe temp at measured depth
  
  dp_dl As Double
  dt_dl As Double
  
  dpdl_g_atmm As Double  ' gravity gradient at measured depth
  dpdl_f_atmm As Double  ' friction gradient at measured depth
  dpdl_a_atmm As Double  ' acceleration gradient at measured depth
  v_sl_msec As Double        ' superficial liquid velosity
  v_sg_msec As Double        ' superficial gas velosity
  h_l_d As Double            ' liquid hold up
  fpat As Double             ' flow pattern code
  thete_deg As Double
  roughness_m As Double
  
  rs_m3m3 As Double         ' ������������ ��� � ����� � ������
  gasfrac As Double         ' ��������� ���������� ���� � ������
  
  Muo_cP As Double          ' �������� ����� � ������
  Muw_cP As Double          ' �������� ���� � ������
  Mug_cP As Double          ' �������� ���� � ������
  MuMix_cP As Double        ' �������� ����� � ������
  
  Rhoo_kgm3 As Double       ' ��������� �����
  Rhow_kgm3 As Double       ' ��������� ����
  rhol_kgm3 As Double       ' ��������� ��������
  Rhog_kgm3 As Double       ' ��������� ����
  rhomix_kgm3 As Double     ' ��������� ����� � ������
  
  Qo_m3day As Double        ' ������ ����� � ������� ��������
  Qw_m3day As Double        ' ������ ���� � ������� ��������
  Qg_m3day As Double        ' ������ ���� � ������� ��������
  
  mo_kgsec As Double        ' �������� ������ ����� � ������� ��������
  mw_kgsec As Double        ' �������� ������ ���� � ������� ��������
  mg_kgsec As Double        ' �������� ������ ���� � ������� ��������
  
  vl_msec As Double         ' �������� �������� ��������
  vg_msec As Double         ' �������� ���� ��������
  
End Type


Public Enum VAL_FROM_SHEET
    VAL_DOUBLE = 0
    VAL_STRING = 1
    VAL_CURVE = 2
    VAL_BOOLEAN = 3
End Enum


Public Enum GAS_INTO_SOLUTION
    GasGoesIntoSolution = 1
    GasnotGoesIntoSolution = 0
End Enum

Public Type CONTEXT_DATA
    Description As String   ' �������� �������� ��������� ����������� ��������
    method As String
    Class As String
    description_out As String ' ������� ��������
End Type

' ������ ������� ��������� - ����� ���� ������������ ��� ������ ������������� ��������
Public Enum MESSAGE_LEVEL
    mlCritical = 0
    mlModerate = 1
    mlStandard = 2
    mlLow = 3
End Enum

' ��� ����������� ��� �������� ��������
Public Enum MESSAGE_TYPE
    msgDataQualityReport = 0
    msgError = 1
    msgWarning = 2
    msgLog = 3
End Enum

Public Enum CALC_RESULTS
    noCurves = 0
    mainCurves = 1
    allCurves = 2
End Enum

Public Function ParamCalcFromTop(Optional hcor As H_CORRELATION = AnsariCor, Optional TempMeth As TEMP_CALC_METHOD = StartEndTemp) As PARAMCALC
    ParamCalcFromTop.correlation = hcor
    ParamCalcFromTop.FlowDirection = FlowAlongCoord
    ParamCalcFromTop.tempMethod = TempMeth
End Function

Public Function ParamCalcFromBottom(Optional hcor As H_CORRELATION = AnsariCor, Optional TempMeth As TEMP_CALC_METHOD = StartEndTemp) As PARAMCALC
    ParamCalcFromBottom.correlation = hcor
    ParamCalcFromBottom.FlowDirection = FlowAgainstCoord
    ParamCalcFromBottom.tempMethod = TempMeth
End Function

Public Function SumPT(PT1 As PTtype, PT2 As PTtype) As PTtype
    SumPT.P_atma = PT1.P_atma + PT2.P_atma
    SumPT.T_C = PT1.T_C + PT2.T_C
End Function
Public Function SubtractPT(PT1 As PTtype, PT2 As PTtype) As PTtype
    SubtractPT.P_atma = PT1.P_atma - PT2.P_atma
    SubtractPT.T_C = PT1.T_C - PT2.T_C
End Function

Public Function SetPT(ByVal P As Double, ByVal t As Double) As PTtype
    SetPT.P_atma = P
    SetPT.T_C = t
End Function


Public Sub RangeChange(rng1, sheet, target)
  ' ��������� ��� ����������� ����������� ��������� �����
 On Error GoTo er1:
  If (target.Address = Range(rng1).Address) And (Sheets(sheet).Range(rng1).Value2 <> target.Value2) Then
    Sheets(sheet).Range(rng1).Value2 = target.Value2
  End If
  Exit Sub
er1:
End Sub

