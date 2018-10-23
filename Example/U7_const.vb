﻿Module U7_const

    '=======================================================================================
    'Unifloc7.3  Testudines                                           khabibullinra@gmail.com
    'Библиотека расчетных модулей по нефтяному инжинирингу
    '2000 - 2018 г
    '
    '=======================================================================================
    ' constant definition module
    '

    '
    ' константы для наименований стандартных переменных используемые при загрузке
    '

    Public Const const_default_wrong_read_num As Double = -999999

    Public Const const_name_Qliq_m3day As String = "Qliq_m3day"
    Public Const const_name_wc_perc As String = "wc_perc"
    Public Const const_name_Pcas_atma As String = "Pcas_atma"
    Public Const const_name_Hdyn_m As String = "Hdyn_m"
    Public Const const_name_Pline_atma As String = "Pline_atma"
    Public Const const_name_Pbuf_atma As String = "Pbuf_atma"
    Public Const const_name_Pwf_atma As String = "Pwf_atma"

    Public Const const_name_gamma_oil As String = "gamma_oil"
    Public Const const_name_gamma_water As String = "gamma_water"
    Public Const const_name_gamma_gas As String = "gamma_gas"
    Public Const const_name_Rp_m3m3 As String = "Rp_m3m3"
    Public Const const_name_Rsb_m3m3 As String = "Rsb_m3m3"
    Public Const const_name_Pb_atma As String = "Pb_atma"
    Public Const const_name_HResMes_m As String = "HResMes_m"
    Public Const const_name_HPumpMes_m As String = "HPumpMes_m"
    Public Const const_name_Dchoke_mm As String = "Dchoke_mm, mm"
    Public Const const_name_Шероховатость_м As String = "Шероховатость_м, м"
    Public Const const_name_ESP_Qliq_m3day As String = "ESP_Qliq_m3day"
    Public Const const_name_ESP_NumStages As String = "ESP_NumStages"
    Public Const const_name_ESP_Freq_Hz As String = "ESP_Freq_Hz"
    Public Const const_name_ESP_Pintake_atma As String = "ESP_Pintake_atma"
    Public Const const_name_Pres_atma As String = "Pres_atma"
    Public Const const_name_PI_m3dayatm As String = "PI_m3dayatm"
    Public Const const_name_ESP_Tintake_C As String = "ESP_Tintake_C"
    Public Const const_name_Tres_C As String = "Tres_C"


    Public Const const_name_HmesCurve As String = "HmesCurve"
    Public Const const_name_DcasCurve As String = "DcasCurve"
    Public Const const_name_DtubCurve As String = "DtubCurve"
    Public Const const_name_TAmbCurve As String = "TAmbCurve"
    '
    '
    '

    'Public Const a = 1

    Public Const str_VLPcurve = "VLPcurve"                      ' кривая оттока -  зависимость забойного давления от дебита жидкости
    Public Const str_HvertCurve = "HvertCurve"                  ' кривая траектории (кривизны) скважины

    Public Const str_DcasCurve = "DcasCurve"                    ' кривая изменения диаметра эксплуатационной колонны
    Public Const str_DtubCurve = "DtubCurve"                    ' кривая изменения диаметра НКТ

    Public Const str_RoughnessCasCurve = "RoughnessCasCurve"    ' кривая изменения шероховатости по трубе эксплуатационной колонны
    Public Const str_RoughnessTubCurve = "RoughnessTubCurve"    ' кривая изменения шероховатости по трубе НКТ

    Public Const str_Hd_Depend_Pwf = "Hd_Depend_Pwf"            ' кривая - зависимость динамического уровня от забойного давления, при заданном давлении в затрубе
    Public Const str_Pan_Depend_Pwf = "Pan_Depend_Pwf"          ' кривая - зависимость затрубного давления от забойного давления

    ' зависимости лин давления и буферного давления от дин уровня логично показывать на одном графике
    Public Const str_Plin_Depend_Pwf = "Plin_Depend_Pwf"        ' кривая - зависимость линейного давления от дин уровня
    Public Const str_Pbuf_Pwf_curve = "Pbuf_Pwf_curve"          ' зависимость буферного давления от дин уровня

    Public Const str_KsepNatQl_curve = "KsepNatQl_curve"         ' зависимость коэффициента сепарации от дебита
    Public Const str_KsepNatRp_curve = "KsepNatRp_curve"         ' зависимость коэффициента сепарации от газового фактора

    Public Const str_KsepTotalQl_curve = "KsepTotalQl_curve"     ' кривая общего коэффициента сепарации от дебита
    Public Const str_KsepTotalRp_curve = "KsepTotalRp_curve"     ' кривая общего коэффициента сепарации от газового фактора
    Public Const str_KsepGasSepQl_curve = "KsepGasSepQl_curve"   ' кривая коэффициента сепарации газосепаратора от дебита
    Public Const str_KsepGasSepRp_curve = "KsepGasSepQl_curve"   ' кривая коэффициента сепарации газосепаратора от газового фактора

    Public Const str_PdisKdegr_curve = "PdisKdegr_curve"         ' кривая зависимости давления на устье от деградации напора УЭЦН


    Public Const str_TambHmes_curve = "TambHmes_curve"           ' профиль температуры окружающего простраства от измеренный координаты

    Public Const str_PtubHmes_curve = "PtubHmes_curve"           ' профиль давления по стволу скважины по ниже НКТ и по НКТ до устья
    Public Const str_TtubHmes_curve = "TtubHmes_curve"           ' профиль температуры по стволу скважины по НКТ

    Public Const str_PcasHmes_curve = "PcasHmes_curve"           ' профиль давления по стволу скважины по ниже НКТ и по затрубу до устья
    Public Const str_TcasHmes_curve = "TcasHmes_curve"           ' профиль температуры по стволу скважины ниже насоса и выше насоса по затрубу

    Public Const str_RstubHmes_curve = "RstubHmes_curve"         ' профиль остаточного содержания газа в нефти по потоку в НКТ
    Public Const str_RscasHmes_curve = "RscasHmes_curve"         ' профиль остаточного содержания газа в нефти по потоку по затрубу

    Public Const str_GasFracTubHmes_curve = "GasFracTubHmes_curve" ' расходное содержание газа в потоке в НКТ
    Public Const str_GasFracCasHmes_curve = "GasFracCasHmes_curve" ' расходное содержание газа в потоке по затрубу

    Public Const str_HlHmes_curve = "HlHmes_curve"         ' Liquid holdup (содержание жидкости) в потоке через НКТ
    Public Const str_HLtubHmes_curve = "HLtubHmes_curve"         ' Liquid holdup (содержание жидкости) в потоке через НКТ
    Public Const str_HLcasHmes_curve = "HLcasHmes_curve"         ' Liquid holdup (содержание жидкости) в потоке по затрубу

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


    Public Const const_T_K_min = 273         ' ниже нуля ничего не считаем?
    Public Const const_TMcCain_K_min = 289         ' ниже нуля ничего не считаем?
    Public Const const_T_K_max = 573         ' выше тоже ничего не считаем?
    Public Const const_T_K_zero_C = 273
    Public Const const_T_C_min = const_T_K_min - const_T_K_zero_C
    Public Const const_T_C_max = const_T_K_max - const_T_K_min
    Public Const aaaa = 1
    Public Const const_Pi As Double = 3.14159265358979
    Public Const const_Tsc_C = 20
    Public Const const_Tsc_K As Double = const_Tsc_C + const_T_K_zero_C ' температура стандартных условиях, К
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

    ' поверхностное натяжение на границе с воздухом (газом) - типовые значения для дефолтных параметров  Н/м
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


    ' набор констант для общих ограничений значений переменных
    Public Const const_gamma_gas_min = 0.5   ' плотность метана 0.59 - предпологаем легче газов не будет
    Public Const const_gamma_gas_max = 2     ' плотность углеводородных газов (гексан) может доходить до 4, но мы считаем что в смеси таких не много должно быть
    Public Const const_gamma_water_min = 0.9 ' плотность воды от 0.9 до 1.5
    Public Const const_gamma_water_max = 1.5
    Public Const const_gamma_oil_min = 0.5   ' плотность нефти
    Public Const const_gamma_oil_max = 1.5

    Public Const const_P_MPa_min = 0
    Public Const const_P_MPa_max = 50
    Public Const const_Salinity_ppm_min = 0
    Public Const const_Salinity_ppm_max = 265000  ' equal to weigh percent salinity 26.5%.  Ограничение по границам применимости корреляций МакКейна
    Public Const const_Rsb_m3m3_min = 0
    Public Const const_Rsb_m3m3_max = 100000 ' Rsb more that 100 000 not allowed
    Public Const const_Ppr_min = 0.002
    Public Const const_Ppr_max = 30
    Public Const const_Tpr_min = 0.7
    Public Const const_Tpr_max = 3
    Public Const const_Z_min = 0.05
    Public Const const_Z_max = 5
    Public Const const_TGeoGrad_C100m = 3   ' геотермальный градиент в градусах на 100 м
    Public Const const_Heps_m = 0.001       ' дельта для корретировки кривой трубы, - примерно соответствует длине сочленения труб
    Public Const const_ESP_length = 1      ' длина УЭЦН по умолчанию

    ' набор констант для перевода единиц измерений в различных размерностях
    Public Const const_convert_atma_Pa = 101325
    Public Const const_convert_Pa_atma = 1 / const_convert_atma_Pa
    ' константа для конверсии единиц давления из atma в Па (в одной atma 101325 Pa)
    Public Const const_convert_kgfcm2_Pa = 98066.5
    ' константа для конверсии единиц давления из atma в MPa
    Public Const const_convert_m3day_bbl = 6.289810569
    Public Const const_convert_gpm_m3day = 5.450992992     ' (US) gallon per minute
    Public Const const_convert_m3day_gpm = 1 / const_convert_gpm_m3day
    Public Const const_convert_m3m3_scfbbl = 5.614583544
    Public Const const_convert_scfbbl_m3m3 = 1 / const_convert_m3m3_scfbbl

    Public Const const_convert_bbl_m3day = 1 / const_convert_m3day_bbl
    Public Const const_conver_day_sec = 86400   ' updated for test  rnt21
    Public Const const_convert_hr_sec = 3600
    Public Const const_convert_m3day_m3sec = 1 / const_conver_day_sec

    ' константа для конверсии единиц объемных расходов из м3/сут в баррели
    Public Const const_conver_sec_day = 1 / const_conver_day_sec
    Public Const const_convert_atma_psi = 14.7
    Public Const const_convert_ft_m = 0.3048
    Public Const const_convert_m_ft = 1 / const_convert_ft_m
    Public Const const_convert_m_mm = 1000
    Public Const const_convert_mm_m = 1 / const_convert_m_mm
    Public Const const_convert_cP_Pasec = 1 / 1000

    Public Const const_convert_HP_W = 745.69987  ' 735.49875  ' метрическая лошадиная сила. следует учесть, что иногда может применяться механическя лошадиная сила (1.013 метрической)
    Public Const const_convert_W_HP = 1 / const_convert_HP_W
    Public Const const_convert_Nm_dynescm = 1000
    Public Const const_convert_lbmft3_kgm3 = 16.01846
    Public Const const_convert_kgm3_lbmft3 = 1 / const_convert_lbmft3_kgm3
    ' pressure gradient conversion factor
    Public Const const_convert_psift_atmm = 1 / const_convert_atma_psi / const_convert_ft_m
    Public Const const_convert_MPa_atma = 1000000 / const_convert_atma_Pa  ' 9.8692
    ' константа для конверсии единиц давления из Мпа в atma
    Public Const const_convert_atma_MPa = 1 / const_convert_MPa_atma ' 0.101325
    ' константа для конверсии единиц давления из atma в MPa
    Public Const const_P_atma_min = const_P_MPa_min * const_convert_MPa_atma
    Public Const const_P_atma_max = const_P_MPa_max * const_convert_MPa_atma
    Public Const MAXIT = 100
    ' константы для расчета многофазного потока
    Public Const const_MaxSegmLen = 100
    Public Const const_n_n = 20
    Public Const const_MaxdP = 10
    Public Const const_minPpipe_atma = 0.9

    Public Const const_well_P_tolerance = 0.05     ' допустимая погрешность при расчете забойного давления в скважине
    Public Const const_P_difference = 0.0001       ' допустимая погрешность при сравнении (в основном) давлений

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
    ' todo заменить, 0, 1, 2 на BeggsBriilCor, AnsariCor, UnifiedCor
    Public Enum H_CORRELATION
        BeggsBriilCor = 0
        AnsariCor = 1
        UnifiedCor = 2
        Gray = 3
        HagedornBrown = 4
        SakharovMokhov = 5
    End Enum
    ' Yukos_standard, иногда PVT_CORRELATION
    ' todo заменить объявление и использование
    Public Enum PVT_CORRELATION
        StandingBased = 0 '
        McCainBased = 1 '
        StraigthLine = 2
    End Enum
    ' 0 - считаем от меньшего давления к наибольшему, по умолчанию.Например,
    ' от устьевого давления считаем против потока до забойного давления
    ' todo заменить объявление и использование
    Public Enum ALONG_FLOW
        AgainstStream = 0   ' подсчет против потока
        AlongStream = 1     ' подсчет по потоку
        AlongStreamZNLF = 2 ' подсчет по потоку через неподвижную жидкость
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

    ' настройки проведения расчетов для трубы и скважины
    Public Structure PARAMCALC
        Dim correlation As H_CORRELATION
        Dim FlowDirection As FLOW_DIRECTION         '   показывает, как поток направлен относительно координат по длине
        'Dim FlowAtStart As FLOW_DIRECTION_AT_START  '   показывает, что точка начального давления соответвует потоку втекающему в трубу (забой)
        'Dim TempDirection As TEMP_DIRECTION         '   если  false  то значит в начальной точке поток вытекает из трубы (устье)
        Dim tempMethod As TEMP_CALC_METHOD          ' метод расчета температуры в трубе
    End Structure

    ' тип для описания термобарических условий  (для расчетов)
    ' используется для передачи данных. Поддерживает только один вариант размерности
    Public Structure PTtype
        Dim P_atma As Double
        Dim T_C As Double
    End Structure

    ' тип для описания параметров работы электрического двигателя
    Public Structure MOTOR_DATA
        Dim U_lin_V As Double       ' напряжение линейное (между фазами)
        Dim I_lin_A As Double       ' ток линейный (в линии)
        Dim U_phase_V As Double     ' напряжение фазное (между фазой и нулем)
        Dim I_phase_A As Double     ' ток фазный (в обмотке)
        Dim f_Hz As Double          ' частота синхронная (вращение поля)
        Dim Eff_d As Double         ' КПД
        Dim Cosphi As Double       ' коэффициент мощности
        Dim S_d As Double           ' проскользывание
        Dim Pshaft_kW As Double     ' мощность на валу механическая
        Dim Pelectr_kW As Double    ' мощность подводимая электрическая
        Dim Mshaft_Nm As Double     ' момент на валу - механический
        Dim load_d As Double        ' загрузка двигателя
    End Structure

    Public Structure PIPE_FLOW_PARAMS
        Dim md_m As Double         ' pipe measured depth (from start - top)
        Dim vd_m As Double         ' pipe vertical depth from start - top
        Dim P_atma As Double        ' pipe pressure at measured depth
        Dim T_C As Double          ' pipe temp at measured depth

        Dim dp_dl As Double
        Dim dt_dl As Double

        Dim dpdl_g_atmm As Double  ' gravity gradient at measured depth
        Dim dpdl_f_atmm As Double  ' friction gradient at measured depth
        Dim dpdl_a_atmm As Double  ' acceleration gradient at measured depth
        Dim v_sl_msec As Double        ' superficial liquid velosity
        Dim v_sg_msec As Double        ' superficial gas velosity
        Dim h_l_d As Double            ' liquid hold up
        Dim fpat As Double             ' flow pattern code
        Dim thete_deg As Double
        Dim roughness_m As Double

        Dim rs_m3m3 As Double         ' растворенный газ в нефти в потоке
        Dim gasfrac As Double         ' расходное содержание газа в потоке

        Dim Muo_cP As Double          ' вязкость нефть в потоке
        Dim Muw_cP As Double          ' вязкость воды в потоке
        Dim Mug_cP As Double          ' вязкость газа в потоке
        Dim MuMix_cP As Double        ' вязкость смеси в потоке

        Dim Rhoo_kgm3 As Double       ' плотность нефти
        Dim Rhow_kgm3 As Double       ' плотность воды
        Dim rhol_kgm3 As Double       ' плотность жидкости
        Dim Rhog_kgm3 As Double       ' плотность газа
        Dim rhomix_kgm3 As Double     ' плотность смеси в потоке

        Dim Qo_m3day As Double        ' расход нефти в рабочих условиях
        Dim Qw_m3day As Double        ' расход воды в рабочих условиях
        Dim Qg_m3day As Double        ' расход газа в рабочих условиях

        Dim mo_kgsec As Double        ' массовый расход нефти в рабочих условиях
        Dim mw_kgsec As Double        ' массовый расход воды в рабочих условиях
        Dim mg_kgsec As Double        ' массовый расход газа в рабочих условиях

        Dim vl_msec As Double         ' скорость жидкости реальная
        Dim vg_msec As Double         ' скорость газа реальная

    End Structure


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

    Public Structure CONTEXT_DATA
        Dim Description As String   ' описание текущего контекста выполняемых действий
        Dim method As String
        Dim Class As String
    Dim description_out As String ' внешний контекст
End Structure

    ' уровни влияния сообщения - могут быть использованы для оценки достоверности расчетов
    Public Enum MESSAGE_LEVEL
            mlCritical = 0
            mlModerate = 1
            mlStandard = 2
            mlLow = 3
        End Enum

        ' тип уведомления при контроле расчетов
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

    Public Function ParamCalcFromTop(Optional hcor As H_CORRELATION = H_CORRELATION.AnsariCor, Optional TempMeth As TEMP_CALC_METHOD = TEMP_CALC_METHOD.StartEndTemp) As PARAMCALC
        ParamCalcFromTop.correlation = hcor
        ParamCalcFromTop.FlowDirection = FLOW_DIRECTION.FlowAlongCoord
        ParamCalcFromTop.tempMethod = TempMeth
    End Function

    Public Function ParamCalcFromBottom(Optional hcor As H_CORRELATION = H_CORRELATION.AnsariCor, Optional TempMeth As TEMP_CALC_METHOD = TEMP_CALC_METHOD.StartEndTemp) As PARAMCALC
        ParamCalcFromBottom.correlation = hcor
        ParamCalcFromBottom.FlowDirection = FLOW_DIRECTION.FlowAgainstCoord
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

    '        Public Sub RangeChange(rng1, sheet, target)
    '            ' процедура для обеспечения синхронного изменения ячеек
    '            On Error GoTo er1
    '            If (target.Address = Range(rng1).Address) And (Sheets(sheet).Range(rng1).Value2 <> target.Value2) Then
    '                Sheets(sheet).Range(rng1).Value2 = target.Value2
    '            End If
    '            Exit Sub
    'er1:
    '        End Sub


End Module
