Public Class CPVT



'=======================================================================================
'Unifloc7.3  Testudines                                           khabibullinra@gmail.com
'Библиотека расчетных модулей по нефтяному инжинирингу
'2000 - 2018 г
'
'=======================================================================================



 ' рабочие условия для которых проводится расчет
 Private p_PTcalc As PTtype                         ' термобарические условия при которых был проведен расчет
 ' базовые параметры флюида
 Private p_gamma_o As Double                        ' плотность нефти удельная
 Private p_gamma_g As Double                        ' плотность газа удельная
 Private p_gamma_w As Double                        ' плотность воды удельная
 Private p_Rsb_m3m3 As Double                       ' газосодержание при давлении насыщения
 ' калибровочные параметры нефти
 Private p_Pb_atma As Double                         ' давление насыщения  (калибровочное значение)
 Private p_Muob_cP As Double                       ' вязкость нефти при давлении насыщения (калибровочное значение)
 Private p_Bob_m3m3 As Double                       ' объемный коэффициент при давлении насыщения
 Private p_Tres_C As Double                         ' пластовая температура при которой заданые значения давления насыщения и объемного коэффициента
 ' расчетные параметры нефти
 Private p_Pbcalc_atma As Double                     ' расчетное значение давления насыщения по корреляции
 Private p_Rs_m3m3 As Double                        ' расчетное значение газосодержания в нефти при текущих условиях
 Private p_Bo_m3m3 As Double                        ' объемный коэффициент нефти при рабочих условиях
 Private p_Muo_cP As Double                         ' вязкость нефти при рабочих условиях
 Private p_Muo_deadoil_cP As Double                 ' вязкость дегазированной нефти
 Private p_copmressibility_o_1atm As Double         ' сжимаемость нефти
' Private p_sigma_o_Nm As Double                     ' поверхностное натяжение
' Private p_sigma_w_Nm As Double                     ' поверхностное натяжение
 Private p_ST_oilgas_dyncm As Double                ' поверхностное натяжение нефть газ
 Private p_ST_watgas_dyncm As Double                ' поверхностное натяжение вода газ
 Private p_ST_liqgas_dyncm As Double
 ' расчетные параметры газа
 Private p_Z As Double                              ' расчетное значение коэффициента сверхсжимаемости
 Private p_Bg_m3m3 As Double                        ' объемный коэффициент газа
 Private p_Mug_cP As Double                         ' вязкость газа при рабочих условиях
 ' расчетные параметры воды
 Private p_Bw_m3m3 As Double                        ' расчетное значение объемного коэффициента воды
 Private p_BwSC_m3m3 As Double
 Private p_Muw_cP As Double                         ' вязкость воды
 Private p_Salinity_ppm As Double                   ' соленость воды
 ' параметры потока
 Private p_Rp_m3m3 As Double                        ' газовый фактор добычной (приведенный к стандартным условиям)
 Private p_wc_fr As Double                          ' объемная доля воды в флюиде
 Private p_Qliq_scm3day As Double                   ' задаем для флюида также и дебиты, это упрощает дальнейшие расчеты расходов в разных условиях
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
 Private p_Qgfree_scm3day As Double                 ' дебит газа, вернее добавка для значения Qgas, в исследовательских целях
 ' набор параметров для температурных расчетов
 Private p_Co_JkgC As Double                        ' oil heat capacity   теплоемкость нефти  Дж/кг/С
 Private p_Cw_JkgC As Double                        ' water heat capacity  теплоемкость воды
 Private p_Cg_JkgC As Double                        ' теплоемкость газа gas heat capacity
 ' настройка модели
 Private p_PVT_correlation As PVT_CORRELATION       ' PVT correlation
 Public calculated As Boolean
 Private p_ZNLF As Boolean                          ' флаг показывает что флюид в режиме барботажа
 Private p_Twh_C As Double                          ' температура для которой строятся кривые
                                                    '  в режиме барботажа газ вспплывает через неподвижный столб жидкости (нефти )
                                                    '  задается дополнительно расход свободного газа через столб жидкости
                                                    '  при этом газ не выделяется из жидкости
                                                    '  корректируется обводненность - в неподвижном потоке может остаться только нефть
 Private c_RsTres_Curve As New TInterpolation       ' кривая газосодержания от давления при пластовой температуре
 Private c_RsT_Curve As New TInterpolation          ' кривая газосодержания от давления при произвольной температуре
 Private c_BoTres_Curve As New TInterpolation       ' кривая объемного коэфициента нефти от давления при пластовой температуре
 Private c_BoT_Curve As New TInterpolation          ' кривая объемного коэфициента нефтиот давления при пластовой температуре
 Private c_MuoTres_Curve As New TInterpolation      ' кривая вязкости нефти от давления при пластовой температуре
 Private c_MuoT_Curve As New TInterpolation         ' кривая вязкости нефти от давления при пластовой температуре
 Private c_MugTres_Curve As New TInterpolation      ' кривая вязкости газа от давления при пластовой температуре
 Private c_MugT_Curve As New TInterpolation         ' кривая вязкости газа от давления при пластовой температуре
    If p_ZNLF Then
    Else
   If val < 0 Then
        addLogMsg "CPVT.wc_fr: обводненность меньше нуля. Значение = " & s(val) & " заменено на 0"
   ElseIf val > 1 Then
        addLogMsg "CPVT.wc_fr: обводненность больше единицы. Значение = " & s(val) & " заменено на 1"
  
 Private Sub calc_wc()
    If (Qliq_scm3day) > 0 Then
    Else
    End If
 End Sub
   If Qgas_m3day < 0 Then Qgas_m3day = 0
   If Qgas_m3day < 0 Then Qgas_rc_m3day = 0
    If Qmix_m3day > 0 Then
    Else
    If Qmix_m3day > 0 Then
    Else
 ' коэффциент Джоуля Томсона для многофазной смеси
    Dim X As Double
    Dim wm As Double
    Dim z1 As Double, z2 As Double
    Dim dTz As Double
    Dim dZdT As Double
    Dim TZdZdT As Double
    If wm > 0 Then
    Else
    End If
 End Property
    Dim msg As String
    If Bo_m3m3 > 0 Then
    Else
        addLogMsg msg
        Err.Raise kErrPVTinput, , msg
    End If
 End Property
    Dim msg As String
    If Bw_m3m3 > 0 Then
    Else
        addLogMsg msg
        Err.Raise kErrPVTinput, , msg
    End If
 End Property
    Dim msg As String
    If Bg_m3m3 > 0 Then
    Else
        addLogMsg msg
        Err.Raise kErrPVTinput, , msg
    End If
 End Property
    If Qmix_m3day > 0 Then
    Else
    If (val > const_gamma_oil_min) And (val < const_gamma_oil_max) Then
    Else
        If (val < const_gamma_oil_min) Then p_gamma_o = const_gamma_oil_min
        If (val > const_gamma_oil_max) Then p_gamma_o = const_gamma_oil_max
        addLogMsg "CPipe.gamma_o: попытка некорректного ввода gamma_o = " & val & " установлено значение = " & p_gamma_o, msgDataQualityReport
 'установка значения плотности газа, влияет на все зависимые флюиды
    If (val > const_gamma_gas_min) And (val < const_gamma_gas_max) Then
    Else
        If (val < const_gamma_gas_min) Then p_gamma_g = const_gamma_gas_min
        If (val > const_gamma_gas_max) Then p_gamma_g = const_gamma_gas_max
        addLogMsg "CPipe.gamma_o: попытка некорректного ввода gamma_o = " & val & " установлено значение = " & p_gamma_g, msgDataQualityReport
    If (val > const_gamma_water_min) And (val < const_gamma_water_max) Then
    Else
        If (val < const_gamma_water_min) Then p_gamma_w = const_gamma_water_min
        If (val > const_gamma_water_max) Then p_gamma_w = const_gamma_water_max
        addLogMsg "CPipe.gamma_o: попытка некорректного ввода gamma_o = " & val & " установлено значение = " & p_gamma_w, msgDataQualityReport
    If (Rpval >= 0) Then
        If p_Rp_m3m3 < p_Rsb_m3m3 Then   ' проверим, что газовый фактор должен быть больше чем газосодержание
            addLogMsg "Газовый фактор при вводе больше газосодержания Rp = " & Format(p_Rp_m3m3, "####0.00") & " < Rsb = " & Format(p_Rsb_m3m3, "#0.00") & ". Газосодержание исправлено", msgDataQualityReport
    Else
        addLogMsg "Попытка установить отрицательное значение газового фактора Rp =" & Format(Rpval, "####0.00"), msgError, kErrPVTinput
    If (Rsbval >= 0) Then
        If p_Rp_m3m3 < p_Rsb_m3m3 Then   ' проверим, что газовый фактор должен быть больше чем газосодержание
            addLogMsg "газосодержания при вводе меньше газового фактора  Rp = " & Format(p_Rp_m3m3, "#0.00") & " < Rsb = " & Format(p_Rsb_m3m3, "#0.00") & ". Газосодержание исправлено", msgDataQualityReport
    Else
        addLogMsg "Попытка установить отрицательное значение газосодержания Rsb =" & Format(Rsbval, "####0.00"), msgError, kErrPVTinput
 
Public Function SetRpRsb(Rpval_m3m3 As Double, Rsbval_m3m3 As Double) As Boolean
' безопасный с точки зрения начисления штрафов способ установки произвольных значений газового фактора в системе
    If Rpval_m3m3 > 0 Then
        If Rpval_m3m3 >= Rsbval_m3m3 Then
            If Rsbval_m3m3 > 0 Then
            Else
            End If
        Else
            addLogMsg "газосодержания при вводе больше газового фактора  Rp = " & Format(p_Rp_m3m3, "#0.00") & " < Rsb = " & Format(p_Rsb_m3m3, "#0.00") & ". Газосодержание исправлено", msgDataQualityReport
        End If
    Else
        If Rpval_m3m3 <= 0 And Rsbval_m3m3 > 0 Then
        Else
        End If
    End If
End Function
    If p_Pb_atma > 0 Then       ' ноль не допустим, это значит что значение отсутствует
    Else
    If (Pbval >= 0) Then
    Else
    If (Boval >= 0) Then
    Else
    If (muoval >= 0) Then
    Else

'===================================================================================
' функции и процедуры
'
'===================================================================================
 
 Private Sub Class_Initialize()
 End Sub
 
 
 Public Sub Init(Optional gg = 0.6, Optional gamma_oil = 0.86, Optional gamma_wat = 1, _
                 Optional rsb_m3m3 = 100, Optional Pb_atma = -1, _
                 Optional Bob_m3m3 = -1, Optional PVTcorr = 0, _
                 Optional Tres_C = 90, Optional Rp_m3m3 = -1, Optional Muob_cP = -1)
        Me.SetRpRsb CDbl(Rp_m3m3), CDbl(rsb_m3m3)
 End Sub
 
 Public Function Clone()
        Dim fl As New CPVT
        Call fl.Copy(Me)
        Set Clone = fl
 End Function
 
 Public Function Copy(fl As CPVT)
        SetRpRsb fl.Rp_m3m3, fl.rsb_m3m3
 End Function

Public Sub Calc_PVT_PT(PT As PTtype)
        Call Calc_PVT(PT.P_atma, PT.T_C)
End Sub

Public Sub Calc_PVT(P_atma As Double, T_C As Double)
' расчет свойств воды нефти и газа при заданных давлении и температуре
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
On Error GoTo err1:    If PVT_CORRELATION = StandingBased Then
            ' дальше ищем калибровочные коэффициенты
            'Calculate bubble point correction factor
            If (Pb_calbr_MPa > 0) Then 'user specified
            Else ' not specified, use from correlations
            End If
            If (Bo_calbr_m3m3 > 0) Then 'Calculate oil formation volume factor correction factor
            Else ' not specified, use from correlations
            End If
            If muo_calibr_cP > 0 Then           ' рассчитаем калибровочный коэффициент для вязкости при давлении насыщения
            Else
            End If
            If P_MPa > p_bi Then 'apply correction to undersaturated oil 'undersaturated oil
            Else 'apply correction to saturated oil
            End If
    End If
    If PVT_CORRELATION = McCainBased Then
            'Calculate bubble point correction factor
            If (Pb_calbr_MPa > 0) Then 'user specifie
            Else ' not specified, use from correlations
            End If
            If (Bo_calbr_m3m3 > 0) Then 'Calculate oil formation volume factor correction factor
            Else ' not specified, use from correlations
            End If
            If P_MPa > p_bi Then 'apply correction to undersaturated oil
            Else 'apply correction to saturated oil
            End If
            If muo_calibr_cP > 0 Then           ' рассчитаем калибровочный коэффициент для вязкости при давлении насыщения
                  If (Rsb_calbr_m3m3 < 350) Then
                  Else
                  End If
            Else
            End If
            If (Rsb_calbr_m3m3 < 350) Then 'Calculate oil viscosity acoording to Standing
            Else 'Calculate according to Begs&Robinson (saturated) and Vasquez&Begs (undersaturated)
               If P_MPa > p_bi Then 'undersaturated oil
               Else 'saturated oil
               End If
            End If
    End If
    If PVT_CORRELATION = 2 Then  'Debug mode. Linear Rs and bo vs P, Pb_calbr_atma should be specified.
         If P_MPa > (p_bi) Then 'undersaturated oil
         Else 'saturate
         End If
         'if Bob_m3m3 is not specified by the user then
         'set Bob_m3m3 so, that oil density, recalculated with Rs_m3m3 would be equal to dead oil density
         If (Bo_calbr_m3m3 < 0) Then
         Else
            If P_MPa > (p_bi) Then 'undersaturated oil
            Else 'saturate
            End If
         End If
         If muo_calibr_cP > 0 Then
         Else
         End If
    End If
    
    Call calc_ST(P_atma, T_C)
'    timeStamp = Time() - timeStamp
'    timePVTtotal = timePVTtotal + timeStamp
    Exit Sub
err1:
    addLogMsg ("CPVT.Calc_PVT: ошибка какая то")
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub calc_ST(ByVal P_atma As Double, ByVal T_C As Double)
' calculate surface tension according Baker Sverdloff correlation

'Расчет коэффициента поверхностного натяжения газ-нефть
    Dim ST68 As Double, ST100 As Double
    Dim STw74 As Double, STw280 As Double
    Dim Tst As Double, Tstw As Double
    Dim STo As Double, STw As Double, ST As Double
    Dim T_F As Double
    Dim P_psia As Double, P_MPa As Double
        If T_F < 68 Then
        Else
            If T_F > 100 Then         End If
        If T_F < 74 Then
        Else
            If T_F > 280 Then         End If
End Sub

Public Function calc_Rs_m3m3(ByVal P_atma As Double, ByVal T_C As Double) As Double
'function calculates solution gas oil ratio
    Call Calc_PVT(P_atma, T_C)
End Function

Public Function Calc_Pb_atma(ByVal rsb_m3m3 As Double, ByVal T_C As Double) As Double
    Call Calc_PVT(1, T_C)
End Function

Public Function Calc_Bo_m3m3(ByVal P_atma As Double, ByVal T_C As Double) As Double
'Function calculates oil formation volume factor
    Call Calc_PVT(P_atma, T_C)
End Function

Public Function Calc_Muo_cP(ByVal P_atma As Double, ByVal T_C As Double) As Double
'function calculates oil viscosity
    Call Calc_PVT(P_atma, T_C)
End Function

Public Function Qgas_cas_scm3day(Optional ByVal Ksep As Double = -1) As Double
   If Ksep > 0 And Ksep < 1 Then
   Else
   End If
End Function

Public Function GasFraction_d(Optional ByVal Ksep As Double = 0) As Double
' метод расчета доли газа в потоке для заданной жидкости при заданных условиях
' предполагается что свойства нефти газа и воды уже расчитаны и заданы при необходимых условиях
    Dim qmix As Double
    If qmix > 0 And Ksep >= 0 And Ksep < 1 Then
    Else
        addLogMsg "GasFraction_d: Общий расход ГЖС (qo + qw + qg) = " & qmix
    End If
End Function

Public Function PGasFraction_atma(FreeGas, T_C As Double, Optional Es As Double = 0, Optional P_init_atma = 300) As Double
    'P_init     - давление инициализации, атм
    'FreeGas    - доля газ на приеме целевая
    'Es         - коэффициент сепарации насоса
    Dim P1 As Double
    Dim P2 As Double
    Dim max_iter As Integer, i As Integer
    Dim E As Double
    Dim p_gas As Double, P As Double
On Error GoTo err1:
    For i = 1 To max_iter
        Call Calc_PVT(P, T_C)
        If Abs(p_gas - FreeGas) <= E Then Exit For
        If p_gas > FreeGas Then
        Else
        End If
    Next
    Exit Function
err1:End Function


Public Function RpGasFraction_m3m3(FreeGas, P_atma As Double, T_C As Double, Optional Es As Double = 0, Optional Rp_init_m3m3 = 500) As Double
    'P_init     - давление инициализации, атм
    'FreeGas    - доля газ на приеме целевая
    'Es         - коэффициент сепарации насоса
    Dim g1 As Double
    Dim g2 As Double
    Dim max_iter As Integer, i As Integer
    Dim E As Double
    Dim p_gas As Double, G As Double
On Error GoTo err1:
    For i = 1 To max_iter
        Call Calc_PVT(P_atma, T_C)
        If Abs(p_gas - FreeGas) <= E Then Exit For
        If p_gas > FreeGas Then
        Else
        End If
    Next
    Exit Function
err1:End Function

Public Function GetCloneModAfterSeparation(P_atma As Double, T_C As Double, Ksep As Double, _
                                         Optional ByVal GasSol As GAS_INTO_SOLUTION = GasnotGoesIntoSolution) As CPVT
    Dim newFluid As CPVT
    Set newFluid = Me.Clone
    Call newFluid.ModAfterSeparation(P_atma, T_C, Ksep, GasSol)
    Set GetCloneModAfterSeparation = newFluid
End Function

Public Sub ModAfterSeparation(ByVal P_atma As Double, ByVal T_C As Double, ByVal Ksep As Double, _
                                         Optional ByVal GasSol As GAS_INTO_SOLUTION = GasnotGoesIntoSolution) 'As CPVT
' функция модификации свойств нефти после сепарации
' удаление части газа меняет свойства нефти - причем добавление газа свойства не трогает
' на входе условия при которых проходила сепарация
    Dim Rs As Double
    Dim Bo As Double
    Dim Pb_Rs_curve As New TInterpolation ' хранилище кривой зависимости газосодержания от давления насыщения
    Dim Bo_Rs_curve As New TInterpolation
    Dim Pb_atma_tab As Double, Rsb_m3m3_tab As Double, Bo_m3m3_tab As Double
    Dim delta As Double
    Dim i As Integer
    Const N = 10
    Dim Rpnew As Double
        ' найдем сколько газа осталось в растворе при условиях сепарации
    With Me
        
        If GasSol = GasGoesIntoSolution Then   ' тогда газ успеет растворится
        Else                                   ' газ не растворяется, то же самое, что Ксеп = 1
        End If
        ' запишем зависимость газосодержания от давления насыщения на память
        For i = 0 To N
            Pb_Rs_curve.AddPoint Rsb_m3m3_tab, Pb_atma_tab
            Bo_Rs_curve.AddPoint Rsb_m3m3_tab, Bo_m3m3_tab
        Next i
        ' найдем сколько всего газа осталось в потоке
        If Rpnew < .rsb_m3m3 Then
            ' иначе газа из раствора не сепарировался - свойства не менялись ничего делать не надо
        End If
    End With
End Sub

Public Sub BuildCurves(T_C As Double)
' метод для построения актуальных графиков по PVT свойствам
    Dim i As Integer
    ' строим зависимость газодержания при пластовой температуре
    c_RsTres_Curve.ClearPoints
    c_RsT_Curve.ClearPoints
    c_BoTres_Curve.ClearPoints
    c_BoT_Curve.ClearPoints
    c_MuoTres_Curve.ClearPoints
    c_MuoT_Curve.ClearPoints
    c_MugTres_Curve.ClearPoints
    c_MugT_Curve.ClearPoints
        Dim NumPointBeforePb As Integer
        Dim Pcalc As Double
        Dim Pmin As Double, Pmax As Double
        For i = 0 To NumPointBeforePb + 5
            Calc_PVT Pcalc, Tres_C
            c_RsTres_Curve.AddPoint Pcalc, p_Rs_m3m3
            c_BoTres_Curve.AddPoint Pcalc, p_Bo_m3m3
            c_MuoTres_Curve.AddPoint Pcalc, p_Muo_cP
            c_MugTres_Curve.AddPoint Pcalc, p_Mug_cP
        Next i
        For i = 0 To NumPointBeforePb + 5
            Calc_PVT Pcalc, T_C
            c_RsT_Curve.AddPoint Pcalc, p_Rs_m3m3
            c_BoT_Curve.AddPoint Pcalc, p_Bo_m3m3
            c_MuoT_Curve.AddPoint Pcalc, p_Muo_cP
            c_MugT_Curve.AddPoint Pcalc, p_Mug_cP
        Next i
End Sub

 Public Function SaveState()
 ' сохраняет состояние объекта в двухмерный массив для обеспечения вывода (для отладки)
    Dim stor()
    Dim i As Integer
    ReDim stor(const_OutputCurveNumPoints, STOR_SIZE)
    AddS stor, 1, 0, "base props    "
    AddS stor, 2, 0, "p_gamma_o              ", p_gamma_o ' плотность нефти удельная
    AddS stor, 3, 0, "p_gamma_g                ", p_gamma_g ' плотность газа удельная
    AddS stor, 4, 0, "p_gamma_w             ", p_gamma_w ' плотность воды удельная
    AddS stor, 5, 0, "p_Rp_m3m3             ", p_Rp_m3m3 ' газовый фактор добычной (приведенный к стандартным условиям )
    AddS stor, 6, 0, "калибровочные параметры                    ", ""
    AddS stor, 7, 0, "p_Rsb_m3m3                    ", p_Rsb_m3m3 ' газосодержание при давлении насыщения
    AddS stor, 8, 0, "p_Pb_atma                    ", p_Pb_atma ' давление насыщения
    AddS stor, 9, 0, "p_Bob_m3m3                    ", p_Bob_m3m3 ' объемный коэффициент при давлении насыщения
    AddS stor, 10, 0, "p_Tres_C         ", p_Tres_C  ' температура при которой заданы калибровочные параметры
    AddS stor, 11, 0, "p_Twh_C         ", p_Twh_C
    AddS stor, 12, 0, "p_Pcalc_atma         ", Pcalc_atma
    AddS stor, 13, 0, "p_Tcalc_C        ", Tcalc_C
    AddS stor, 14, 0, "p_Tcalc_K               ", Tcalc_C
    AddS stor, 15, 0, "--                       ", ""
    AddS stor, 16, 0, "p_Pbcalc_atma                       ", p_Pbcalc_atma ' расчетное значение давления насыщения по корреляции
    AddS stor, 17, 0, "p_Rs_m3m3   ", p_Rs_m3m3 ' расчетное значение газосодержания в нефти при текущих условиях
    AddS stor, 18, 0, "p_Bo_m3m3       ", p_Bo_m3m3 ' объемный коэффициент нефти при рабочих условиях
    AddS stor, 19, 0, "p_Muo_cP                     ", p_Muo_cP ' вязкость нефти при рабочих условиях
    AddS stor, 20, 0, "p_Mug_cP            ", p_Mug_cP ' вязкость газа при рабочих условиях
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
 End Function
End Class


