Imports System
Imports ExcelDna.Integration
Imports ExcelDna.ComInterop
Imports System.Runtime.InteropServices

<ComVisible(True)>
<ClassInterface(ClassInterfaceType.AutoDual)>
<ProgId("Example_COM_lib")>
Public Class COMLibrary

    Public Function add(x As Double, y As Double) As Double
        add = x + y
    End Function

End Class

<ComVisible(False)>
Public Class ExcelAddin
    Implements IExcelAddIn

    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ComServer.DllRegisterServer()
    End Sub

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        ComServer.DllUnregisterServer()
    End Sub
End Class

Public Class Example

    Shared Function unf_Bubblepoint_Valko_McCainSI(rsb_m3m3 As Double, gamma_gas As Double, Temperature_K As Double, gamma_oil As Double)
        Const Min_rsb As Double = 1.8
        Const Max_rsb As Double = 800

        Dim Rsb_old As Double
        Dim API As Double, z1 As Double, z2 As Double, z3 As Double, z4 As Double, z As Double, lnpb As Double

        Rsb_old = rsb_m3m3

        If (rsb_m3m3 < Min_rsb) Then
            rsb_m3m3 = Min_rsb
        End If

        If (rsb_m3m3 > Max_rsb) Then
            rsb_m3m3 = Max_rsb
        End If

        API = 141.5 / gamma_oil - 131.5

        z1 = -4.814074834 + 0.7480913 * Math.Log(rsb_m3m3) + 0.1743556 * Math.Log(rsb_m3m3) ^ 2 - 0.0206 * Math.Log(rsb_m3m3) ^ 3
        z2 = 1.27 - 0.0449 * API + 4.36 * 10 ^ (-4) * API ^ 2 - 4.76 * 10 ^ (-6) * API ^ 3
        z3 = 4.51 - 10.84 * gamma_gas + 8.39 * gamma_gas ^ 2 - 2.34 * gamma_gas ^ 3
        z4 = -7.2254661 + 0.043155 * Temperature_K - 8.5548 * 10 ^ (-5) * Temperature_K ^ 2 + 6.00696 * 10 ^ (-8) * Temperature_K ^ 3
        z = z1 + z2 + z3 + z4
        lnpb = 2.498006 + 0.713 * z + 0.0075 * z ^ 2

        unf_Bubblepoint_Valko_McCainSI = 2.718282 ^ lnpb

        If (Rsb_old < Min_rsb) Then
            unf_Bubblepoint_Valko_McCainSI = (unf_Bubblepoint_Valko_McCainSI - 0.1013) * Rsb_old / Min_rsb + 0.1013
        End If

        If (Rsb_old > Max_rsb) Then
            unf_Bubblepoint_Valko_McCainSI = (unf_Bubblepoint_Valko_McCainSI - 0.1013) * Rsb_old / Max_rsb + 0.1013
        End If
    End Function

    <ExcelFunction(Description:="My first .NET function")>
    Shared Function unf_Density_McCainSI(Pressure_MPa As Double, gamma_gas As Double, Temperature_K As Double,
                                  gamma_oil As Double, Rs_m3_m3 As Double, BP_Pressure_MPa As Double, Compressibility As Double) As Double
        Dim API As Double, ropo As Double, pm As Double, pmmo As Double, epsilon As Double
        Dim i As Integer, counter As Integer
        Dim a0, a1, a2, a3, a4, a5
        Dim roa As Double
        API = 141.5 / gamma_oil - 131.5
        'limit input range to Rs = 800, Pb =1000
        If (Rs_m3_m3 > 800) Then
            Rs_m3_m3 = 800
            BP_Pressure_MPa = unf_Bubblepoint_Valko_McCainSI(Rs_m3_m3, gamma_gas, Temperature_K, gamma_oil)
        End If
        ropo = 845.8 - 0.9 * Rs_m3_m3
        pm = ropo
        pmmo = 0
        epsilon = 0.000001
        i = 0
        counter = 0
        Const MaxIter As Integer = 100
        While (Math.Abs(pmmo - pm) > epsilon And counter < MaxIter)

            i = i + 1

            pmmo = pm
            a0 = -799.21
            a1 = 1361.8
            a2 = -3.70373
            a3 = 0.003
            a4 = 2.98914
            a5 = -0.00223

            roa = a0 + a1 * gamma_gas + a2 * gamma_gas * ropo + a3 * gamma_gas * ropo ^ 2 + a4 * ropo + a5 * ropo ^ 2
            ropo = (Rs_m3_m3 * gamma_gas + 818.81 * gamma_oil) / (0.81881 + Rs_m3_m3 * gamma_gas / roa)
            pm = ropo
            counter = counter + 1
            ' Debug.Assert counter < 20
        End While
        Dim dpp As Double, pbs As Double, dpt As Double
        If Pressure_MPa <= BP_Pressure_MPa Then
            dpp = (0.167 + 16.181 * (10 ^ (-0.00265 * pm))) * (2.32328 * Pressure_MPa) - 0.16 * (0.299 + 263 * (10 ^ (-0.00376 * pm))) * (0.14503774 * Pressure_MPa) ^ 2
            pbs = pm + dpp
            dpt = (0.04837 + 337.094 * pbs ^ (-0.951)) * (1.8 * Temperature_K - 520) ^ 0.938 - (0.346 - 0.3732 * (10 ^ (-0.001 * pbs))) * (1.8 * Temperature_K - 520) ^ 0.475
            pm = pbs - dpt
            unf_Density_McCainSI = pm
        Else
            dpp = (0.167 + 16.181 * (10 ^ (-0.00265 * pm))) * (2.32328 * BP_Pressure_MPa) - 0.16 * (0.299 + 263 * (10 ^ (-0.00376 * pm))) * (0.14503774 * BP_Pressure_MPa) ^ 2
            pbs = pm + dpp
            dpt = (0.04837 + 337.094 * pbs ^ (-0.951)) * (1.8 * Temperature_K - 520) ^ 0.938 - (0.346 - 0.3732 * (10 ^ (-0.001 * pbs))) * (1.8 * Temperature_K - 520) ^ 0.475
            pm = pbs - dpt
            unf_Density_McCainSI = pm * Math.Exp(Compressibility * (Pressure_MPa - BP_Pressure_MPa))
        End If
    End Function


End Class
