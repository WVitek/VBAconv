Attribute VB_Name = "u_math"
'=======================================================================================
'Unifloc7.3  Testudines                                           khabibullinra@gmail.com
'���������� ��������� ������� �� ��������� �����������
'2000 - 2018 �
'
'=======================================================================================
' �������������� �������
'
'
'
'


Option Explicit

Public Const MachineEpsilon = 5E-16
Public Const MaxRealNumber = 1E+300
Public Const MinRealNumber = 1E-300

Private Const BigNumber As Double = 1E+70
Private Const SmallNumber As Double = 1E-70


Public Type Complex
    X As Double
    Y As Double
End Type

'Data types


Public Type RCommState
    Stage As Long
    BA() As Boolean
    IA() As Long
    RA() As Double
    CA() As Complex
End Type

Public Type ALGLIBDataset
    NIn As Long
    NOut As Long
    NClasses As Long
    
    Trn() As Double
    Tst() As Double
    val() As Double
    AllDataset() As Double
    
    TrnSize As Long
    TstSize As Long
    ValSize As Long
    TotalSize As Long
End Type




Public Type ODESolverState
    N As Long
    M As Long
    XScale As Double
    h As Double
    eps As Double
    FracEps As Boolean
    YC() As Double
    EScale() As Double
    XG() As Double
    SolverType As Long
    X As Double
    Y() As Double
    DY() As Double
    YTbl() As Double
    RepTerminationType As Long
    RepNFEV As Long
    YN() As Double
    YNS() As Double
    RKA() As Double
    RKC() As Double
    RKCS() As Double
    RKB() As Double
    RKK() As Double
    RState As RCommState
End Type


Public Type ODESolverReport
    NFEV As Long
    TerminationType As Long
End Type




'Global constants
Private Const ODESolverMaxGrow As Double = 3#
Private Const ODESolverMaxShrink As Double = 10#

Function tand(ang) As Double
 tand = Tan(ang / 180 * const_Pi)
End Function


Function cosd(ang) As Double
 cosd = Cos(ang / 180 * const_Pi)
End Function

Function sind(ang) As Double
 sind = Sin(ang / 180 * const_Pi)
End Function

Public Function min(a As Double, B As Double) As Double
  If a < B Then min = a Else min = B
End Function

Public Function max(a As Double, B As Double) As Double
  If a > B Then max = a Else max = B
End Function

Public Function log10(X As Double) As Double
 log10 = Log(X) / Log(10#)
End Function

Public Function isEqual(a As Double, B As Double) As Double
    Const eps = const_P_difference
    isEqual = False
    If Abs(a - B) < eps Then isEqual = True
End Function

Public Function isGreater(a As Double, B As Double) As Double
    Const eps = const_P_difference
    isGreater = False
    If (a - B) > eps Then isGreater = True
End Function
 
Public Function isBetween(a As Double, a0 As Double, a1 As Double)
    isBetween = False
    If ((a <= a0) And (a >= a1)) Or ((a >= a0) And (a <= a1)) Then isBetween = True
End Function
 
Public Function ArcCos(ByVal X As Double) As Double
    Dim t As Double
    t = Sqr(1 - X * X)
    If t < SmallNumber Then
        ArcCos = Atn(BigNumber * Sgn(-X)) + 2 * Atn(1)
    Else
        ArcCos = Atn(-X / t) + 2 * Atn(1)
    End If
End Function
 

Public Function ArcSin(ByVal X As Double) As Double
    Dim t As Double
    t = Sqr(1 - X * X)
    If t < SmallNumber Then
        ArcSin = Atn(BigNumber * Sgn(X))
    Else
        ArcSin = Atn(X / t)
    End If
End Function

 
Public Function ArcTan(X As Double) As Double
  'Inverse Tangent
    On Error Resume Next
        ArcTan = Atn(X) '* (180 / PI)
    On Error GoTo 0
End Function


Public Function MaxReal(ByVal M1 As Double, ByVal M2 As Double) As Double
    If M1 > M2 Then
        MaxReal = M1
    Else
        MaxReal = M2
    End If
End Function

Public Function MinReal(ByVal M1 As Double, ByVal M2 As Double) As Double
    If M1 < M2 Then
        MinReal = M1
    Else
        MinReal = M2
    End If
End Function

Public Function MaxInt(ByVal M1 As Long, ByVal M2 As Long) As Long
    If M1 > M2 Then
        MaxInt = M1
    Else
        MaxInt = M2
    End If
End Function

Public Function MinInt(ByVal M1 As Long, ByVal M2 As Long) As Long
    If M1 < M2 Then
        MinInt = M1
    Else
        MinInt = M2
    End If
End Function




Public Function SinH(ByVal X As Double) As Double
    SinH = (Exp(X) - Exp(-X)) / 2
End Function

Public Function CosH(ByVal X As Double) As Double
    CosH = (Exp(X) + Exp(-X)) / 2
End Function

Public Function TanH(ByVal X As Double) As Double
    Dim t As Double
    If X > 0 Then
        t = Exp(-X)
        t = t * t
        TanH = (1 - t) / (1 + t)
    Else
        t = Exp(X)
        t = t * t
        TanH = (t - 1) / (t + 1)
    End If
End Function


Public Function Power(ByVal Base As Double, ByVal Exponent As Double) As Double
    Power = Base ^ Exponent
End Function

Public Function Square(ByVal X As Double) As Double
    Square = X * X
End Function


Public Function Ceil(ByVal X As Double) As Double
    Ceil = -Int(-X)
End Function

Public Function RandomInteger(ByVal X As Long) As Long
    RandomInteger = Int(Rnd() * X)
End Function

Public Function Atn2(ByVal Y As Double, ByVal X As Double) As Double
    If SmallNumber * Abs(Y) < Abs(X) Then
        If X < 0 Then
            If Y = 0 Then
                Atn2 = const_Pi
            Else
                Atn2 = Atn(Y / X) + const_Pi * Sgn(Y)
            End If
        Else
            Atn2 = Atn(Y / X)
        End If
    Else
        Atn2 = Sgn(Y) * const_Pi / 2
    End If
End Function


' ===========================================================================================
'  math_complex
' ===========================================================================================

'=======================================================================================
'Unifloc7.3  Testudines                                           khabibullinra@gmail.com
'���������� ��������� ������� �� ��������� �����������
'2000 - 2018 �
'
'=======================================================================================





Public Function C_Complex(ByVal X As Double) As Complex
    Dim result As Complex

    result.X = X
    result.Y = 0

    C_Complex = result
End Function


Public Function AbsComplex(ByRef z As Complex) As Double
    Dim result As Double
    Dim W As Double
    Dim XABS As Double
    Dim YABS As Double
    Dim v As Double

    XABS = Abs(z.X)
    YABS = Abs(z.Y)
    W = MaxReal(XABS, YABS)
    v = MinReal(XABS, YABS)
    If v = 0 Then
        result = W
    Else
        result = W * Sqr(1 + Square(v / W))
    End If

    AbsComplex = result
End Function


Public Function C_Opposite(ByRef z As Complex) As Complex
    Dim result As Complex

    result.X = -z.X
    result.Y = -z.Y

    C_Opposite = result
End Function


Public Function Conj(ByRef z As Complex) As Complex
    Dim result As Complex

    result.X = z.X
    result.Y = -z.Y

    Conj = result
End Function


Public Function CSqr(ByRef z As Complex) As Complex
    Dim result As Complex

    result.X = Square(z.X) - Square(z.Y)
    result.Y = 2 * z.X * z.Y

    CSqr = result
End Function


Public Function C_Add(ByRef z1 As Complex, ByRef z2 As Complex) As Complex
    Dim result As Complex

    result.X = z1.X + z2.X
    result.Y = z1.Y + z2.Y

    C_Add = result
End Function


Public Function C_Mul(ByRef z1 As Complex, ByRef z2 As Complex) As Complex
    Dim result As Complex

    result.X = z1.X * z2.X - z1.Y * z2.Y
    result.Y = z1.X * z2.Y + z1.Y * z2.X

    C_Mul = result
End Function


Public Function C_AddR(ByRef z1 As Complex, ByVal r As Double) As Complex
    Dim result As Complex

    result.X = z1.X + r
    result.Y = z1.Y

    C_AddR = result
End Function


Public Function C_MulR(ByRef z1 As Complex, ByVal r As Double) As Complex
    Dim result As Complex

    result.X = z1.X * r
    result.Y = z1.Y * r

    C_MulR = result
End Function


Public Function C_Sub(ByRef z1 As Complex, ByRef z2 As Complex) As Complex
    Dim result As Complex

    result.X = z1.X - z2.X
    result.Y = z1.Y - z2.Y

    C_Sub = result
End Function


Public Function C_SubR(ByRef z1 As Complex, ByVal r As Double) As Complex
    Dim result As Complex

    result.X = z1.X - r
    result.Y = z1.Y

    C_SubR = result
End Function


Public Function C_RSub(ByVal r As Double, ByRef z1 As Complex) As Complex
    Dim result As Complex

    result.X = r - z1.X
    result.Y = -z1.Y

    C_RSub = result
End Function


Public Function C_Div(ByRef z1 As Complex, ByRef z2 As Complex) As Complex
    Dim result As Complex
    Dim a As Double
    Dim B As Double
    Dim c As Double
    Dim d As Double
    Dim E As Double
    Dim f As Double

    a = z1.X
    B = z1.Y
    c = z2.X
    d = z2.Y
    If Abs(d) < Abs(c) Then
        E = d / c
        f = c + d * E
        result.X = (a + B * E) / f
        result.Y = (B - a * E) / f
    Else
        E = c / d
        f = d + c * E
        result.X = (B + a * E) / f
        result.Y = (-a + B * E) / f
    End If

    C_Div = result
End Function


Public Function C_DivR(ByRef z1 As Complex, ByVal r As Double) As Complex
    Dim result As Complex

    result.X = z1.X / r
    result.Y = z1.Y / r

    C_DivR = result
End Function


Public Function C_RDiv(ByVal r As Double, ByRef z2 As Complex) As Complex
    Dim result As Complex
    Dim a As Double
    Dim c As Double
    Dim d As Double
    Dim E As Double
    Dim f As Double

    a = r
    c = z2.X
    d = z2.Y
    If Abs(d) < Abs(c) Then
        E = d / c
        f = c + d * E
        result.X = a / f
        result.Y = -(a * E / f)
    Else
        E = c / d
        f = d + c * E
        result.X = a * E / f
        result.Y = -(a / f)
    End If

    C_RDiv = result
End Function


Public Function C_Equal(ByRef z1 As Complex, ByRef z2 As Complex) As Boolean
    Dim result As Boolean

    result = z1.X = z2.X And z1.Y = z2.Y

    C_Equal = result
End Function


Public Function C_NotEqual(ByRef z1 As Complex, _
         ByRef z2 As Complex) As Boolean
    Dim result As Boolean

    result = z1.X <> z2.X Or z1.Y <> z2.Y

    C_NotEqual = result
End Function

Public Function C_EqualR(ByRef z1 As Complex, ByVal r As Double) As Boolean
    Dim result As Boolean

    result = z1.X = r And z1.Y = 0

    C_EqualR = result
End Function


Public Function C_NotEqualR(ByRef z1 As Complex, _
         ByVal r As Double) As Boolean
    Dim result As Boolean

    result = z1.X <> r Or z1.Y <> 0

    C_NotEqualR = result
End Function





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Copyright 2009 by Sergey Bochkanov (ALGLIB project).
'
'>>> SOURCE LICENSE >>>
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation (www.fsf.org); either version 2 of the
'License, or (at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'A copy of the GNU General Public License is available at
'http://www.fsf.org/licensing/licenses
'
'>>> END OF LICENSE >>>
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'Routines
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Cash-Karp adaptive ODE solver.
'
'This subroutine solves ODE  Y'=f(Y,x)  with  initial  conditions  Y(xs)=Ys
'(here Y may be single variable or vector of N variables).
'
'INPUT PARAMETERS:
'    Y       -   initial conditions, array[0..N-1].
'                contains values of Y[] at X[0]
'    N       -   system size
'    X       -   points at which Y should be tabulated, array[0..M-1]
'                integrations starts at X[0], ends at X[M-1],  intermediate
'                values at X[i] are returned too.
'                SHOULD BE ORDERED BY ASCENDING OR BY DESCENDING!!!!
'    M       -   number of intermediate points + first point + last point:
'                * M>2 means that you need both Y(X[M-1]) and M-2 values at
'                  intermediate points
'                * M=2 means that you want just to integrate from  X[0]  to
'                  X[1] and don't interested in intermediate values.
'                * M=1 means that you don't want to integrate :)
'                  it is degenerate case, but it will be handled correctly.
'                * M<1 means error
'    Eps     -   tolerance (absolute/relative error on each  step  will  be
'                less than Eps). When passing:
'                * Eps>0, it means desired ABSOLUTE error
'                * Eps<0, it means desired RELATIVE error.  Relative errors
'                  are calculated with respect to maximum values of  Y seen
'                  so far. Be careful to use this criterion  when  starting
'                  from Y[] that are close to zero.
'    H       -   initial  step  lenth,  it  will  be adjusted automatically
'                after the first  step.  If  H=0,  step  will  be  selected
'                automatically  (usualy  it  will  be  equal  to  0.001  of
'                min(x[i]-x[j])).
'
'OUTPUT PARAMETERS
'    State   -   structure which stores algorithm state between  subsequent
'                calls of OdeSolverIteration. Used for reverse communication.
'                This structure should be passed  to the OdeSolverIteration
'                subroutine.
'
'SEE ALSO
'    AutoGKSmoothW, AutoGKSingular, AutoGKIteration, AutoGKResults.
'
'
'  -- ALGLIB --
'     Copyright 01.09.2009 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ODESolverRKCK(ByRef Y() As Double, _
         ByVal N As Long, _
         ByRef X() As Double, _
         ByVal M As Long, _
         ByVal eps As Double, _
         ByVal h As Double, _
         ByRef State As ODESolverState)

    Call ODESolverInit(0#, Y, N, X, M, eps, h, State)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'One iteration of ODE solver.
'
'Called after inialization of State structure with OdeSolverXXX subroutine.
'See HTML docs for examples.
'
'INPUT PARAMETERS:
'    State   -   structure which stores algorithm state between subsequent
'                calls and which is used for reverse communication. Must be
'                initialized with OdeSolverXXX() call first.
'
'If subroutine returned False, algorithm have finished its work.
'If subroutine returned True, then user should:
'* calculate F(State.X, State.Y)
'* store it in State.DY
'Here State.X is real, State.Y and State.DY are arrays[0..N-1] of reals.
'
'  -- ALGLIB --
'     Copyright 01.09.2009 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ODESolverIteration(ByRef State As ODESolverState) As Boolean
    Dim result As Boolean
    Dim N As Long
    Dim M As Long
    Dim i As Long
    Dim j As Long
    Dim K As Long
    Dim XC As Double
    Dim v As Double
    Dim h As Double
    Dim H2 As Double
    Dim GridPoint As Boolean
    Dim Err As Double
    Dim MaxGrowPow As Double
    Dim KLimit As Long
    Dim i_ As Long

    
    '
    ' Reverse communication preparations
    ' I know it looks ugly, but it works the same way
    ' anywhere from C++ to Python.
    '
    ' This code initializes locals by:
    ' * random values determined during code
    '   generation - on first subroutine call
    ' * values from previous call - on subsequent calls
    '
    If State.RState.Stage >= 0# Then
        N = State.RState.IA(0#)
        M = State.RState.IA(1#)
        i = State.RState.IA(2#)
        j = State.RState.IA(3#)
        K = State.RState.IA(4#)
        KLimit = State.RState.IA(5#)
        GridPoint = State.RState.BA(0#)
        XC = State.RState.RA(0#)
        v = State.RState.RA(1#)
        h = State.RState.RA(2#)
        H2 = State.RState.RA(3#)
        Err = State.RState.RA(4#)
        MaxGrowPow = State.RState.RA(5#)
    Else
        N = -983#
        M = -989#
        i = -834#
        j = 900#
        K = -287#
        KLimit = 364#
        GridPoint = False
        XC = -338#
        v = -686#
        h = 912#
        H2 = 585#
        Err = 497#
        MaxGrowPow = -271#
    End If
    If State.RState.Stage = 0# Then
        GoTo lbl_0
    End If
    
    '
    ' Routine body
    '
    
    '
    ' prepare
    '
    If State.RepTerminationType <> 0# Then
        result = False
        ODESolverIteration = result
        Exit Function
    End If
    N = State.N
    M = State.M
    h = State.h
    ReDim State.Y(0 To N - 1)
    ReDim State.DY(0 To N - 1)
    MaxGrowPow = Power(ODESolverMaxGrow, 5#)
    State.RepNFEV = 0#
    
    '
    ' some preliminary checks for internal errors
    ' after this we assume that H>0 and M>1
    '
    
    '
    ' choose solver
    '
    If State.SolverType <> 0# Then
        GoTo lbl_1
    End If
    
    '
    ' Cask-Karp solver
    ' Prepare coefficients table.
    ' Check it for errors
    '
    ReDim State.RKA(0 To 6# - 1)
    State.RKA(0#) = 0#
    State.RKA(1#) = 1# / 5#
    State.RKA(2#) = 3# / 10#
    State.RKA(3#) = 3# / 5#
    State.RKA(4#) = 1#
    State.RKA(5#) = 7# / 8#
    ReDim State.RKB(0 To 6# - 1, 0 To 5# - 1)
    State.RKB(1#, 0#) = 1# / 5#
    State.RKB(2#, 0#) = 3# / 40#
    State.RKB(2#, 1#) = 9# / 40#
    State.RKB(3#, 0#) = 3# / 10#
    State.RKB(3#, 1#) = -(9# / 10#)
    State.RKB(3#, 2#) = 6# / 5#
    State.RKB(4#, 0#) = -(11# / 54#)
    State.RKB(4#, 1#) = 5# / 2#
    State.RKB(4#, 2#) = -(70# / 27#)
    State.RKB(4#, 3#) = 35# / 27#
    State.RKB(5#, 0#) = 1631# / 55296#
    State.RKB(5#, 1#) = 175# / 512#
    State.RKB(5#, 2#) = 575# / 13824#
    State.RKB(5#, 3#) = 44275# / 110592#
    State.RKB(5#, 4#) = 253# / 4096#
    ReDim State.RKC(0 To 6# - 1)
    State.RKC(0#) = 37# / 378#
    State.RKC(1#) = 0#
    State.RKC(2#) = 250# / 621#
    State.RKC(3#) = 125# / 594#
    State.RKC(4#) = 0#
    State.RKC(5#) = 512# / 1771#
    ReDim State.RKCS(0 To 6# - 1)
    State.RKCS(0#) = 2825# / 27648#
    State.RKCS(1#) = 0#
    State.RKCS(2#) = 18575# / 48384#
    State.RKCS(3#) = 13525# / 55296#
    State.RKCS(4#) = 277# / 14336#
    State.RKCS(5#) = 1# / 4#
    ReDim State.RKK(0 To 6# - 1, 0 To N - 1)
    
    '
    ' Main cycle consists of two iterations:
    ' * outer where we travel from X[i-1] to X[i]
    ' * inner where we travel inside [X[i-1],X[i]]
    '
    ReDim State.YTbl(0 To M - 1, 0 To N - 1)
    ReDim State.EScale(0 To N - 1)
    ReDim State.YN(0 To N - 1)
    ReDim State.YNS(0 To N - 1)
    XC = State.XG(0#)
    For i_ = 0# To N - 1# Step 1
        State.YTbl(0#, i_) = State.YC(i_)
    Next i_
    For j = 0# To N - 1# Step 1
        State.EScale(j) = 0#
    Next j
    i = 1#
lbl_3:
    If i > M - 1# Then
        GoTo lbl_5
    End If
    
    '
    ' begin inner iteration
    '
lbl_6:
    If False Then
        GoTo lbl_7
    End If
    
    '
    ' truncate step if needed (beyond right boundary).
    ' determine should we store X or not
    '
    If XC + h >= State.XG(i) Then
        h = State.XG(i) - XC
        GridPoint = True
    Else
        GridPoint = False
    End If
    
    '
    ' Update error scale maximums
    '
    ' These maximums are initialized by zeros,
    ' then updated every iterations.
    '
    For j = 0# To N - 1# Step 1
        State.EScale(j) = MaxReal(State.EScale(j), Abs(State.YC(j)))
    Next j
    
    '
    ' make one step:
    ' 1. calculate all info needed to do step
    ' 2. update errors scale maximums using values/derivatives
    '    obtained during (1)
    '
    ' Take into account that we use scaling of X to reduce task
    ' to the form where x[0] < x[1] < ... < x[n-1]. So X is
    ' replaced by x=xscale*t, and dy/dx=f(y,x) is replaced
    ' by dy/dt=xscale*f(y,xscale*t).
    '
    For i_ = 0# To N - 1# Step 1
        State.YN(i_) = State.YC(i_)
    Next i_
    For i_ = 0# To N - 1# Step 1
        State.YNS(i_) = State.YC(i_)
    Next i_
    K = 0#
lbl_8:
    If K > 5# Then
        GoTo lbl_10
    End If
    
    '
    ' prepare data for the next update of YN/YNS
    '
    State.X = State.XScale * (XC + State.RKA(K) * h)
    For i_ = 0# To N - 1# Step 1
        State.Y(i_) = State.YC(i_)
    Next i_
    For j = 0# To K - 1# Step 1
        v = State.RKB(K, j)
        For i_ = 0# To N - 1# Step 1
            State.Y(i_) = State.Y(i_) + v * State.RKK(j, i_)
        Next i_
    Next j
    State.RState.Stage = 0#
    GoTo lbl_rcomm
lbl_0:
    State.RepNFEV = State.RepNFEV + 1#
    v = h * State.XScale
    For i_ = 0# To N - 1# Step 1
        State.RKK(K, i_) = v * State.DY(i_)
    Next i_
    
    '
    ' update YN/YNS
    '
    v = State.RKC(K)
    For i_ = 0# To N - 1# Step 1
        State.YN(i_) = State.YN(i_) + v * State.RKK(K, i_)
    Next i_
    v = State.RKCS(K)
    For i_ = 0# To N - 1# Step 1
        State.YNS(i_) = State.YNS(i_) + v * State.RKK(K, i_)
    Next i_
    K = K + 1#
    GoTo lbl_8
lbl_10:
    
    '
    ' estimate error
    '
    Err = 0#
    For j = 0# To N - 1# Step 1
        If Not State.FracEps Then
            
            '
            ' absolute error is estimated
            '
            Err = MaxReal(Err, Abs(State.YN(j) - State.YNS(j)))
        Else
            
            '
            ' Relative error is estimated
            '
            v = State.EScale(j)
            If v = 0# Then
                v = 1#
            End If
            Err = MaxReal(Err, Abs(State.YN(j) - State.YNS(j)) / v)
        End If
    Next j
    
    '
    ' calculate new step, restart if necessary
    '
    If MaxGrowPow * Err <= State.eps Then
        H2 = ODESolverMaxGrow * h
    Else
        H2 = h * Power(State.eps / Err, 0.2)
    End If
    If H2 < h / ODESolverMaxShrink Then
        H2 = h / ODESolverMaxShrink
    End If
    If Err > State.eps Then
        h = H2
        GoTo lbl_6
    End If
    
  '  If H > 10 Then H = 10
    '
    ' advance position
    '
    XC = XC + h
    For i_ = 0# To N - 1# Step 1
        State.YC(i_) = State.YN(i_)
    Next i_
    
    '
    ' update H
    '
    h = H2
    
   '   If H > 10 Then H = 10
  
    '
    ' break on grid point
    '
    If GridPoint Then
        GoTo lbl_7
    End If
    GoTo lbl_6
lbl_7:
    
    '
    ' save result
    '
    For i_ = 0# To N - 1# Step 1
        State.YTbl(i, i_) = State.YC(i_)
    Next i_
    i = i + 1#
    GoTo lbl_3
lbl_5:
    State.RepTerminationType = 1#
    result = False
    ODESolverIteration = result
    Exit Function
lbl_1:
    result = False
    ODESolverIteration = result
    Exit Function
    
    '
    ' Saving state
    '
lbl_rcomm:
    result = True
    State.RState.IA(0#) = N
    State.RState.IA(1#) = M
    State.RState.IA(2#) = i
    State.RState.IA(3#) = j
    State.RState.IA(4#) = K
    State.RState.IA(5#) = KLimit
    State.RState.BA(0#) = GridPoint
    State.RState.RA(0#) = XC
    State.RState.RA(1#) = v
    State.RState.RA(2#) = h
    State.RState.RA(3#) = H2
    State.RState.RA(4#) = Err
    State.RState.RA(5#) = MaxGrowPow

    ODESolverIteration = result
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ODE solver results
'
'Called after OdeSolverIteration returned False.
'
'INPUT PARAMETERS:
'    State   -   algorithm state (used by OdeSolverIteration).
'
'OUTPUT PARAMETERS:
'    M       -   number of tabulated values, M>=1
'    XTbl    -   array[0..M-1], values of X
'    YTbl    -   array[0..M-1,0..N-1], values of Y in X[i]
'    Rep     -   solver report:
'                * Rep.TerminationType completetion code:
'                    * -2    X is not ordered  by  ascending/descending  or
'                            there are non-distinct X[],  i.e.  X[i]=X[i+1]
'                    * -1    incorrect parameters were specified
'                    *  1    task has been solved
'                * Rep.NFEV contains number of function calculations
'
'  -- ALGLIB --
'     Copyright 01.09.2009 by Bochkanov Sergey
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ODESolverResults(ByRef State As ODESolverState, _
         ByRef M As Long, _
         ByRef XTbl() As Double, _
         ByRef YTbl() As Double, _
         ByRef Rep As ODESolverReport)
    Dim v As Double
    Dim i As Long
    Dim i_ As Long

    Rep.TerminationType = State.RepTerminationType
    If Rep.TerminationType > 0# Then
        M = State.M
        Rep.NFEV = State.RepNFEV
        ReDim XTbl(0 To State.M - 1)
        v = State.XScale
        For i_ = 0# To State.M - 1# Step 1
            XTbl(i_) = v * State.XG(i_)
        Next i_
        ReDim YTbl(0 To State.M - 1, 0 To State.N - 1)
        For i = 0# To State.M - 1# Step 1
            For i_ = 0# To State.N - 1# Step 1
                YTbl(i, i_) = State.YTbl(i, i_)
            Next i_
        Next i
    Else
        Rep.NFEV = 0#
    End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Internal initialization subroutine
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ODESolverInit(ByVal SolverType As Long, _
         ByRef Y() As Double, _
         ByVal N As Long, _
         ByRef X() As Double, _
         ByVal M As Long, _
         ByVal eps As Double, _
         ByVal h As Double, _
         ByRef State As ODESolverState)
    Dim i As Long
    Dim v As Double
    Dim i_ As Long

    
    '
    ' Prepare RComm
    '
    ReDim State.RState.IA(0# To 5#)
    ReDim State.RState.BA(0# To 0#)
    ReDim State.RState.RA(0# To 5#)
    State.RState.Stage = -1#
    
    '
    ' check parameters.
    '
    If N <= 0# Or M < 1# Or eps = 0# Then
        State.RepTerminationType = -1#
        Exit Sub
    End If
    If h < 0# Then
        h = -h
    End If
    
    '
    ' quick exit if necessary.
    ' after this block we assume that M>1
    '
    If M = 1# Then
        State.RepNFEV = 0#
        State.RepTerminationType = 1#
        ReDim State.YTbl(0 To 1# - 1, 0 To N - 1)
        For i_ = 0# To N - 1# Step 1
            State.YTbl(0#, i_) = Y(i_)
        Next i_
        ReDim State.XG(0 To M - 1)
        For i_ = 0# To M - 1# Step 1
            State.XG(i_) = X(i_)
        Next i_
        Exit Sub
    End If
    
    '
    ' check again: correct order of X[]
    '
    If X(1#) = X(0#) Then
        State.RepTerminationType = -2#
        Exit Sub
    End If
    For i = 1# To M - 1# Step 1
        If X(1#) > X(0#) And X(i) <= X(i - 1#) Or X(1#) < X(0#) And X(i) >= X(i - 1#) Then
            State.RepTerminationType = -2#
            Exit Sub
        End If
    Next i
    
    '
    ' auto-select H if necessary
    '
    If h = 0# Then
        v = Abs(X(1#) - X(0#))
        For i = 2# To M - 1# Step 1
            v = MinReal(v, Abs(X(i) - X(i - 1#)))
        Next i
        h = 0.001 * v
    End If
    
    '
    ' store parameters
    '
    State.N = N
    State.M = M
    State.h = h
    State.eps = Abs(eps)
    State.FracEps = eps < 0#
    ReDim State.XG(0 To M - 1)
    For i_ = 0# To M - 1# Step 1
        State.XG(i_) = X(i_)
    Next i_
    If X(1#) > X(0#) Then
        State.XScale = 1#
    Else
        State.XScale = -1#
        For i_ = 0# To M - 1# Step 1
            State.XG(i_) = -1 * State.XG(i_)
        Next i_
    End If
    ReDim State.YC(0 To N - 1)
    For i_ = 0# To N - 1# Step 1
        State.YC(i_) = Y(i_)
    Next i_
    State.SolverType = SolverType
    State.RepTerminationType = 0#
End Sub




