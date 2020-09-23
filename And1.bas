Attribute VB_Name = "And1"
Dim I6 As Integer

Public Sub DD1()
On Error GoTo 4
Randomize Timer
4 I = Abs(RNDD(14, False))  'some possible math functions for use
I1 = RNDD(5, False)
I2 = RNDD(5, False)
I3 = RNDD(8, False)
i4 = RNDD(2, False)
If I <> 8 And I <> 1 Then N1 = Abs(RNDD(N2(I), True))
3 RR0 = Abs(RNDD(RR1(I) * 1.05, True))
If RR0 < (RR1(I) / 10) Then GoTo 3
For X = 0 To Form3.ScaleWidth / 1.99
For Y = 0 To Form3.ScaleHeight / 1.99
XX = Form3.ScaleWidth - X
YY = Form3.ScaleHeight - Y
If (X > XX) Then
SLEEP1
Form3.Cls
Dec
End If

If i4 = 1 Then Call Prev1
If i4 = 2 Then Call Prev2

If I3 = 1 Then A = (A1 Xor A2 Xor A3 Xor A4)
If I3 = 2 Then A = (A1 Or A2 Or A3 Or A4)
If I3 = 3 Then A = ((Not A1) Xor (Not A2) Xor (Not A3) Xor (Not A4))
If I3 = 4 Then A = (A1 Eqv A2 Eqv A3 Eqv A4)
If I3 = 5 Then A = ((Not A1) Eqv (Not A2) Eqv (Not A3) Eqv (Not A4))
If I3 = 6 Then A = (A1 And A2 And A3 And A4)

If I3 = 7 Then A = (A1 Eqv A2 Xor A3 Eqv A4)
If I3 = 8 Then A = (A1 Xor A2 Eqv A3 Xor A4)

A = Sgn2(A, I1)
If I2 = 2 Or I = 4 Then A = Abs(A)

'RR0 = RR1(I): N1 = N2(I)
If I = 1 Then A = SZ(CIS(A)) * RR0
If I = 2 Then A = SZ(CIS(A * N1)) * RR0
If I = 3 Then A = SZ(CIS((A * N1) * N1)) * RR0
If I = 4 Then A = SZ(CIS(A - N1)) * RR0
If I = 5 Then A = SZ(CIS(A + N1)) * RR0
If I = 6 Then A = SZ(CIS((A - N1) * N1)) * RR0
If I = 7 Then A = SZ(CIS((A + N1) * N1)) * RR0


If I = 8 Then A = CIS(A) * RR0
If I = 9 Then A = CIS(A * N1) * RR0
If I = 10 Then A = CIS((A * N1) * N1) * RR0
If I = 11 Then A = CIS(A - N1) * RR0
If I = 12 Then A = CIS(A + N1) * RR0
If I = 13 Then A = CIS((A - N1) * N1) * RR0
If I = 14 Then A = CIS((A + N1) * N1) * RR0
A = A ^ 2
SetPixel Form3.hdc, XX, Y, A
SetPixel Form3.hdc, X, YY, A          ' .....
SetPixel Form3.hdc, X, Y, A          ' .....
SetPixel Form3.hdc, XX, YY, A          '.....
Next Y: DoEvents: Next X
End Sub


Public Sub DD2()
On Error GoTo 4
Randomize Timer
4 I = Abs(RNDD(14, False))  'some possible math functions for use
I1 = RNDD(5, False)
I2 = RNDD(5, False)
I3 = RNDD(18, False)
i4 = RNDD(2, False)
If I <> 8 And I <> 1 Then N1 = Abs(RNDD(N2(I), True))
3 RR0 = Abs(RNDD(RR1(I) * 1.05, True))
If RR0 < (RR1(I) / 10) Then GoTo 3
For X = 0 To Form3.ScaleWidth / 1.99
For Y = 0 To Form3.ScaleHeight / 1.99
XX = Form3.ScaleWidth - X
YY = Form3.ScaleHeight - Y
If (X > XX) Then
SLEEP1
Form3.Cls
Dec
End If

If i4 = 1 Then Call Prevv1
If i4 = 2 Then Call Prevv2

If I3 = 1 Then A = (A1 Xor A2 Xor A3 Xor A4)
If I3 = 2 Then A = (A1 Or A2 Or A3 Or A4)
If I3 = 3 Then A = ((Not A1) Xor (Not A2) Xor (Not A3) Xor (Not A4))
If I3 = 4 Then A = (A1 Eqv A2 Eqv A3 Eqv A4)

If I3 = 5 Then A = ((Not A1) Eqv (Not A2) Eqv (Not A3) Eqv (Not A4))
If I3 = 6 Then A = (A1 And A2 And A3 And A4)
If I3 = 7 Then A = (A1 Eqv A2 Xor A3 Eqv A4)
If I3 = 8 Then A = (A1 Xor A2 Eqv A3 Xor A4)

If I3 = 9 Then A = (A1 Eqv A2 Imp A3 Eqv A4)
If I3 = 10 Then A = (A1 Xor A2 Imp A3 Xor A4)
If I3 = 11 Then A = (A1 And A2 Imp A3 And A4)
If I3 = 12 Then A = (A1 Or A2 Imp A3 Or A4)

If I3 = 13 Then A = (A1 Or A2 Xor A3 Or A4)
If I3 = 14 Then A = (A1 Or A2 Eqv A3 Or A4)
If I3 = 15 Then A = (A1 And A2 Or A3 And A4)

If I3 = 15 Then A = (A1 Or A2 Imp A3 Or A4)
If I3 = 17 Then A = (A1 Or A2 And A3 Or A4)
If I3 = 18 Then A = (A1 + A2) * (A3 + A4)


A = Sgn2(A, I1)
If I2 = 2 Or I = 4 Then A = Abs(A)

'RR0 = RR1(I): N1 = N2(I)
If I = 1 Then A = SZ(CIS(A)) * RR0
If I = 2 Then A = SZ(CIS(A * N1)) * RR0
If I = 3 Then A = SZ(CIS((A * N1) * N1)) * RR0
If I = 4 Then A = SZ(CIS(A - N1)) * RR0
If I = 5 Then A = SZ(CIS(A + N1)) * RR0
If I = 6 Then A = SZ(CIS((A - N1) * N1)) * RR0
If I = 7 Then A = SZ(CIS((A + N1) * N1)) * RR0


If I = 8 Then A = CIS(A) * RR0
If I = 9 Then A = CIS(A * N1) * RR0
If I = 10 Then A = CIS((A * N1) * N1) * RR0
If I = 11 Then A = CIS(A - N1) * RR0
If I = 12 Then A = CIS(A + N1) * RR0
If I = 13 Then A = CIS((A - N1) * N1) * RR0
If I = 14 Then A = CIS((A + N1) * N1) * RR0
A = A ^ 2
SetPixel Form3.hdc, XX, Y, A
SetPixel Form3.hdc, X, YY, A          ' .....
SetPixel Form3.hdc, X, Y, A          ' .....
SetPixel Form3.hdc, XX, YY, A          '.....
Next Y: DoEvents: Next X
End Sub


