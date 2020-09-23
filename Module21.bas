Attribute VB_Name = "Module21"
Public Sub Prev1()
A1 = (Sgn2(X Eqv Y, I5) + Sgn2(XX And YY, I5))
A2 = (Sgn2(XX Eqv Y, I5) + Sgn2(X And YY, I5))
A3 = (Sgn2(X Eqv YY, I5) + Sgn2(XX And Y, I5))
A4 = (Sgn2(XX Eqv YY, I5) + Sgn2(X And Y, I5))
End Sub

Public Sub Prev2()
A1 = (Sgn2(X Or Y, I5) + Sgn2(XX Xor YY, I5))
A2 = (Sgn2(XX Or Y, I5) + Sgn2(X Xor YY, I5))
A3 = (Sgn2(X Or YY, I5) + Sgn2(XX Xor Y, I5))
A4 = (Sgn2(XX Or YY, I5) + Sgn2(X Xor Y, I5))
End Sub

