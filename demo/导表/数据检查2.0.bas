Attribute VB_Name = "���ݼ��"
Sub ���ݼ��()
Dim YqV As Integer, YqD As Integer, YqT As Integer, YqP As Integer, Yq_VD_Z As Integer, YqZ As Integer '��������
Dim NumYK
Dim X_Tem As Single, X_HB As Single
'����ÿ����������
YqV = 5  '���ټ�
YqD = 3 '�����
YqT = 1 '�¶ȴ�����
YqP = 1 'ѹ��������
YqZ = YqV + YqD + YqT + YqP
Yq_VD_Z = YqV + YqD
'����ÿ����������ռ������
NumYK = 4
'�����ֳ����μ���ƽ���¶�
X_Tem = -2     '��ƽ���¶� ���϶�
'X_HB = 1509     '����������
X_HB = Split(Cells(5, 1), " ")(2)

ReDim Vspeed(YqV) As Integer
ReDim HVgao(YqV) As String
Dim Rng As Range

ReDim Ddirect(YqD) As Integer
ReDim HDgao(YqD) As Integer

ReDim YqTP(YqT + YqP) As Integer
ReDim YqTPm(YqT + YqP) As String
ReDim Cha(YqT + YqP) As Single


ReDim Yqm(YqZ) As String
Dim Vlow As Single, Vtop As Single, Dlow As Single, Dtop As Single, Plow As Single, Ptop As Single, Tlow As Single, Ttop As Single
ReDim Fanw(YqZ, 2) As Single
Dim col As Integer, Tcol As Integer



Dim Snum As Long
Set Rng = Range("e3")  'ͳ��������ʼ�У��������ϲ���������������ٴ�3��2�п�ʼ
V_X_Qrow = Rng.Row
V_X_Qcolumn = Rng.Column
D_X_Qrow = Rng.Row
D_X_Qcolumn = Rng.Column + 6 + UBound(Vspeed)
V_Q_Qrow = Rng.Row - 1
V_Q_Qcolumn = D_X_Qcolumn + UBound(Ddirect) + 2
Y_Q_Qrow = Rng.Row - 1
Y_Q_Qcolumn = V_Q_Qcolumn + 3
Z_F_Qrow = Rng.Row - 2
Z_F_Qcolumn = Y_Q_Qcolumn + 3
Z_B_Qrow = Rng.Row - 2
Z_B_Column = Z_F_Qcolumn + 3 + UBound(Yqm)
V_XX_Qrow = Rng.Row - 1
V_XX_Qcolumn = Z_B_Column + 4
Z_SZH_Qrow = Rng.Row
Z_SZH_QColumn = V_XX_Qcolumn + UBound(Vspeed) + 3

Application.ScreenUpdating = False
  '��ȡӦ�����ݸ���
  Snum = Range("a14").CurrentRegion.Rows.Count - 1


'������������
For i = 1 To YqV
    mn = Split(Cells(13, 2 + (i - 1) * NumYK), " ")
    For j = 1 To UBound(mn) - 1
        Yqm(i - 1) = Yqm(i - 1) & mn(j)
    Next j
Next i

For i = 1 To YqD
    mn = Split(Cells(13, 2 + (YqV + i - 1) * NumYK), " ")
    Yqm(YqV + i - 1) = mn(1) & "D"
Next i

Yqm(YqV + YqD) = "�¶�"
Yqm(YqV + YqD + 1) = "��ѹ"


'����仯��ֵ
Vlow = 0 '������Сֵ
Vtop = 40 '�������ֵ
Dlow = 0 '������Сֵ
Dtop = 360.1 '�������ֵ
Plow = 94 / (10 ^ (X_HB / 18400 * (1 + (1 / 273) * (X_Tem + X_HB / 200 * 0.6)))) '��ѹ��СֵkPa�� ���躣������100�����½���0.6�棬ѹ�߹�ʽ
Ptop = 106 / (10 ^ (X_HB / 18400 * (1 + (1 / 273) * (X_Tem + X_HB / 200 * 0.6))))  '��ѹ���ֵkPa
Tlow = -40 '�¶���Сֵ���϶�
Ttop = 50 '�¶����ֵ���϶�

'����仯��ֵ
Cha(0) = 5 'Сʱ�¶ȱ仯��ֵ
Cha(1) = 1 'Сʱ��ѹ�仯��ֵ

For i = 0 To YqV - 1

Fanw(i, 0) = Vlow    '���ٺ���Χ
Fanw(i, 1) = Vtop
Vspeed(i) = i * NumYK + 2 '����������
HVgao(i) = Replace(Replace(Replace(Yqm(i), "A", ""), "B", ""), "m", "") '������
Next i

For i = 0 To YqD - 1
Fanw(i + YqV, 0) = Dlow   '�������Χ
Fanw(i + YqV, 1) = Dtop
Ddirect(i) = (i + YqV) * NumYK + 2    '����������
HDgao(i) = CInt(Replace(Replace(Replace(Yqm(i + YqV), "A", ""), "B", ""), "D", "")) 'CInt(Left(Yqm(i + YqV), 2))   '�������ڸ߶�
Next i

For i = 0 To YqT - 1
Fanw(i + YqV + YqD, 0) = Tlow    '�¶Ⱥ���Χ
Fanw(i + YqV + YqD, 1) = Ttop
YqTP(i) = (i + YqV + YqD) * NumYK + 2   '�¶�������  *****ע���¶���ѹ��˳��
YqTPm(i) = Yqm(YqV + YqD)
Next i

Tcol = YqTP(0) '�¶�������  *****�������ݼ����

For i = 0 To YqP - 1
Fanw(i + YqV + YqD + YqT, 0) = Plow    '��ѹ����Χ
Fanw(i + YqV + YqD + YqT, 1) = Ptop
YqTP(YqT + i) = (i + YqV + YqD + YqT) * NumYK + 2   '��ѹ������
YqTPm(i + YqT) = Yqm(YqV + YqD + YqT)
Next i

'ͳ�Ʋ��������
Dim Str_CV$, Str_CX$, Str_CT$, Str_CP$
Cells(Z_SZH_Qrow - 1, Z_SZH_QColumn) = "����λ��"
Range(Cells(Z_SZH_Qrow, Z_SZH_QColumn), Cells(Z_SZH_Qrow + 1, Z_SZH_QColumn)).Merge
Cells(Z_SZH_Qrow, Z_SZH_QColumn) = Replace(Split(Range("a3"), " ")(3), "?", "��") & Split(Range("a3"), " ")(4)
Range(Cells(Z_SZH_Qrow + 2, Z_SZH_QColumn), Cells(Z_SZH_Qrow + 3, Z_SZH_QColumn)).Merge
Cells(Z_SZH_Qrow + 2, Z_SZH_QColumn) = Replace(Split(Range("a4"), " ")(3), "?", "��") & Split(Range("a3"), " ")(4)
Cells(Z_SZH_Qrow - 1, Z_SZH_QColumn + 1) = "�߳�(m)"
Range(Cells(Z_SZH_Qrow, Z_SZH_QColumn + 1), Cells(Z_SZH_Qrow + 3, Z_SZH_QColumn + 1)).Merge
Cells(Z_SZH_Qrow, Z_SZH_QColumn + 1) = X_HB
Cells(Z_SZH_Qrow - 1, Z_SZH_QColumn + 2) = "�߶�(m)"
Range(Cells(Z_SZH_Qrow, Z_SZH_QColumn + 2), Cells(Z_SZH_Qrow + 3, Z_SZH_QColumn + 2)).Merge
Cells(Z_SZH_Qrow, Z_SZH_QColumn + 2) = HVgao(0)
Cells(Z_SZH_Qrow - 1, Z_SZH_QColumn + 3) = "��Ŀ�����"
'Str_CV = Yqm(0)
'For i = 1 To YqV - 1
'Str_CV = Str_CV & "/" & Yqm(i)
'Next
'Cells(Z_SZH_Qrow, Z_SZH_QColumn + 3) = "���٣�" & Str_CV
'Str_CX = Yqm(YqV)
'For i = YqV + 1 To YqV + YqD - 1
'Str_CX = Str_CX & "/" & Yqm(i)
'Next
'Cells(Z_SZH_Qrow + 1, Z_SZH_QColumn + 3) = "����" & Str_CX
Str_CV = HVgao(0)
For i = 1 To YqV - 1
Str_CV = Str_CV & "/" & HVgao(i)
Next
Cells(Z_SZH_Qrow, Z_SZH_QColumn + 3) = "���٣�" & Str_CV
Str_CX = HDgao(0)
For i = 1 To YqD - 1
Str_CX = Str_CX & "/" & HDgao(i)
Next
Cells(Z_SZH_Qrow + 1, Z_SZH_QColumn + 3) = "����" & Str_CX

Cells(Z_SZH_Qrow + 2, Z_SZH_QColumn + 3) = "���£�" & Replace(Split(Cells(13, YqTP(0)), " ")(1), "m", "")
Cells(Z_SZH_Qrow + 3, Z_SZH_QColumn + 3) = "��ѹ��" & Replace(Split(Cells(13, YqTP(1)), " ")(1), "m", "")
Cells(Z_SZH_Qrow - 1, Z_SZH_QColumn + 4) = "����ʱ��"
Range(Cells(Z_SZH_Qrow, Z_SZH_QColumn + 4), Cells(Z_SZH_Qrow + 3, Z_SZH_QColumn + 4)).Merge
Cells(Z_SZH_Qrow, Z_SZH_QColumn + 4) = Split(Range("a14"), " ")(0) & " �� " & Split(Cells(Snum + 13, 1), " ")(0)

'ͳ�Ƹ����ٲ㲻�������������
For i = 0 To UBound(Vspeed) - 2
For j = 1 + i To UBound(Vspeed) - 1
If CInt(HVgao(j)) = CInt(HVgao(i)) Then
Cells(V_X_Qrow + i, V_X_Qcolumn + j).Value = ""
Else
Cells(V_X_Qrow + i, V_X_Qcolumn + j).Value = �������������Լ��(Vspeed(i), Vspeed(j), (CInt(HVgao(i)) - CInt(HVgao(j))) \ 10, Snum)
End If
Next
Next
'���Ʊ��
For j = 0 To UBound(Vspeed)
For i = 0 To UBound(Vspeed) - 1
If j = 0 Then
Cells(V_X_Qrow - 1 + j, i + V_X_Qcolumn).Value = Yqm(i)
ElseIf i < j Then
Cells(V_X_Qrow - 1 + j, i + V_X_Qcolumn - 1).Value = Cells(V_X_Qrow - 1 + i, j + V_X_Qcolumn - 1).Value
Else
End If
Next
Next
'ͳ�Ƹ����ٲ㲻�������������
For i = 0 To UBound(Vspeed) - 1
Cells(V_X_Qrow + i, 5 + UBound(Vspeed)).Value = ͳ���в���������(Vspeed(i), Snum)
Cells(V_X_Qrow + i, 6 + UBound(Vspeed)).Value = ͳ����ʵ������(Vspeed(i), Snum)
Next
Cells(V_X_Qrow - 1, 5 + UBound(Vspeed)) = "��������Բ��������ݺϼ�"
Cells(V_X_Qrow - 2, 6 + UBound(Vspeed)) = "����Ӧ�����ݺϼ�"
Cells(V_X_Qrow - 1, 6 + UBound(Vspeed)) = "ͳ����ʵ������"
Cells(V_X_Qrow - 1, 7 + UBound(Vspeed)) = "ʵ������������"
Cells(V_X_Qrow - 1, 8 + UBound(Vspeed)) = "ȱ������"
'ͳ�Ƹ����ٲ�ʵ������������
Cells(V_X_Qrow - 2, 7 + UBound(Vspeed)).Value = Snum
For i = 0 To UBound(Vspeed) - 1
Cells(V_X_Qrow + i, 7 + UBound(Vspeed)).Value = Cells(V_X_Qrow + i, 6 + UBound(Vspeed)).Value / Snum
Cells(V_X_Qrow + i, 8 + UBound(Vspeed)).Value = Snum - Cells(V_X_Qrow + i, 6 + UBound(Vspeed)).Value
Next

'ͳ�Ƹ�����㲻�������������
For i = 0 To UBound(Ddirect) - 2
For j = 1 + i To UBound(Ddirect) - 1
If HDgao(j) = HDgao(i) Then
Cells(D_X_Qrow + i, D_X_Qcolumn + j).Value = ""
Else
Cells(D_X_Qrow + i, D_X_Qcolumn + j).Value = �������������Լ��(Ddirect(i), Ddirect(j), (HDgao(i) - HDgao(j)) * 1.125, Snum)
End If
Next
Next
'���Ʊ��
For j = 0 To UBound(Ddirect)
For i = 0 To UBound(Ddirect) - 1
If j = 0 Then
Cells(D_X_Qrow - 1 + j, i + D_X_Qcolumn).Value = HDgao(i) & "D"
ElseIf i < j Then
Cells(D_X_Qrow - 1 + j, i + D_X_Qcolumn - 1).Value = Cells(D_X_Qrow - 1 + i, j + D_X_Qcolumn - 1).Value
Else
End If
Next
Next
'ͳ�Ƹ�����㲻�������������
For i = 0 To UBound(Ddirect) - 1
Cells(D_X_Qrow + i, D_X_Qcolumn + UBound(Ddirect)).Value = ͳ���в���������(Ddirect(i), Snum)
Next
Cells(D_X_Qrow - 1, D_X_Qcolumn + UBound(Ddirect)) = "��������Բ��������ݺϼ�"

'ͳ�Ƹ�����ټƲ�����仯��������
For i = 0 To UBound(Vspeed) - 1

Cells(V_Q_Qrow + i + 1, V_Q_Qcolumn + 1).Value = ��������ٱ仯���Ƽ��(Vspeed(i), 6, Snum)

Next
'���Ʊ��
For i = 0 To UBound(Vspeed) - 1
Cells(V_Q_Qrow + i + 1, V_Q_Qcolumn).Value = Yqm(i)
Next
Cells(V_Q_Qrow, V_Q_Qcolumn + 1) = "���ٱ仯���Ʋ��������ݺϼ�"

'ͳ�Ƹ��¶ȼ���ѹ�Ʋ�����仯��������
For i = 0 To UBound(YqTP) - 1

Cells(Y_Q_Qrow + i + 1, Y_Q_Qcolumn + 1).Value = �����������ֵ���(YqTP(i), Cha(0), Snum)

Next
'���Ʊ��
For i = 0 To UBound(YqTP)
Cells(Y_Q_Qrow + i + 1, Y_Q_Qcolumn).Value = YqTPm(i)
Next
Cells(Y_Q_Qrow, Y_Q_Qcolumn + 1) = "�¶���ѹ�仯���Ʋ��������ݺϼ�"

'ͳ�Ƹ�����������仯��������

For i = 0 To YqZ - 1
col = i * 4 + 2
Cells(Z_F_Qrow + 1, Z_F_Qcolumn + i + 1).Value = ���������귶Χ���(col, Fanw(i, 0), Fanw(i, 1), Snum)

Next
'���Ʊ��
For i = 0 To YqZ - 1
Cells(Z_F_Qrow, Z_F_Qcolumn + i + 1).Value = Yqm(i)
Next
Cells(Z_F_Qrow + 1, Z_F_Qcolumn) = "�������ݷ�Χ������ϼ�"

'ͳ�Ƹ�������ټƱ�������

For i = 0 To Yq_VD_Z - 1
col = i * NumYK + 2
Cells(Z_B_Qrow + i + 1, Z_B_Column + 1).Value = ����������������(col, Tcol, Snum)
Cells(Z_B_Qrow + i + 1, Z_B_Column + 2).Value = Round(Cells(Z_B_Qrow + i + 1, Z_B_Column + 1).Value / Snum * 100, 2)
Next
'���Ʊ��
For i = 0 To Yq_VD_Z - 1
Cells(Z_B_Qrow + i + 1, Z_B_Column).Value = Yqm(i)
Next
Cells(Z_B_Qrow, Z_B_Column + 1) = "�������ݺϼ�"
Cells(Z_B_Qrow, Z_B_Column + 2) = "��������ռ�ȣ�%��"

'�����������Χ����������
For i = 0 To YqZ - 1
col = i * NumYK + 2
    For j = 14 To Snum + 12
    If Cells(j, col).style = "ע��" Or Cells(j, col).style = "����" Then
    Range(Cells(j, col), Cells(j, col + 3)).ClearContents
    End If
    Next
Next

'ͳ�Ƹ����ټ������ϵ��
For i = 0 To UBound(Vspeed) - 1
For j = 1 + i To UBound(Vspeed) - 1

Cells(V_XX_Qrow + j, V_XX_Qcolumn + i + 1).Value = Round(Application.Correl(Range(Cells(14, Vspeed(i)), Cells(14 + 8783, Vspeed(i))), Range(Cells(14, Vspeed(j)), Cells(14 + 8783, Vspeed(j)))), 3)

Next
Next
'���Ʊ��
For i = 0 To UBound(Vspeed) - 1
For j = 0 To UBound(Vspeed) - 1
If i = 0 Then
Cells(V_XX_Qrow, V_XX_Qcolumn + 1 + j).Value = Yqm(j)
ElseIf j = 0 Then
Cells(V_XX_Qrow + i, V_XX_Qcolumn).Value = Yqm(i)
ElseIf i <= j Then
Cells(V_XX_Qrow + i, V_XX_Qcolumn + 1 + j).Value = "/"
End If
Next
Next
Cells(V_XX_Qrow, V_XX_Qcolumn).Value = "�߶�"


Application.ScreenUpdating = True
End Sub
Function �������������Լ��(V1 As Integer, V2 As Integer, Cha As Integer, Snum As Long)
Dim n As Integer, Row As Integer, Col1 As Integer, Col2 As Integer 'ͳ�����߶ȷ��ٲ�����������
Row = 14
Col1 = V1
Col2 = V2
n = 0
For i = 0 To Snum - 1
If ��������Լ��(Cells(Row + i, Col1).Value, Cells(Row + i, Col2).Value, Cha) = 1 Then
    If Cells(Row + i, Col2).Value < 0.5 Then
    Else
    Cells(Row + i, Col1).style = "��"
    End If
    If Cells(Row + i, Col1).Value < 0.5 Then
    Else
     Cells(Row + i, Col2).style = "��"
     End If
 Cells(Row + i, 3).style = "��"
 n = n + 1
End If
Next
�������������Լ�� = n
End Function
Function ��������Լ��(V1 As Single, V2 As Single, Cha As Integer)
If Abs(V1 - V2) >= Cha And V1 <> 0 And V2 <> 0 Then
��������Լ�� = 1
Else
��������Լ�� = 0
End If
End Function
Function ͳ���в���������(col As Integer, Snum As Long)
n = 0
For i = 14 To 13 + Snum
If Cells(i, col).style = "��" Then
n = n + 1
End If
Next
ͳ���в��������� = n
End Function
Function ͳ����ʵ������(col As Integer, Snum As Long)
n = 0
For i = 14 To 13 + Snum
If Cells(i, col).Value <> "" Then
n = n + 1
End If
Next
ͳ����ʵ������ = n
End Function

Function �������������Լ��(V1 As Integer, V2 As Integer, Cha As Single, Snum As Long)
Dim n As Integer, Row As Integer, Col1 As Integer, Col2 As Integer 'ͳ�����߶ȷ��򲻺���������
Row = 14
Col1 = V1
Col2 = V2
n = 0
For i = 0 To Snum - 1
If ��������Լ��(Cells(Row + i, Col1).Value, Cells(Row + i, Col2).Value, Cha) = 1 Then
   If Cells(Row + i, Col2 + 1).Value < 0.4 And Cells(Row + i, Col1 + 1).Value >= 0.4 Then
   Else
    Cells(Row + i, Col1).style = "��"
    End If
   If Cells(Row + i, Col1 + 1).Value < 0.4 And Cells(Row + i, Col2 + 1).Value >= 0.4 Then
   Else
     Cells(Row + i, Col2).style = "��"
     End If
 Cells(Row + i, 3).style = "��"
 n = n + 1
End If
Next
�������������Լ�� = n
End Function
Function ��������Լ��(V1 As Single, V2 As Single, Cha As Single)
If Abs(V1 - V2) >= Cha And V1 <> 0 And V2 <> 0 And Abs(V1 - V2) <= 360 - Cha Then
��������Լ�� = 1
Else
��������Լ�� = 0
End If
End Function
Function ��������ٱ仯���Ƽ��(V1 As Integer, Cha As Single, Snum As Long)
Dim n As Integer, Row As Integer, Col1 As Integer, Col2 As Integer 'ͳ�����߶ȷ��򲻺���������
Row = 14
Col1 = V1
'Col2 = V2
n = 0
For i = 0 To Snum - 2
If ���ٱ仯���Ƽ��(Cells(Row + i, Col1).Value, Cells(Row + i + 1, Col1).Value, Cha) = 1 Then
Cells(Row + i + 1, Col1).style = "����"
Cells(Row + i, Col1).style = "����"
Cells(Row + i + 1, 3).style = "��"
Cells(Row + i, 3).style = "��"
 n = n + 1
End If
Next
��������ٱ仯���Ƽ�� = n
End Function
Function ���ٱ仯���Ƽ��(V1 As Single, V2 As Single, Cha As Single)
If Abs(V1 - V2) >= Cha And V1 <> 0 And V2 <> 0 Then
���ٱ仯���Ƽ�� = 1
Else
���ٱ仯���Ƽ�� = 0
End If
End Function

Function �����������ֵ���(V1 As Integer, Cha As Single, Snum)
Dim n As Integer, Row As Integer, Col1 As Integer, Col2 As Integer
Row = 14
Col1 = V1
n = 0
For i = 0 To Snum - 2
If ��ֵ���(Cells(Row + i, Col1).Value, Cells(Row + i + 1, Col1).Value, Cha) = 1 Then
Cells(Row + i + 1, Col1).style = "����"
Cells(Row + i, Col1).style = "����"
Cells(Row + i + 1, 3).style = "��"
Cells(Row + i, 3).style = "��"
 n = n + 1
End If
Next
�����������ֵ��� = n
End Function
Function ��ֵ���(V1 As Single, V2 As Single, Cha As Single)
If Abs(V1 - V2) >= Cha And V1 <> 0 And V2 <> 0 Then
��ֵ��� = 1
Else
��ֵ��� = 0
End If
End Function

Function ���������귶Χ���(col As Integer, Low As Single, Top As Single, Snum As Long)
Dim n As Integer, Row As Integer
Row = 14
n = 0
For i = 0 To Snum - 1 '������������
If ��Χ���(Cells(Row + i, col).Value, Low, Top) = 1 And Cells(Row + i, col).Value <> "" Then
Cells(Row + i, col).style = "����"
Cells(Row + i, 3).style = "��"
 n = n + 1
End If
Next
���������귶Χ��� = n
End Function
Function ��Χ���(Shu As Single, Low As Single, Top As Single) '����Χ���ڵ�����Сֵ��С�����ֵ
If Shu >= Top Or Shu < Low Then
��Χ��� = 1
Else
��Χ��� = 0
End If
End Function

Function ����������������(col As Integer, ColT As Integer, Snum)
Dim n As Integer, Row As Integer
Row = 14
n = 0
For i = 0 To Snum - 1 '������������
If �������(Cells(Row + i, col + 1).Value, Cells(Row + i, ColT).Value) = 1 And Cells(Row + i, col).Value <> "" And (Cells(Row + i, col).Value = Cells(Row + i + 1, col).Value Or Cells(Row + i, col).Value = Cells(Row + i - 1, col).Value) Then
Cells(Row + i, col).style = "ע��"
Cells(Row + i, 3).style = "��"
 n = n + 1
End If
Next
���������������� = n
End Function
Function �������(Sd As Double, Tem As Double) '�¶ȵ���0.5����׼ƫ�����0
If Sd < 0.2 And Tem <= 0.7 Then
������� = 1
Else
������� = 0
End If
End Function
Function dtodfm(du As Single)
dtodfm = Int(du) & "��" & Int((du - Int(du)) * 60) & Chr(39) & ((du - Int(du)) * 60 - Int((du - Int(du)) * 60)) * 60 & Chr(34)
End Function

