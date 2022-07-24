Public Declare Function GetTickCount64 Lib "kernel32.dll" () As Long

sub getBaesick()
  TotalLen = Int(Range("B3:B3").Text)
  PassLen = Int(Range("C3:C3").Text)
  BaesickLen = Int(Range("D3:D3").Text)
  
  if PassLen + BaesickLen >= TotalLen Then
    MsgBox("가용 인원이 없습니다.")
  Else
    Do While 1
      ' 자정으로부터 지난 초를 의미하는 Timer 값으로 seed 값 지정 (소수점 2번쨰 자리로 계속 바뀌는 값)
      ' 직접 바꾸지 않으면 seed 값이 고정되어 있어 똑같은 결과가 나오므로 주기적으로 변경
      Randomize GetTickCount64() + Timer
      
      ' randomRnd = Rnd()*1000 mod (TotalLen + 1) [1]
      
      ' Int( ( upperbound - lowerbound + 1 ) * Rnd + lowerbound )
      randomRnd = Int((TotalLen) * Rnd() + 1) ' [2]
      
      ' 이름 명단에서 랜덤으로 하나 선택
      randomValue = Range("B6:B" + Str(TotalLen + 5)).Item(randomRnd, 1)
      
      For i = 1 To PassLen
        ' 열외에 있는 이름인지 확인
        if Range("C6:C" + Str(PassLen + 5)).Item(i, 1) = randomValue Then
          checkValue = 1
        End if
      Next
      
      For i = 1 To BaesickLen
        ' 배식조 뽑힌 인원에 있는 이름인지 확인
        if Range("D6:D" + Str(BaesickLen + 5)).Item(i, 1) = randomValue Then
          checkValue = 1
        End if 
      Next
      
      ' 둘 다 없으면 출력
      if checkValue = 0 Then
        With Range("H12:H12")
          .Value = randomValue
          .Font.Size = 25
          .HorizontalAlignment = xlHAlignCenter
          .Select
        End With
        
        Range("D" + Str(6 + BaesickLen) + ":D" + Str(6 + BaesickLen)).Value = randomValue
          
        ' 뽑힌 순서 확인 및 랜덤 난수 테스트
        Range("E" + Str(5 + randomRnd) + ":E" + Str(5 + randomRnd)).Value = BaesickLen + 1
        Exit Do
      End if
          
      ' 변수를 재사용하기 위하여 0으로 초기화
      checkValue = 0
    Loop
  End If
End Sub
