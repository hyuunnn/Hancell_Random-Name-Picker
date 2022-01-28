sub getBaesick()
  'Randomize Rnd()
  
  TotalLen = Int(Range("B3:B3").Text)
  PassLen = Int(Range("C3:C3").Text)
  BaesickLen = Int(Range("D3:D3").Text)
  
  if PassLen + BaesickLen >= TotalLen Then
    MsgBox("가용 인원이 없습니다.")
  Else
    Do While 1
      ' randomRnd = Rnd()*1000 mod (TotalLen + 1) [1]
      ' Int( ( upperbound - lowerbound + 1 ) * Rnd + lowerbound )
      randomRnd = Int((TotalLen) * Rnd() + 1) ' [2]
      
      ' 0이면 다시 뽑게하기 (0은 5번째 라인[Name]을 선택한다. 6번째 라인 이후에 있는 데이터들을 사용해야함)
      if randomRnd = 0 Then
        checkValue = 1
      End if
      
      ' TotalArr 변수에서 랜덤으로 하나 선택
      randomValue = Range("B6:B" + Str(TotalLen + 5)).Item(randomRnd, 1)
      
      For i = 1 To PassLen
        ' PassArr에 있는 이름인지 확인
        if Range("C6:C" + Str(PassLen + 5)).Item(i, 1) = randomValue Then
          checkValue = 1
        End if
      Next
      
      For i = 1 To BaesickLen
        ' BaesickArr에 있는 이름인지 확인
        if Range("D6:D" + Str(BaesickLen + 5)).Item(i, 1) = randomValue Then
          checkValue = 1
        End if 
      Next
      
      ' PassArr, BaesickArr에 없으면 출력
      if checkValue = 0 Then
        With Range("H12:H12")
          .Value = randomValue
          .Font.Size = 25
          .HorizontalAlignment = xlHAlignCenter
          .Select
        End With
        
        Range("D" + Str(6 + BaesickLen) + ":D" + Str(6 + BaesickLen)).Value = randomValue
          
        ' 뽑힌 순서 확인 용도 (랜덤 난수 테스트용) - 필요 없는 기능이면 코드 앞에 rem 사용하면 된다.
        Range("E" + Str(5 + randomRnd) + ":E" + Str(5 + randomRnd)).Value = BaesickLen + 1
        Exit Do
      End if
          
      ' 변수를 재사용하기 위하여 0으로 초기화
      checkValue = 0
    Loop
  End If
End Sub
