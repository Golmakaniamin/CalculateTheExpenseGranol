Attribute VB_Name = "KontrolGardeshmes"
Dim infogardesh(29, 1) As String

Public Sub amin1(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form9.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 + 1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub

Public Sub amin2(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form10.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 + 1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub

Public Sub amin3(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form11.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 + 1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub

Public Sub amin4(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form13.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 + 1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub


Public Sub amin5(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form1.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 + 1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub


Public Sub amin6(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form14.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 + 1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub




Public Sub amin7(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form16.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 + 1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub




Public Sub amin8(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form17.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 + 1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub



Public Sub amin9(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form18.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 + 1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub



Public Sub amin10(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form19.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 + 1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub




Public Sub amin11(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form20.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 + 1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub


Public Sub amin12(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form21.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 + 1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub



Public Sub amin13(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form22.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 + 1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub


