Sub 作成_Click()
  On Error GoTo Error
  Dim module As New ExcelComModule
  module.SetProviderName ("Kintone")
  module.SetConnectionString ("User=*****;Password=*****;Url=https://*****.cybozu.com")
  Cursor = Application.Cursor
  Application.Cursor = xlWait
  Dim nameArray, valueArray
  With Worksheets("テンプレート_Macro")
    '必須項目チェック
    If Cells(2, "F").Value = Empty Then
      Err.Description = "見積り番号をセットしてください"
      GoTo Error
    End If
    'セル初期化
    Cells(4, "M").MergeArea.ClearContents '見積番号
    Cells(5, "M").MergeArea.ClearContents '見積日
    Cells(8, "A").MergeArea.ClearContents '宛名
    Cells(30, "B").MergeArea.ClearContents '備考
    Cells(2, "M").MergeArea.ClearContents 'RecordId (非表示)
    For i = 18 To 27
      Cells(i, "B").MergeArea.ClearContents '見積明細(型番)
      Cells(i, "D").MergeArea.ClearContents '見積明細(商品名)
      Cells(i, "J").MergeArea.ClearContents '見積明細(単価)
      Cells(i, "L").MergeArea.ClearContents '見積明細(数量)
    Next i
    '見積書 取得
    Query = "SELECT * FROM 見積書 WHERE 見積番号 = '" & Range("F2").Value & "'"
    result = module.Select(Query, nameArray, valueArray)
    If Not module.EOF Then
      Cells(4, "M").Value = module.GetValue(3) '見積番号
      Cells(5, "M").Value = module.GetValue(13) '見積日
      Cells(8, "A").Value = module.GetValue(8) '宛名
      Cells(30, "B").Value = module.GetValue(11) '備考
      Cells(2, "M").Value = module.GetValue(0) 'RecordId (非表示)
    Else
      Err.Description = "見積り番号が見つかりませんでした"
      GoTo Error
    End If
    '見積明細 取得
    Query = "SELECT * FROM 見積書_見積明細 WHERE 見積書Id = '" & Range("O2").Value & "' ORDER BY Id"
    result = module.Select(Query, nameArray, valueArray)
    i = 18
    While (Not module.EOF)
      Cells(i, "B").Value = module.GetValue(8) '見積明細(型番)
      Cells(i, "D").Value = module.GetValue(7) '見積明細(商品名)
      Cells(i, "J").Value = module.GetValue(6) '見積明細(単価)
      Cells(i, "L").Value = module.GetValue(5) '見積明細(数量)
      module.MoveNext
      i = i + 1
    Wend
    MsgBox "完成"
  End With
  Application.Cursor = Cursor
  module.Close
  Exit Sub
Error:
  MsgBox "ERROR: " & Err.Description
  Application.Cursor = Cursor
  module.Close
End Sub
