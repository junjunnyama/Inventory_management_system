

Sub inputPic(ByVal target As Excel.Range, ByVal Tool As String)
    
    Dim myF As Variant      '写真のファイルアドレス
    Dim strD As String      '写真のディレクトリ
    Dim strF As String      'ディレクトリ内のファイル
    Dim text As String      '作業用レジスタ
    
    '=====画像取得=====
    'ファイルの場所取得
    strD = ActiveWorkbook.Path & "\image\"
    strF = Dir(strD)
    
    Do While strF <> ""
        text = strF
        text = WorksheetFunction.Asc(text)
        text = Replace(text, " ", "")
        If text Like Tool & ".jpg" Or text Like Tool & ".png" Or text Like Tool & ".JPG" Then
            myF = strD & strF
            Exit Do
        End If
        strF = Dir()
    Loop
    If myF = Empty Then
        target.Value = "対象の画像が存在しません。" & vbCrLf & "対象アイテム：" & Tool
        Exit Sub
    End If
    
    '=====画像添付=====
    '指定座標への画像添付
    With ActiveSheet.Shapes.AddPicture( _
        Filename:=myF, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        left:=target.left, _
        top:=target.top, _
        Width:=-1, _
        Height:=-1 _
        )
    
        '縦横比固定
        .LockAspectRatio = msoTrue
        
        '縮尺補正
        .ScaleHeight (target.Height - 3) / .Height, msoFalse
        If .Width > target.Width Then
            .ScaleWidth (target.Width - 3) / .Width, msoFalse
        End If
        
        '中央へ調整
        .top = target.top + target.Height / 2 - .Height / 2 + 0.5
        .left = target.left + target.Width / 2 - .Width / 2 + 0.5
        
    End With
    
End Sub