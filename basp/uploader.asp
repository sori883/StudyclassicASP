<%
  nLength = Request.TotalBytes
	ByteArray = Request.BinaryRead( nLength )

  Response.Write nLength & " バイトのデータが読み込まれました......<br>"

  ' ファイルサイズでエラー処理
  If nLength > 7000 Then
    Response.Write "ファイルサイズが大きすぎます"
    Response.End
  End If

  Set basp = Server.CreateObject("basp21")

  ' ファイル名取得
  file = basp.FormFileName(ByteArray, "fileinput")
  filename = Mid(file, InstrRev(file,"\")+1)
  Response.write(file)
  Response.write(filename)
  ' ファイル保存
  nRet = basp.FormSaveAs(ByteArray, "fileinput", Server.MapPath("uploads/" & filename))

  ' エラー処理
  Select Case nRet
		Case -1
			Response.Write "渡されたパラメータは、１バイトの配列ではありません<br>"
		Case -2
			Response.Write "フォームの中に名前がありません<br>"
		Case -3
			Response.Write "ファイルを作成できません<br>"
		Case -4
			Response.Write "ファイルの書込みが失敗しました<br>"
		Case Else
			Response.Write nRet & " バイトのファイルを作成しました<br>"
			Response.Write "読み込まれたファイル名は、" & filename & "です<br>"
	End Select
%>