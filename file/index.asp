<%@LANGUAGE=VBSCRIPT %>
<% 'On Error Resume Next %>
<html>
<head>
<meta charset="shift_jis">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<!-- <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8"> -->
</head>
<body>
<%
  Set objFso = Server.CreateObject("Scripting.FileSystemObject")
  
  ' ファイルの保存場所を作成

  folderpath = "C:\inetpub\wwwroot\file\testfiles"
  fullpath = objFso.BuildPath(folderpath, "test.csv")

  ' ファイルの存在確認
  If Not objFso.FileExists(fullpath) Then
    Response.Write "ファイルが存在しません"
    Response.End
  End If


  ' 第2パラメータ→ 1:読み取り専用、2:書き込み専用、8:ファイルの最後に書き込み
  ' 第3パラメータ→ True(規定値):新しいファイルを作成する、False:新しいファイルを作成しない
  ' 第4パラメータ→ 0(規定値):ASCII ファイルとして開く、-1:Unicode ファイルとして開く、-2:システムの既定値で開く
  Set objFile = objFso.OpenTextFile(fullpath, 1, False, -2)


  If Err.Number > 0 Then
    ' エラーがある場合
    Response.Write "Open Error"
  Else
    ' ファイルの最終行まで表示する
    Do Until objFile.AtEndOfStream
      Response.Write objFile.ReadLine & "<br>"
    Loop
  End If

  objFile.Close

 ' ファイルの内容を全て削除して書き込み
  Set objFileEdit = objFso.OpenTextFile(fullpath, 2, True, -2)
  If Err.Number > 0 Then
    Response.Write "Open Error"
  Else
      objFileEdit.writeLine "追加する文字列です。"
  End If

  objFileEdit.Close

  ' ファイルに文字を追加
  Set objFileAdd = objFso.OpenTextFile(fullpath, 8, True, -2)
  If Err.Number > 0 Then
    Response.Write "Open Error"
  Else
      objFileAdd.writeLine "追加する文字列です。"
  End If

  objFileAdd.Close
  
  Set objFile = Nothing
  Set objFso = Nothing
%>
</body>
</html>