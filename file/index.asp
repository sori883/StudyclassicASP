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
  
  ' �t�@�C���̕ۑ��ꏊ���쐬

  folderpath = "C:\inetpub\wwwroot\file\testfiles"
  fullpath = objFso.BuildPath(folderpath, "test.csv")

  ' �t�@�C���̑��݊m�F
  If Not objFso.FileExists(fullpath) Then
    Response.Write "�t�@�C�������݂��܂���"
    Response.End
  End If


  ' ��2�p�����[�^�� 1:�ǂݎ���p�A2:�������ݐ�p�A8:�t�@�C���̍Ō�ɏ�������
  ' ��3�p�����[�^�� True(�K��l):�V�����t�@�C�����쐬����AFalse:�V�����t�@�C�����쐬���Ȃ�
  ' ��4�p�����[�^�� 0(�K��l):ASCII �t�@�C���Ƃ��ĊJ���A-1:Unicode �t�@�C���Ƃ��ĊJ���A-2:�V�X�e���̊���l�ŊJ��
  Set objFile = objFso.OpenTextFile(fullpath, 1, False, -2)


  If Err.Number > 0 Then
    ' �G���[������ꍇ
    Response.Write "Open Error"
  Else
    ' �t�@�C���̍ŏI�s�܂ŕ\������
    Do Until objFile.AtEndOfStream
      Response.Write objFile.ReadLine & "<br>"
    Loop
  End If

  objFile.Close

 ' �t�@�C���̓��e��S�č폜���ď�������
  Set objFileEdit = objFso.OpenTextFile(fullpath, 2, True, -2)
  If Err.Number > 0 Then
    Response.Write "Open Error"
  Else
      objFileEdit.writeLine "�ǉ����镶����ł��B"
  End If

  objFileEdit.Close

  ' �t�@�C���ɕ�����ǉ�
  Set objFileAdd = objFso.OpenTextFile(fullpath, 8, True, -2)
  If Err.Number > 0 Then
    Response.Write "Open Error"
  Else
      objFileAdd.writeLine "�ǉ����镶����ł��B"
  End If

  objFileAdd.Close
  
  Set objFile = Nothing
  Set objFso = Nothing
%>
</body>
</html>