<%
  nLength = Request.TotalBytes
	ByteArray = Request.BinaryRead( nLength )

  Response.Write nLength & " �o�C�g�̃f�[�^���ǂݍ��܂�܂���......<br>"

  ' �t�@�C���T�C�Y�ŃG���[����
  If nLength > 7000 Then
    Response.Write "�t�@�C���T�C�Y���傫�����܂�"
    Response.End
  End If

  Set basp = Server.CreateObject("basp21")

  ' �t�@�C�����擾
  file = basp.FormFileName(ByteArray, "fileinput")
  filename = Mid(file, InstrRev(file,"\")+1)
  Response.write(file)
  Response.write(filename)
  ' �t�@�C���ۑ�
  nRet = basp.FormSaveAs(ByteArray, "fileinput", Server.MapPath("uploads/" & filename))

  ' �G���[����
  Select Case nRet
		Case -1
			Response.Write "�n���ꂽ�p�����[�^�́A�P�o�C�g�̔z��ł͂���܂���<br>"
		Case -2
			Response.Write "�t�H�[���̒��ɖ��O������܂���<br>"
		Case -3
			Response.Write "�t�@�C�����쐬�ł��܂���<br>"
		Case -4
			Response.Write "�t�@�C���̏����݂����s���܂���<br>"
		Case Else
			Response.Write nRet & " �o�C�g�̃t�@�C�����쐬���܂���<br>"
			Response.Write "�ǂݍ��܂ꂽ�t�@�C�����́A" & filename & "�ł�<br>"
	End Select
%>