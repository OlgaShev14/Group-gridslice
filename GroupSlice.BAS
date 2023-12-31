'��� ��������� ���������� �������� "grid slice" ����� �������� ����� ��� ���� grid ������ � ��������� ����������
' � ���� ���������� ������� �������� ����� � ������ �����, ���������� ������, � ������� ����� ������������ ����, � ���������� � �������
' ���������� ����� DAT ����������� � ��� �� ���������� ��� ���� �� �������, ��� � �������� �����

' ���� ����� ������� ��������� ������ �� ����������, ���������, ��� ������� ������ ����������

Sub Main
Dim SurferApp As Object
Set SurferApp = CreateObject("Surfer.Application")
SurferApp.Visible = True
Debug.Clear
BlankFile$ = GetFilePath( "Yaroslavl_Arkhangelsk.bln", "bln",path1, "Select BLN file") '������� �����, ����� ������� ��������� ��������
If BlankFile$ ="" Then End
''''''''''''' User Variables ''''''''''''''''''

file_extension	= "grd" '������� ���������� ������, � ������� ��������� ��������
file_directory	= "\\tsclient\C\Users\o.shevaldysheva\Downloads\test_groupslice" '������� ���������� � ��������� �������


'''''''''''''''''''''''''''''''''''''''''''''''
	Set surf = CreateObject("surfer.application")
	surf.Visible = True '�������� ���������� ��������� ����� ���� ����� �� ������ ������������ ��� ������� �����

	'���������, ��� ���������� ����� �� ����� ������ ����� � � ���� � ���������� � ������� ���� ����������� "\"
	file_extension	= LCase(Right(file_extension,(Len(file_extension) - InStrRev(file_extension,"."))))
	If  Len(file_directory)-InStrRev(file_directory,"\") <> 0 Then file_directory = file_directory + "\"
	data_file = Dir( file_directory  + "*." + file_extension)
	On Error GoTo FileError
	While data_file <> ""

		'�������� ��������� �����
		OutDataFile$ = file_directory + Left(data_file, Len(data_file)-(Len(data_file)-InStrRev(data_file,".")+1) ) + ".dat"

		'���������� ������� "slice"
		SurferApp.GridSlice(InGrid:=file_directory + data_file, BlankFile:=BlankFile$, OutDataFile:=OutDataFile$)

		Debug.Print data_file
		data_file = Dir() '������� � ���������� �����
	Wend
	surf.Quit
	Exit Sub

	'������� ��������� �� ������ ��� ������� �����, � �������� �� ��� ������ ����
	FileError:
	Debug.Print  "Error:	" + data_file + "						" + Err.Description
	Resume Next
End Sub
