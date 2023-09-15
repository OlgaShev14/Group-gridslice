'Эта программа производит операцию "grid slice" вдоль заданной линии для всех grid файлов в указанной директории
' В коде необходимо указать название файла с линией среза, расширение файлов, с которых будет производится срез, и директорию с файлами
' Полученные файлы DAT сохраняются в той же директории под теми же именами, что и исходные файлы

' Если после запуска программы ничего не происходит, убедитесь, что указана верная директория

Sub Main
Dim SurferApp As Object
Set SurferApp = CreateObject("Surfer.Application")
SurferApp.Visible = True
Debug.Clear
BlankFile$ = GetFilePath( "Yaroslavl_Arkhangelsk.bln", "bln",path1, "Select BLN file") 'Задайте линию, вдоль которой снимаются значения
If BlankFile$ ="" Then End
''''''''''''' User Variables ''''''''''''''''''

file_extension	= "grd" 'Задайте расширение файлов, с которых снимаются значения
file_directory	= "\\tsclient\C\Users\o.shevaldysheva\Downloads\test_groupslice" 'Задайте директорию с исходными файлами


'''''''''''''''''''''''''''''''''''''''''''''''
	Set surf = CreateObject("surfer.application")
	surf.Visible = True 'Прогресс выполнения программы может быть виден на панели инструментов для каждого файла

	'Убедитесь, что расширение файла не имеет лишней точки и в пути к директории с файлами есть закрывающая "\"
	file_extension	= LCase(Right(file_extension,(Len(file_extension) - InStrRev(file_extension,"."))))
	If  Len(file_directory)-InStrRev(file_directory,"\") <> 0 Then file_directory = file_directory + "\"
	data_file = Dir( file_directory  + "*." + file_extension)
	On Error GoTo FileError
	While data_file <> ""

		'Название выходного файла
		OutDataFile$ = file_directory + Left(data_file, Len(data_file)-(Len(data_file)-InStrRev(data_file,".")+1) ) + ".dat"

		'Выполнение функции "slice"
		SurferApp.GridSlice(InGrid:=file_directory + data_file, BlankFile:=BlankFile$, OutDataFile:=OutDataFile$)

		Debug.Print data_file
		data_file = Dir() 'переход к следующему файлу
	Wend
	surf.Quit
	Exit Sub

	'Выводит сообщение об ошибке для каждого файла, с которого не был сделан срез
	FileError:
	Debug.Print  "Error:	" + data_file + "						" + Err.Description
	Resume Next
End Sub
