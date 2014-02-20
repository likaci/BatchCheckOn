;ComObjError(false)
goto SetType
SetType:
	FileSelectFile, FileList, M, , 选择要处理的表格, *.xls
	if FileList =
		Return
	A_NowTime := A_Now
	StringSplit,FileListArray,FileList,`n

	Loop % FileListArray0 - 1
	{
		B_Index := A_Index + 1
		T_Index := A_Index 
		ExcelFilePath := FileListArray1 "\" FileListArray%B_Index%
		TrayTip,% "开始处理" A_Index "/" FileListArray0 - 1,  % ExcelFilePath

		try	
		{
		excelS := ComObjCreate("Excel.Application")
		}
		catch e
		{
			MsgBox 系统中没有发现Excel, 程序将退出.
			MyExit()
		}
		try
		{
			excelS.Workbooks.Open(ExcelFilePath)
		}
		catch e
		{
			MsgBox 打开 %ExcelFilePath% 出错,文件损坏?
			MyExit()
		}
		excelS.Visible := false
		excels.Columns("D:D").NumberFormatLocal := "G/通用格式"

		excelTpath := A_ScriptDir "\考勤模板.xls"
		excelT := ComObjCreate("Excel.Application")
		try
		{
		excelT.Workbooks.Open(excelTpath)
		}
		catch e
		{
			MsgBox 模板文件丢失,请确保 考勤模板.xls 与 程序在同一文件夹下
			MyExit()
		}
		excelT.Visible := false

		dayPre := 0
		ManPre := ""
		Loop
		{
			if A_Index = 1
				continue
			cellMan := excelS.cells(A_Index,2)
			department := excelS.cells(2,1).Value
			Man := cellMan.Value
			if A_Index = 2
				ManPre := Man
			if Man != %ManPre%
			{
				TargetDir :=  FileListArray1 "\" department month "考勤记录"
				IfNotExist % TargetDir
					FileCreateDir, % TargetDir
				SaveASPathFull := TargetDir "\" manPre month "考勤记录.xls"

				excelT.cells(2,1).Value := "     部门：" department "                    姓名：" manPre "                     月份：" year "年" month "月"
				excelT.cells(48,1).Value := "   分管领导审核：                                                        " A_YYYY "年" A_MM "月" A_DD "日    "
				excelT.Sheets("模板").Name := ManPre
				try
				{
				excelT.ActiveWorkbook.SaveAS(SaveASPathFull)
				}
				catch e
				{
					excelT.ActiveWorkbook.Saved := 1
				}
				excelT.Quit
				If Man != 
				{
					manPre := man
					timeCount := 1
					excelT := ComObjCreate("Excel.Application")
					excelT.Workbooks.Open(excelTpath)
					excelT.Visible := False
					TrayTip,正在处理 %department%,%ManPre%
				}
			}
			if Man = 
				break	
			cellTime := excelS.cells(A_Index,4)
			date = 19000101000000
			date += cellTime.Value, days
			date += -2, days
			;month year
			year := SubStr(date,1,4)
			month := SubStr(date,5,2)
			day := SubStr(date,7,2)
			time := SubStr(date,9,4)
			if day != %dayPre%
				timeCount := 1
			else
				timeCount++
			if day <= 10
			{
				dayCount := day - 1
				dayColumn:= 2
			}
			if day > 10 && day <=20
			{
				dayCount := day -11
				dayColumn:= 5
			}
			if day > 20
			{
				dayCount := day -21
				dayColumn:= 8
			}
			cell3 := excelT.cells(3+dayCount*4+timeCount,dayColumn)
			cell3.Value := CellTime.Value
			dayPre := day
		}
		
		;excelS.ActiveWindow.View := 3
		;SaveASPathFull := SaveASPath . FileListArray%B_Index% . ".xlsx"
 		;excelS.ActiveWorkbook.SaveAs(SaveASPathFull,51)
		excelT.Quit
		excelS.ActiveWorkbook.Saved := 1
		excelS.Quit
	}
	MsgBox 1,处理完成,处理完成
return

MyExit()
{

	excelT.ActiveWorkbook.Saved := 1
	excelT.Quit
	excelS.ActiveWorkbook.Saved := 1
	excelS.Quit
	ExitApp
}
