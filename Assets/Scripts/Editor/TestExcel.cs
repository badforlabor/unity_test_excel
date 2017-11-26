using System.Collections;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using UnityEngine;


public class TestExcel
{
	public static void Test1()
	{
		var fi = new FileInfo(@"c:\workbooks\myworkbook.xlsx");
		using (var p = new ExcelPackage(fi))
		{
			//Get the Worksheet created in the previous codesample. 
			var ws=p.Workbook.Worksheets["MySheet"];
			//Set the cell value using row and column.
				ws.Cells[2, 1].Value = "This is cell B1. It is set to bolds";
			//The style object is used to access most cells formatting and styles.
			ws.Cells[2, 1].Style.Font.Bold=true;
			//Save and close the package.
			p.Save();
		}
	}
}
