using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System.IO;
using System;
using OfficeOpenXml;

[CustomEditor(typeof(TestExcel))]
public class TestExcelInspector : Editor
{
    public override void OnInspectorGUI()
    {
        base.OnInspectorGUI();

        if (GUILayout.Button("测试读取Excel"))
        {
            var path = Application.dataPath + "/../驿站.xlsx";

            var fi = new FileInfo(path);
            using (var p = new ExcelPackage(fi))
            {
                //Get the Worksheet created in the previous codesample. 
                var ws = p.Workbook.Worksheets["驿站信息"];
                int maxcol = 1000;   // 1000列
                int maxrow = 10000;   // 1000行

                bool err = false;
                int i = 1;
                for (i = 1; !err && i < maxrow; i++)
                {
                    string line = "";

                    for (int j = 1; j < maxcol; j++)
                    { 
                        try
                        {
                            string v = ws.Cells[i, j].Text;
                            if (v.Trim().Length == 0)
                            {
                                break;
                            }
                            else if (v.StartsWith("#"))
                            {
                                // 此行为注释！
                                line += v;
                                break;
                            }
                            else
                            {
                                line += v + ","; 
                            }
                        }
                        catch(Exception)
                        {
                            err = true;
                            break;
                        }
                    }

                    // 此行没数据，那么说明都处理完毕了。
                    if (line.Length == 0)
                        break;

                    Debug.logger.Log("此行数据：" + line);
                }
                Debug.logger.Log("一共读取的行数：" + (i-1));
            }   
        }

    }

    public static void Test1()
    {
        var fi = new FileInfo(@"c:\workbooks\myworkbook.xlsx");
        using (var p = new ExcelPackage(fi))
        {
            //Get the Worksheet created in the previous codesample. 
            var ws = p.Workbook.Worksheets["MySheet"];
            //Set the cell value using row and column.
            ws.Cells[2, 1].Value = "This is cell B1. It is set to bolds";
            //The style object is used to access most cells formatting and styles.
            ws.Cells[2, 1].Style.Font.Bold = true;
            //Save and close the package.
            p.Save();
        }
    }

}
