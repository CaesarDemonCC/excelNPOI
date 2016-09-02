using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Collections;
using System.IO;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using System.Resources;
using System.Globalization;
using System.Reflection;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Vevisoft.Excel.Core;
namespace ExcelSX 
{
    static class Program 
    {
        static void Main()
        {
            Console.WriteLine("默认读取当前路径下的excel.xlsx文件");
            Console.WriteLine("请输入txt文件名，例如txt1");

            cc c=new cc();
            c.run(Console.ReadLine());
            Console.WriteLine("输入任意键退出...");
            Console.ReadLine();
        }
        
    }
    public class cc
    {
        public void run(string text)
        {
            var filename = "./"+ text +".txt";
            var importCore = new ExcelImportCore();  
            importCore.LoadFile("./excel.xlsx");  
            var ds = importCore.GetAllTables(false);  
            Console.WriteLine(ds.Tables[0].Rows.Count);
            
            var allTable = ds.Tables[0];
            DataTable outputTable = new DataTable();
            outputTable = allTable.Copy();
            outputTable.Rows.Clear();
            outputTable.ImportRow(allTable.Rows[0]);//这是加入的是第一行

            List<string> excelArr = new List<string>();
            for (var i = 0; i < ds.Tables[0].Rows.Count; i++ )
            {
                excelArr.Add(allTable.Rows[i][0].ToString());
            }
            Console.WriteLine("excel.xlsx文件总行数：" +excelArr.Count);

            List<string> txtArr = new List<string>();
            //从头到尾以流的方式读出文本文件
            //该方法会一行一行读出文本
            using (System.IO.StreamReader sr = new System.IO.StreamReader(filename))
            {
                string str;
                while ((str = sr.ReadLine()) != null)
                {
                    txtArr.Add(str);
                }
            }
            Console.WriteLine("txt文件总行数：" +txtArr.Count);

            List<string> r = txtArr.Intersect(excelArr).ToList();
            Console.WriteLine("相同ID的总行数：" + r.Count);
            //List<string> r2 = txtArr.Except(excelArr).ToList();
            //Console.WriteLine("不同的总行数：" + r2.Count);

            
            Console.WriteLine("根据相同ID开始查找...");
            for (var i = 0; i < allTable.Rows.Count; i++ )
            {
             if(r.Exists(p => p == allTable.Rows[i][0].ToString()))
                 outputTable.ImportRow(allTable.Rows[i]);
            }
            Console.WriteLine("根据相同ID查找的总行数：" + outputTable.Rows.Count);

            Console.WriteLine(" 开始写入新的excel，名称为：" + text + '-' + outputTable.Rows.Count + "-excel.xlsx");

             
            var exportCore = new ExportExcelCore();  
                exportCore.RenderDataTableToExcel("./"+text + '-' + outputTable.Rows.Count + "-excel.xlsx", outputTable, "1", null);  
                //exportCore.RenderDataTableToExcelHasTemplate(sdiag.FileName, file +"-excel.xlsx", dt, 3, 1);  
           
        }
    }
}