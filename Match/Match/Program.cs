using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;
using Excel=Microsoft.Office.Interop.Excel;
using System.IO;
using static LogSystem.LogSystem;

namespace Match
{
    public class Logfile:IComparable
    {
        private string type;
        private UInt64 counts;

        public string Type { get => type; set => type = value; }
        public ulong Counts { get => counts; set => counts = value; }
        public int CompareTo(object o)
        {
            if (o is Logfile)
            {
                Logfile OtherLogfile = o as Logfile;
                return Convert.ToInt32(Type.CompareTo(OtherLogfile.Type));
            }
            else
            {
                throw new ArgumentException("排序错误");
            }
        }
        public Logfile()
        {
            Type = string.Empty;
            Counts = new UInt64();
        }

    }
    public class Program
    {
        static string FileOut = @"D:\all_countOut.xls";
        static string FileIn = @"D:\all_count.xls";
        private static int ConTroNum = 6;
        private static int DifferNum = 10;

        public static int ConTroNum1 { get => ConTroNum; set => ConTroNum = value; }
        public static int DifferNum1 { get => DifferNum; set => DifferNum = value; }

        //static void Main(string[] args)
        //{
        //    StartDeel();
        //}
        public static void StartDeel()
        {
            //try
            //{
            List<Logfile> MyLogList = FromExcel(FileIn);
            DeelRepetion(ref MyLogList,ConTroNum1);
            SortList(ref MyLogList);
            OutToExcel(MyLogList);
            //}
            //catch(Exception excep)
            //{
            //    
            //};
            
        }
        public static void ModifyPathIn(string PathIn)
        {
            FileIn = PathIn;
        }
        public static void ModifyPathOut(string PathOut)
        {
            FileOut = PathOut;
        }

        public static void SortList(ref List<Logfile> MyLogList)
        {
            MyLogList.Sort();

        }
        //从Excel表查询，这里使用的是LinqToExcel
        public static List<Logfile> FromExcel(string pathfile)
        {
            var excel = new ExcelQueryFactory(pathfile);
            excel.AddMapping<Logfile>(x => x.Type, "Type");//添加映射
            excel.AddMapping<Logfile>(x => x.Counts, "Counts");
            //Counts Report是sheet的名字，必须对应
            var worksheet = excel.Worksheet<Logfile>("Counts Report").ToList();
            return worksheet;
        }
        //导出到Excel表，使用的Excel=Microsoft.Office.Interop.Excel
        public static void OutToExcel(List<Logfile> worksheet)
        {
            Excel.Application eApp = new Excel.Application();
            Excel.Workbooks workbooks = eApp.Workbooks;
            Excel.Workbook workbook = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet WorksheetOut = (Excel.Worksheet)workbook.Worksheets[1];
            //添加元素
            for (int i = 0; i < worksheet.Count; i++)
            {
                WorksheetOut.Cells[i + 1, 1] = worksheet[i].Type;
                WorksheetOut.Cells[i + 1, 2] = worksheet[i].Counts.ToString();
            }
            //设置格式
            Excel.Range range = WorksheetOut.Range[WorksheetOut.Cells[1, 1], WorksheetOut.Cells[worksheet.Count, 2]];
            Excel.Range AllColumn = WorksheetOut.Columns;
            AllColumn.Columns.ColumnWidth = 100;
            Excel.Range AllRow = WorksheetOut.Rows;
            AllRow.Rows.RowHeight = 13.5;

            eApp.Visible = true;
            if (File.Exists(FileOut))
            {
                try
                {
                    File.Delete(FileOut);
                }
                catch (IOException e)
                {
                    Console.WriteLine($"{e.Message}");
                }
            }
            try
            {
                workbook.SaveAs(FileOut);
            }
            catch(Exception e)
            {
                LOG(NOW, 2, e.Message, "导出excel失败");
            }
        }




        public static void DeelRepetion(ref List<Logfile> MyLogList,int NumContr)
        {
            for (int i = 0; i < MyLogList.Count; i++)
            {
                for (int j = i + 1; j < MyLogList.Count; j++)
                {
                    if (MyLogList[i] != null && MyLogList[j] != null)
                        if (CompareString(MyLogList[i].Type, MyLogList[j].Type,NumContr,DifferNum1))
                        {
                            MyLogList[i].Counts += MyLogList[j].Counts;
                            MyLogList[j] = null;

                        }
                }
            }
            MyLogList.RemoveAll(e => e == null);
        }

        public static bool CompareStringByRatio(string str1, string str2, double Ratio,int DifferNum)
        {
            if ((str1.Length - str2.Length) > DifferNum)
                return false;
            return CountHammingDitance(str1, str2) < Ratio ;

        }
        public static bool CompareString(string str1,string str2,int ConctrNum,int DifferNum)
        {
            if ((str1.Length - str2.Length) > DifferNum)
                return false;
            return CountStringHamming(str1, str2) < ConctrNum;
        }
        public static double CountHammingDitance(string str1,string str2)
        {
            double DifNumL = 0;
            double DifNumR = 0;
            double DifNumResult = 0;
            int str1HalfLen = str1.Length / 2;
            int str2HalfLen = str2.Length / 2;

            int MidNum = Math.Min(str1HalfLen, str2HalfLen);

            //int MinLen = Math.Min();
            for (int i = 0; i < MidNum; i++)
            {
                if (str1[i] != str2[i])
                    DifNumL++;
            }
            for (int j = str1.Length - 1, k = str2.Length - 1; j > MidNum && k > MidNum; j--, k--)
            {
                if (str1[j] != str2[k])
                {
                    DifNumR++;
                }
            }
            DifNumResult = Math.Min(DifNumL, DifNumR)/MidNum;
            //Console.WriteLine(DifNumResult);
            return DifNumResult;
        }

        public static int CountStringHamming(string str1,string str2)
        {
            int DifNumL = 0;
            int DifNumR = 0;
            int DifNumResult = 0;
            int str1HalfLen = str1.Length / 2;
            int str2HalfLen = str2.Length / 2;

            int MidNum = Math.Min(str1HalfLen, str2HalfLen);

            //int MinLen = Math.Min();
            for(int i=0;i<MidNum;i++)
            {
                if (str1[i] != str2[i])
                    DifNumL++;
            }
            for (int j = str1.Length-1,k = str2.Length-1;j > MidNum && k > MidNum;j--,k--)
            {
                if(str1[j]!=str2[k])
                {
                    DifNumR++;
                }
            }
            DifNumResult = Math.Min(DifNumL, DifNumR);
            Console.WriteLine(DifNumResult);
            return DifNumResult;
        }
    }
}
