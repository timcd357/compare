using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Util;
using System.Diagnostics;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace Compare
{
    class Program
    {
        static void Main(string[] args)
        {
            //if (args.Length > 0)
            //{
            //    FileInfo file = new FileInfo(args[0]);
            //    Console.WriteLine(file.Directory);
            //    List<string> list = ReadFromExcelFile(file.FullName);
            List<string> list = ReadFromExcelFile("C:\\Users\\timcd\\Desktop\\1.XLS");
            Console.Write(list);
                if (list != null)
                {
                WriteToExcel("C:\\Users\\timcd\\Desktop\\2.xlsx", list);
                //WriteToExcel(file.Directory + "\\" + "2.xlsx", list);
                    Console.WriteLine("处理完成");
                }
                else
                {
                    Console.WriteLine("处理失败");
                }

                Console.ReadKey();
            //}
        }
        public static List<string> ReadFromExcelFile(string filePath)
        {
            IWorkbook wk = null;
            string extension = Path.GetExtension(filePath);
            try
            {
                FileStream fs = File.OpenRead(filePath);
                if (extension.Equals(".XLS"))
                {
                    //把xls文件中的数据写入wk中
                    wk = new HSSFWorkbook(fs);
                }
                else
                {
                    //把xlsx文件中的数据写入wk中
                    wk = new XSSFWorkbook(fs);
                }

                fs.Close();
                //读取当前表数据
                ISheet sheet = wk.GetSheetAt(0);

                IRow row = null;  //读取当前行数据

                List<string> list = new List<string>();
                //Console.WriteLine(sheet.LastRowNum);
                for (int i = 1; i <= sheet.LastRowNum; i++)  //LastRowNum 是当前表的总行数-1（注意）
                {
                    row = sheet.GetRow(i);  //读取当前行数据
                    if (row != null)
                    {
                        StringBuilder v = new StringBuilder("");
                        //LastCellNum 是当前行的总列数
                        if(!(row.GetCell(18) + "").ToString().Substring(0, 3).Equals("(ES"))
                        {
                            continue;
                        }
                        for (int j = 0; j < row.LastCellNum; j++)
                        {
                            //读取该行的第j列数据
                            string value = (row.GetCell(j)+"").ToString();                        
                            if (j == 9)
                            {
                                v.Append("厂商：" + value);
                                //Console.WriteLine(v);
                                continue;
                            }else if (j == 18)
                            {
                                v.Append( "|规格：" + value.Substring(value.IndexOf(")")+1));
                                //Console.WriteLine(v);
                                continue;
                            }else if (j == 24)
                            {
                                v.Append( "|数量：" + value);
                                list.Add(v.ToString());
                                Console.WriteLine(v);
                                break;
                            }
                           
                        }
                        Console.WriteLine("\n");
                    }
                    //Console.ReadKey();
                }
                //Console.WriteLine(list.Count);
                return list;
            }

            catch (Exception e)
            {
                //只在Debug模式下才输出
                Console.WriteLine(e.Message);
                return null;
            }
        }
        public static void WriteToExcel(string filePath, List<string> list)
        {
            //创建工作薄  
            IWorkbook wb;

            string extension = Path.GetExtension(filePath);
            //根据指定的文件格式创建对应的类
            if (extension.Equals(".xls"))
            {
                wb = new HSSFWorkbook();
            }
            else
            {
                wb = new XSSFWorkbook();
            }

            ICellStyle titleStyle = wb.CreateCellStyle();//样式
            IFont font1 = wb.CreateFont();//字体
            //font1.IsBold = true;
            font1.FontName = "宋体";
            font1.FontHeightInPoints = 22;
            titleStyle.SetFont(font1);
            titleStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;//文字水平对齐方式
            titleStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直对齐方式       

            ICellStyle commonStyle = wb.CreateCellStyle();//样式
            IFont font2 = wb.CreateFont();//字体
            font2.FontHeightInPoints = 12;
            font2.FontName = "宋体";
            commonStyle.SetFont(font2);//样式里的字体设置具体的字体样式                                       
            commonStyle.FillPattern = FillPattern.SolidForeground;
            commonStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;//文字水平对齐方式
            commonStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直对齐方式
            commonStyle.WrapText = true;

            ICellStyle tableStyle = wb.CreateCellStyle();//样式
            IFont font3 = wb.CreateFont();//字体
            font3.FontHeightInPoints = 10;
            font3.FontName = "宋体";
            tableStyle.SetFont(font3);//样式里的字体设置具体的字体样式                                       
            tableStyle.FillPattern = FillPattern.SolidForeground;
            tableStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;//文字水平对齐方式
            tableStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直对齐方式
            tableStyle.WrapText = true;
            tableStyle.BorderBottom = BorderStyle.Thin;
            tableStyle.BorderLeft = BorderStyle.Thin;
            tableStyle.BorderRight = BorderStyle.Thin;
            tableStyle.BorderTop = BorderStyle.Thin;

            ICellStyle dateStyle = wb.CreateCellStyle();//样式
            dateStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;//文字水平对齐方式
            dateStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直对齐方式
                                                                                     //设置数据显示格式
            IDataFormat dataFormatCustom = wb.CreateDataFormat();
            dateStyle.DataFormat = dataFormatCustom.GetFormat("yyyy-MM-dd");

            //创建一个表单
            ISheet sheet = wb.CreateSheet("Sheet0");
            //设置列宽
            int[] columnWidth = { 10, 10, 20, 10 };
            for (int i = 0; i < columnWidth.Length; i++)
            {
                //设置列宽度，256*字符数，因为单位是1/256个字符
                sheet.SetColumnWidth(i, 256 * columnWidth[i]);
            }

            //测试数据
            int columnCount = 1;

            object[,] data = {
        {"型号（只）"},
        {"81120"},
        {"81140"},
        {"81160"},
        {"81180"},
        {"81200"},
        {"81220"},
        {"81240"},
        {"81260"},
        {"81280"},
        {"81300"},
        {"81325"},
        {"81350"},
        {"81375"},
        {"81400"},
        {"801120"},
        {"801140"},
        {"801160"},
        {"801180"},
        {"801200"},
        {"801220"},
        {"801240"},
        {"801260"},
        {"801280"},
        {"801300"},
        {"801325"},
        {"801350"},
        {"801375"},
        {"801400"},
        {"28180"},
        {"28200"},
        {"28220"},
        {"28240"},
        {"28260"},
        {"28280"},
        {"28300"},
        {"28325"},
        {"28350"},
        {"28375"},
        {"28400"},
       
        };
            int rowCount = data.Length;
            IRow row;
            ICell cell;

            for (int i = 0; i < rowCount; i++)
            {
                row = sheet.CreateRow(i);//创建第i行
                for (int j = 0; j < columnCount; j++)
                {
                    cell = row.CreateCell(j);//创建第j列
                    //cell.CellStyle = j % 2 == 0 ? style1 : style2;
                    //根据数据类型设置不同类型的cell
                    object obj = data[i, j];
                    SetCellValue(cell, data[i, j]);
                    //如果是日期，则设置日期显示的格式
                    if (obj.GetType() == typeof(DateTime))
                    {
                        cell.CellStyle = dateStyle;
                    }
                    cell.CellStyle = tableStyle;
                    //如果要根据内容自动调整列宽，需要先setCellValue再调用
                    sheet.AutoSizeColumn(j);
                }
            }
            //string now = formDate();
            List<string> factoryList = new List<string>();
            for (int i = 0; i < list.Count; i++)
            {
                string s = list[i];
                string f = s.Substring(3, s.IndexOf("|")-3);//客户                         
                if (!factoryList.Contains(f))
                {
                    factoryList.Add(f);
                }                                        
            }

            for(int i = 0; i < factoryList.Count; i++)
            {
                string f = factoryList[i];
                row = sheet.GetRow(0);
                AddString(row.LastCellNum, row, f, tableStyle);
                //addString(row.LastCellNum + 4, row, s, tableStyle);
            }

            for (int i = 0; i < list.Count; i++)
            {
                string s = list[i];
                string f = s.Substring(3, s.IndexOf("|") - 3);//客户
                string t = s.Substring(s.IndexOf(")") + 1, s.LastIndexOf("|") - s.IndexOf(")") - 2);//型号
                int n = int.Parse(s.Substring(s.LastIndexOf("：") + 1));
                int fn = GetNumInList(factoryList, f);
                int tn = GetNumInArray(data, t);
                cell = sheet.GetRow(tn).GetCell(fn);
                Console.WriteLine(fn+"||"+tn);
                if (cell != null)
                {
                    SetCellValue(cell, n + int.Parse(cell.ToString()));
                }
                else
                {
                    //SetCellValue(cell, n);
                    sheet.GetRow(tn).CreateCell(fn).SetCellValue(n);
                }
                
                sheet.AutoSizeColumn(i);
            }

            try
            {
                FileStream fs = File.OpenWrite(filePath);
                wb.Write(fs);//向打开的这个Excel文件中写入表单并保存。  
                fs.Close();
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }
        public static string FormDate()
        {
            DateTime currentTime;
            //  string strYMD = currentTime.ToString("y");
            string strDate;
            //   strYMD = System.DateTime.Now.ToString();//2019/5/29 10:14:39
            //  strYMD = System.DateTime.Now.ToString("yyyyMMddhhmmss");//20190529101211
            //  strYMD = System.DateTime.Now.ToString("d");//显示格式：2019/5/29
            currentTime = DateTime.Now;
            strDate = currentTime.ToString("yyyy-MM-dd");

            return strDate;
        }
        private static void AddString(int col, IRow _row,  object s, ICellStyle style)
        {
            ICell cell;
            Console.WriteLine(s);
            cell = _row.CreateCell(col);//创建第j列       

            cell.CellStyle = style;
            SetCellValue(cell, s);

        }
        public static void SetCellValue(ICell cell, object obj)
        {
            if (obj.GetType() == typeof(int))
            {
                cell.SetCellValue((int)obj);
            }
            else if (obj.GetType() == typeof(double))
            {
                cell.SetCellValue((double)obj);
            }
            else if (obj.GetType() == typeof(IRichTextString))
            {
                cell.SetCellValue((IRichTextString)obj);
            }
            else if (obj.GetType() == typeof(string))
            {
                cell.SetCellValue(obj.ToString());
            }
            else if (obj.GetType() == typeof(DateTime))
            {
                cell.SetCellValue((DateTime)obj);
            }
            else if (obj.GetType() == typeof(bool))
            {
                cell.SetCellValue((bool)obj);
            }
            else
            {
                cell.SetCellValue(obj.ToString());
            }
        }

        public static int GetNumInList(List<string> l,string s)
        {
            int i = 0;
            foreach(string ss in l)
            {
                i++;
                if (s.Equals(ss))
                {
                    break;
                }
            }
            return i;
        }

        public static int GetNumInArray(Array l, string s)
        {
            int i = 0;
            foreach (string ss in l)
            {
                i++;
                if (s.Equals(ss))
                {
                    i--;
                    break;
                }
            }
            return i;
        }
    }
}
