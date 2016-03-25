using DoExcel.Model;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DoExcel.Helper
{
    class ExcelHelper
    {
        public void ReadFromExcelFile(string filePath)
        {
            IWorkbook wk = null;
            string extension = System.IO.Path.GetExtension(filePath);
            try
            {
                FileStream fs = File.OpenRead(filePath);
                //把XLS文件中数据写入wk中
                if(extension.Equals(".xls"))
                {
                    wk = new HSSFWorkbook(fs);
                }
                else
                {
                    wk = new XSSFWorkbook(fs);
                }

                fs.Close();

                ISheet sheet = wk.GetSheetAt(0);
                IRow row = sheet.GetRow(0);

                int offset = 0;
                for(int i = 0;i <= sheet.LastRowNum;i++)
                {
                    row = sheet.GetRow(i);
                    if(row != null)
                    {
                        for(int j = 0;j < row.LastCellNum; j++)
                        {
                            ICell cellValue = row.GetCell(j);
                            if(cellValue != null)
                            {
                                Console.Write(cellValue.ToString().ToString() + " / j = " + j);
                            }
                            else
                            {
                                Console.Write("j = " + j + " / ISNULL");
                            }
                        }
                        Console.WriteLine("\n");
                    }
                }

            }catch(Exception e){
                Console.WriteLine(e.Message);
            }
        }

        public void WriteToExcel(string filePath)
        {
            //创建工作薄  
            IWorkbook wb;
            string extension = System.IO.Path.GetExtension(filePath);
            //根据指定的文件格式创建对应的类
            if (extension.Equals(".xls"))
            {
                wb = new HSSFWorkbook();
            }
            else
            {
                wb = new XSSFWorkbook();
            }

            IFont normalFont = wb.CreateFont();//字体
            normalFont.FontName = "宋体";
            normalFont.FontHeight = 220.0;

            IFont headFont = wb.CreateFont();//字体
            headFont.FontName = "宋体";
            headFont.FontHeight = 400.0;
            headFont.Boldweight = (short)FontBoldWeight.Bold;

            IFont fontTitle = wb.CreateFont();//字体
            fontTitle.FontName = "宋体";
            fontTitle.FontHeight = 240.0;
            fontTitle.Boldweight = (short)FontBoldWeight.Bold;//字体

            ICellStyle style1 = wb.CreateCellStyle();//样式
            style1.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;//文字水平对齐方式
            style1.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直对齐方式
            style1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style1.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style1.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style1.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            style1.WrapText = true;//自动换行
            style1.SetFont(normalFont);

            ICellStyle headstyle = wb.CreateCellStyle();//样式
            headstyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;//文字水平对齐方式
            headstyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直对齐方式
                                                                                     //设置边框
            headstyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            headstyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            headstyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            headstyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            headstyle.WrapText = true;//自动换行
            headstyle.SetFont(headFont);

            ICellStyle style2 = wb.CreateCellStyle();//样式
            style2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;//文字水平对齐方式
            style2.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直对齐方式
                                                                                  //设置边框
            style2.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style2.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style2.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style2.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            style2.WrapText = true;//自动换行
            style2.SetFont(normalFont);

            ICellStyle dateStyle = wb.CreateCellStyle();//样式
            dateStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;//文字水平对齐方式
            dateStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直对齐方式
                                                                                     //设置数据显示格式
            IDataFormat dataFormatCustom = wb.CreateDataFormat();
            dateStyle.DataFormat = dataFormatCustom.GetFormat("yyyy-MM-dd HH:mm:ss");

            //
            ICellStyle styleTitle = wb.CreateCellStyle();//样式
            styleTitle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;//文字水平对齐方式
            styleTitle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直对齐方式
                                                                                      //设置边框
            styleTitle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            styleTitle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            styleTitle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            styleTitle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            styleTitle.WrapText = true;//自动换行

            styleTitle.SetFont(fontTitle);

            //创建一个表单
            ISheet sheet = wb.CreateSheet("Sheet0");
            //设置列宽
            int[] columnWidth = { 10, 10, 20, 10 };
            for (int i = 0; i < columnWidth.Length; i++)
            {
                //设置列宽度，256*字符数，因为单位是1/256个字符
                sheet.SetColumnWidth(i, 256 * 14);

            }


            //测试数据
            int rowCount = 16, columnCount = 6;
            object[,] data = {
        {"abcdef", "", "", "", "", ""},
        {"","", "", "", "", ""},
        {"创建日期：2016年01","", "", "", "", ""},
        {"基本信息","", "", "", "", ""},
        {"分支公司","", "保单号", "", "缴费金额", ""},
        {"销售渠道","", "生效日期", "", "缴费年限", ""},
        {"主要险种","", "投保人姓名", "", "投保人日期", ""},
        {"联系电话","", "联系地址", "", "", ""},
        {"首期信息","", "", "", "", ""},
        {"签单人员","", "签单率", "", "签单业务员在职情况", ""},
        {"销售网点","", "", "", "机构继续率", ""},
        {"售卖过程","", "", "", "", ""},
        {"客户购买原因","", "", "", "", ""},
        {"客户经济状况","", "", "", "", ""},
        {"续期服务轨迹","", "", "", "", ""},
        {"服务时间（年月日）","服务人员", "服务内容（拜访过程)", "服务效果（拜访结果）", "服务需求（资源需求）", "核查（人员及结果）"},
    };

            IRow row;
            ICell cell;

            for (int i = 0; i < rowCount; i++)
            {
                row = sheet.CreateRow(i);//创建第i行
                for (int j = 0; j < columnCount; j++)
                {
                    cell = row.CreateCell(j);//创建第j列
                    if(i == 2)
                    {
                        cell.CellStyle = style2;
                    }
                    else if (i == 0 || i == 1)
                    {
                        cell.CellStyle = headstyle;
                    }
                    else if( i == 3 || i== 14 || i == 8)
                    {
                        cell.CellStyle = styleTitle;
                    }
                    //else if( i == 10 || i == 11 || i == 12 || i == 13)
                    //{

                    //}
                    else
                    {
                        cell.CellStyle = style1;
                    }
                    //根据数据类型设置不同类型的cell
                    object obj = data[i, j];
                    SetCellValue(cell, data[i, j]);
                    //如果是日期，则设置日期显示的格式
                    if (obj.GetType() == typeof(DateTime))
                    {
                        cell.CellStyle = dateStyle;
                    }
                    //如果要根据内容自动调整列宽，需要先setCellValue再调用
                    //sheet.AutoSizeColumn(j);
                }
            }

            //合并单元格，如果要合并的单元格中都有数据，只会保留左上角的
            //CellRangeAddress(0, 2, 0, 0)，合并0-2行，0-0列的单元格
            CellRangeAddress region = new CellRangeAddress(0, 1, 0, 5);
            sheet.AddMergedRegion(region);

            CellRangeAddress region1 = new CellRangeAddress(2, 2, 0, 5);
            sheet.AddMergedRegion(region1);

            CellRangeAddress region2 = new CellRangeAddress(3, 3, 0, 5);
            sheet.AddMergedRegion(region2);

            CellRangeAddress region8 = new CellRangeAddress(8, 8, 0, 5);
            sheet.AddMergedRegion(region8);

            CellRangeAddress region14 = new CellRangeAddress(14, 14, 0, 5);
            sheet.AddMergedRegion(region14);

            CellRangeAddress region3 = new CellRangeAddress(10, 10, 1, 3);
            sheet.AddMergedRegion(region3);

            CellRangeAddress region4 = new CellRangeAddress(11, 11, 1, 5);
            sheet.AddMergedRegion(region4);

            CellRangeAddress region5 = new CellRangeAddress(12, 12, 1, 5);
            sheet.AddMergedRegion(region5);

            CellRangeAddress region6 = new CellRangeAddress(13, 13, 1, 5);
            sheet.AddMergedRegion(region6);

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

        public void CopyExcel(string filePath, Person person)
        {
            File.Copy("D:/c#project/vbProject/excel/model.xls", filePath);

            EditExcel(filePath, person);
        }

        public void EditExcel(string filePath, Person person)
        {
            //创建工作薄  
            IWorkbook wb;
            string extension = System.IO.Path.GetExtension(filePath);

            FileStream fs = File.OpenRead(filePath);
            //把XLS文件中数据写入wk中
            if (extension.Equals(".xls"))
            {
                wb = new HSSFWorkbook(fs);
            }
            else
            {
                wb = new XSSFWorkbook(fs);
            }



            ISheet sheet = wb.GetSheetAt(0);

            IRow row;
            ICell cell;

            int[] rowArr = { 1,2,3};
            int[] colArr = { 2,3,4};

            //foreach(int i in rowArr)
            //{
            row = sheet.GetRow(4);
            cell = row.GetCell(1);

            //}



            //object obj = data[i, j];
            SetCellValue(cell, "libin");
            //如果是日期，则设置日期显示的格式
            //if (obj.GetType() == typeof(DateTime))
            //{
            //    cell.CellStyle = dateStyle;
            //}

            try
            {
                FileStream file = File.OpenWrite(filePath);
                wb.Write(file);//向打开的这个Excel文件中写入表单并保存。  
                file.Close();
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
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
    }
}
