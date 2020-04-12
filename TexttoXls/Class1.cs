/*
 litao 20200410
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Text.RegularExpressions;
//using ZXing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NPOI.CSS;
using NPOI.HSSF.Util;
using Microsoft.Office.Interop;
using NPOI.SS.Converter;
using System.Text;

//操作json


using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Converters;
using NPOI.XSSF.UserModel;

using NPOI.HSSF.Record;
namespace TexttoXls
{
    [Guid("FD22E8C5-07A7-4DCC-AF93-5B33E867AF6A")]

    public interface IConvertXls
    {
        [DispId(1)]
        void openxls(string xls);
        void closexls();
        void CreateSheet(int nums);
        string GetCell(int k, int mrow, int mcol);
        void InsertCell(int k, int mrow, int mcol, string text);
        void InsertNumCell(int k, int mrow, int mcol, string text);
        string GetSheetNums();
        void RemoveOneRow(int k, int mrow);
        void RemoveOneCol(int k, int mrow, int mcol);
        void InsertRow(int k, int mrow);
        void InsertPicture(int k, int startrow, int startcol, int lastrow, int lastcol, string PicturePath);
        void HideCol(int k, int mcol, bool ishide);
        void ChangeSheetName(int k, string sheetname);
        void SetColor(int k, int mrow, int mcol, short R, short G, short B);
        void SetCellColumnWidth(int k, int mcol, float size);
        void SetCellRowHeight(int k, int mrow, short size);
        void SetCellStyle(int k, int mrow, int mcol, string CssStyle); //增加设置单元格样式
        void SetCellRangeAddress(int k, int rowstart, int rowend, int colstart, int colend); //合并单元格

        string GetPathByXlsToHTML(string strFile);
        string ExcelToHtml(int i);
        //增加excel to json
        string XlsToJson(string xls);    //excel 转 json
        
        void Insertbase64Picture(int k, int startrow, int startcol, int lastrow, int lastcol, int anchorType, string base64);
        string Getbase64PictureTest(int k);
    }

    [Guid("34F268AE-FDA9-4757-92ED-DF6AEB7D490E")]
    [ClassInterface(ClassInterfaceType.None)]

    public partial class ConvertXls : IConvertXls
    {
        private HSSFWorkbook wb = null;
        private string xlsfile = "";

        public void openxls(string xls)
        {
            FileStream file = new FileStream(xls, FileMode.Open, FileAccess.Read);
            wb = new HSSFWorkbook(file);
            HSSFPalette palette = wb.GetCustomPalette();
            //调色板实例

            palette.SetColorAtIndex((short)8, (byte)0, (byte)0, (byte)0);
            xlsfile = xls;
            file.Close();

        }

        public void CreateSheet(int nums)
        {
            for (int i = 1; i <= nums; i = i + 1)
            {
                wb.CreateSheet("sheet" + i.ToString());
            }

        }

        public string ExcelToHtml(int i)
        {
            if (i >= wb.NumberOfSheets) return "";
            ISheet sheet = wb.GetSheetAt(i);
            IWorkbook workbook = sheet.Workbook;
            ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter();

            // 设置输出参数
            excelToHtmlConverter.OutputColumnHeaders = false;
            excelToHtmlConverter.OutputHiddenColumns = false;
            excelToHtmlConverter.OutputHiddenRows = false;
            excelToHtmlConverter.OutputLeadingSpacesAsNonBreaking = true;
            excelToHtmlConverter.OutputRowNumbers = false;
            excelToHtmlConverter.UseDivsToSpan = true;

            // 处理的Excel文件
            excelToHtmlConverter.ProcessWorkbook(workbook);

            //添加表格样式
            /*excelToHtmlConverter.Document.InnerXml =
                excelToHtmlConverter.Document.InnerXml.Insert(
                    excelToHtmlConverter.Document.InnerXml.IndexOf("<head>", 0) + 6,
                    @"<style>table, td, th{border:1px solid green;}th{background-color:green;color:white;}</style>"
                );*/

            //方法一
            return excelToHtmlConverter.Document.InnerXml;
        }

        public string GetPathByXlsToHTML(string strFile)
        {
            if (string.IsNullOrEmpty(strFile))
            {
                return "0";//没有文件
            }

            //实例化Excel  
            Microsoft.Office.Interop.Excel.Application repExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

            //打开文件，n.FullPath是文件路径  
            workbook = repExcel.Application.Workbooks.Open(strFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];

            //ISheet sheet = wb.GetSheetAt(0);
            //IRow row = sheet.GetRow(0);
            //ICell cell = row.GetCell(0);
            //IWorkbook mywb = cell.Sheet.Workbook;

            //worksheet = (Microsoft.Office.Interop.Excel.Worksheet)mywb.GetSheetAt(0);
            //给文件重新起名
            string filename = System.DateTime.Now.Year.ToString() + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Day.ToString() +
            System.DateTime.Now.Hour.ToString() + System.DateTime.Now.Minute.ToString() + System.DateTime.Now.Second.ToString();

            string strFileFolder = "D:\\HGSoftware\\001_美蝶设计软件工厂版190807\\Python3\\Server\\Dll\\";
            DateTime dt = DateTime.Now;
            //以yyyymmdd形式生成子文件夹名
            string strFileSubFolder = dt.Year.ToString();
            strFileSubFolder += (dt.Month < 10) ? ("0" + dt.Month.ToString()) : dt.Month.ToString();
            strFileSubFolder += (dt.Day < 10) ? ("0" + dt.Day.ToString()) : dt.Day.ToString();
            string strFilePath = strFileFolder + strFileSubFolder + "\\";
            // 判断指定目录下是否存在文件夹，如果不存在，则创建 
            if (!Directory.Exists(strFilePath))
            {
                // 创建up文件夹 
                Directory.CreateDirectory(strFilePath);
            }
            string ConfigPath = (strFilePath + filename + ".html");    //输出完整路径
            //MessageBox.Show(ConfigPath);
            object savefilename = (object)ConfigPath;

            object ofmt = Microsoft.Office.Interop.Excel.XlFileFormat.xlHtml;
            //进行另存为操作    
            workbook.SaveAs(savefilename, ofmt, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            object osave = false;
            //逐步关闭所有使用的对象  
            workbook.Close(osave, Type.Missing, Type.Missing);
            repExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            worksheet = null;
            //垃圾回收  
            GC.Collect();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            workbook = null;
            GC.Collect();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(repExcel.Application.Workbooks);
            GC.Collect();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(repExcel);
            repExcel = null;
            GC.Collect();
            //依据时间杀灭进程  
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process p in process)
            {
                if (DateTime.Now.Second - p.StartTime.Second > 0 && DateTime.Now.Second - p.StartTime.Second < 5)
                {
                    p.Kill();
                }
            }

            return (strFilePath + filename + ".html");
        }

        public void closexls()
        {
            FileStream FileStreamfile = new FileStream(xlsfile, FileMode.Create);
            int sheetcount = wb.NumberOfSheets;

            for (int i = 0; i < sheetcount; i++)
            {
                ISheet sheet = wb.GetSheetAt(i);
                sheet.ForceFormulaRecalculation = true;
            }
            //for (int i = sheetcount-1; i > 0; i--)
            //{
            //    ISheet sheet = wb.GetSheetAt(i);
            //    sheet.ForceFormulaRecalculation = true;
            //}

            wb.Write(FileStreamfile);
            FileStreamfile.Close();
            wb = null;
        }
        //合并单元格
        public void SetCellRangeAddress(int k, int rowstart, int rowend, int colstart, int colend)
        {
            k = k - 1;
            rowstart = rowstart - 1;
            rowend = rowend - 1;
            colstart = colstart - 1;
            colend = colend - 1;
            string Result = "";
            ISheet sheet = wb.GetSheetAt(k);
            CellRangeAddress cellRangeAddress = new CellRangeAddress(rowstart, rowend, colstart, colend);
            sheet.AddMergedRegion(cellRangeAddress);
        }

        public void SetCellStyle(int k, int mrow, int mcol, string CssStyle)
        {
            k = k - 1;
            string Result = "";
            ISheet sheet = wb.GetSheetAt(k);
            mrow = mrow - 1;
            mcol = mcol - 1;

            IRow row = sheet.GetRow(mrow);
            if (row == null)
            {
                row = sheet.CreateRow(mrow);
            }
            ICell cell = row.GetCell(mcol); //|| (cell.CellType != CellType.String)
            if (cell == null)
            {
                cell = row.CreateCell(mcol);
            }

            cell.CSS(CssStyle);

        }

        public void SetColor(int k, int mrow, int mcol, short R, short G, short B)
        {
            k = k - 1;
            string Result = "";
            if (wb == null)
            {
                return;
            }
            ISheet sheet = wb.GetSheetAt(k);
            mrow = mrow - 1;
            mcol = mcol - 1;

            IRow row = sheet.GetRow(mrow);
            if (row == null)
            {
                row = sheet.CreateRow(mrow);
            }
            ICell cell = row.GetCell(mcol); //|| (cell.CellType != CellType.String)
            if ((cell == null))
            {
                cell = row.CreateCell(mcol);
            }
            ICellStyle s = wb.CreateCellStyle();
            //HSSFColor.GetIndexHash();

            HSSFPalette palette = wb.GetCustomPalette();
            HSSFColor hssFColor = palette.FindColor((Byte)R, (Byte)G, (Byte)B);
            s.FillForegroundColor = hssFColor.Indexed;
            s.FillPattern = FillPattern.SolidForeground;
            cell.CellStyle = s;

        }
        //设置列宽
        public void SetCellColumnWidth(int k, int mcol, float size)
        {
            k = k - 1;
            string Result = "";
            if (wb == null)
            {
                return;
            }
            ISheet sheet = wb.GetSheetAt(k);
            mcol = mcol - 1;
            sheet.SetColumnWidth(mcol, (int)((size + 0.63) * 256));
        }
        //设置行高
        public void SetCellRowHeight(int k, int mrow, short size)
        {
            k = k - 1;
            string Result = "";
            if (wb == null)
            {
                return;
            }
            ISheet sheet = wb.GetSheetAt(k);
            mrow = mrow - 1;

            IRow row = sheet.GetRow(mrow);
            if (row == null)
            {
                row = sheet.CreateRow(mrow);
            }

            row.Height = (short)(size * 20);

        }

        public void ChangeSheetName(int k, string sheetname)
        {
            k = k - 1;
            if (wb == null)
            {
                return;
            }
            wb.SetSheetName(k, sheetname);
        }
        //获取第k sheet 第mrow 行， 第mcol 列内容
        //C# 和delphi 行号差1
        //C# 和delphi 列号差1
        public string GetCell(int k, int mrow, int mcol)
        {
            k = k - 1;
            string Result = "";
            if (wb == null)
            {
                return Result;
            }
            ISheet sheet = wb.GetSheetAt(k);
            mrow = mrow - 1;
            mcol = mcol - 1;

            IRow row = sheet.GetRow(mrow);
            if (row == null)
            {
                return Result;
            }
            ICell cell = row.GetCell(mcol); //|| (cell.CellType != CellType.String)
            if ((cell == null))
            {
                return Result;
            }
            Result = cell.StringCellValue;
            return Result;
        }

        public void InsertCell(int k, int mrow, int mcol, string text)
        {
            k = k - 1;
            if (wb == null)
            {
                return;
            }
            ISheet sheet = wb.GetSheetAt(k);
            mrow = mrow - 1;
            mcol = mcol - 1;

            IRow row = sheet.GetRow(mrow);
            if (row == null)
            {
                row = sheet.CreateRow(mrow);
            }
            ICell cell = row.GetCell(mcol);
            if ((cell == null))
            {
                cell = row.CreateCell(mcol);
                cell.SetCellType(CellType.String);
            }
            switch (cell.CellType)
            {
                case CellType.Blank:
                    cell.SetCellValue(cell.StringCellValue);
                    break;
                case CellType.Boolean:
                    cell.SetCellValue(cell.BooleanCellValue);
                    break;
                case CellType.String:
                    cell.SetCellValue(text);
                    break;

                case CellType.Numeric:
                    cell.SetCellValue(Convert.ToInt32(text));
                    break;
            }
        }

        public void InsertNumCell(int k, int mrow, int mcol, string text)
        {
            k = k - 1;
            if (wb == null)
            {
                return;
            }
            ISheet sheet = wb.GetSheetAt(k);
            mrow = mrow - 1;
            mcol = mcol - 1;

            IRow row = sheet.GetRow(mrow);
            if (row == null)
            {

            }
            ICell cell = row.GetCell(mcol);
            if ((cell == null))
            {

            }
            cell.SetCellValue(Convert.ToDouble(text));

        }
        //获取sheet数目
        public string GetSheetNums()
        {
            if (wb == null)
            {
                return "";
            }
            string s;
            int sheetcount = wb.NumberOfSheets;
            s = sheetcount.ToString();
            return s;
        }
        //删除第k sheet 第mrow 行，操作
        public void RemoveOneRow(int k, int mrow)
        {
            if (wb == null)
            {
                return;
            }
            k = k - 1;
            mrow = mrow - 1;
            ISheet sheet = wb.GetSheetAt(k);
            IRow row = sheet.GetRow(mrow);
            RemoveRowMergedRegion(sheet, row.RowNum);
            sheet.RemoveRow(row);
            try
            {

                sheet.ShiftRows(mrow + 1, sheet.LastRowNum, -1, true, false);
            }
            catch
            {
            }

        }

        public void HideCol(int k, int mcol, bool ishide)
        {
            if (wb == null)
            {
                return;
            }
            k = k - 1;
            mcol = mcol - 1;
            ISheet sheet = wb.GetSheetAt(k);

            try
            {

                sheet.SetColumnHidden(mcol, ishide);
            }
            catch
            {
            }
        }

        public void RemoveOneCol(int k, int mrow, int mcol)
        {
            if (wb == null)
            {
                return;
            }
            k = k - 1;
            mrow = mrow - 1;
            mcol = mcol - 1;
            ISheet sheet = wb.GetSheetAt(k);

            for (int i = 0; i <= mrow; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                ICell cell = row.GetCell(mcol);//|| (cell.CellType != CellType.String)
                if ((cell == null)) continue;
                row.RemoveCell(cell);
            }
        }

        private void RemoveRowMergedRegion(ISheet sheet, int row)
        {
            if (wb == null)
            {
                return;
            }
            while (true)
            {
                bool finded = false;
                for (int i = sheet.NumMergedRegions - 1; i >= 0; i--)
                {
                    CellRangeAddress ca = sheet.GetMergedRegion(i);
                    if (ca.LastRow < row) return;
                    if ((ca.FirstRow <= row) && (ca.LastRow >= row))
                    {
                        finded = true;
                        sheet.RemoveMergedRegion(i);
                        break;
                    }
                }
                if (!finded) return;
            }
        }

        public void InsertRow(int k, int mrow)
        {
            if (wb == null)
            {
                return;
            }
            k = k - 1;
            ISheet sheet = wb.GetSheetAt(k);

            List<string> formulas = new List<string>();
            int startrow = 0;

            startrow = mrow - 1;
            int endrow = sheet.LastRowNum;

            //备份公式
            formulas.Clear();
            IRow row = sheet.GetRow(startrow);
            for (int j = 0; j <= row.LastCellNum; j++)
            {
                ICell cell = row.GetCell(j);
                if ((cell == null) || (cell.CellType != CellType.Formula))
                {
                    formulas.Add("");
                }
                else
                {
                    formulas.Add(cell.CellFormula);
                }
            }
            sheet.ShiftRows(startrow + 1, endrow, 1, true, false);
            CopyOneRow(sheet, startrow, startrow + 1, ref formulas);

        }

        private void CopyOneRow(ISheet sheet, int srcnum, int destnum, ref List<string> formulas)
        {
            if (wb == null)
            {
                return;
            }
            IRow srcrow = sheet.GetRow(srcnum);
            IRow destrow = sheet.GetRow(destnum);
            if (srcrow == null) return;
            if (destrow == null) destrow = sheet.CreateRow(destnum);

            destrow.Height = srcrow.Height;
            for (int i = 0; i < srcrow.LastCellNum; i++)
            {
                ICell oldCell = srcrow.GetCell(i);
                ICell newCell = destrow.CreateCell(i);
                if (oldCell == null)
                {
                    newCell = null;
                    continue;
                }
                newCell.CellStyle = oldCell.CellStyle;
                if (newCell.CellComment != null)
                {
                    newCell.CellComment = oldCell.CellComment;
                }
                if (oldCell.Hyperlink != null)
                {
                    newCell.Hyperlink = oldCell.Hyperlink;
                }
                newCell.SetCellType(oldCell.CellType);

                switch (oldCell.CellType)
                {
                    case CellType.Blank:
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                    case CellType.Boolean:
                        newCell.SetCellValue(oldCell.BooleanCellValue);
                        break;
                    case CellType.Error:
                        newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                        break;
                    case CellType.Formula:
                        newCell.SetCellFormula(formulas[i]);
                        break;
                    case CellType.Numeric:
                        newCell.SetCellValue(oldCell.NumericCellValue);
                        break;
                    case CellType.String:
                        newCell.SetCellValue(oldCell.RichStringCellValue);
                        break;
                    case CellType.Unknown:
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                }
            }
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress cellRangeAddress = sheet.GetMergedRegion(i);

                if (cellRangeAddress.FirstRow == srcrow.RowNum)
                {
                    int firstcol = cellRangeAddress.FirstColumn;
                    if (firstcol < 0) firstcol = 0;
                    int lastcol = cellRangeAddress.LastColumn;
                    if (lastcol < 0) lastcol = 0;
                    CellRangeAddress newCellRangeAddress = new CellRangeAddress(destrow.RowNum,
                                                                                (destrow.RowNum +
                                                                                 (cellRangeAddress.LastRow -
                                                                                  cellRangeAddress.FirstRow)),
                                                                                firstcol,
                                                                                lastcol);
                    sheet.AddMergedRegion(newCellRangeAddress);
                }
            }
        }

        //插入门图片
        public void InsertPicture(int k, int startrow, int startcol, int lastrow, int lastcol, string PicturePath)//, float PictuteWidth, float PictureHeight
        {
            if (wb == null)
            {
                return;
            }
            k = k - 1;
            ISheet sheet = wb.GetSheetAt(k);
            startrow = startrow - 1;
            startcol = startcol - 1;
            //IRow row = sheet.GetRow(startrow);
            //int rowline = 1;//从第二行开始(索引从0开始)
            //IRow row = sheet.CreateRow(startrow);
            //设置行高 ,excel行高度每个像素点是1/20
            //if (row == null) row = sheet.CreateRow(startrow);
            //row.Height = 80 * 20;
            //填入生产单号
            //row.CreateCell(0, CellType.String).SetCellValue("litao");
            //将图片文件读入一个字符串
            //byte[] bytes = System.IO.File.ReadAllBytes(PicturePath);
            //int pictureIdx = wb.AddPicture(bytes, PictureType.WMF);
            FileStream file = new FileStream(PicturePath, FileMode.Open, FileAccess.Read);
            byte[] buffer;
            buffer = new byte[file.Length];
            file.Read(buffer, 0, (int)file.Length);
            file.Close();


            string extension = Path.GetExtension(PicturePath);
            //MessageBox.Show(extension.ToLower());
            int pictureIdx = 0;
            if (extension.ToLower() == ".jpg")
            {
                pictureIdx = wb.AddPicture(buffer, PictureType.JPEG);
            }
            else if (extension.ToLower() == ".wmf")
            {
                pictureIdx = wb.AddPicture(buffer, PictureType.WMF);
            }
            else if (extension.ToLower() == ".wpg")
            {
                pictureIdx = wb.AddPicture(buffer, PictureType.WPG);
            }
            else if (extension.ToLower() == ".bmp")
            {
                pictureIdx = wb.AddPicture(buffer, PictureType.BMP);
            }
            else if (extension.ToLower() == ".png")
            {
                pictureIdx = wb.AddPicture(buffer, PictureType.PNG);
            }
            else if (extension.ToLower() == ".pict")
            {
                pictureIdx = wb.AddPicture(buffer, PictureType.PICT);
            }
            else if (extension.ToLower() == ".dib")
            {
                pictureIdx = wb.AddPicture(buffer, PictureType.DIB);
            }
            else if (extension.ToLower() == ".emf")
            {
                pictureIdx = wb.AddPicture(buffer, PictureType.EMF);
            }
            else if (extension.ToLower() == ".gif")
            {
                pictureIdx = wb.AddPicture(buffer, PictureType.GIF);
            }
            else if (extension.ToLower() == ".tiff")
            {
                pictureIdx = wb.AddPicture(buffer, PictureType.TIFF);
            }
            else if (extension.ToLower() == ".eps")
            {
                pictureIdx = wb.AddPicture(buffer, PictureType.EPS);
            }
            else
            {
                pictureIdx = wb.AddPicture(buffer, PictureType.Unknown);
            }
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
            // 插图片的位置  HSSFClientAnchor（dx1,dy1,dx2,dy2,col1,row1,col2,row2) 后面再作解释
            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, startcol, startrow, lastcol, lastrow);
            //把图片插到相应的位置
            HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
            //pict.Resize();
            /*int dy1 = 0;
            int dy2 = 0;
            int dx1 = 0;
            int dx2 = 0;
            HSSFClientAnchor anchor;
            anchor = new HSSFClientAnchor(dx1, dy1, dx2, dy2, startcol, startrow, startcol+1, startrow+1);
            anchor.AnchorType = (AnchorType)2;
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
            byte[] buffer;
            int pictureIndex;
         
            FileStream file = new FileStream(PicturePath, FileMode.Open, FileAccess.Read);
            buffer = new byte[file.Length];
            file.Read(buffer, 0, (int)file.Length);
            file.Close();
            pictureIndex = wb.AddPicture(buffer, PictureType.WMF);
            HSSFPicture picture = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIndex);
            picture.Resize();*/


        }

        public void deletePicture(int k, int i)
        {
            if (wb == null)
            {
                return;
            }
            i = i - 1;
            ISheet sheet = wb.GetSheetAt(k);
            //sheet.GetAllPictures();
            System.Collections.IList pictures = wb.GetAllPictures();
            int w = pictures.Count;
            Console.WriteLine("Count=" + w);
            pictures.Remove(i);

            //foreach (HSSFPictureData pic in pictures)
            //{
            //    pic.

            //}
        }


    
        


        public string XlsToJson(string xls)
        {
            string Result = "";
            //文件不存在就直接退出
            if ((string.IsNullOrEmpty(xls)) || (! File.Exists(xls)))
            {
                return "";//没有文件
            }
 
            JObject staff = new JObject();
            FileStream file = new FileStream(xls, FileMode.Open, FileAccess.Read);
            HSSFWorkbook mywk = new HSSFWorkbook(file);
            xlsfile = xls;
            string fileType = Path.GetExtension(xls).ToLower();
            string fileName = Path.GetFileName(xls).ToLower();
            staff.Add("type", fileType);
            staff.Add("fileName", fileName);
            file.Close();

            JArray sheets = new JArray();


            for (int k = 0; k < mywk.NumberOfSheets; k++)
            {
                JObject onesheet = new JObject();
                ISheet sheet = mywk.GetSheetAt(k);
               
                string sheetName = mywk.GetSheetName(k);    //读取当前表数据
                onesheet.Add("sheetName", sheetName);
                JObject data = new JObject();
                JObject rowheightobj = new JObject();
                JObject colwidthobj = new JObject();

                for (int i = 0; i <= sheet.LastRowNum; i++)
                {
                    JObject rowobj = new JObject();
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {

                        short rowheight = (short)(row.Height / 20);
                        rowheightobj.Add("L" + i.ToString(), rowheight);
                        for (int j = 0; j <= row.LastCellNum; j++)
                        {
                            

                            ICell cell = row.GetCell(j);
                            if (i == 0)
                            {
                                float ColumnWidth = (float)((float)(sheet.GetColumnWidth(j)) / 256 - 0.63);
                                colwidthobj.Add("C" + j.ToString(), ColumnWidth);
                            }

                            if (cell != null)
                            {
                                string style = GetCellStyle(cell, mywk);
                                                             
                                JObject cellobj = new JObject();
                                cellobj.Add("Text", cell.ToString());
                                cellobj.Add("style", style);
                                
                                rowobj.Add("C" + j.ToString(), cellobj);

                                Dimension dimension;
                                bool result = IsMergeCell(sheet, i, j, out dimension);
                                if (result)
                                {
                                    cellobj.Add("rowSpan", dimension.RowSpan.ToString());
                                    cellobj.Add("columnSpan", dimension.ColumnSpan.ToString());
                                    if ((i == dimension.FirstRowIndex) && (j == dimension.FirstColumnIndex))
                                    {
                                        cellobj.Add("_mergeCount", dimension.ColumnSpan-1);
                                    }
                                }
                            }
                        }
                    }
                    data.Add("L" + i.ToString(), rowobj);
                }
                onesheet.Add("data", data);  // 添加data
                onesheet.Add("RowHeight", rowheightobj);
                onesheet.Add("ColumnWidth", colwidthobj);
                //JArray pictures = new JArray();
                string pictures = Getbase64PictureTest1(sheet);
                onesheet.Add("pictures", pictures);
                sheets.Add(onesheet);
                

            }
            staff.Add("sheets", sheets);
            Result = staff.ToString();
            
            return Result;


        }

        


        /// <summary>
        /// 判断指定单元格是否为合并单元格，并且输出该单元格的维度
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="dimension">单元格维度</param>
        /// <returns>返回是否为合并单元格的布尔(Boolean)值</returns>





    }


    

    /*public interface IXlsToJson
    {
        [DispId(1)]
        void openxls(string xls);
        void closexls();
        
    }

    public class XlsToJson : IXlsToJson
    {
        private HSSFWorkbook mywk = null;
        private string xlsfile = "";

        public void XlsToJson(string xls)
        {
            FileStream file = new FileStream(xls, FileMode.Open, FileAccess.Read);
            mywk = new HSSFWorkbook(file);
            xlsfile = xls;
            file.Close();
            for (int i = 0; i < mywk.NumberOfSheets; i++)
            {
                ISheet sheet = mywk.GetSheetAt(i);
                string sheetName = mywk.GetSheetName(i);    //读取当前表数据
                for (int j = 0; j <= sheet.LastRowNum; j++)
                {
                    IRow row = sheet.GetRow(j);
                    if (row != null)
                    {
                        for (int k = 0; k <= row.LastCellNum; k++)
                        {
                            ICell cell = row.GetCell(k);
                            if (cell != null)
                            {
                                MessageBox.Show(cell.ToString());
                            }
                        }
                    }
                }

            }
            

        }
        public void closexls()
        {
            FileStream FileStreamfile = new FileStream(xlsfile, FileMode.Create);
            int sheetcount = wb.NumberOfSheets;

            for (int i = 1; i < sheetcount; i++)
            {
                ISheet sheet = wb.GetSheetAt(i);
                sheet.ForceFormulaRecalculation = true;
            }
            //for (int i = sheetcount-1; i > 0; i--)
            //{
            //    ISheet sheet = wb.GetSheetAt(i);
            //    sheet.ForceFormulaRecalculation = true;
            //}

            wb.Write(FileStreamfile);
            FileStreamfile.Close();
            wb = null;
        }
    }*/
}
