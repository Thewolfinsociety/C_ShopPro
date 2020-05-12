
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using NPOI.XSSF.UserModel;
using Newtonsoft.Json.Linq;
using NPOI.SS.Util;
using ZXing;

namespace TexttoXls
{
    using System;
    using System.Collections;
    using System.Text.RegularExpressions;
    using System.Text;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;

    public class MyQRCode
    {
        private byte[] BitmapToBytes(Bitmap bitmap)
        {
            byte[] byteImage = { };
            try
            {
                MemoryStream stream = new MemoryStream();
                bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Bmp);
                byteImage = new Byte[stream.Length];
                byteImage = stream.ToArray();
                stream.Close();
            }
            catch
            {
            }
            return byteImage;
        }
        public void CreateQRCode(HSSFWorkbook wb, ISheet sheet, string value, int rownum, int colnum)
        {
            //todo 计算二维码的大小和位置
            int sheetmergecount = sheet.NumMergedRegions;
            rownum = rownum - 1;
            colnum = colnum - 1;
            for (int k = sheetmergecount - 1; k >= 0; k--)
            {
                CellRangeAddress ca = sheet.GetMergedRegion(k);
                if ((rownum >= ca.FirstRow) && (rownum <= ca.LastRow) && (colnum >= ca.FirstColumn) && (colnum <= ca.LastColumn))
                {
                    BarcodeWriter writer = new BarcodeWriter();
                    writer.Format = BarcodeFormat.QR_CODE;
                    writer.Options.Hints.Add(EncodeHintType.CHARACTER_SET, "UTF-8");//编码问题
                    writer.Options.Hints.Add(EncodeHintType.ERROR_CORRECTION, ZXing.QrCode.Internal.ErrorCorrectionLevel.H);
                    writer.Options.Height = writer.Options.Width = 256;
                    writer.Options.Margin = 1;//设置边框
                    ZXing.Common.BitMatrix bm = writer.Encode(value);
                    Bitmap img = writer.Write(bm);
                    byte[] buffer = BitmapToBytes(img);

                    if (buffer == null) break;
                    HSSFClientAnchor anchor;
                    anchor = new HSSFClientAnchor(0, 0, 0, 0, ca.FirstColumn, ca.FirstRow, ca.LastColumn + 1, ca.LastRow + 1);

                    anchor.AnchorType = (AnchorType)2;
                    HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
                    int pictureIndex = wb.AddPicture(buffer, PictureType.JPEG);
                    HSSFPicture picture = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIndex);
                    break;
                }
                else
                {
                    BarcodeWriter writer = new BarcodeWriter();
                    writer.Format = BarcodeFormat.QR_CODE;
                    writer.Options.Hints.Add(EncodeHintType.CHARACTER_SET, "UTF-8");//编码问题
                    writer.Options.Hints.Add(EncodeHintType.ERROR_CORRECTION, ZXing.QrCode.Internal.ErrorCorrectionLevel.H);
                    writer.Options.Height = writer.Options.Width = 256;
                    writer.Options.Margin = 1;//设置边框
                    ZXing.Common.BitMatrix bm = writer.Encode(value);
                    Bitmap img = writer.Write(bm);
                    byte[] buffer = BitmapToBytes(img);

                    if (buffer == null) break;
                    HSSFClientAnchor anchor;
                    anchor = new HSSFClientAnchor(0, 0, 0, 0, colnum, rownum, colnum + 1, rownum + 1);

                    anchor.AnchorType = (AnchorType)2;
                    HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
                    int pictureIndex = wb.AddPicture(buffer, PictureType.JPEG);
                    HSSFPicture picture = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIndex);
                    break;

                }
            }

            if (sheetmergecount == 0)
            {
                BarcodeWriter writer = new BarcodeWriter();
                writer.Format = BarcodeFormat.QR_CODE;
                writer.Options.Hints.Add(EncodeHintType.CHARACTER_SET, "UTF-8");//编码问题
                writer.Options.Hints.Add(EncodeHintType.ERROR_CORRECTION, ZXing.QrCode.Internal.ErrorCorrectionLevel.H);
                writer.Options.Height = writer.Options.Width = 256;
                writer.Options.Margin = 1;//设置边框
                ZXing.Common.BitMatrix bm = writer.Encode(value);
                Bitmap img = writer.Write(bm);
                byte[] buffer = BitmapToBytes(img);

                if (buffer == null) return;
                HSSFClientAnchor anchor;
                anchor = new HSSFClientAnchor(0, 0, 0, 0, colnum, rownum, colnum + 1, rownum + 1);

                anchor.AnchorType = (AnchorType)2;
                HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
                int pictureIndex = wb.AddPicture(buffer, PictureType.JPEG);
                HSSFPicture picture = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIndex);
            }
        }
    }

    public class PicturesInfo
    {
        public int MinRow { get; set; }
        public int MaxRow { get; set; }
        public int MinCol { get; set; }
        public int MaxCol { get; set; }
        public AnchorType AnchorType { get; set; }
        public Byte[] PictureData { get; private set; }

        public PicturesInfo(int minRow, int maxRow, int minCol, int maxCol, Byte[] pictureData, AnchorType AnchorType)
        {
            this.MinRow = minRow;
            this.MaxRow = maxRow;
            this.MinCol = minCol;
            this.MaxCol = maxCol;
            this.PictureData = pictureData;
            this.AnchorType = AnchorType;
        }
    }

    public static class NpoiExtend
    {
        public static List<PicturesInfo> GetAllPictureInfos(this ISheet sheet)
        {
            return sheet.GetAllPictureInfos(null, null, null, null);
        }

        public static List<PicturesInfo> GetAllPictureInfos(this ISheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal = true)
        {
            if (sheet is HSSFSheet)
            {
                return GetAllPictureInfos((HSSFSheet)sheet, minRow, maxRow, minCol, maxCol, onlyInternal);
            }
            else if (sheet is XSSFSheet)
            {
                return GetAllPictureInfos((XSSFSheet)sheet, minRow, maxRow, minCol, maxCol, onlyInternal);
            }
            else
            {
                throw new Exception("未处理类型，没有为该类型添加：GetAllPicturesInfos()扩展方法！");
            }
        }

        private static List<PicturesInfo> GetAllPictureInfos(HSSFSheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal)
        {
            List<PicturesInfo> picturesInfoList = new List<PicturesInfo>();

            var shapeContainer = sheet.DrawingPatriarch as HSSFShapeContainer;
            if (null != shapeContainer)
            {
                var shapeList = shapeContainer.Children;
                foreach (var shape in shapeList)
                {
                    if (shape is HSSFPicture && shape.Anchor is HSSFClientAnchor)
                    {
                        var picture = (HSSFPicture)shape;
                        //Boolean isnofill = picture.IsNoFill;
                        Console.WriteLine("ShapeType=" + picture.ShapeType);
                        var anchor = (HSSFClientAnchor)shape.Anchor;

                        if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, onlyInternal))
                        {

                            Console.WriteLine("AnchorType=" + anchor.AnchorType);
                            picturesInfoList.Add(new PicturesInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, picture.PictureData.Data, anchor.AnchorType));
                        }
                    }
                }
            }

            return picturesInfoList;
        }

        private static List<PicturesInfo> GetAllPictureInfos(XSSFSheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal)
        {
            List<PicturesInfo> picturesInfoList = new List<PicturesInfo>();

            var documentPartList = sheet.GetRelations();
            foreach (var documentPart in documentPartList)
            {
                if (documentPart is XSSFDrawing)
                {
                    var drawing = (XSSFDrawing)documentPart;
                    var shapeList = drawing.GetShapes();
                    foreach (var shape in shapeList)
                    {
                        if (shape is XSSFPicture)
                        {
                            var picture = (XSSFPicture)shape;
                         
                            var anchor = picture.GetPreferredSize();

                            if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, onlyInternal))
                            {
                                picturesInfoList.Add(new PicturesInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, picture.PictureData.Data, anchor.AnchorType));
                            }
                        }
                    }
                }
            }

            return picturesInfoList;
        }

        private static bool IsInternalOrIntersect(int? rangeMinRow, int? rangeMaxRow, int? rangeMinCol, int? rangeMaxCol,
            int pictureMinRow, int pictureMaxRow, int pictureMinCol, int pictureMaxCol, bool onlyInternal)
        {
            int _rangeMinRow = rangeMinRow ?? pictureMinRow;
            int _rangeMaxRow = rangeMaxRow ?? pictureMaxRow;
            int _rangeMinCol = rangeMinCol ?? pictureMinCol;
            int _rangeMaxCol = rangeMaxCol ?? pictureMaxCol;

            if (onlyInternal)
            {
                return (_rangeMinRow <= pictureMinRow && _rangeMaxRow >= pictureMaxRow &&
                        _rangeMinCol <= pictureMinCol && _rangeMaxCol >= pictureMaxCol);
            }
            else
            {
                return ((Math.Abs(_rangeMaxRow - _rangeMinRow) + Math.Abs(pictureMaxRow - pictureMinRow) >= Math.Abs(_rangeMaxRow + _rangeMinRow - pictureMaxRow - pictureMinRow)) &&
                (Math.Abs(_rangeMaxCol - _rangeMinCol) + Math.Abs(pictureMaxCol - pictureMinCol) >= Math.Abs(_rangeMaxCol + _rangeMinCol - pictureMaxCol - pictureMinCol)));
            }
        }
    }

    public partial class ConvertXls : IConvertXls
    {
        public string Getbase64PictureTest1(ISheet sheet)
        {

            List<PicturesInfo> picturesInfoList = sheet.GetAllPictureInfos();
            JArray picturesInfoListObj = new JArray();
            foreach (var picturesInfo in picturesInfoList)
            {
                JObject picturesInfoObj = new JObject();
                picturesInfoObj.Add("startrow", picturesInfo.MinRow);
                picturesInfoObj.Add("startcol", picturesInfo.MinCol);
                picturesInfoObj.Add("endrow", picturesInfo.MaxRow);
                picturesInfoObj.Add("endcol", picturesInfo.MaxCol);
                picturesInfoObj.Add("picturedata", Convert.ToBase64String(picturesInfo.PictureData));
                switch (picturesInfo.AnchorType)
                {
                    case AnchorType.MoveAndResize:
                        picturesInfoObj.Add("AnchorType", 0);
                        break;
                    case AnchorType.MoveDontResize:
                        picturesInfoObj.Add("AnchorType", 2);
                        break;
                    case AnchorType.DontMoveAndResize:
                        picturesInfoObj.Add("AnchorType", 3);
                        break;
                    default:
                        break;
                }
                picturesInfoListObj.Add(picturesInfoObj);
            }

            return picturesInfoListObj.ToString();
        }
        // 获取图片信息
        public string Getbase64PictureTest(int k)
        {
            k = k - 1;
            ISheet sheet = wb.GetSheetAt(k);
            List<PicturesInfo> picturesInfoList = sheet.GetAllPictureInfos();
            JArray picturesInfoListObj = new JArray();
            foreach (var picturesInfo in picturesInfoList)
            {
                JObject picturesInfoObj = new JObject();
                picturesInfoObj.Add("startrow", picturesInfo.MinRow);
                picturesInfoObj.Add("startcol", picturesInfo.MinCol);
                picturesInfoObj.Add("endrow", picturesInfo.MaxRow);
                picturesInfoObj.Add("endcol", picturesInfo.MaxCol);
                picturesInfoObj.Add("picturedata", Convert.ToBase64String(picturesInfo.PictureData));
                picturesInfoListObj.Add(picturesInfoObj);
            }

            return picturesInfoListObj.ToString();
        }
        // 获取图片信息
        public JArray Getbase64Picture(ISheet sheet)
        {
            List<PicturesInfo> picturesInfoList = sheet.GetAllPictureInfos();
            JArray picturesInfoListObj = new JArray();
            foreach (var picturesInfo in picturesInfoList)
            {
                JObject picturesInfoObj = new JObject();
                picturesInfoObj.Add("startrow", picturesInfo.MinRow);
                picturesInfoObj.Add("startcol", picturesInfo.MinCol);
                picturesInfoObj.Add("endrow", picturesInfo.MaxRow);
                picturesInfoObj.Add("endcol", picturesInfo.MaxCol);
                picturesInfoObj.Add("picturedata", Convert.ToBase64String(picturesInfo.PictureData));
                picturesInfoListObj.Add(picturesInfoObj);
            }

            return picturesInfoListObj;
        }

        public void InsertMyQRCode(int k, int startrow, int startcol, string value)
        {
            if (wb == null)
            {
                return;
            }
            k = k - 1;
            ISheet sheet = wb.GetSheetAt(k);
            MyQRCode mrcode = new MyQRCode();
            mrcode.CreateQRCode(wb, sheet, value, startrow, startcol);
        }
        //插入base64图片数据
        public void Insertbase64Picture(int k, int startrow, int startcol, int lastrow, int lastcol, int anchorType, string base64)
        {
            if (wb == null)
            {
                return;
            }
            k = k - 1;
            ISheet sheet = wb.GetSheetAt(k);
            HSSFClientAnchor anchor;
            anchor = new HSSFClientAnchor(24, 24, 0, 0, startcol, startrow, lastcol, lastrow);
            anchor.AnchorType = (AnchorType)2;
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();

            byte[] arr = Convert.FromBase64String(base64);
            MemoryStream ms = new MemoryStream(arr);

            System.Drawing.Bitmap bmp = new Bitmap(ms);

            MemoryStream ms2 = new MemoryStream();
            bmp.Save(ms2, System.Drawing.Imaging.ImageFormat.Png);
            byte[] buffer = ms2.GetBuffer();
            ms2.Close();
            int pictureIndex = wb.AddPicture(buffer, PictureType.PNG);
            switch (anchorType)
            {
                case 0:
                    anchor.AnchorType = AnchorType.MoveAndResize;
                    break;
                case 2:
                    anchor.AnchorType = AnchorType.MoveDontResize;
                    break;
                case 3:
                    anchor.AnchorType = AnchorType.DontMoveAndResize;
                    break;
                default:
                    break;
            }
            

            HSSFPicture picture = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIndex);
            picture.IsNoFill = false;
          
            //picture.
        }

        //-------------------------插入图片--------------------------------
        public void InsertSlidingPicture2Sheet(ISheet sheet, string base64)
        {
            int lr = sheet.LastRowNum;  //行数
            for (int i = 0; i <= lr; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                {
                    continue;
                }
                int lc = row.LastCellNum;//列数
                for (int j = 0; j <= lc; j++)
                {
                    ICell cell = row.GetCell(j);
                    if ((cell != null) && (cell.CellType == CellType.String))
                    {
                        string cellstr = cell.ToString();
                        if (cellstr.Contains("SlidingPicture"))
                        {
                            string value = cellstr.Replace("SlidingPicture", "");
                            if (value == "")
                            {
                                continue;
                            }
                            //解析图片的位置
                            string[] strs = value.Split(',');
                            if (strs.Length < 3)
                            {
                                continue;
                            }
                            int col1 = j;
                            int row1 = i;
                            int col2 = 0;
                            int.TryParse(strs[1], out col2);
                            col2 += col1;
                            int row2 = 0;
                            int.TryParse(strs[2], out row2);
                            row2 += row1;

                            HSSFClientAnchor anchor;
                            anchor = new HSSFClientAnchor(24, 24, 0, 0, col1, row1, col2, row2);
                            anchor.AnchorType = (AnchorType)2;
                            HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();

                            byte[] arr = Convert.FromBase64String(base64);
                            MemoryStream ms = new MemoryStream(arr);
                           
                            System.Drawing.Bitmap bmp = new Bitmap(ms);

                            MemoryStream ms2 = new MemoryStream();
                            bmp.Save(ms2, System.Drawing.Imaging.ImageFormat.Png);
                            byte[] buffer = ms2.GetBuffer();
                            ms2.Close();
                            int pictureIndex = wb.AddPicture(buffer, PictureType.PNG);
                            HSSFPicture picture = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIndex);
                        }
                    }
                }
            }
        }
    }
}