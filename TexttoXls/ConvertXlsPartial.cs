/**
 进行分文件
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;

//MessageBox.Show(extension.ToLower());
//20200227

using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Record;

namespace TexttoXls
{
    public partial class ConvertXls : IConvertXls
    {
        private string returnfontcolor(short color)
        {

            string Result = "font - color:";
            HSSFPalette mypalette = new HSSFPalette(new PaletteRecord());
            HSSFColor hssFColor = mypalette.GetColor(color);
            if (hssFColor == null) return "";
            byte[] rgb = hssFColor.RGB;
            int r = rgb[0];
            int g = rgb[1];
            int b = rgb[2];
            string R = Convert.ToString(r, 16);
            if (R == "0")
                R = "00";
            string G = Convert.ToString(g, 16);
            if (G == "0")
                G = "00";
            string B = Convert.ToString(b, 16);
            if (B == "0")
                B = "00";
            Result = Result + "#" + R + G + B + ";";
            return Result;
        }

        private string ConvertHorizontalAlignmentToString(NPOI.SS.UserModel.HorizontalAlignment alignment)
        {
            string Result = "text-align:";
            switch (alignment)
            {
                case NPOI.SS.UserModel.HorizontalAlignment.Left:
                    return Result + "LEFT;";
                case NPOI.SS.UserModel.HorizontalAlignment.Center:
                    return Result + "CENTER;";
                case NPOI.SS.UserModel.HorizontalAlignment.CenterSelection:
                    return Result + "CENTER_SELECTION;";
                case NPOI.SS.UserModel.HorizontalAlignment.Right:
                    return Result + "RIGHT;";
                case NPOI.SS.UserModel.HorizontalAlignment.Distributed:
                    return Result + "DISTRIBUTED;";
                case NPOI.SS.UserModel.HorizontalAlignment.Fill:
                    return Result + "FILL;";
                case NPOI.SS.UserModel.HorizontalAlignment.Justify:
                    return Result + "JUSTIFY;";

                default:
                    return "";
            }
        }

        private string ConvertBorderStyleToString(NPOI.SS.UserModel.BorderStyle boderstyle)
        {
            switch (boderstyle)
            {
                case NPOI.SS.UserModel.BorderStyle.Thin:
                    return "THIN";
                case NPOI.SS.UserModel.BorderStyle.Medium:
                    return "MEDIUM";
                case NPOI.SS.UserModel.BorderStyle.Dashed:
                    return "DASHED";
                case NPOI.SS.UserModel.BorderStyle.Hair:
                    return "HAIR";
                case NPOI.SS.UserModel.BorderStyle.Thick:
                    return "THICK";
                case NPOI.SS.UserModel.BorderStyle.Double:
                    return "DOUBLE";
                case NPOI.SS.UserModel.BorderStyle.Dotted:
                    return "DOTTED";
                case NPOI.SS.UserModel.BorderStyle.MediumDashed:
                    return "MEDIUMDASHED";
                case NPOI.SS.UserModel.BorderStyle.DashDot:
                    return "DASHDOT";
                case NPOI.SS.UserModel.BorderStyle.MediumDashDot:
                    return "MEDIUMDASHDOT";
                case NPOI.SS.UserModel.BorderStyle.DashDotDot:
                    return "DASHDOTDOT";
                case NPOI.SS.UserModel.BorderStyle.MediumDashDotDot:
                    return "MEDIUMDASHDOTDOT";
                case NPOI.SS.UserModel.BorderStyle.SlantedDashDot:
                    return "SLANTEDDASHDOT";
                default:
                    return "None";
            }
        }
        //边框分解 上 右 下 左
        private string GetBoderStyle(ICellStyle cellstyle)
        {
            NPOI.SS.UserModel.BorderStyle boderstyle = cellstyle.BorderTop;
            string Result = "border-type:";
            string topboderstyle = ConvertBorderStyleToString(cellstyle.BorderTop);
            string rightboderstyle = ConvertBorderStyleToString(cellstyle.BorderRight);
            string bomboderstyle = ConvertBorderStyleToString(cellstyle.BorderBottom);
            string leftboderstyle = ConvertBorderStyleToString(cellstyle.BorderLeft);
            if ((topboderstyle == "") && (rightboderstyle == "") && (bomboderstyle == "") && (leftboderstyle == ""))
            {
                return "";
            }
            Result = Result + topboderstyle + " " + rightboderstyle + " " + bomboderstyle + " " + leftboderstyle + ";";
            return Result;

        }

        private string GetUnderline(FontUnderlineType fontunderlinetype)
        {
            switch (fontunderlinetype)
            {
                case FontUnderlineType.Single:
                    return "font-underline:SINGLE;";
                case FontUnderlineType.Double:
                    return "font-underline:DOUBLE;";
                case FontUnderlineType.SingleAccounting:
                    return "font-underline:SINGLEACCOUNTING;";
                case FontUnderlineType.DoubleAccounting:
                    return "font-underline:DOUBLEACCOUNTING;";
                case FontUnderlineType.None:
                    return "";
                default:
                    return "";
            }
        }

        private string GetCellStyle(ICell cell, IWorkbook mywk)
        {
            
            ICellStyle cellStyle = cell.CellStyle;
            
            IFont font = cellStyle.GetFont(mywk);
            //Console.WriteLine(cellStyle.FontIndex.ToString());
            //Console.WriteLine(font.Index.ToString()+cell.ToString());
            //Console.WriteLine(font.ToString());
            string Result = "";
            short weight = font.Boldweight;   //字体加粗
            if (weight == 700)
            {
                Result = Result + "font-weight:bold;";
            }
            else if (weight == 400)
            {
                Result = Result + "font-weight:normal;";
            }
            short color = font.Color;    //字体颜色
            Result = Result + returnfontcolor(color);

            string fontname = font.FontName;    //字体类型
            Result = Result + "font-name:" + fontname + ";";

            double fontsize = font.FontHeightInPoints;    //字体大小
            Result = Result + "font-size:" + fontsize.ToString() + ";";

            string textalign = ConvertHorizontalAlignmentToString(cellStyle.Alignment); //居中对齐
            Result = Result + textalign;

            string bordertype = GetBoderStyle(cellStyle); //边框
            Result = Result + bordertype;
            
            string fontunderline = GetUnderline(font.Underline); //下划线
            Result = Result + fontunderline;

            if (cellStyle.WrapText)
            {
                Result = Result + "WrapText:True";
            }

            return Result;
        }

        private struct Dimension
        {
            /// <summary>
            /// 含有数据的单元格(通常表示合并单元格的第一个跨度行第一个跨度列)，该字段可能为null
            /// </summary>
            public ICell DataCell;

            /// <summary>
            /// 行跨度(跨越了多少行)
            /// </summary>
            public int RowSpan;

            /// <summary>
            /// 列跨度(跨越了多少列)
            /// </summary>
            public int ColumnSpan;

            /// <summary>
            /// 合并单元格的起始行索引
            /// </summary>
            public int FirstRowIndex;

            /// <summary>
            /// 合并单元格的结束行索引
            /// </summary>
            public int LastRowIndex;

            /// <summary>
            /// 合并单元格的起始列索引
            /// </summary>
            public int FirstColumnIndex;

            /// <summary>
            /// 合并单元格的结束列索引
            /// </summary>
            public int LastColumnIndex;
        }

        private bool IsMergeCell(ISheet sheet, int rowIndex, int columnIndex, out Dimension dimension)
        {
            dimension = new Dimension
            {
                DataCell = null,
                RowSpan = 1,
                ColumnSpan = 1,
                FirstRowIndex = rowIndex,
                LastRowIndex = rowIndex,
                FirstColumnIndex = columnIndex,
                LastColumnIndex = columnIndex
            };

            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                sheet.IsMergedRegion(range);

                //这种算法只有当指定行列索引刚好是合并单元格的第一个跨度行第一个跨度列时才能取得合并单元格的跨度
                //if (range.FirstRow == rowIndex && range.FirstColumn == columnIndex)
                //{
                //    dimension.DataCell = sheet.GetRow(range.FirstRow).GetCell(range.FirstColumn);
                //    dimension.RowSpan = range.LastRow - range.FirstRow + 1;
                //    dimension.ColumnSpan = range.LastColumn - range.FirstColumn + 1;
                //    dimension.FirstRowIndex = range.FirstRow;
                //    dimension.LastRowIndex = range.LastRow;
                //    dimension.FirstColumnIndex = range.FirstColumn;
                //    dimension.LastColumnIndex = range.LastColumn;
                //    break;
                //}

                if ((rowIndex >= range.FirstRow && range.LastRow >= rowIndex) && (columnIndex >= range.FirstColumn && range.LastColumn >= columnIndex))
                {
                    dimension.DataCell = sheet.GetRow(range.FirstRow).GetCell(range.FirstColumn);
                    dimension.RowSpan = range.LastRow - range.FirstRow + 1;
                    dimension.ColumnSpan = range.LastColumn - range.FirstColumn + 1;
                    dimension.FirstRowIndex = range.FirstRow;
                    dimension.LastRowIndex = range.LastRow;
                    dimension.FirstColumnIndex = range.FirstColumn;
                    dimension.LastColumnIndex = range.LastColumn;
                    break;
                }
            }

            bool result;
            if (rowIndex >= 0 && sheet.LastRowNum > rowIndex)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (columnIndex >= 0 && row.LastCellNum > columnIndex)
                {
                    ICell cell = row.GetCell(columnIndex);
                    result = cell.IsMergedCell;

                    if (dimension.DataCell == null)
                    {
                        dimension.DataCell = cell;
                    }
                }
                else
                {
                    result = false;
                }
            }
            else
            {
                result = false;
            }

            return result;
        }
    }
}
