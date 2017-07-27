using EPPlus.Html.Html;
using EPPlus.Html.Model;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Converters
{
    internal static class ExcelToStyle
    {
        internal static IStyling ToStyle(this ExcelWorksheet excelSheet)
        {
            IStyling styling = new Styling();
            return styling;
        }

        internal static IStyling ToStyle(this ExcelRow excelRow)
        {
            IStyling styling = excelRow.Style.ToStyle();
            styling.Height = excelRow.Height;
            return styling;
        }

        internal static IStyling ToStyle(this ExcelRangeBase excelRange)
        {
            IStyling styling = excelRange.Style.ToStyle();
            if (excelRange.Columns == 1 && excelRange.Rows == 1)
            {
                var excelColumn = excelRange.Worksheet.Column(excelRange.Start.Column);
                styling.Width = excelColumn.Width;
            }
            return styling;
        }

        internal static HAlign GetHAlign(this ExcelStyle excelStyle)
        {
            switch (excelStyle.HorizontalAlignment)
            {
                case ExcelHorizontalAlignment.Right:
                    return HAlign.Right;
                case ExcelHorizontalAlignment.Left:
                    return HAlign.Left;
                case ExcelHorizontalAlignment.Center:
                    return HAlign.Center;
                default:
                    return HAlign.Left;
            }
        }

        internal static IStyling ToStyle(this ExcelStyle excelStyle)
        {
            IStyling styling = new Styling()
            {
                Background = excelStyle.Fill.BackgroundColor.ToHexCode(),
                HAlign = excelStyle.GetHAlign()
            };
            excelStyle.Border.AddToStyling(styling);
            excelStyle.Font.AddToStyling(styling);
            return styling;
        }

        internal static void AddToStyling(this ExcelFont excelFont, IStyling styling)
        {
            if (excelFont.Bold)
            {
                styling.IsBold = true;
            }
            styling.FontFamily = excelFont.Name;
            styling.FontSize = excelFont.Size;
            styling.Color = excelFont.Color.ToHexCode();
        }

        internal static void AddToStyling(this Border border, IStyling styling)
        {
            if (border.Top.Style != ExcelBorderStyle.None)
            {
                styling.BorderTop = border.Top.ToBorderStyle();
            }
            if (border.Bottom.Style != ExcelBorderStyle.None)
            {
                styling.BorderBottom = border.Bottom.ToBorderStyle();
            }
            if (border.Right.Style != ExcelBorderStyle.None)
            {
                styling.BorderRight = border.Right.ToBorderStyle();
            }
            if (border.Left.Style != ExcelBorderStyle.None)
            {
                styling.BorderLeft = border.Left.ToBorderStyle();
            };
        }

        internal static IBorderStyle ToBorderStyle(this ExcelBorderItem borderItem)
        {
            string color = (borderItem.Color.Rgb != null)
                   ? borderItem.Color.ToHexCode()
                   : "black";
            IBorderStyle borderStyle = new BorderStyle() { Color = color };
            switch (borderItem.Style)
            {
                case ExcelBorderStyle.Thin:
                    borderStyle.LineStyle = "solid";
                    borderStyle.Width = 1;
                    break;
                case ExcelBorderStyle.Thick:
                    borderStyle.LineStyle = "solid";
                    borderStyle.Width = 2;
                    break;
                case ExcelBorderStyle.None:
                    borderStyle.LineStyle = "none";
                    break;
                default:
                    borderStyle.LineStyle = "solid";
                    borderStyle.Width = 0 + 1;
                    break;
            }
            return borderStyle;
        }
    }
}
