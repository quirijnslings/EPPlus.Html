
using EPPlus.Html.Converters;
using EPPlus.Html.Html;
using EPPlus.Html.Json;
using EPPlus.Html.Model;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using OfficeOpenXml;
using System;
using System.Linq;
using System.Text;

namespace EPPlus.Html
{
    public static class EPPlusExtensions
    {
        public static string ToHtml(this ExcelWorksheet sheet)
        {
            return sheet.ToHtml(HtmlExportConfiguration.Default());
        }

        public static string ToHtml(this ExcelWorksheet sheet, bool consolidateStyles)
        {
            return sheet.ToHtml(HtmlExportConfiguration.Default(), consolidateStyles);
        }

        public static string ToHtml(this ExcelWorksheet sheet, HtmlExportConfiguration configuration)
        {
            return sheet.ToHtml(configuration, false);
        }

        public static string ToHtml(this ExcelWorksheet sheet, HtmlExportConfiguration configuration, bool consolidateStyles)
        {
            int lastRow = sheet.Dimension.Rows;
            int lastCol = sheet.Dimension.Columns;

            HtmlElement htmlTable = new HtmlElement("table");
            htmlTable.Attributes["cellspacing"] = 0;
            htmlTable.Styles["white-space"] = "nowrap";

            //render rows
            for (int row = 1; row <= lastRow; row++)
            {
                ExcelRow excelRow = sheet.Row(row);

                HtmlElement htmlRow = htmlTable.AddChild("tr");
                if (configuration.Height)
                {
                    htmlRow.Styles.Update(excelRow.ToCss(configuration));
                }

                for (int col = 1; col <= lastCol; col++)
                {
                    ExcelRange excelCell = sheet.Cells[row, col];
                    HtmlElement htmlCell = htmlRow.AddChild("td");
                    htmlCell.Content = excelCell.Text;
                    htmlCell.Styles.Update(excelCell.ToCss(configuration));
                }
            }

            // consolidate styles into classes
            if (consolidateStyles)
            {
                htmlTable.ClassTable = new CssClassTable();
                htmlTable.ConsolidateStyles(htmlTable.ClassTable);
            }

            return htmlTable.ToString();
        }

        public static string ToJson(this ExcelWorksheet sheet)
        {
            return ToJson(sheet, JsonStyle.AsIs);
        }
        public static string ToJson(this ExcelWorksheet sheet, JsonStyle jsonStyle)
        {
            ITable table = ConvertToTable(sheet);
            IContractResolver contractResolver = null;
            Formatting formatting = Formatting.None;

            switch (jsonStyle)
            {
                case JsonStyle.Indented:
                    formatting = Formatting.Indented;
                    break;
                case JsonStyle.Lc:
                    contractResolver = new LowercaseContractResolver();
                    break;
                case JsonStyle.LcIndented:
                    contractResolver = new LowercaseContractResolver();
                    formatting = Formatting.Indented;
                    break;
                case JsonStyle.LcFirst:
                    contractResolver = new LowercaseFirstContractResolver();
                    break;
                case JsonStyle.LcFirstIndented:
                    contractResolver = new LowercaseFirstContractResolver();
                    formatting = Formatting.Indented;
                    break;
            }

            if (contractResolver == null)
            {
                return JsonConvert.SerializeObject(table, formatting);
            }
            var settings = new JsonSerializerSettings()
            {
                ContractResolver = contractResolver
            };
            return JsonConvert.SerializeObject(table, formatting, settings);
        }

        private static ITable ConvertToTable(this ExcelWorksheet sheet)
        {
            int lastRow = sheet.Dimension.Rows;
            int lastCol = sheet.Dimension.Columns;

            ITable table = new Table();
            table.Styling = sheet.ToStyle();

            //render rows
            for (int rowNr = 1; rowNr <= lastRow; rowNr++)
            {
                ExcelRow excelRow = sheet.Row(rowNr);
                IRow row = new Row();
                table.Rows.Add(row);
                row.Styling = excelRow.ToStyle();

                for (int col = 1; col <= lastCol; col++)
                {
                    ExcelRange excelCell = sheet.Cells[rowNr, col];
                    ICell cell = new Cell();
                    row.Cells.Add(cell);
                    cell.Text = excelCell.Text;
                    cell.Styling = excelCell.ToStyle();
                }
            }
            return table;
        }

        public static void ConsolidateStyles(this HtmlElement elmt, CssClassTable classTable)
        {
            if (elmt.Styles.Any())
            {
                StringBuilder sb = new StringBuilder();
                foreach (var style in elmt.Styles)
                {
                    sb.Append($"{style.Key}:{style.Value};");
                }

                var styleDef = sb.ToString();
                var className = classTable.GetOrAddClassForStyle(styleDef);
                elmt.InlineClasses.Add(className);
            }
            foreach (var child in elmt.Children)
            {
                ConsolidateStyles(child, classTable);
            }
        }

        public static CssInlineStyles ToCSS(this ExcelStyles styles)
        {
            throw new NotImplementedException();
        }
    }
}
