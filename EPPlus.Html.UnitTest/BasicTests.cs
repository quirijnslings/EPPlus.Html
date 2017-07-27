using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml;
using System.Text;
using EPPlus.Html;


namespace Test
{
    [TestClass]
    public class BasicTests
    {

        [TestMethod]
        public void OptionsNone()
        {
            ExcelWorksheet ws = TestHelper.GetWorkSheet();
            ExportExcel(ws, HtmlExportConfiguration.Minimal());
        }

        [TestMethod]
        public void OptionsAll()
        {
            ExcelWorksheet ws = TestHelper.GetWorkSheet();
            ExportExcel(ws, HtmlExportConfiguration.Default());
        }

        [TestMethod]
        public void OptionsBorders()
        {
            ExcelWorksheet ws = TestHelper.GetWorkSheet();
            ExportExcel(ws, new HtmlExportConfiguration() { Borders = true });
        }

        [TestMethod]
        public void OptionsFill()
        {
            ExcelWorksheet ws = TestHelper.GetWorkSheet();
            ExportExcel(ws, new HtmlExportConfiguration() { Fill = true });
        }

        [TestMethod]
        public void OptionsBordersAndFill()
        {
            ExcelWorksheet ws = TestHelper.GetWorkSheet();
            ExportExcel(ws, new HtmlExportConfiguration() { Borders = true, Fill = true });
        }

        [TestMethod]
        public void OptionsNotSpecified()
        {
            ExcelWorksheet ws = TestHelper.GetWorkSheet();
            var html1 = ws.ToHtml(true);
            Assert.IsNotNull(html1, $"html generated without configuration is null");
            var html2 = ExportExcel(ws, HtmlExportConfiguration.Default());
            Assert.AreEqual(html1, html2, "HtmlExportConfiguration unspecified is not the same as HtmlExportConfiguration.Default()");
        }

        private string ExportExcel(ExcelWorksheet ws, HtmlExportConfiguration configuration)
        {
            var html = ws.ToHtml(configuration, true);
            Assert.IsNotNull(html, $"html generated with HtmlExportConfiguration {configuration.ToString()} is null");
            TestHelper.CreateFile(configuration.ToString(), "html", html);
            return html;
        }
    }
}
