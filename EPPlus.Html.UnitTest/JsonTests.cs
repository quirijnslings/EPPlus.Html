using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml;
using System.Text;
using EPPlus.Html;
using EPPlus.Html.Json;

namespace Test
{
    [TestClass]
    public class JsonTests
    {
        [TestMethod]
        public void ExportAsIs()
        {
            ExcelWorksheet ws = TestHelper.GetWorkSheet();
            var json = ws.ToJson();
            Assert.IsNotNull(json);
            TestHelper.CreateFile("asis", "json", json);
        }

        [TestMethod]
        public void ExportIndented()
        {
            ExcelWorksheet ws = TestHelper.GetWorkSheet();
            var json = ws.ToJson(JsonStyle.Indented);
            Assert.IsNotNull(json);
            TestHelper.CreateFile("indented", "json", json);
        }

        [TestMethod]
        public void ExportLowerCase()
        {
            ExcelWorksheet ws = TestHelper.GetWorkSheet();
            var json = ws.ToJson(JsonStyle.Lc);
            Assert.IsNotNull(json);
            TestHelper.CreateFile("lc", "json", json);
        }

        [TestMethod]
        public void ExportLowerCaseIndented()
        {
            ExcelWorksheet ws = TestHelper.GetWorkSheet();
            var json = ws.ToJson(JsonStyle.LcIndented);
            Assert.IsNotNull(json);
            TestHelper.CreateFile("lcindented", "json", json);
        }
        [TestMethod]
        public void ExportLowerCaseFirst()
        {
            ExcelWorksheet ws = TestHelper.GetWorkSheet();
            var json = ws.ToJson(JsonStyle.LcFirst);
            Assert.IsNotNull(json);
            TestHelper.CreateFile("lcfirst", "json", json);
        }
        [TestMethod]
        public void ExportLowerCaseFirstIndented()
        {
            ExcelWorksheet ws = TestHelper.GetWorkSheet();
            var json = ws.ToJson(JsonStyle.LcFirstIndented);
            Assert.IsNotNull(json);
            TestHelper.CreateFile("lcfirstindented", "json", json);
        }
    }
}
