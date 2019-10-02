using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;
using Xenirio.Component.Gutenberg.Extensions;

namespace Xenirio.Component.Gutenberg.Test
{
    [TestClass]
    public class ReportGeneratorExtensionSpec
    {
        [TestMethod]
        public void Should_Set_Json_Token_Value()
        {
            var json = JObject.FromObject(new
            {
                Type = "Mock",
                Header = new
                {
                    Title = "Foo"
                }
            });
            var results = json.DeserializeAndFlatten();
            Assert.AreEqual("Mock", ((JValue)results["Type"]).Value);
            Assert.AreEqual("Foo", ((JValue)results["Header.Title"]).Value);
        }

        [TestMethod]
        public void Should_Set_Json_Token_Values()
        {
            var json = JObject.FromObject(new
            {
                Content = new
                {
                    Addresses = new string[]
                    {
                        "Foo",
                        "Bar"
                    }
                }
            });
            var results = json.DeserializeAndFlatten();
            var addresses = results["Content.Addresses"];
            Assert.AreEqual(addresses.Type, JTokenType.Array);
            CollectionAssert.AreEqual(new string[] { "Foo", "Bar" }, addresses.Children().Select(t => ((JValue)t).Value.ToString()).ToArray());
        }

        [TestMethod]
        public void Should_Set_Json_Token_Table()
        {
            var json = JObject.FromObject(new
            {
                Content = new
                {
                    Financial = new string[][]
                    {
                        new string[] {"1.00", "2.00", "3.00"},
                        new string[] {"2.00", "3.00", "1.00"},
                        new string[] {"3.00", "2.00", "1.00"},
                    }
                }
            });
            var results = json.DeserializeAndFlatten();
            var fin = results["Content.Financial"];
            Assert.AreEqual(fin.Type, JTokenType.Array);
            Assert.AreEqual(fin[0].Type, JTokenType.Array);
            CollectionAssert.AreEqual(new string[] { "1.00", "2.00", "3.00" }, fin[0].Children().Select(t => ((JValue)t).Value.ToString()).ToArray());
        }

        [TestMethod]
        public void Should_Set_ReportGenerator_Json_Object()
        {
            var json = JObject.FromObject(new
            {
                Header = new
                {
                    Title = "Foo"
                },
                Content = new
                {
                    Type = "Mock",
                    Addresses = new string[]
                    {
                        "Foo",
                        "Bar"
                    },
                    Financial = new string[][]
                    {
                        new string[] {"1.00", "2.00", "3.00"},
                        new string[] {"2.00", "3.00", "1.00"},
                        new string[] {"3.00", "2.00", "1.00"},
                    }
                }
            });
            var template = Environment.CurrentDirectory + @"\Resources\SampleJsonTemplate.docx";
            var report = new ReportGenerator(template);
            report.setJsonObject(json);
            var outputPath = string.Format(@"{0}\{1}{2}.docx", Path.GetDirectoryName(template), Path.GetFileNameWithoutExtension(template), DateTime.Now.Ticks);
            report.GenerateToFile(outputPath);
            Assert.IsTrue(File.Exists(outputPath));
        }
    }
}
