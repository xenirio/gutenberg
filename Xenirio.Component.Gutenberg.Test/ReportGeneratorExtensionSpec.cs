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

        [TestMethod]
        public void Should_Support_Paragraph_Style()
        {
            var json = JObject.FromObject(new
            {
                Content = new
                {
                    Should = new
                    {
                        Red = "This is Red sentense.",
                        Green = "This is Green sentense.",
                        Blue = "This is Blue sentense.",
                        Bold = "This is Bold sentense.",
                        Italic = "This is Italic sentense."
                    }
                },
                Style = new
                {
                    Content = new
                    {
                        Should = new
                        {
                            Red = new
                            {
                                Color = "#FF0000"
                            },
                            Green = new
                            {
                                Color = "#00FF00"
                            },
                            Blue = new
                            {
                                Color = "#0000FF"
                            },
                            Bold = new
                            {
                                Bold = true
                            },
                            Italic = new
                            {
                                Italic = true
                            }
                        }
                    }
                }
            });
            var template = Environment.CurrentDirectory + @"\Resources\SampleStyling.docx";
            var report = new ReportGenerator(template);
            report.setJsonObject(json);
            var outputPath = string.Format(@"{0}\{1}{2}.docx", Path.GetDirectoryName(template), Path.GetFileNameWithoutExtension(template), DateTime.Now.Ticks);
            report.GenerateToFile(outputPath);
            Assert.IsTrue(File.Exists(outputPath));
        }

        [TestMethod]
        public void Should_Support_Table_Style()
        {
            var json = JObject.FromObject(new
            {
                Content = new
                {
                    Should = new
                    {
                        Red = "This is Red sentense.",
                        Green = "This is Green sentense.",
                        Blue = "This is Blue sentense.",
                        Bold = "This is Bold sentense.",
                        Italic = "This is Italic sentense."
                    },
                    Table = new
                    {
                        Should = new
                        {
                            Fancy = new string[][]
                            {
                                new string[] { "This is Red cell", "This is Green cell", "This is Blue cell"},
                                new string[] { "This is Red Bold cell", "This is Green Bold cell", "This is Bule Bold cell"},
                                new string[] { "This is Red Bold Italic cell", "This is Green Bold Italic cell", "This is Blue Bold Italic cell"},
                                new string[] { "This is cell", "This is cell", "This is cell"},
                            }
                        }
                    }
                },
                Style = new
                {
                    Content = new
                    {
                        Should = new
                        {
                            Red = new
                            {
                                Color = "#FF0000"
                            },
                            Green = new
                            {
                                Color = "#00FF00"
                            },
                            Blue = new
                            {
                                Color = "#0000FF"
                            },
                            Bold = new
                            {
                                Bold = true
                            },
                            Italic = new
                            {
                                Italic = true
                            }
                        },
                        Table = new
                        {
                            Should = new
                            {
                                Fancy = new object[][]
                                {
                                    new object[]
                                    {
                                        new { Color = "#FF0000" },
                                        new { Color = "#00FF00" },
                                        new { Color = "#0000FF" },
                                    },
                                    new object[]
                                    {
                                        new { Color = "#FF0000", Bold = true },
                                        new { Color = "#00FF00", Bold = true },
                                        new { Color = "#0000FF", Bold = true },
                                    },
                                    new object[]
                                    {
                                        new { Color = "#FF0000", Bold = true, Italic = true },
                                        new { Color = "#00FF00", Bold = true, Italic = true },
                                        new { Color = "#0000FF", Bold = true, Italic = true },
                                    },
                                    new object[]
                                    {
                                        new {},
                                        new {},
                                        new {},
                                    }
                                }
                            }
                        }
                    }
                }
            });
            var template = Environment.CurrentDirectory + @"\Resources\SampleStyling.docx";
            var report = new ReportGenerator(template);
            report.setJsonObject(json);
            var outputPath = string.Format(@"{0}\{1}{2}.docx", Path.GetDirectoryName(template), Path.GetFileNameWithoutExtension(template), DateTime.Now.Ticks);
            report.GenerateToFile(outputPath);
            Assert.IsTrue(File.Exists(outputPath));
        }
    }
}
