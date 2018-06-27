using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using Xenirio.Component.Gutenberg.Model;


namespace Xenirio.Component.Gutenberg.Test
{
    [TestClass]
    public class FragmentedVariableSpec
    {
        [TestMethod]
        public void Incorrect_VariableName_ThrowsArgumentOut()
        {
            try
            {
                var sourcefile = Environment.CurrentDirectory + @"\Resources\SampleFragmentedVariable.docx";
                var outfile = Environment.CurrentDirectory + @"\Resources\SampleFragmentedVariableTest.docx";
                if (File.Exists(outfile))
                    File.Delete(outfile);
                File.Copy(sourcefile, outfile);
                var document = new ReportDocument();

                document.InjectReportElement(new ReportLabel() { Key = "Entity.Names", Value = "IPO of Ezra Holdings Limited" });

                document.Save(outfile);
            }
            catch (Exception e)
            {
                Assert.AreEqual("Variables not match with Fields", e.Message);
            }
        }

        [TestMethod]
        public void Incorrect_VariableType_ThrowsArgumentOut()
        {
            try
            {
                var sourcefile = Environment.CurrentDirectory + @"\Resources\SampleFragmentedVariable.docx";
                var outfile = Environment.CurrentDirectory + @"\Resources\SampleFragmentedVariableTest.docx";
                if (File.Exists(outfile))
                    File.Delete(outfile);
                File.Copy(sourcefile, outfile);
                var document = new ReportDocument();

                document.InjectReportElement(new ReportTable()
                {
                    Key = "Entity.Name",
                    Elements = new ReportLabel[][] {
                        new ReportLabel[] {
                            new ReportLabel() { Value = "Coopers" },
                            new ReportLabel() { Value = "Firm" },
                            new ReportLabel() { Value = "Chiang Mai" },
                            new ReportLabel() { Value = "John" },
                            new ReportLabel() { Value = "Good" },
                            new ReportLabel() { Value = "Chiang Mai" },
                            new ReportLabel() { Value = "John" },
                            new ReportLabel() { Value = "Good" },
                        }
                    }
                });
                document.Save(outfile);
            }
            catch(Exception e)
            {
                Assert.AreEqual("Specified argument was out of the range of valid values.\r\nParameter name: index", e.Message);
            }
           
        }

        [TestMethod]
        public void Duplicate_VariableName_ThrowArgumentOut()
        {
            try
            {
                var sourcefile = Environment.CurrentDirectory + @"\Resources\SampleFragmentedVariable.docx";
                var outfile = Environment.CurrentDirectory + @"\Resources\SampleFragmentedVariableTest.docx";
                if (File.Exists(outfile))
                    File.Delete(outfile);
                File.Copy(sourcefile, outfile);
                var document = new ReportDocument();

                document.InjectReportElement(new ReportLabel() { Key = "Entity.Name", Value = "IPO of Ezra Holdings Limited" });
                document.InjectReportElement(new ReportLabel() { Key = "Entity.Name", Value = "DBS Banking" });

                document.Save(outfile);

            }
            catch(Exception e)
            {
                Assert.AreEqual("An item with the same key has already been added.", e.Message);
            }

        }

        [TestMethod]
        public void Incorrect_VariableSize_ThrowArgumentOut()
        {
            try
            {
                var sourcefile = Environment.CurrentDirectory + @"\Resources\SampleFragmentedVariable.docx";
                var outfile = Environment.CurrentDirectory + @"\Resources\SampleFragmentedVariableTest.docx";
                if (File.Exists(outfile))
                    File.Delete(outfile);
                File.Copy(sourcefile, outfile);
                var document = new ReportDocument();

                document.InjectReportElement(new ReportLabels() { Key = "Entity.Name", Values = new string[] { "IPO of Ezra Holdings Limited", "IPO of Ezra" } });


                document.Save(outfile);

            }
            catch (Exception e)
            {
                Assert.AreEqual("An item with the same key has already been added.", e.Message);
            }

        }




    }

}
