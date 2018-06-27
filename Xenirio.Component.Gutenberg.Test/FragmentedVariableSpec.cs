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
        public void Should_Merge_Fragmented_Variable_Of_Page_Break_Table()
        {
            var sourcefile = Environment.CurrentDirectory + @"\Resources\SampleFragmentedVariable.docx";
            var outfile = Environment.CurrentDirectory + @"\Resources\SampleFragmentedVariableTest.docx";
            if (File.Exists(outfile))
                File.Delete(outfile);
            File.Copy(sourcefile, outfile);
            var document = new ReportDocument();
            document.InjectReportElement(new ReportLabel() { Key = "Entity.Name", Value = "IPO of Ezra Holdings Limited & Coopers" });
            document.Save(outfile);
        }
    }

}
