using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Handshakes.Api.Report.Tests
{
	[TestClass]
	public class ReportDocumentSpec
	{
		private ReportDocument document;

		[TestInitialize]
		public void Initialize()
		{
			var sourcefile = Environment.CurrentDirectory + @"\Resources\Sample.docx";
			var outfile = Environment.CurrentDirectory + @"\Resources\SampleTest.docx";
			if (File.Exists(outfile))
				File.Delete(outfile);
			File.Copy(sourcefile, outfile);
			document = new ReportDocument(outfile);
		}

		[TestMethod]
		public void Should_Save_Document()
		{
			document.InjectReportLabel("Header.Entity.Name", "Coopers");
			document.InjectReportLabel("Content", "ABCD");
			document.Save();
		}
	}
}
