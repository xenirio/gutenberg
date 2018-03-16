using Handshakes.Api.Report.Converter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Handshakes.Api.Report.Tests
{
	[TestClass]
	public class PDFReportConverterSpec
	{
		[TestMethod]
		public void Should_Convert_Document()
		{
			var sourcefile = Environment.CurrentDirectory + @"\Resources\Sample.docx";

			PDFReportConverter.ConvertToFile(sourcefile);
		}
	}
}
