using Handshakes.Api.Report.Converter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace Handshakes.Api.Report.Tests
{
	[TestClass]
	public class PDFReportConverterSpec
	{
		[TestMethod]
		public void Should_Convert_Document()
		{
			var sourcefile = Environment.CurrentDirectory + @"\Resources\Sample.docx";
			var outputPath = string.Format(@"{0}\{1}.pdf", Path.GetDirectoryName(sourcefile), Path.GetFileNameWithoutExtension(sourcefile));

			PDFReportConverter.ConvertToFile(sourcefile, outputPath);
		}
	}
}
