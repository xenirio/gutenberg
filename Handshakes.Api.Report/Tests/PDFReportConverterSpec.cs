using Handshakes.Api.Report.Converter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
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
			var outputPath = string.Format(@"{0}\{1}.pdf", Path.GetDirectoryName(sourcefile), Path.GetFileNameWithoutExtension(sourcefile));

			PDFReportConverter.ConvertToFile(sourcefile, outputPath);
		}
	}
}
