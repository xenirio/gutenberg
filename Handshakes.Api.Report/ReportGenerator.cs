using Handshakes.Api.Report.Converter;
using Handshakes.Api.Report.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Handshakes.Api.Report
{
	public enum OutputFormat
	{
		PDF = 1
	}
	public class ReportGenerator
	{
		private string _templatePath;
		private ReportDocument _document;

		public ReportGenerator(string templatePath)
		{
			_document = new ReportDocument();
			_templatePath = templatePath;
		}

		public void setParagraph(string key, string content)
		{
			_document.InjectReportElement(new ReportLabel() { Key = key, Value = content });
		}

		public void setParagraphs(string key, string[] content)
		{
			_document.InjectReportElement(new ReportLabels() { Key = key, Values = content });
		}

		public void setTable(string key, string[][] content)
		{
			_document.InjectReportElement(new ReportTable() { Key = key, Values = content });
		}

		private string saveDocumentToFile(string filePath)
		{
			var fileName = Path.GetFileNameWithoutExtension(filePath);
			var outFile = string.Format(@"{0}\{1}{2}{3}", Path.GetDirectoryName(filePath), fileName, DateTime.Now.Ticks, Path.GetExtension(filePath));
			if (File.Exists(outFile))
				File.Delete(outFile);
			File.Copy(filePath, outFile);
			_document.Save(outFile);
			return outFile;
		}

		private byte[] saveDocumentToBytes(string filePath)
		{
			return _document.Save(File.ReadAllBytes(filePath));
		}

		public byte[] GenerateToByte(OutputFormat format = OutputFormat.PDF)
		{
			byte[] bytes = null;
			var outFile = saveDocumentToBytes(_templatePath);
			switch (format) {
				case OutputFormat.PDF:
					bytes = PDFReportConverter.ConvertToByte(outFile);
					break;
			}
			return bytes;
		}

		public void GenerateToFile(string outputPath, OutputFormat format = OutputFormat.PDF)
		{
			var outFile = saveDocumentToBytes(_templatePath);
			switch (format)
			{
				case OutputFormat.PDF:
					PDFReportConverter.ConvertToFile(outFile, outputPath);
					break;
			}
		}
	}
}
