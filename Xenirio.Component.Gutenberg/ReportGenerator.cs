using Xenirio.Component.Gutenberg.Converter;
using Xenirio.Component.Gutenberg.Model;
using System;
using System.IO;
using System.Collections.Generic;

namespace Xenirio.Component.Gutenberg
{
	public enum OutputFormat
	{
        Word = 0,
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

		public void setTableParagraph(string key, string[][] content)
		{
            var values = new List<ReportElement[]>();
            for (var i = 0; i < content.Length; i++)
            {
                var labels = new List<ReportLabel>();
                for (var j = 0; j < content[i].Length; j++)
                {
                    labels.Add(new ReportLabel() { Value = content[i][j] });
                }
                values.Add(labels.ToArray());
            }

            _document.InjectReportElement(new ReportTable() { Key = key, Elements = values.ToArray() });
		}

        public void setImage(string key, byte[] content)
        {
            _document.InjectReportElement(new ReportImage() { Key = key, Value = content });
        }

        public void setTableImage(string key, byte[][][] content)
        {
            var values = new List<ReportElement[]>();
            for (var i = 0; i < content.Length; i++)
            {
                var labels = new List<ReportImage>();
                for (var j = 0; j < content[i].Length; j++)
                {
                    labels.Add(new ReportImage() { Value = content[i][j] });
                }
                values.Add(labels.ToArray());
            }

            _document.InjectReportElement(new ReportTable() { Key = key, Elements = values.ToArray() });
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

		public byte[] GenerateToByte(OutputFormat format = OutputFormat.Word)
		{
			byte[] bytes = null;
			var outFile = saveDocumentToBytes(_templatePath);
			switch (format) {
				case OutputFormat.PDF:
					bytes = PDFReportConverter.ConvertToByte(outFile);
					break;
                default:
                    return outFile;

            }
			return bytes;
		}

		public void GenerateToFile(string outputPath, OutputFormat format = OutputFormat.Word)
		{
			switch (format)
			{
				case OutputFormat.PDF:
                    var outFile = saveDocumentToBytes(_templatePath);
                    PDFReportConverter.ConvertToFile(outFile, outputPath);
					break;
                default:
                    saveDocumentToFile(_templatePath);
                    break;
            }
		}
	}
}
