using Xenirio.Component.Gutenberg.Model;
using System;
using System.IO;
using System.Collections.Generic;

namespace Xenirio.Component.Gutenberg
{
	public enum OutputFormat
	{
        Word = 0
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

        private void saveDocumentToFile(string filePath)
		{
			if (File.Exists(filePath))
				File.Delete(filePath);
			File.Copy(_templatePath, filePath);
			_document.Save(filePath);
		}

		public byte[] GenerateToByte(OutputFormat format = OutputFormat.Word)
		{
			var outBytes = _document.Save(File.ReadAllBytes(_templatePath));
            switch (format) {
                default:
                    return outBytes;
            }
		}

		public void GenerateToFile(string outputPath, OutputFormat format = OutputFormat.Word)
		{
			switch (format)
			{
                default:
                    saveDocumentToFile(outputPath);
                    break;
            }
		}
	}
}
