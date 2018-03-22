using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Handshakes.Api.Report.Model;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handshakes.Api.Report
{
	internal class ReportDocument
	{
		private Dictionary<string, IReportReplaceable> variables = new Dictionary<string, IReportReplaceable>();

		public void InjectReportElement(ReportElement element)
		{
			variables.Add(string.Format(@"DOCVARIABLE  {0}", element.Key), (IReportReplaceable)element);
		}

		private WordprocessingDocument ReplaceDocument(WordprocessingDocument document)
		{
			// replace header variable
			var headers = document.MainDocumentPart.HeaderParts.ToList();
			foreach (var header in headers)
			{
				Replaces(header.Header.Descendants<FieldCode>().ToArray());
			}

			// replace body variable
			Replaces(document.MainDocumentPart.RootElement.Descendants<FieldCode>().ToArray());

			// replace footer variable
			var footers = document.MainDocumentPart.FooterParts.ToList();
			foreach (var footer in footers)
			{
				Replaces(footer.Footer.Descendants<FieldCode>().ToArray());
			}
			return document;
		}

		public void Save(string filePath)
		{
			if (!File.Exists(filePath))
				throw new FileNotFoundException("File not found", filePath);
			var document = WordprocessingDocument.Open(filePath, true);
			ReplaceDocument(document);
			document.Close();
		}

		public byte[] Save(byte[] bytes)
		{
			using (MemoryStream mem = new MemoryStream())
			{
				mem.Write(bytes, 0, bytes.Length);
				using (WordprocessingDocument document = WordprocessingDocument.Open(mem, true))
				{
					ReplaceDocument(document);
					document.Save();
					return mem.ToArray();
				}
			}
		}

		private void Replaces(FieldCode[] fields)
		{
			foreach (var field in fields)
			{
				var key = field.Text.Trim();
				if (variables.ContainsKey(key))
				{
					variables[key].Replace((Run)field.Parent);
				}
			}
		}
	}
}
