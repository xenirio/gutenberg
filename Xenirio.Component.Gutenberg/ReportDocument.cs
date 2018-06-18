using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xenirio.Component.Gutenberg.Model;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Xenirio.Component.Gutenberg
{
	internal class ReportDocument
	{
		private Dictionary<string, IReportReplaceable> variables = new Dictionary<string, IReportReplaceable>();

		public void InjectReportElement(ReportElement element)
		{
			variables.Add(string.Format(@"DOCVARIABLE  {0}", element.Key), (IReportReplaceable)element);
		}
        
		private void ReplaceDocument(WordprocessingDocument document)
		{
			// replace header variable
			var headers = document.MainDocumentPart.HeaderParts.ToList();
			foreach (var header in headers)
			{
                var headerFields = header.Header.Descendants<FieldCode>().ToArray();
                Replace(headerFields);
                Clean(headerFields);

            }

            // replace body variable
            var bodyFields = document.MainDocumentPart.RootElement.Descendants<FieldCode>().ToArray();
            Replace(bodyFields);
            Clean(bodyFields);

            // replace footer variable
            var footers = document.MainDocumentPart.FooterParts.ToList();
			foreach (var footer in footers)
			{
                var footerFields = footer.Footer.Descendants<FieldCode>().ToArray();
                Replace(footerFields);
                Clean(footerFields);
			}

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
					document.Close();
					return mem.ToArray();
				}
			}
		}

		private void Replace(FieldCode[] fields)
		{
            foreach (var field in fields)
            {
                var key = field.Text.Trim();
                IReportReplaceable variable;
                if (variables.ContainsKey(key))
                {
                    variable = variables[key];
                    variable.Replace((Run)field.Parent);
                }
            }
        }

        private void Clean(FieldCode[] fields)
        {
            foreach (var field in fields)
            {
                var key = field.Text.Trim();
                var varKey = key.Replace("DOCVARIABLE", "").Trim();
                if (varKey.StartsWith("Template"))
                {
                    field.Ancestors<TableRow>().First().Remove();
                }
                else
                {
                    (new ReportLabel() { Key = key, Value = "" }).Replace((Run)field.Parent);
                }
            }
        }
	}
}
