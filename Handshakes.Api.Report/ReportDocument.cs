using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Handshakes.Api.Report.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Handshakes.Api.Report
{
	internal class ReportDocument
	{
		private WordprocessingDocument document;
		private Dictionary<string, IReportReplaceable> variables = new Dictionary<string, IReportReplaceable>();

		public ReportDocument(string filePath)
		{
			if (!File.Exists(filePath))
				throw new FileNotFoundException("File not found", filePath);
			document = WordprocessingDocument.Open(filePath, true);
		}

		public void InjectReportElement(ReportElement element)
		{
			variables.Add(string.Format(@"DOCVARIABLE  {0}", element.Key), (IReportReplaceable)element);
		}

		public void Save()
		{
			// replace header variable
			var header = document.MainDocumentPart.HeaderParts.FirstOrDefault();
			if(header != null)
			{
				Replaces(header.Header.Descendants<FieldCode>().ToArray());
			}

			// replace body variable
			Replaces(document.MainDocumentPart.RootElement.Descendants<FieldCode>().ToArray());
			document.Close();
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
