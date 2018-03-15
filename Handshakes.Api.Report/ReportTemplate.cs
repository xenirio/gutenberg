using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Handshakes.Api.Report
{
	internal class ReportTemplate
	{
		WordprocessingDocument document;

		public ReportTemplate(string filePath)
		{
			if (!File.Exists(filePath))
				throw new FileNotFoundException("File not found", filePath);
			document = WordprocessingDocument.Open(filePath, false);
		}

		public string[] DocVariables
		{
			get
			{
				var variables = new List<string>();
				var headerContent = string.Join("", document.MainDocumentPart.HeaderParts.Select(h => h.Header.InnerText));
				var bodyContent = document.MainDocumentPart.Document.InnerText;
				Regex regexText = new Regex("\\{.*?\\}");
				foreach (Match match in regexText.Matches(headerContent + bodyContent))
				{
					foreach (Capture capture in match.Captures)
					{
						variables.Add(capture.Value);
					}
				}
				return variables.ToArray();
			}
		}
	}
}
