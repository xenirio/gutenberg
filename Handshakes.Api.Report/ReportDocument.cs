using DocumentFormat.OpenXml.Packaging;
using Handshakes.Api.Report.Models;
using System;
using System.Collections.Generic;
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
		private List<ReportElement> variables = new List<ReportElement>();

		public ReportDocument(string filePath)
		{
			if (!File.Exists(filePath))
				throw new FileNotFoundException("File not found", filePath);
			document = WordprocessingDocument.Open(filePath, true);
		}

		public void InjectReportLabel(string key, string value)
		{
			variables.Add(new ReportLabel() { Key = key, Value = value });
		}

		public void Save()
		{
			string header = string.Empty;
			var headerPart = document.MainDocumentPart.HeaderParts.FirstOrDefault();
			if (headerPart != null) {
				using (StreamReader reader = new StreamReader(headerPart.GetStream()))
				{
					header = reader.ReadToEnd();
				}
			}

			string body = string.Empty;
			using (StreamReader sr = new StreamReader(document.MainDocumentPart.GetStream()))
			{
				body = sr.ReadToEnd();
			}

			foreach (var variable in variables)
			{
				if (variable.GetType() == typeof(ReportLabel))
				{
					var label = (ReportLabel)variable;
					body = body.Replace(label.Key, label.Value);
					header = header.Replace(label.Key, label.Value);
				}
			}
			using (StreamWriter writer = new StreamWriter(document.MainDocumentPart.GetStream(FileMode.Create)))
			{
				writer.Write(body);
			}
			if (!string.IsNullOrWhiteSpace(header))
			{
				using (StreamWriter writer = new StreamWriter(document.MainDocumentPart.HeaderParts.First().GetStream(FileMode.Create)))
				{
					writer.Write(header);
				}
			}
		}
	}
}
