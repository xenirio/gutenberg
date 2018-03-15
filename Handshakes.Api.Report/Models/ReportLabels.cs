using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace Handshakes.Api.Report.Models
{
	public class ReportLabels : ReportElement, IReportReplaceable
	{
		public string[] Values
		{
			get { return GetValue<string[]>(); }
			set { SetValue(value); }
		}

		public OpenXmlCompositeElement[] Replace(OpenXmlCompositeElement element)
		{
			var paragraphs = new List<Paragraph>();
			foreach (var v in Values)
			{
				var label = new ReportLabel() { Key = "", Value = v };
				paragraphs.Add((Paragraph)label.Replace(element)[0]);
			}
			return paragraphs.ToArray();
		}
	}
}
