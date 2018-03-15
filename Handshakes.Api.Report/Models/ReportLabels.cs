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

		public void Replace(Run element)
		{
			var tmpParagraph = (Paragraph)element.Parent.Clone();
			var elementAt = element.Parent;
			var root = element.Parent.Parent;
			for (var i = 0; i < Values.Length; i++)
			{
				Paragraph paragraph = (Paragraph)element.Parent;
				if (i > 0)
				{
					paragraph = tmpParagraph;
					root.InsertAfter(paragraph, elementAt);
					elementAt = paragraph;
				}
				var label = new ReportLabel() { Key = "", Value = Values[i] };
				label.Replace((Run)paragraph.Descendants<FieldCode>().ToArray()[0].Parent);
			}
		}
	}
}
