using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace Handshakes.Api.Report.Model
{
	internal class ReportLabels : ReportElement, IReportReplaceable
	{
		public string[] Values
		{
			get { return GetValue<string[]>(); }
			set { SetValue(value); }
		}

		public void Replace(Run element)
		{
			var paragraph = (Paragraph)element.Parent;
			paragraph.Descendants<Run>().ToList().ForEach(c => c.Remove());
			var elementAt = paragraph;
			var root = paragraph.Parent;
			for (var i = 0; i < Values.Length; i++)
			{
				var tmpParagraph = (Paragraph)paragraph.Clone();
				root.InsertAfter(tmpParagraph, elementAt);
				elementAt = tmpParagraph;
				tmpParagraph.AppendChild(new Run(new Text(Values[i])));
			}
			paragraph.Remove();
		}
	}
}
