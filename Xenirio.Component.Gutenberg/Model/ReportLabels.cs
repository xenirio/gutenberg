using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Xenirio.Component.Gutenberg.Model
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
