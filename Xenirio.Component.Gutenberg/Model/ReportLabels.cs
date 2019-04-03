using System;
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
                var token = Values[i].Split(new string[] { "\r\n" }, StringSplitOptions.None);
                for (var j = 0; j < token.Length; j++)
                {
                    tmpParagraph.AppendChild(new Run(new Text(token[j])));
                    if (j < token.Length - 1)
                        tmpParagraph.AppendChild(new Run(new Break()));
                }
				
			}
			paragraph.Remove();
		}
	}
}
