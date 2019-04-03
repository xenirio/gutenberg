using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace Xenirio.Component.Gutenberg.Model
{
    public class ReportLabel : ReportElement, IReportReplaceable
	{
		public string Value
		{
			get { return GetValue<string>(); }
			set { SetValue(value); }
		}

		public void Replace(Run element)
		{
			var paragraph = (Paragraph)element.Parent;
            if (paragraph != null)
            {
                paragraph.RemoveAllChildren<Run>();
                var token = Value.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                for (var i = 0; i < token.Length; i++)
                {
                    paragraph.AppendChild(new Run(new Text(token[i])));
                    if (i < token.Length - 1)
                        paragraph.AppendChild(new Run(new Break()));
                }
			}
		}
	}
}
