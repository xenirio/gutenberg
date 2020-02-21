using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace Xenirio.Component.Gutenberg.Model
{
	public class ReportLabelStyle
	{
		public string Color { get; set; }
		public bool? Bold { get; set; }
		public bool? Italic { get; set; }
	}

    public class ReportLabel : ReportElement, IReportReplaceable
	{
		public string Value
		{
			get { return GetValue<string>(); }
			set { SetValue(value); }
		}

		public ReportLabelStyle Style { get; set; }

		public void Replace(Run element)
		{
			var paragraph = (Paragraph)element.Parent;
            if (paragraph != null)
            {
                paragraph.RemoveAllChildren<Run>();
                var token = Value.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                for (var i = 0; i < token.Length; i++)
                {
					var text = new Run(new Text(token[i]));
					if (Style != null) {
						text.RunProperties = new RunProperties();
						if (!string.IsNullOrWhiteSpace(Style.Color))
						{
							text.RunProperties.Color = new Color()
							{
								Val = Style.Color
							};
						}
						if (Style.Bold == true)
							text.RunProperties.Bold = new Bold();
						if (Style.Italic == true)
							text.RunProperties.Italic = new Italic();
					}
					paragraph.AppendChild(text);
                    if (i < token.Length - 1)
                        paragraph.AppendChild(new Run(new Break()));
                }
			}
		}
	}
}
