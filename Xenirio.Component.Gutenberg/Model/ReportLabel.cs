using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;

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
				//element.ElementsBefore().SelectMany(e => e.Descendants<FieldChar>()).FirstOrDefault(c => c.FieldCharType == FieldCharValues.Begin)?.Parent.Remove();
				element.PreviousSibling<Run>()?.Remove();
				element.NextSibling<Run>()?.Remove();
				var token = Value.Split(new string[] { "\r\n" }, StringSplitOptions.None);
				for (var i = 0; i < token.Length; i++)
                {
					var text = (Run)element.Clone(); //new Run(new Text(token[i]));
					text.RemoveAllChildren<FieldCode>();
					text.AppendChild(new Text(token[i]));
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
					if (i == 0)
					{
						paragraph.ReplaceChild(text, element);
					}
					else
					{
						if (i < token.Length - 1)
							paragraph.AppendChild(new Run(new Break()));
						paragraph.AppendChild(text);
					}
				}
			}
		}
	}
}
