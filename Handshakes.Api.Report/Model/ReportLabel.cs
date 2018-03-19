using DocumentFormat.OpenXml.Wordprocessing;

namespace Handshakes.Api.Report.Model
{
	internal class ReportLabel : ReportElement, IReportReplaceable
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
				paragraph.AppendChild(new Run(new Text(Value)));
			}
		}
	}
}
