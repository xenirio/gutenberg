using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace Handshakes.Api.Report.Models
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
				paragraph.AppendChild(new Run(new Text(Value)));
			}
		}
	}
}
