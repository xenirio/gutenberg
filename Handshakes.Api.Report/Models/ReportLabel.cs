using System;
using System.Collections.Generic;
using System.Text;

namespace Handshakes.Api.Report.Models
{
	public class ReportLabel : ReportElement
	{
		public string Value
		{
			get { return GetValue<string>(); }
			set { SetValue(value); }
		}
	}
}
