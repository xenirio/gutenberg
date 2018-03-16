using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace Handshakes.Api.Report.Model
{
	internal class ReportImage : ReportElement, IReportReplaceable
	{
		public byte[] Value
		{
			get { return GetValue<byte[]>(); }
			set { SetValue(value); }
		}

		public void Replace(Run element)
		{
			throw new NotImplementedException();
		}
	}
}
