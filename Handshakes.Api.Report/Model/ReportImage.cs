using System;
using DocumentFormat.OpenXml.Wordprocessing;

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
