using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Xenirio.Component.Gutenberg.Model
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
