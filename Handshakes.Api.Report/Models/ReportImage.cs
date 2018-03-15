using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace Handshakes.Api.Report.Models
{
	public class ReportImage : ReportElement, IReportReplaceable
	{
		public byte[] Value
		{
			get { return GetValue<byte[]>(); }
			set { SetValue(value); }
		}

		public OpenXmlCompositeElement[] Replace(OpenXmlCompositeElement element)
		{
			throw new NotImplementedException();
		}
	}
}
