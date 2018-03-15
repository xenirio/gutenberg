using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Handshakes.Api.Report.Models
{
	interface IReportReplaceable
	{
		OpenXmlCompositeElement[] Replace(OpenXmlCompositeElement element);
	}
}
