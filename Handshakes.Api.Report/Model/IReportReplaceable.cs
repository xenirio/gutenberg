using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Handshakes.Api.Report.Model
{
	interface IReportReplaceable
	{
		void Replace(Run element);
	}
}
