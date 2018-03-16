using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Handshakes.Api.Report.Converter
{
	interface IReportConvertable
	{
		byte[] Convert(string filePath);
	}
}
