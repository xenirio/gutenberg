using DocumentFormat.OpenXml.Wordprocessing;

namespace Handshakes.Api.Report.Model
{
	interface IReportReplaceable
	{
		void Replace(Run element);
	}
}
