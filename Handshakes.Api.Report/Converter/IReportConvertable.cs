namespace Handshakes.Api.Report.Converter
{
	interface IReportConvertable
	{
		byte[] Convert(string filePath);
	}
}
