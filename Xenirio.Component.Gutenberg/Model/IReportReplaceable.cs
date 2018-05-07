using DocumentFormat.OpenXml.Wordprocessing;

namespace Xenirio.Component.Gutenberg.Model
{
	interface IReportReplaceable
	{
		void Replace(Run element);
    }
}
