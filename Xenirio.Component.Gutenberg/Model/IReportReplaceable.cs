using DocumentFormat.OpenXml.Wordprocessing;

namespace Xenirio.Component.Gutenberg.Model
{
	public interface IReportReplaceable
	{
		void Replace(Run element);
    }
}
