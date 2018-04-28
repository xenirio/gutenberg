using Spire.Doc;
using System.IO;

namespace Xenirio.Component.Gutenberg.Converter
{
	internal class PDFReportConverter : IReportConvertable
	{
		public byte[] Convert(string filePath)
		{
			//Load Document
			Document document = new Document();
			document.LoadFromFile(filePath);

			byte[] pdf = null;
			//Convert Word to PDF
			using (MemoryStream stream = new MemoryStream())
			{
				document.SaveToStream(stream, FileFormat.PDF);
				//save to byte array
				pdf = stream.ToArray();
			}
			return pdf;
		}

		public static byte[] ConvertToByte(string filePath)
		{
			return new PDFReportConverter().Convert(filePath);
		}

		public static byte[] ConvertToByte(byte[] file)
		{
			//Load Document
			Document document = new Document();
			using (MemoryStream stream = new MemoryStream(file))
			{
				document.LoadFromStream(stream, FileFormat.Auto);
			}

			//Convert Word to PDF
			using (MemoryStream stream = new MemoryStream())
			{
				document.SaveToStream(stream, FileFormat.PDF);
				//save to byte array
				return stream.ToArray();
			}
		}

		public static void ConvertToFile(string filePath, string outputPath)
		{
			var outFile = string.Format(@"{0}\{1}.pdf", Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath));

			//Load Document
			Document document = new Document();
			document.LoadFromFile(filePath);

			document.SaveToFile(outFile, FileFormat.PDF);
		}

		public static void ConvertToFile(byte[] file, string outputPath)
		{
			//Load Document
			Document document = new Document();
			using (MemoryStream stream = new MemoryStream(file))
			{
				document.LoadFromStream(stream, FileFormat.Auto);
			}

			document.SaveToFile(outputPath, FileFormat.PDF);
		}
	}
}
