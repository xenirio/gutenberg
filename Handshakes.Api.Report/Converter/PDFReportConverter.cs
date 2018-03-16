using Spire.Doc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Handshakes.Api.Report.Converter
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

		public static void ConvertToFile(string filePath)
		{
			var outFile = string.Format(@"{0}\{1}.pdf", Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath));

			//Load Document
			Document document = new Document();
			document.LoadFromFile(filePath);

			document.SaveToFile(outFile, FileFormat.PDF);
		}
	}
}
