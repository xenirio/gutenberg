using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Handshakes.Api.Report.Model
{
	internal class ReportTable : ReportElement, IReportReplaceable
	{
		public string[][] Values
		{
			get { return GetValue<string[][]>(); }
			set { SetValue(value); }
		}

		public void Replace(Run element)
		{
			var tr = (TableRow)element.Parent.Parent.Parent;
			tr.Descendants<Run>().ToList().ForEach(c => c.Remove());
			var table = (Table)tr.Parent;
			for (var i = 0; i < Values.Length; i++)
			{
				var tmpTr = (TableRow)tr.Clone();
				table.AppendChild(tmpTr);
				for (var j = 0; j < Values[i].Length; j++)
				{
					var value = Values[i][j];
					tmpTr.ChildElements.Skip(1).ElementAt(j).GetFirstChild<Paragraph>().AppendChild(new Run(new Text(value)));
				}
			}
			tr.Remove();
		}
	}
}
