using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Xenirio.Component.Gutenberg.Model
{
	internal class ReportTable : ReportElement, IReportReplaceable
	{
		public ReportElement[][] Elements
		{
			get { return GetValue<ReportElement[][]>(); }
			set { SetValue(value); }
		}

		public virtual void Replace(Run element)
		{
            var tr = (TableRow)element.Parent.Parent.Parent;
			tr.Descendants<Run>().ToList().ForEach(c => c.Remove());
			var table = (Table)tr.Parent;
			for (var i = 0; i < Elements.Length; i++)
			{
				var tmpTr = (TableRow)tr.Clone();
				table.AppendChild(tmpTr);
				for (var j = 0; j < Elements[i].Length; j++)
				{
                    var elem = Elements[i][j];
                    elem.Key = Key;
                    Run run = tmpTr.Descendants<TableCell>().ElementAt(j).GetFirstChild<Paragraph>().AppendChild<Run>(new Run());
                    ((IReportReplaceable)elem).Replace(run);
				}
			}
			tr.Remove();
		}
	}

    internal class ReportTableComplex : ReportTable
    {
        public override void Replace(Run element)
        {
            var table = (Table)element.Parent.Parent.Parent.Parent;
            for (var i = 0; i < Elements.Length; i++)
            {
                var tds = new List<TableCell>();
                for (var j = 0; j < Elements[i].Length; j++)
                {
                    var td = table.Descendants<TableRow>().ElementAt(i).Descendants<TableCell>().ElementAt(j);
                    var elem = Elements[i][j];
                    
                    if (elem.GetType() == typeof(ReportTableComplex))
                    {
                        var tmpTd = (TableCell)td.CloneNode(true);
                        tds.Add(tmpTd);
                        td = tmpTd;
                    }
                    Run run = td.GetFirstChild<Paragraph>().AppendChild<Run>(new Run());
                    ((IReportReplaceable)elem).Replace(run);
                }
                var tr = new TableRow(tds);
                table.AppendChild(tr);
            }
        }
    }
}
