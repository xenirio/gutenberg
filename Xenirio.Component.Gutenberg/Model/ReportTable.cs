using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
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
            var body = element.Parent;
            while (body.GetType() != typeof(Body))
                body = body.Parent;
            var parent = element.Parent.Parent;
            var table = element.Parent.Parent.Parent.Parent;
            table.Descendants<TableRow>().First().Remove();
            for (var i = 0; i < Elements.Length; i++)
            {
                var tr = new TableRow();
                table.AppendChild(tr);
                for (var j = 0; j < Elements[i].Length; j++)
                {
                    var elem = Elements[i][j];
                    var template = (TableCell)body.Descendants<FieldCode>()
                        .Where(f => f.Text.Trim() == string.Format(@"DOCVARIABLE  {0}", elem.Key)).Single().Parent.Parent.Parent.Clone();
                    tr.AppendChild(template);
                    var run = (Run)template.Descendants<FieldCode>().Single().Parent;
                    ((IReportReplaceable)elem).Replace(run);
                }
            }
        }
    }
}
