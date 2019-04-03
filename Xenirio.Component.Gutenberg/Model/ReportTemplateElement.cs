using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Xenirio.Component.Gutenberg.Model
{
    public class ReportTemplateElement : ReportElement, IReportReplaceable
    {
        public string TemplateKey { get; set; }
        public IReportReplaceable[][] Value { get; set; } 

        public void Replace(Run element)
        {
            throw new NotImplementedException();
        }
    }
}
