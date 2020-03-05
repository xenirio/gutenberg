using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Xenirio.Component.Gutenberg
{
    public interface ReportDocumentContext 
    {
        ReportTemplate GetTemplate(string templateName);
    }
}
