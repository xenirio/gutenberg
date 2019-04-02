using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xenirio.Component.Gutenberg.Model;

namespace Xenirio.Component.Gutenberg
{
    internal class ReportTemplate
    {
        private Body rootElement;

        internal string Name { get; set; }

        internal ReportTemplate(string templateName)
        {
            Name = templateName;
        }

        internal void Apply(OpenXmlElement replacedElement, Dictionary<string, IReportReplaceable>[] templateValues)
        {
            var paragraphs = new List<Paragraph>();
            foreach (var rowValue in templateValues) {
                var dRootElement = rootElement.Clone() as Body;
                paragraphs.AddRange(dRootElement.Descendants<Paragraph>());

                foreach (var paragraph in paragraphs.Where(p => p.Descendants<FieldCode>().Any()))
                {
                    var code = paragraph.InnerText.Trim();
                    paragraph.RemoveAllChildren<Run>();
                    paragraph.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }));
                    paragraph.Append(new Run(new FieldCode(code)));
                    paragraph.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.End }));
                }

                var fields = paragraphs.SelectMany(p => p.Descendants<FieldCode>());
                replace(fields.ToArray(), rowValue);
                clean(fields.ToArray(), rowValue);

                foreach (var child in dRootElement.ChildElements)
                {
                    replacedElement.InsertBeforeSelf(child.CloneNode(true));
                }
            }
            replacedElement.Remove();
        }

        private void replace(FieldCode[] fields, Dictionary<string, IReportReplaceable> rowValue)
        {
            var codes = fields.Select(f => f.Text.Replace("DOCVARIABLE", "").Trim()).Distinct().ToArray();
            var intersectFields = rowValue.Select(v => v.Key).Intersect(codes);
            if (rowValue.Count > intersectFields.Count())
                throw new KeyNotFoundException("Variables not match with Fields");
            foreach (var field in fields)
            {
                var key = field.Text.Replace("DOCVARIABLE", "").Trim();
                IReportReplaceable variable;
                if (rowValue.ContainsKey(key))
                {
                    variable = rowValue[key];
                    variable.Replace((Run)field.Parent);
                }
            }
        }

        private void clean(FieldCode[] fields, Dictionary<string, IReportReplaceable> rowValue)
        {
            foreach (var field in fields)
            {
                var key = field.Text.Trim();
                if (!rowValue.ContainsKey(key))
                {
                    var elem = field.Ancestors<Paragraph>().Single();
                    if (elem.Parent.GetType() == typeof(TableCell))
                    {
                        elem.Parent.Parent.Remove();
                    }
                    else
                    {
                        (new ReportLabel() { Key = key, Value = "" }).Replace((Run)field.Parent);
                    }
                }
            }
        }

        internal void Compile(WordprocessingDocument document)
        {
            var fields = document.MainDocumentPart.RootElement.Descendants<FieldCode>();
            var startSectionName = string.Format("DOCVARIABLETemplateSection.{0}.Start", Name);
            var endSectionName = string.Format("DOCVARIABLETemplateSection.{0}.End", Name);
            var removingNode = new List<OpenXmlElement>();
            Paragraph start = null;
            Paragraph end = null;
            foreach (var f in fields)
            {
                var innerText = f.InnerText.Replace(" ", "");
                if (innerText == startSectionName)
                {
                    var cur = f as OpenXmlElement;
                    while (cur.GetType() != typeof(Paragraph))
                        cur = cur.Parent;
                    start = cur as Paragraph;
                }
                else if (innerText == endSectionName)
                {
                    var cur = f as OpenXmlElement;
                    while (cur.GetType() != typeof(Paragraph))
                        cur = cur.Parent;
                    end = cur as Paragraph;
                }
            }

            if (start != null && end != null)
            {
                removingNode.Add(start);
                removingNode.Add(end);

                var templateBody = new Body();
                var cur = start.NextSibling();
                while (cur != null && cur != end)
                {
                    templateBody.Append(cur.CloneNode(true));
                    removingNode.Add(cur);
                    cur = cur.NextSibling();
                }
                rootElement = templateBody;

                // Remove Successfully Compiled Template From Document
                removingNode.ForEach(n =>
                {
                    n.Remove();
                });
            } else
                throw new Exception($"Cannot create template name {Name}");
        }
    }
}
