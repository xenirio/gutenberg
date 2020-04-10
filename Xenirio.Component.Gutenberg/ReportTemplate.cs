using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xenirio.Component.Gutenberg.Model;

namespace Xenirio.Component.Gutenberg
{
    public class ReportTemplate
    {
        private Body rootElement;
        private ReportDocumentContext context;

        public string Name { get; set; }

        public ReportTemplate(ReportDocumentContext context, string templateName)
        {
            Name = templateName;
            this.context = context;
        }

        public void Apply(OpenXmlElement replacedElement, IReportReplaceable[][] templateValues)
        {
            var templateValuesDicts = new List<Dictionary<string, IReportReplaceable>>();
            var mainPart = getMainDocumentPart(replacedElement);
            foreach (var row in templateValues)
            {
                var rowDict = new Dictionary<string, IReportReplaceable>();
                foreach(var item in row)
                {
                    rowDict.Add((item as ReportElement).Key, item);
                    if (item.GetType() == typeof(ReportImage))
                    {
                        (item as ReportImage).InjectDocPart(mainPart);
                    }
                }
                templateValuesDicts.Add(rowDict);
            }
 
            foreach (var rowValue in templateValuesDicts) {
                var paragraphs = new List<Paragraph>();
                var dRootElement = rootElement.Clone() as Body;
                paragraphs.AddRange(dRootElement.Descendants<Paragraph>());

                var dirtyChilds = new List<OpenXmlElement>();
                foreach (var paragraph in paragraphs.Where(p => p.Descendants<FieldCode>().Any()))
                {
                    List<string> codes = null;
                    foreach (var elem in paragraph.Elements().SelectMany(e => e.ChildElements))
                    {
                        if (elem.GetType() == typeof(FieldChar))
                        {
                            var fChar = (FieldChar)elem;
                            if (fChar.FieldCharType == FieldCharValues.Begin)
                                codes = new List<string>();
                            if (fChar.FieldCharType == FieldCharValues.End)
                            {
                                var code = string.Join("", codes).Trim();
                                var fieldCode = new FieldCode(code);
                                var target = (Run)elem.Parent.PreviousSibling<Run>().Clone();
                                target.ReplaceChild(fieldCode, target.GetFirstChild<FieldCode>());
                                paragraph.InsertBefore(target, elem.Parent);
                                if (!code.StartsWith("DOCVARIABLE"))
                                    dirtyChilds.RemoveRange(dirtyChilds.Count - codes.Count, codes.Count);
                                codes = null;
                            }
                        }
                        else
                        {
                            if (codes != null && elem.GetType() == typeof(FieldCode))
                            {
                                codes.Add(elem.InnerText);
                                dirtyChilds.Add(elem.Parent);
                            }
                        }
                    }
                    foreach (var dirty in dirtyChilds)
                    {
                        paragraph.RemoveChild(dirty);
                    }
                    dirtyChilds.Clear();
                }

                //foreach (var paragraph in paragraphs.Where(p => p.Descendants<FieldCode>().Any()))
                //{
                //    var code = paragraph.InnerText.Trim();
                //    paragraph.RemoveAllChildren<Run>();
                //    paragraph.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }));
                //    paragraph.Append(new Run(new FieldCode(code)));
                //    paragraph.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.End }));
                //}

                var fields = paragraphs.SelectMany(p => p.Descendants<FieldCode>()).Where(p => p.InnerText.Contains("DOCVARIABLE"));
                replace(fields.ToArray(), rowValue);
                clean(fields.ToArray(), rowValue);

                foreach (var child in dRootElement.ChildElements)
                {
                    replacedElement.InsertBeforeSelf(child.CloneNode(true));
                }
            }
            replacedElement.Remove();
        }

        private OpenXmlPart getMainDocumentPart(OpenXmlElement element)
        {
            return
            element?.Ancestors<Document>()?.FirstOrDefault()?.MainDocumentPart as OpenXmlPart ??
            element?.Ancestors<Header>()?.FirstOrDefault()?.HeaderPart as OpenXmlPart ??
            element?.Ancestors<Footer>()?.FirstOrDefault()?.FooterPart as OpenXmlPart;
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
                    if (variable is ReportTemplateElement)
                    {
                        var templateElement = variable as ReportTemplateElement;
                        var template = context.GetTemplate(templateElement.TemplateKey);
                        if (template != null)
                        {
                            template.Apply(field.Ancestors<Paragraph>().Single(), templateElement.Value);
                        }
                        else
                            throw new KeyNotFoundException($"Not found template key {templateElement.TemplateKey}");
                    }
                    else 
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
                    if (elem.Parent == null)
                        continue;
                    if (elem.Parent.GetType() == typeof(TableCell))
                    {
                        try
                        {
                            elem.Parent.Parent.Remove();
                        }
                        catch
                        {
                            // Ignore in case cannot remove
                        }
                    }
                    else
                    {
                        (new ReportLabel() { Key = key, Value = "" }).Replace((Run)field.Parent);
                    }
                }
            }
        }

        public void Compile(WordprocessingDocument document)
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
