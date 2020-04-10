using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xenirio.Component.Gutenberg.Model;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System;
using DocumentFormat.OpenXml;

namespace Xenirio.Component.Gutenberg
{
    public class ReportDocument : ReportDocumentContext
    {
        private Dictionary<string, IReportReplaceable> variables = new Dictionary<string, IReportReplaceable>();
        private Dictionary<string, ReportTemplate> templates = new Dictionary<string, ReportTemplate>();

        public ReportTemplate GetTemplate(string templateName)
        {
            if (templates.ContainsKey(templateName))
                return templates[templateName];
            else
                return null;
        }

        public void InjectReportElement(ReportElement element)
        {
            variables.Add(string.Format(@"DOCVARIABLE  {0}", element.Key), (IReportReplaceable)element);
        }

        public void RegisterTemplate(string templateName)
        {
            templates.Add(templateName, new ReportTemplate(this, templateName));
        }

        public void Save(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("File not found", filePath);
            var document = WordprocessingDocument.Open(filePath, true);
            CompileTemplate(document);
            ReplaceDocument(document);
            document.Close();
        }

        public byte[] Save(byte[] bytes)
        {
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(bytes, 0, bytes.Length);
                using (WordprocessingDocument document = WordprocessingDocument.Open(mem, true))
                {
                    CompileTemplate(document);
                    ReplaceDocument(document);
                    document.Close();
                    return mem.ToArray();
                }
            }
        }

        private void CompileTemplate(WordprocessingDocument document)
        {
            foreach (var template in templates.Values)
            {
                template.Compile(document);
            }
        }

        private void ReplaceDocument(WordprocessingDocument document)
        {
            var paragraphs = new List<Paragraph>();

            // header variable
            paragraphs.AddRange(document.MainDocumentPart.HeaderParts.SelectMany(h => h.Header.Descendants<Paragraph>()));
            // body variable
            paragraphs.AddRange(document.MainDocumentPart.RootElement.Descendants<Paragraph>());
            // footer variable
            paragraphs.AddRange(document.MainDocumentPart.FooterParts.SelectMany(f => f.Footer.Descendants<Paragraph>()));

            //foreach (var paragraph in paragraphs.Where(p => p.Descendants<FieldCode>().Any()))
            //{
            //    var codes = paragraph.SelectMany(p => p.Descendants<FieldCode>()).ToList();
            //    paragraph.RemoveAllChildren<Run>();
            //    foreach (var code in codes)
            //    {
            //        paragraph.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }));
            //        paragraph.Append(new Run(new FieldCode(code.InnerText.Trim())));
            //        paragraph.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.End }));
            //    }
            //}

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
                            var fieldCode = new Run(new FieldCode(code));
                            paragraph.InsertBefore(fieldCode, elem.Parent);
                            if(!code.StartsWith("DOCVARIABLE"))
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

            var fields = paragraphs.SelectMany(p => p.Descendants<FieldCode>()).Where(p => p.InnerText.Contains("DOCVARIABLE"));
            Replace(fields.ToArray());
            Clean(fields.ToArray());
        }

        private void Replace(FieldCode[] fields)
        {
            var codes = fields.Select(f => f.Text.Trim()).Distinct().ToArray();
            var intersectFields = variables.Select(v => v.Key).Intersect(codes);
            if (variables.Count > intersectFields.Count())
            {
                var missingFields = variables.Keys.Except(intersectFields);
                throw new KeyNotFoundException($"Variables not match with Fields. Missing Fields = {string.Join(",", missingFields)}");
            }
            foreach (var field in fields)
            {
                var key = field.Text.Trim();
                IReportReplaceable variable;
                if (variables.ContainsKey(key))
                {
                    variable = variables[key];
                    if (variable is ReportTemplateElement)
                    {
                        var element = variable as ReportTemplateElement;
                        if (templates.ContainsKey(element.TemplateKey))
                        {
                            templates[element.TemplateKey].Apply(field.Ancestors<Paragraph>().Single(), element.Value);
                        }
                        else
                            throw new KeyNotFoundException($"Not found template key {element.TemplateKey}");
                    }
                    else
                        variable.Replace((Run)field.Parent);
                }
            }
        }

        private void Clean(FieldCode[] fields)
        {
            foreach (var field in fields)
            {
                var key = field.Text.Trim();
                if (!key.Contains("DOCVARIABLE"))
                    continue;
                if (!variables.ContainsKey(key))
                {
                    var elem = field.Ancestors<Paragraph>().Single();
                    if (elem.Parent.GetType() == typeof(TableCell))
                    {
                        try
                        {
                            elem.Parent.Parent.Remove();
                        } catch
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
    }
}
