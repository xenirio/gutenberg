using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xenirio.Component.Gutenberg.Model;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Xenirio.Component.Gutenberg
{
    public class ReportDocument
    {
        private Dictionary<string, IReportReplaceable> variables = new Dictionary<string, IReportReplaceable>();

        public void InjectReportElement(ReportElement element)
        {
            variables.Add(string.Format(@"DOCVARIABLE  {0}", element.Key), (IReportReplaceable)element);
        }

        public void Save(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("File not found", filePath);
            var document = WordprocessingDocument.Open(filePath, true);
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
                    ReplaceDocument(document);
                    document.Close();
                    return mem.ToArray();
                }
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

            foreach (var paragraph in paragraphs.Where(p => p.Descendants<FieldCode>().Any()))
            {
                var code = paragraph.InnerText.Trim();
                paragraph.RemoveAllChildren<Run>();
                paragraph.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }));
                paragraph.Append(new Run(new FieldCode(code)));
                paragraph.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.End }));
            }

            var fields = paragraphs.SelectMany(p => p.Descendants<FieldCode>());
            Replace(fields.ToArray());
            Clean(fields.ToArray());
        }

        private void Replace(FieldCode[] fields)
        {
            var codes = fields.Select(f => f.Text.Trim()).Distinct().ToArray();
            var intersectFields = variables.Select(v => v.Key).Intersect(codes);
            if (variables.Count > intersectFields.Count())
                throw new KeyNotFoundException("Variables not match with Fields");
            foreach (var field in fields)
            {
                var key = field.Text.Trim();
                IReportReplaceable variable;
                if (variables.ContainsKey(key))
                {
                    variable = variables[key];
                    variable.Replace((Run)field.Parent);
                }
            }
        }

        private void Clean(FieldCode[] fields)
        {
            foreach (var field in fields)
            {
                var key = field.Text.Trim();
                if (!variables.ContainsKey(key))
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
    }
}
