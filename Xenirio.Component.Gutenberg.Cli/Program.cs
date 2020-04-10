using CommandLine;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xenirio.Component.Gutenberg.Extensions;

namespace Xenirio.Component.Gutenberg.Cli
{
    class Program
    {
        [Verb("create", HelpText = "Create file")]
        public class CreateOptions
        {
            [Option("file", Required = false, HelpText = "Output file name.")]
            public string Output { get; set; }

            [Option("from", Required = true, HelpText = "Template file.")]
            public string Template { get; set; }

            [Option("with", Required = true, HelpText = "Source of data file.")]
            public string Source { get; set; }
        }

        static void Main(string[] args)
        {
            var result = CommandLine.Parser.Default.ParseArguments<CreateOptions>(args)
                         .MapResult(
                            (CreateOptions opts) =>
                            {
                                var template = opts.Template;
                                var output = opts.Output;
                                if (string.IsNullOrEmpty(output))
                                    output = string.Format("{0}\\{1}.docx", Environment.CurrentDirectory, Path.GetFileNameWithoutExtension(template));
                                var report = new ReportGenerator(string.Format("{0}\\{1}", Environment.CurrentDirectory, template));
                                var json = JObject.Parse(File.ReadAllText(string.Format("{0}\\{1}", Environment.CurrentDirectory, opts.Source)));
                                report.setJsonObject(json);
                                report.GenerateToFile(output);
                                return 0;
                            },
                            errs => 1
                         );
        }
    }
}
