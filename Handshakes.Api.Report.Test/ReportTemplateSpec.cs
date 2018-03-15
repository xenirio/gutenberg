using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Handshakes.Api.Report.Test
{
	[TestClass]
	public class ReportTemplateSpec
	{
		private ReportTemplate template;

		[TestInitialize]
		public void Initialize()
		{
			template = new ReportTemplate(Environment.CurrentDirectory + @"\Resources\Sample.docx");
		}

		[TestMethod]
		public void Should_Open_Report_Template()
		{
			Assert.IsNotNull(template);
		}

		[TestMethod]
		public void Should_Get_Report_Template_Variables()
		{
			var variables = template.DocVariables;
			Assert.IsTrue(variables.Length > 0);
		}
	}
}
