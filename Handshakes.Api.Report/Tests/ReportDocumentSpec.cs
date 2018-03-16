using Handshakes.Api.Report.Model;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Handshakes.Api.Report.Tests
{
	[TestClass]
	public class ReportDocumentSpec
	{
		[TestMethod]
		public void Should_Save_Document()
		{
			var sourcefile = Environment.CurrentDirectory + @"\Resources\Sample.docx";
			var outfile = Environment.CurrentDirectory + @"\Resources\SampleTest.docx";
			if (File.Exists(outfile))
				File.Delete(outfile);
			File.Copy(sourcefile, outfile);
			var document = new ReportDocument(outfile);

			document.InjectReportElement(new ReportLabel() { Key = "Header.Entity.Name", Value = "IPO of Ezra Holdings Limited" });
			document.InjectReportElement(new ReportLabel() { Key = "Content.Entity.Name", Value = "IPO of Ezra Holdings Limited" });
			document.InjectReportElement(new ReportLabels() { Key = "Content.Entity.Names", Values = new string[] { "IPO of Ezra Holdings Limited", "IPO of Ezra" } });
			document.InjectReportElement(new ReportLabel() { Key = "Content.Entity.Remark", Value = @"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed sollicitudin blandit massa, sit amet ornare odio aliquet a. Aliquam ac porttitor lacus. Duis molestie felis convallis, volutpat orci et, ornare tellus. Proin sed risus sit amet nunc faucibus pharetra. Curabitur imperdiet, lectus vitae accumsan semper, justo magna hendrerit lacus, eget iaculis felis quam non nunc. Vestibulum condimentum congue lectus, et consequat augue venenatis sit amet. Vivamus id efficitur nunc, sed suscipit lacus. Vivamus eleifend vestibulum tortor id ullamcorper. Phasellus nec fringilla mauris, eu aliquam nulla. In hac habitasse platea dictumst. Nullam venenatis tristique hendrerit. Maecenas a tellus euismod sapien elementum lobortis. Donec a felis imperdiet, cursus lorem at, porttitor odio.

Morbi eget lobortis sapien.Integer non fermentum ante.Mauris non fermentum ipsum,
				porttitor mollis erat.Phasellus massa nibh,
				vehicula vestibulum pharetra id,
				posuere quis sapien.In ullamcorper velit diam,
				et luctus nulla auctor at.Morbi tempus sollicitudin ex,
				lobortis efficitur quam gravida ornare.Donec et vestibulum sapien.Donec fermentum lectus enim,
				nec consequat metus varius a.Vestibulum dignissim vestibulum eros,
				vel euismod nisi condimentum non.

Ut maximus dui arcu,
				volutpat bibendum lorem convallis vel.Morbi iaculis at augue ornare efficitur.Nulla eu convallis odio.Duis quis leo euismod,
				varius ex blandit,
				luctus quam.Aenean pharetra magna non placerat sodales.Nulla sed purus et ipsum molestie faucibus at ac nunc.Pellentesque imperdiet euismod mauris nec varius.Fusce mollis condimentum mi,
				eget sodales lectus tempor quis.Pellentesque eleifend cursus nibh,
				et sodales lacus tempus et.Nam condimentum varius porta.

Suspendisse potenti.Morbi in varius nisi.Nunc nec diam metus.In ullamcorper sit amet elit in malesuada.Mauris aliquam nibh ornare quam iaculis euismod.Maecenas porta mollis diam,
				quis rhoncus ante maximus vel.Mauris nec tincidunt elit.Sed consequat commodo dignissim.Nunc eu egestas enim.Morbi et magna ut nisl imperdiet consectetur eu eu diam.Praesent vel ante consectetur,
				sodales erat ac,
				finibus risus.Praesent quis suscipit odio.

Cras vel suscipit ex.Fusce quis egestas ex.Nunc mattis arcu sit amet felis ultricies condimentum.Sed eget placerat nulla,
				vitae ultricies ex.Duis bibendum luctus mauris lacinia dapibus.Vivamus consectetur lacus laoreet consequat ultrices.Donec ligula eros,
				pretium ut magna id,
				pharetra aliquet velit.Duis auctor semper nisi." });
			document.InjectReportElement(new ReportLabel() { Key = "Content.Entity.Remark.Chinese", Value = @"Lorem存有悲坐阿梅德，consectetur adipiscing ELIT。然而，关心群众坐，很多足球恨香蕉。信用评级机构的sed arcu交流porttitor拉克丝。作业电视台的足球成就，而临床检查，以适应区域。但是，现在很多笑声微波下巴颤抖的。 Curabitur imperdiet，lectus简历accumsan森佩尔胡斯托麦格纳hendrerit ullamcorper英里，ID iaculis猫，他们现在要做的不是。规划毕业酱制造和摄影宣传消毒胡萝卜。这使我们现在可是好东西生活。 Eleifend Vivamus前庭tortor ID ullamcorper。 Phasellus NEC燕雀mauris，欧盟aliquam法无。再融资。局伤心继电器消毒。性能SAPIEN SAPIEN无线网元的政策。直到在，创新的硬件足球融资运行LOREM。

足球SAPIEN需要纸箱。整数暖锋。发酵存有mauris非，porttitor油树了。 Phasellus accumsan DUI NIBH，vehicula前庭pharetra ID，posuere QUIS SAPIEN。笔记本电脑的视频电缆，如笔者在零。 Morbi tempor sollicitudin出来，它变得比孕妇ornare lobortis。拱门直到SAPIEN。 Donec发酵lectus enim，从eget欧盟turpis NEC consequat。制造赌波观光，或直到酱性能。" });
			document.Save();
		}

		[TestMethod]
		public void Should_Save_Table_Document()
		{
			var sourcefile = Environment.CurrentDirectory + @"\Resources\SampleTable.docx";
			var outfile = Environment.CurrentDirectory + @"\Resources\SampleTableTest.docx";
			if (File.Exists(outfile))
				File.Delete(outfile);
			File.Copy(sourcefile, outfile);
			var document = new ReportDocument(outfile);

			document.InjectReportElement(new ReportTable()
			{
				Key = "Content.Table.Entity",
				Values = new string[][] {
					new string[] { "Coopers", "Firm", "Chiang Mai", "John", "Good" },
					new string[] { "John", "Person", "USA", "John", "Good" }
				}
			});
			document.Save();
		}

		[TestMethod]
		public void Should_Save_EntityInfo()
		{
			var sourcefile = Environment.CurrentDirectory + @"\Resources\Templates\TemplateEntityInfo.docx";
			var outfile = Environment.CurrentDirectory + @"\Resources\EntityInfo.docx";
			if (File.Exists(outfile))
				File.Delete(outfile);
			File.Copy(sourcefile, outfile);
			var document = new ReportDocument(outfile);
		}
	}
}
