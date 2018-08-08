using Xenirio.Component.Gutenberg.Model;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace Xenirio.Component.Gutenberg.Tests
{
    [TestClass]
    public class ReportDocumentSpec
    {
        [TestMethod]
        public void Should_Replace_Label_Element()
        {
            var sourcefile = Environment.CurrentDirectory + @"\Resources\Sample.docx";
            var outfile = Environment.CurrentDirectory + @"\Resources\SampleTest.docx";
            if (File.Exists(outfile))
                File.Delete(outfile);
            File.Copy(sourcefile, outfile);
            var document = new ReportDocument();

            document.InjectReportElement(new ReportLabel() { Key = "Header.Entity.Name", Value = "IPO of Ezra Holdings Limited" });
            document.InjectReportElement(new ReportLabel() { Key = "Footer.Creator", Value = "Vee" });
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
            document.Save(outfile);
        }

        [TestMethod]
        public void Should_Replace_Table_Element()
        {
            var sourcefile = Environment.CurrentDirectory + @"\Resources\SampleTable.docx";
            var outfile = Environment.CurrentDirectory + @"\Resources\SampleTableTest.docx";
            if (File.Exists(outfile))
                File.Delete(outfile);
            File.Copy(sourcefile, outfile);
            var document = new ReportDocument();

            document.InjectReportElement(new ReportTable()
            {
                Key = "Content.Table.Entity",
                Elements = new ReportLabel[][] {
                    new ReportLabel[] {
                        new ReportLabel() { Value = "Coopers" },
                        new ReportLabel() { Value = "Firm" },
                        new ReportLabel() { Value = "Chiang Mai" },
                        new ReportLabel() { Value = "John" },
                        new ReportLabel() { Value = "Good" },
                    },
                    new ReportLabel[] {
                        new ReportLabel() { Value = "John" },
                        new ReportLabel() { Value = "Person" },
                        new ReportLabel() { Value = "USA" },
                        new ReportLabel() { Value = "John" },
                        new ReportLabel() { Value = "Good" }
                    }
                }
            });
            document.Save(outfile);
        }

        [TestMethod]
        public void Should_Replace_Image_Element()
        {
            var sourcefile = Environment.CurrentDirectory + @"\Resources\SampleImage.docx";
            var outfile = Environment.CurrentDirectory + @"\Resources\SampleImageTest.docx";
            if (File.Exists(outfile))
                File.Delete(outfile);
            File.Copy(sourcefile, outfile);
            var document = new ReportDocument();

            document.InjectReportElement(new ReportImage()
            {
                Key = "Content.Image.Logo",
                Value = File.ReadAllBytes(Environment.CurrentDirectory + @"\Resources\logo.png")
            });
            document.InjectReportElement(new ReportImage()
            {
                Key = "Content.Image.Banner",
                Value = File.ReadAllBytes(Environment.CurrentDirectory + @"\Resources\banner.png")
            });
            document.Save(outfile);
        }

        [TestMethod]
        public void Should_Replace_Image_Table()
        {
            var sourcefile = Environment.CurrentDirectory + @"\Resources\SampleImageTable.docx";
            var outfile = Environment.CurrentDirectory + @"\Resources\SampleImageTableTest.docx";
            if (File.Exists(outfile))
                File.Delete(outfile);
            File.Copy(sourcefile, outfile);
            var document = new ReportDocument();

            var byteLogo = File.ReadAllBytes(Environment.CurrentDirectory + @"\Resources\logo.png");
            document.InjectReportElement(new ReportTable()
            {
                Key = "Content.Image.Logo",
                Elements = new ReportImage[][]{
                    new ReportImage[] {
                        new ReportImage(){ Value = byteLogo }
                    },
                    new ReportImage[] {
                        new ReportImage(){ Value = byteLogo }
                    },
                    new ReportImage[] {
                        new ReportImage(){ Value = byteLogo }
                    }
                }
            });
            document.InjectReportElement(new ReportImage()
            {
                Key = "Content.Image.Banner",
                Value = File.ReadAllBytes(Environment.CurrentDirectory + @"\Resources\banner.png")
            });
            document.Save(outfile);
        }

        [TestMethod]
        public void Should_Replace_Composite_Table()
        {
            var sourcefile = Environment.CurrentDirectory + @"\Resources\SampleCompositeTable.docx";
            var outfile = Environment.CurrentDirectory + @"\Resources\SampleCompositeTableTest.docx";
            if (File.Exists(outfile))
                File.Delete(outfile);
            File.Copy(sourcefile, outfile);
            var document = new ReportDocument();

            var byteLogo = File.ReadAllBytes(Environment.CurrentDirectory + @"\Resources\logo.png");
            document.InjectReportElement(new ReportTable()
            {
                Key = "Content.Image.Logo",
                Elements = new ReportElement[][]{
                    new ReportElement[] {
                        new ReportImage(){ Value = byteLogo },
                        new ReportLabel(){ Value = "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum." },
                        new ReportLabel(){ Value = "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum." }
                    },
                    new ReportElement[] {
                        new ReportImage(){ Value = byteLogo },
                        new ReportLabel(){ Value = "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum." },
                        new ReportLabel(){ Value = "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum." }
                    },
                    new ReportElement[] {
                        new ReportImage(){ Value = byteLogo },
                        new ReportLabel(){ Value = "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum." },
                        new ReportLabel(){ Value = "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum." }
                    },
                }
            });
            document.InjectReportElement(new ReportImage()
            {
                Key = "Content.Image.Banner",
                Value = File.ReadAllBytes(Environment.CurrentDirectory + @"\Resources\banner.png")
            });
            document.Save(outfile);
        }

        [TestMethod]
        public void Should_Replace_Complex_Table()
        {
            var sourcefile = Environment.CurrentDirectory + @"\Resources\SampleTableAdvance.docx";
            var outfile = Environment.CurrentDirectory + @"\Resources\SampleTableAdvanceTest.docx";
            if (File.Exists(outfile))
                File.Delete(outfile);
            File.Copy(sourcefile, outfile);

            var byteBanner = File.ReadAllBytes(Environment.CurrentDirectory + @"\Resources\banner.png");
            var document = new ReportDocument();
            document.InjectReportElement(new ReportTableComplex()
            {
                Key = "Content.Table.Rule",
                Elements = new ReportElement[][] {
                    new ReportLabel[] {
                        new ReportLabel(){ Key = "Template.Rule.Title", Value = "Rule 1" }
                    },
                    new ReportLabel[] {
                        new ReportLabel(){ Key = "Template.Rule.Entity.Name", Value = "ABC" }
                    },
                    new ReportImage[] {
                        new ReportImage(){ Key = "Template.Rule.Entity.Map", Value = byteBanner }
                    },
                    new ReportImage[] {
                        new ReportImage(){ Key = "Template.Rule.Entity.Map", Value = byteBanner }
                    },
                    new ReportImage[] {
                        new ReportImage(){ Key = "Template.Rule.Entity.Map", Value = byteBanner }
                    },
                    new ReportLabel[] {
                        new ReportLabel(){ Key = "Template.Rule.Entity.Name", Value = "DEF" }
                    },
                    new ReportLabel[] {
                        new ReportLabel(){ Key = "Template.Rule.Title", Value = "Rule 2" }
                    },
                    new ReportLabel[] {
                        new ReportLabel(){ Key = "Template.Rule.Entity.Name", Value = "GHI" }
                    },
                }
            });
            document.Save(outfile);
        }

        [TestMethod]
        public void Should_Replace_Apply_Table()
        {
            var sourcefile = Environment.CurrentDirectory + @"\Resources\SampleApplyTable.docx";
            var outfile = Environment.CurrentDirectory + @"\Resources\SampleApplyTableTest.docx";
            if (File.Exists(outfile))
                File.Delete(outfile);
            File.Copy(sourcefile, outfile);

            var document = new ReportDocument();
            document.InjectReportElement(new ReportTable()
            {
                Key = "Content.Rule.Header",
                Elements = new ReportLabel[][] {
                    new ReportLabel[] {
                        new ReportLabel() { Value = "Application ID" },
                        new ReportLabel() { Value = "Company UEN" }
                    }
                }
            });
            document.InjectReportElement(new ReportTable()
            {
                Key = "Content.Rule.Title2",
                Elements = new ReportLabel[][] {
                    new ReportLabel[] {
                        new ReportLabel() { Value = "The following are applicants/claimants with the same phone number" }
                    }
                }
            });
            document.InjectReportElement(new ReportTable()
            {
                Key = "Content.Rule.Table2",
                Elements = new ReportLabel[][] {
                    new ReportLabel[] {
                        new ReportLabel() { Value = "000" },
                        new ReportLabel() { Value = "AAA" }
                    },
                    new ReportLabel[] {
                        new ReportLabel() { Value = "001" },
                        new ReportLabel() { Value = "AAB" }
                    },
                    new ReportLabel[] {
                        new ReportLabel() { Value = "002" },
                        new ReportLabel() { Value = "AAC" }
                    }
                }
            });
            document.Save(outfile);
        }
    }
}
