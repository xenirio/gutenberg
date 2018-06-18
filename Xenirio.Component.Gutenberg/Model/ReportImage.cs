using System;
using System.IO;
using System.Linq;
using System.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Xenirio.Component.Gutenberg.Model
{
	public class ReportImage : ReportElement, IReportReplaceable
	{
		public byte[] Value
		{
			get { return GetValue<byte[]>(); }
			set { SetValue(value); }
		}

		public void Replace(Run element)
		{
            var paragraph = (Paragraph)element.Parent;
            if (paragraph != null)
            {
                var mainPart = getMainDocumentPart(paragraph);

                ImagePart imagePart;
                if (mainPart.GetType() == typeof(MainDocumentPart))
                    imagePart = ((MainDocumentPart)mainPart).AddImagePart(ImagePartType.Png);
                else if (mainPart.GetType() == typeof(HeaderPart))
                    imagePart = ((HeaderPart)mainPart).AddImagePart(ImagePartType.Png);
                else
                    imagePart = ((FooterPart)mainPart).AddImagePart(ImagePartType.Png);

                using (var stream = new MemoryStream(Value))
                {
                    imagePart.FeedData(stream);
                    var img = new Bitmap(stream);
                    var widthPx = img.Width;
                    var heightPx = img.Height;
                    var horzRezDpi = img.HorizontalResolution;
                    var vertRezDpi = img.VerticalResolution;
                    const int emusPerInch = 914400;
                    const int emusPerCm = 360000;
                    var maxWidthCm = 16.51;
                    var widthEmus = (long)(widthPx / horzRezDpi * emusPerInch);
                    var heightEmus = (long)(heightPx / vertRezDpi * emusPerInch);
                    var maxWidthEmus = (long)(maxWidthCm * emusPerCm);
                    if (widthEmus > maxWidthEmus)
                    {
                        var ratio = (heightEmus * 1.0m) / widthEmus;
                        widthEmus = maxWidthEmus;
                        heightEmus = (long)(widthEmus * ratio);
                    }
                    var run = runImage(mainPart.GetIdOfPart(imagePart), widthEmus, heightEmus);
                    paragraph.RemoveAllChildren<Run>();
                    paragraph.AppendChild(run);
                }
            }
        }

        private OpenXmlPart getMainDocumentPart(OpenXmlElement element)
        {
            return
            element?.Ancestors<Document>()?.FirstOrDefault()?.MainDocumentPart as OpenXmlPart ??
            element?.Ancestors<Header>()?.FirstOrDefault()?.HeaderPart as OpenXmlPart ??
            element?.Ancestors<Footer>()?.FirstOrDefault()?.FooterPart as OpenXmlPart;
        }

        private Run runImage(string relationshipId, long width, long height)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = width, Cy = height },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = Key
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = string.Format("{0}.png", Key.ToLower().Replace(".", "-"))
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri = string.Format("{{{0}}}", Guid.NewGuid())
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.None
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = width, Cy = height }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });
            return new Run(element);
        }
    }
}
