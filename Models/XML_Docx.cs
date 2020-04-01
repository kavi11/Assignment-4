﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection.Metadata;
using System.Text;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXML.Models
{
    class XML_Docx
    {
        public static void CreateWordprocessingDocument(string filepath, byte[] myimage, List<Student> students)
        {
            // Create a document by supplying the filepath. 
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                foreach (var student in students)
                {
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text("My Name is "+ student.FirstName + " " + student.LastName));

                    /*Display My Picture on my Page*/
                    if (student.IsMe == true)
                    {
                        FTP.DownloadFile(Constants.Location.FTP.ImageUrl, Constants.Location.DataFolderImage);
                        using (FileStream stream = new FileStream(Constants.Location.DataFolderImage, FileMode.Open))
                        {
                            imagePart.FeedData(stream);
                        }
                        AddImageToBody(wordDocument, mainPart.GetIdOfPart(imagePart));
                    }
                    Paragraph PageBreakParagraph = new Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Break() { Type = BreakValues.Page }));
                    body.Append(PageBreakParagraph);
                }
            }
        }
        /// <summary>
        /// Take the image from above and defines it in xml format for insertion into .docx.
        /// </summary>  
        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
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
                             Name = "My Picture"
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
                                             Name = "myimage.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                       new PIC.BlipFill(
                                 new A.Blip(
                                     new A.BlipExtensionList(
                                         new A.BlipExtension()
                                         {
                                             Uri =
                                               "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                         })
                                 )
                                 {
                                     Embed = relationshipId,
                                     CompressionState =
                                     A.BlipCompressionValues.Print
                                 },
                                 new A.Stretch(
                                     new A.FillRectangle())),
                             new PIC.ShapeProperties(
                                 new A.Transform2D(
                                     new A.Offset() { X = 0L, Y = 0L },
                                     new A.Extents() { Cx = 990000L, Cy = 792000L }),
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

            // Append the reference to the body. The element should be in 
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }
    }
}
