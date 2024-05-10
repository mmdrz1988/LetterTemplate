// See https://aka.ms/new-console-template for more information

using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System.Reflection.Metadata;
using System.Text.RegularExpressions;

namespace LetterTemplate
{
	class Program
	{
        static void Main(string[] args)
        {

            string ResultPath = @"E:\result.docx";
            string loadPath = @"E:\result.docx";
            string ExistText = "SautinSoft.Document 2024.4.24 trial. Get your free 30-day key instantly: https://sautinsoft.com/start-for-free/";

            string SavePdfPath = @"E:\replaced.pdf";
            string SavePdfPath2 = @"E:\";
            string picPath = @"E:\12.png";
            //string SavePdfPath = @"E:\Replaced_signature.pdf";

            //DocumentCore dc = new DocumentCore();
            //DocumentBuilder db = new DocumentBuilder(dc);
            //db.StartBookmark("BookMark1");
            //db.Write("This is <BookMark> 1 and this is signature <signature>");
            //db.EndBookmark("BookMark1");
            //db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);
            //dc.Save(ResultPath);


            //FindTextAndReplaceAnyText(loadPath, "BookMark", "Target Text", SavePdfPath);
            //FindTextAndReplaceImage(loadPath, picPath, SavePdfPath2);
            HeadersAndFooters(null);
            //FindTextAndReplaceAnyText(loadPath, ExistText, "s", SavePdfPath, true);
        }
        public static void FindTextAndReplaceAnyText(string loadPath, string ExistText, string TargetText, string SavePdfPath, bool IsCopyRightContent=false)
        {
            DocumentCore dc = DocumentCore.Load(loadPath);
            Regex regex = new Regex(@$"{ExistText}", RegexOptions.IgnoreCase);
            //Find "Bean" and Replace everywhere to "Joker"

            foreach (ContentRange item in dc.Content.Find(regex))
            {
                if (IsCopyRightContent)
                {
                    item.Delete();
                }
                else
                {
                    item.Replace(@$"{TargetText}");
                }
                
            }

            string savePath = Path.ChangeExtension(SavePdfPath ,".replaced.pdf");
            dc.Save(loadPath);
            //dc.Save(loadPath, new PdfSaveOptions());
            //System.Diagnostics.Process.Start(savePath);
        }

        public static void FindTextAndReplaceImage(string loadPath, string picPath , string SavePdfPath)
        {

            // Load a document intoDocumentCore.
            DocumentCore dc = DocumentCore.Load(loadPath);

            //Find "<signature>" Text and Replace everywhere with the "Smile.png"
            // Please note, Reverse() makes sure that action replace not affects to Find().
            Regex regex = new Regex(@"<signature>", RegexOptions.IgnoreCase);
            Picture picture = new Picture(dc, InlineLayout.Inline(new Size(50, 50)), picPath);
            foreach (ContentRange item in dc.Content.Find(regex).Reverse())
            {
                item.Replace(picture.Content);
            }

            // Save our document into PDF format.
            dc.Save(loadPath);
            //dc.Save(SavePdfPath, new PdfSaveOptions());

            // Open the result for demonstration purposes.
            //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(loadPath) { UseShellExecute = true });
            //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(SavePdfPath) { UseShellExecute = true });
        }

        public static void HeadersAndFooters(string loadPath)
        {
            // distroy priview error
            string ExistText = "SautinSoft.Document 2024.4.24 trial. Get your free 30-day key instantly: https://sautinsoft.com/start-for-free/";

            //string documentPath = @"HeadersAndFooters.docx";
            string documentPath = @"E:\result.docx";
            // Let's create a simple document.
            //DocumentCore dc = DocumentCore.Load(documentPath);
            DocumentCore dc = new DocumentCore();
            // Add a new section in the document.
            Section s = new Section(dc);
            dc.Sections.Add(s);

            // Let's add a paragraph with text.
            Paragraph p = new Paragraph(dc);
            dc.Sections[0].Blocks.Add(p);

            p.ParagraphFormat.Alignment = HorizontalAlignment.Justify;
            p.Content.Start.Insert("سلام چطوری این پیام به صورت آزمایشی ایجاد گردیده است " +
                                   "شما می توانید از ساهتار این متن استفاده و نامه خود را ارسال نمایید شماره 894839 مربوط به نامه می باشد", new CharacterFormat() { Size = 12, FontName = "Arial" });
            p.ParagraphFormat.Alignment = HorizontalAlignment.Right;
            //p.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            // Create a new header with formatted text.
            HeaderFooter header = new HeaderFooter(dc, HeaderFooterType.HeaderDefault);
            header.Content.Start.Insert("Shrek and Donkey travel to the castle and split up to find Fiona.", new CharacterFormat() { Size = 14.0 });

            // Add the header into HeadersFooters collection of the 1st section.
            s.HeadersFooters.Add(header);

            // Create a new footer with some text and image.
            HeaderFooter footer = new HeaderFooter(dc, HeaderFooterType.FooterDefault);

            // Create a paragraph to insert it into the footer.
            Paragraph par = new Paragraph(dc);
            par.Content.Start.Insert("Shrek and Donkey travel to the castle and split up to find Fiona. ", new CharacterFormat() { Size = 14.0 });
            par.ParagraphFormat.Alignment = HorizontalAlignment.Left;

            // Insert image into the paragraph.
            double wPt = LengthUnitConverter.Convert(2, LengthUnit.Centimeter, LengthUnit.Point);
            double hPt = LengthUnitConverter.Convert(2, LengthUnit.Centimeter, LengthUnit.Point);

            Picture pict = new Picture(dc, Layout.Inline(new Size(wPt, hPt)), @"E:\12.jpg");
            par.Inlines.Add(pict);

            // Add the paragraph into the Blocks collection of the footer.
            footer.Blocks.Add(par);

            // Finally, add the footer into 1st section (HeadersFooters collection).
            s.HeadersFooters.Add(footer);
            dc.Save(documentPath);

            // Save the document into DOCX format.
            FindTextAndReplaceAnyText(documentPath, ExistText, " ", "", true);



            // Open the result for demonstration purposes.
            //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }





        //static void Main(string[] args)
        //{

        //    string ResultPath = @"E:\result.docx";
        //    DocumentCore dc = new DocumentCore();
        //    DocumentBuilder db = new DocumentBuilder(dc);
        //    db.StartBookmark("BookMark1");
        //    db.Write("This is <BookMark> 1");
        //    db.EndBookmark("BookMark1");
        //    db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);
        //    dc.Save(ResultPath);

        //    string loadPath = @"E:\result.docx";
        //    string picPath = @"E:\12.png";
        //    string SavePdfPath = @"E:\Replaced_signature.pdf";
        //    //// Load a document intoDocumentCore.
        //    //DocumentCore doc = DocumentCore.Load(loadPath);

        //    ////Find "<signature>" Text and Replace everywhere with the "Smile.png"
        //    //// Please note, Reverse() makes sure that action replace not affects to Find().
        //    //Regex regex = new Regex(@"<BookMark>", RegexOptions.IgnoreCase);
        //    //Picture picture = new Picture(doc, InlineLayout.Inline(new Size(50, 50)), pictPath);
        //    //foreach (ContentRange item in doc.Content.Find(regex).Reverse())
        //    //{
        //    //    item.Replace(picture.Content);
        //    //}

        //    //// Save our document into PDF format.
        //    //string savePath = @"E:\Replaced_signature.pdf";
        //    //doc.Save(savePath, new PdfSaveOptions());
        //    //// Open the result for demonstration purposes.
        //    //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(loadPath) { UseShellExecute = true });
        //    FindTextAndReplaceImage(loadPath, picPath, SavePdfPath);

        //}
    }
}


