// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace Consultant
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("Hello, World! hhh");
            SaveWordProcessingDocument();
        }
        public static void SaveWordProcessingDocument()
        {
            string[] mybookmark = 
            {
                "From" ,"To","Content"
            };

            byte[] byteArray = File.ReadAllBytes("C:\\Users\\mmdr.sasani\\source\\repos\\LetterTemplate\\Consultant\\content\\test1.docx");
            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length);
                using (var wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    // Do work here
                    var mainPart = wordDoc.MainDocumentPart;
                    if (mainPart != null)
                    {
                        var bookmarks = mainPart.Document.Body.Descendants<BookmarkStart>();
                        if (!bookmarks.Any())
                        {
                            Console.WriteLine("No bookmarks found in the document.");
                        }
                        else
                        {
                            IDictionary<String, BookmarkStart> bookmarkMap =
                                    new Dictionary<String, BookmarkStart>();
                            foreach (var bookmarkStart in bookmarks)
                            {
                                Console.WriteLine($"Bookmark Name: {bookmarkStart.Name}");
                                bookmarkMap[bookmarkStart.Name] = bookmarkStart;
                            }
                            foreach (BookmarkStart bookmarkStart in bookmarkMap.Values)
                            {
                                Run bookmarkText = bookmarkStart.NextSibling<Run>();

                                foreach (var mb in mybookmark)
                                {
                                    if (bookmarkStart.Name == mb)
                                    {
                                        Console.WriteLine($"It's correctly : {bookmarkStart.Name}");

                                    }

                                }
                                if (bookmarkStart.Name == "From")
                                {
                                    Console.WriteLine($"Bookmark From replacement: {bookmarkStart.Name}");
                                    if (bookmarkText != null)
                                    {
                                        bookmarkText.GetFirstChild<Text>().Text = "امیر محمد رضایی";
                                    }
                                }
                                else if (bookmarkStart.Name == "To")
                                {
                                    if (bookmarkText != null)
                                    {
                                        bookmarkText.GetFirstChild<Text>().Text = "محمدرضا ساسانی";
                                    }
                                }
                                else if (bookmarkStart.Name == "Content")
                                {
                                    if (bookmarkText != null)
                                    {
                                        bookmarkText.GetFirstChild<Text>().Text = "گوگل و محتوای کپی و تکراری، مثل کارد و پنیرند و به‌هیچ‌وجه آب‌شان به یک جوب نمی‌رود!\r\n\r\nمانند سایر موتورهای جست‌وجو، گوگل هم نمی‌خواهد ۱۰ محتوای یکسان و مشابه را در صفحهٔ نتایج جست‌وجو به کاربر ارائه دهد؛ پس سعی می‌کند محتواهای منحصربه‌فرد را شناسایی کند و به آن‌ها احترام بگذارد.\r\n\r\nوقتی گوگل با چند محتوای کپی مواجه شود، بین آن‌ها گیر می‌کند و نمی‌تواند بفهمد کدام را به کاربر پیشنهاد دهد و اعتبار لینک‌ها و رتبهٔ بهتر را باید به کدام یکی اعطا کند. این اتفاق به‌شدت به رتبه‌بندی سایت شما آسیب می‌زند.\r\n\r\nبرای اینکه پول خودتان را خرج مقالهٔ کپی‌شده نکنید و مخاطبان خود را به‌خاطر آن از دست دهید، بهتر است به سراغ ابزارهای تشخیص محتوای کپی بروید.\r\n\r\nقبل از اینکه ۱۰ ابزار کاربردی و مناسب برای محتواهای فارسی و انگلیسی را به شما معرفی کنیم، بیایید ببینیم اصلاً به چه چیزی محتوای کپی می‌گوییم!";
                                    }
                                }
                                else
                                {
                                    if (bookmarkText != null)
                                    {
                                        bookmarkText.GetFirstChild<Text>().Text =" ";
                                    }
                                }

                            }
                        }
                    }
                    // ...
                    //GetBookmarksFromDocx(byteArray);
                    wordDoc.MainDocumentPart.Document.Save(); // won't update the original file 
                }
                // Save the file with the new name
                stream.Position = 0;
                Guid guid = Guid.NewGuid();
                File.WriteAllBytes($@"C:\Users\mmdr.sasani\source\repos\LetterTemplate\Consultant\content\Letter\{guid}.docx", stream.ToArray());
            }
        }
        //------------------------------------------------------------------------
        //----------------------------------------------------------------------
        public static void GetBookmarksFromDocx(string filePath)
        {
            //string strDoc = @"C:\Users\mmdr.sasani\source\repos\LetterTemplate\Consultant\content\test.docx";
            //GetBookmarksFromDocx(strDoc);
            //string strDoc = @"E:\resulttest.docx";
            //string txt = "Append text in body - OpenAndAddToWordprocessingStream";
            //Stream stream = File.Open(strDoc, FileMode.Open);
            //OpenAndAddToWordprocessingStream(stream, txt);
            //stream.Close();
            //----------------------------------------------------
            //GetBookmarksFromDocx("E:\\test.docx");
            //OpenWordprocessingDocumentReadonly(@"E:\WordDocTest\Test.docx");
            try
            {
                using var wordDoc = WordprocessingDocument.Open(filePath, false);
                var mainPart = wordDoc.MainDocumentPart;
                if (mainPart != null)
                {
                    var bookmarks = mainPart.Document.Body.Descendants<BookmarkStart>();
                    if (!bookmarks.Any())
                    {
                        Console.WriteLine("No bookmarks found in the document.");
                    }
                    else
                    {
                        foreach (var bookmarkStart in bookmarks)
                        {
                            Console.WriteLine($"Bookmark Name: {bookmarkStart.Name}");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No bookmarks found in the document.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        //-----------------------------------------------------------------------
        static void OpenAndAddToWordprocessingStream(Stream stream, string txt)
        {
            // Open a WordProcessingDocument based on a stream.
            WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, true);

            // Assign a reference to the existing document body.
            MainDocumentPart mainDocumentPart = wordprocessingDocument.MainDocumentPart ?? wordprocessingDocument.AddMainDocumentPart();
            mainDocumentPart.Document ??= new Document();
            Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());

            // Add new text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new DocumentFormat.OpenXml.Drawing.Text(txt));

            // Dispose the document handle.
            wordprocessingDocument.Dispose();

            // Caller must close the stream.


            //  save to another file 

        }
        //------------------------------------------------------------------------

        public static void OpenWordprocessingDocumentReadonly(string filepath)
        {
            // Open a WordprocessingDocument based on a filepath.
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Open(filepath, false))
            {
                // Assign a reference to the existing document body.  
                Body body = wordDocument.MainDocumentPart.Document.Body;
                //var paras = body.OfType<Paragraph>();
                var paras = body.OfType<Paragraph>()
                            .Where(p => p.ParagraphProperties != null &&
                            p.ParagraphProperties.ParagraphStyleId != null &&
                            p.ParagraphProperties.ParagraphStyleId.Val.Value.Contains("Heading1")).ToList();
                foreach (var para in paras)
                {
                    //richTextBox1.Text += para.NextSibling().InnerText + "\n";
                }
                Console.WriteLine(body.InnerText);
                //wordDocument.Close();
            }
        }        
    }
}
//IDictionary<String, BookmarkStart> bookmarkMap =
//    new Dictionary<String, BookmarkStart>();

//foreach (BookmarkStart bookmarkStart in file.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
//{
//    bookmarkMap[bookmarkStart.Name] = bookmarkStart;
//}

//foreach (BookmarkStart bookmarkStart in bookmarkMap.Values)
//{
//    Run bookmarkText = bookmarkStart.NextSibling<Run>();
//    if (bookmarkText != null)
//    {
//        bookmarkText.GetFirstChild<Text>().Text = "blah";
//    }
//}