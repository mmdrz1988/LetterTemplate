// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Runtime.CompilerServices;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace Consultant
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("Hello, World! hhh");
            string filePath = "C:\\Users\\mmdr.sasani\\source\\repos\\LetterTemplate\\Consultant\\content\\test1.docx";
            List<string> values = new List<string>() 
            {
               "محمدرضا ساسانی","حمیدرضا تقوی","شاید زندگی پیدا کردن بخش هایی از خودمان در تکه تکه هایی از یک کل منسجم است" +
               "\n" +
               "پیدا کردنی که به اندازه یک عمر طول میکشد" +
               "و تکه هایی که همه جا حضور دارند و فقط کافی است ما در مسیرشان قرار بگیریم ، آن وقت اگر هوشیار باشیم." +
               "شاید بخش هایی از خودمان را ببینیم. بخش هایی از خودمان را در رابطه ها و در آدم هایی که تجربه می کنیم"," i"
            };
            var Result = GetBookmarksFromDoc(filePath);
            Console.WriteLine(Result);
            SaveWordProcessingDocument(filePath, Result , values);
        }
        public static void SaveWordProcessingDocument(string filePath ,List<string> Result , List<string> value)
        {

            byte[] byteArray = File.ReadAllBytes(filePath);
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
                                int i =0 ;
                                foreach (var mb in Result)
                                {
                                    
                                    if (bookmarkStart.Name == mb && bookmarkText != null)
                                    {
                                        bookmarkText.GetFirstChild<Text>().Text = value[i];

                                        Console.WriteLine($"It's correctly : {bookmarkStart.Name}");

                                    }
                                    i++;

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
        public static List<string> GetBookmarksFromDoc(string filePath)
        {
            List<string> BookmarkAchive = new List<string>();
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
                            BookmarkAchive.Add(bookmarkStart.Name);
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
            return BookmarkAchive;
        }

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