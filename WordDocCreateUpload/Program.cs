
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;


using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Runtime.CompilerServices;

namespace WordDocCreateUpload
{
    internal class Program
    {
        private static GraphAPI? _graphApi; 

        static async Task Main()//string[] args)
        {
            /*
             * Steps
             *  - Init Client
             *  - Get Drive ID and Root ID (often reused)
             *  - Check if folder Exists, if not create it
             *  - Enter sentence for Document
             *  - Enter Name for Document
             *      - Check name does not exist in folder 
             *      - If it does, enter name for document
             *  - Upload document to folder 
             */
            
            try
            {
                var settings = Settings.LoadSettings();
                _graphApi = await GraphHelper.CreateAsync(settings);
            }
            catch (InvalidOperationException ex) {
                DisplayErrorExit(ex.Message);
            }

            string folderName = "TestDocs";

            string folderID = await _graphApi!.GetFolderID(folderName);

            Console.WriteLine(folderID);

            Console.WriteLine("Enter sentence for doc:");
            var docContents = Console.ReadLine();

            MemoryStream wordDocStream = CreateWordDocStream(docContents!);

            List<DriveItem>? items;
            string docName;
            bool itemExists;
            do
            {
                Console.WriteLine("Enter Name for word doc:");
                docName = Console.ReadLine()!;
                docName += ".docx"; //add a check, so if the last is .docx don't include
                items = await _graphApi!.GetChildItems(folderID);
                itemExists = GraphHelper.ItemNameExists(items!, docName) != null;
                if (itemExists) { Console.WriteLine($"Item with the name {docName} already exists. Please enter a different name"); }
            } while (itemExists);

            await _graphApi!.UploadWordDoc(wordDocStream,docName, folderID);
            Console.WriteLine("Doc Created");

            Console.ReadKey(true);
        }

        static void DisplayErrorExit(string error)
        {
            Console.WriteLine(error);
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey(true);
            Environment.Exit(0);
        }
        
        static MemoryStream CreateWordDocStream(string text)
        {
            MemoryStream stream = new ();
            using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text(text));
            }
            stream.Position = 0;
            return stream;
        }
    }
}