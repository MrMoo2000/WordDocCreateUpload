
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
            try
            {
                var settings = Settings.LoadSettings();
                _graphApi = await GraphHelper.CreateAsync(settings);
            }
            catch (InvalidOperationException ex) {
                DisplayErrorExit(ex.Message);
            }

            string folderID = await GetFolderID();

            await CreateWordDoc(folderID);

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

        static async Task<string> GetFolderID()
        {
            ConsoleKeyInfo userChoice;
            string folderName;
            string? folderID;

            do
            {
                Console.WriteLine("Enter Folder name");
                folderName = Console.ReadLine()!;
                folderID = await _graphApi!.FolderIDAtRoot(folderName!);
                if (folderID != null)
                {
                    Console.WriteLine($"Folder {folderName} exists, use it? [y/n]");
                    userChoice = Console.ReadKey(true);
                }
                else
                {
                    Console.WriteLine($"Folder {folderName} does not exist, create it? [y/n]");
                    userChoice = Console.ReadKey(true);
                    if (userChoice.Key == ConsoleKey.Y)
                    {
                        folderID = await _graphApi.CreateNewFolderAtRoot(folderName);
                    }
                }
            } while (userChoice.Key != ConsoleKey.Y);

            _ = folderID ?? throw new NullReferenceException("Folder ID Returned as Null");

            return folderID;
        }

        static async Task CreateWordDoc(string folderID)
        {
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
                items = await _graphApi!.GetChildItems(folderID!);
                itemExists = GraphHelper.ItemNameExists(items!, docName) != null;
                if (itemExists) { Console.WriteLine($"Item with the name {docName} already exists. Please enter a different name"); }
            } while (itemExists);

            await _graphApi!.UploadWordDoc(wordDocStream, docName, folderID!);
            Console.WriteLine("Doc Created");
        }
    }
}