using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordDocCreateUpload.Menu;

namespace WordDocCreateUpload
{
    internal class CreateWordDocCommand : MenuItem
    {

        public override async Task<IMenuItem?> navigate()
        {
            _ = Program.GraphApi ?? throw new NullReferenceException("Graph API not set before calling folder command");
            await CreateWordDoc();
            return getParent();
        }

        static async Task CreateWordDoc()
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
                items = await Program.GraphApi!.GetChildItems();
                itemExists = GraphHelper.ItemNameExists(items!, docName) != null;
                if (itemExists) { Console.WriteLine($"Item with the name {docName} already exists. Please enter a different name"); }
            } while (itemExists);

            await Program.GraphApi!.UploadWordDoc(wordDocStream, docName);
            Console.WriteLine("Doc Created");
        }
        static MemoryStream CreateWordDocStream(string text)
        {
            MemoryStream stream = new();
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
