using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Microsoft.Graph.Models;
using WordDocCreateUpload.Menu;

namespace WordDocCreateUpload
{
    internal class CreateWordDocMenu : MenuItem
    {
        /// <summary>
        /// Navigate to menu, call CreateWordDocFunction
        /// </summary>
        /// <returns>Parent MenuItem</returns>
        /// <exception cref="NullReferenceException"></exception>
        public override async Task<IMenuItem?> navigate()
        {
            _ = Program.GraphApi ?? throw new NullReferenceException("Graph API not set before calling folder command");
            await CreateWordDoc();
            return getParent();
        }
        /// <summary>
        /// Creates word doc in OneDrive
        /// </summary>
        /// <returns>Task</returns>
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
                if(docName.Length >= 5 && docName.Substring(docName.Length - 5, 5) != ".docx")
                {
                    docName += ".docx"; 
                }
                items = await Program.GraphApi!.GetChildItems();
                itemExists = ItemNameExists(items!, docName) != null;
                if (itemExists) { Console.WriteLine($"Item with the name {docName} already exists. Please enter a different name"); }
            } while (itemExists);

            try
            {
                await Program.GraphApi!.UploadWordDoc(wordDocStream, docName);
                Console.WriteLine("Doc Created. Press any key to return.");
            }
            catch(Exception ex)
            {
                Console.WriteLine($"Error, unable to create document: {ex.Message} ");
                Console.WriteLine("Press any key to return.");
            }
            Console.ReadKey(true);
        }
        /// <summary>
        /// Creates word document stream with single run of text
        /// </summary>
        /// <param name="text">single run of text to add</param>
        /// <returns>MemoryStream containg word document</returns>
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

        static DriveItem? ItemNameExists(List<DriveItem> items, string itemName)
        {
            foreach (DriveItem item in items)
            {
                if (item.Name == itemName)
                {
                    return item;
                }
            }
            return null;
        }
    }
}
