
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
        private static GraphServiceClient? _userClient;

        private static string _driveId;
        private static string _driveRootId; 

        static async Task Main()//string[] args)
        {
            try
            {
                var settings = Settings.LoadSettings();
                InitializeGraph(settings);
                _driveId = await GetDriveID();
                _driveRootId = await GetDriveRootID();
            }
            catch (InvalidOperationException ex) {
                DisplayErrorExit(ex.Message);
            }

            string folderName = "TestDocs";

            string folderID = await GetFolderID(folderName);

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
                items = await GetChildItems(folderID);
                itemExists = ItemNameExists(items, docName) == null ? false : true;
                if (itemExists) { Console.WriteLine($"Item with the name {docName} already exists. Please enter a different name"); }
            } while (itemExists);

            await UploadWordDoc(wordDocStream,docName, folderID);
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
            MemoryStream stream = new MemoryStream();
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

        static void InitializeGraph(Settings settings)
        {
            InitializeGraphForUserAuth(settings,
                (info, cancel) =>
                {
                    // Display the device code message to
                    // the user. This tells them
                    // where to go to sign in and provides the
                    // code to use.
                    Console.WriteLine(info.Message);
                    return Task.FromResult(0);
                });
        }

        async static Task<string> GetDriveID()
        {
            var driveItem = await _userClient.Me.Drive.GetAsync();
            return driveItem.Id;
        }
        async static Task<string> GetDriveRootID()
        {
            var root = await _userClient.Drives[_driveId].Root.GetAsync();
            return root.Id;
        }

        public static void InitializeGraphForUserAuth(Settings settings,
            Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
        {

            DeviceCodeCredential deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
                settings.TenantId, settings.ClientId);

            _userClient = new GraphServiceClient(deviceCodeCredential, settings.GraphUserScopes);
        }

        async static Task<string> GetFolderID(string folderName)
        {
            string? folderID = await FolderIDAtRoot(folderName);
            folderID ??= await CreateNewFolderAtRoot(folderName);

            return folderID;
        }

        async static Task<List<DriveItem>?> GetChildItems(string itemId)
        {
            var children = await _userClient.Drives[_driveId].Items[itemId].Children.GetAsync();
            return children.Value;
        }

        // Change to return the ID of folder, null ID if does not exist 
        async static Task<string?> FolderIDAtRoot(string folderName)
        {
            var children = await _userClient.Drives[_driveId].Items[_driveRootId].Children.GetAsync();

            var item = ItemNameExists(children.Value, folderName);

            return item == null ? null : item.Id;
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

        async static Task<string> CreateNewFolderAtRoot(string folderName)
        {
            DriveItem newFolder = new DriveItem()
            {
                Name = folderName,
                Folder = new Folder()
            };

            var folder = await _userClient.Drives[_driveId].Items[_driveRootId].Children.PostAsync(newFolder);

            return folder.Id;

        }

        async static Task UploadWordDoc(MemoryStream docStream, string docName, string folderId)
        {
            await _userClient.Drives[_driveId].Items[folderId].ItemWithPath(docName).Content.PutAsync(docStream);
        }
    }
}