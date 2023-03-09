
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;


using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;




namespace WordDocCreateUpload
{
    internal class Program
    {
        private static GraphServiceClient? _userClient;

        static async Task Main()//string[] args)
        {
            try
            {
                var settings = Settings.LoadSettings();
                InitializeGraph(settings);
            }
            catch (InvalidOperationException ex) {
                DisplayErrorExit(ex.Message);
            }

            string folderName = "TestDocssz";

            // We get the ID of the folder, then we pass that to the word doc handling... 
            string folderID = await GetFolderID(folderName);

            Console.WriteLine(folderID);


            Console.WriteLine("Enter sentence for doc:");
            var docContents = Console.ReadLine();

            MemoryStream wordDocStream = CreateWordDocStream(docContents!);


            Console.WriteLine("Enter Name for word doc:");

            var docName = Console.ReadLine();
            // Could validate doc name before creation as well

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

        public static void InitializeGraphForUserAuth(Settings settings,
            Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
        {

            DeviceCodeCredential deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
                settings.TenantId, settings.ClientId);

            _userClient = new GraphServiceClient(deviceCodeCredential, settings.GraphUserScopes);
        }

        /*
         * 
         * Get all DOC names in one drive folder 
         * If any names match, return false
         * Else, true 
         */

        // Change to return ID of folder 
        async static Task<string> GetFolderID(string folderName)
        {
            string? folderID = await FolderIDAtRoot(folderName);
            folderID ??= await CreateNewFolderAtRoot(folderName);

            return folderID;
        }

        // Change to return the ID of folder, null ID if does not exist 
        async static Task<string> FolderIDAtRoot(string folderName)
        {
            var driveItem = await _userClient.Me.Drive.GetAsync();
            var root = await _userClient.Drives[driveItem.Id].Root.GetAsync();
            var children = await _userClient.Drives[driveItem.Id].Items[root.Id].Children.GetAsync();

            foreach (var item in children.Value)
            {
                if(item.Folder != null && item.Folder.GetType() == typeof(Folder) && item.Name == folderName)
                {
                    return item.Id;
                }
            }
            return null;
        }

        // Change to return ID of folder 
        async static Task<string> CreateNewFolderAtRoot(string folderName)
        {
            var driveItem = await _userClient.Me.Drive.GetAsync();
            var root = await _userClient.Drives[driveItem.Id].Root.GetAsync();

            DriveItem newFolder = new DriveItem()
            {
                Name = folderName,
                Folder = new Folder()
            };

            var folder = await _userClient.Drives[driveItem.Id].Items[root.Id].Children.PostAsync(newFolder);

            return folder.Id;

        }
    }
}