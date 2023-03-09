
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;


using Azure.Identity;
using Microsoft.Graph;




namespace WordDocCreateUpload
{
    internal class Program
    {
        public string folderName = "TestDocs";
        private static GraphServiceClient? _userClient;

        static void Main()//string[] args)
        {
            try
            {
                var settings = Settings.LoadSettings();
                InitializeGraph(settings);
            }
            catch (InvalidOperationException ex) {
                DisplayErrorExit(ex.Message);
            }


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
        async Task<bool> DriveTest()
        {
            await Task.Delay(100);
            return false;
        }
    }
}