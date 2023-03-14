using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Microsoft.Graph.Models;
using WordDocCreateUpload.Menu;

namespace WordDocCreateUpload
{
    internal class Program
    {
        public static GraphAPI? GraphApi; 
        public static IMenuController? MenuController;
        public static IMenuItem? FolderMenu;

        static async Task Main()
        {
            try
            {
                var settings = Settings.LoadSettings();
                GraphApi = await GraphHelper.CreateAsync(settings);
            }
            catch (InvalidOperationException ex) {
                DisplayErrorExit(ex.Message);
                Environment.Exit(0);

            }

            IMenuItem mainMenu = new MenuItem().setName("WordDoc Create and Upload");

            new CreateWordDocCommand().setName($"Create Word Doc").createParentLink(mainMenu);
            //new GetFolderCommand().setName($"Set Upload Destination - Current: Root").createParentLink(mainMenu);

            FolderMenu = new ChangeFolderMenu(GraphApi.DriveRoot).setName($"Change Upload Destination - Current Target: {GraphApi.TargetDriveItem.Name}").createParentLink(mainMenu);
            MenuController = new MenuController().SetMainMenu(mainMenu).AddExitToMainMenu();
            await MenuController.Start();

            Console.ReadKey(true);
        }

        static void DisplayErrorExit(string error)
        {
            Console.WriteLine(error);
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey(true);
            Environment.Exit(0);
        }
    }
}