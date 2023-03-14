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
            catch (Exception ex) {
                DisplayErrorExit(ex.Message);
                Environment.Exit(0);
            }

            IMenuItem mainMenu = new MenuItem().setName("WordDoc Create and Upload");

            new CreateWordDocCommand().setName($"Create Word Doc").createParentLink(mainMenu);

            FolderMenu = new ChangeFolderMenu().setName($"Change Upload Destination - {GetFormmatedCurentTarget()}").createParentLink(mainMenu);
            new FolderMenu(GraphApi.DriveRoot).createParentLink(FolderMenu);

            MenuController = new MenuController().SetMainMenu(mainMenu).AddExitToMainMenu();
            await MenuController.Start();
        }

        static void DisplayErrorExit(string error)
        {
            Console.WriteLine(error);
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey(true);
            Environment.Exit(0);
        }

        public static string GetFormmatedCurentTarget()
        {
            return $"Current Target:[yellow] {GraphApi.TargetDriveItem.Name}[/]";
        }
    }
}