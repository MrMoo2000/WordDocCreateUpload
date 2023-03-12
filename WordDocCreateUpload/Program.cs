using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Microsoft.Graph.Models;
using WordDocCreateUpload.SpectreMenu;

namespace WordDocCreateUpload
{
    internal class Program
    {
        public static GraphAPI? GraphApi; 

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
            new GetFolderCommand().setName($"Set Upload Destination - Current: Root").createParentLink(mainMenu);
            IMenuController menuController = new MenuController().SetMainMenu(mainMenu).AddExitToMainMenu();
            await menuController.Start();

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