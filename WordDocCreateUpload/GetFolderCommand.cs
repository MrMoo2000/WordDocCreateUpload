using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordDocCreateUpload.SpectreMenu;
using Microsoft.Graph.Models;

namespace WordDocCreateUpload
{
    internal class GetFolderCommand : MenuItem
    {
        public override async Task<IMenuItem> navigate()
        {
            _ = Program.GraphApi ?? throw new NullReferenceException("Graph API not set before calling folder command");

            DriveItem folder = await GetFolderID();
            setName($"Set Upload Destination - Current: {folder.Name}");
            Program.GraphApi.TargetDriveItem = folder;
            return getParent();
        }
        static async Task<DriveItem> GetFolderID()
        {
            ConsoleKeyInfo userChoice;
            string folderName;
            DriveItem? folder;

            do
            {
                Console.WriteLine("Enter Folder name");
                folderName = Console.ReadLine()!;
                folder = await Program.GraphApi!.FolderAtRoot(folderName!);
                if (folder != null)
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
                        folder = await Program.GraphApi.CreateNewFolderAtRoot(folderName);
                    }
                }
            } while (userChoice.Key != ConsoleKey.Y);

            _ = folder ?? throw new NullReferenceException("Folder Returned as Null");

            return folder;
        }
    }
}
