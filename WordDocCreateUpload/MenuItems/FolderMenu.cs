using DocumentFormat.OpenXml.Drawing;
using Microsoft.Graph.Models;
using Spectre.Console;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordDocCreateUpload.Menu;

namespace WordDocCreateUpload
{
    internal class FolderMenu : MenuItem
    {
        private DriveItem _driveItem;
        public FolderMenu(DriveItem driveItem)
        {
            _driveItem = driveItem;
            setName($"[yellow]{_driveItem.Name}[/]");
        }

        public override async Task<IMenuItem?> navigate()
        {
            _ = Program.GraphApi ?? throw new NullReferenceException("Graph API not set before calling folder command");

            AnsiConsole.MarkupLine(Program.GetFormmatedCurentTarget());

            removeAllChildren(); // Clear any existing 

            // Add a Set Upload Destination menu item that returns to main menu... 
            // It needs a reference to the main menu and to return that 
            if(_driveItem.Id != Program.GraphApi.DriveRoot.Id)
            {
                addChild(new MenuItem().setName(MenuConstants.BACK_STRING));
            }
            addChild(new SetUploadDestinationMenu());

            // Set GraphAPI to this drive Item 
            if(Program.GraphApi.TargetDriveItem != _driveItem)
            {
                Program.GraphApi.TargetDriveItem = _driveItem;
                Program.FolderMenu.setName($"Change Upload Destination - Current Target:[yellow] {Program.GraphApi.TargetDriveItem.Name}[/]");
            }

            // SO, get all folder items of this item and add them as children... 
            List<DriveItem>? children = await Program.GraphApi.GetChildItems();

            // If not, filter children for just folders...
            List<DriveItem> childFolders = children.Where(c => c.Folder != null).ToList();

            // Set create children menu items...
            foreach (DriveItem childFolder in childFolders)
            {
                var newFolder = new FolderMenu(childFolder);
                newFolder.createParentLink(this);
            }
            return await base.navigate();
        }
    }
}
