using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordDocCreateUpload.Menu;

namespace WordDocCreateUpload
{
    internal class ChangeFolderMenu : MenuItem
    {

        private DriveItem _driveItem;
    /* Steps
     *  - Get all folders in current GraphAPI Target
     *  - Return each folder as an ID 
     *  - Selecting a folder
     *      - Sets Graph API as that target 
     *      - Navigates to that folder, displaying all folders underneath 
     *  - Add Back button that returns up one level, unless ID == root ID 
     *  - Add Set Upload Destination button that returns to main menu 
     */
        public ChangeFolderMenu(DriveItem driveItem)
        {
            _driveItem = driveItem;
            setName(_driveItem.Name);
        }

        public override async Task<IMenuItem?> navigate()
        {
            _ = Program.GraphApi ?? throw new NullReferenceException("Graph API not set before calling folder command");

            removeAllChildren(); // Clear any existing 

            // Add a Set Upload Destination menu item that returns to main menu... 
            // It needs a reference to the main menu and to return that 
            addChild(new MenuItem().setName(MenuConstants.BACK_STRING));
            addChild(new SetUploadDestinationMenu());

            // Set GraphAPI to this drive Item 
            if(Program.GraphApi.TargetDriveItem != _driveItem)
            {
                Program.GraphApi.TargetDriveItem = _driveItem;
                Program.FolderMenu.setName($"Change Upload Destination - Current Target: {Program.GraphApi.TargetDriveItem.Name}");
            }


            // SO, get all folder items of this item and add them as children... 
            List<DriveItem>? children = await Program.GraphApi.GetChildItems();

            // Check if children is null or empty... that means it's an empty folder 
            _ = children ?? throw new Exception("Need to handle empty folders");

            // If not, filter children for just folders...
            List<DriveItem> childFolders = children; // Add filtering of folders 


            // Set create children menu items...
            foreach (DriveItem childFolder in childFolders)
            {
                var newFolder = new ChangeFolderMenu(childFolder);
                newFolder.createParentLink(this);
                //addChild(new ChangeFolderMenu(childFolder));
            }


            return await base.navigate();
        }
    }
}
