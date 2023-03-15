using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordDocCreateUpload.Menu;

namespace WordDocCreateUpload
{
    internal class SetUploadDestinationMenu : MenuItem
    {
        /// <summary>
        /// Constructor will set name
        /// </summary>
        public SetUploadDestinationMenu() { setName("[cyan]Set Upload Destination[/]"); }
        /// <summary>
        /// Navigate, will return main menu
        /// </summary>
        /// <returns>Main Menu</returns>
        public async override Task<IMenuItem?> navigate()
        {
            return await Task.Run(Program.MenuController.GetMainMenu);
        }
    }
}
