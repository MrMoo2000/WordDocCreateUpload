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
        public SetUploadDestinationMenu() { setName("Set Upload Destination"); }
        public async override Task<IMenuItem?> navigate()
        {
            return await Task.Run(Program.MenuController.GetMainMenu);
        }
    }
}
