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
        public async override Task<IMenuItem?> navigate()
        {
            var child = getChildren().First().Value;
            return await child.navigate();
        }
    }
}
