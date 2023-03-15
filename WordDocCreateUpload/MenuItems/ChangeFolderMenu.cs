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
        /// <summary>
        /// Menu Item for under main menu. Needed for a different name
        /// Calls and navigates to first child
        /// </summary>
        /// <returns>Task of MenuItem of first child</returns>
        public async override Task<IMenuItem?> navigate()
        {
            var child = getChildren().First().Value;
            return await child.navigate();
        }
    }
}
