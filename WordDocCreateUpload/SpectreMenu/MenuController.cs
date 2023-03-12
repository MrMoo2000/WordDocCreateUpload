using System;
using System.Collections.Generic;
using System.Text;

namespace WordDocCreateUpload.SpectreMenu
{
    public class MenuController : IMenuController
    {
        private IMenuItem? _mainMenu;

        public IMenuController AddExitToMainMenu()
        {
            _ = _mainMenu ?? throw new NullReferenceException("Main Menu not set - cannot add exit menu item");

            int index = _mainMenu.getChildren().Count;
            new ExitMenuItem().createParentLink(_mainMenu, index);
            return this;
        }

        public IMenuController SetMainMenu(IMenuItem mainMenu)
        {
            _mainMenu = mainMenu;
            return this;
        }
        public async Task Start()
        {
            _ = _mainMenu ?? throw new NullReferenceException("Main Menu not set - cannot add exit menu item");

            IMenuItem? currentMenu = _mainMenu;
            do
            {
                Console.Clear();
                currentMenu = await currentMenu!.navigate();
            } while (0 == 0);
        }
    }
}
