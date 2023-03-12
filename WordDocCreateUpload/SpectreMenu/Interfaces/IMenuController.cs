using System.Threading;

namespace WordDocCreateUpload.SpectreMenu
{
    public interface IMenuController
    {
        public IMenuController SetMainMenu(IMenuItem mainMenu);
        public IMenuController AddExitToMainMenu();
        public Task Start();
    }
}
