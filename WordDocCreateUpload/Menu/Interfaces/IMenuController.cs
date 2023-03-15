using System.Threading;

namespace WordDocCreateUpload.Menu
{
    public interface IMenuController
    {
        public IMenuItem GetMainMenu();
        public IMenuController SetMainMenu(IMenuItem mainMenu);
        public IMenuController AddExitToMainMenu();
        public Task Start();
    }
}
