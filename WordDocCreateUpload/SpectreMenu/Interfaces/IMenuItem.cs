using System;
using System.Collections.Generic;

namespace WordDocCreateUpload.SpectreMenu
{
    delegate void renameMenuItem(IMenuItem item);
    public interface IMenuItem
    {
        Task<IMenuItem> navigate();
        string getName();
        bool childNameExists(string childName);
        IMenuItem getParent();
        Dictionary<string, IMenuItem> getChildren();
        IMenuItem createParentLink(IMenuItem parent, int index = 0);
        IMenuItem setName(string itemName);
        void addChild(IMenuItem child, int index = 0);
        void removeChild(IMenuItem child);
        void displayBackOnlyOption();
    }
}