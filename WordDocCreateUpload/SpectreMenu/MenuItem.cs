using System.Data;
using Spectre.Console;

namespace WordDocCreateUpload.SpectreMenu
{
    public class MenuItem : IMenuItem
    {
        private Dictionary<string, IMenuItem> _children;
        private List<string> _childrenMenuItems;
        private string _itemName;
        private IMenuItem _parent;
        public Guid InstanceID { get; private set; }

        public MenuItem()
        {
            _children = new Dictionary<string, IMenuItem>();

            _childrenMenuItems = new List<string>();
        }
        public void addChild(IMenuItem child, int index = 0)
        {
            if (child.getName() == null) { throw new NullReferenceException($"Failed to add child to {_itemName}- child name was not set"); }
            _children.Add(child.getName(), child);
            _childrenMenuItems.Insert(index, child.getName());
        }
        public void removeChild(IMenuItem child)
        {
            _children.Remove(child.getName());
            _childrenMenuItems.Remove(child.getName());
        }
        public IMenuItem createParentLink(IMenuItem parent, int index = 0)
        {
            parent.addChild(this, index);
            _parent = parent;
            return this;
        }
        public string getName()
        {
            return _itemName;
        }
        public IMenuItem getParent()
        {
            return _parent;
        }
        public Dictionary<string, IMenuItem> getChildren()
        {
            return _children;
        }

        public bool childNameExists(string childName)
        {
            foreach (KeyValuePair<string, IMenuItem> child in _children)
            {
                if (childName.Equals(child.Key))
                {
                    return true;
                }
            }
            return false;
        }


        public async virtual Task<IMenuItem> navigate()
        {
            string navSelection = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title(_itemName)
                    .PageSize(10)
                    .MoreChoicesText("[grey](Move up and down to reveal more choices)[/]")
                    .AddChoices(_childrenMenuItems));

            if (navSelection.Equals(MenuConstants.BACK_STRING))
            {
                return _parent;
            }
            else
            {
                return _children[navSelection];
            }
        }
        public IMenuItem setName(string newItemName)
        {
            if (_parent != null)
            {
                if (_parent.childNameExists(newItemName))
                {
                    throw new DuplicateNameException($"Failed to rename {_itemName} Parent already has item named {newItemName}");
                }
                _parent.removeChild(this);
                _itemName = newItemName;
                _parent.addChild(this);
            }
            else
            {
                _itemName = newItemName;
            }
            return this;
        }
        public void displayBackOnlyOption()
        {
            AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title(_itemName)
                    .AddChoices(MenuConstants.BACK_STRING));
        }
    }
}


