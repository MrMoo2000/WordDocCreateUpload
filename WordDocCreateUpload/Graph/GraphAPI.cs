using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace WordDocCreateUpload
{
    internal class GraphAPI
    {
        /// <summary>
        /// Configured GraphServiceClient 
        /// </summary>
        private GraphServiceClient _userClient;
        /// <summary>
        /// Reference to users drive 
        /// </summary>
        private readonly Drive _userDrive;
        /// <summary>
        /// Reference root drive item 
        /// </summary>
        public DriveItem DriveRoot { get; private set; }
        /// <summary>
        /// Drive item to be maniuplated when APIs called
        /// </summary>
        public DriveItem TargetDriveItem;

        /// <summary>
        /// Requires configured GraphServiceClient 
        /// </summary>
        /// <param name="userClient">Configured GraphServiceClient </param>
        /// <param name="driveId">Drive ID of the user</param>
        /// <param name="driveRootId">Root item ID</param>
        public GraphAPI(GraphServiceClient userClient, Drive userDrive, DriveItem driveRoot)
        {
            _userClient = userClient;
            _userDrive = userDrive;
            DriveRoot = driveRoot;
            TargetDriveItem = DriveRoot;
        }
        /// <summary>
        /// Create a new folder
        /// </summary>
        /// <param name="folderName">Name of the new folder</param>
        /// <returns>DriveItem of folder</returns>
        /// <exception cref="NullReferenceException"></exception>
        public async Task<DriveItem> CreateNewFolder(string folderName)
        {
            DriveItem newFolder = new()
            {
                Name = folderName,
                Folder = new Folder()
            };

            DriveItem? folder = await _userClient.Drives[_userDrive!.Id].Items[TargetDriveItem!.Id].Children.PostAsync(newFolder);

            _ = folder ?? throw new NullReferenceException("Could not get folder");

            return folder;
        }
        /// <summary>
        /// Gets child items of an item ID 
        /// </summary>
        /// <returns>Child Items of target drive item</returns>
        /// <exception cref="NullReferenceException"></exception>
        public async Task<List<DriveItem>?> GetChildItems()
        {
            var children = await _userClient.Drives[_userDrive!.Id].Items[TargetDriveItem!.Id].Children.GetAsync();
            _ = children?.Value ?? throw new NullReferenceException($"Could not get children of {TargetDriveItem.Id}");
            return children.Value;
        }
        /// <summary>
        /// Uploads a word document to target drive item
        /// </summary>
        /// <param name="docStream">Stream containg word document</param>
        /// <param name="docName">Name of Word Doc</param>
        /// <returns></returns>
        public async Task UploadWordDoc(MemoryStream docStream, string docName)
        {
            await _userClient.Drives[_userDrive!.Id].Items[TargetDriveItem!.Id].ItemWithPath(docName).Content.PutAsync(docStream);
        }

    }
}
