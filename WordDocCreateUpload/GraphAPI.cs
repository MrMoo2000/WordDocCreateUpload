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
        /// Drive ID of the user
        /// </summary>
        private readonly string? _driveId;
        /// <summary>
        /// Item ID of the root
        /// </summary>
        private readonly string? _driveRootId;
        /// <summary>
        /// Requires configured GraphServiceClient 
        /// </summary>
        /// <param name="userClient">Configured GraphServiceClient </param>
        /// <param name="driveId">Drive ID of the user</param>
        /// <param name="driveRootId">Root item ID</param>
        public GraphAPI(GraphServiceClient userClient, string driveId, string driveRootId)
        {
            _userClient = userClient;
            _driveId = driveId;
            _driveRootId = driveRootId;
        }

        /// <summary>
        /// Creates a new folder at drive root
        /// </summary>
        /// <param name="folderName">Name of folder to be created</param>
        /// <returns>ID of newly created folder</returns>
        /// <exception cref="System.NullReferenceException"></exception>
        public async Task<string> CreateNewFolderAtRoot(string folderName)
        {
            DriveItem newFolder = new()
            {
                Name = folderName,
                Folder = new Folder()
            };

            var folder = await _userClient.Drives[_driveId].Items[_driveRootId].Children.PostAsync(newFolder);

            _ = folder?.Id ?? throw new System.NullReferenceException("Could not get folder ID");

            return folder.Id;
        }
        /// <summary>
        /// Gets folder ID by name at root of drive 
        /// </summary>
        /// <param name="folderName">Name of folder to check for</param>
        /// <returns>Returns folder ID or null if not found</returns>
        /// <exception cref="NullReferenceException"></exception>
        public async Task<string?> FolderIDAtRoot(string folderName)
        {
            var children = await _userClient.Drives[_driveId].Items[_driveRootId].Children.GetAsync();

            _ = children?.Value ?? throw new NullReferenceException($"Could not get children of {folderName}");

            var item = GraphHelper.ItemNameExists(children.Value, folderName);

            return item?.Id;
        }
        /// <summary>
        /// Gets child items of an item ID 
        /// </summary>
        /// <param name="itemId">Item ID to get child items of</param>
        /// <returns>Child Items of an Item ID</returns>
        /// <exception cref="NullReferenceException"></exception>
        public async Task<List<DriveItem>?> GetChildItems(string itemId)
        {
            var children = await _userClient.Drives[_driveId].Items[itemId].Children.GetAsync();
            _ = children?.Value ?? throw new NullReferenceException($"Could not get children of {itemId}");
            return children.Value;
        }
        /// <summary>
        /// Uploads a word document to a specific itemID
        /// </summary>
        /// <param name="docStream">Stream containg word document</param>
        /// <param name="docName">Name of Word Doc</param>
        /// <param name="itemID">ID of item to upload under</param>
        /// <returns></returns>
        public async Task UploadWordDoc(MemoryStream docStream, string docName, string itemID)
        {
            await _userClient.Drives[_driveId].Items[itemID].ItemWithPath(docName).Content.PutAsync(docStream);
        }

    }
}
