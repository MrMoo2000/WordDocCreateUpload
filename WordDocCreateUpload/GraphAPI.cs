using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.Workbook.Functions.Beta_Dist;
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
        private readonly Drive? _userDrive;
        /// <summary>
        /// Item ID of the root
        /// </summary>
        private readonly DriveItem? _driveRoot;


        public DriveItem? TargetDriveItem;

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
            _driveRoot = driveRoot;
            TargetDriveItem = _driveRoot;
        }

        public async Task<DriveItem> CreateNewFolderAtRoot(string folderName)
        {
            DriveItem newFolder = new()
            {
                Name = folderName,
                Folder = new Folder()
            };

            var folder = await _userClient.Drives[_userDrive.Id].Items[_driveRoot.Id].Children.PostAsync(newFolder);

            _ = folder ?? throw new System.NullReferenceException("Could not get folder ID");

            return folder;
        }

        public async Task<DriveItem?> FolderAtRoot(string folderName)
        {
            var children = await _userClient.Drives[_userDrive.Id].Items[_driveRoot.Id].Children.GetAsync();

            _ = children?.Value ?? throw new NullReferenceException($"Could not get children of {folderName}");

            var item = GraphHelper.ItemNameExists(children.Value, folderName);

            return item;
        }
        /// <summary>
        /// Gets child items of an item ID 
        /// </summary>
        /// <returns>Child Items of target drive item</returns>
        /// <exception cref="NullReferenceException"></exception>
        public async Task<List<DriveItem>?> GetChildItems()
        {
            var children = await _userClient.Drives[_userDrive.Id].Items[TargetDriveItem.Id].Children.GetAsync();
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
            await _userClient.Drives[_userDrive.Id].Items[TargetDriveItem.Id].ItemWithPath(docName).Content.PutAsync(docStream);
        }


    }
}
