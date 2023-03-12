using Azure.Identity;
using DocumentFormat.OpenXml.EMMA;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocCreateUpload
{
    internal class GraphAPI
    {
        private GraphServiceClient _userClient;

        private readonly string? _driveId;
        private readonly string? _driveRootId;
        private string? _targetFolderId;

        public GraphAPI(GraphServiceClient userClient, string driveId, string driveRootId)
        {
            _userClient = userClient;
            _driveId = driveId;
            _driveRootId = driveRootId;
        }

        /*
        public async Task<string> GetFolderID(string folderName)
        {
            string? folderID = await FolderIDAtRoot(folderName);
            //folderID ??= await CreateNewFolderAtRoot(folderName);

            return folderID;
        }*/

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

        public async Task<string?> FolderIDAtRoot(string folderName)
        {
            var children = await _userClient.Drives[_driveId].Items[_driveRootId].Children.GetAsync();

            _ = children?.Value ?? throw new NullReferenceException($"Could not get children of {folderName}");

            var item = GraphHelper.ItemNameExists(children.Value, folderName);

            return item?.Id;
        }
        public async Task<List<DriveItem>?> GetChildItems(string itemId)
        {
            var children = await _userClient.Drives[_driveId].Items[itemId].Children.GetAsync();
            _ = children?.Value ?? throw new NullReferenceException($"Could not get children of {itemId}");
            return children.Value;
        }
        public async Task UploadWordDoc(MemoryStream docStream, string docName, string folderId)
        {
            await _userClient.Drives[_driveId].Items[folderId].ItemWithPath(docName).Content.PutAsync(docStream);
        }

    }
}
