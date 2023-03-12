using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tavis.UriTemplates;

namespace WordDocCreateUpload
{
    internal class GraphHelper
    {
        static GraphServiceClient InitializeGraph(Settings settings)
        {
            return InitializeGraphForUserAuth(settings,
                (info, cancel) =>
                {
                    // Display the device code message to
                    // the user. This tells them
                    // where to go to sign in and provides the
                    // code to use.
                    Console.WriteLine(info.Message);
                    return Task.FromResult(0);
                });
        }
        static GraphServiceClient InitializeGraphForUserAuth(Settings settings,
            Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
        {

            DeviceCodeCredential deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
                settings.TenantId, settings.ClientId);

            return new GraphServiceClient(deviceCodeCredential, settings.GraphUserScopes);
        }
        public static async Task<GraphAPI> CreateAsync(Settings settings)
        {
            var userClient = InitializeGraph(settings);
            var driveId = await GetDriveID(userClient);
            var driveRootId = await GetDriveRootID(userClient,driveId);
            var result = new GraphAPI(userClient,driveId,driveRootId);
            return result;
        }

        async static Task<string> GetDriveID(GraphServiceClient userClient)
        {
            var driveItem = await userClient.Me.Drive.GetAsync();
            _ = driveItem?.Id ?? throw new NullReferenceException("Set Drive ID returned null - could not get drive ID");
            return driveItem.Id;
        }
        async static Task<string> GetDriveRootID(GraphServiceClient userClient, string driveId)
        {
            var root = await userClient.Drives[driveId].Root.GetAsync();
            _ = root?.Id ?? throw new NullReferenceException("Set Drive Root ID returned null - could not get root ID");
            return root.Id;
        }

        public static DriveItem? ItemNameExists(List<DriveItem> items, string itemName)
        {
            foreach (DriveItem item in items)
            {
                if (item.Name == itemName)
                {
                    return item;
                }
            }
            return null;
        }

    }
}
