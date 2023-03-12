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
        /// <summary>
        /// Creates Graph API class with drive ID and drive root ID set
        /// </summary>
        /// <param name="settings">Settings with client and tenant IDs</param>
        /// <returns>Graph API instance</returns>
        public static async Task<GraphAPI> CreateAsync(Settings settings)
        {
            var userClient = InitializeGraph(settings);
            var driveId = await GetDriveID(userClient);
            var driveRootId = await GetDriveRootID(userClient, driveId);
            var result = new GraphAPI(userClient, driveId, driveRootId);
            return result;
        }
        /// <summary>
        /// Prompts user to authenticate
        /// </summary>
        /// <param name="settings">Settings with client and tenant IDs</param>
        /// <returns>GraphServiceClient</returns>
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
        /// <summary>
        /// Creates GraphServiceClient with credentials
        /// </summary>
        /// <param name="settings">Settings with client and tenant IDs</param>
        /// <param name="deviceCodePrompt">Authenticated credential</param>
        /// <returns>GraphServiceClient</returns>
        static GraphServiceClient InitializeGraphForUserAuth(Settings settings,
            Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
        {

            DeviceCodeCredential deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
                settings.TenantId, settings.ClientId);

            return new GraphServiceClient(deviceCodeCredential, settings.GraphUserScopes);
        }
        /// <summary>
        /// Gets the drive ID of the authenticated user 
        /// </summary>
        /// <param name="userClient">Authenticated GraphServiceClient</param>
        /// <returns>Drive ID of authenticated user</returns>
        /// <exception cref="NullReferenceException"></exception>
        async static Task<string> GetDriveID(GraphServiceClient userClient)
        {
            var driveItem = await userClient.Me.Drive.GetAsync();
            _ = driveItem?.Id ?? throw new NullReferenceException("Set Drive ID returned null - could not get drive ID");
            return driveItem.Id;
        }
        /// <summary>
        /// Gets the root ID of the drive ID of the authenticated use
        /// </summary>
        /// <param name="userClient">Authenticated GraphServiceClient</param>
        /// <param name="driveId">driveId of authenticated user</param>
        /// <returns>Root Item ID of authenticated user</returns>
        /// <exception cref="NullReferenceException"></exception>
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
