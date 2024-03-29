﻿using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

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
            var userDrive = await GetDrive(userClient);
            var driveRoot = await GetDriveRoot(userClient, userDrive);
            var result = new GraphAPI(userClient, userDrive, driveRoot);
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
        /// Get drive of authenticated user
        /// </summary>
        /// <param name="userClient">Authenticated GraphServiceClient</param>
        /// <returns>Drive of authenticiated user</returns>
        /// <exception cref="NullReferenceException"></exception>
        async static Task<Drive> GetDrive(GraphServiceClient userClient)
        {
            Drive? driveItem = await userClient.Me.Drive.GetAsync();
            _ = driveItem ?? throw new NullReferenceException("Set Drive returned null - could not get drive");
            return driveItem;
        }
        /// <summary>
        /// Gets the DriveItem that is the root of the users drive
        /// </summary>
        /// <param name="userClient">Authenticated GraphServiceClient</param>
        /// <param name="userDrive">Drive of authenticated user</param>
        /// <returns>root as DriveItem</returns>
        /// <exception cref="NullReferenceException"></exception>
        async static Task<DriveItem> GetDriveRoot(GraphServiceClient userClient, Drive userDrive)
        {
            DriveItem? root = await userClient.Drives[userDrive.Id].Root.GetAsync();
            _ = root ?? throw new NullReferenceException("Set Drive Root returned null - could not get root");
            return root;
        }
    }
}
