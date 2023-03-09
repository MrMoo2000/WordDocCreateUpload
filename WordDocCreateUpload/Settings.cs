using Microsoft.Extensions.Configuration;



namespace WordDocCreateUpload
{
    internal class Settings
    {
        public string? ClientId { get; set; }
        public string? TenantId { get; set; }
        public string[] GraphUserScopes =
        {
            "user.read",
            "Files.Read"
        };
        public static Settings LoadSettings()
        {
            IConfiguration config = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: false)
                .AddJsonFile($"appsettings.Development.json", optional: true)
                .Build();

            return config.GetRequiredSection("Settings").Get<Settings>() ??
                throw new Exception("Could not load app settings");
        }

    }
}
