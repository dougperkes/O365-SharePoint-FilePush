using System;
using System.Configuration;

namespace O365SharePointFilePush.Utils
{
    public class SettingsHelper
    {
        public static string ClientId { get; } = ConfigurationManager.AppSettings["ida:ClientId"] ?? ConfigurationManager.AppSettings["ida:ClientID"];
        public static string AppKey { get; } = ConfigurationManager.AppSettings["ida:ClientSecret"] ?? ConfigurationManager.AppSettings["ida:AppKey"] ?? ConfigurationManager.AppSettings["ida:Password"];
        public static string AuthorizationUri { get; } = "https://login.windows.net";
        public static string Authority { get; } = "https://login.windows.net/common/"; //$"https://login.windows.net/{ConfigurationManager.AppSettings["ida:TenantId"]}/";
        public static string AadGraphResourceId { get; } = "https://graph.windows.net";
        public static string DiscoveryServiceResourceId { get; } = "https://api.office.com/discovery/";
        public static Uri DiscoveryServiceEndpointUri { get; } = new Uri("https://api.office.com/discovery/v1.0/me/");
        public static string ClaimTypeObjectIdentifier { get; } = "http://schemas.microsoft.com/identity/claims/objectidentifier";
    }

}