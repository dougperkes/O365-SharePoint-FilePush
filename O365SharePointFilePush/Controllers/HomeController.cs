using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.SharePoint.Client;
using O365SharePointFilePush.Models;
using O365SharePointFilePush.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Xml.Linq;

namespace O365SharePointFilePush.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [Authorize]
        public async Task<ActionResult> Sites()
        {
            List<SearchResult> results = new List<SearchResult>();

            var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            var tenantId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;

            AuthenticationContext authContext = new AuthenticationContext(string.Format("{0}/{1}", SettingsHelper.AuthorizationUri, tenantId), new ADALTokenCache(signInUserId));

            try
            {
                DiscoveryClient discClient = new DiscoveryClient(SettingsHelper.DiscoveryServiceEndpointUri,
                    async () =>
                    {
                        var authResultDisc = await authContext.AcquireTokenSilentAsync(SettingsHelper.DiscoveryServiceResourceId, new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey), new UserIdentifier(userObjectId, UserIdentifierType.UniqueId));

                        return authResultDisc.AccessToken;
                    });

                var dcr = await discClient.DiscoverCapabilityAsync("RootSite");

                ViewBag.ResourceId = dcr.ServiceResourceId;

                var authResultSharePoint = await authContext.AcquireTokenSilentAsync(dcr.ServiceResourceId, 
                    new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey), new UserIdentifier(userObjectId, UserIdentifierType.UniqueId));

                var sharePointToken = authResultSharePoint.AccessToken;

                results.Add(new SearchResult
                {
                    Title = "Root Site",
                    Path = dcr.ServiceResourceId,
                });

                var query = "/search/query?querytext='contentclass:sts_site'&trimduplicates=true&rowlimit=50&SelectProperties='WebTemplate,Title,Path,SiteLogo,contentclass'";

                await ExecuteSearchQuery(results, dcr, sharePointToken, query);

            }
            catch (AdalException exception)
            {
                //handle token acquisition failure
                if (exception.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                {
                    authContext.TokenCache.Clear();

                    ViewBag.ErrorMessage = "AuthorizationRequired";
                }
            }

            return View(results);
        }

        [Authorize]
        public async Task<ActionResult> Libraries(string path)
        {
            //List<SearchResult> results = new List<SearchResult>();
            LibrariesViewModel vm = new LibrariesViewModel();

            var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            var tenantId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;

            AuthenticationContext authContext = new AuthenticationContext(string.Format("{0}/{1}", SettingsHelper.AuthorizationUri, tenantId), new ADALTokenCache(signInUserId));

            try
            {
                DiscoveryClient discClient = new DiscoveryClient(SettingsHelper.DiscoveryServiceEndpointUri,
                    async () =>
                    {
                        var authResultDisc = await authContext.AcquireTokenSilentAsync(SettingsHelper.DiscoveryServiceResourceId, new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey), new UserIdentifier(userObjectId, UserIdentifierType.UniqueId));

                        return authResultDisc.AccessToken;
                    });

                var dcr = await discClient.DiscoverCapabilityAsync("RootSite");

                ViewBag.ResourceId = dcr.ServiceResourceId;

                var authResultSharePoint = await authContext.AcquireTokenSilentAsync(dcr.ServiceResourceId,
                    new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey), new UserIdentifier(userObjectId, UserIdentifierType.UniqueId));

                var sharePointToken = authResultSharePoint.AccessToken;

                ClientContext ctx = new ClientContext(path);
                ctx.ExecutingWebRequest += (sender, e) => {
                    e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + sharePointToken;
                };

                var allWebs = ctx.Web.Webs;
                var webs = ctx.LoadQuery(
                    allWebs.Include(web => web.Title, web => web.Url)
                    .Where(web => web.WebTemplate != "APP" && web.WebTemplate != "ACCSVC")
                    );
                var lists = ctx.Web.Lists;
                var docLibs = ctx.LoadQuery(
                    lists.Include(list => list.Title, list => list.Id, list => list.BaseType)
                    .Where(list => list.BaseTemplate == 101 && !list.Hidden)
                    );

                //ctx.Load(libs, splists => splists.Include(list => list.Title, list => list.Id, list => list.BaseType);
                ctx.ExecuteQuery();
                vm.SubSites = webs;
                vm.DocumentLibraries = docLibs;
            }
            catch (AdalException exception)
            {
                //handle token acquisition failure
                if (exception.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                {
                    authContext.TokenCache.Clear();

                    ViewBag.ErrorMessage = "AuthorizationRequired";
                }
            }

            return View(vm);

        }

        [Authorize]
        public ActionResult Upload(string path, Guid docLibId)
        {
            var filePath = Server.MapPath("~/Content/SampleFiles");
            var files = new System.IO.DirectoryInfo(filePath).GetFileSystemInfos();
            ViewBag.Path = path;
            ViewBag.DocLibId = docLibId;
            return View(files);
        }


        [Authorize]
        public async Task<ActionResult> UploadFile(string path, Guid docLibId, string fileName)
        {
            var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            var tenantId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;

            var directoryPath = Server.MapPath("~/Content/SampleFiles");
            var theFileToUpload = new System.IO.FileInfo(System.IO.Path.Combine(directoryPath, fileName));            

            AuthenticationContext authContext = new AuthenticationContext(string.Format("{0}/{1}", SettingsHelper.AuthorizationUri, tenantId), new ADALTokenCache(signInUserId));

            try
            {
                DiscoveryClient discClient = new DiscoveryClient(SettingsHelper.DiscoveryServiceEndpointUri,
                    async () =>
                    {
                        var authResultDisc = await authContext.AcquireTokenSilentAsync(SettingsHelper.DiscoveryServiceResourceId, new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey), new UserIdentifier(userObjectId, UserIdentifierType.UniqueId));

                        return authResultDisc.AccessToken;
                    });

                var dcr = await discClient.DiscoverCapabilityAsync("RootSite");

                ViewBag.ResourceId = dcr.ServiceResourceId;

                var authResultSharePoint = await authContext.AcquireTokenSilentAsync(dcr.ServiceResourceId,
                    new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey), new UserIdentifier(userObjectId, UserIdentifierType.UniqueId));

                var sharePointToken = authResultSharePoint.AccessToken;

                ClientContext ctx = new ClientContext(path);
                ctx.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + sharePointToken;
                };

                //Get the document library we want to upload to
                var docLib = ctx.Web.Lists.GetById(docLibId);

                ctx.Load(docLib.RootFolder);
                ctx.ExecuteQuery();

                byte[] bytes = System.IO.File.ReadAllBytes(theFileToUpload.FullName);

                ctx.RequestTimeout = 3600000;

                FileCreationInformation fci = new FileCreationInformation();
                fci.Content = bytes;
                fci.Url = $"{docLib.RootFolder.ServerRelativeUrl}/{theFileToUpload.Name}";
                fci.Overwrite = true;

                var newFile = docLib.RootFolder.Files.Add(fci);
                ctx.Load(newFile);
                ctx.ExecuteQuery();

                var absoluteUrl = new Uri(ctx.Url)
                    .GetLeftPart(UriPartial.Authority) + fci.Url;

                ViewBag.DocumentUrl = absoluteUrl;

            }
            catch (AdalException exception)
            {
                //handle token acquisition failure
                if (exception.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                {
                    authContext.TokenCache.Clear();

                    ViewBag.ErrorMessage = "AuthorizationRequired";
                }
            }

            return View();

        }

        private static async Task ExecuteSearchQuery(List<SearchResult> results, CapabilityDiscoveryResult dcr, string sharePointToken, string query)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + sharePointToken);
            client.DefaultRequestHeaders.Add("Accept", "application/json; odata=verbose");
            using (HttpResponseMessage response = await client.GetAsync(new Uri(dcr.ServiceEndpointUri + query, UriKind.Absolute)))
            {
                if (response.IsSuccessStatusCode)
                {
                    XElement root = SearchModel.Json2Xml(await response.Content.ReadAsStringAsync());
                    var items = root.Descendants("RelevantResults").Elements("Table").Elements("Rows").Elements("results").Elements("item");
                    foreach (var item in items)
                    {
                        //loop through the properties returned for this item
                        var newItem = new SearchResult();
                        foreach (var prop in item.Descendants("item"))
                        {
                            if (prop.Element("Key").Value == "Title")
                                newItem.Title = prop.Element("Value").Value;
                            else if (prop.Element("Key").Value == "Path")
                                newItem.Path = prop.Element("Value").Value;
                            else if (prop.Element("Key").Value == "SiteLogo")
                                newItem.SiteLogo = prop.Element("Value").Value;
                            else if (prop.Element("Key").Value == "contentclass")
                                newItem.ContentClass = prop.Element("Value").Value;
                        }

                        //only return site collections in primary domain...not the onedrive or public domains
                        //this would probably be better placed in the original search query
                        if (newItem.Path.ToLower().Contains(dcr.ServiceResourceId.ToLower()))
                            results.Add(newItem);
                    }
                }
            }
        }
    }
}