using Microsoft.SharePoint.Client;
using System.Collections.Generic;

namespace O365SharePointFilePush.Models
{
    public class LibrariesViewModel
    {
        public IEnumerable<List> DocumentLibraries { get; set; }
        public IEnumerable<Web> SubSites { get; set; }
    }
}