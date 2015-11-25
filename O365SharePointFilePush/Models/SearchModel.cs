using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace O365SharePointFilePush.Models
{
    public class SearchModel
    {
        // convert JSON response data into XML for easier consumption from C#
        public static XElement Json2Xml(string json)
        {
            using (XmlDictionaryReader reader = JsonReaderWriterFactory.CreateJsonReader(
                    Encoding.UTF8.GetBytes(json),
                    XmlDictionaryReaderQuotas.Max))
            {
                return XElement.Load(reader);
            }
        }
    }
}