using System.Xml.Linq;

namespace Geradeaus.Visio
{
    public class Relationships
    {
        public const string Document = "http://schemas.microsoft.com/visio/2010/relationships/document";
        public const string Masters = "http://schemas.microsoft.com/visio/2010/relationships/masters";
        public const string Master = "http://schemas.microsoft.com/visio/2010/relationships/master";
        public const string Pages = "http://schemas.microsoft.com/visio/2010/relationships/pages";
        public const string Page = "http://schemas.microsoft.com/visio/2010/relationships/page";
        public const string CoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        public const string ExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    }

    public class Namespaces
    {
        public static readonly XNamespace Main = "http://schemas.microsoft.com/office/visio/2012/main";
        public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static readonly XNamespace Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        public static readonly XNamespace Ep = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        public static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
        public static readonly XNamespace Dcterms = "http://purl.org/dc/terms/";
    }
}
