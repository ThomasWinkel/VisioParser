using Newtonsoft.Json;
using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;

namespace Geradeaus.Visio
{
    public class VisioParser
    {
        public string FullPath { get; }
        public VisioModel VisioModel { get; internal set; }
        public VisioParser(string fullPath)
        {
            FullPath = fullPath;
        }

        public void Parse()
        {
            using (FileStream fileStream = new FileStream(FullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (Package package = Package.Open(fileStream, FileMode.Open, FileAccess.Read))
                {
                    ParseDocument(package);
                }
            }
        }

        public void ExportJson(string fullPath)
        {
            using (StreamWriter file = File.CreateText(fullPath))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(file, VisioModel);
            }
        }

        private void ParseDocument(Package package)
        {
            PackagePart documentPart = VsdxTools.GetPackageParts(package, RelationshipTypes.Document).ElementAtOrDefault(0);
            if (documentPart == null) return;

            XDocument document = VsdxTools.GetXMLFromPart(documentPart);
            var documentSheetElement = document.Root.Element(VsdxTools.ns + "DocumentSheet");
            if (documentSheetElement == null) return;

            VisioModel = new VisioModel();
            VisioModel.Metadata.ExportTime = DateTime.UtcNow;
            VisioModel.Metadata.VisioParserVersion = Assembly.GetEntryAssembly().GetName().Version.ToString();
            VisioModel.Document.UserRows = VsdxTools.ParseUserSection(documentSheetElement);
            VisioModel.Document.PropRows = VsdxTools.ParsePropertySection(documentSheetElement);
            VisioModel.Document.Masters = VsdxTools.ParseMasters(package, documentPart);
            VisioModel.Document.Pages = VsdxTools.ParsePages(package, documentPart);
        }
    }
}
