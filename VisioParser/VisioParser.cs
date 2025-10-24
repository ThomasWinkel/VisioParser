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
        public VisioModel VisioModel { get; set; }
        public bool InheritFromMasters { get; set; } = true;
        public VisioParser() { }

        public void ParseVsdx(string vsdxPath)
        {
            using (FileStream fileStream = new FileStream(vsdxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (Package package = Package.Open(fileStream, FileMode.Open, FileAccess.Read))
                {
                    ParseDocument(package);
                }
            }

            if (VisioModel != null)
            {
                VisioModel.Document.Name = Path.GetFileName(vsdxPath);
            }
        }

        public void ExportJson(string jsonPath)
        {
            using (StreamWriter file = File.CreateText(jsonPath))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(file, VisioModel);
            }
        }

        public void ImportJson(string jsonPath)
        {
            VisioModel = null;

            using (StreamReader file = File.OpenText(jsonPath))
            {
                JsonSerializer serializer = new JsonSerializer();
                VisioModel = (VisioModel)serializer.Deserialize(file, typeof(VisioModel));
            }
        }

        private void ParseDocument(Package package)
        {
            VisioModel = null;

            PackagePart documentPart = VsdxTools.GetPackageParts(package, Relationships.Document).ElementAtOrDefault(0);
            if (documentPart == null) return;

            XDocument document = VsdxTools.GetXMLFromPart(documentPart);
            var documentSheetElement = document.Root.Element(Namespaces.Main + "DocumentSheet");
            if (documentSheetElement == null) return;

            VisioModel = new VisioModel();
            VisioModel.Metadata.ExportTime = DateTime.UtcNow;
            VisioModel.Metadata.VisioParserVersion = Assembly.GetEntryAssembly().GetName().Version.ToString();
            VsdxTools.GetDocumentProperties(package, VisioModel.Document);
            VisioModel.Document.UserRows = VsdxTools.ParseUserSection(documentSheetElement);
            VisioModel.Document.PropRows = VsdxTools.ParsePropertySection(documentSheetElement);
            VisioModel.Document.Masters = VsdxTools.ParseMasters(package, documentPart);
            if (InheritFromMasters)
            {
                VisioModel.Document.Pages = VsdxTools.ParsePages(package, documentPart, VisioModel.Document.Masters);
            }
            else
            {
                VisioModel.Document.Pages = VsdxTools.ParsePages(package, documentPart, null);
            }
        }
    }
}
