using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace Geradeaus.Visio
{
    public static class VsdxTools
    {
        public static readonly XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
        public static readonly XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        public static Dictionary<int, Master> ParseMasters(Package package, PackagePart documentPart)
        {
            Dictionary<int, Master> masters = new Dictionary<int, Master>();
            PackagePart mastersPart = GetPackageParts(package, documentPart, RelationshipTypes.Masters).ElementAtOrDefault(0);
            if (mastersPart == null) return masters;

            XDocument mastersDocument = GetXMLFromPart(mastersPart);
            var masterElements = mastersDocument.Root.Elements(ns + "Master");

            foreach (var masterElement in masterElements)
            {
                string relId = masterElement.Element(ns + "Rel")?.Attribute(rNs + "id")?.Value;
                if (string.IsNullOrEmpty(relId)) continue;
                PackageRelationship masterRelationship = mastersPart.GetRelationship(relId);
                Uri pageUri = PackUriHelper.ResolvePartUri(mastersPart.Uri, masterRelationship.TargetUri);
                PackagePart masterPart = package.GetPart(pageUri);
                Master master = ParseMaster(masterPart, masterElement);
                masters[master.Id] = master;
            }

            return masters;
        }

        public static Master ParseMaster(PackagePart masterPart, XElement masterElement)
        {
            XDocument masterDocument = GetXMLFromPart(masterPart);
            var shapeElement = masterDocument.Root.Element(ns + "Shapes")?.Element(ns + "Shape");

            return new Master
            {
                Id = int.Parse(masterElement.Attribute("ID")?.Value),
                Name = masterElement.Attribute("Name")?.Value,
                NameU = masterElement.Attribute("NameU")?.Value,
                Text = masterElement.Element(ns + "Text")?.Value,
                UserRows = ParseUserSection(shapeElement),
                PropRows = ParsePropertySection(shapeElement)
            };
        }

        public static Dictionary<int, Page> ParsePages(Package package, PackagePart documentPart)
        {
            Dictionary<int, Page> pages = new Dictionary<int, Page>();
            PackagePart pagesPart = GetPackageParts(package, documentPart, RelationshipTypes.Pages).ElementAtOrDefault(0);
            if (pagesPart == null) return pages;

            XDocument pagesDocument = GetXMLFromPart(pagesPart);
            var pageElements = pagesDocument.Root.Elements(ns + "Page");

            foreach (var pageElement in pageElements)
            {
                string relId = pageElement.Element(ns + "Rel")?.Attribute(rNs + "id")?.Value;
                if (string.IsNullOrEmpty(relId)) continue;
                PackageRelationship pageRelationship = pagesPart.GetRelationship(relId);
                Uri pageUri = PackUriHelper.ResolvePartUri(pagesPart.Uri, pageRelationship.TargetUri);
                PackagePart pagePart = package.GetPart(pageUri);
                Page page = ParsePage(pagePart, pageElement);
                pages[page.Id] = page;
            }

            return pages;
        }

        public static Page ParsePage(PackagePart pagePart, XElement pageElement)
        {
            Page page = new Page();
            page.Id = int.Parse(pageElement.Attribute("ID")?.Value);
            page.NameU = pageElement.Attribute("NameU")?.Value;
            page.Name = pageElement.Attribute("Name")?.Value;

            var pageSheetElement = pageElement.Element(ns + "PageSheet");
            page.UserRows = ParseUserSection(pageSheetElement);
            page.PropRows = ParsePropertySection(pageSheetElement);
            page.Layers = ParseLayerSection(pageSheetElement);

            XDocument pageDocument = GetXMLFromPart(pagePart);
            var shapeElements = pageDocument.Root.Element(ns + "Shapes")?.Elements(ns + "Shape") ?? Enumerable.Empty<XElement>();

            if (shapeElements.Any()) page.Shapes = new Dictionary<int, Shape>();

            foreach (var shapeElement in shapeElements)
            {
                Shape shape = ParseShape(shapeElement);
                page.Shapes[shape.Id] = shape;
            }

            page.Connects = ParseConnects(pageDocument);

            return page;
        }

        public static Shape ParseShape(XElement shapeElement)
        {
            return new Shape
            {
                Id = int.Parse(shapeElement.Attribute("ID")?.Value),
                Name = shapeElement.Attribute("Name")?.Value,
                NameU = shapeElement.Attribute("NameU")?.Value,
                NameId = shapeElement.Attribute("NameID")?.Value,
                Text = shapeElement.Element(ns + "Text")?.Value,
                Master = shapeElement.Attribute("Master")?.Value,
                UserRows = ParseUserSection(shapeElement),
                PropRows = ParsePropertySection(shapeElement)
            };
        }

        public static List<Connect> ParseConnects(XDocument pageDocument)
        {
            return pageDocument?.Root.Element(ns + "Connects")?.Elements(ns + "Connect").Select(
                con => new Connect
                {
                    FromSheet = int.Parse((string)con.Attribute("FromSheet")),
                    FromCell = (string)con.Attribute("FromCell"),
                    FromPart = int.Parse((string)con.Attribute("FromPart")),
                    ToSheet = int.Parse((string)con.Attribute("ToSheet")),
                    ToCell = (string)con.Attribute("ToCell"),
                    ToPart = int.Parse((string)con.Attribute("ToPart"))
                }).ToList() ?? new List<Connect>();
        }

        public static Dictionary<int, Layer> ParseLayerSection(XElement pageSheetElement)
        {
            return pageSheetElement?
                .Elements(ns + "Section")
                .FirstOrDefault(s => (string)s.Attribute("N") == "Layer")
                ?.Elements(ns + "Row")
                .Where(row => !string.IsNullOrEmpty((string)row.Attribute("IX")))
                .ToDictionary(
                    row => int.Parse((string)row.Attribute("IX")),
                    row => new Layer
                    {
                        Index = int.Parse((string)row.Attribute("IX")),
                        Name = (string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Name")
                            ?.Attribute("V"),
                        NameU = (string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "NameUniv")
                            ?.Attribute("V"),
                        Visible = (string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Visible")
                            ?.Attribute("V") == "1",
                        Print = (string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Print")
                            ?.Attribute("V") == "1",
                        Active = (string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Active")
                            ?.Attribute("V") == "1",
                        Lock = (string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Lock")
                            ?.Attribute("V") == "1",
                        Snap = (string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Snap")
                            ?.Attribute("V") == "1",
                        Glue = (string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Glue")
                            ?.Attribute("V") == "1",
                        Color = int.Parse((string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Color")
                            ?.Attribute("V")),
                        ColorTrans = double.Parse((string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Color")
                            ?.Attribute("V"))
                    })
                ?? new Dictionary<int, Layer>();
        }

        public static Dictionary<string, UserRow> ParseUserSection(XElement element)
        {
            return element?
                .Elements(ns + "Section")
                .FirstOrDefault(s => (string)s.Attribute("N") == "User")
                ?.Elements(ns + "Row")
                .Where(row => !string.IsNullOrEmpty((string)row.Attribute("N")))
                .ToDictionary(
                    row => (string)row.Attribute("N"),
                    row => new UserRow
                    {
                        Value = (string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Value")
                            ?.Attribute("V"),
                        Prompt = (string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Prompt")
                            ?.Attribute("V")
                    })
                ?? new Dictionary<string, UserRow>();
        }

        public static Dictionary<string, PropRow> ParsePropertySection(XElement element)
        {
            return element?
                .Elements(ns + "Section")
                .FirstOrDefault(s => (string)s.Attribute("N") == "Property")
                ?.Elements(ns + "Row")
                .Where(row => !string.IsNullOrEmpty((string)row.Attribute("N")))
                .ToDictionary(
                    row => (string)row.Attribute("N"),
                    row => new PropRow
                    {
                        Value = (string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Value")
                            ?.Attribute("V"),
                        Prompt = (string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Prompt")
                            ?.Attribute("V"),
                        Label = (string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Label")
                            ?.Attribute("V"),
                        Format = (string)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Format")
                            ?.Attribute("V"),
                        Type = (int?)row.Elements(ns + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Type")
                            ?.Attribute("V")
                    })
                ?? new Dictionary<string, PropRow>();
        }

        public static List<PackagePart> GetPackageParts(Package filePackage, string relationshipType)
        {
            List<PackagePart> packageParts = new List<PackagePart>();

            foreach (var packageRelationship in filePackage.GetRelationshipsByType(relationshipType))
            {
                Uri uri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), packageRelationship.TargetUri);
                packageParts.Add(filePackage.GetPart(uri));
            }

            return packageParts;
        }

        public static List<PackagePart> GetPackageParts(Package filePackage, PackagePart sourcePart, string relationshipType)
        {
            List<PackagePart> packageParts = new List<PackagePart>();

            foreach (var packageRelationship in sourcePart.GetRelationshipsByType(relationshipType))
            {
                Uri uri = PackUriHelper.ResolvePartUri(sourcePart.Uri, packageRelationship.TargetUri);
                packageParts.Add(filePackage.GetPart(uri));
            }

            return packageParts;
        }

        public static XDocument GetXMLFromPart(PackagePart packagePart)
        {
            using (Stream partStream = packagePart.GetStream())
            {
                return XDocument.Load(partStream);
            }
        }
    }

    public class RelationshipTypes
    {
        public const string Document = "http://schemas.microsoft.com/visio/2010/relationships/document";
        public const string Masters = "http://schemas.microsoft.com/visio/2010/relationships/masters";
        public const string Master = "http://schemas.microsoft.com/visio/2010/relationships/master";
        public const string Pages = "http://schemas.microsoft.com/visio/2010/relationships/pages";
        public const string Page = "http://schemas.microsoft.com/visio/2010/relationships/page";
    }
}
