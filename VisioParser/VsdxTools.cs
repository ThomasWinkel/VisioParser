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
        public static Dictionary<int, Master> ParseMasters(Package package, PackagePart documentPart)
        {
            Dictionary<int, Master> masters = new Dictionary<int, Master>();
            PackagePart mastersPart = GetPackageParts(package, documentPart, Relationships.Masters).ElementAtOrDefault(0);
            if (mastersPart == null) return masters;

            XDocument mastersDocument = GetXMLFromPart(mastersPart);
            var masterElements = mastersDocument.Root.Elements(Namespaces.Main + "Master");

            foreach (var masterElement in masterElements)
            {
                string relId = masterElement.Element(Namespaces.Main + "Rel")?.Attribute(Namespaces.R + "id")?.Value;
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
            var shapeElement = masterDocument.Root.Element(Namespaces.Main + "Shapes")?.Element(Namespaces.Main + "Shape");

            return new Master
            {
                Id = int.Parse(masterElement.Attribute("ID")?.Value),
                Name = masterElement.Attribute("Name")?.Value,
                NameU = masterElement.Attribute("NameU")?.Value,
                Text = masterElement.Element(Namespaces.Main + "Text")?.Value,
                UserRows = ParseUserSection(shapeElement),
                PropRows = ParsePropertySection(shapeElement)
            };
        }

        public static Dictionary<int, Page> ParsePages(Package package, PackagePart documentPart, Dictionary<int, Master> masters)
        {
            Dictionary<int, Page> pages = new Dictionary<int, Page>();
            PackagePart pagesPart = GetPackageParts(package, documentPart, Relationships.Pages).ElementAtOrDefault(0);
            if (pagesPart == null) return pages;

            XDocument pagesDocument = GetXMLFromPart(pagesPart);
            var pageElements = pagesDocument.Root.Elements(Namespaces.Main + "Page");

            foreach (var pageElement in pageElements)
            {
                string relId = pageElement.Element(Namespaces.Main + "Rel")?.Attribute(Namespaces.R + "id")?.Value;
                if (string.IsNullOrEmpty(relId)) continue;
                PackageRelationship pageRelationship = pagesPart.GetRelationship(relId);
                Uri pageUri = PackUriHelper.ResolvePartUri(pagesPart.Uri, pageRelationship.TargetUri);
                PackagePart pagePart = package.GetPart(pageUri);
                Page page = ParsePage(pagePart, pageElement, masters);
                pages[page.Id] = page;
            }

            return pages;
        }

        public static Page ParsePage(PackagePart pagePart, XElement pageElement, Dictionary<int, Master> masters)
        {
            Page page = new Page();
            page.Id = int.Parse(pageElement.Attribute("ID")?.Value);
            page.NameU = pageElement.Attribute("NameU")?.Value;
            page.Name = pageElement.Attribute("Name")?.Value;

            var pageSheetElement = pageElement.Element(Namespaces.Main + "PageSheet");
            page.UserRows = ParseUserSection(pageSheetElement);
            page.PropRows = ParsePropertySection(pageSheetElement);
            page.Layers = ParseLayerSection(pageSheetElement);

            XDocument pageDocument = GetXMLFromPart(pagePart);
            var shapeElements = pageDocument.Root.Element(Namespaces.Main + "Shapes")?.Elements(Namespaces.Main + "Shape") ?? Enumerable.Empty<XElement>();

            if (shapeElements.Any()) page.Shapes = new Dictionary<int, Shape>();

            foreach (var shapeElement in shapeElements)
            {
                Shape shape = ParseShape(shapeElement, masters);
                page.Shapes[shape.Id] = shape;
            }

            page.Connects = ParseConnects(pageDocument);

            return page;
        }

        public static Shape ParseShape(XElement shapeElement, Dictionary<int, Master> masters)
        {
            Shape shape = new Shape
            {
                Id = int.Parse(shapeElement.Attribute("ID")?.Value),
                Name = shapeElement.Attribute("Name")?.Value,
                NameU = shapeElement.Attribute("NameU")?.Value,
                NameId = shapeElement.Attribute("NameID")?.Value,
                Text = shapeElement.Element(Namespaces.Main + "Text")?.Value,
                Master = (int?)shapeElement.Attribute("Master")
            };

            if (masters == null || shape.Master == null || !masters.ContainsKey(shape.Master.Value))
            {
                shape.UserRows = ParseUserSection(shapeElement);
                shape.PropRows = ParsePropertySection(shapeElement);
                return shape;
            }

            shape.UserRows = ParseUserSection(shapeElement, masters[shape.Master.Value].UserRows);
            shape.PropRows = ParsePropertySection(shapeElement, masters[shape.Master.Value].PropRows);

            return shape;
        }

        public static List<Connect> ParseConnects(XDocument pageDocument)
        {
            return pageDocument?.Root.Element(Namespaces.Main + "Connects")?.Elements(Namespaces.Main + "Connect").Select(
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
                .Elements(Namespaces.Main + "Section")
                .FirstOrDefault(s => (string)s.Attribute("N") == "Layer")
                ?.Elements(Namespaces.Main + "Row")
                .Where(row => !string.IsNullOrEmpty((string)row.Attribute("IX")))
                .ToDictionary(
                    row => int.Parse((string)row.Attribute("IX")),
                    row => new Layer
                    {
                        Index = int.Parse((string)row.Attribute("IX")),
                        Name = (string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Name")
                            ?.Attribute("V"),
                        NameU = (string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "NameUniv")
                            ?.Attribute("V"),
                        Visible = (string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Visible")
                            ?.Attribute("V") == "1",
                        Print = (string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Print")
                            ?.Attribute("V") == "1",
                        Active = (string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Active")
                            ?.Attribute("V") == "1",
                        Lock = (string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Lock")
                            ?.Attribute("V") == "1",
                        Snap = (string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Snap")
                            ?.Attribute("V") == "1",
                        Glue = (string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Glue")
                            ?.Attribute("V") == "1",
                        Color = int.Parse((string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Color")
                            ?.Attribute("V")),
                        ColorTrans = double.Parse((string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "ColorTrans")
                            ?.Attribute("V"))
                    })
                ?? new Dictionary<int, Layer>();
        }

        public static Dictionary<string, UserRow> ParseUserSection(XElement element)
        {
            return element?
                .Elements(Namespaces.Main + "Section")
                .FirstOrDefault(s => (string)s.Attribute("N") == "User")
                ?.Elements(Namespaces.Main + "Row")
                .Where(row => !string.IsNullOrEmpty((string)row.Attribute("N")))
                .ToDictionary(
                    row => (string)row.Attribute("N"),
                    row => new UserRow
                    {
                        Value = (string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Value")
                            ?.Attribute("V"),
                        Prompt = (string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Prompt")
                            ?.Attribute("V")
                    })
                ?? new Dictionary<string, UserRow>();
        }

        public static Dictionary<string, UserRow> ParseUserSection(XElement element, Dictionary<string, UserRow> masterUserRows)
        {
            Dictionary<string, UserRow> userRows;

            if (masterUserRows == null)
            {
                userRows = new Dictionary<string, UserRow>();
            }
            else
            {
                userRows = masterUserRows.ToDictionary(
                    entry => entry.Key,
                    entry => new UserRow
                    {
                        Value = entry.Value.Value,
                        Prompt = entry.Value.Prompt
                    });
            }

            var rowElements = element?
                .Elements(Namespaces.Main + "Section")
                .FirstOrDefault(s => (string)s.Attribute("N") == "User")
                ?.Elements(Namespaces.Main + "Row");

            if (rowElements == null || !rowElements.Any())
                return masterUserRows;

            foreach (var rowElement in rowElements)
            {
                string rowName = (string)rowElement.Attribute("N");

                if ((string)rowElement.Attribute("Del") == "1")
                {
                    if (userRows.ContainsKey(rowName))
                    {
                        userRows.Remove(rowName);
                    }
                    continue;
                }

                string value = (string)rowElement.Elements(Namespaces.Main + "Cell")
                    .FirstOrDefault(c => (string)c.Attribute("N") == "Value")
                    ?.Attribute("V");
                string prompt = (string)rowElement.Elements(Namespaces.Main + "Cell")
                    .FirstOrDefault(c => (string)c.Attribute("N") == "Prompt")
                    ?.Attribute("V");

                if (!userRows.ContainsKey(rowName))
                {
                    userRows[rowName] = new UserRow
                    {
                        Value = value,
                        Prompt = prompt
                    };
                    continue;
                }

                if (value != null)
                {
                    userRows[rowName].Value = value;
                }
                if (prompt != null)
                {
                    userRows[rowName].Prompt = prompt;
                }
            }

            return userRows;
        }

        public static Dictionary<string, PropRow> ParsePropertySection(XElement element)
        {
            return element?
                .Elements(Namespaces.Main + "Section")
                .FirstOrDefault(s => (string)s.Attribute("N") == "Property")
                ?.Elements(Namespaces.Main + "Row")
                .Where(row => !string.IsNullOrEmpty((string)row.Attribute("N")))
                .ToDictionary(
                    row => (string)row.Attribute("N"),
                    row => new PropRow
                    {
                        Value = (string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Value")
                            ?.Attribute("V"),
                        Prompt = (string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Prompt")
                            ?.Attribute("V"),
                        Label = (string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Label")
                            ?.Attribute("V"),
                        Format = (string)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Format")
                            ?.Attribute("V"),
                        Type = (int?)row.Elements(Namespaces.Main + "Cell")
                            .FirstOrDefault(c => (string)c.Attribute("N") == "Type")
                            ?.Attribute("V")
                    })
                ?? new Dictionary<string, PropRow>();
        }

        public static Dictionary<string, PropRow> ParsePropertySection(XElement element, Dictionary<string, PropRow> masterPropRows)
        {
            Dictionary<string, PropRow> propRows;

            if (masterPropRows == null)
            {
                propRows = new Dictionary<string, PropRow>();
            }
            else
            {
                propRows = masterPropRows.ToDictionary(
                    entry => entry.Key,
                    entry => new PropRow
                    {
                        Label = entry.Value.Label,
                        Prompt = entry.Value.Prompt,
                        Type = entry.Value.Type,
                        Format = entry.Value.Format,
                        Value = entry.Value.Value
                    });
            }

            var rowElements = element?
                .Elements(Namespaces.Main + "Section")
                .FirstOrDefault(s => (string)s.Attribute("N") == "Property")
                ?.Elements(Namespaces.Main + "Row");

            if (rowElements == null || !rowElements.Any())
                return masterPropRows;

            foreach (var rowElement in rowElements)
            {
                string rowName = (string)rowElement.Attribute("N");

                if ((string)rowElement.Attribute("Del") == "1")
                {
                    if (propRows.ContainsKey(rowName))
                    {
                        propRows.Remove(rowName);
                    }
                    continue;
                }

                string label = (string)rowElement.Elements(Namespaces.Main + "Cell")
                    .FirstOrDefault(c => (string)c.Attribute("N") == "Label")
                    ?.Attribute("V");
                string prompt = (string)rowElement.Elements(Namespaces.Main + "Cell")
                    .FirstOrDefault(c => (string)c.Attribute("N") == "Prompt")
                    ?.Attribute("V");
                int? type = (int?)rowElement.Elements(Namespaces.Main + "Cell")
                    .FirstOrDefault(c => (string)c.Attribute("N") == "Type")
                    ?.Attribute("V");
                string format = (string)rowElement.Elements(Namespaces.Main + "Cell")
                    .FirstOrDefault(c => (string)c.Attribute("N") == "Format")
                    ?.Attribute("V");
                string value = (string)rowElement.Elements(Namespaces.Main + "Cell")
                    .FirstOrDefault(c => (string)c.Attribute("N") == "Value")
                    ?.Attribute("V");


                if (!propRows.ContainsKey(rowName))
                {
                    propRows[rowName] = new PropRow
                    {
                        Label = label,
                        Prompt = prompt,
                        Type = type,
                        Format = format,
                        Value = value
                    };
                    continue;
                }
                if (label != null) propRows[rowName].Label = label;
                if (prompt != null) propRows[rowName].Prompt = prompt;
                if (type != null) propRows[rowName].Type = type;
                if (format != null) propRows[rowName].Format = format;
                if (value != null) propRows[rowName].Value = value;
            }

            return propRows;
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

        public static void GetDocumentProperties(Package package, Document document)
        {
            PackagePart corePropertiesPart = GetPackageParts(package, Relationships.CoreProperties).ElementAtOrDefault(0);
            if (corePropertiesPart != null)
            {
                XDocument corePropertiesDocument = GetXMLFromPart(corePropertiesPart);
                document.Title = corePropertiesDocument.Root.Element(Namespaces.Dc + "title")?.Value;
                document.Subject = corePropertiesDocument.Root.Element(Namespaces.Dc + "subject")?.Value;
                document.Description = corePropertiesDocument.Root.Element(Namespaces.Dc + "description")?.Value;
                document.Creator = corePropertiesDocument.Root.Element(Namespaces.Dc + "creator")?.Value;
                document.Category = corePropertiesDocument.Root.Element(Namespaces.Cp + "category")?.Value;
                document.Keywords = corePropertiesDocument.Root.Element(Namespaces.Cp + "keywords")?.Value;
                document.Language = corePropertiesDocument.Root.Element(Namespaces.Dc + "language")?.Value;
                document.TimeCreated = corePropertiesDocument.Root.Element(Namespaces.Dcterms + "created")?.Value;
                document.TimeEdited = corePropertiesDocument.Root.Element(Namespaces.Dcterms + "modified")?.Value;
            }
            
            PackagePart extendedPropertiesPart = GetPackageParts(package, Relationships.ExtendedProperties).ElementAtOrDefault(0);
            if (extendedPropertiesPart != null)
            {
                XDocument extendedPropertiesDocument = GetXMLFromPart(extendedPropertiesPart);
                document.Manager = extendedPropertiesDocument.Root.Element(Namespaces.Ep + "Manager")?.Value;
                document.Company = extendedPropertiesDocument.Root.Element(Namespaces.Ep + "Company")?.Value;
                document.AppVersion = extendedPropertiesDocument.Root.Element(Namespaces.Ep + "AppVersion")?.Value;
            }
        }
    }
}
