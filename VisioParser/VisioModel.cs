using System;
using System.Collections.Generic;

namespace Geradeaus.Visio
{
    public class Connect
    {
        public int FromSheet { get; set; }
        public string FromCell { get; set; }
        public int FromPart { get; set; }
        public int ToSheet { get; set; }
        public string ToCell { get; set; }
        public int ToPart { get; set; }
    }

    public class WitConnect
    {
        public int Index { get; set; }
        public int ToSheet { get; set; }
        public string ToPoint { get; set; }
        public string ToPointD { get; set; }
    }

    public class ConnectionPoint
    {
        public string D { get; set; }
    }

    public class UserRow
    {
        public string Value { get; set; }
        public string Prompt { get; set; }
    }

    public class PropRow
    {
        public string Label { get; set; }
        public string Prompt { get; set; }
        public int? Type { get; set; }
        public string Format { get; set; }
        public string Value { get; set; }
    }

    public class Shape
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string NameU { get; set; }
        public string NameId { get; set; }
        public int? Master { get; set; }
        public string Text { get; set; }
        public Dictionary<string, UserRow> UserRows { get; set; }
        public Dictionary<string, PropRow> PropRows { get; set; }
        public Dictionary<string, ConnectionPoint> ConnectionPoints { get; set; }
        public List<WitConnect> WitConnects { get; set; }
        public List<Layer> Layers { get; set; }
    }

    public class Layer
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public string NameU { get; set; }
        public bool Visible { get; set; }
        public bool Print { get; set; }
        public bool Active { get; set; }
        public bool Lock { get; set; }
        public bool Snap { get; set; }
        public bool Glue { get; set; }
        public int Color { get; set; }
        public double ColorTrans { get; set; }
    }

    public class Page
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string NameU { get; set; }
        public Dictionary<string, UserRow> UserRows { get; set; }
        public Dictionary<string, PropRow> PropRows { get; set; }
        public Dictionary<int, Shape> Shapes { get; set; }
        public Dictionary<int, Layer> Layers { get; set; }
        public List<Connect> Connects { get; set; }
    }

    public class Master
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string NameU { get; set; }
        public string Text { get; set; }
        public Dictionary<string, UserRow> UserRows { get; set; }
        public Dictionary<string, PropRow> PropRows { get; set; }
    }

    public class Document
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public string Subject { get; set; }
        public string Description { get; set; }
        public string Creator { get; set; }
        public string Manager { get; set; }
        public string Company { get; set; }
        public string Category { get; set; }
        public string Keywords { get; set; }
        public string Language { get; set; }
        public string TimeCreated { get; set; }
        public string TimeEdited { get; set; }
        public string AppVersion { get; set; }
        public Dictionary<string, UserRow> UserRows { get; set; }
        public Dictionary<string, PropRow> PropRows { get; set; }
        public Dictionary<int, Master> Masters { get; set; }
        public Dictionary<int, Page> Pages { get; set; }
    }

    public class Metadata
    {
        public DateTime ExportTime { get; set; }
        public string VisioParserVersion { get; set; }
    }

    public class VisioModel
    {
        public Metadata Metadata { get; set; } = new Metadata();
        public Document Document { get; set; } = new Document();
    }
}
