using System.IO;
using System.IO.Compression;
using System.Text;

namespace OfficeMetadataEditor.Tests;

public enum Kind { Docx, Xlsx, Pptx }

/// <summary>
/// Builds real .docx / .xlsx / .pptx packages as raw zips, and reads them back the
/// same way. Nothing here goes through System.IO.Packaging, so a test can tell the
/// difference between "the app wrote what it meant to" and "the library round-tripped
/// its own output" - which is the only way to assert on the scrub.
/// </summary>
public static class Opc
{
    public const string ExtendedPropertiesRel =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    public const string CustomPropertiesRel =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    public const string ThumbnailRel =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";

    /// <summary>The first eight bytes of every OLE2 compound file - i.e. of a
    /// pre-2007 .doc, .xls or .ppt.</summary>
    public static readonly byte[] OleMagic = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    private static (string Part, string ContentType, string Body) MainPart(Kind kind) => kind switch
    {
        Kind.Docx => ("word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
            "<w:body><w:p><w:r><w:t>Hello world</w:t></w:r></w:p></w:body></w:document>"),
        Kind.Xlsx => ("xl/workbook.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" /></sheets></workbook>"),
        Kind.Pptx => ("ppt/presentation.xml",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            "<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" />"),
        _ => throw new ArgumentOutOfRangeException(nameof(kind))
    };

    public static string Extension(Kind kind) => kind switch
    {
        Kind.Docx => ".docx",
        Kind.Xlsx => ".xlsx",
        Kind.Pptx => ".pptx",
        _ => throw new ArgumentOutOfRangeException(nameof(kind))
    };

    public const string DefaultCore =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<cp:coreProperties " +
        "xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" " +
        "xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
        "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" " +
        "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
        "<dc:title>Quarterly Report</dc:title>" +
        "<dc:subject>Finance</dc:subject>" +
        "<dc:creator>Alice Author</dc:creator>" +
        "<cp:keywords>confidential;internal</cp:keywords>" +
        "<dc:description>Not for circulation</dc:description>" +
        "<cp:lastModifiedBy>Bob Editor</cp:lastModifiedBy>" +
        "<cp:revision>7</cp:revision>" +
        "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2026-03-04T09:15:00Z</dcterms:created>" +
        "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">2026-05-06T17:45:00Z</dcterms:modified>" +
        "<cp:category>Reports</cp:category>" +
        "<cp:contentStatus>Draft</cp:contentStatus>" +
        "<cp:version>2.1</cp:version>" +
        "<dc:language>en-GB</dc:language>" +
        "<dc:identifier>DOC-00042</dc:identifier>" +
        "</cp:coreProperties>";

    public const string DefaultApp =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" " +
        "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
        "<Template>Normal.dotm</Template>" +
        "<TotalTime>4321</TotalTime>" +
        "<Pages>12</Pages>" +
        "<Words>3000</Words>" +
        "<Company>Acme Holdings Ltd</Company>" +
        "<Manager>Carol Manager</Manager>" +
        "<HyperlinkBase>C:\\Users\\alice\\Documents</HyperlinkBase>" +
        "<AppVersion>16.0000</AppVersion>" +
        "</Properties>";

    public const string DefaultCustom =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" " +
        "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
        "<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"2\" name=\"ClientCode\">" +
        "<vt:lpwstr>ACME-001</vt:lpwstr></property></Properties>";

    /// <summary>A byte pattern that is not valid UTF-8 and not valid XML, used as the
    /// embedded image so that "the body survived" is a claim about bytes.</summary>
    public static readonly byte[] MediaBytes =
        { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0xFF, 0xFE, 0x00, 0x01, 0x80, 0x7F };

    public static void Build(
        string path,
        Kind kind = Kind.Docx,
        string? core = null,
        string? app = DefaultApp,
        string? custom = DefaultCustom,
        bool thumbnail = true)
    {
        var (mainPart, mainType, mainBody) = MainPart(kind);
        core ??= DefaultCore;

        var types = new StringBuilder();
        types.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        types.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        types.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        types.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        types.Append("<Default Extension=\"png\" ContentType=\"image/png\"/>");
        types.Append("<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>");
        types.Append($"<Override PartName=\"/{mainPart}\" ContentType=\"{mainType}\"/>");
        types.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
        if (app is not null)
            types.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
        if (custom is not null)
            types.Append("<Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/>");
        types.Append("</Types>");

        var rels = new StringBuilder();
        rels.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        rels.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        rels.Append($"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"{mainPart}\"/>");
        rels.Append("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>");
        if (app is not null)
            rels.Append($"<Relationship Id=\"rId3\" Type=\"{ExtendedPropertiesRel}\" Target=\"docProps/app.xml\"/>");
        if (custom is not null)
            rels.Append($"<Relationship Id=\"rId4\" Type=\"{CustomPropertiesRel}\" Target=\"docProps/custom.xml\"/>");
        if (thumbnail)
            rels.Append($"<Relationship Id=\"rId5\" Type=\"{ThumbnailRel}\" Target=\"docProps/thumbnail.jpeg\"/>");
        rels.Append("</Relationships>");

        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
        if (File.Exists(path)) File.Delete(path);

        using var zip = new ZipArchive(File.Create(path), ZipArchiveMode.Create);
        AddText(zip, "[Content_Types].xml", types.ToString());
        AddText(zip, "_rels/.rels", rels.ToString());
        AddText(zip, mainPart, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + mainBody);
        AddText(zip, "docProps/core.xml", core);
        if (app is not null) AddText(zip, "docProps/app.xml", app);
        if (custom is not null) AddText(zip, "docProps/custom.xml", custom);
        if (thumbnail) AddBytes(zip, "docProps/thumbnail.jpeg", new byte[] { 0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10 });
        AddBytes(zip, "media/image1.png", MediaBytes);
    }

    private static void AddText(ZipArchive zip, string name, string content) =>
        AddBytes(zip, name, Encoding.UTF8.GetBytes(content));

    private static void AddBytes(ZipArchive zip, string name, byte[] content)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var s = entry.Open();
        s.Write(content, 0, content.Length);
    }

    // ---- reading back -------------------------------------------------------

    public static string[] Parts(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        return zip.Entries.Select(e => e.FullName).ToArray();
    }

    public static bool HasPart(string path, string part) => Parts(path).Contains(part);

    public static byte[]? ReadPart(string path, string part)
    {
        using var zip = ZipFile.OpenRead(path);
        var entry = zip.GetEntry(part);
        if (entry is null) return null;
        using var s = entry.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    public static string? ReadText(string path, string part)
    {
        var bytes = ReadPart(path, part);
        return bytes is null ? null : Encoding.UTF8.GetString(bytes);
    }

    /// <summary>The package-level relationship types present in _rels/.rels.</summary>
    public static string[] RelationshipTypes(string path)
    {
        var xml = ReadText(path, "_rels/.rels") ?? "";
        return System.Text.RegularExpressions.Regex
            .Matches(xml, "Type=\"([^\"]+)\"")
            .Select(m => m.Groups[1].Value)
            .ToArray();
    }
}
