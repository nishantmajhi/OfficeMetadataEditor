using System.IO;

namespace OfficeMetadataEditor.Tests;

public static class Program
{
    public static int Main()
    {
        var dir = Path.Combine(Path.GetTempPath(), "OfficeMetadataEditor.Tests", Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(dir);

        Console.WriteLine("Office Metadata Editor - test suite");
        Console.WriteLine($"scratch: {dir}");

        try
        {
            ModelTests.Run();
            MetadataServiceTests.Run(dir);
            RecentFilesTests.Run(dir);
            ViewModelTests.Run();
        }
        finally
        {
            try { Directory.Delete(dir, true); } catch { /* scratch only */ }
        }

        return T.Report();
    }
}
