namespace OfficeMetadataEditor.Tests;

/// <summary>
/// The whole test framework. Deliberately not xunit/nunit: the project had no test
/// dependencies and this keeps it that way, so `dotnet run` is the entire story.
/// Every assertion prints a line; the process exit code is the number of failures.
/// </summary>
public static class T
{
    private static int _passed;
    private static readonly List<string> Failures = new();
    private static string _section = "";

    public static void Section(string name)
    {
        _section = name;
        Console.WriteLine();
        Console.WriteLine($"-- {name}");
    }

    public static void Ok(string what, bool condition)
    {
        if (condition)
        {
            _passed++;
            Console.WriteLine($"   ok   {what}");
        }
        else
        {
            Failures.Add($"{_section}: {what}");
            Console.WriteLine($"   FAIL {what}");
        }
    }

    public static void Equal<TValue>(string what, TValue expected, TValue actual)
    {
        var same = EqualityComparer<TValue>.Default.Equals(expected, actual);
        if (same)
        {
            _passed++;
            Console.WriteLine($"   ok   {what}");
        }
        else
        {
            Failures.Add($"{_section}: {what}");
            Console.WriteLine($"   FAIL {what}");
            Console.WriteLine($"          expected: {Show(expected)}");
            Console.WriteLine($"          actual:   {Show(actual)}");
        }
    }

    /// <summary>Asserts the action throws TException (or a subclass) and hands the
    /// exception back so the caller can assert on its message.</summary>
    public static TException Throws<TException>(string what, Action action)
        where TException : Exception
    {
        try
        {
            action();
        }
        catch (TException ex)
        {
            _passed++;
            Console.WriteLine($"   ok   {what}");
            return ex;
        }
        catch (Exception ex)
        {
            Failures.Add($"{_section}: {what}");
            Console.WriteLine($"   FAIL {what}");
            Console.WriteLine($"          expected: {typeof(TException).Name}");
            Console.WriteLine($"          actual:   {ex.GetType().Name}: {ex.Message}");
            return null!;
        }

        Failures.Add($"{_section}: {what}");
        Console.WriteLine($"   FAIL {what}");
        Console.WriteLine($"          expected: {typeof(TException).Name}");
        Console.WriteLine($"          actual:   nothing was thrown");
        return null!;
    }

    public static void DoesNotThrow(string what, Action action)
    {
        try
        {
            action();
            _passed++;
            Console.WriteLine($"   ok   {what}");
        }
        catch (Exception ex)
        {
            Failures.Add($"{_section}: {what}");
            Console.WriteLine($"   FAIL {what}");
            Console.WriteLine($"          threw: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private static string Show<TValue>(TValue v) => v switch
    {
        null => "(null)",
        string s => $"\"{s}\"",
        DateTime d => d.ToString("o") + $" (Kind={d.Kind})",
        _ => v.ToString() ?? "(null)"
    };

    public static int Report()
    {
        Console.WriteLine();
        Console.WriteLine(new string('=', 60));
        if (Failures.Count == 0)
        {
            Console.WriteLine($"PASS  {_passed} checks, 0 failures");
            return 0;
        }

        Console.WriteLine($"FAIL  {_passed} passed, {Failures.Count} failed:");
        foreach (var f in Failures) Console.WriteLine($"        {f}");
        return Failures.Count;
    }
}
