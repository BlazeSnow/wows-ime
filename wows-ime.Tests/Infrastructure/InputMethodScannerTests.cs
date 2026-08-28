using wows_ime.Core.Infrastructure;
using wows_ime.Core.Models;

namespace wows_ime.Tests.Infrastructure;

public sealed class InputMethodScannerTests
{
    private static readonly string[] KnownWarningCodes =
    [
        InputMethodScanner.CoInitializeFailedWarning,
        InputMethodScanner.CreateProfilesObjectFailedWarning,
        InputMethodScanner.GetLanguageListFailedWarning,
        InputMethodScanner.GetLanguageListEmptyWarning,
        InputMethodScanner.ComExceptionWarning,
        InputMethodScanner.GenericExceptionWarning
    ];

    [Fact]
    public void Scan_ReturnsResult_WithoutThrowing()
    {
        var result = new InputMethodScanner().Scan();

        Assert.NotNull(result);
        Assert.NotNull(result.Candidates);
    }

    [Fact]
    public void Scan_WarningCode_WhenPresent_IsAKnownWarningCode()
    {
        var result = new InputMethodScanner().Scan();

        Assert.True(
            result.WarningCode is null || KnownWarningCodes.Contains(result.WarningCode),
            $"Unexpected warning code: {result.WarningCode}");
    }

    [Fact]
    public void Scan_Candidates_HaveNamesAndDefinedCategories()
    {
        var result = new InputMethodScanner().Scan();

        foreach (var candidate in result.Candidates)
        {
            Assert.False(string.IsNullOrWhiteSpace(candidate.DisplayName));
            Assert.True(Enum.IsDefined(candidate.Category));
        }
    }

    [Fact]
    public void Scan_Candidates_AreSortedByNameAndUnique()
    {
        var result = new InputMethodScanner().Scan();

        var names = result.Candidates.Select(candidate => candidate.DisplayName).ToList();
        Assert.Equal(names.Count, names.Distinct(StringComparer.OrdinalIgnoreCase).Count());
        var sorted = names.OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase).ToList();
        Assert.Equal(sorted, names);
    }
}
