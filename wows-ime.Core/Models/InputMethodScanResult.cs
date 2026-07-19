namespace wows_ime.Core.Models;

public sealed record InputMethodScanResult(
    IReadOnlyList<ScannedImeCandidate> Candidates,
    string? WarningCode = null,
    IReadOnlyList<string>? WarningArguments = null);
