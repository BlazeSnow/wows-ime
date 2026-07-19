namespace wows_ime.Core.Models;

public sealed record ScannedImeCandidate(string DisplayName, ImeCategory Category, int Confidence);
