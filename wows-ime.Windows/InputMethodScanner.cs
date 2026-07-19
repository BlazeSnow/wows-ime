using System.Runtime.InteropServices;
using wows_ime.Core.Abstractions;
using wows_ime.Core.Models;

namespace wows_ime.Windows;

public sealed class InputMethodScanner : IInputMethodScanner
{
    public const string CoInitializeFailedWarning = "Tsf/CoInitializeFailed";
    public const string CreateProfilesObjectFailedWarning = "Tsf/CreateProfilesObjectFailed";
    public const string GetLanguageListFailedWarning = "Tsf/GetLanguageListFailed";
    public const string GetLanguageListEmptyWarning = "Tsf/GetLanguageListEmpty";
    public const string ComExceptionWarning = "Tsf/ComException";
    public const string GenericExceptionWarning = "Tsf/GenericException";

    private const uint CoinitApartmentThreaded = 0x2;
    private const uint ClsctxInprocServer = 0x1;

    public InputMethodScanResult Scan()
    {
        var candidates = new Dictionary<string, ScannedImeCandidate>(StringComparer.OrdinalIgnoreCase);
        var result = ReadCandidatesFromTsf();
        foreach (var candidate in result.Candidates)
        {
            UpsertCandidate(candidates, candidate);
        }

        return new InputMethodScanResult(
            candidates.Values.OrderBy(item => item.DisplayName, StringComparer.CurrentCultureIgnoreCase).ToList(),
            result.WarningCode,
            result.WarningArguments);
    }

    private static InputMethodScanResult ReadCandidatesFromTsf()
    {
        var candidates = new List<ScannedImeCandidate>();
        var coInitHr = CoInitializeEx(IntPtr.Zero, CoinitApartmentThreaded);
        var shouldUninitialize = coInitHr is 0 or 1;
        if (coInitHr < 0 && coInitHr != unchecked((int)0x80010106))
        {
            return Warning(candidates, CoInitializeFailedWarning, FormatHResult(coInitHr));
        }

        try
        {
            var profilesPtr = CreateInputProcessorProfilesCom();
            if (profilesPtr == IntPtr.Zero)
            {
                return Warning(candidates, CreateProfilesObjectFailedWarning);
            }

            try
            {
                var hr = GetLanguageList(profilesPtr, out var langPtr, out var langCount);
                if (hr < 0 || langPtr == IntPtr.Zero || langCount == 0)
                {
                    return hr < 0
                        ? Warning(candidates, GetLanguageListFailedWarning, FormatHResult(hr))
                        : Warning(candidates, GetLanguageListEmptyWarning);
                }

                try
                {
                    for (var i = 0; i < langCount; i++)
                    {
                        var langId = (ushort)Marshal.ReadInt16(langPtr, checked((int)i * sizeof(short)));
                        if (!IsTargetLanguageProfile(langId))
                        {
                            continue;
                        }

                        hr = EnumLanguageProfiles(profilesPtr, langId, out var enumProfilesPtr);
                        if (hr < 0 || enumProfilesPtr == IntPtr.Zero)
                        {
                            continue;
                        }

                        try
                        {
                            while (true)
                            {
                                var items = new TfLanguageProfile[1];
                                hr = EnumLanguageProfilesNext(enumProfilesPtr, 1, items, out var fetched);
                                if (hr != 0 || fetched == 0)
                                {
                                    break;
                                }

                                var item = items[0];
                                var enabledHr = IsEnabledLanguageProfile(
                                    profilesPtr,
                                    ref item.Clsid,
                                    item.LanguageId,
                                    ref item.ProfileGuid,
                                    out var enabled);
                                if (enabledHr < 0 || enabled == 0)
                                {
                                    continue;
                                }

                                var name = GetTsfProfileDescription(profilesPtr, item);
                                if (string.IsNullOrWhiteSpace(name))
                                {
                                    continue;
                                }

                                name = name.Trim();
                                if (IsNoiseImeName(name))
                                {
                                    continue;
                                }

                                var category = InferCategoryFromLangId(item.LanguageId)
                                    ?? InferCategoryFromName(name)
                                    ?? ImeCategory.ChineseSimplified;
                                candidates.Add(new ScannedImeCandidate(name, category, 10));
                            }
                        }
                        finally
                        {
                            _ = Marshal.Release(enumProfilesPtr);
                        }
                    }
                }
                finally
                {
                    CoTaskMemFree(langPtr);
                }
            }
            finally
            {
                _ = Marshal.Release(profilesPtr);
            }
        }
        catch (COMException ex)
        {
            return Warning(candidates, ComExceptionWarning, FormatHResult(ex.HResult), ex.Message);
        }
        catch (Exception ex)
        {
            return Warning(candidates, GenericExceptionWarning, ex.Message);
        }
        finally
        {
            if (shouldUninitialize)
            {
                CoUninitialize();
            }
        }

        return new InputMethodScanResult(candidates);
    }

    private static InputMethodScanResult Warning(
        IReadOnlyList<ScannedImeCandidate> candidates,
        string warningCode,
        params string[] warningArguments) =>
        new(candidates, warningCode, warningArguments);

    private static string FormatHResult(int hresult) => $"0x{hresult:X8}";

    private static IntPtr CreateInputProcessorProfilesCom()
    {
        var clsid = new Guid("33C53A50-F456-4884-B049-85FD643ECFED");
        var iid = new Guid("1F02B6C5-7842-4EE6-8A0B-9A24183A95CA");
        return CoCreateInstance(ref clsid, IntPtr.Zero, ClsctxInprocServer, ref iid, out var ptr) < 0 ? IntPtr.Zero : ptr;
    }

    private static string? GetTsfProfileDescription(IntPtr profilesPtr, TfLanguageProfile item)
    {
        var hr = GetLanguageProfileDescription(profilesPtr, ref item.Clsid, item.LanguageId, ref item.ProfileGuid, out var bstrPtr);
        if (hr < 0 || bstrPtr == IntPtr.Zero)
        {
            return null;
        }

        try
        {
            return Marshal.PtrToStringBSTR(bstrPtr);
        }
        finally
        {
            SysFreeString(bstrPtr);
        }
    }

    private static void UpsertCandidate(IDictionary<string, ScannedImeCandidate> candidates, ScannedImeCandidate candidate)
    {
        if (!candidates.TryGetValue(candidate.DisplayName, out var existing) || candidate.Confidence > existing.Confidence)
        {
            candidates[candidate.DisplayName] = candidate;
        }
    }

    private static bool IsTargetLanguageProfile(ushort languageId) => languageId is
        0x0804 or 0x0404 or 0x0C04 or 0x1004 or 0x1404 or 0x0411;

    private static bool IsNoiseImeName(string name) =>
        name.Contains("输入体验", StringComparison.OrdinalIgnoreCase) ||
        name.Contains("Input Experience", StringComparison.OrdinalIgnoreCase);

    private static ImeCategory? InferCategoryFromLangId(ushort languageId) => languageId switch
    {
        0x0804 or 0x1004 => ImeCategory.ChineseSimplified,
        0x0404 or 0x0C04 or 0x1404 => ImeCategory.ChineseTraditional,
        0x0411 => ImeCategory.Japanese,
        _ => null
    };

    private static ImeCategory? InferCategoryFromName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return null;
        }

        if (name.Contains("速成", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("倉頡", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("仓颉", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("注音", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("Quick", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("Cangjie", StringComparison.OrdinalIgnoreCase))
        {
            return ImeCategory.ChineseTraditional;
        }

        if (name.Contains("Japanese", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("日文", StringComparison.OrdinalIgnoreCase))
        {
            return ImeCategory.Japanese;
        }

        return name.Contains("拼音", StringComparison.OrdinalIgnoreCase) ||
               name.Contains("五笔", StringComparison.OrdinalIgnoreCase)
            ? ImeCategory.ChineseSimplified
            : null;
    }

    private static int GetLanguageList(IntPtr profilesPtr, out IntPtr languageIds, out uint count) =>
        GetVtableDelegate<TfGetLanguageListDelegate>(profilesPtr, 15)(profilesPtr, out languageIds, out count);

    private static int EnumLanguageProfiles(IntPtr profilesPtr, ushort languageId, out IntPtr enumerator) =>
        GetVtableDelegate<TfEnumLanguageProfilesDelegate>(profilesPtr, 16)(profilesPtr, languageId, out enumerator);

    private static int GetLanguageProfileDescription(IntPtr profilesPtr, ref Guid clsid, ushort languageId, ref Guid profileGuid, out IntPtr description) =>
        GetVtableDelegate<TfGetLanguageProfileDescriptionDelegate>(profilesPtr, 12)(profilesPtr, ref clsid, languageId, ref profileGuid, out description);

    private static int EnumLanguageProfilesNext(IntPtr enumerator, uint count, TfLanguageProfile[] profiles, out uint fetched) =>
        GetVtableDelegate<TfEnumLanguageProfilesNextDelegate>(enumerator, 4)(enumerator, count, profiles, out fetched);

    private static int IsEnabledLanguageProfile(IntPtr profilesPtr, ref Guid clsid, ushort languageId, ref Guid profileGuid, out int enabled) =>
        GetVtableDelegate<TfIsEnabledLanguageProfileDelegate>(profilesPtr, 18)(profilesPtr, ref clsid, languageId, ref profileGuid, out enabled);

    private static T GetVtableDelegate<T>(IntPtr comPtr, int methodIndex) where T : Delegate
    {
        var vtable = Marshal.ReadIntPtr(comPtr);
        var methodPtr = Marshal.ReadIntPtr(vtable, methodIndex * IntPtr.Size);
        return Marshal.GetDelegateForFunctionPointer<T>(methodPtr);
    }

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfGetLanguageListDelegate(IntPtr @this, out IntPtr languageIds, out uint count);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfEnumLanguageProfilesDelegate(IntPtr @this, ushort languageId, out IntPtr enumerator);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfGetLanguageProfileDescriptionDelegate(IntPtr @this, ref Guid clsid, ushort languageId, ref Guid profileGuid, out IntPtr description);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfEnumLanguageProfilesNextDelegate(IntPtr @this, uint count, [Out] TfLanguageProfile[] profiles, out uint fetched);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfIsEnabledLanguageProfileDelegate(IntPtr @this, ref Guid clsid, ushort languageId, ref Guid profileGuid, out int enabled);

    [DllImport("ole32.dll")]
    private static extern int CoCreateInstance(ref Guid clsid, IntPtr outer, uint context, ref Guid iid, out IntPtr instance);

    [DllImport("ole32.dll")]
    private static extern void CoTaskMemFree(IntPtr memory);

    [DllImport("ole32.dll")]
    private static extern int CoInitializeEx(IntPtr reserved, uint coInit);

    [DllImport("ole32.dll")]
    private static extern void CoUninitialize();

    [DllImport("oleaut32.dll")]
    private static extern void SysFreeString(IntPtr bstr);

    [StructLayout(LayoutKind.Sequential)]
    private struct TfLanguageProfile
    {
        public Guid Clsid;
        public ushort LanguageId;
        public Guid CategoryId;
        public int IsActive;
        public Guid ProfileGuid;
    }
}
