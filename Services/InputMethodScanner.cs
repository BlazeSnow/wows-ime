using System.Runtime.InteropServices;

namespace wows_ime.Services;

internal sealed record ScannedImeCandidate(string DisplayName, ImeCategory Category, int Confidence);

internal static class InputMethodScanner
{
    private const uint COINIT_APARTMENTTHREADED = 0x2;
    private const uint CLSCTX_INPROC_SERVER = 0x1;

    internal static IEnumerable<ScannedImeCandidate> ReadCandidates(out string? warning)
    {
        warning = null;
        var candidates = new Dictionary<string, ScannedImeCandidate>(StringComparer.OrdinalIgnoreCase);
        foreach (var candidate in ReadCandidatesFromTsf(out warning))
        {
            UpsertCandidate(candidates, candidate);
        }

        return candidates.Values.OrderBy(item => item.DisplayName, StringComparer.CurrentCultureIgnoreCase);
    }

    private static IEnumerable<ScannedImeCandidate> ReadCandidatesFromTsf(out string? warning)
    {
        warning = null;
        var candidates = new List<ScannedImeCandidate>();
        var coInitHr = CoInitializeEx(IntPtr.Zero, COINIT_APARTMENTTHREADED);
        var shouldUninitialize = coInitHr == 0 || coInitHr == 1;
        if (coInitHr < 0 && coInitHr != unchecked((int)0x80010106))
        {
            warning = AppResources.Format("Tsf/CoInitializeFailed", $"0x{coInitHr:X8}");
            return candidates;
        }

        try
        {
            var profilesPtr = CreateInputProcessorProfilesCom();
            if (profilesPtr == IntPtr.Zero)
            {
                warning = AppResources.Get("Tsf/CreateProfilesObjectFailed");
                return candidates;
            }

            try
            {
                var hr = GetLanguageList(profilesPtr, out var langPtr, out var langCount);
                if (hr < 0 || langPtr == IntPtr.Zero || langCount == 0)
                {
                    warning = hr < 0
                        ? AppResources.Format("Tsf/GetLanguageListFailed", $"0x{hr:X8}")
                        : AppResources.Get("Tsf/GetLanguageListEmpty");
                    return candidates;
                }

                try
                {
                    for (var i = 0; i < langCount; i++)
                    {
                        var langId = (ushort)Marshal.ReadInt16(langPtr, (int)i * sizeof(short));
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
                                var items = new TF_LANGUAGEPROFILE[1];
                                hr = EnumLanguageProfilesNext(enumProfilesPtr, 1, items, out var fetched);
                                if (hr != 0 || fetched == 0)
                                {
                                    break;
                                }

                                var item = items[0];
                                var enabledHr = IsEnabledLanguageProfile(
                                    profilesPtr,
                                    ref item.clsid,
                                    item.langid,
                                    ref item.guidProfile,
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

                                name = NormalizeImeDisplayName(name);
                                if (IsNoiseImeName(name))
                                {
                                    continue;
                                }

                                var category = InferCategoryFromLangId(item.langid)
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
            warning = AppResources.Format("Tsf/ComException", $"0x{ex.HResult:X8}", ex.Message);
            return candidates;
        }
        catch (Exception ex)
        {
            warning = AppResources.Format("Tsf/GenericException", ex.Message);
            return candidates;
        }
        finally
        {
            if (shouldUninitialize)
            {
                CoUninitialize();
            }
        }

        return candidates;
    }

    private static IntPtr CreateInputProcessorProfilesCom()
    {
        // CLSID_TF_InputProcessorProfiles
        var clsid = new Guid("33C53A50-F456-4884-B049-85FD643ECFED");
        var iid = new Guid("1F02B6C5-7842-4EE6-8A0B-9A24183A95CA");
        var hr = CoCreateInstance(ref clsid, IntPtr.Zero, CLSCTX_INPROC_SERVER, ref iid, out var ptr);
        if (hr < 0)
        {
            return IntPtr.Zero;
        }

        return ptr;
    }

    private static string? GetTsfProfileDescription(IntPtr profilesPtr, TF_LANGUAGEPROFILE item)
    {
        var hr = GetLanguageProfileDescription(profilesPtr, ref item.clsid, item.langid, ref item.guidProfile, out var bstrPtr);
        if (hr >= 0 && bstrPtr != IntPtr.Zero)
        {
            try
            {
                var tsfDescription = Marshal.PtrToStringBSTR(bstrPtr);
                if (!string.IsNullOrWhiteSpace(tsfDescription))
                {
                    return tsfDescription;
                }
            }
            finally
            {
                SysFreeString(bstrPtr);
            }
        }

        return null;
    }

    private static void UpsertCandidate(IDictionary<string, ScannedImeCandidate> candidates, ScannedImeCandidate candidate)
    {
        if (!candidates.TryGetValue(candidate.DisplayName, out var existing) || candidate.Confidence > existing.Confidence)
        {
            candidates[candidate.DisplayName] = candidate;
        }
    }

    private static bool IsTargetLanguageProfile(ushort? langId) => langId is
        0x0804 or // zh-CN
        0x0404 or // zh-TW
        0x0C04 or // zh-HK
        0x1004 or // zh-SG
        0x1404 or // zh-MO
        0x0411;   // ja-JP

    private static string NormalizeImeDisplayName(string name)
    {
        if (name.Contains("wetype", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, AppResources.Get("Ime/Weixin"), StringComparison.OrdinalIgnoreCase))
        {
            return AppResources.Get("Ime/Weixin");
        }

        return name.Trim();
    }

    private static bool IsNoiseImeName(string name) =>
        name.Contains("输入体验", StringComparison.OrdinalIgnoreCase) ||
        name.Contains("Input Experience", StringComparison.OrdinalIgnoreCase);

    private static ImeCategory? InferCategoryFromLangId(ushort? langId) => langId switch
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

        if (name.Contains("拼音", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("五笔", StringComparison.OrdinalIgnoreCase) ||
            name.Contains(AppResources.Get("Ime/Weixin"), StringComparison.OrdinalIgnoreCase))
        {
            return ImeCategory.ChineseSimplified;
        }

        return null;
    }

    private static int GetLanguageList(IntPtr profilesPtr, out IntPtr langPtr, out uint langCount)
    {
        // ITfInputProcessorProfiles::GetLanguageList is vtable slot 15 (IUnknown + 12 methods before it).
        var fn = GetVtableDelegate<TfGetLanguageListDelegate>(profilesPtr, 15);
        return fn(profilesPtr, out langPtr, out langCount);
    }

    private static int EnumLanguageProfiles(IntPtr profilesPtr, ushort langId, out IntPtr enumProfilesPtr)
    {
        // ITfInputProcessorProfiles::EnumLanguageProfiles is vtable slot 16.
        var fn = GetVtableDelegate<TfEnumLanguageProfilesDelegate>(profilesPtr, 16);
        return fn(profilesPtr, langId, out enumProfilesPtr);
    }

    private static int GetLanguageProfileDescription(IntPtr profilesPtr, ref Guid clsid, ushort langId, ref Guid profileGuid, out IntPtr bstrPtr)
    {
        // ITfInputProcessorProfiles::GetLanguageProfileDescription is vtable slot 12.
        var fn = GetVtableDelegate<TfGetLanguageProfileDescriptionDelegate>(profilesPtr, 12);
        return fn(profilesPtr, ref clsid, langId, ref profileGuid, out bstrPtr);
    }

    private static int EnumLanguageProfilesNext(IntPtr enumProfilesPtr, uint count, TF_LANGUAGEPROFILE[] buffer, out uint fetched)
    {
        // IEnumTfLanguageProfiles::Next is vtable slot 4.
        var fn = GetVtableDelegate<TfEnumLanguageProfilesNextDelegate>(enumProfilesPtr, 4);
        return fn(enumProfilesPtr, count, buffer, out fetched);
    }

    private static int IsEnabledLanguageProfile(IntPtr profilesPtr, ref Guid clsid, ushort langId, ref Guid profileGuid, out int enabled)
    {
        // ITfInputProcessorProfiles::IsEnabledLanguageProfile is vtable slot 18.
        var fn = GetVtableDelegate<TfIsEnabledLanguageProfileDelegate>(profilesPtr, 18);
        return fn(profilesPtr, ref clsid, langId, ref profileGuid, out enabled);
    }

    private static T GetVtableDelegate<T>(IntPtr comPtr, int methodIndex) where T : Delegate
    {
        var vtable = Marshal.ReadIntPtr(comPtr);
        var methodPtr = Marshal.ReadIntPtr(vtable, methodIndex * IntPtr.Size);
        return Marshal.GetDelegateForFunctionPointer<T>(methodPtr);
    }

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfGetLanguageListDelegate(IntPtr @this, out IntPtr ppLangId, out uint pulCount);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfEnumLanguageProfilesDelegate(IntPtr @this, ushort langid, out IntPtr ppEnum);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfGetLanguageProfileDescriptionDelegate(
        IntPtr @this,
        ref Guid rclsid,
        ushort langid,
        ref Guid guidProfile,
        out IntPtr pbstrProfile);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfEnumLanguageProfilesNextDelegate(
        IntPtr @this,
        uint ulCount,
        [Out] TF_LANGUAGEPROFILE[] pProfile,
        out uint pcFetch);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TfIsEnabledLanguageProfileDelegate(
        IntPtr @this,
        ref Guid rclsid,
        ushort langid,
        ref Guid guidProfile,
        out int pfEnable);

    [DllImport("ole32.dll")]
    private static extern int CoCreateInstance(
        ref Guid rclsid,
        IntPtr pUnkOuter,
        uint dwClsContext,
        ref Guid riid,
        out IntPtr ppv);

    [DllImport("ole32.dll")]
    private static extern void CoTaskMemFree(IntPtr pv);

    [DllImport("ole32.dll")]
    private static extern int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);

    [DllImport("ole32.dll")]
    private static extern void CoUninitialize();

    [DllImport("oleaut32.dll")]
    private static extern void SysFreeString(IntPtr bstr);

    [StructLayout(LayoutKind.Sequential)]
    private struct TF_LANGUAGEPROFILE
    {
        public Guid clsid;
        public ushort langid;
        public Guid catid;
        public int fActive;
        public Guid guidProfile;
    }
}
