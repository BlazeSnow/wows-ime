using wows_ime.Core.Models;

namespace wows_ime.Core.Abstractions;

public interface IInputMethodScanner
{
    InputMethodScanResult Scan();
}
