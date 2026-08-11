#include <atlbase.h>
#include <atlcom.h>
#include <gdiplus.h>

#include <string>

#include "FormulaOleContract.h"
#include "FormulaOleObject.h"
#include "resource.h"

namespace
{
void ConfigureProcessDpiAwareness() noexcept
{
    // The OLE LocalServer converts the PowerPoint-generated enhanced metafile
    // to legacy CF_METAFILEPICT on demand. GetWinMetaFileBits derives that WMF
    // resolution from this process's reference DC, so leaving the LocalServer
    // DPI-unaware can turn a 125/150/175%-DPI PowerPoint preview into a small
    // upper-left rendering inside an otherwise correctly sized OLE host box.
    // Set the awareness before ATL/OLE/GDI+ touch any display DC. Resolve the
    // newer API dynamically so the binary still has a safe legacy fallback.
    using SetProcessDpiAwarenessContextFn = BOOL(WINAPI*)(HANDLE);
    HMODULE user32 = GetModuleHandleW(L"user32.dll");
    auto setContext = user32 == nullptr
        ? nullptr
        : reinterpret_cast<SetProcessDpiAwarenessContextFn>(
            GetProcAddress(user32, "SetProcessDpiAwarenessContext"));
    if (setContext != nullptr)
    {
        // DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 is the pseudo-handle -4.
        if (setContext(reinterpret_cast<HANDLE>(-4)))
            return;
    }
    SetProcessDPIAware();
}

void TraceModule(const char* message) noexcept
{
    wchar_t localApplicationData[32768] = {};
    const DWORD length = GetEnvironmentVariableW(
        L"LOCALAPPDATA",
        localApplicationData,
        static_cast<DWORD>(std::size(localApplicationData)));
    if (length == 0 || length >= std::size(localApplicationData))
        return;
    const std::wstring root = std::wstring(localApplicationData, length) + L"\\VisualTeX\\office";
    const std::wstring marker = root + L"\\ole-server-trace.enabled";
    if (GetFileAttributesW(marker.c_str()) == INVALID_FILE_ATTRIBUTES)
        return;
    const std::wstring path = root + L"\\ole-server-trace.log";
    HANDLE file = CreateFileW(
        path.c_str(),
        FILE_APPEND_DATA,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        nullptr,
        OPEN_ALWAYS,
        FILE_ATTRIBUTE_NORMAL,
        nullptr);
    if (file == INVALID_HANDLE_VALUE)
        return;
    SYSTEMTIME now = {};
    GetLocalTime(&now);
    char line[512] = {};
    const int characters = sprintf_s(
        line,
        "%04u-%02u-%02u %02u:%02u:%02u.%03u pid=%lu tid=%lu %s\r\n",
        now.wYear,
        now.wMonth,
        now.wDay,
        now.wHour,
        now.wMinute,
        now.wSecond,
        now.wMilliseconds,
        GetCurrentProcessId(),
        GetCurrentThreadId(),
        message);
    if (characters > 0)
    {
        DWORD written = 0;
        WriteFile(file, line, static_cast<DWORD>(characters), &written, nullptr);
    }
    CloseHandle(file);
}
} // namespace

class CVisualTeXFormulaOleServerModule final
    : public ATL::CAtlExeModuleT<CVisualTeXFormulaOleServerModule>
{
public:
    DECLARE_REGISTRY_APPID_RESOURCEID(
        IDR_FORMULAOLESERVER,
        "{3C72FF7F-B04A-4FD0-AA7D-61D110D8B3C1}")
    DECLARE_LIBID(LIBID_VisualTeXFormulaOleLib)

    HRESULT PreMessageLoop(int showCommand) noexcept
    {
        TraceModule("PreMessageLoop enter");
        HoldStartupGraceLock();
        Gdiplus::GdiplusStartupInput startupInput;
        if (Gdiplus::GdiplusStartup(&gdiplusToken_, &startupInput, nullptr) != Gdiplus::Ok)
        {
            TraceModule("GdiplusStartup failed");
            ReleaseStartupGraceLockImmediately();
            return E_FAIL;
        }
        TraceModule("GdiplusStartup succeeded");
        const HRESULT result = __super::PreMessageLoop(showCommand);
        TraceModule(SUCCEEDED(result) ? "ATL PreMessageLoop succeeded" : "ATL PreMessageLoop failed");
        if (FAILED(result))
        {
            Gdiplus::GdiplusShutdown(gdiplusToken_);
            gdiplusToken_ = 0;
            ReleaseStartupGraceLockImmediately();
        }
        else
        {
            ArmStartupGraceRelease();
        }
        return result;
    }

    HRESULT PostMessageLoop() noexcept
    {
        TraceModule("PostMessageLoop enter");
        const HRESULT result = __super::PostMessageLoop();
        if (gdiplusToken_ != 0)
        {
            Gdiplus::GdiplusShutdown(gdiplusToken_);
            gdiplusToken_ = 0;
        }
        return result;
    }

private:
    static constexpr DWORD StartupGraceMilliseconds = 15000;

    void HoldStartupGraceLock() noexcept
    {
        if (InterlockedCompareExchange(&startupGraceHeld_, 1, 0) != 0)
            return;
        Lock();
        TraceModule("startup grace lock acquired");
    }

    void ArmStartupGraceRelease() noexcept
    {
        HANDLE thread = CreateThread(
            nullptr,
            0,
            &ReleaseStartupGraceProc,
            this,
            0,
            nullptr);
        if (thread == nullptr)
        {
            TraceModule("startup grace thread failed");
            ReleaseStartupGraceLockImmediately();
            return;
        }
        CloseHandle(thread);
        TraceModule("startup grace release armed");
    }

    void ReleaseStartupGraceLockImmediately() noexcept
    {
        if (InterlockedExchange(&startupGraceHeld_, 0) == 0)
            return;
        Unlock();
        TraceModule("startup grace lock released");
    }

    static DWORD WINAPI ReleaseStartupGraceProc(void* context) noexcept
    {
        Sleep(StartupGraceMilliseconds);
        auto* module = static_cast<CVisualTeXFormulaOleServerModule*>(context);
        module->ReleaseStartupGraceLockImmediately();
        return 0;
    }

    ULONG_PTR gdiplusToken_ = 0;
    volatile LONG startupGraceHeld_ = 0;
};

CVisualTeXFormulaOleServerModule _AtlModule;

extern "C" int WINAPI wWinMain(HINSTANCE, HINSTANCE, PWSTR, int showCommand)
{
    ConfigureProcessDpiAwareness();
    TraceModule("wWinMain enter");
    const int result = _AtlModule.WinMain(showCommand);
    TraceModule("wWinMain exit");
    return result;
}
