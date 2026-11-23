#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <shobjidl.h>
#include <shlwapi.h>
#include <cwchar>
#include <new>

static HMODULE moduleHandle = nullptr;
static volatile LONG liveObjects = 0;

static const CLSID commandId = {
    0xa5f9e465, 0x48e4, 0x49ec,
    {0x98, 0x06, 0x54, 0x27, 0x02, 0xd6, 0x1f, 0x2c}
};

static PCWSTR ExtensionOf(PCWSTR name) {
    PCWSTR extension = wcsrchr(name, L'.');
    return extension ? extension : L"";
}

static bool IsSupported(PCWSTR extension) {
    return _wcsicmp(extension, registeredExtension) == 0;
}

static HRESULT GetSingleItem(IShellItemArray* selection, IShellItem** item) {
    *item = nullptr;
    if (!selection) return E_INVALIDARG;

    DWORD count = 0;
    HRESULT result = selection->GetCount(&count);
    if (FAILED(result)) return result;
    if (count != 1) return E_INVALIDARG;

    return selection->GetItemAt(0, item);
}

static HRESULT Launch(PCWSTR format, PCWSTR selectedFile) {
    wchar_t directory[32768];
    DWORD length = GetModuleFileNameW(moduleHandle, directory, ARRAYSIZE(directory));
    if (!length) return HRESULT_FROM_WIN32(GetLastError());
    if (length >= ARRAYSIZE(directory))
        return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);

    wchar_t* separator = wcsrchr(directory, L'\\');
    if (!separator) return E_UNEXPECTED;
    *separator = L'\0';

    wchar_t python[32768];
    wchar_t command[32768];

    if (swprintf_s(python, ARRAYSIZE(python),
                   L"%ls\\ProjectConvert.exe", directory) < 0)
        return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);

    if (swprintf_s(command, ARRAYSIZE(command),
                   L"\"%ls\" %ls \"%ls\"",
                   python, format, selectedFile) < 0)
        return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);

    STARTUPINFOW startup{};
    startup.cb = sizeof(startup);
    PROCESS_INFORMATION process{};

    if (!CreateProcessW(python, command, nullptr, nullptr, FALSE,
                        CREATE_NO_WINDOW, nullptr, directory, &startup, &process))
        return HRESULT_FROM_WIN32(GetLastError());

    CloseHandle(process.hThread);
    CloseHandle(process.hProcess);
    return S_OK;
}

class ConvertCommand final : public IExplorerCommand {
    volatile LONG references = 1;

public:
    int child;
    explicit ConvertCommand(int index = -1) : child(index) { InterlockedIncrement(&liveObjects); }
    ~ConvertCommand() { InterlockedDecrement(&liveObjects); }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID id, void** object) override {
        if (!object) return E_POINTER;
        *object = nullptr;
        if (id != IID_IUnknown && id != __uuidof(IExplorerCommand))
            return E_NOINTERFACE;

        *object = static_cast<IExplorerCommand*>(this);
        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() override {
        return InterlockedIncrement(&references);
    }

    ULONG STDMETHODCALLTYPE Release() override {
        LONG remaining = InterlockedDecrement(&references);
        if (!remaining) delete this;
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE GetTitle(IShellItemArray*, LPWSTR* title) override {
        return SHStrDupW(child < 0 ? L"AnyConvert" : choices[child].title, title);
    }

    HRESULT STDMETHODCALLTYPE GetIcon(IShellItemArray*, LPWSTR* icon) override {
        wchar_t location[32768];
        DWORD size = GetModuleFileNameW(moduleHandle, location, ARRAYSIZE(location));
        if (!size || size >= ARRAYSIZE(location)) return E_FAIL;
        wchar_t* slash = wcsrchr(location, L'\\');
        if (!slash) return E_FAIL;
        if (wcscpy_s(slash + 1, ARRAYSIZE(location) - (slash + 1 - location), L"AnyConvert.ico") != 0) return E_FAIL;
        return SHStrDupW(location, icon);
    }

    HRESULT STDMETHODCALLTYPE GetToolTip(IShellItemArray*, LPWSTR* tip) override {
        return SHStrDupW(L"Choose an output format", tip);
    }

    HRESULT STDMETHODCALLTYPE GetCanonicalName(GUID* name) override {
        *name = commandId;
        if (child >= 0) name->Data4[7] += static_cast<BYTE>(child + 1);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetState(IShellItemArray* selection, BOOL,
                                       EXPCMDSTATE* state) override {
        // File type matching is declared in the package manifest. Explorer
        // can query an individual command before resolving its selected item.
        *state = ECS_ENABLED;
        if (selection) {
            DWORD count = 0;
            if (SUCCEEDED(selection->GetCount(&count)) && count != 1)
                *state = ECS_HIDDEN;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Invoke(IShellItemArray* selection, IBindCtx*) override {
        IShellItem* item = nullptr;
        HRESULT result = GetSingleItem(selection, &item);
        if (FAILED(result)) return result;

        PWSTR path = nullptr;
        result = item->GetDisplayName(SIGDN_FILESYSPATH, &path);
        item->Release();

        if (SUCCEEDED(result)) {
            result = path && IsSupported(ExtensionOf(path))
                && child >= 0 ? Launch(choices[child].format, path) : E_INVALIDARG;
        }

        CoTaskMemFree(path);
        return result;
    }

    HRESULT STDMETHODCALLTYPE GetFlags(EXPCMDFLAGS* flags) override {
        *flags = child < 0 ? ECF_HASSUBCOMMANDS : ECF_DEFAULT;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE EnumSubCommands(IEnumExplorerCommand** commands) override;
};

class FormatEnumerator final : public IEnumExplorerCommand {
    volatile LONG refs = 1;
    ULONG cursor;
public:
    explicit FormatEnumerator(ULONG index = 0) : cursor(index) { InterlockedIncrement(&liveObjects); }
    ~FormatEnumerator() { InterlockedDecrement(&liveObjects); }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID id, void** out) override {
        if (!out) return E_POINTER;
        *out = nullptr;
        if (id != IID_IUnknown && id != __uuidof(IEnumExplorerCommand)) return E_NOINTERFACE;
        *out = static_cast<IEnumExplorerCommand*>(this);
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return InterlockedIncrement(&refs); }
    ULONG STDMETHODCALLTYPE Release() override {
        ULONG n = InterlockedDecrement(&refs);
        if (!n) delete this;
        return n;
    }
    HRESULT STDMETHODCALLTYPE Next(ULONG count, IExplorerCommand** out, ULONG* fetched) override {
        if (!out || (!fetched && count != 1)) return E_POINTER;
        ULONG n = 0;
        if (fetched) *fetched = 0;
        while (n < count && cursor < ARRAYSIZE(choices)) {
            out[n] = new (std::nothrow) ConvertCommand(cursor);
            if (!out[n]) {
                cursor -= n;
                while (n) out[--n]->Release();
                return E_OUTOFMEMORY;
            }
            ++n;
            ++cursor;
        }
        if (fetched) *fetched = n;
        return n == count ? S_OK : S_FALSE;
    }
    HRESULT STDMETHODCALLTYPE Skip(ULONG count) override {
        ULONG remaining = ARRAYSIZE(choices) - cursor;
        if (count > remaining) { cursor += remaining; return S_FALSE; }
        cursor += count;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Reset() override { cursor = 0; return S_OK; }
    HRESULT STDMETHODCALLTYPE Clone(IEnumExplorerCommand** out) override {
        if (!out) return E_POINTER;
        *out = new (std::nothrow) FormatEnumerator(cursor);
        return *out ? S_OK : E_OUTOFMEMORY;
    }
};

HRESULT ConvertCommand::EnumSubCommands(IEnumExplorerCommand** out) {
    if (!out) return E_POINTER;
    *out = nullptr;
    if (child >= 0) return E_NOTIMPL;
    *out = new (std::nothrow) FormatEnumerator;
    return *out ? S_OK : E_OUTOFMEMORY;
}

class CommandFactory final : public IClassFactory {
    volatile LONG references = 1;

public:
    CommandFactory() { InterlockedIncrement(&liveObjects); }
    ~CommandFactory() { InterlockedDecrement(&liveObjects); }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID id, void** object) override {
        if (!object) return E_POINTER;
        *object = nullptr;
        if (id != IID_IUnknown && id != IID_IClassFactory) return E_NOINTERFACE;

        *object = static_cast<IClassFactory*>(this);
        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() override {
        return InterlockedIncrement(&references);
    }

    ULONG STDMETHODCALLTYPE Release() override {
        LONG remaining = InterlockedDecrement(&references);
        if (!remaining) delete this;
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* outer, REFIID id,
                                             void** object) override {
        if (!object) return E_POINTER;
        *object = nullptr;
        if (outer) return CLASS_E_NOAGGREGATION;

        auto* command = new (std::nothrow) ConvertCommand();
        if (!command) return E_OUTOFMEMORY;

        HRESULT result = command->QueryInterface(id, object);
        command->Release();
        return result;
    }

    HRESULT STDMETHODCALLTYPE LockServer(BOOL lock) override {
        if (lock) InterlockedIncrement(&liveObjects);
        else InterlockedDecrement(&liveObjects);
        return S_OK;
    }
};

#pragma comment(linker, "/EXPORT:DllGetClassObject")
STDAPI DllGetClassObject(REFCLSID id, REFIID interfaceId, void** object) {
    if (!object) return E_POINTER;
    *object = nullptr;
    if (id != commandId) return CLASS_E_CLASSNOTAVAILABLE;

    auto* factory = new (std::nothrow) CommandFactory();
    if (!factory) return E_OUTOFMEMORY;

    HRESULT result = factory->QueryInterface(interfaceId, object);
    factory->Release();
    return result;
}

#pragma comment(linker, "/EXPORT:DllCanUnloadNow")
STDAPI DllCanUnloadNow() {
    return InterlockedCompareExchange(&liveObjects, 0, 0) == 0 ? S_OK : S_FALSE;
}

BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) moduleHandle = instance;
    return TRUE;
}
