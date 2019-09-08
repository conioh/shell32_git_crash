#include <filesystem>


#include <Windows.h>
#include <winerror.h>


#import <shell32.dll>


class COM_Initializer {
    HRESULT hr;

public:
    COM_Initializer(COINIT co_init) : hr(CoInitializeEx(nullptr, co_init)) {}
    ~COM_Initializer() {
        if (SUCCEEDED(hr)) {
            CoUninitialize();
        }
    }
};


void invoke_verb(const std::filesystem::path& shortcut_path) {
    COM_Initializer COM_init(COINIT_APARTMENTTHREADED);

    Shell32::IShellDispatchPtr shell_application("Shell.Application");

    auto folder_name = shortcut_path.parent_path().wstring();
    auto shell_folder = shell_application->NameSpace(folder_name.c_str());

    auto file_name = shortcut_path.filename().wstring();
    auto shell_item = shell_folder->ParseName(file_name.c_str());

    _bstr_t verb_name = "open";
    auto hr = shell_item->InvokeVerb(verb_name);
}


int main() {
    std::wstring target_path = LR"(C:\Windows\notepad.exe)";
    invoke_verb(target_path);
    
    Sleep(60000);

    return 0;
}
