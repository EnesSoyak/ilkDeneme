#include <windows.h>
#include <string>
#include <vector>
#include <thread>
#include <chrono>
#include <iostream>
#include <unordered_map>

void typeTextFull(const std::wstring& text) {
    HKL layout = LoadKeyboardLayoutW(L"0000041F", KLF_ACTIVATE); // Türkçe Q klavye

    for (wchar_t ch : text) {
        // if (ch == L'i' || ch == L'I') {
        //     // Override Türkçe klavye hatası
        //     WORD vkCode = 0x49; // 'i' veya 'I' tuş kodu
        //     bool shift = (ch == L'I');

        //     std::vector<INPUT> inputs;

        //     if (shift) {
        //         INPUT shiftDown = {};
        //         shiftDown.type = INPUT_KEYBOARD;
        //         shiftDown.ki.wVk = VK_SHIFT;
        //         inputs.push_back(shiftDown);
        //     }

        //     INPUT keyDown = {};
        //     keyDown.type = INPUT_KEYBOARD;
        //     keyDown.ki.wVk = vkCode;
        //     inputs.push_back(keyDown);

        //     INPUT keyUp = keyDown;
        //     keyUp.ki.dwFlags = KEYEVENTF_KEYUP;
        //     inputs.push_back(keyUp);

        //     if (shift) {
        //         INPUT shiftUp = {};
        //         shiftUp.type = INPUT_KEYBOARD;
        //         shiftUp.ki.wVk = VK_SHIFT;
        //         shiftUp.ki.dwFlags = KEYEVENTF_KEYUP;
        //         inputs.push_back(shiftUp);
        //     }

        //     SendInput(inputs.size(), inputs.data(), sizeof(INPUT));
        //     std::this_thread::sleep_for(std::chrono::milliseconds(80));
        //     continue;
        // }

        SHORT vk = VkKeyScanExW(ch, layout);

        if (vk != -1) {
            BYTE vkCode = LOBYTE(vk);
            BYTE shift = HIBYTE(vk) & 1;

            std::vector<INPUT> inputs;

            if (shift) {
                INPUT shiftDown = {};
                shiftDown.type = INPUT_KEYBOARD;
                shiftDown.ki.wVk = VK_SHIFT;
                inputs.push_back(shiftDown);
            }

            INPUT keyDown = {};
            keyDown.type = INPUT_KEYBOARD;
            keyDown.ki.wVk = vkCode;
            inputs.push_back(keyDown);

            INPUT keyUp = keyDown;
            keyUp.ki.dwFlags = KEYEVENTF_KEYUP;
            inputs.push_back(keyUp);

            if (shift) {
                INPUT shiftUp = {};
                shiftUp.type = INPUT_KEYBOARD;
                shiftUp.ki.wVk = VK_SHIFT;
                shiftUp.ki.dwFlags = KEYEVENTF_KEYUP;
                inputs.push_back(shiftUp);
            }

            SendInput(inputs.size(), inputs.data(), sizeof(INPUT));
            std::this_thread::sleep_for(std::chrono::milliseconds(80));
            continue;
        }

        // Türkçe özel karakterler elle haritalanır
        WORD vkCode = 0;
        bool shift = false;

        switch (ch) {
            case L'ç': vkCode = 0xBA; shift = false; break;
            case L'Ç': vkCode = 0xBA; shift = true;  break;
            case L'ş': vkCode = 0xDE; shift = false; break;
            case L'Ş': vkCode = 0xDE; shift = true;  break;
            case L'ö': vkCode = 0xC0; shift = false; break;
            case L'Ö': vkCode = 0xC0; shift = true;  break;
            case L'ü': vkCode = 0xDC; shift = false; break;
            case L'Ü': vkCode = 0xDC; shift = true;  break;
            case L'ğ': vkCode = 0xBD; shift = false; break;
            case L'Ğ': vkCode = 0xBD; shift = true;  break;
            case L'ı': vkCode = 0xF0; shift = false; break;
            case L'i': vkCode = 0xBA; shift = false;  break;
            case L'I': vkCode = 0x49; shift = false; break;
            case L'İ': vkCode = 0x49; shift = true;  break;
            default: continue;
        }

        std::vector<INPUT> inputs;

        if (shift) {
            INPUT shiftDown = {};
            shiftDown.type = INPUT_KEYBOARD;
            shiftDown.ki.wVk = VK_SHIFT;
            inputs.push_back(shiftDown);
        }

        INPUT keyDown = {};
        keyDown.type = INPUT_KEYBOARD;
        keyDown.ki.wVk = vkCode;
        inputs.push_back(keyDown);

        INPUT keyUp = keyDown;
        keyUp.ki.dwFlags = KEYEVENTF_KEYUP;
        inputs.push_back(keyUp);

        if (shift) {
            INPUT shiftUp = {};
            shiftUp.type = INPUT_KEYBOARD;
            shiftUp.ki.wVk = VK_SHIFT;
            shiftUp.ki.dwFlags = KEYEVENTF_KEYUP;
            inputs.push_back(shiftUp);
        }

        SendInput(inputs.size(), inputs.data(), sizeof(INPUT));
        std::this_thread::sleep_for(std::chrono::milliseconds(80));
    }
}


int main() {
    ActivateKeyboardLayout(LoadKeyboardLayoutA("00000409", KLF_ACTIVATE), KLF_SETFORPROCESS);
    // 1. Excel'i başlat
    STARTUPINFO si = { sizeof(si) };
    PROCESS_INFORMATION pi;
    CreateProcess(NULL, (LPSTR)"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE",
                  NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi);

    // 2. Excel penceresi hazır olana kadar bekle
    HWND hwnd = NULL;
    while (!hwnd) {
        hwnd = FindWindow("XLMAIN", NULL);
        Sleep(200);
    }

    // 3. Pencereyi öne getir
    SetForegroundWindow(hwnd);
    Sleep(500);

    // 4. Pencere konumunu al ve hücreye tıklat
    RECT rect;
    GetWindowRect(hwnd, &rect);
    int x = rect.left + 250;
    int y = rect.top + 250;
    SetCursorPos(x, y);
    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
    Sleep(1000); 

   
    std::vector<std::wstring> values = {
        L"Ali", L"Veli", L"Çiğdem", L"İzmir", L"Şule", L"Ğazi", L"Ümit", L"Islak"
    };

    for (const auto& word : values) {
        typeTextFull(word);

        // Enter tuşu
        INPUT enter[2] = {};
        enter[0].type = INPUT_KEYBOARD;
        enter[0].ki.wVk = VK_RETURN;
        enter[1] = enter[0];
        enter[1].ki.dwFlags = KEYEVENTF_KEYUP;
        SendInput(2, enter, sizeof(INPUT));

        std::this_thread::sleep_for(std::chrono::milliseconds(150));
    }
    typeTextFull(L"i");


    return 0;
}
