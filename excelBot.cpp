#include <windows.h>
#include <string>
#include <vector>
#include <thread>
#include <chrono>
// Enter’a bas
void pressEnter() {
    INPUT input[2] = {};
    input[0].type = INPUT_KEYBOARD;
    input[0].ki.wVk = VK_RETURN;
    input[1] = input[0];
    input[1].ki.dwFlags = KEYEVENTF_KEYUP;

    SendInput(2, input, sizeof(INPUT));
}


void typeTextWithShift(const std::string& text) {
    for (char ch : text) {
        SHORT vk = VkKeyScanA(ch);
        if (vk == -1) continue;

        bool shift = HIBYTE(vk) & 1;

        INPUT inputs[4] = {};
        int count = 0;

        if (shift) {
            inputs[count].type = INPUT_KEYBOARD;
            inputs[count].ki.wVk = VK_SHIFT;
            count++;
        }

        inputs[count].type = INPUT_KEYBOARD;
        inputs[count].ki.wVk = LOBYTE(vk);
        count++;

        inputs[count] = inputs[count - 1];
        inputs[count].ki.dwFlags = KEYEVENTF_KEYUP;
        count++;

        if (shift) {
            inputs[count].type = INPUT_KEYBOARD;
            inputs[count].ki.wVk = VK_SHIFT;
            inputs[count].ki.dwFlags = KEYEVENTF_KEYUP;
            count++;
        }

        SendInput(count, inputs, sizeof(INPUT));
        std::this_thread::sleep_for(std::chrono::milliseconds(80));
    }
}
int main() {
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
    Sleep(1000); // Hücre odaklansın

    // 5. Yazılacak veriler
    std::vector<std::string> values = {"Ali", "Veli", "Ayse", "Fatma"};

    for (const auto& word : values) {
        typeTextWithShift(word);

        // Enter tuşu
        INPUT enter[2] = {};
        enter[0].type = INPUT_KEYBOARD;
        enter[0].ki.wVk = VK_RETURN;
        enter[1] = enter[0];
        enter[1].ki.dwFlags = KEYEVENTF_KEYUP;
        SendInput(2, enter, sizeof(INPUT));

        std::this_thread::sleep_for(std::chrono::milliseconds(150));
    }

    return 0;
}
