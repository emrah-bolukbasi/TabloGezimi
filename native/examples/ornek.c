#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <stdio.h>
#include "kaynak/tablo.h"

typedef int (__cdecl *SatirIslevi)(const TabloU16 *const *,
    const TabloU16 *const *, int, const int *, int, int, TabloU16 *, int);

int main(void)
{
#ifdef _WIN64
    const wchar_t *yol = L".\\x64\\TabloOkuma.dll";
#else
    const wchar_t *yol = L".\\x86\\TabloOkuma.dll";
#endif
    HMODULE dll = LoadLibraryW(yol);
    if (!dll) return 1;

    SatirIslevi satir = (SatirIslevi)GetProcAddress(dll, "Tablo_Satir");
    const TabloU16 *basliklar[] = {
        (const TabloU16 *)L"\x0130sim", (const TabloU16 *)L"Numara",
        (const TabloU16 *)L"S\x0131ra"
    };
    const TabloU16 *hucreler[] = {
        (const TabloU16 *)L"Emrah", (const TabloU16 *)L"2121",
        (const TabloU16 *)L"23"
    };
    TabloU16 cikti[128];
    int sonuc = satir ? satir(basliklar, hucreler, 3, NULL, 0,
                               TABLO_BASLIK_KONUSMA, cikti, 128) : 0;
    if (sonuc > 0 && sonuc <= 128) {
        DWORD yazilan;
        WriteConsoleW(GetStdHandle(STD_OUTPUT_HANDLE), cikti, sonuc - 1, &yazilan, NULL);
        puts("");
    }
    FreeLibrary(dll);
    return sonuc > 0 && sonuc <= 128 ? 0 : 2;
}
