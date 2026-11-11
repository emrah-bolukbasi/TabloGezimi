#include "tablo.h"

#define SINIR 1048576

int Tablo_ABI(void) { return 1; }

int Tablo_Satir(const TabloU16 *const *basliklar,
                const TabloU16 *const *hucreler, int sutun_sayisi,
                const int *secim, int secim_sayisi, int baslik_oku,
                TabloU16 *cikti, int kapasite)
{
    if (!hucreler || sutun_sayisi < 1 || sutun_sayisi > 4096 ||
        secim_sayisi < 0 || secim_sayisi > 4096 ||
        (secim_sayisi && !secim) || kapasite < 0 ||
        (!cikti && kapasite) || (baslik_oku < TABLO_BASLIKSIZ || baslik_oku > TABLO_BASLIK_KONUSMA))
        return TABLO_GECERSIZ;

    int adet = secim_sayisi ? secim_sayisi : sutun_sayisi;
    int uzunluk = 0;
    for (int i = 0; i < adet; ++i) {
        int sutun = secim_sayisi ? secim[i] : i;
        if (sutun < 0 || sutun >= sutun_sayisi)
            return TABLO_SUTUN_HATASI;
        const TabloU16 *baslik = baslik_oku && basliklar ? basliklar[sutun] : 0;
        const TabloU16 *hucre = hucreler[sutun];
        if (i) uzunluk += 2;
        if (baslik && *baslik) {
            uzunluk += baslik_oku == TABLO_BASLIK_IKI_NOKTA ? 2 : 1;
            while (*baslik++) {
                if (++uzunluk >= SINIR) return TABLO_COK_UZUN;
            }
        }
        if (hucre && *hucre) {
            while (*hucre++) {
                if (++uzunluk >= SINIR) return TABLO_COK_UZUN;
            }
        } else {
            uzunluk += 3;
        }
        if (uzunluk >= SINIR) return TABLO_COK_UZUN;
    }

    int gerekli = uzunluk + 1;
    if (!cikti || kapasite < gerekli) return gerekli;

    for (int i = 0; i < adet; ++i) {
        int sutun = secim_sayisi ? secim[i] : i;
        const TabloU16 *baslik = baslik_oku && basliklar ? basliklar[sutun] : 0;
        const TabloU16 *hucre = hucreler[sutun];
        if (i) {
            *cikti++ = ',';
            *cikti++ = ' ';
        }
        if (baslik && *baslik) {
            while (*baslik) *cikti++ = *baslik++;
            if (baslik_oku == TABLO_BASLIK_IKI_NOKTA) *cikti++ = ':';
            *cikti++ = ' ';
        }
        if (hucre && *hucre) {
            while (*hucre) *cikti++ = *hucre++;
        } else {
            *cikti++ = 'B';
            *cikti++ = 'o';
            *cikti++ = 0x015f;
        }
    }
    *cikti = 0;
    return gerekli;
}
