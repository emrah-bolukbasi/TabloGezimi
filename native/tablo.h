#ifndef TABLO_H
#define TABLO_H
#ifdef __cplusplus
extern "C" {
#endif
typedef unsigned short TabloU16;
enum { TABLO_GECERSIZ = -1, TABLO_SUTUN_HATASI = -2, TABLO_COK_UZUN = -3 };
enum { TABLO_BASLIKSIZ = 0, TABLO_BASLIK_IKI_NOKTA = 1, TABLO_BASLIK_KONUSMA = 2 };
int Tablo_ABI(void);
int Tablo_Satir(const TabloU16 *const *basliklar,
                const TabloU16 *const *hucreler, int sutun_sayisi,
                const int *secim, int secim_sayisi, int baslik_oku,
                TabloU16 *cikti, int kapasite);
#ifdef __cplusplus
}
#endif
#endif
