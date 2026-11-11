"""Masaüstü Microsoft Excel için, yalnızca etkin hücrede çalışan tablo katmanı.

Excel'in COM otomasyonu üzerinden CurrentRegion (bitişik veri bloğu) kullanılır.
Bu modül Excel Online veya LibreOffice için COM desteği iddia etmez.
"""

import api
import controlTypes
from comHelper import getActiveObject
from .i18n import _

MAX_SATIR = 100000
MAX_SUTUN = 256


class ExcelHatasi(Exception):
    pass


def excel_hucresinde_mi():
    """Yalnızca gerçek Excel hücresinde gezinme tuşlarını yakala.

    Şerit, formül çubuğu, düzenleme kutusu ve diyalogların ok tuşları serbest.
    """
    odak = api.getFocusObject()
    try:
        return (
            odak.appModule.appName.lower() == "excel"
            and odak.role in (controlTypes.Role.TABLECELL, controlTypes.Role.DATAITEM)
        )
    except (AttributeError, RuntimeError):
        return False


def temizle(deger):
    if deger is None:
        return ""
    return " ".join(str(deger).split())


class ExcelTablosu:
    def __init__(self):
        if not excel_hucresinde_mi():
            raise ExcelHatasi(_("Excel hücresinde değilsiniz"))
        odak = api.getFocusObject()
        try:
            self.excel = getActiveObject("Excel.Application", dynamic=True, appModule=odak.appModule)
        except (AttributeError, ValueError, OSError, RuntimeError):
            # Yardımcı bağlayıcı bulunmayan sürümlerde standart COM yolunu dene.
            self.excel = getActiveObject("Excel.Application", dynamic=True)
        try:
            self.sayfa = self.excel.ActiveSheet
            self.aktif = self.excel.ActiveCell
            self.aralik = self.aktif.CurrentRegion
            self.ilk_satir = int(self.aralik.Row)
            self.ilk_sutun = int(self.aralik.Column)
            self.satir_sayisi = int(self.aralik.Rows.Count)
            self.sutun_sayisi = int(self.aralik.Columns.Count)
            self.satir = int(self.aktif.Row)
            self.sutun = int(self.aktif.Column)
            self._baslik_onbellegi = {}
        except Exception as hata:
            raise ExcelHatasi(_("Excel çalışma sayfasına erişilemedi")) from hata
        if self.satir_sayisi < 1 or self.sutun_sayisi < 1:
            raise ExcelHatasi(_("Tablo bulunamadı"))
        # Tek başına boş bir hücre tablo değildir.
        if self.satir_sayisi == 1 and self.sutun_sayisi == 1 and not self.deger(self.satir, self.sutun):
            raise ExcelHatasi(_("Bu hücrede tablo bulunamadı"))

    @property
    def son_satir(self):
        return self.ilk_satir + self.satir_sayisi - 1

    @property
    def son_sutun(self):
        return self.ilk_sutun + self.sutun_sayisi - 1

    def deger(self, satir, sutun):
        # Text, sayının Excel'deki görünen biçimini (para birimi, tarih vb.) korur.
        return temizle(self.sayfa.Cells(satir, sutun).Text)

    def baslik(self, sutun):
        if sutun not in self._baslik_onbellegi:
            deger = self.deger(self.ilk_satir, sutun) if self.satir_sayisi > 1 else ""
            self._baslik_onbellegi[sutun] = deger or _("Sütun {num}").format(num=sutun - self.ilk_sutun + 1)
        return self._baslik_onbellegi[sutun]

    def hizli_metin_bloku(self, tur):
        """Tek COM çağrısı: yalnızca metin/boş hücrelerden oluşan alanlar.

        Numeric, tarih veya formül sonucu sayılar için görünür Excel .Text
        biçimi farklı olacağından None döndürülür. Büyük bloklarda COM'un
        tek seferde NVDA'yı uzun süre durdurmaması için güvenlik sınırı var.
        """
        alan_hucreleri = self.satir_sayisi * (self.sutun_sayisi if tur == "tablo" else 1)
        if alan_hucreleri > 50000 or alan_hucreleri < 2:
            return None
        try:
            if tur == "tablo":
                aralik = self.aralik
                sutun_sayisi = self.sutun_sayisi
            else:
                aralik = self.sayfa.Range(
                    self.sayfa.Cells(self.ilk_satir, self.sutun),
                    self.sayfa.Cells(self.son_satir, self.sutun),
                )
                sutun_sayisi = 1
            veri = aralik.Value2
            # Excel Range.Value2: çok hücreli alanlar 2-boyutlu tuple.
            if not isinstance(veri, (tuple, list)) or len(veri) != self.satir_sayisi:
                return None
            sonuc = []
            for satir in veri:
                if not isinstance(satir, (tuple, list)) or len(satir) != sutun_sayisi:
                    return None
                if any(deger is not None and not isinstance(deger, str) for deger in satir):
                    return None
                sonuc.append(tuple(temizle(deger) for deger in satir))
            return tuple(sonuc)
        except Exception:
            return None

    def hucre(self, satir=None, sutun=None):
        satir = self.satir if satir is None else satir
        sutun = self.sutun if sutun is None else sutun
        if satir == self.ilk_satir and self.satir_sayisi > 1:
            return "", self.deger(satir, sutun)
        return self.baslik(sutun), self.deger(satir, sutun)

    def satir_verisi(self, satir):
        if self.sutun_sayisi > MAX_SUTUN:
            raise ExcelHatasi(_("Bu veri alanında {count} sütundan fazlası var").format(count=MAX_SUTUN))
        basliklar, degerler = [], []
        for sutun in range(self.ilk_sutun, self.son_sutun + 1):
            baslik, deger = self.hucre(satir, sutun)
            basliklar.append(baslik)
            degerler.append(deger)
        return basliklar, degerler

    def ilerle(self, satir_farki, sutun_farki):
        yeni_satir, yeni_sutun = self.satir + satir_farki, self.sutun + sutun_farki
        if not self.ilk_satir <= yeni_satir <= self.son_satir or not self.ilk_sutun <= yeni_sutun <= self.son_sutun:
            raise ExcelHatasi(_("Tablo sınırı"))
        self.sayfa.Cells(yeni_satir, yeni_sutun).Select()
        self.satir, self.sutun = yeni_satir, yeni_sutun
        return self.satir_verisi(yeni_satir) if satir_farki else tuple([x] for x in self.hucre(yeni_satir, yeni_sutun))

    def sutun_degerleri(self):
        if self.satir_sayisi > MAX_SATIR:
            raise ExcelHatasi(_("Bu sütun tek seferde kopyalamak için çok büyük"))
        return [self.deger(r, self.sutun) or _("Boş") for r in range(self.ilk_satir, self.son_satir + 1)]

    def tablo_satirlari(self):
        if self.satir_sayisi > MAX_SATIR or self.sutun_sayisi > MAX_SUTUN:
            raise ExcelHatasi(_("Bu tablo tek seferde kopyalamak için çok büyük"))
        for r in range(self.ilk_satir, self.son_satir + 1):
            yield self.satir_verisi(r)
