import platform
import threading
import time

import api
import controlTypes
import core
import globalPluginHandler
import inputCore
import queueHandler
import speech
import textInfos
import ui
from documentBase import DocumentWithTableNavigation, _Axis, _Movement
from logHandler import log
from scriptHandler import script

from .i18n import _
from .dll_arayuzu import satir_oku, yukle
from .excel_destegi import ExcelTablosu, ExcelHatasi, excel_hucresinde_mi


HAREKETLER = {
    "kb:downarrow": (_Movement.NEXT, _Axis.ROW),
    "kb:uparrow": (_Movement.PREVIOUS, _Axis.ROW),
    "kb:rightarrow": (_Movement.NEXT, _Axis.COLUMN),
    "kb:leftarrow": (_Movement.PREVIOUS, _Axis.COLUMN),
}


def hucre_bilgisi(info, konum):
    for alan in reversed(info.getTextWithFields()):
        if not isinstance(alan, textInfos.FieldCommand) or alan.command != "controlStart":
            continue
        veri = alan.field
        if (
            veri.get("table-id") == konum.tableID
            and veri.get("table-rownumber") == konum.row
            and veri.get("table-columnnumber") == konum.col
        ):
            return veri
    return {}




def _satirlari_birlestir(satirlar):
    """İş parçacığında yalnızca saf Python metinlerini birleştir."""
    return "\r\n".join(satirlar)


def _hizli_excel_satirlari(tur, blok, basliklar, ilk_satir):
    """COM ve NVDA nesnesi içermeyen hazır metin bloğunu biçimlendir."""
    if tur == "sutun":
        return [satir[0] or _("Boş") for satir in blok]
    cikti = []
    bos = _("Boş")
    for index, satir in enumerate(blok):
        baslik_var = not (ilk_satir and index == 0)
        cikti.append(", ".join(
            (f"{baslik} {deger or _('Boş')}" if baslik_var and baslik else deger or _("Boş"))
            for baslik, deger in zip(basliklar, satir)
        ))
    return cikti

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.acik = False
        self.dll = None
        # En son başlayan gezinmenin sıra numarası.
        self.gezinme_numarasi = 0
        self.kopyalama_numarasi = 0
        self.bekleyen_excel_okuma = None
        self.excel_duyuru_planlandi = False
        self._excel_tuslar_bagli = False
        self._sonlandirildi = False
        if platform.machine().lower() in {"amd64", "x86", "i386", "i686", "x86_64"}:
            try:
                self.dll = yukle()
            except (OSError, RuntimeError):
                log.warning("Tablo Gezimi DLL yüklenemedi.", exc_info=True)
        inputCore.decide_executeGesture.register(self.ok_tusunu_yakala)

    def terminate(self):
        self._sonlandirildi = True
        self.acik = False
        self.gezinme_numarasi += 1
        self.kopyalama_numarasi += 1
        self.bekleyen_excel_okuma = None
        self.excel_duyuru_planlandi = False
        self._excel_tuslarini_ayarla(False)
        inputCore.decide_executeGesture.unregister(self.ok_tusunu_yakala)
        super().terminate()

    @script(
        description=_("Tablo gezimini açar veya kapatır"),
        category=_("Tablo Gezimi"),
        gesture="kb:alt+shift+o",
    )
    def script_tabloGeziminiDegistir(self, gesture):
        self.acik = not self.acik
        self.gezinme_numarasi += 1
        self.kopyalama_numarasi += 1
        self.bekleyen_excel_okuma = None
        self.excel_duyuru_planlandi = False
        speech.cancelSpeech()
        self._excel_tuslarini_ayarla(self.acik and excel_hucresinde_mi())
        ui.message(_("Tablo gezimi açıldı") if self.acik else _("Tablo gezimi kapatıldı"))

    def _excel_tuslarini_ayarla(self, bagla):
        """Yalnızca Excel hücresi odaktayken okları NVDA komutuna bağla.

        Eski decide_executeGesture kancası yalnızca NVDA'nın zaten işlediği
        tuşları görebiliyordu. Excel'deki çıplak oklar sıkça doğrudan Windows'a
        gönderildiği için satır okuma işlevi hiç çağrılmıyordu.
        """
        if bool(bagla) == self._excel_tuslar_bagli:
            return
        if bagla:
            self.bindGesture("kb:downArrow", "excelAsagi")
            self.bindGesture("kb:upArrow", "excelYukari")
            self.bindGesture("kb:rightArrow", "excelSaga")
            self.bindGesture("kb:leftArrow", "excelSola")
            self._excel_tuslar_bagli = True
        else:
            # clearGestureBindings() tüm kısayolları (Alt+Shift+O dahil)
            # siliyordu. Yalnızca Excel için dinamik eklenen okları kaldır.
            for tus in ("kb:downArrow", "kb:upArrow", "kb:rightArrow", "kb:leftArrow"):
                try:
                    self.removeGestureBinding(tus)
                except LookupError:
                    pass
            self._excel_tuslar_bagli = False

    def _excel_ok_komutu(self, gesture, hareket, eksen):
        if not self.acik or not excel_hucresinde_mi():
            self._excel_tuslarini_ayarla(False)
            gesture.send()
            return
        self.gezinme_numarasi += 1
        numara = self.gezinme_numarasi
        speech.cancelSpeech()
        # Satır okuma, ExcelTablosu kaynağıyla
        # aynıdır; yalnızca bir sonraki/önceki satır seçilir.
        self.excel_gezin(hareket, eksen, numara)

    def script_excelAsagi(self, gesture):
        self._excel_ok_komutu(gesture, _Movement.NEXT, _Axis.ROW)

    def script_excelYukari(self, gesture):
        self._excel_ok_komutu(gesture, _Movement.PREVIOUS, _Axis.ROW)

    def script_excelSaga(self, gesture):
        self._excel_ok_komutu(gesture, _Movement.NEXT, _Axis.COLUMN)

    def script_excelSola(self, gesture):
        self._excel_ok_komutu(gesture, _Movement.PREVIOUS, _Axis.COLUMN)

    def aktif_belge(self):
        """NVDA'nın kendisinin tablo gezinmesi yapabildiği belgeyi bul.

        PDF okuyucularında tablo navigasyonu kimi zaman bir sanal belge
        (treeInterceptor), kimi zaman doğrudan odaktaki NVDAObject üzerinde
        sağlanır. İkincisinde ``isReady`` / ``passThrough`` özellikleri
        bulunmayabilir. Bu özellikleri zorunlu tutmak PDF'yi sessizce
        reddediyordu. Doğrulama için NVDA'nın kendi hücre koordinatları
        yöntemini kullanıyoruz; HTML/Word davranışı değişmiyor.
        """
        odak = api.getFocusObject()
        adaylar = []
        gorulen = set()

        def ekle(nesne):
            if nesne is not None and id(nesne) not in gorulen:
                gorulen.add(id(nesne))
                adaylar.append(nesne)

        # Birincil yol: web, HTML ve PDF sanal belgeleri.
        ekle(getattr(odak, "treeInterceptor", None))
        # PDF / Word UIA okuyucuları: tablo doğrudan odak nesnesindedir.
        ekle(odak)
        # Bazı PDF görüntüleyicilerinde caret ile focus ayrı nesnelerdir.
        try:
            caret = api.getCaretObject()
        except (AttributeError, RuntimeError):
            caret = None
        if caret is not None and caret is not odak:
            ekle(getattr(caret, "treeInterceptor", None))
            ekle(caret)

        for belge in adaylar:
            if not isinstance(belge, DocumentWithTableNavigation):
                continue
            try:
                # isReady sanal belgede vardır, doğrudan NVDAObject'te
                # bulunması şart değildir.
                if getattr(belge, "isReady", True) is False:
                    continue
                if getattr(belge, "passThrough", False):
                    continue
                belge._getTableCellCoords(belge.selection)
                return belge
            except (LookupError, NotImplementedError, AttributeError, RuntimeError):
                continue
        return None

    def ok_tusunu_yakala(self, gesture):
        # Tuş eşlemesi yapılmadan önce çalışan NVDA giriş kancası.
        # Excel için dinamik eklenen ok kısayollarının tarayıcı ve PDF'ye
        # taşınmasını engelle. Tarayıcı için orijinal çalışan yöntem kullanılır.
        if not excel_hucresinde_mi() and self._excel_tuslar_bagli:
            self._excel_tuslarini_ayarla(False)
        if not self.acik or inputCore.manager.isInputHelpActive:
            return True
        hareket = next(
            (HAREKETLER[k] for k in gesture.normalizedIdentifiers if k in HAREKETLER),
            None,
        )
        if hareket is None:
            return True
        # Excel'in ok tuşları kendi doğrudan bağlı scriptleriyle çalışır.
        if excel_hucresinde_mi():
            if not self._excel_tuslar_bagli:
                self._excel_tuslarini_ayarla(True)
            return True
        try:
            belge = self.aktif_belge()
        except Exception:
            log.debugWarning("Tablo Gezimi belge konumunu inceleyemedi.", exc_info=True)
            return True
        if belge is None:
            return True
        # Burada _tableFindNewCell ÇAĞIRMA: bu kod giriş iş parçacığında
        # çalışır ve sanal belgenin seçim verisini kararsız hale getirebilir.
        # Asıl gezinme NVDA olay kuyruğunda tek kez yürütülür.
        self.gezinme_numarasi += 1
        numara = self.gezinme_numarasi
        queueHandler.queueFunction(
            queueHandler.eventQueue, speech.cancelSpeech, _immediate=True,
        )
        queueHandler.queueFunction(
            queueHandler.eventQueue, self.gezin, belge, *hareket, numara,
            _immediate=True,
        )
        return False

    def excel_gezin(self, hareket, eksen, numara):
        if not self.acik or numara != self.gezinme_numarasi or not excel_hucresinde_mi():
            return
        try:
            excel = ExcelTablosu()
            fark = 1 if hareket == _Movement.NEXT else -1
            # Metni hedef hücreye geçmeden önce hazırla. Excel'in seçim
            # değişikliği, NVDA'nın odak nesnesini geçici olarak değiştirebilir.
            yeni_satir = excel.satir + (fark if eksen == _Axis.ROW else 0)
            yeni_sutun = excel.sutun + (fark if eksen == _Axis.COLUMN else 0)
            if not excel.ilk_satir <= yeni_satir <= excel.son_satir or not excel.ilk_sutun <= yeni_sutun <= excel.son_sutun:
                raise ExcelHatasi(_("Tablo sınırı"))
            if eksen == _Axis.ROW:
                basliklar, degerler = excel.satir_verisi(yeni_satir)
            else:
                baslik, deger = excel.hucre(yeni_satir, yeni_sutun)
                basliklar, degerler = [baslik], [deger]
            metin = self.metin_hazirla(basliklar, degerler)
            # Seçim değişikliğinin NVDA'ya yaptığı tek hücre duyurusu,
            # hazırlanan tam satırın konuşmasını iptal etmesin.
            self.excel_duyuru_planlandi = False
            self.bekleyen_excel_okuma = (numara, metin)
            excel.sayfa.Cells(yeni_satir, yeni_sutun).Select()
            # Excel'de odak / seçim duyuruları farklı NVDA sürümlerinde
            # farklı olaylarla gelir. Bu nedenle son duyuru planı olaydan
            # bağımsız olarak da çalışır.
            # Sabit 450 ms bekleme hızlı gezinmeyi duraksatıyordu.
            # Seçim olayının normal yolunu tercih et; olay gelmezse kısa
            # süreli yedek duyuru devreye girsin.
            core.callLater(115, self.excel_gezinmeyi_duyur, numara)
            log.debug("Tablo Gezimi Excel satır hazırlandı: satır=%s, sütun=%s, uzunluk=%s",
                      yeni_satir, yeni_sutun, len(metin))
        except ExcelHatasi as hata:
            self.bekleyen_excel_okuma = None
            if self.acik and numara == self.gezinme_numarasi:
                ui.message(str(hata))
        except Exception:
            self.bekleyen_excel_okuma = None
            log.exception("Tablo Gezimi Excel hücresinde gezinemedi.")
            if self.acik and numara == self.gezinme_numarasi:
                ui.message(_("Excel hücresi okunamadı"))

    def excel_gezinmeyi_duyur(self, numara):
        """Excel seçim olayı tamamlandıktan sonra satırı tek mesaj olarak duyur.

        Seçimden hemen sonraki odak rolünü kontrol etmiyoruz: NVDA, Excel
        seçimini işlerken odak nesnesini geçici olarak başka role geçirebilir.
        Bu, önceki sürümün tüm satır duyurusunu sessizce atlamasına neden
        olabiliyordu. Yalnızca Excel uygulamasının halen odakta olduğunu doğrula.
        """
        bekleyen = self.bekleyen_excel_okuma
        if (not self.acik or numara != self.gezinme_numarasi
                or bekleyen is None or bekleyen[0] != numara):
            return
        # Başka uygulamaya geçildiyse eski Excel verisini duyurma.
        try:
            odak = api.getFocusObject()
            if getattr(getattr(odak, "appModule", None), "appName", "").lower() != "excel":
                self.bekleyen_excel_okuma = None
                return
        except (AttributeError, RuntimeError):
            self.bekleyen_excel_okuma = None
            return
        self.bekleyen_excel_okuma = None
        speech.cancelSpeech()
        ui.message(bekleyen[1])

    def event_gainFocus(self, obj, nextHandler):
        # Excel hücresine girince okları yakala, hücreden veya uygulamadan
        # çıkıldığında tüm okları normal işleyişlerine geri ver.
        try:
            excel_mi = (
                self.acik
                and getattr(getattr(obj, "appModule", None), "appName", "").lower() == "excel"
                and obj.role in (controlTypes.Role.TABLECELL, controlTypes.Role.DATAITEM)
            )
        except (AttributeError, RuntimeError):
            excel_mi = False
        self._excel_tuslarini_ayarla(excel_mi)
        # Excel'in standart odak / hücre duyurusu önce işlenir.
        # Ardından bekleyen satır konuşması üstünlüğü alır.
        nextHandler()
        self.excel_secim_olayi(obj)

    def event_nameChange(self, obj, nextHandler):
        nextHandler()
        self.excel_secim_olayi(obj)

    def excel_secim_olayi(self, obj):
        bekleyen = self.bekleyen_excel_okuma
        if not self.acik or bekleyen is None:
            return
        if getattr(getattr(obj, "appModule", None), "appName", "").lower() != "excel":
            return
        if self.excel_duyuru_planlandi:
            return
        self.excel_duyuru_planlandi = True
        core.callLater(55, self.excel_secim_olayi_duyur, bekleyen[0])

    def excel_secim_olayi_duyur(self, numara):
        self.excel_duyuru_planlandi = False
        self.excel_gezinmeyi_duyur(numara)

    def _kopyalamayi_arka_planda_tamamla(self, satirlar, numara, aciklama, hazir=False):
        """Arka planda yalnızca normal Python dizeleri işle; pano NVDA kuyruğunda."""
        # Diziyi dondur: işçi başlarken NVDA iş parçacığı da değiştirmesin.
        sabit = tuple(satirlar)

        def isci():
            try:
                if hazir:
                    metin = sabit[0]
                else:
                    metin = _satirlari_birlestir(sabit)
            except Exception:
                log.exception("Tablo Gezimi arka plan metin birleştirme hatası.")
                return
            queueHandler.queueFunction(
                queueHandler.eventQueue, self._kopyalamayi_bitir,
                metin, numara, aciklama,
            )

        threading.Thread(target=isci, name="TabloGezimiMetin", daemon=True).start()

    def _kopyalamayi_bitir(self, metin, numara, aciklama):
        if self._sonlandirildi or not self.acik or self.kopyalama_numarasi != numara:
            return
        try:
            if api.copyToClip(metin):
                ui.message(aciklama)
            else:
                ui.message(_("Kopyalanamadı"))
        except Exception:
            log.exception("Tablo Gezimi panoya yazamadı.")
            ui.message(_("Kopyalanamadı"))

    def excel_kopyala(self, tur):
        try:
            excel = ExcelTablosu()
            if tur == "hucre":
                metin = excel.deger(excel.satir, excel.sutun)
            elif tur == "satir":
                metin = self.metin_hazirla(*excel.satir_verisi(excel.satir))
            else:
                # Büyük Excel alanları kopyalanırken NVDA olay kuyruğuna
                # küçük parçalar halinde iş bırakılır.
                self.kopyalama_numarasi += 1
                numara = self.kopyalama_numarasi
                if tur == "tablo":
                    if excel.satir_sayisi > 100000 or excel.sutun_sayisi > 256:
                        raise ExcelHatasi(_("Bu tablo tek seferde kopyalamak için çok büyük"))
                    baslangic, bitis = excel.ilk_satir, excel.son_satir
                else:
                    if excel.satir_sayisi > 100000:
                        raise ExcelHatasi(_("Bu sütun tek seferde kopyalamak için çok büyük"))
                    baslangic, bitis = excel.ilk_satir, excel.son_satir
                # Metin ağırlıklı Excel tablolarını tek COM isteğiyle al.
                # Sayı/tarih biçimi varsa eski .Text yolunu kullan ve
                # Excel'deki görünür biçimi kaybetme.
                blok = excel.hizli_metin_bloku(tur)
                if blok is not None:
                    basliklar = (
                        [excel.baslik(s) for s in range(excel.ilk_sutun, excel.son_sutun + 1)]
                        if tur == "tablo" else []
                    )
                    ilk_satir = tur == "tablo" and excel.satir_sayisi > 1
                    def hizli_isci():
                        try:
                            satirlar = _hizli_excel_satirlari(tur, blok, basliklar, ilk_satir)
                            metin = _satirlari_birlestir(satirlar)
                        except Exception:
                            log.exception("Tablo Gezimi hızlı Excel biçimlendirmesi başarısız.")
                            return
                        queueHandler.queueFunction(
                            queueHandler.eventQueue, self._kopyalamayi_bitir,
                            metin, numara, _("{count} satır panoya kopyalandı").format(count=len(blok)),
                        )
                    threading.Thread(target=hizli_isci, name="TabloGezimiExcelMetin", daemon=True).start()
                    return
                self.excel_kopyalama_parcasi(excel, tur, baslangic, bitis, [], numara)
                return
            if not metin:
                ui.message(_("Bulunulan hücre boş"))
            else:
                ui.message(_("Panoya kopyalandı") if api.copyToClip(metin) else _("Kopyalanamadı"))
        except ExcelHatasi as hata:
            ui.message(str(hata))
        except Exception:
            log.exception("Tablo Gezimi Excel kopyalama başarısız.")
            ui.message(_("Excel verileri kopyalanamadı"))

    def excel_kopyalama_parcasi(self, excel, tur, baslangic, bitis, satirlar, numara):
        if not self.acik or self.kopyalama_numarasi != numara or not excel_hucresinde_mi():
            return
        try:
            en_fazla = min(baslangic + 40, bitis + 1)
            bitis_zamani = time.perf_counter() + 0.014
            son = baslangic
            for satir in range(baslangic, en_fazla):
                if tur == "tablo":
                    satirlar.append(self.metin_hazirla(*excel.satir_verisi(satir)))
                else:
                    satirlar.append(excel.deger(satir, excel.sutun) or _("Boş"))
                son = satir + 1
                if time.perf_counter() >= bitis_zamani:
                    break
            if son <= bitis:
                queueHandler.queueFunction(
                    queueHandler.eventQueue, self.excel_kopyalama_parcasi,
                    excel, tur, son, bitis, satirlar, numara,
                )
            else:
                self._kopyalamayi_arka_planda_tamamla(
                    satirlar, numara, _("{count} satır panoya kopyalandı").format(count=len(satirlar)),
                )
        except Exception:
            log.exception("Tablo Gezimi Excel kopyalama sırasında hata.")
            ui.message(_("Excel verileri kopyalanamadı"))

    def sutun_basligi(self, belge, info, konum, onbellek=None):
        if onbellek is not None and konum.col in onbellek:
            return onbellek[konum.col]
        baslik = hucre_bilgisi(info, konum).get("table-columnheadertext")
        if baslik and baslik.strip():
            sonuc = " ".join(baslik.split())
            if onbellek is not None:
                onbellek[konum.col] = sonuc
            return sonuc
        try:
            ilk = belge._getTableCellAt(konum.tableID, info, 1, konum.col)
            ilk_konum = belge._getTableCellCoords(ilk)
        except LookupError:
            sonuc = _("Sütun {num}").format(num=konum.col)
            if onbellek is not None:
                onbellek[konum.col] = sonuc
            return sonuc
        if hucre_bilgisi(ilk, ilk_konum).get("role") == controlTypes.Role.TABLECOLUMNHEADER:
            baslik = " ".join(ilk.text.split())
            if baslik:
                if onbellek is not None:
                    onbellek[konum.col] = baslik
                return baslik
        sonuc = _("Sütun {num}").format(num=konum.col)
        if onbellek is not None:
            onbellek[konum.col] = sonuc
        return sonuc

    def hucre_verisi(self, belge, info, konum, onbellek=None):
        deger = " ".join(info.text.split())
        rol = hucre_bilgisi(info, konum).get("role")
        if rol in (controlTypes.Role.TABLECOLUMNHEADER, controlTypes.Role.TABLEROWHEADER):
            return "", deger
        return self.sutun_basligi(belge, info, konum, onbellek), deger

    def gezin(self, belge, hareket, eksen, numara):
        try:
            if not self.acik or self.aktif_belge() is not belge:
                return
            try:
                konum, info, secim = belge._tableFindNewCell(hareket, eksen, raiseOnEdge=True)
            except LookupError:
                # Son/ilk satırdan sonra belge imlecini tablo DIŞINA taşı.
                # Eklenti açık kalır; normal belge gezintisi devam eder.
                self.belgede_tablodan_cik(belge, hareket, eksen)
                return
            if eksen == _Axis.ROW:
                basliklar, degerler = self.satir_verisi(belge, info, konum, secim.trueRow)
            else:
                baslik, deger = self.hucre_verisi(belge, info, konum)
                basliklar, degerler = [baslik], [deger]
            metin = self.metin_hazirla(basliklar, degerler)
            hedef = info.copy()
            hedef.collapse()
            belge.selection = hedef
            belge._lastTableSelection = secim
            # Kullanıcı bu arada başka bir ok tuşuna bastıysa eski satırın
            # konuşmasını başlatma. Seçim değişimi ise sırasıyla korunur.
            if self.acik and numara == self.gezinme_numarasi:
                speech.cancelSpeech()
                ui.message(metin)
        except Exception:
            log.exception("Tablo Gezimi hücreyi okuyamadı.")
            if self.acik and numara == self.gezinme_numarasi:
                ui.message(_("Tablo hücresi okunamadı"))

    def belgede_tablodan_cik(self, belge, hareket, eksen):
        """Tablo sınırında sanal belge imlecini normal metne geçir.

        Satır sınırında satır birimi, sütun sınırında karakter birimi
        kullanılır. Yalnızca mevcut tablonun dışına çıkana kadar ilerler;
        böylece son hücrede kilitli kalınmaz.
        """
        try:
            konum = belge._getTableCellCoords(belge.selection)
            info = belge.selection.copy()
            info.collapse()
            birim = textInfos.UNIT_LINE if eksen == _Axis.ROW else textInfos.UNIT_CHARACTER
            yon = 1 if hareket == _Movement.NEXT else -1
            # Bir hücre birkaç görsel satıra yayılabilir; yine de sonsuz
            # tarama yapmamak için hareket sayısını sınırla.
            for _ in range(256):
                if not info.move(birim, yon):
                    break
                try:
                    yeni = belge._getTableCellCoords(info)
                except LookupError:
                    yeni = None
                if yeni is None or yeni.tableID != konum.tableID:
                    belge.selection = info
                    belge._lastTableSelection = None
                    # NVDA'nın normal belge konuşma yöntemini kullan.
                    duyuru = info.copy()
                    duyuru.expand(textInfos.UNIT_LINE)
                    speech.speakTextInfo(duyuru, reason=controlTypes.OutputReason.CARET)
                    return
            # Belgenin kendi sınırına gelindiyse kullanıcıya geri bildirim ver.
            ui.message(_("Belge sınırı"))
        except Exception:
            log.debugWarning("Tablo Gezimi tablodan çıkışı tamamlayamadı.", exc_info=True)
            ui.message(_("Tablo sınırı"))

    def mevcut_satir(self, belge):
        """Son gezinilen satırı, sütun başlıkları ve hücreleriyle döndür."""
        info = belge.selection
        konum = belge._getTableCellCoords(info)
        secim = getattr(belge, "_lastTableSelection", None)
        satir = getattr(secim, "trueRow", None)
        if (
            satir is None
            or konum.row != getattr(secim, "lastRow", konum.row)
        ):
            satir = konum.row
        basliklar, degerler = self.satir_verisi(belge, info, konum, satir)
        return self.metin_hazirla(basliklar, degerler)

    @script(
        description=_("Bulunulan tablo satırını panoya kopyalar (yalnızca tablo modunda)"),
        category=_("Tablo Gezimi"),
        gesture="kb:alt+shift+uparrow",
    )
    def script_satirKopyala(self, gesture):
        if not self.acik:
            gesture.send()
            return
        if excel_hucresinde_mi():
            self.excel_kopyala("satir")
            return
        try:
            belge = self.aktif_belge()
            if belge is None:
                ui.message(_("Tablo içinde değilsiniz"))
                return
            metin = self.mevcut_satir(belge)
            ui.message(_("Satır panoya kopyalandı") if api.copyToClip(metin) else _("Satır kopyalanamadı"))
        except Exception:
            log.exception("Tablo Gezimi satırı panoya kopyalayamadı.")
            ui.message(_("Satır kopyalanamadı"))

    @script(
        description=_("Tablonun tüm satırlarını panoya kopyalar (yalnızca tablo modunda)"),
        category=_("Tablo Gezimi"),
        gesture="kb:alt+shift+downarrow",
    )
    def script_tabloKopyala(self, gesture):
        if not self.acik:
            gesture.send()
            return
        if excel_hucresinde_mi():
            self.excel_kopyala("tablo")
            return
        try:
            belge = self.aktif_belge()
            if belge is None:
                ui.message(_("Tablo içinde değilsiniz"))
                return
            info = belge.selection
            konum = belge._getTableCellCoords(info)
            satir_sayisi, sutun_sayisi = belge._getTableDimensions(info)
            if not isinstance(satir_sayisi, int) or satir_sayisi < 1:
                ui.message(_("Tablonun satır sayısı belirlenemedi"))
                return
            if not isinstance(sutun_sayisi, int) or sutun_sayisi < 1:
                ui.message(_("Tablonun sütun sayısı belirlenemedi"))
                return
            if satir_sayisi > 100000 or sutun_sayisi > 4096:
                ui.message(_("Bu tablo tek seferde kopyalamak için çok büyük"))
                return
            self.kopyalama_numarasi += 1
            numara = self.kopyalama_numarasi
            # Satırları parça parça okuyarak uzun tablolarda NVDA'yı kilitleme.
            self.tablo_kopyalama_parcasi(belge, info, konum.tableID,
                                        satir_sayisi, sutun_sayisi, 1, [], numara, {})
        except Exception:
            log.exception("Tablo Gezimi tüm tabloyu panoya kopyalayamadı.")
            ui.message(_("Tablo kopyalanamadı"))

    def tablo_kopyalama_parcasi(self, belge, info, tablo_id, satir_sayisi,
                               sutun_sayisi, baslangic, satirlar, numara, onbellek):
        if not self.acik or numara != self.kopyalama_numarasi:
            return
        try:
            dur = time.perf_counter() + 0.014
            son = baslangic
            for satir in range(baslangic, min(satir_sayisi + 1, baslangic + 40)):
                hucre = None
                for sutun in range(1, sutun_sayisi + 1):
                    try:
                        hucre = belge._getTableCellAt(tablo_id, info, satir, sutun)
                        break
                    except LookupError:
                        continue
                if hucre is None:
                    satirlar.append(_("Boş"))
                else:
                    yer = belge._getTableCellCoords(hucre)
                    basliklar, degerler = self.satir_verisi(belge, hucre, yer, satir, onbellek)
                    satirlar.append(self.metin_hazirla(basliklar, degerler))
                son = satir + 1
                if time.perf_counter() >= dur:
                    break
            siradaki = son
            if siradaki <= satir_sayisi:
                queueHandler.queueFunction(queueHandler.eventQueue,
                    self.tablo_kopyalama_parcasi, belge, info, tablo_id,
                    satir_sayisi, sutun_sayisi, siradaki, satirlar, numara, onbellek)
                return
            self._kopyalamayi_arka_planda_tamamla(
                satirlar, numara, _("Tablonun {count} satırı panoya kopyalandı").format(count=satir_sayisi),
            )
        except Exception:
            log.exception("Tablo Gezimi tüm tabloyu panoya kopyalayamadı.")
            ui.message(_("Tablo kopyalanamadı"))

    @script(
        description=_("Bulunulan sütunun tüm hücrelerini panoya kopyalar (yalnızca tablo modunda)"),
        category=_("Tablo Gezimi"),
        gesture="kb:alt+shift+leftarrow",
    )
    def script_sutunKopyala(self, gesture):
        if not self.acik:
            gesture.send()
            return
        if excel_hucresinde_mi():
            self.excel_kopyala("sutun")
            return
        try:
            belge = self.aktif_belge()
            if belge is None:
                ui.message(_("Tablo içinde değilsiniz"))
                return
            info = belge.selection
            konum = belge._getTableCellCoords(info)
            satir_sayisi, _ = belge._getTableDimensions(info)
            if not isinstance(satir_sayisi, int) or not 1 <= satir_sayisi <= 100000:
                ui.message(_("Tablonun satır sayısı belirlenemedi veya çok büyük"))
                return
            self.kopyalama_numarasi += 1
            numara = self.kopyalama_numarasi
            self.sutun_kopyalama_parcasi(
                belge, info, konum.tableID, konum.col, satir_sayisi,
                1, [], set(), numara,
            )
        except Exception:
            log.exception("Tablo Gezimi sütunu panoya kopyalayamadı.")
            ui.message(_("Sütun kopyalanamadı"))

    def sutun_kopyalama_parcasi(self, belge, info, tablo_id, sutun,
                               satir_sayisi, baslangic, degerler, gorulen, numara):
        if not self.acik or numara != self.kopyalama_numarasi:
            return
        try:
            dur = time.perf_counter() + 0.014
            son = baslangic
            for satir in range(baslangic, min(satir_sayisi + 1, baslangic + 80)):
                son = satir + 1
                try:
                    hucre = belge._getTableCellAt(tablo_id, info, satir, sutun)
                    yer = belge._getTableCellCoords(hucre)
                except LookupError:
                    # Birleşik veya eksik hücrelerin olduğu tablolarda devam et.
                    continue
                if yer.tableID != tablo_id:
                    continue
                kimlik = (yer.row, yer.col)
                if kimlik in gorulen:
                    continue
                gorulen.add(kimlik)
                degerler.append(" ".join(hucre.text.split()) or _("Boş"))
                if time.perf_counter() >= dur:
                    break
            siradaki = son
            if siradaki <= satir_sayisi:
                queueHandler.queueFunction(
                    queueHandler.eventQueue, self.sutun_kopyalama_parcasi,
                    belge, info, tablo_id, sutun, satir_sayisi,
                    siradaki, degerler, gorulen, numara,
                )
                return
            if not degerler:
                ui.message(_("Bu sütunda kopyalanacak hücre bulunamadı"))
            else:
                self._kopyalamayi_arka_planda_tamamla(
                    degerler, numara, _("Sütunun {count} hücresi panoya kopyalandı").format(count=len(degerler)),
                )
        except Exception:
            log.exception("Tablo Gezimi sütunu panoya kopyalayamadı.")
            ui.message(_("Sütun kopyalanamadı"))

    @script(
        description=_("Bulunulan hücrenin yalnızca içeriğini panoya kopyalar (yalnızca tablo modunda)"),
        category=_("Tablo Gezimi"),
        gesture="kb:alt+shift+rightarrow",
    )
    def script_hucreKopyala(self, gesture):
        if not self.acik:
            gesture.send()
            return
        if excel_hucresinde_mi():
            self.excel_kopyala("hucre")
            return
        try:
            belge = self.aktif_belge()
            if belge is None:
                ui.message(_("Tablo içinde değilsiniz"))
                return
            # belge.selection, gezin() içinde hücrenin başına daraltılmış
            # (collapse) bir TextInfo'dur. Bu nedenle selection.text boş
            # olabilir. Sütun kopyalamada olduğu gibi hücrenin TAM
            # TextInfo'sunu tablo koordinatlarından yeniden al.
            secim = belge.selection
            konum = belge._getTableCellCoords(secim)
            hucre = belge._getTableCellAt(
                konum.tableID, secim, konum.row, konum.col,
            )
            metin = " ".join(hucre.text.split())
            if not metin:
                ui.message(_("Bulunulan hücre boş"))
                return
            ui.message(_("Hücre panoya kopyalandı") if api.copyToClip(metin) else _("Hücre kopyalanamadı"))
        except Exception:
            log.exception("Tablo Gezimi hücreyi panoya kopyalayamadı.")
            ui.message(_("Hücre kopyalanamadı"))

    def satir_verisi(self, belge, info, konum, satir, onbellek=None):
        try:
            _, sutun_sayisi = belge._getTableDimensions(info)
        except (LookupError, NotImplementedError):
            sutun_sayisi = 4096
        if not isinstance(sutun_sayisi, int) or sutun_sayisi < 1:
            sutun_sayisi = 4096

        basliklar, degerler = [], []
        gorulen = set()
        sutun = 1
        while sutun <= min(sutun_sayisi, 4096):
            try:
                hucre = belge._getTableCellAt(konum.tableID, info, satir, sutun)
                yer = belge._getTableCellCoords(hucre)
            except LookupError:
                if sutun_sayisi == 4096 and basliklar:
                    break
                sutun += 1
                continue
            kimlik = (yer.row, yer.col)
            if yer.tableID == konum.tableID and kimlik not in gorulen:
                gorulen.add(kimlik)
                baslik, deger = self.hucre_verisi(belge, hucre, yer, onbellek)
                basliklar.append(baslik)
                degerler.append(deger)
            sutun = max(sutun + 1, yer.col + yer.colSpan)
        if not basliklar:
            baslik, deger = self.hucre_verisi(belge, info, konum, onbellek)
            return [baslik], [deger]
        return basliklar, degerler

    def metin_hazirla(self, basliklar, degerler):
        if self.dll is not None:
            try:
                return satir_oku(self.dll, basliklar, [d if d else _("Boş") for d in degerler], baslik_oku=2)
            except (ValueError, RuntimeError):
                log.debugWarning("Tablo Gezimi Python biçimleyicisine geçti.", exc_info=True)
        return ", ".join(
            f"{baslik} {deger or _('Boş')}" if baslik else deger or _("Boş")
            for baslik, deger in zip(basliklar, degerler)
        )
