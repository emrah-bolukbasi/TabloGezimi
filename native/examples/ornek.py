"""Windows: py -3 ornek.py | Dogrulama: py -3 ornek.py --test"""
import ctypes as c
from pathlib import Path
import sys

U16 = c.c_uint16
Metin = c.POINTER(U16)
Dizin = Path(__file__).resolve().parent


def yukle(yol=None):
    if yol is None:
        mimari = "x64" if c.sizeof(c.c_void_p) == 8 else "x86"
        yol = Dizin / mimari / "TabloOkuma.dll"
    dll = c.CDLL(str(yol))
    dll.Tablo_ABI.argtypes = []
    dll.Tablo_ABI.restype = c.c_int
    if dll.Tablo_ABI() != 1:
        raise RuntimeError("Uyumsuz TabloOkuma DLL sürümü.")
    dll.Tablo_Satir.argtypes = [
        c.POINTER(Metin), c.POINTER(Metin), c.c_int,
        c.POINTER(c.c_int), c.c_int, c.c_int, Metin, c.c_int,
    ]
    dll.Tablo_Satir.restype = c.c_int
    return dll


def metin_dizisi(metinler):
    tamponlar = []
    for metin in metinler:
        if metin is None:
            tamponlar.append(None)
            continue
        if "\0" in metin:
            raise ValueError("Hücre metni NUL karakteri içeremez.")
        veri = metin.encode("utf-16-le") + b"\0\0"
        tamponlar.append((U16 * (len(veri) // 2)).from_buffer_copy(veri))
    return (Metin * len(tamponlar))(*tamponlar)


def satir_oku(dll, basliklar, hucreler, secim=(), baslik_oku=True):
    if basliklar is not None and len(basliklar) != len(hucreler):
        raise ValueError("Başlık ve hücre sayısı eşit olmalı.")
    if any(not isinstance(i, int) or not 0 <= i < len(hucreler) for i in secim):
        raise ValueError("Geçersiz sütun numarası.")
    b = None if basliklar is None else metin_dizisi(basliklar)
    h = metin_dizisi(hucreler)
    s = (c.c_int * len(secim))(*secim) if secim else None
    arg = (b, h, len(hucreler), s, len(secim), int(baslik_oku))
    uzunluk = dll.Tablo_Satir(*arg, None, 0)
    if uzunluk <= 0:
        raise ValueError(f"Satır hazırlanamadı. Hata kodu {uzunluk}.")
    cikti = (U16 * uzunluk)()
    sonuc = dll.Tablo_Satir(*arg, cikti, uzunluk)
    if sonuc != uzunluk:
        raise RuntimeError(f"Beklenmeyen DLL sonucu {sonuc}.")
    return bytes(cikti)[:-2].decode("utf-16-le")


def dogrula(dll):
    assert satir_oku(dll, ["Ürün", "Fiyat"], ["Çay 🍵", "120 TL"], [1, 0]) == "Fiyat: 120 TL, Ürün: Çay 🍵"
    assert satir_oku(dll, ["İsim", "Numara", "Sıra"], ["Emrah", "2121", "23"], baslik_oku=2) == "İsim Emrah, Numara 2121, Sıra 23"
    assert satir_oku(dll, None, [None, ""], baslik_oku=False) == "Boş, Boş"
    assert satir_oku(dll, ["", None], ["A", "B"]) == "A, B"
    h = metin_dizisi(["Çay"])
    cikti = (U16 * 2)(77, 88)
    assert dll.Tablo_Satir(None, h, 1, None, 0, 0, cikti, 2) == 4
    assert list(cikti) == [77, 88]
    secim = (c.c_int * 1)(1)
    assert dll.Tablo_Satir(None, h, 1, secim, 1, 0, None, 0) == -2
    assert dll.Tablo_Satir(None, None, 0, None, 0, 0, None, 0) == -1
    h = metin_dizisi(["x" * 1048576])
    assert dll.Tablo_Satir(None, h, 1, None, 0, 0, None, 0) == -3
    print("Bütün kontroller geçti.")


if __name__ == "__main__":
    dll = yukle()
    if "--test" in sys.argv:
        dogrula(dll)
    else:
        print(satir_oku(dll, ["İsim", "Numara", "Sıra"], ["Emrah", "2121", "23"], baslik_oku=2))
        basliklar = ["Mağaza", "Fiyat", "Stok", "Kargo"]
        satirlar = [
            ["Hepsiburada", "17.499 TL", "Stokta", "Ücretsiz"],
            ["Trendyol", "17.249 TL", "Stokta", "50 TL"],
            ["Amazon", "16.999 TL", "Stokta yok", None],
        ]
        for satir in satirlar:
            print(satir_oku(dll, basliklar, satir, [0, 1, 2]))
