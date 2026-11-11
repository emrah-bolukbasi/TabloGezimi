TABLO OKUMA DLL'Sİ — SÜRÜM 1.1

Windows için hazır dosyalar
  x64\TabloOkuma.dll — 64 bit Python ve uygulamalar
  x86\TabloOkuma.dll — 32 bit Python ve uygulamalar

DLL mimarisini Windows'un değil, DLL'yi çağıran uygulamanın bit
sayısına göre seçin. Örnek Python betiği bu seçimi otomatik yapar.
ZIP'in tamamını bir klasöre çıkarın. O klasörde Komut İstemi açın.

  py -3 ornek.py
  py -3 ornek.py --test

İlk komut örnek fiyat tablosunun satırlarını konsola yazar. NVDA ile
konsoldan okuyabilirsiniz. İkinci komut DLL üzerinden doğrulama yapar.
Python örneği için ek pip paketi gerekmez.

UYGULAMAYA EKLEME
ornek.py içindeki yukle ve satir_oku işlevlerini kullanabilirsiniz.
Örnek çağrı:
  satir_oku(dll, ["Mağaza", "Fiyat"], ["Trendyol", "17.249 TL"], [1, 0])
Sonuç:
  Fiyat: 17.249 TL, Mağaza: Trendyol

Sütun numaraları sıfırdan başlar. Seçimin sırası korunur. Boş seçim
bütün sütunları okur. Başlıkları kapatmak için baslik_oku=False verin. Konuşma biçimi için
baslik_oku=2 verin. Örnek: “İsim Emrah, Numara 2121, Sıra 23”.
Boş veya eksik bir hücre “Boş” olarak döner. Boş başlık atlanır.

NVDA Eklentisiyle Kullanım
DLL bir satırın metnini hazırlar. Tabloyu web sayfasından ya da başka
bir programdan okuma, ok tuşlarını yönetme ve metni NVDA'ya okutma
işlerini ileride eklenti yapabilir. Uygulama da kendi tablosunun
verilerini verip dönen metni bir erişilebilir durum alanına aktarabilir.
Bu paket DLL aşamasını ve konsol örneğini içerir.

C VE C++ ARAYÜZÜ
kaynak\tablo.h bildirimleri içerir. İki işlev de C çağrı düzenini
kullanır (x86'da cdecl). DLL'yi tam yoluyla yükleyip Tablo_ABI ve
Tablo_Satir adlarını çözümleyebilirsiniz. Tablo_ABI sürüm 1 döndürür.

Tablo_Satir girdileri:
  basliklar: sütun başına bir UTF-16 metin işaretçisi veya NULL.
  hucreler: sütun başına bir UTF-16 metin işaretçisi. Tekil işaretçiler
    NULL olabilir. Her metin sıfır kod birimiyle sonlanır.
  sutun_sayisi: 1–4096.
  secim: sıfır tabanlı sütun indisleri. Tekrarlar korunur.
  secim_sayisi: 0–4096. Sıfırsa secim işaretçisi kullanılmaz.
  baslik_oku: 0 başlıksız, 1 iki noktalı başlık, 2 konuşma biçimi.
    Kip 2, DLL 1.1 ile eklendi. Çağrı imzası ve ABI numarası aynı.
  cikti, kapasite: çıktı dizisi ve UTF-16 kod birimi kapasitesi.

Önce cikti=NULL ve kapasite=0 ile gereken kapasiteyi alın. Pozitif
sonuç sonlandırıcıyı da içerir. Yeterli bir diziyle tekrar çağırın.
Kapasite yetersizse gereken sayı döner ve dizi değişmez.
-1 geçersiz parametre, -2 geçersiz sütun indisi, -3 uzunluk sınırı.
Sonlandırıcı dahil üst sınır 1.048.576 UTF-16 kod birimidir.

Belleği çağıran uygulama ayırır. Giriş dizileri ve işaretçileri çağrı
boyunca geçerli ve sabit kalmalıdır. Çıktı belleği girdilerle
çakışmamalıdır. Bağımsız tamponlarla eşzamanlı çağrı yapılabilir.

DİĞER DİLLERDEN KULLANIM
C: ZIP'ten çıkan ana klasörde Visual Studio x64 veya x86 Native
Tools Komut İstemi açın. Şu komutlarla örneği çalıştırın:
  cl /nologo /W4 ornek.c
  ornek.exe

C#: ZIP'ten çıkan ana klasörde .NET 8 veya üstü Windows SDK ile:
  dotnet new console -o TabloOrnek
  copy /Y Ornek.cs TabloOrnek\Program.cs
  dotnet run --project TabloOrnek -- "%CD%"

C++ kaynak\tablo.h başlığını doğrudan kullanabilir. Rust ve
Delphi gibi diller aynı C ABI ile çağrı yapabilir. x86 için cdecl,
Windows x64 için standart Windows çağrı düzenini seçin. İşaretçi
dizileri, sıfırla biten UTF-16 metinleri gösterir.

KAYNAKTAN DERLEME
Visual Studio C++ Build Tools ile x64 veya x86 Native Tools
Komut İstemi açıp kaynak\derle.bat çalıştırın. Üretilen DLL
kaynak klasöründe olur. İki mimariyi ayrı komut istemlerinde derleyin.

DOĞRULAMA
Kaynak kodun satır oluşturma testleri Linux'ta geçti. Windows
DLL'lerinin mimarileri ve dışa aktarımları denetlendi.
Windows ve NVDA üzerinde çalıştırma testi yapılmadı.
