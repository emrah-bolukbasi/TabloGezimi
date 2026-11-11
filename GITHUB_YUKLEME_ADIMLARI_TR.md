# GitHub’a yükleme: NVDA ile adım adım

Depo: https://github.com/emrah-bolukbasi/TabloGezimi

1. Bu ZIP dosyasını Windows’ta bir klasöre çıkar. ZIP’in kendisini GitHub köküne koyma.
2. Depoda “Dosya ekle” (Add file) > “Dosya yükle” (Upload files) seç.
3. ZIP içeriğinin **kökündeki** dosyaları ve `globalPlugins`, `locale`, `doc`, `native` klasörlerini, yolları korunacak şekilde yükle. GitHub web arayüzünde klasör yükleme NVDA ile güçse GitHub Desktop/Git kullan; klasörleri iç içe yanlış yerleştirme.
4. GitHub’da var olan README.md dosyasını paketteki README.md ile değiştir.
5. Aşağıdaki yolların doğru göründüğünü kontrol et: LICENSE, README.md, manifest.ini, native/tablo.c, globalPlugins/tabloGezimi/__init__.py.
6. Değişiklikleri ana dala kaydet (Commit changes).
7. Sonra Releases > Draft a new release: etiket `v1.1.10`, başlık `Tablo Gezimi 1.1.10`.
8. Kurulabilir `TabloGezimi-1.1.10_GPLv2.nvda-addon` dosyasını Release varlıklarına (assets) yükle; ZIP kaynak paketini release yerine kaynak depoda açılmış dosyalar olarak bulundur.
9. Resmi NVDA mağazası başvurusundan önce bağlantı, manifest sürümleri, lisans ve NVDA üzerindeki işlevleri doğrula.

Uyarı: Bu çalışma GitHub’a otomatik yükleme yapmaz. Kullanıcı hesabında kaydetme ve yayınlama işlemleri ayrıca gereklidir.
