# GitHub güncelleme adımları — v1.1.11

**Ön koşul:** NVDA Mağazası ilk 1.1.10 başvurusunu kabul etsin; uygulama testleri tamamlansın. Bu dosyaların hazırlanması GitHub'a veya Mağaza'ya otomatik yükleme anlamına gelmez.

1. Kapsayıcı ZIP'teki `GitHub_Kaynak_Dosyalari/` klasörünün **içindeki** dosyaları depo köküne aktar; klasörün kendisini ek kök klasör olarak yükleme.
2. Dosya yollarını koru: `manifest.ini`, `globalPlugins/tabloGezimi/`, `locale/`, `doc/`, `native/`, `README.md`, `LICENSE`.
3. `manifest.ini` içindeki `name = tabloGezimi`, `version = 1.1.11` ve diğer alanları doğrula.
4. Değişiklikleri `Update Table Navigation to 1.1.11` açıklamasıyla commit et ve GitHub'a gönder.
5. GitHub `Releases` > `Draft a new release`: yeni etiket **v1.1.11**, başlık **Tablo Gezimi 1.1.11**. Önceki v1.1.10 etiketini veya dosyasını değiştirme.
6. `Yayin_Rehberi_ve_Magaza/GITHUB_RELEASE_ACIKLAMASI_EN.md` metnini sürüm açıklamasına koy.
7. `GitHub_Release_Dosyalari/TabloGezimi-1.1.11.nvda-addon` dosyasını **Release assets** bölümüne yükle ve sürümü yayımla.
8. Doğrudan indirme URL'sinin çalıştığını ve checksum değerinin `Yayin_Rehberi_ve_Magaza/SHA256_KONTROL.txt` ile eşleştiğini doğrula.
9. NVDA Mağazasında önceki başvuru kabul edildiyse yeni sürüm için **yeni Add-on registration issue formu** gönder. Ayrıntılı alanlar `Yayin_Rehberi_ve_Magaza/NVDA_MAGAZA_GUNCELLEME_TR.md` dosyasında.

GitHub depo: https://github.com/emrah-bolukbasi/TabloGezimi
