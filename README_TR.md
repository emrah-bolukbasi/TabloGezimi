# Tablo Gezimi — Türkçe kullanım kılavuzu

**Sürüm:** 1.1.11 (yayına hazırlık)  
**Geliştirici:** Emrah Bölükbaşı

Bu sürüm, GitHub'da yayımlanmış 1.1.10'dan sonra geliştirilen ve yerel olarak 1.1.17'ye kadar ilerleyen tüm özelliklerin 1.1.11 adıyla birleştirilmiş hâlidir.

## Tablo gezinmesi

- **Alt + Shift + O**: Tablo gezinmesini açar/kapatır. Açık/kapalı tercihi NVDA yeniden başlatıldığında korunur.
- **Ctrl + Shift + Alt + Yukarı Ok**: 1. ve 2. gezinme modları arasında geçer.
  - **1. Mod:** Yukarı/aşağı oklarla satırın tamamını; sağ/sol oklarla tek hücreyi okur.
  - **2. Mod:** Dört ok tuşuyla yalnızca gidilen hücrenin içeriğini okur. Excel, desteklenen tarayıcı ve erişilebilir PDF tablolarında kullanılabilir.
  - Tablo gezimi her yeniden etkinleştirildiğinde **1. mod** seçilir.
- **Ctrl + Shift + Alt + Aşağı Ok**: Hücreler arasında konuşma duraklaması için Türkçe giriş penceresi açar. **0–5000 ms** girilebilir; varsayılan **150 ms**. Bu tuş NVDA **Girdi Hareketleri** bölümünden değiştirilebilir.

## Kopyalama kısayolları

| Kısayol | İşlem |
| --- | --- |
| Alt + Shift + Yukarı Ok | Satırı kopyala |
| Alt + Shift + Aşağı Ok | Tabloyu kopyala |
| Alt + Shift + Sol Ok | Sütunu kopyala |
| Alt + Shift + Sağ Ok | Hücreyi kopyala |

Kopyalama kısayolları her iki modda da çalışır ve tablo gezimi açık olmalıdır.

## Excel'e özel iyileştirmeler

Ok tuşlarıyla veri bölgesinin dışındaki hücrelere geçilebilir. Boş hücreye geldiğinde, örneğin **A5**, hücrenin adresi seslendirilir. Ancak toplu satır/sütun/tablo okuma ve kopyalama işlemleri boş değerleri atlar. Excel'in bitişik veri bölgesi (`CurrentRegion`) temel alınır; tamamen boş satır veya sütunlarla ayrılan bloklar ayrı değerlendirilir. Dolu alanlar arasında boş ayırıcı bulunmuyorsa ayrı tabloların tamamını güvenilir biçimde ayırmak mümkün olmayabilir.

## Dikkat edilmesi gerekenler

- Excel desteği Windows'ta kurulu Microsoft Excel'in COM sistemini kullanır.
- Tarayıcı/PDF tabloları, NVDA'nın erişebildiği tablo bilgilerini sağlamalıdır. Taranmış/resim PDF'ler için OCR dâhil değildir.
- Yeni mod değişikliği ve süre penceresindeki bazı mesajlar kodda Türkçe yazılmıştır; yeni metinlerin diğer dillere çevirisi yayın öncesi gözden geçirilmelidir.
- Bildirilen NVDA uyumluluğu gerçek ortamda ayrıca sınanmalıdır.

## Lisans

GPL-2.0-only. Telif hakkı (C) 2026 Emrah Bölükbaşı. Ayrıntılar [LICENSE](LICENSE) ve [COPYRIGHT.md](COPYRIGHT.md) dosyalarında.
