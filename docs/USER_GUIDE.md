# ReportIQ - Kullanıcı Kılavuzu 🛡️

ReportIQ, siber güvenlik zafiyet tarama verilerini Excel formatından alıp, görsel grafikler ve analizlerle zenginleştirilmiş profesyonel Word raporlarına dönüştüren bir masaüstü uygulamasıdır.

## 🚀 Başlangıç

### Kurulum

1. **Python Yükleyin**: Bilgisayarınızda Python 3.9 veya üzeri bir sürümün yüklü olduğundan emin olun.
2. **Bağımlılıkları Yükleyin**:
   Terminal veya komut satırını açın ve proje klasöründe şu komutu çalıştırın:
   ```bash
   pip install -r requirements.txt
   ```

### Uygulamayı Çalıştırma

Proje dizininde aşağıdaki komutu çalıştırın:
```bash
python main.py
```

---

## 🖥️ Arayüz Kullanımı

Uygulama açıldığında koyu tema (dark mode) ile karşılaşacaksınız. Arayüz 3 ana panele ayrılmıştır:

1. **Sol Panel**: Dosya seçimi ve Filtreler
2. **Orta Panel**: Rapor Bölümleri (Grafik Seçimi)
3. **Sağ Panel**: Veri Önizleme

### Adım 1: Excel Dosyası Yükleme

1. Sol üstteki **"Gözat"** (Browse) butonuna tıklayın.
2. Bilgisayarınızdan zafiyet verilerini içeren Excel dosyasını (`.xlsx` veya `.xls`) seçin.
3. Dosya yüklendiğinde "✓ Dosya Yüklendi" mesajını göreceksiniz ve filtreler otomatik olarak dolacaktır.

**Not**: Excel dosyanızdaki sütun isimleri otomatik olarak algılanır. (Örn: TICKETID, PRIORITY, STATUS vb.)

### Adım 2: Filtreleme (Opsiyonel)

Yüklenen veriler üzerinde istediğiniz filtreleri uygulayabilirsiniz:

- **SLA Status**: Sadece SLA süresi geçenleri (Out of SLA) veya SLA içindekileri (In SLA) seçebilirsiniz.
- **Status**: Zafiyet durumuna göre filtreleme yapabilirsiniz (Örn: Sadece PENDING ve QUEUED olanlar).
- **Priority**: Kritiklik seviyesine göre (High, Critical) filtreleyebilirsiniz.
- **Tool/Source**: Hangi araçla tarandığına göre (TenableSC, NessusAgent vb.) seçim yapabilirsiniz.
- **Year**: Belirli bir yılda oluşturulan kayıtları filtreleyebilirsiniz.
- **Department**: Belirli departmanlara odaklanabilirsiniz.

Filtreleri değiştirdiğinizde Sağ Paneldeki **Veri Önizleme** ve kayıt sayıları (Toplam/Filtrelenen) anlık olarak güncellenir.

### Adım 3: Rapor İçeriğini Seçme

Orta panelde, oluşturulacak Word raporunda yer almasını istediğiniz analizleri seçin:

- **📊 Yıllara Göre Açık Zafiyet**: Yıllık dağılım grafiği.
- **🎯 Priority Dağılımı**: Kritik seviye pasta grafiği.
- **👥 Line Manager Kırılımı**: Hangi yöneticide ne kadar zafiyet var.
- **🏢 Departman Kırılımı**: Departman bazlı dağılım.
- **🔧 Tool Kırılımı**: Tarama kaynaklarına göre dağılım.
- **⏰ SLA Durumu**: SLA uyumluluk oranı.
- **📈 Trend Analizi**: Zaman içindeki artış/azalış trendi.
- **🔥 Top 10 Zafiyet**: En sık görülen 10 zafiyet ve detaylı açıklamaları. (Sözlükten otomatik açıklama ve çözüm önerisi eklenir)
- **💻 IP Bazlı Yoğunluk**: En çok zafiyet barındıran IP adresleri.
- **📅 Ortalama Çözüm Süresi**: Kapatılan zafiyetlerin ortalama kapatılma gün sayısı.
- **⚠️ SLA Aşım Analizi**: SLA süresi ne kadar aşılmış analizi.

Hepsini seçmek için **"Tümünü Seç"** butonunu kullanabilirsiniz.

### Adım 4: Rapor Oluşturma

1. Alt kısımdaki **"🚀 Rapor Oluştur"** butonuna tıklayın.
2. Açılan pencerede raporu kaydetmek istediğiniz konumu ve dosya adını belirleyin.
3. Uygulama grafikleri oluşturup Word belgesini hazırlarken bekleyin (İlerleme çubuğunu takip edebilirsiniz).
4. İşlem bittiğinde "Rapor kaydedildi" mesajı çıkacaktır.

---

## 🌍 Dil Seçeneği

Uygulamanın sağ üst köşesindeki **TR | EN** düğmesi ile arayüzü ve rapor dilini Türkçe veya İngilizce olarak değiştirebilirsiniz.
- **TR**: Arayüz ve Rapor çıktıları Türkçe olur.
- **EN**: Arayüz ve Rapor çıktıları İngilizce olur.

---

## 🛠️ Sorun Giderme

- **Dosya yüklenmiyor**: Excel dosyanızın bozuk olmadığından ve `.xlsx` formatında olduğundan emin olun.
- **Veri bulunamadı hatası**: Seçtiğiniz filtreler çok kısıtlayıcı olabilir (Örn: 2024 yılında olup durumu CLOSED olan kayıt yoksa). Filtreleri temizleyip tekrar deneyin.
- **Grafikler boş çıkıyor**: İlgili analiz için veride eksiklik olabilir (Örn: Tarih sütunu yoksa Trend grafiği çıkmaz).

---


## 📞 Destek

Herhangi bir hata veya öneri için geliştirici ekibiyle iletişime geçebilirsiniz.

ReportIQ 🛡️ - Güvenliğiniz İçin Akıllı Raporlama

