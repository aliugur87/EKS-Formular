# EKS Form Doldurucu Pro

Almanya'daki serbest meslek sahipleri için EKS (Einkommen aus selbständiger Tätigkeit) formunu doldurma sürecini otomatikleştiren ve hızlandıran bir masaüstü uygulaması. Bu araç, BWA (Betriebswirtschaftliche Auswertung) dosyalarından verileri okur, akıllıca EKS alanlarıyla eşleştirir ve doldurulmuş Excel formunu dışa aktarır.



---

## ✨ Özellikler

-   **Otomatik BWA Analizi:** Standart BWA Excel dosyalarını otomatik olarak okur ve aylık verileri ayrıştırır.
-   **Akıllı Eşleştirme:** BWA hesap kodlarını ilgili EKS form alanlarıyla yüksek doğrulukla eşleştirir.
-   **Yapay Zeka Desteği (Claude AI):** Bilinmeyen BWA hesapları için otomatik EKS alanı önerileri sunar.
-   **Veri Yönetimi:** Müşteri bilgilerini ve geçmiş BWA yüklemelerini güvenli bir şekilde saklar.
-   **Düzenlenebilir Sonuçlar:** Eşleştirme sonuçlarını dışa aktarmadan önce doğrudan arayüz üzerinden düzenlemeye olanak tanır.
-   **Çoklu Dil Desteği:** Almanca ve Türkçe dil seçenekleri sunar.
-   **Tek Tıkla Dışa Aktarma:** Tüm verileri doğrudan resmi EKS Excel şablonuna doldurur ve kaydedilmeye hazır hale getirir.

---

## 🚀 Kullanım

Bu uygulamayı kullanmak için Python veya herhangi bir kütüphane kurmanıza gerek yoktur.

1.  **Uygulamayı İndirin:** Projenin [Releases (Sürümler) sayfasına](https://github.com/aliugur87/eks-1909/releases) gidin.
2.  En son sürümün altındaki "Assets" bölümünden `.exe` uzantılı dosyayı indirin.
3.  İndirdiğiniz `EKS_Form_Doldurucu.exe` dosyasına çift tıklayarak uygulamayı başlatın.

### Yapay Zeka Özelliğini Aktif Etme

Yapay zeka destekli eşleştirme önerilerini kullanmak için:
1.  Uygulama içindeki **Ayarlar (⚙️)** menüsünü açın.
2.  `settings.json` adında bir dosya oluşturulacaktır. Bu dosyayı bir metin düzenleyici ile açın.
3.  Kendi [Anthropic Claude](https://www.anthropic.com/) API anahtarınızı ilgili alana girin ve kaydedin.

---

## 🛠️ Geliştirme

Projeyi yerel olarak geliştirmek isterseniz:

1.  Depoyu klonlayın: `git clone https://github.com/aliugur87/eks-1909.git`
2.  Gerekli kütüphaneleri yükleyin: `pip install -r requirements.txt`
3.  Uygulamayı çalıştırın: `python form_doldurucu.py`