# Optik Kaplama & Malzeme Mühendisi Araştırma Ajanı

Optik kaplama ve malzeme mühendisliği alanında derlenmiş akademik PDF'lerden
yüksek doğrulukla akademik rapor taslakları üreten bir RAG (Retrieval-Augmented
Generation) aracı.

> **Not:** Bu araç bir LLM'i "fine-tune" etmez. Bunun yerine, eklediğiniz
> PDF'lerden bir bilgi tabanı (vektör veritabanı) oluşturur ve raporları
> SADECE bu kaynaklara dayandırarak üretir. Yeni makale ekledikçe bilgi
> tabanı büyür ("kendini eğitme" = kanıt tabanının sürekli genişlemesi),
> ama model ağırlıkları değişmez. Bu yaklaşım, halüsinasyonu önemli ölçüde
> azaltır çünkü her iddia bir kaynağa (dosya + sayfa) bağlanır.

## Kurulum

```bash
pip install -r coating_agent/requirements.txt
export ANTHROPIC_API_KEY="sk-ant-..."
```

## Kullanım

### 1. Akademik makaleleri ekleyin

PDF dosyalarını `coating_agent/source_pdfs/` klasörüne koyun.

### 2. Bilgi tabanını güncelleyin (ingest)

```bash
python coating_agent/ingest.py
```

Sadece yeni veya değişmiş PDF'ler işlenir (hash kontrolü ile tespit edilir).

**Otomatik mod (her yeni makaleyi otomatik işler):**

```bash
python coating_agent/watch_and_ingest.py --interval 30
```

Bu komut `source_pdfs/` klasörünü periyodik olarak tarar; siz çalışırken
klasöre yeni bir PDF bıraktığınızda bilgi tabanı otomatik güncellenir.

### 3. Rapor üretin

```bash
python coating_agent/report_generator.py "Ta2O5/SiO2 ince film kaplamalarda lazer hasar eşiği"
```

Çıktı: `coating_agent/output/<tarih>_<konu>.pdf`

Rapor üretilirken kaynaklarda yeterli bilgi bulunamazsa, ajan bunu tahmin
etmek yerine açıkça belirtir.

## academic-research-skills ile birlikte kullanım (opsiyonel)

Daha kapsamlı atıf doğrulama (Semantic Scholar + OpenAlex + Crossref
çapraz kontrolü) ve akran inceleme istiyorsanız, Claude Code'a
[academic-research-skills](https://github.com/imbad0202/academic-research-skills)
plugin'ini kurabilirsiniz:

```
/plugin marketplace add Imbad0202/academic-research-skills
/plugin install academic-research-skills
```

Kurulumdan sonra:

1. `report_generator.py` ile ürettiğiniz taslak raporu (`output/` altındaki PDF
   veya rapor metni) Claude Code'da açın.
2. Plugin'in **Academic Pipeline** veya **Academic Paper Reviewer** skill'ini
   bu taslak üzerinde çalıştırın — atıf doğrulama, iddia denetimi ve akran
   inceleme aşamalarından geçirir.

> Lisans notu: academic-research-skills **CC-BY-NC 4.0** ile lisanslanmıştır
> (ticari olmayan kullanım + atıf gerektirir). Ticari bir bağlamda
> kullanmadan önce lisans şartlarını gözden geçirin.

## Klasör Yapısı

```
coating_agent/
├── config.py             # ayarlar (API key, dizinler, chunk boyutu)
├── ingest.py              # PDF -> chunk -> embedding -> ChromaDB
├── watch_and_ingest.py    # source_pdfs/ klasörünü izleyip otomatik ingest eder
├── report_generator.py    # konudan rapor taslağı üretir
├── pdf_writer.py          # raporu PDF'e döker
├── source_pdfs/           # buraya akademik PDF'lerinizi koyun (commit edilmez)
├── data/                  # vektör veritabanı + manifest (commit edilmez)
└── output/                # üretilen rapor PDF'leri (commit edilmez)
```
