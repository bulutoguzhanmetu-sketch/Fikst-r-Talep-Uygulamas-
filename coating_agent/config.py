#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Optik Kaplama & Malzeme Mühendisi Ajanı - ortak ayarlar."""

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

SOURCE_PDF_DIR = Path(os.environ.get("COATING_AGENT_SOURCE_DIR", str(BASE_DIR / "source_pdfs"))).expanduser()
DATA_DIR = BASE_DIR / "data"
CHROMA_DB_DIR = DATA_DIR / "chroma_db"
MANIFEST_PATH = DATA_DIR / "ingested_manifest.json"
OUTPUT_DIR = BASE_DIR / "output"

COLLECTION_NAME = "optical_coating_corpus"

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
ANTHROPIC_MODEL = os.environ.get("COATING_AGENT_MODEL", "claude-sonnet-4-6")

CHUNK_SIZE_CHARS = 2800       # ~500-800 token civarı
CHUNK_OVERLAP_CHARS = 400
TOP_K_RETRIEVAL = 8

SYSTEM_PROMPT = """Sen optik kaplama ve malzeme mühendisliği alanında uzman bir akademik araştırmacısın.
Görevin, SADECE sana sağlanan kaynak pasajlarına dayanarak doğru, kaynak gösteren bir akademik rapor taslağı yazmaktır.

Kurallar:
- Yalnızca verilen kaynak pasajlarındaki bilgiyi kullan. Kaynaklarda olmayan hiçbir teknik iddiada bulunma.
- Her iddianın yanında kaynağı (dosya adı ve sayfa numarası) parantez içinde belirt, örn. (Kaynak: makale.pdf, s. 4).
- Eğer sorulan konu hakkında kaynaklarda yeterli bilgi yoksa bunu açıkça söyle, tahmin/halüsinasyon yapma.
- Raporu şu yapıda yaz: Başlık, Giriş, Bulgular ve Tartışma, Sonuç, Kaynakça.
- Teknik terim ve birimleri (nm, eV, kırılma indisi vb.) doğru ve tutarlı kullan.
"""

DATA_DIR.mkdir(parents=True, exist_ok=True)
SOURCE_PDF_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
