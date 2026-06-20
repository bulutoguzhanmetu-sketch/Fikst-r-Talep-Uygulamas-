#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PDF -> metin -> chunk -> embedding -> ChromaDB ingest scripti.

Kullanım:
    python ingest.py [pdf_klasoru]

pdf_klasoru verilmezse config.SOURCE_PDF_DIR kullanılır.
Sadece yeni veya değişmiş (hash farklı) PDF'ler işlenir.
"""

import hashlib
import json
import sys

import chromadb
from pypdf import PdfReader

import config


def file_hash(path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(65536), b""):
            h.update(block)
    return h.hexdigest()


def load_manifest() -> dict:
    if config.MANIFEST_PATH.exists():
        return json.loads(config.MANIFEST_PATH.read_text(encoding="utf-8"))
    return {}


def save_manifest(manifest: dict) -> None:
    config.MANIFEST_PATH.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def extract_pages(pdf_path) -> list[str]:
    reader = PdfReader(str(pdf_path))
    return [page.extract_text() or "" for page in reader.pages]


def chunk_page_text(text: str) -> list[str]:
    text = text.strip()
    if not text:
        return []
    chunks = []
    step = config.CHUNK_SIZE_CHARS - config.CHUNK_OVERLAP_CHARS
    for start in range(0, len(text), step):
        chunk = text[start:start + config.CHUNK_SIZE_CHARS].strip()
        if chunk:
            chunks.append(chunk)
        if start + config.CHUNK_SIZE_CHARS >= len(text):
            break
    return chunks


def remove_existing_chunks(collection, source_name: str) -> None:
    existing = collection.get(where={"source": source_name})
    if existing and existing.get("ids"):
        collection.delete(ids=existing["ids"])


def ingest_pdf(collection, pdf_path) -> int:
    source_name = pdf_path.name
    remove_existing_chunks(collection, source_name)

    pages = extract_pages(pdf_path)
    ids, documents, metadatas = [], [], []
    for page_no, page_text in enumerate(pages, start=1):
        for chunk_idx, chunk in enumerate(chunk_page_text(page_text)):
            ids.append(f"{source_name}::p{page_no}::c{chunk_idx}")
            documents.append(chunk)
            metadatas.append({"source": source_name, "page": page_no})

    if documents:
        collection.add(ids=ids, documents=documents, metadatas=metadatas)
    return len(documents)


def main() -> None:
    pdf_dir = config.SOURCE_PDF_DIR
    if len(sys.argv) > 1:
        from pathlib import Path
        pdf_dir = Path(sys.argv[1]).expanduser()

    pdf_files = sorted(pdf_dir.glob("*.pdf"))
    if not pdf_files:
        print(f"Klasörde PDF bulunamadı: {pdf_dir}")
        return

    client = chromadb.PersistentClient(path=str(config.CHROMA_DB_DIR))
    collection = client.get_or_create_collection(config.COLLECTION_NAME)

    manifest = load_manifest()
    processed, skipped = 0, 0

    for pdf_path in pdf_files:
        digest = file_hash(pdf_path)
        if manifest.get(pdf_path.name) == digest:
            skipped += 1
            continue

        n_chunks = ingest_pdf(collection, pdf_path)
        manifest[pdf_path.name] = digest
        processed += 1
        print(f"İşlendi: {pdf_path.name} ({n_chunks} parça)")

    save_manifest(manifest)
    print(f"\nToplam: {processed} dosya işlendi, {skipped} dosya değişmediği için atlandı.")


if __name__ == "__main__":
    main()
