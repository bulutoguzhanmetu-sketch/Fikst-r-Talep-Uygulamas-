#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Verilen bir konu için ChromaDB'den ilgili pasajları çekip Claude API ile
kaynak gösteren bir akademik rapor taslağı üretir ve PDF olarak kaydeder.

Kullanım:
    python report_generator.py "Ta2O5/SiO2 ince film kaplamalarda lazer hasar eşiği"
"""

import sys
from datetime import datetime

import anthropic
import chromadb

import config
from pdf_writer import write_report_pdf


def retrieve_context(topic: str) -> list[dict]:
    client = chromadb.PersistentClient(path=str(config.CHROMA_DB_DIR))
    collection = client.get_or_create_collection(config.COLLECTION_NAME)

    if collection.count() == 0:
        return []

    results = collection.query(query_texts=[topic], n_results=config.TOP_K_RETRIEVAL)
    documents = results.get("documents", [[]])[0]
    metadatas = results.get("metadatas", [[]])[0]

    return [
        {"text": doc, "source": meta.get("source", "bilinmiyor"), "page": meta.get("page", "?")}
        for doc, meta in zip(documents, metadatas)
    ]


def build_user_prompt(topic: str, passages: list[dict]) -> str:
    if not passages:
        return (
            f"Konu: {topic}\n\n"
            "Bilgi tabanında bu konuyla ilgili hiçbir kaynak pasaj bulunamadı. "
            "Lütfen raporu, kaynaklarda yeterli bilgi bulunmadığını açıkça belirterek yaz."
        )

    passage_blocks = []
    for i, p in enumerate(passages, start=1):
        passage_blocks.append(
            f"[Kaynak {i}: {p['source']}, sayfa {p['page']}]\n{p['text']}"
        )

    return (
        f"Konu: {topic}\n\n"
        "Aşağıdaki kaynak pasajlarına dayanarak yapılandırılmış bir akademik rapor taslağı yaz.\n\n"
        + "\n\n---\n\n".join(passage_blocks)
    )


def generate_report(topic: str) -> str:
    if not config.ANTHROPIC_API_KEY:
        raise SystemExit("ANTHROPIC_API_KEY ortam değişkeni ayarlanmamış.")

    passages = retrieve_context(topic)
    user_prompt = build_user_prompt(topic, passages)

    client = anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)
    message = client.messages.create(
        model=config.ANTHROPIC_MODEL,
        max_tokens=4096,
        system=config.SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}],
    )
    return "".join(block.text for block in message.content if block.type == "text")


def main() -> None:
    if len(sys.argv) < 2:
        print('Kullanım: python report_generator.py "<konu>"')
        sys.exit(1)

    topic = sys.argv[1]
    report_text = generate_report(topic)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_topic = "".join(c if c.isalnum() else "_" for c in topic)[:60]
    output_path = config.OUTPUT_DIR / f"{timestamp}_{safe_topic}.pdf"

    write_report_pdf(topic, report_text, output_path)
    print(f"Rapor oluşturuldu: {output_path}")


if __name__ == "__main__":
    main()
