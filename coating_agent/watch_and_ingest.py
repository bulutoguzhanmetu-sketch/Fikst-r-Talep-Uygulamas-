#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
source_pdfs/ klasörünü izler; yeni veya değişmiş bir PDF düştüğünde
otomatik olarak ingest.main() çalıştırır.

Bu, "kendini eğitme" mekanizmasıdır: kullanıcı her yeni akademik makaleyi
klasöre attığında bilgi tabanı otomatik güncellenir.

Kullanım:
    python watch_and_ingest.py [--interval SANIYE]

Bağımlılık eklemeden (watchdog kütüphanesi olmadan) çalışacak şekilde
basit bir polling (periyodik tarama) yaklaşımı kullanılır.
"""

import argparse
import time

import config
import ingest


def poll_loop(interval: int) -> None:
    print(f"İzleniyor: {config.SOURCE_PDF_DIR} (her {interval} sn'de bir kontrol)")
    print("Çıkmak için Ctrl+C.")
    try:
        while True:
            ingest.main()
            time.sleep(interval)
    except KeyboardInterrupt:
        print("\nİzleme durduruldu.")


def main() -> None:
    parser = argparse.ArgumentParser(description="PDF klasörünü izleyip otomatik ingest eder.")
    parser.add_argument("--interval", type=int, default=30, help="Tarama aralığı (saniye, varsayılan: 30)")
    args = parser.parse_args()
    poll_loop(args.interval)


if __name__ == "__main__":
    main()
