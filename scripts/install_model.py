"""Download the same GGUF used by the existing local index; verify SHA-256."""

import hashlib
import os
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
URL = "https://huggingface.co/ggml-org/embeddinggemma-300m-qat-q8_0-GGUF/resolve/main/embeddinggemma-300m-qat-Q8_0.gguf"
SHA256 = "6fa0c02a9c302be6f977521d399b4de3a46310a4f2621ee0063747881b673f67"


def main():
    dest = ROOT / "models/embedding.gguf"
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        with dest.open("rb") as f:
            actual = hashlib.file_digest(f, "sha256").hexdigest()
        if actual != SHA256:
            raise RuntimeError(
                "Existing model checksum does not match; inspect it before replacing"
            )
        print("Existing embedding model verified")
        return
    partial = dest.with_suffix(".download")
    request = urllib.request.Request(
        URL, headers={"User-Agent": "OutlookMCP-installer/1"}
    )
    with (
        urllib.request.urlopen(request, timeout=60) as source,
        partial.open("wb") as target,
    ):
        while block := source.read(1024 * 1024):
            target.write(block)
    with partial.open("rb") as f:
        actual = hashlib.file_digest(f, "sha256").hexdigest()
    if actual != SHA256:
        raise RuntimeError("Downloaded model checksum mismatch; file was not activated")
    os.replace(partial, dest)
    print("Embedding model downloaded and verified")


if __name__ == "__main__":
    main()
