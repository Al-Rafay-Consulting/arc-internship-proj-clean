import os
import json
import re

SOURCE_DIR = "local_cosmos_db/source_files"
OUTPUT_DIR = "local_cosmos_db/documents"

os.makedirs(OUTPUT_DIR, exist_ok=True)


def extract_lease_id(text):
    match = re.search(r"Lease ID:\s*([A-Za-z0-9\-]+)", text)
    if match:
        return match.group(1)
    return None


def ingest_file(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        text = f.read()

    filename = os.path.basename(filepath)

    lease_id = extract_lease_id(text)

    document = {
        "id": lease_id if lease_id else filename.replace(".txt", ""),
        "source_file": filename,
        "document_type": "lease",
        "full_text": text
    }

    output_path = os.path.join(
        OUTPUT_DIR,
        f"{document['id']}.json"
    )

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(document, f, indent=4)

    print(f"Ingested {filename}")


def main():
    for file in os.listdir(SOURCE_DIR):
        if file.endswith(".txt"):
            ingest_file(os.path.join(SOURCE_DIR, file))


if __name__ == "__main__":
    main()