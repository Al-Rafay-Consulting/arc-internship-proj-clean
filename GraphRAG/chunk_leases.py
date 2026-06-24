import os
import json
import re

DOCUMENTS_DIR = "local_cosmos_db/documents"
CHUNKS_DIR = "local_cosmos_db/chunks"

os.makedirs(CHUNKS_DIR, exist_ok=True)

SECTION_PATTERN = re.compile(
    r'([A-Za-z ]+):\n(.*?)(?=\n[A-Za-z ]+:\n|\Z)',
    re.DOTALL
)


def create_chunks(document):
    document_id = document["id"]
    text = document["full_text"]

    matches = SECTION_PATTERN.findall(text)

    chunks = []

    for idx, (section, content) in enumerate(matches, start=1):

        chunk = {
            "chunk_id": f"{document_id}_chunk_{idx}",
            "document_id": document_id,
            "section": section.strip(),
            "text": content.strip()
        }

        chunks.append(chunk)

    return chunks


def process_document(filepath):

    with open(filepath, "r", encoding="utf-8") as f:
        document = json.load(f)

    chunks = create_chunks(document)

    for chunk in chunks:

        output_file = os.path.join(
            CHUNKS_DIR,
            f"{chunk['chunk_id']}.json"
        )

        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(chunk, f, indent=4)

    print(
        f"{document['id']} -> {len(chunks)} chunks created"
    )


def main():

    for file in os.listdir(DOCUMENTS_DIR):

        if file.endswith(".json"):

            process_document(
                os.path.join(DOCUMENTS_DIR, file)
            )


if __name__ == "__main__":
    main()