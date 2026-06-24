import os
import json
from groq import Groq
from dotenv import load_dotenv

load_dotenv()

# =========================
# CONFIG
# =========================

CHUNKS_DIR = "local_cosmos_db/chunks"
OUTPUT_DIR = "local_cosmos_db"

os.makedirs(OUTPUT_DIR, exist_ok=True)

client = Groq(
    api_key=os.getenv("GROQ_API_KEY")
)

MODEL = "llama-3.3-70b-versatile"

# =========================
# PROMPT
# =========================

SYSTEM_PROMPT = """
You are an expert knowledge graph extraction system.

Extract entities and relationships from the provided lease text.

Return ONLY valid JSON.

Rules:

1. Entity IDs must be lowercase.
2. Replace spaces with underscores.
3. Use stable IDs.

Examples:

Tenant -> tenant
Landlord -> landlord
5 Day Grace Period -> grace_period_5_days
$250 Late Fee -> late_fee_250

Return exactly:

{
  "nodes": [
    {
      "id": "",
      "name": "",
      "type": ""
    }
  ],
  "edges": [
    {
      "source": "",
      "relationship": "",
      "target": ""
    }
  ]
}

No markdown.
No explanations.
No extra text.
"""

# =========================
# EXTRACTION FUNCTION
# =========================

def extract_graph(text):

    response = client.chat.completions.create(
        model=MODEL,
        temperature=0,
        messages=[
            {
                "role": "system",
                "content": SYSTEM_PROMPT
            },
            {
                "role": "user",
                "content": text
            }
        ]
    )

    content = response.choices[0].message.content.strip()

    if content.startswith("```json"):
        content = content.replace("```json", "")
        content = content.replace("```", "").strip()

    return json.loads(content)

# =========================
# MAIN
# =========================

def main():

    all_nodes = {}
    all_edges = []

    files = [
        f for f in os.listdir(CHUNKS_DIR)
        if f.endswith(".json")
    ]

    print(f"Found {len(files)} chunks")

    for file in files:

        chunk_path = os.path.join(CHUNKS_DIR, file)

        try:

            with open(chunk_path, "r", encoding="utf-8") as f:
                chunk = json.load(f)

            print(f"Processing {chunk['chunk_id']}")

            graph_data = extract_graph(
                chunk["text"]
            )

            # -------------------------
            # Nodes
            # -------------------------

            for node in graph_data.get("nodes", []):

                node["chunk_id"] = chunk["chunk_id"]
                node["document_id"] = chunk["document_id"]

                node_id = node["id"]

                if node_id not in all_nodes:
                    all_nodes[node_id] = node

            # -------------------------
            # Edges
            # -------------------------

            for edge in graph_data.get("edges", []):

                edge["chunk_id"] = chunk["chunk_id"]
                edge["document_id"] = chunk["document_id"]

                all_edges.append(edge)

        except Exception as e:

            print(
                f"Failed on {file}: {str(e)}"
            )

    # =========================
    # SAVE NODES
    # =========================

    nodes_output = os.path.join(
        OUTPUT_DIR,
        "nodes.json"
    )

    with open(
        nodes_output,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            list(all_nodes.values()),
            f,
            indent=4
        )

    # =========================
    # SAVE EDGES
    # =========================

    edges_output = os.path.join(
        OUTPUT_DIR,
        "edges.json"
    )

    with open(
        edges_output,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            all_edges,
            f,
            indent=4
        )

    print()
    print("=" * 50)
    print(f"Nodes: {len(all_nodes)}")
    print(f"Edges: {len(all_edges)}")
    print("Saved nodes.json")
    print("Saved edges.json")
    print("=" * 50)


if __name__ == "__main__":
    main()