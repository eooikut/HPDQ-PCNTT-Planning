import json, os, shutil
# === Utils ===
METADATA_FILE = "metadata.json"
def load_metadata():
    if os.path.exists(METADATA_FILE):
        with open(METADATA_FILE, "r") as f:
            return json.load(f)
    return []

def save_metadata(data):
    with open(METADATA_FILE, "w") as f:
        json.dump(data, f, indent=2)
def get_lsx_by_id(lsx_id):
    metadata = load_metadata()
    for d in metadata:
        if d.get("type") == "lsx" and d.get("id") == lsx_id:
            return d
    return None
def update_metadata(lsx_id, new_entry):
    metadata = load_metadata()
    updated = []
    for d in metadata:
        if d.get("type") == "lsx" and d.get("id") == lsx_id:
            updated.append(new_entry)
        else:
            updated.append(d)
    save_metadata(updated)

def delete_metadata(lsx_id):
    metadata = load_metadata()
    metadata = [lsx for lsx in metadata if lsx["id"] != lsx_id]
    save_metadata(metadata)
    # xóa folder chứa file
    folder = os.path.join("uploads", lsx_id)
    if os.path.exists(folder):
        shutil.rmtree(folder)
