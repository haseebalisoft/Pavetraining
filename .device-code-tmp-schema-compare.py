import json, csv, re
from pathlib import Path

schema = Path("src/lib/schema/sharepointSchema.ts").read_text(encoding="utf-8")
field_blocks = {}
for m in re.finditer(r"const (\w+)Fields = \{([^}]+)\}", schema, re.S):
    name = m.group(1)
    fields = {}
    for fm in re.finditer(r'(\w+):\s*"([^"]+)"', m.group(2)):
        fields[fm.group(1)] = fm.group(2)
    field_blocks[name] = fields

list_map = {
    "company": "Company List",
    "workforce": "Workforce List",
    "trainingMatrix": "Training Matrix",
    "nporsRegister": "NPORS Register",
    "eusrRegister": "EUSR Register",
    "nrswaRegister": "NRSWA Register",
    "inHouseCertificates": "In-House Certificates Register",
    "nvqRegister": "NVQ Register",
    "customerDocuments": "Customer Documents",
    "events": "Events",
    "offersPromotions": "Offers / Promotions",
    "permissions": "Permissions List",
    "trainingCourseCategories": "Training Course Categories",
    "trainingManagerLogs": "Training Manager Logs",
}

d = json.load(open("pave-schema-export/sharepoint-schema.json", encoding="utf-8"))
export_lists = {L["ListTitle"]: L for L in d["Lists"]}

csv_fields = {}
with open("pave-schema-export/sharepoint-fields.csv", encoding="utf-8") as f:
    for row in csv.DictReader(f):
        csv_fields.setdefault(row["ListTitle"], []).append(row)

print("EXPORT LISTS:", len(export_lists))
print("CODE LISTS:", len(list_map))
print()
print("=== Lists in export NOT in code schema ===")
code_names = set(list_map.values())
for t in sorted(export_lists):
    L = export_lists[t]
    if t not in code_names and str(L.get("Hidden", "")).lower() != "true":
        print(f"  {t} | {L.get('ListId')} | template={L.get('BaseTemplate')} items={L.get('ItemCount')}")

print()
print("=== Lists in code NOT in export ===")
for k, v in list_map.items():
    if v not in export_lists:
        print(" ", v)

print()
print("=== Field gaps ===")
for key, list_name in list_map.items():
    code_fields = field_blocks.get(key, {})
    export_internals = {r["InternalName"] for r in csv_fields.get(list_name, [])}
    missing = []
    for prop, internal in code_fields.items():
        if internal == "ID":
            continue
        if internal not in export_internals:
            missing.append(f"{prop}={internal}")

    code_internals = set(code_fields.values())
    extras = []
    for r in csv_fields.get(list_name, []):
        inn = r["InternalName"]
        if inn.startswith("_"):
            continue
        if r.get("FromBaseType") == "True" or r.get("Hidden") == "True":
            continue
        if inn in ("ID", "Title", "Modified", "Created", "Author", "Editor", "Attachments", "ContentType", "ContentTypeId", "Edit", "LinkTitleNoMenu", "LinkTitle", "SelectTitle", "InstanceID", "Order", "GUID", "WorkflowVersion", "FileLeafRef", "UniqueId", "SyncClientId", "ProgId", "ScopeId", "File_x0020_Type", "MetaInfo", "HTML_x0020_File_x0020_Type", "FileRef", "FileDirRef", "Last_x0020_Modified", "Created_x0020_Date", "FSObjType", "PermMask", "PrincipalCount", "Restricted", "OriginatorId", "NoExecute", "ContentVersion", "SortBehavior", "DocIcon", "ServerUrl", "EncodedAbsUrl", "BaseName", "FileSizeDisplay", "LinkFilenameNoMenu", "LinkFilename", "ItemChildCount", "FolderChildCount", "AppAuthor", "AppEditor", "ComplianceAssetId", "owshiddenversion", "_UIVersion", "_UIVersionString", "SMTotalSize", "SMLastModifiedDate", "SMTotalFileStreamSize", "SMTotalFileCount", "ParentUniqueId", "ParentVersionString", "ParentLeafName", "DocConcurrencyNumber", "StreamHash", "AccessPolicy"):
            continue
        if inn not in code_internals:
            extras.append(f"{r.get('Title')}={inn} type={r.get('TypeAsString')}")

    print(f"\n[{list_name}] id={export_lists.get(list_name, {}).get('ListId', '?')}")
    if missing:
        print("  MISSING IN EXPORT:", ", ".join(missing))
    else:
        print("  All code internal names found in export")
    if extras:
        print("  IN EXPORT NOT IN CODE:")
        for e in extras:
            print("   -", e)
