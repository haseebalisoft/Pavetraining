import csv
from collections import defaultdict

fields = defaultdict(list)
with open("pave-schema-export/sharepoint-fields.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        if r["Hidden"] == "True":
            continue
        fields[r["ListTitle"]].append(r)

checks = {
    "Workforce List": ["TrainingManager", "Trainingmanager", "CandidateName", "CompanyName"],
    "NPORS Register": [
        "NPORSNumber", "On_x002f_Number", "TrainingOutcome", "CustomerVisible",
        "AssessorTrainer", "OutcomeDate", "OutcomeNotes", "CandidateName", "CompanyName",
        "NPORSCategory", "Expiry", "NoviceorEwt", "Notes", "TESTER",
    ],
    "EUSR Register": [
        "CompanyName", "EUSRNumber", "TrainingAddress", "TrainingOutcome",
        "CustomerVisible", "AssessorTrainer", "OutcomeDate", "OutcomeNotes",
        "EusrCategory", "CardType", "Expiry", "CandidateName", "BatchNumber",
    ],
    "NRSWA Register": [
        "CompanyName", "SWQRNumber", "TrainingAddress", "TrainingOutcome",
        "CustomerVisible", "AssessorTrainer", "OutcomeDate", "OutcomeNotes",
        "StreetworksCategory", "Course", "Expirydate", "CandidateName",
    ],
    "In-House Certificates Register": [
        "CourseCategory", "CertificateCategory", "TrainingAddress", "TrainingOutcome",
        "CustomerVisible", "AssessorTrainer", "OutcomeDate", "OutcomeNotes",
        "CourseDate", "ExpiryDate", "CUSTOMER", "CandidateName", "CompanyName",
    ],
    "NVQ Register": [
        "CustomerVisible", "TrainingOutcome", "AssessorTrainer", "OutcomeDate",
        "OutcomeNotes", "NvqTitle", "BoltonNvq", "StageofNvq", "CustomerUpdateNotes",
        "Internal_x0020_Notes", "Notes",
    ],
    "Events": [
        "EventCompany", "Company", "Customer_x0020_Visible", "CustomerVisible",
        "TrainingAddress", "OutlookEventId", "OutlookCalendarId", "OutlookICalUid",
        "SyncStatus", "SyncDirection", "LastSyncedAt", "LastSyncSource", "SyncError",
        "DoNotSync", "SyncHash", "EventDate", "EndDate", "Location", "Description",
        "Trainer",
    ],
    "Permissions List": [
        "CompanyLookupId", "Company", "UserEmail", "RoleType", "Status",
        "AccessScope", "CanView", "CanDownload", "CanEdit", "Departments",
        "DepartmentsAllowed", "ReceiveExpiryNotifications",
    ],
    "Training Course Categories": None,
}

for list_name, wanted in checks.items():
    print("====", list_name)
    allf = fields.get(list_name, [])
    by_lower = {r["InternalName"].lower(): r for r in allf}
    if wanted is None:
        for r in allf:
            if r["InternalName"] in ("ContentType", "Attachments", "Edit", "LinkTitleNoMenu", "LinkTitle", "SelectTitle", "InstanceID", "Order", "GUID", "WorkflowVersion", "FileLeafRef", "UniqueId", "SyncClientId", "ProgId", "ScopeId", "File_x0020_Type", "MetaInfo", "HTML_x0020_File_x0020_Type", "FileRef", "FileDirRef", "Last_x0020_Modified", "Created_x0020_Date", "FSObjType", "PermMask", "PrincipalCount", "Restricted", "OriginatorId", "NoExecute", "ContentVersion", "SortBehavior", "DocIcon", "ServerUrl", "EncodedAbsUrl", "BaseName", "FileSizeDisplay", "LinkFilenameNoMenu", "LinkFilename", "ItemChildCount", "FolderChildCount", "AppAuthor", "AppEditor", "ComplianceAssetId", "owshiddenversion", "ID", "Modified", "Created", "Author", "Editor"):
                continue
            print(f"  {r['InternalName']} | {r['FieldTitle']} | {r['TypeAsString']} | Hidden={r['Hidden']}")
    else:
        for w in wanted:
            hit = by_lower.get(w.lower())
            if hit:
                same = hit["InternalName"] == w
                print(f"  {'OK' if same else 'CASE'} want={w} found={hit['InternalName']} type={hit['TypeAsString']} title={hit['FieldTitle']}")
            else:
                print(f"  MISS want={w}")
        print("  -- ALL NON-SYSTEM --")
        skip_prefix = ("_",)
        skip = {"ContentType", "Attachments", "Edit", "LinkTitleNoMenu", "LinkTitle", "SelectTitle", "InstanceID", "Order", "GUID", "WorkflowVersion", "FileLeafRef", "UniqueId", "SyncClientId", "ProgId", "ScopeId", "File_x0020_Type", "MetaInfo", "HTML_x0020_File_x0020_Type", "FileRef", "FileDirRef", "Last_x0020_Modified", "Created_x0020_Date", "FSObjType", "PermMask", "PrincipalCount", "Restricted", "OriginatorId", "NoExecute", "ContentVersion", "SortBehavior", "DocIcon", "ServerUrl", "EncodedAbsUrl", "BaseName", "FileSizeDisplay", "LinkFilenameNoMenu", "LinkFilename", "ItemChildCount", "FolderChildCount", "AppAuthor", "AppEditor", "ComplianceAssetId", "owshiddenversion", "ID", "Modified", "Created", "Author", "Editor", "ContentTypeId", "_UIVersionString", "LinkTitle"}
        for r in allf:
            if r["InternalName"] in skip or r["InternalName"].startswith("_"):
                continue
            print(f"  {r['InternalName']} | {r['FieldTitle']} | {r['TypeAsString']}")
    print()
