# SharePoint Sync CLI

A command-line tool to manually **sync** files between two SharePoint directories:  
- **KIS SharePoint** (internal)
- **Client SharePoint** (external)

It includes features such as:

✅ **Syncing missing or outdated files**  
✅ **Checking for differences between docx files**  
✅ **Skipping unchanged files**  
✅ **Excluding specific files and directories from sync**  

## 📌 **Installation**

### **1️⃣ Clone the Repository**
```sh
git clone TODO
cd sharepoint-sync
```

### **2️⃣ Install Dependencies**
This script requires Python 3.7 or later. Install dependencies with:
```sh
pip install click python-docx
```

## 🚀 **Usage Guide**

### **1️⃣ Setting Up a Profile**
A profile stores the paths for your **KIS SharePoint** and **Client SharePoint**.

Run the following command and provide the required paths:
```sh
python sharepoint_sync.py setup
```
You’ll be asked:
- **Profile name** (e.g., `client_project_A`)
- **KIS SharePoint directory** (e.g., `/Users/username/SharePoint/KIS`)
- **Client SharePoint directory** (e.g., `/Users/username/SharePoint/Client`)

### **2️⃣ Syncing Files**
To sync files between SharePoints, run:
```sh
python sharepoint_sync.py sync client_project_A
```
This will:
- **Compare files** in both SharePoint directories.
- **Skip unchanged files** (same modification date).
- **Show differences** for modified files.
- **Prompt to copy** newer files from one directory to the other.

If a file is missing in one directory, you'll be asked:
```
File missing in Client SharePoint: report.docx. Copy from KIS SharePoint? [y/N]
```

If a file is updated, you’ll see:
```
File exists in both locations:
  KIS SharePoint last modified: 2025-02-01 14:30:00
  Client SharePoint last modified: 2025-01-29 10:15:00
Differences detected
Copy newer version from KIS SharePoint to Client SharePoint? [y/N]
```

---

## ⚠️ **Excluding Files and Folders**

### **1️⃣ Exclude a Specific File**
To permanently ignore a file from syncing, add it to the ignore list:
```sh
python sharepoint_sync.py exclude-file .DS_Store
```
Now, any file named `.DS_Store` will be ignored.

### **2️⃣ Exclude a Directory**
To ignore an entire folder (and its subfolders), use:
```sh
python sharepoint_sync.py exclude-dir node_modules
```
Now, any folder named `node_modules` inside your SharePoints **will be skipped**.

---

## 🔄 **Managing Exclusions**
You can manually check or modify the **exclusion lists** in:
```
~/.sharepoint_sync_profiles.json
```
This JSON file stores:
- **Profiles** (SharePoint paths)
- **Excluded files**
- **Excluded directories**

Example:
```json
{
    "profiles": {
        "client_project_A": {
            "kis_dir": "/Users/username/SharePoint/KIS",
            "client_dir": "/Users/username/SharePoint/Client"
        }
    },
    "excluded_files": [".DS_Store"],
    "excluded_dirs": ["controlled"]
}
```

---

## 🛠 **Troubleshooting**
- **Error: `FileNotFoundError`**  
  Ensure your SharePoint folders exist before running `sync`.  
- **Error: `ValueError` (Path issue)**  
  This happens if your files exist outside the expected directories. Run `exclude-file` or `exclude-dir` to ignore them.  
- **Skipping files but you want them included?**  
  Remove them from `excluded_files` or `excluded_dirs` in the JSON config.  