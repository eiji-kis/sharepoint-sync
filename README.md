

# 🚀 SharePoint Sync CLI

## 🌟 Overview
**SharePoint Sync CLI** is a command-line tool for synchronizing files between two SharePoints (e.g., **KIS** and **Client SharePoints**). It provides powerful features such as:

- Excluding specific files or directories
- Detecting moved or updated files
- Displaying diffs for **Word documents** (`.docx`)
- Interactive copy/move operations
- Managing **multiple profiles** for different SharePoints

---

## 🔧 Requirements
Make sure you have the following installed before using the CLI:

🔹 **Python 3**
🔹 Required Python packages:
  - `python-docx`
  - `pyfiglet`
  - `click`
  - `colorlog`

📌 **Install dependencies with:**
```sh
pip install python-docx pyfiglet click colorlog
```

---

## 📥 Installation
Clone the repository and navigate to the project directory:
```sh
git clone https://github.com/eiji-kis/sharepoint-sync.git
cd sharepoint-sync
```
Make the script executable (optional):
```sh
chmod +x sharepoint_sync.py
```

---

## 🛠️ Usage
Run the script using:
```sh
python sharepoint_sync.py <command> [options]
```

or, if you make the script executable:
```sh
./sharepoint_sync.py <command> [options]
```

### 🔹 Available Commands

#### 1️⃣ **Setup a Profile**

> Before setting up a profile, make sure you have OneDrive installed and both SharePoints are avaliable locally on your machine

Before syncing files, **set up a profile**:
```sh
python sharepoint_sync.py setup
```
You will be prompted to enter:
- 📌 **Profile name** (e.g., `client_name`)
- 📌 **Root directory** of KIS SharePoint
- 📌 **Root directory** of the Client SharePoint

This stores your profile configuration in `~/.sharepoint_sync_profiles.json`.

#### 2️⃣ **Sync Files**

Before synchronizing two SharePoints, the files on your local should be up to date.

To achieve this, hit the **Sync** button on SharePoint web to pull in changes others may have made to the documents:

![alt text](/img/sync.png)

To **sync files** between SharePoints:
```sh
python sharepoint_sync.py sync <profile>
```
✨ This will:
- Identify **missing files** in each directory
- Detect **moved or updated** files
- **Prompt for interactive actions**
- Display **diffs** for `.docx` files 📄

#### 3️⃣ **Exclude Directories**
To **exclude a directory** from the sync process:
```sh
python sharepoint_sync.py exclude_dir <dir_name>
```

#### 4️⃣ **Exclude Files**
To **exclude a file** from the sync process:
```sh
python sharepoint_sync.py exclude_file <file_name>
```

---

## ⚡ Features
- **Profile Management** – Configure multiple SharePoint pairs for different clients.
- **Diff Viewer** – Shows content differences in `.docx` files.
- **Interactive Actions** – Prompt-based confirmations for moving or copying files.
- **Exclusion Support** – Exclude specific files or directories from sync.
- **Logging** – Colorized logging for better visibility. 🎨

---

## 📝 Logging & Follow-Ups
🔹 **All actions and skipped operations are logged.**
🔹 **Follow-ups** for manual review are stored in `follow_up_tasks.md`. 📜

---

## 🎯 Example Workflow
```sh
# Set up a profile
python sharepoint_sync.py setup

# Sync files for a profile
python sharepoint_sync.py sync client_name

# Exclude a directory
python sharepoint_sync.py exclude_dir "old_files"

# Exclude a file
python sharepoint_sync.py exclude_file "confidential.docx"
```

---

## ⚡ Managing Excluded Dirs and Files
To view all **excluded directories and files**, check your configuration file:
```sh
cat ~/.sharepoint_sync_profiles.json
```

To **remove a directory or file** from exclusion, edit the JSON file manually or override it using:
```sh
python sharepoint_sync.py exclude_dir <dir_name> --remove
python sharepoint_sync.py exclude_file <file_name> --remove
```

You can also reset the exclusion list by running:
```sh
python sharepoint_sync.py reset_exclusions
```

---

## 🛠️ Troubleshooting
⚠️ If an error occurs while reading `.docx` files, **ensure `python-docx` is installed.**
⚠️ **Check that both SharePoint directories exist before syncing.**
⚠️ Run with `--verbosity debug` to see **detailed logs**:
```sh
python sharepoint_sync.py sync <profile> --verbosity debug
```

---

## 🤝 Contributions
Pull requests and feature suggestions are **welcome**! 🚀 Feel free to open an issue if you encounter any bugs or need improvements.