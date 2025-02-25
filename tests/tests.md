# SharePoint Sync Command - Test Plan

## 📌 **1. Command Invocation Tests**
### **1.1 Verify Command Runs Without Errors**
- Run: `python sync.py profile_name`
- Expected: Command executes successfully without crashes.

> ✅ passed

###  **1.2 Verify `--verbosity` Option**
- Run with `--verbosity debug`
- Expected: Debug logs appear.
- Run with `--verbosity info`
- Expected: Only important logs appear.

> ✅ passed

---

## 📌 **2. Profile & Configuration Tests**
###  **2.1 Profile Exists**
- Run: `python sync.py non_existent_profile`
- Expected: `"Profile 'non_existent_profile' not found. Run 'setup' first."`

> ✅ passed

###  **2.2 Load Profiles Correctly**
- Ensure `profiles.json` loads and contains valid SharePoint paths.

> ✅ passed

###  **2.3 Handle Missing Configuration Keys**
- Test missing keys (`kis_dir`, `client_dir`, etc.)
- Expected: Graceful error handling.

> ✅ passed

---

## 📌 **3. Directory & File Comparison Tests**
###  **3.1 Handle Empty Directories**
- If both SharePoints are empty, log `"KIS and <profile> SharePoints are already synched 😎"`

> ✅ passed

###  **3.2 Detect New Files in KIS**
- Create a new file in `kis_dir` but not in `client_dir`.
- Run sync.
- Expected: Prompt user to copy file to `client_dir`.

> ✅ passed

###  **3.3 Detect New Files in Client SharePoint**
- Create a new file in `client_dir` but not in `kis_dir`.
- Run sync.
- Expected: Prompt user to copy file to `kis_dir`.

> ✅ passed

###  **3.4 Detect Moved Files**
- Move a file within `kis_dir` or `client_dir`.
- Run sync.
- Expected: Log file movement, prompt user to move it in the other SharePoint.

> ✅ passed

###  **3.5 Detect Updated Files**
- Modify a file in `kis_dir` but keep `client_dir` version unchanged.
- Run sync.
- Expected: Prompt user to update the outdated file.

> ✅ passed

---

## 📌 **4. File Synchronization Tests**
###  **4.1 Copy New Files**
- Run sync and confirm copying a new file.
- Expected: File appears in target directory with same content.

> ✅ passed

###  **4.2 Move Files Across SharePoints**
- Run sync and confirm moving a file.
- Expected: File path updates in both SharePoints.

> ✅ passed

###  **4.3 Overwrite Updated Files**
- Modify a file and confirm sync overwrites outdated version.
- Expected: Target file gets replaced.

> ✅ passed

###  **4.4 Handle DOCX File Differences**
- Modify a `.docx` file and run sync.
- Expected: `show_file_diff()` should display changes.

> ✅ passed

---

## 📌 **5. Exclusion Rules Tests**
###  **5.1 Ignore Excluded Files**
- Add a file to `excluded_files` in `profiles.json`.
- Run sync.
- Expected: The file should not appear in sync operations.

> ✅ passed

###  **5.2 Ignore Excluded Directories**
- Add a folder to `excluded_dirs` in `profiles.json`.
- Run sync.
- Expected: The folder should be ignored.

> ✅ passed

---

## 📌 **6. Edge Case & Error Handling Tests**

###  **6.1 Handle File Permission Issues**
- Restrict write permissions in `client_dir` or `kis_dir`.
- Run sync.
- Expected: Log permission error.

> ✅ passed

###  **6.2 Handle Read-Only Files**
- Mark a file as read-only and run sync.
- Expected: Log a warning and prompt user.

> ✅ passed
