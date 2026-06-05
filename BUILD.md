# Building the Standalone Executable (.exe)

This guide explains how to compile the graphical user interface (`gui.py`) and automation script (`automate.py`) into a single, standalone executable that can be run on any computer without needing Python or manual setups.

---

## 💻 Building for Windows (.exe)

Since PyInstaller compiles executables for the host operating system, **you must run these commands on a Windows machine** to generate the `.exe` file.

### Step 1: Install Python
Ensure Python 3.10+ is installed on the Windows machine. (Make sure to check the box "Add Python to PATH" during installation).

### Step 2: Install PyInstaller and Dependencies
Open Command Prompt (`cmd`) in this project folder and run:
```cmd
pip install -r requirements.txt
pip install pyinstaller
```

*Note: If `requirements.txt` is not present, manually install the dependencies:*
```cmd
pip install playwright openpyxl
```

### Step 3: Compile using PyInstaller
Run the following packaging command in your terminal:
```cmd
pyinstaller --onefile --noconsole --collect-all playwright --name "SIDH_Automator" gui.py
```

#### What this command does:
- `--onefile`: Packages everything into a single, clean `.exe` file.
- `--noconsole`: Hides the command prompt window, launching only the beautiful GUI.
- `--collect-all playwright`: Bundles all internal Playwright drivers and configurations.
- `--name "SIDH_Automator"`: Sets the executable name.

### Step 4: Access your Software
Once the compilation completes, you will find the standalone executable inside the newly created **`dist`** folder:
`dist/SIDH_Automator.exe`

---

## 🐧 Building for Linux (Binary)

To compile the application into a standalone binary for Linux:

1. Install PyInstaller in your environment:
   ```bash
   pip install pyinstaller
   ```
2. Compile:
   ```bash
   pyinstaller --onefile --noconsole --collect-all playwright --name "SIDH_Automator" gui.py
   ```
3. Run the binary:
   ```bash
   ./dist/SIDH_Automator
   ```

---

## 🌟 Built-in Premium Features

The packaged application is designed to be user-friendly for resale:
1. **Auto-Chromium Installer**: On the first launch, the software will automatically detect if Chromium is missing and install it in the background dynamically without user intervention.
2. **Dynamic File Picker**: Users can easily browse and load their Excel sheet.
3. **Smart Parsing**: The app auto-extracts the 7-digit Batch ID from the Excel filename.
4. **Interactive Logs**: The embedded terminal in the GUI displays execution logs in real-time.
5. **Session Retention**: Logins are cached locally inside batch-specific directories (`.playwright_session_{batch_id}`) so users don't have to log in on every run for a batch, and multiple batches can run concurrently without session lock conflicts.
