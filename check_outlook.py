"""Comprehensive Outlook diagnostic"""
import sys
import os
import subprocess

print("=" * 70)
print("  OUTLOOK INSTALLATION DIAGNOSTIC")
print("=" * 70)

# Check 1: Outlook executable
print("\n[1] Checking for Outlook executable...")
outlook_paths = [
    r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
    r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
    r"C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE",
]

found_outlook = None
for path in outlook_paths:
    if os.path.exists(path):
        print(f"  ✓ Found: {path}")
        found_outlook = path
        break

if not found_outlook:
    print("  ✗ Outlook executable NOT found in standard locations")
    print("\n  Checking if Outlook is in PATH...")
    try:
        result = subprocess.run(["where", "outlook.exe"], 
                              capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            print(f"  ✓ Found in PATH: {result.stdout.strip()}")
            found_outlook = result.stdout.strip().split('\n')[0]
        else:
            print("  ✗ Not found in PATH either")
    except:
        print("  ✗ Could not check PATH")

# Check 2: Windows Store version warning
print("\n[2] Checking for Windows Store Outlook...")
store_outlook = r"C:\Program Files\WindowsApps"
if os.path.exists(store_outlook):
    try:
        apps = os.listdir(store_outlook)
        outlook_apps = [a for a in apps if 'outlook' in a.lower()]
        if outlook_apps:
            print("  ⚠ WARNING: Found Windows Store Outlook:")
            for app in outlook_apps:
                print(f"    - {app}")
            print("  ⚠ Windows Store Outlook does NOT support COM automation!")
            print("  ⚠ You need the desktop version from Microsoft 365/Office")
    except PermissionError:
        print("  (Cannot check - permission denied)")

# Check 3: pywin32
print("\n[3] Checking pywin32...")
try:
    import win32com.client
    print("  ✓ pywin32 is installed")
except ImportError:
    print("  ✗ pywin32 NOT installed")
    sys.exit(1)

# Check 4: Try COM connection
print("\n[4] Testing COM connection to Outlook.Application...")
try:
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    print(f"  ✓ SUCCESS! Connected to Outlook")
    try:
        version = outlook.Version
        print(f"  ✓ Outlook version: {version}")
    except:
        print("  ✓ Connected but cannot get version")
    outlook = None
except Exception as e:
    print(f"  ✗ FAILED: {e}")
    error_code = str(e)
    if "-2147221005" in error_code or "Invalid class string" in error_code:
        print("\n  This error means:")
        print("    • Outlook is not installed, OR")
        print("    • You have Windows Store Outlook (unsupported), OR")
        print("    • Outlook's COM registration is broken")

# Check 5: Try EnsureDispatch
print("\n[5] Testing EnsureDispatch method...")
try:
    import win32com.client
    outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    print("  ✓ EnsureDispatch worked!")
    outlook = None
except Exception as e:
    print(f"  ✗ EnsureDispatch failed: {e}")

# Check 6: Registry check
print("\n[6] Checking Windows Registry for Outlook...")
try:
    import winreg
    try:
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Outlook.Application")
        print("  ✓ Outlook.Application is registered in Windows Registry")
        winreg.CloseKey(key)
    except FileNotFoundError:
        print("  ✗ Outlook.Application NOT found in Registry")
        print("     This confirms Outlook COM is not registered!")
except Exception as e:
    print(f"  ✗ Could not check registry: {e}")

print("\n" + "=" * 70)
print("  SUMMARY")
print("=" * 70)

if found_outlook:
    print(f"✓ Outlook executable: {found_outlook}")
else:
    print("✗ Outlook executable: NOT FOUND")

print("\nRecommendations:")
print("1. If you have Windows Store Outlook → Uninstall it and install")
print("   Microsoft 365/Office desktop version")
print("2. If Outlook is installed → Try repairing Office installation")
print("3. Alternative → Use SMTP with OAuth2 (complex)")
print("4. Alternative → Use mailto: to create draft (manual send)")
print("=" * 70)
