"""
autostart.py – Legger Notater til Windows-oppstart og lager skrivebords-snarvei.
"""

import os, sys, winreg
from pathlib import Path


def add_to_windows_startup():
    """Legger start.bat til Windows-oppstart via register."""
    notater_dir = Path(__file__).parent.parent
    start_bat   = notater_dir / "start.bat"

    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0, winreg.KEY_SET_VALUE
        )
        winreg.SetValueEx(key, "Notater", 0, winreg.REG_SZ, str(start_bat))
        winreg.CloseKey(key)
        print(f"[Autostart] Lagt til i Windows-oppstart: {start_bat}")
        return True
    except Exception as e:
        print(f"[Autostart] Registerfeil: {e}")
        return False


def create_desktop_shortcut():
    """Lager .lnk-snarvei på skrivebordet."""
    import win32com.client as win32

    notater_dir = Path(__file__).parent.parent
    start_bat   = notater_dir / "start.bat"
    icon_file   = notater_dir / "assets" / "icon.ico"

    # Finn skrivebord
    desktop = Path(os.environ.get("USERPROFILE", "")) / "OneDrive - Osloskolen" / "Skrivebord"
    if not desktop.exists():
        desktop = Path(os.environ.get("USERPROFILE", "")) / "Desktop"

    lnk_path = desktop / "Notater.lnk"

    try:
        shell = win32.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(str(lnk_path))
        shortcut.TargetPath       = str(start_bat)
        shortcut.WorkingDirectory = str(notater_dir)
        shortcut.Description      = "Notater – Matteassistent"
        if icon_file.exists():
            shortcut.IconLocation = str(icon_file)
        shortcut.Save()
        print(f"[Autostart] Snarvei opprettet: {lnk_path}")
        return True
    except Exception as e:
        print(f"[Autostart] Snarvei-feil: {e}")
        return False


def run():
    print("\n=== Setter opp autostart og snarvei ===")
    add_to_windows_startup()
    create_desktop_shortcut()
    print("[Autostart] Ferdig.")


if __name__ == "__main__":
    run()
