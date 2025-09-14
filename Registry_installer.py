import winreg
import os
import tkinter as tk
from tkinter import messagebox


def trouver_bat():
    """Tente de localiser ResumeNotion.bat automatiquement"""
    chemins_possibles = []

    # 1. Même dossier que le script
    chemins_possibles.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "ResumeNotion.bat"))

    # 2. Dossier courant
    chemins_possibles.append(os.path.join(os.getcwd(), "ResumeNotion.bat"))

    # 3. Bureau & Documents
    userprofile = os.environ.get("USERPROFILE", "")
    if userprofile:
        chemins_possibles.append(os.path.join(userprofile, "Desktop", "ResumeNotion.bat"))
        chemins_possibles.append(os.path.join(userprofile, "Documents", "ResumeNotion.bat"))

    # Vérifie chaque chemin
    for chemin in chemins_possibles:
        if os.path.exists(chemin):
            return chemin

    return None


def installer():
    """Installe le protocole si ResumeNotion.bat est trouvé"""
    chemin = trouver_bat()

    if not chemin:
        messagebox.showerror("Erreur", "Impossible de trouver ResumeNotion.bat automatiquement !")
        return

    program_path = os.path.normpath(chemin)
    protocol_name = "ResumeNotion"

    try:
        # Création de la clé principale
        key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, protocol_name)
        winreg.SetValueEx(key, None, 0, winreg.REG_SZ, "Resume Notion")
        winreg.SetValueEx(key, "URL Protocol", 0, winreg.REG_SZ, "")
        winreg.CloseKey(key)

        # Actions (open, edit, print)
        actions = ["open", "edit", "print"]
        for action in actions:
            command_key_path = f"{protocol_name}\\shell\\{action}\\command"
            command_key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, command_key_path)
            winreg.SetValueEx(command_key, None, 0, winreg.REG_SZ, f'"{program_path}" "%1"')
            winreg.CloseKey(command_key)

        messagebox.showinfo("Succès", f"Protocole {protocol_name} créé avec succès !\nFichier trouvé : {chemin}")

    except PermissionError:
        messagebox.showerror("Erreur", "Tu dois exécuter ce script en tant qu'administrateur !")
    except Exception as e:
        messagebox.showerror("Erreur", str(e))


# --- Lancement direct ---
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Masque la fenêtre Tkinter
    installer()
    root.destroy()
