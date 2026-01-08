import shutil
import os
import sys

def clean_gen_py():
    # Der Standardpfad für den Cache
    user_temp = os.environ.get('LOCALAPPDATA')
    if not user_temp:
        print("Konnte LOCALAPPDATA nicht finden.")
        return

    gen_py_path = os.path.join(user_temp, 'Temp', 'gen_py')
    
    print(f"Suche Cache in: {gen_py_path}")

    if os.path.exists(gen_py_path):
        try:
            shutil.rmtree(gen_py_path)
            print("✅ Cache erfolgreich gelöscht! ('gen_py' Ordner entfernt)")
        except PermissionError:
            print("❌ Zugriff verweigert. Bitte schließe ALLE Excel-Instanzen und versuche es erneut.")
        except Exception as e:
            print(f"❌ Fehler: {e}")
    else:
        print("ℹ️ Kein Cache gefunden (Ordner existiert nicht). Das ist okay.")

if __name__ == "__main__":
    clean_gen_py()