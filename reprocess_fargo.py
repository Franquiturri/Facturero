import sys; sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import shutil
import cargar_real_consumido as crc

fname = "Real Consumido_Fargo_Q1_2026_Omnet_V5.xlsx"
bak   = "Real Consumido_Fargo_Q1_2026_Omnet_V5_BACKUP.xlsx"

path     = crc.RC_DIR / fname
bak_path = crc.RC_DIR / bak

print(f"Restaurando desde backup...")
shutil.copy2(bak_path, path)
print(f"OK — {fname} restaurado")

print(f"\n{'='*55}\n{fname}\n{'='*55}")
crc.process(path, "Fargo")
