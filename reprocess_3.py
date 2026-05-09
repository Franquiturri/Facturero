import sys; sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import cargar_real_consumido as crc

targets = [
    ("Real Consumido_Rapiditas_CampañaRegional_Feb-Mar.xlsx", "Rapiditas"),
    ("Real Consumido_Salmas_V7.xlsx", "Salmas"),
    ("Real Consumido_BackToSchol_Feb-Abr.xlsx", "Bimbo_BTS"),
]
for fname, adv in targets:
    path = crc.RC_DIR / fname
    print(f"\n{'='*55}\n{fname}\n{'='*55}")
    crc.process(path, adv)
