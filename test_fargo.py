import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import importlib.util, pathlib

# Load the module
spec = importlib.util.spec_from_file_location(
    "crc",
    r"c:\Users\franco.faustin\OneDrive - OneWorkplace\Escritorio\Claude Van Damme\Diferenciador de facturas\cargar_real_consumido.py"
)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

path = mod.RC_DIR / "Real Consumido_Fargo_Q1_2026_Omnet_V5.xlsx"
print(f"Procesando solo Fargo: {path.name}")
mod.process(path, "Fargo")
print("\n✅ Test Fargo completado. Revisá el archivo y verificá que esté correcto.")
