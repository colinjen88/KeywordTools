import sys
import traceback

print(f"Python executable: {sys.executable}")
print(f"Python version: {sys.version}")

try:
    import tkcalendar
    print("SUCCESS: tkcalendar imported successfully.")
    print(f"tkcalendar file: {tkcalendar.__file__}")
except Exception:
    print("ERROR: Failed to import tkcalendar.")
    traceback.print_exc()

try:
    import babel
    print(f"babel version: {babel.__version__}")
    print(f"babel file: {babel.__file__}")
except Exception:
    print("ERROR: Failed to import babel (dependency of tkcalendar).")
    traceback.print_exc()
