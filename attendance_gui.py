from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if SRC.exists() and str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from attendance.gui import main


def _smoke_test() -> None:
    import numpy
    import pandas

    if not numpy.__version__ or not pandas.__version__:
        raise RuntimeError("No se pudieron cargar las dependencias de datos.")


if __name__ == "__main__":
    if "--smoke-test" in sys.argv:
        _smoke_test()
    else:
        main()
