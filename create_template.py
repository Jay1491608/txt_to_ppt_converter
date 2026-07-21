"""Create a starter template from python-pptx defaults for local testing."""

from pathlib import Path

# pyrefly: ignore [missing-import]
from pptx import Presentation

def main() -> None:
    templates_dir = Path(__file__).resolve().parent / "templates"
    templates_dir.mkdir(exist_ok=True)
    output = templates_dir / "default.pptx"

    prs = Presentation()
    prs.save(str(output))
    print(f"Created starter template: {output}")
    print("Replace with your corporate .pptx/.potx for production use.")


if __name__ == "__main__":
    main()
