import json
from pathlib import Path


SOURCE = Path.home() / "Downloads" / "kings_reign_timeline_prophets.html"
OUT_DIR = Path(__file__).resolve().parents[1] / "data"
OUT_FILE = OUT_DIR / "timeline_enriched.json"


def extract_balanced(text: str, marker: str, opener: str, closer: str) -> str:
    start = text.index(marker)
    start = text.index(opener, start)
    depth = 0
    in_string = False
    escaped = False
    for index in range(start, len(text)):
        char = text[index]
        if in_string:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
            continue
        if char == '"':
            in_string = True
        elif char == opener:
            depth += 1
        elif char == closer:
            depth -= 1
            if depth == 0:
                return text[start : index + 1]
    raise ValueError(f"Could not find balanced block for {marker}")


def main() -> None:
    text = SOURCE.read_text(encoding="utf-8")
    data = json.loads(extract_balanced(text, "const DATA", "{", "}"))
    data["prophets"] = json.loads(extract_balanced(text, "const PROPHETS", "[", "]"))

    OUT_DIR.mkdir(exist_ok=True)
    OUT_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(OUT_FILE)
    print(f"rulers={len(data['rulers'])}, judah={len(data['cards']['judah'])}, israel={len(data['cards']['israel'])}, prophets={len(data['prophets'])}")


if __name__ == "__main__":
    main()
