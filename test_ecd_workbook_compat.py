from __future__ import annotations

from collections import Counter
from pathlib import Path

import bulletin_generator as bg


ROOT = Path(__file__).resolve().parents[2]
REFERENCE = ROOT / "work" / "reference_ecd"


def fake_extract(url: str, headline: str = "", source: str = "") -> tuple[str, str, str]:
    return (
        "This is test article text used to verify the raw ECD news links workbook without making network requests.",
        "",
        "Extracted test headline",
    )


def fake_recover(headline: str, source: str) -> tuple[str, str, str]:
    return (
        "This is recovered test article text for a row whose Link cell did not contain a URL.",
        headline or "Recovered test headline",
        "https://example.com/recovered-article",
    )


def inspect(path: Path, monkeypatch_extract: bool = False) -> None:
    if monkeypatch_extract:
        bg.extract_article_from_url = fake_extract
        bg.recover_article_without_url = fake_recover
    items, warnings = bg.read_news_items(path)
    sections = Counter(item.section for item in items)
    print(f"\n{path.name}")
    print("items", len(items))
    print("sections", dict(sections))
    print("warnings", len(warnings))
    print("first", items[0].section, items[0].headline, items[0].source, items[0].url[:60])


def main() -> None:
    inspect(REFERENCE / "news_links.xlsx", monkeypatch_extract=True)
    inspect(REFERENCE / "unzipped" / "ECD" / "news.xlsx", monkeypatch_extract=False)


if __name__ == "__main__":
    main()
