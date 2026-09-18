from __future__ import annotations


def ask(prompt: str, *, default: str | None = None) -> str:
    suffix = f" [{default}]" if default else ""
    raw = input(f"{prompt}{suffix}: ").strip()
    return raw if raw else (default or "")


def ask_yes_no(prompt: str, *, default: bool = False) -> bool:
    hint = "Y/n" if default else "y/N"
    raw = input(f"{prompt} [{hint}]: ").strip().lower()
    if not raw:
        return default
    return raw in {"y", "yes"}


def ask_multi_select(options: list[str]) -> list[str]:
    for i, option in enumerate(options, start=1):
        print(f"  {i}. {option}")
    raw = input("Enter numbers separated by commas (blank for none): ").strip()
    if not raw:
        return []

    chosen: list[str] = []
    for part in raw.split(","):
        part = part.strip()
        if not part.isdigit():
            continue
        index = int(part) - 1
        if 0 <= index < len(options):
            chosen.append(options[index])
    return chosen
