from __future__ import annotations

import json
import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "mathtype-word"


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def read_text(path: Path) -> str:
    require(path.is_file(), f"Missing required file: {path.relative_to(ROOT)}")
    return path.read_text(encoding="utf-8-sig")


def validate_skill_metadata() -> None:
    text = read_text(SKILL / "SKILL.md")
    match = re.match(r"\A---\r?\n(.*?)\r?\n---\r?\n", text, re.DOTALL)
    require(match is not None, "SKILL.md needs YAML frontmatter delimited by ---")
    frontmatter = match.group(1)

    name_match = re.search(r"(?m)^name:\s*(\S+)\s*$", frontmatter)
    description_match = re.search(r"(?m)^description:\s*(.+?)\s*$", frontmatter)
    require(name_match is not None, "SKILL.md frontmatter is missing name")
    require(description_match is not None, "SKILL.md frontmatter is missing description")
    require(name_match.group(1) == "mathtype-word", "Skill name must be mathtype-word")
    require(
        re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", name_match.group(1)) is not None,
        "Skill name must use lowercase letters, digits, and single hyphens",
    )
    description = description_match.group(1)
    require(20 <= len(description) <= 1024, "Skill description has an unexpected length")
    require("<" not in description and ">" not in description, "Description cannot contain angle brackets")

    agent_yaml = read_text(SKILL / "agents" / "openai.yaml")
    require("$mathtype-word" in agent_yaml, "Agent default prompt must name $mathtype-word")


def validate_layout() -> None:
    required = {
        "SKILL.md",
        "VERSION",
        "LICENSE",
        "THIRD_PARTY_NOTICES.md",
        "agents/openai.yaml",
        "references/manifest.md",
        "references/latex-compatibility.md",
        "scripts/insert-equations.ps1",
        "scripts/WordComBridge.cs",
    }
    present = {
        path.relative_to(SKILL).as_posix()
        for path in SKILL.rglob("*")
        if path.is_file()
    }
    missing = required - present
    require(not missing, f"Missing distributable files: {sorted(missing)}")

    forbidden_suffixes = {".docx", ".doc", ".pdf", ".dll", ".exe", ".wll", ".dotm", ".pyc"}
    forbidden = sorted(
        str(path.relative_to(SKILL))
        for path in SKILL.rglob("*")
        if path.is_file() and path.suffix.lower() in forbidden_suffixes
    )
    require(not forbidden, f"Proprietary, generated, or binary files found in skill: {forbidden}")

    local_path_pattern = re.compile(r"(?:[A-Za-z]:\\Users\\|mathtypde解放双手)", re.IGNORECASE)
    for path in SKILL.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in {".md", ".ps1", ".cs", ".yaml", ".yml"}:
            continue
        require(
            local_path_pattern.search(read_text(path)) is None,
            f"Machine-specific path leaked into {path.relative_to(ROOT)}",
        )


def validate_version_and_example() -> None:
    version = read_text(ROOT / "VERSION").strip()
    require(
        re.fullmatch(r"\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?", version) is not None,
        "VERSION must be a semantic version",
    )
    require(read_text(SKILL / "VERSION").strip() == version, "Root and skill VERSION files differ")
    require(read_text(SKILL / "LICENSE") == read_text(ROOT / "LICENSE"), "Root and skill LICENSE files differ")
    require(
        read_text(SKILL / "THIRD_PARTY_NOTICES.md") == read_text(ROOT / "THIRD_PARTY_NOTICES.md"),
        "Root and skill third-party notices differ",
    )

    example = json.loads(read_text(ROOT / "examples" / "equations.example.json"))
    require(set(example) <= {"input", "output", "pdf", "equations"}, "Example has unknown top-level fields")
    require({"input", "output", "equations"} <= set(example), "Example is missing required fields")
    require(isinstance(example["equations"], list) and example["equations"], "Example equations must be nonempty")

    locators: set[str] = set()
    for index, item in enumerate(example["equations"], start=1):
        require(set(item) <= {"anchor", "bookmark", "latex", "mode"}, f"Equation {index} has unknown fields")
        require(("anchor" in item) ^ ("bookmark" in item), f"Equation {index} needs exactly one locator")
        require(item.get("mode") in {"inline", "display", "right-numbered"}, f"Equation {index} has an invalid mode")
        require(isinstance(item.get("latex"), str) and item["latex"].strip(), f"Equation {index} has empty LaTeX")
        locator = item.get("anchor", item.get("bookmark"))
        require(locator not in locators, f"Duplicate example locator: {locator}")
        locators.add(locator)


def validate_local_markdown_links() -> None:
    link_pattern = re.compile(r"\[[^\]]+\]\(([^)]+)\)")
    for path in [ROOT / "README.md", ROOT / "README.zh-CN.md", SKILL / "SKILL.md"]:
        text = read_text(path)
        for target in link_pattern.findall(text):
            if re.match(r"^[a-z]+://", target, re.IGNORECASE) or target.startswith("#"):
                continue
            link_path = (path.parent / target.split("#", 1)[0]).resolve()
            require(link_path.exists(), f"Broken local link in {path.name}: {target}")


def main() -> int:
    validate_skill_metadata()
    validate_layout()
    validate_version_and_example()
    validate_local_markdown_links()
    print("Release metadata, layout, example manifest, and local links are valid.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (AssertionError, json.JSONDecodeError, UnicodeError) as exc:
        print(f"validation failed: {exc}", file=sys.stderr)
        raise SystemExit(1)
