"""
Paste a TOML object into the console and get back a Python dict you can copy.

Usage:
    python toml_to_dict.py

Paste your TOML text, then on a new line type END (or press Ctrl+D on
Mac/Linux, Ctrl+Z then Enter on Windows) to finish input.
"""

import sys
import tomllib
from pprint import pformat


def read_multiline_input() -> str:
    print("Paste your TOML below.")
    print("When done, type END on its own line (or Ctrl+D / Ctrl+Z+Enter):\n")

    lines = []
    try:
        while True:
            line = input()
            if line.strip() == "END":
                break
            lines.append(line)
    except EOFError:
        pass  # Ctrl+D / Ctrl+Z was used instead of typing END

    return "\n".join(lines)


def main():
    toml_text = read_multiline_input()

    if not toml_text.strip():
        print("\nNo input received.")
        sys.exit(1)

    try:
        data = tomllib.loads(toml_text)
    except tomllib.TOMLDecodeError as e:
        print(f"\nCould not parse TOML: {e}")
        sys.exit(1)

    print("\n--- Python dict (copy below) ---\n")
    print(pformat(data))


if __name__ == "__main__":
    main()
