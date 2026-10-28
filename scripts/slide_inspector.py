"""Read-only legacy adapter for pmg.validator.StaticValidator."""
import argparse
import json
import os
import sys

if __package__ in (None, ""):
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def audit_presentation(pptx_path: str):
    from pmg.validator import StaticValidator
    report = StaticValidator().validate(pptx_path)
    print(json.dumps(report, indent=2, ensure_ascii=False))
    print("Static inspection only; playback certification is NOT_VERIFIED.")
    return report["is_valid"]


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx")
    sys.exit(0 if audit_presentation(parser.parse_args().pptx) else 1)
