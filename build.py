"""Generate the IRRBB PowerPoint report: python build.py"""
import argparse
import json
import sys
from pathlib import Path

from pptx import Presentation
from pptx.exc import PackageNotFoundError

from irrbb.calc import calculate
from irrbb.data import InputError, load_inputs
from irrbb.deck import DeckBuilder
from irrbb.pptx_helpers import TemplateError

ROOT = Path(__file__).resolve().parent


def list_layouts(template):
    for layout in Presentation(template).slide_layouts:
        print(layout.name)
        for placeholder in layout.placeholders:
            print(f"    idx {placeholder.placeholder_format.idx}: {placeholder.placeholder_format.type} ({placeholder.name})")


def main():
    parser = argparse.ArgumentParser(description="Generate the IRRBB PowerPoint report from the Excel input workbook.")
    parser.add_argument("--data", help="input workbook; overrides config.json")
    parser.add_argument("--template", help=".pptx template; overrides config.json")
    parser.add_argument("--output", help="output .pptx path; default is a dated file in the output folder")
    parser.add_argument("--list-layouts", action="store_true", help="print the template's layouts and placeholders, then exit")
    args = parser.parse_args()

    config = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
    config["template"] = str(ROOT / (args.template or config["template"]))
    data_path = ROOT / (args.data or config["data"])

    try:
        if args.list_layouts:
            list_layouts(config["template"])
            return 0
        inputs = load_inputs(data_path)
        results = calculate(inputs)
        if args.output:
            output = Path(args.output)
        else:
            output = ROOT / config["output_dir"] / f"IRRBB_Report_{inputs.reporting_date:%Y-%m-%d}.pptx"
        output.parent.mkdir(parents=True, exist_ok=True)
        slides = DeckBuilder(inputs, results, config).build(output)
    except (InputError, TemplateError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1
    except PackageNotFoundError:
        print(f"ERROR: template not found or not a .pptx file: {config['template']}", file=sys.stderr)
        return 1
    except FileNotFoundError as error:
        print(f"ERROR: file not found: {error.filename}", file=sys.stderr)
        return 1
    except PermissionError as error:
        print(f"ERROR: cannot write {error.filename}. Close it in PowerPoint and run again.", file=sys.stderr)
        return 1

    print(f"Built {slides} slides: {output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
