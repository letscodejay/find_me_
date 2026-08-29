"""
diagnose.py
===========
Reads every DataStage export under INPUTS/analyzer_INPUTS, runs the parser over
each one, and prints what the report generator needs in order to handle them.

    python diagnose.py
    python diagnose.py --dir "some\\other\\folder"
    python diagnose.py --full          also dumps every property name it saw

Nothing is modified. No credentials and no data values are printed - only
structural names: record types, stage types, property names and link shapes.

The long version is written to diagnose_report.txt.
"""

from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

DEFAULT_DIR = "INPUTS/analyzer_INPUTS"

LINES: list[str] = []
DETAIL: list[str] = []


def out(text: str = "") -> None:
    print(text)
    LINES.append(text)


def detail(text: str) -> None:
    DETAIL.append(text)


def head(title: str) -> None:
    out("")
    out("=" * 74)
    out(f" {title}")
    out("=" * 74)


def wrap(label: str, items, width: int = 56) -> None:
    """Print a long list across several lines, keeping the label column clean."""
    line, first = "", True
    for item in items:
        if len(line) + len(item) + 2 > width:
            out(f" {label if first else '':<15}{line.rstrip(', ')}")
            line, first = "", False
        line += f"{item}, "
    if line:
        out(f" {label if first else '':<15}{line.rstrip(', ')}")


def load_parser():
    """Import generate_doc, stubbing anything unrelated that is missing."""
    import importlib
    import types

    for name, attrs in {
        "crewai": ("Agent", "Crew", "Process", "Task"),
        "langchain_openai": ("AzureChatOpenAI",),
        "docx2pdf": ("convert",),
        "pythoncom": (),
        "win32com": (),
        "win32com.client": (),
    }.items():
        try:
            importlib.import_module(name)
        except Exception:
            module = types.ModuleType(name)
            for attr in attrs:
                setattr(module, attr, object)
            sys.modules[name] = module
    if "win32com" in sys.modules and "win32com.client" in sys.modules:
        sys.modules["win32com"].client = sys.modules["win32com.client"]

    sys.path.insert(0, str(Path.cwd()))
    try:
        from Backend.Analyzer import generate_doc as G      # type: ignore
    except Exception:
        import generate_doc as G                            # type: ignore
    G.LLM_ENABLED = False
    return G


def tag(el) -> str:
    return el.tag.rsplit("}", 1)[-1] if isinstance(el.tag, str) else ""


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", default=DEFAULT_DIR)
    ap.add_argument("--full", action="store_true")
    args = ap.parse_args()

    folder = Path(args.dir)
    if not folder.is_dir():
        found = [p for p in Path.cwd().rglob("analyzer_INPUTS") if p.is_dir()]
        folder = found[0] if found else Path(".")
    files = sorted(p for p in folder.rglob("*.xml"))

    out("DATASTAGE EXPORT DIAGNOSTIC")
    out(f" folder   {folder}")
    out(f" files    {len(files)}")
    if not files:
        out(" No .xml files found. Pass --dir with the right folder.")
        return 1

    G = load_parser()
    from lxml import etree

    record_types: Counter = Counter()
    stage_types: Counter = Counter()
    prop_names: Counter = Counter()
    collections: Counter = Counter()
    partner_shapes: Counter = Counter()
    link_type_values: Counter = Counter()
    props_by_role: dict[str, set] = {}
    object_report: list[tuple] = []
    problems: list[str] = []

    # ---------------------------------------------------------------- per file
    head("PER FILE")
    out(f" {'file':<30}{'job':>4}{'rec':>5}{'stg':>5}{'lnk':>5}"
        f"{'src':>4}{'ref':>4}{'xfm':>4}{'tgt':>4}{'unc':>4}{'path':>5}{'warn':>5}")
    out(" " + "-" * 72)

    for path in files:
        name = path.name[:29]
        try:
            root = etree.parse(str(path), etree.XMLParser(recover=True, huge_tree=True)).getroot()
        except Exception as exc:
            out(f" {name:<30} UNPARSEABLE  {exc}")
            problems.append(f"{path.name}: unparseable ({exc})")
            continue

        for rec in root.iter():
            if tag(rec) != "Record":
                continue
            record_types[rec.get("Type") or "(none)"] += 1
            for child in rec:
                t = tag(child)
                if t == "Property" and child.get("Name"):
                    key = child.get("Name")
                    prop_names[key] += 1
                    if key == "StageType":
                        stage_types["".join(child.itertext()).strip()] += 1
                    elif key in ("Partner", "PartnerLink", "LinkPartner"):
                        raw = "".join(child.itertext()).strip()
                        partner_shapes["pipe" if "|" in raw else "plain"] += 1
                    elif key == "LinkType":
                        link_type_values["".join(child.itertext()).strip()] += 1
                elif t == "Collection" and child.get("Name"):
                    collections[child.get("Name")] += 1

        try:
            jobs = G.DataStageParser(path).parse_all_jobs()
        except Exception as exc:
            out(f" {name:<30} PARSER FAILED  {type(exc).__name__}: {exc}")
            problems.append(f"{path.name}: parser failed ({exc})")
            continue

        recs = sum(len(j.all_records) for j in jobs)
        for job in jobs:
            g = G.analyze(job)
            out(f" {name:<30}{len(jobs):>4}{len(job.all_records):>5}{len(g.stages):>5}"
                f"{len(g.links):>5}{len(g.sources):>4}{len(g.references):>4}"
                f"{len(g.transformations):>4}{len(g.targets):>4}{len(g.data_objects):>4}"
                f"{len(g.paths):>5}{len(job.warnings):>5}")
            name = ""    # only label the first job of a multi-job file

            if g.stages and not g.links:
                problems.append(f"{path.name}: {len(g.stages)} stages but NO links resolved")
            if g.data_objects:
                problems.append(
                    f"{path.name}: {len(g.data_objects)} stage(s) unclassified "
                    f"({', '.join(s.display_name for s in g.data_objects[:3])})")
            for w in job.warnings[:3]:
                detail(f"{path.name} warning: {w}")

            for stage in g.stages:
                props_by_role.setdefault(stage.role, set()).update(
                    k for k, v in stage.properties.items() if v)
            for stage in g.sources + g.targets:
                if len(object_report) < 8:
                    blobs = [(k, len(v)) for k, v in stage.properties.items()
                             if v and v.lstrip().startswith("<")]
                    object_report.append(
                        (stage.display_name, stage.display_type, stage.role,
                         G.stage_object(stage),
                         sorted(k for k, v in stage.properties.items() if v),
                         blobs, stage.properties))

    # ------------------------------------------------------------- aggregates
    head("RECORD TYPES ACROSS ALL FILES")
    handled = (G.RECORD_TYPES_JOB | G.RECORD_TYPES_STAGE | G.RECORD_TYPES_INPUT
               | G.RECORD_TYPES_OUTPUT | G.RECORD_TYPES_ANNOTATION
               | G.RECORD_TYPES_IGNORED)
    wrap("handled", [f"{k} {v}" for k, v in record_types.most_common() if k in handled])
    unknown = [f"{k} {v}" for k, v in record_types.most_common() if k not in handled]
    if unknown:
        out(" NOT HANDLED - these records are ignored by the parser:")
        wrap("", unknown)
        problems.append(f"unhandled record types: {', '.join(u.split()[0] for u in unknown)}")
    else:
        out(" every record type is handled")

    head("STAGE TYPES ACROSS ALL FILES")
    wrap("types", [f"{k} {v}" for k, v in stage_types.most_common(24)])
    lookups = [k for k in stage_types if any(h in k.lower() for h in G.REFERENCE_CONSUMER_HINTS)]
    out(f" treated as reference consumers: {', '.join(lookups) or 'none'}")
    joins = [k for k in stage_types if "join" in k.lower() or "merge" in k.lower()]
    if joins:
        out(f" join/merge present (secondary input is data, not reference): {', '.join(joins)}")

    head("LINKS")
    out(f" Partner format   pipe 'stage|port': {partner_shapes['pipe']}   "
        f"plain 'port': {partner_shapes['plain']}")
    out(f" LinkType property values: "
        f"{dict(link_type_values) if link_type_values else 'property not present'}")
    if not link_type_values:
        out("   -> reference links are detected from the port index on a lookup stage")

    head("OBJECT COLUMN - where the table or file name lives")
    for nm, ty, role, resolved, keys, blobs, allprops in object_report:
        out(f" {nm[:30]:<31}{ty[:18]:<19}{role}")
        out(f"   resolved -> {resolved if resolved else 'NOT FOUND'}")
        wrap("   props", keys[:14] if not blobs else keys[:10], width=54)
        for bk, blen in blobs:
            out(f"   xml blob '{bk}' ({blen} chars), tags inside:")
            try:
                sub = etree.fromstring(allprops[bk].encode("utf-8", "ignore"),
                                       etree.XMLParser(recover=True))
                tags = []
                for el in sub.iter():
                    t = tag(el)
                    if t and t not in tags:
                        tags.append(t)
                wrap("     ", tags[:26], width=52)
            except Exception as exc:
                out(f"     could not parse blob: {exc}")
        detail(f"{nm} all properties: {sorted(allprops)}")

    head("PROPERTY NAMES BY ROLE")
    for role in ("source", "reference", "transformation", "target", "unclassified"):
        if role in props_by_role:
            wrap(role, sorted(props_by_role[role])[: (60 if args.full else 14)])

    if collections:
        head("COLLECTIONS (column and schema data)")
        wrap("names", [f"{k} {v}" for k, v in collections.most_common(14)])

    head("PROBLEMS")
    if problems:
        for p in dict.fromkeys(problems):
            out(f" - {p}")
    else:
        out(" none")

    out("")
    out("=" * 74)
    Path("diagnose_report.txt").write_text(
        "\n".join(LINES) + "\n\nDETAIL\n" + "\n".join(DETAIL), encoding="utf-8")
    out(" Long version in diagnose_report.txt")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
