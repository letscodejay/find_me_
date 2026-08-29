"""
generate_doc.py
===============
DataStage workflow XML  ->  analysis  ->  PDF report.

    .xml  ->  DataStageParser  ->  WorkflowGraph  ->  Narrative  ->  PDF

Rendering is done directly with ReportLab, which is a pure Python wheel: no
system libraries, no browser, nothing to install beyond pip. Page geometry and
column widths are given in points, so the layout is identical everywhere.

One input file producing one report gives a .pdf; anything more gives a .zip.

    SECTION A   Tunables
    SECTION B   Parser
    SECTION C   Analysis
    SECTION D   Narrative
    SECTION E   Rendering
    SECTION F   Orchestration and CLI
"""

from __future__ import annotations

import io
import json
import logging
import os
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Sequence

from crewai import Agent, Crew, Process, Task
from langchain_openai import AzureChatOpenAI
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4, LETTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import (
    BaseDocTemplate, Frame, KeepTogether, PageBreak, PageTemplate, Paragraph,
    Preformatted, Spacer, Table, TableStyle,
)

log = logging.getLogger("generate_doc")
HERE = Path(__file__).resolve().parent


# =============================================================================
# SECTION A   TUNABLES
# =============================================================================

# --- Files --------------------------------------------------------------------
DEFAULT_OUTPUT_DIR = "reports"

# Where selected exports live. document_generate() gets a project and a file
# name, not a path, so each of these is tried in turn.
PROJECT_ROOTS = (
    "INPUTS/analyzer_INPUTS", "INPUTS", "uploads", "data", "projects",
    "Backend/uploads", ".",
)
PROJECT_FILE_SEARCH_DEPTH = 4

# --- Page ---------------------------------------------------------------------
PAGE_SIZE = "A4"                     # "A4" or "LETTER"
MARGIN_TOP_MM = 20.0
MARGIN_BOTTOM_MM = 18.0
MARGIN_SIDE_MM = 20.0

# --- Palette ------------------------------------------------------------------
COLOR_CORAL = "#C4503C"
COLOR_CORAL_DEEP = "#A63D2B"
COLOR_CORAL_LINE = "#E8C3BA"
COLOR_INK = "#17242E"
COLOR_INK_DIM = "#5D6E78"
COLOR_INK_FAINT = "#8A979E"
COLOR_RULE = "#DCDFDC"
COLOR_RULE_SOFT = "#EBEDEA"
COLOR_HEAD_BG = "#F1F0EC"
COLOR_KEY_BG = "#F7F6F2"
COLOR_BAND_BG = "#F4F3EF"
COLOR_PANEL_BG = "#EEF4F4"
COLOR_PANEL_LINE = "#B6D2D4"
COLOR_ROLE_SOURCE = "#2F6F4E"
COLOR_ROLE_REFERENCE = "#5B4B8A"
COLOR_ROLE_TRANSFORM = "#8A5A00"
COLOR_ROLE_TARGET = "#8C3A3A"

# The PDF base-14 fonts. They are part of the PDF specification, so nothing is
# installed, nothing is embedded, and the metrics are identical on every reader.
FONT_SANS = "Helvetica"
FONT_SANS_BOLD = "Helvetica-Bold"
FONT_SERIF = "Times-Roman"
FONT_SERIF_ITALIC = "Times-Italic"
FONT_MONO = "Courier"
FONT_MONO_BOLD = "Courier-Bold"

# --- Type scale (pt) ----------------------------------------------------------
SIZE_SECTION_NUMBER = 7.0
SIZE_SECTION_TITLE = 13.0
SIZE_BODY = 9.5
SIZE_TABLE = 8.5
SIZE_TABLE_HEAD = 7.5
SIZE_CAPTION = 7.0
SIZE_DIAGRAM = 7.5
SIZE_RUNNING = 7.5

# --- Report wording -----------------------------------------------------------
REPORT_TITLE = "Workflow Analysis Report"
REPORT_SUBTITLE = "IBM InfoSphere DataStage"
TEXT_NOT_IDENTIFIED = "Not identified from file"
TEXT_NOT_APPLICABLE = "Not applicable"

SECTION_STARTS_NEW_PAGE = True
SKIP_EMPTY_SECTIONS = True
MARK_GENERATED_PROSE = True          # a rule beside model-written text

# --- Content limits -----------------------------------------------------------
MAX_DESCRIPTION_CHARS = 340
MAX_PATH_CELL_CHARS = 260
MAX_PROSE_PARAGRAPHS = 3
MAX_PROSE_CHARS = 1500
TRUNCATION_MARK = "…"

# --- Diagram ------------------------------------------------------------------
MAX_STAGES_FOR_DIAGRAM = 40
MAX_DIAGRAM_LINE_CHARS = 78          # fits the content width at SIZE_DIAGRAM
# Courier has no box-drawing glyphs, so the chart is drawn in ASCII. "box" needs
# a TrueType font registered with ReportLab first.
DIAGRAM_STYLE = "ascii"
GLANCE_DIAGRAM_STAGES = 6            # Section 1 shows the main chain only

# --- Analysis limits ----------------------------------------------------------
MAX_PATHS = 200
MAX_PATH_DEPTH = 50

# --- LLM ----------------------------------------------------------------------
LLM_ENABLED = True
LLM_TEMPERATURE = 0.0
LLM_TIMEOUT_SECONDS = 180
LLM_MAX_RETRIES = 1
LLM_CONFIG_MODULES = (
    "Modules.Libraries", "modules.Libraries", "Modules.libraries",
    "modules.libraries", "Libraries",
)
LLM_INSTANCE_NAMES = ("llm", "LLM", "azure_llm", "chat_model", "client", "openai_client")
LLM_CONFIG_NAMES = {
    "deployment": ("AZURE_OPENAI_DEPLOYMENT", "AZURE_DEPLOYMENT", "DEPLOYMENT_NAME",
                   "AZURE_OPENAI_DEPLOYMENT_NAME", "deployment_name", "MODEL_NAME",
                   "AZURE_DEPLOYMENT_NAME"),
    "endpoint": ("AZURE_OPENAI_ENDPOINT", "AZURE_ENDPOINT", "OPENAI_API_BASE",
                 "AZURE_OPENAI_API_BASE", "azure_endpoint", "ENDPOINT"),
    "api_key": ("AZURE_OPENAI_API_KEY", "AZURE_API_KEY", "OPENAI_API_KEY",
                "AZURE_OPENAI_KEY", "api_key", "API_KEY"),
    "api_version": ("AZURE_OPENAI_API_VERSION", "OPENAI_API_VERSION", "API_VERSION",
                    "api_version"),
}
AZURE_API_VERSION_DEFAULT = "2024-08-01-preview"


# =============================================================================
# SECTION B   PARSER
# =============================================================================
# XML to a typed model. No formatting, no graph reasoning.
#
# Record types are recognised structurally, not from a fixed list. A list is
# what let 46 HashedFileStage records be dropped in silence; anything carrying a
# StageType is treated as a stage even when its type name is new.

RECORD_TYPES_JOB = {"DSJob", "JobDefn", "Job", "JobDefinition"}
RECORD_TYPES_ANNOTATION = {"Annotation", "CAnnotation", "Note"}
# Definitions and designer furniture, never workflow objects.
RECORD_TYPES_IGNORED = {
    "StageType", "ContainerView", "JobView", "TableDef", "TableDefinition",
    "Routine", "SharedContainerDef", "ParameterSet", "DataElement",
}
# Known names, kept so an export that omits StageType still resolves.
RECORD_TYPES_STAGE = {
    "CustomStage", "TransformerStage", "HashedFileStage", "SequentialFileStage",
    "ODBCStage", "OracleStage", "Stage", "ContainerStage", "CContainer",
    "CTrxStage", "ServerStage", "ParallelStage", "AggregatorStage", "SortStage",
}
RECORD_TYPES_INPUT = {
    "CustomInput", "TrxInput", "HashedInput", "CTrxInput", "StageInput", "Input",
    "ODBCInput", "OracleInput", "SeqInput",
}
RECORD_TYPES_OUTPUT = {
    "CustomOutput", "TrxOutput", "HashedOutput", "CTrxOutput", "StageOutput",
    "Output", "ODBCOutput", "OracleOutput", "SeqOutput",
}

REFERENCE_CONSUMER_HINTS = ("lookup",)
NAME_PROPERTY_CANDIDATES = ("Name", "StageName", "LinkName", "Identifier")
TYPE_PROPERTY_CANDIDATES = ("StageType", "OLEType", "StageTypeName", "Type")
PARTNER_PROPERTY_CANDIDATES = ("Partner", "PartnerLink", "LinkPartner")

_PORT_ID_RE = re.compile(r"^(?P<stage>.*?S\d+)P(?P<port>\d+)$", re.IGNORECASE)


def _parse_partner(value: str | None) -> tuple[str | None, str | None]:
    """Partner may be 'V0S2P1' or 'V0S2|V0S2P1'. Returns (stage, port)."""
    if not value:
        return None, None
    parts = [p.strip() for p in value.replace("\\", "|").split("|") if p.strip()]
    if not parts:
        return None, None
    if len(parts) >= 2:
        return parts[0], parts[-1]
    return None, parts[0]


class ParseError(Exception):
    """The file could not be read as a DataStage export."""


class NoJobsFound(ParseError):
    pass


@dataclass
class Stage:
    identifier: str
    name: str | None
    stage_type: str | None
    properties: dict[str, str] = field(default_factory=dict)
    collections: dict[str, list[dict[str, str]]] = field(default_factory=dict)
    record_type: str = ""
    recognised: bool = True          # False when the type was inferred, not known
    annotation: str | None = None
    role: str = "unclassified"
    inputs: list[str] = field(default_factory=list)
    outputs: list[str] = field(default_factory=list)

    @property
    def display_name(self) -> str:
        return self.name or TEXT_NOT_IDENTIFIED

    @property
    def display_type(self) -> str:
        return self.stage_type or self.record_type or TEXT_NOT_IDENTIFIED

    def prop(self, *names: str, default: str | None = None) -> str | None:
        for n in names:
            v = self.properties.get(n)
            if v:
                return v
        return default


@dataclass
class Link:
    name: str | None
    from_stage: str
    to_stage: str
    from_port: str
    to_port: str
    to_port_index: int | None = None
    is_reference: bool = False
    is_reject: bool = False
    column_count: int | None = None
    properties: dict[str, str] = field(default_factory=dict)

    @property
    def display_name(self) -> str:
        return self.name or TEXT_NOT_IDENTIFIED

    @property
    def kind(self) -> str:
        if self.is_reference:
            return "Reference"
        return "Reject" if self.is_reject else "Stream"


@dataclass
class ParsedJob:
    job_name: str
    source_file: str
    root_record: dict[str, Any]
    job_record: dict[str, Any]
    all_records: dict[str, dict[str, Any]]
    stages: list[Stage]
    links: list[Link]
    annotations: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    inferred_types: set[str] = field(default_factory=set)


class DataStageParser:
    """Reads a DataStage XML export into one ParsedJob per job in the file."""

    def __init__(self, file_path: str | Path):
        self.path = Path(file_path)
        if not self.path.is_file():
            raise ParseError(f"File not found: {self.path}")

    def parse_all_jobs(self) -> list[ParsedJob]:
        root = self._load_xml()
        jobs = self._find_job_elements(root)
        if not jobs:
            raise NoJobsFound(
                f"No <Job> elements found in {self.path.name}. "
                "The file may not be a DataStage XML export.")
        return [self._parse_job(el) for el in jobs]

    # -- XML ------------------------------------------------------------------

    def _load_xml(self):
        try:
            from lxml import etree
            return etree.parse(
                str(self.path), etree.XMLParser(recover=True, huge_tree=True)).getroot()
        except ParseError:
            raise
        except Exception as exc:
            raise ParseError(f"Could not parse {self.path.name}: {exc}") from exc

    @staticmethod
    def _tag(el) -> str:
        t = el.tag
        return t.rsplit("}", 1)[-1] if isinstance(t, str) else ""

    def _iter_children(self, el, *names: str):
        wanted = {n.lower() for n in names}
        for child in el:
            if self._tag(child).lower() in wanted:
                yield child

    def _find_job_elements(self, root) -> list:
        found = [el for el in root.iter() if self._tag(el) == "Job"]
        if found:
            return found
        if any(self._tag(e) == "Record" for e in root.iter()):
            return [root]        # some exports hold a single implicit job
        return []

    @staticmethod
    def _text_of(el) -> str:
        try:
            return "".join(p for p in el.itertext() if p).strip()
        except Exception:
            return (el.text or "").strip()

    def _read_properties(self, rec) -> dict[str, str]:
        props: dict[str, str] = {}
        for prop in self._iter_children(rec, "Property"):
            name = prop.get("Name") or prop.get("name")
            if not name:
                continue
            value = self._text_of(prop)
            if name not in props or (not props[name] and value):
                props[name] = value
        return props

    def _read_collections(self, rec) -> dict[str, list[dict[str, str]]]:
        out: dict[str, list[dict[str, str]]] = {}
        for coll in self._iter_children(rec, "Collection"):
            name = coll.get("Name") or coll.get("name") or "Collection"
            out[name] = [self._read_properties(sub)
                         for sub in self._iter_children(coll, "SubRecord", "Record")]
        return out

    # -- job ------------------------------------------------------------------

    def _parse_job(self, job_el) -> ParsedJob:
        job_record = {
            "identifier": job_el.get("Identifier") or "",
            "date_modified": job_el.get("DateModified") or "",
            "time_modified": job_el.get("TimeModified") or "",
        }
        all_records = self._index_records(job_el)
        root_record = next(
            (r for r in all_records.values() if r["type"] in RECORD_TYPES_JOB),
            {"identifier": "", "type": "", "properties": {}, "collections": {}})

        job_name = (self._first_prop(root_record, *NAME_PROPERTY_CANDIDATES)
                    or job_record["identifier"] or self.path.stem)

        stages, inferred = self._build_stages(all_records)
        links, warnings = self._resolve_links(all_records, stages)
        self._attach_ports(stages, links)

        return ParsedJob(
            job_name=job_name, source_file=self.path.name,
            root_record=root_record, job_record=job_record, all_records=all_records,
            stages=stages, links=links,
            annotations=self._collect_annotations(all_records),
            warnings=warnings, inferred_types=inferred,
        )

    def _index_records(self, job_el) -> dict[str, dict[str, Any]]:
        records: dict[str, dict[str, Any]] = {}
        anon = 0
        for rec in job_el.iter():
            if self._tag(rec) != "Record":
                continue
            ident = rec.get("Identifier") or rec.get("identifier")
            if not ident:
                anon += 1
                ident = f"__anon_{anon}"
            records[ident] = {
                "identifier": ident,
                "type": rec.get("Type") or rec.get("type") or "",
                "properties": self._read_properties(rec),
                "collections": self._read_collections(rec),
            }
        return records

    @staticmethod
    def _first_prop(record: dict[str, Any] | None, *names: str) -> str | None:
        if not record:
            return None
        props = record.get("properties", {})
        for n in names:
            if props.get(n):
                return props[n]
        return None

    # -- structural classification -------------------------------------------

    @staticmethod
    def _is_port(record: dict[str, Any]) -> bool:
        rtype = record["type"]
        if rtype in RECORD_TYPES_INPUT or rtype in RECORD_TYPES_OUTPUT:
            return True
        if any(record["properties"].get(p) for p in PARTNER_PROPERTY_CANDIDATES):
            return True
        return bool(re.search(r"(input|output)$", rtype, re.IGNORECASE))

    @staticmethod
    def _is_output_port(record: dict[str, Any]) -> bool:
        if record["type"] in RECORD_TYPES_OUTPUT:
            return True
        if record["type"] in RECORD_TYPES_INPUT:
            return False
        return bool(re.search(r"output$", record["type"], re.IGNORECASE))

    @classmethod
    def _stage_kind(cls, record: dict[str, Any]) -> str:
        """'known', 'inferred' or '' - inferred means it looks like a stage but
        its record type is not one we have seen before."""
        rtype = record["type"]
        if rtype in RECORD_TYPES_IGNORED or rtype in RECORD_TYPES_JOB \
                or rtype in RECORD_TYPES_ANNOTATION:
            return ""
        if cls._is_port(record):
            return ""
        if rtype in RECORD_TYPES_STAGE:
            return "known"
        # Anything carrying a stage type, or named like a stage, is a stage.
        if record["properties"].get("StageType") or rtype.lower().endswith("stage"):
            return "inferred"
        return ""

    def _build_stages(self, records: dict[str, dict[str, Any]]) -> tuple[list[Stage], set[str]]:
        stages: list[Stage] = []
        inferred: set[str] = set()
        for rec in records.values():
            kind = self._stage_kind(rec)
            if not kind:
                continue
            if kind == "inferred":
                inferred.add(rec["type"])
            stages.append(Stage(
                identifier=rec["identifier"],
                name=self._first_prop(rec, *NAME_PROPERTY_CANDIDATES),
                stage_type=self._first_prop(rec, *TYPE_PROPERTY_CANDIDATES) or rec["type"],
                properties=dict(rec["properties"]),
                collections=dict(rec["collections"]),
                record_type=rec["type"],
                recognised=(kind == "known"),
            ))
        stages.sort(key=lambda s: self._identifier_sort_key(s.identifier))
        return stages, inferred

    @staticmethod
    def _identifier_sort_key(ident: str) -> tuple:
        nums = [int(n) for n in re.findall(r"\d+", ident)]
        return (len(nums) == 0, nums, ident)

    @classmethod
    def _owner_of(cls, port_id: str) -> str | None:
        m = _PORT_ID_RE.match(port_id or "")
        return m.group("stage") if m else None

    @classmethod
    def _port_index(cls, port_id: str) -> int | None:
        m = _PORT_ID_RE.match(port_id or "")
        return int(m.group("port")) if m else None

    # -- links ----------------------------------------------------------------
    # An edge is a pair of port records joined by Partner. Walking outputs only
    # means each edge is produced exactly once.

    def _resolve_links(self, records, stages) -> tuple[list[Link], list[str]]:
        stage_ids = {s.identifier for s in stages}
        links: list[Link] = []
        warnings: list[str] = []

        for rec in records.values():
            if not self._is_port(rec) or not self._is_output_port(rec):
                continue

            out_id = rec["identifier"]
            raw = self._first_prop(rec, *PARTNER_PROPERTY_CANDIDATES)
            partner_stage, partner_port = _parse_partner(raw)
            if not partner_stage and not partner_port:
                warnings.append(f"Output port {out_id} declares no partner link.")
                continue

            to_stage = partner_stage or self._owner_of(partner_port or "")
            back = records.get(partner_port or "", {})
            from_stage = _parse_partner(
                self._first_prop(back, *PARTNER_PROPERTY_CANDIDATES))[0] \
                or self._owner_of(out_id)

            if not from_stage or not to_stage:
                warnings.append(
                    f"Could not determine the stages either side of {out_id} -> {raw}.")
                continue
            if from_stage not in stage_ids or to_stage not in stage_ids:
                missing = from_stage if from_stage not in stage_ids else to_stage
                warnings.append(
                    f"Link {out_id} -> {raw} names {missing}, which this export does not "
                    "describe as a stage. This normally indicates a shared container "
                    "boundary.")
                continue

            in_rec = records.get(partner_port or "", {})
            links.append(Link(
                name=(self._first_prop(rec, *NAME_PROPERTY_CANDIDATES)
                      or self._first_prop(in_rec, *NAME_PROPERTY_CANDIDATES)),
                from_stage=from_stage, to_stage=to_stage,
                from_port=out_id, to_port=partner_port or "",
                to_port_index=self._port_index(partner_port or ""),
                column_count=self._column_count(rec) or self._column_count(in_rec),
                properties=dict(rec["properties"]),
            ))

        self._flag_link_kinds(links, stages)
        return links, warnings

    @staticmethod
    def _column_count(record: dict[str, Any]) -> int | None:
        for name, rows in (record.get("collections") or {}).items():
            if "column" in name.lower():
                return len(rows)
        return None

    def _flag_link_kinds(self, links: list[Link], stages: list[Stage]) -> None:
        """Reference and reject are derivations - DataStage flags neither."""
        by_id = {s.identifier: s for s in stages}
        for link in links:
            explicit = (link.properties.get("LinkType") or "").strip().lower()
            if explicit in ("reference", "lookup"):
                link.is_reference = True
            elif explicit in ("reject", "rejects"):
                link.is_reject = True

            name = (link.name or "").lower()
            if any(k in name for k in ("rej", "err", "bad")):
                link.is_reject = True

            target = by_id.get(link.to_stage)
            if link.is_reference or not target or not target.stage_type:
                continue
            if any(h in target.stage_type.lower() for h in REFERENCE_CONSUMER_HINTS) \
                    and (link.to_port_index or 1) > 1:
                link.is_reference = True

    def _attach_ports(self, stages: list[Stage], links: list[Link]) -> None:
        by_id = {s.identifier: s for s in stages}
        for link in links:
            if link.from_stage in by_id:
                by_id[link.from_stage].outputs.append(link.display_name)
            if link.to_stage in by_id:
                by_id[link.to_stage].inputs.append(link.display_name)

    def _collect_annotations(self, records) -> list[str]:
        notes = []
        for rec in records.values():
            if rec["type"] not in RECORD_TYPES_ANNOTATION:
                continue
            text = self._first_prop(rec, "AnnotationText", "Text", "Description")
            if text:
                notes.append(text)
        return notes


# =============================================================================
# SECTION C   ANALYSIS
# =============================================================================
# Stages and links to derived facts: roles, paths, findings.

ROLE_SOURCE = "source"
ROLE_REFERENCE = "reference"
ROLE_TRANSFORMATION = "transformation"
ROLE_TARGET = "target"
ROLE_UNCLASSIFIED = "unclassified"

ROLE_LABELS = {
    ROLE_SOURCE: "Source",
    ROLE_REFERENCE: "Reference",
    ROLE_TRANSFORMATION: "Transformation",
    ROLE_TARGET: "Target",
    ROLE_UNCLASSIFIED: TEXT_NOT_APPLICABLE,
}


@dataclass
class Finding:
    ref: str
    category: str
    obj: str
    text: str


@dataclass
class WorkflowGraph:
    job: ParsedJob
    sources: list[Stage] = field(default_factory=list)
    references: list[Stage] = field(default_factory=list)
    transformations: list[Stage] = field(default_factory=list)
    targets: list[Stage] = field(default_factory=list)
    data_objects: list[Stage] = field(default_factory=list)
    paths: list[list[str]] = field(default_factory=list)
    paths_truncated: bool = False
    total_path_count: int = 0
    has_cycles: bool = False
    findings: list[Finding] = field(default_factory=list)

    @property
    def stages(self) -> list[Stage]:
        return self.job.stages

    @property
    def links(self) -> list[Link]:
        return self.job.links

    def outgoing(self, stage: Stage) -> list[Link]:
        return [l for l in self.links if l.from_stage == stage.identifier]

    def name_of(self, stage_id: str) -> str:
        for s in self.stages:
            if s.identifier == stage_id:
                return s.display_name
        return stage_id


def analyze(job: ParsedJob) -> WorkflowGraph:
    graph = WorkflowGraph(job=job)
    _classify_roles(job)
    for s in job.stages:
        {ROLE_SOURCE: graph.sources, ROLE_REFERENCE: graph.references,
         ROLE_TRANSFORMATION: graph.transformations, ROLE_TARGET: graph.targets,
         ROLE_UNCLASSIFIED: graph.data_objects}[s.role].append(s)

    graph.paths, graph.paths_truncated, graph.has_cycles, graph.total_path_count = \
        _enumerate_paths(job, graph)
    graph.findings = _build_findings(job, graph)
    return graph


def _classify_roles(job: ParsedJob) -> None:
    """Roles come from link degree - DataStage carries no source/target flag."""
    in_deg = {s.identifier: 0 for s in job.stages}
    out_deg = {s.identifier: 0 for s in job.stages}
    for l in job.links:
        if l.from_stage in out_deg:
            out_deg[l.from_stage] += 1
        if l.to_stage in in_deg:
            in_deg[l.to_stage] += 1

    for s in job.stages:
        i, o = in_deg[s.identifier], out_deg[s.identifier]
        if i == 0 and o == 0:
            s.role = ROLE_UNCLASSIFIED
        elif i == 0:
            out_links = [l for l in job.links if l.from_stage == s.identifier]
            s.role = ROLE_REFERENCE if out_links and all(
                l.is_reference for l in out_links) else ROLE_SOURCE
        elif o == 0:
            s.role = ROLE_TARGET
        else:
            s.role = ROLE_TRANSFORMATION


def _enumerate_paths(job, graph) -> tuple[list[list[str]], bool, bool, int]:
    """Every walk from a source or reference to a target, with three guards:
    cycles, depth, and volume."""
    adjacency: dict[str, list[str]] = {s.identifier: [] for s in job.stages}
    for l in job.links:
        if l.from_stage in adjacency:
            adjacency[l.from_stage].append(l.to_stage)

    names = {s.identifier: s.display_name for s in job.stages}
    targets = {s.identifier for s in graph.targets}
    starts = [s.identifier for s in graph.sources + graph.references]

    paths: list[list[str]] = []
    truncated = cycles = False
    total = 0

    def walk(node: str, trail: list[str], on_path: set[str]) -> None:
        nonlocal truncated, cycles, total
        if len(trail) > MAX_PATH_DEPTH:
            truncated = True
            return
        if node in targets or not adjacency.get(node):
            total += 1
            if len(paths) < MAX_PATHS:
                paths.append([names.get(n, n) for n in trail])
            else:
                truncated = True
            return
        for nxt in adjacency.get(node, []):
            if nxt in on_path:
                cycles = True
                continue
            walk(nxt, trail + [nxt], on_path | {nxt})

    for start in starts:
        walk(start, [start], {start})
    return paths, truncated, cycles, total


def _consistency_problems(job: ParsedJob, graph: WorkflowGraph) -> list[str]:
    """Invariants that should never fail. Better for the report to say something
    is wrong than to print a confident number nobody can trust."""
    problems: list[str] = []
    by_id = {s.identifier: s for s in graph.stages}
    names = {s.display_name for s in graph.stages}

    if len(by_id) != len(graph.stages):
        problems.append("Two or more stages share an identifier.")

    classified = (len(graph.sources) + len(graph.references)
                  + len(graph.transformations) + len(graph.targets))
    if classified + len(graph.data_objects) != len(graph.stages):
        problems.append(
            f"Stage counts disagree: {classified} classified plus "
            f"{len(graph.data_objects)} unclassified is {classified + len(graph.data_objects)}, "
            f"but the inventory lists {len(graph.stages)}.")

    in_deg = {i: 0 for i in by_id}
    out_deg = {i: 0 for i in by_id}
    for l in graph.links:
        if l.from_stage not in by_id or l.to_stage not in by_id:
            problems.append(f"Link {l.display_name} connects a stage the report does not list.")
            break
        out_deg[l.from_stage] += 1
        in_deg[l.to_stage] += 1

    for s in graph.stages:
        if len(s.inputs) != in_deg.get(s.identifier, 0) or \
                len(s.outputs) != out_deg.get(s.identifier, 0):
            problems.append(
                f"The link counts shown for {s.display_name} do not match the link inventory.")
            break

    for s in graph.sources + graph.references:
        if in_deg.get(s.identifier):
            problems.append(f"{s.display_name} is listed as an input but has incoming links.")
            break
    for s in graph.targets:
        if out_deg.get(s.identifier):
            problems.append(f"{s.display_name} is listed as a target but has outgoing links.")
            break

    for path in graph.paths:
        unknown = [n for n in path if n not in names]
        if unknown:
            problems.append(f"A path names {unknown[0]}, which the inventory does not list.")
            break

    return problems


def _build_findings(job: ParsedJob, graph: WorkflowGraph) -> list[Finding]:
    findings: list[Finding] = []

    def add(category: str, obj: str, text: str) -> None:
        findings.append(Finding(f"F-{len(findings) + 1:02d}", category, obj, text))

    for problem in _consistency_problems(job, graph):
        add("Internal inconsistency", TEXT_NOT_APPLICABLE,
            problem + " This report should not be relied on until it is resolved.")

    for w in job.warnings:
        m = re.search(r"\b([A-Za-z0-9_]*S\d+(?:P\d+)?)\b", w)
        add("Unresolved reference", m.group(1) if m else TEXT_NOT_APPLICABLE, w)

    unnamed = [s for s in graph.stages if s.name is None]
    if unnamed:
        add("Incomplete metadata", ", ".join(s.identifier for s in unnamed[:4]),
            f"{len(unnamed)} object(s) carry no technical name in this export and are listed "
            f"as {TEXT_NOT_IDENTIFIED.lower()}.")

    no_object = [s for s in graph.sources + graph.targets if not stage_object(s)]
    if no_object:
        add("Incomplete metadata", ", ".join(s.identifier for s in no_object[:4]),
            f"{len(no_object)} source or target record no table or file name in this export, "
            "so their Object entries read as not identified.")

    for s in graph.data_objects:
        label = s.name or f"The object {s.identifier}"
        add("Isolated object", s.identifier,
            f"{label} declares no links and takes no part in any data path.")

    if job.inferred_types:
        add("Unfamiliar record type", ", ".join(sorted(job.inferred_types)),
            "These record types were not previously known but carry stage properties, so they "
            "are documented as stages. Confirm they are stages rather than definitions.")

    if graph.has_cycles:
        add("Cyclic route", TEXT_NOT_APPLICABLE,
            "Cyclic routes were found. Affected branches were terminated at the point of "
            "repetition, so the path list is not the complete set.")
    if graph.paths_truncated:
        add("Enumeration limit", TEXT_NOT_APPLICABLE,
            f"{graph.total_path_count} paths were found in total and the first "
            f"{len(graph.paths)} are listed.")
    if not graph.has_cycles and not graph.paths_truncated:
        add("Completeness", TEXT_NOT_APPLICABLE,
            "Every record type was recognised, every stage classified, and path enumeration "
            "completed with no cyclic routes.")
    return findings


# --- what a source or target reads from or writes to --------------------------

OBJECT_PROPERTY_CANDIDATES = (
    "TableName", "Table_name", "TableNameInput", "Table", "TargetTable",
    "FileName", "Filename", "File", "FilePath", "Path", "Directory",
    "SelectStatement", "Select_Statement", "SQLStatement", "SQL", "Query",
    "Source", "DataSource", "DatabaseName", "SchemaName",
)
OBJECT_XML_TAGS = ("TableName", "Table", "SelectStatement", "SQL", "Query",
                   "FileName", "File", "Path", "TargetTable")
# MetaBag is absent on purpose: it holds column metadata and encoded internals.
OBJECT_COLLECTION_HINTS = ("Properties", "PropertyList", "StageProperties", "Usage")
COLLECTION_KEY_FIELDS = ("Name", "PropertyName", "Key", "Id")
COLLECTION_VALUE_FIELDS = ("Value", "PropertyValue", "Data", "Text")

_OBJECT_ALLOWED = re.compile(r"^[\w\s.,;:*=<>()\[\]{}'\"/\\$#@%+?!|-]+$")


def _looks_like_object(text: str) -> bool:
    """A table name, path or query has separators. A long run without one is an
    encoded value, not something to print."""
    text = (text or "").strip()
    if not (2 <= len(text) <= 4000):
        return False
    if any(ord(c) < 32 and c not in "\t\n\r" for c in text):
        return False
    if not _OBJECT_ALLOWED.match(text):
        return False
    if sum(c.isalpha() for c in text) < max(2, len(text) * 0.30):
        return False
    unbroken = max((len(r) for r in re.split(r"[\s_./\\-]+", text) if r), default=0)
    return unbroken <= 40


def _tidy_object(text: str) -> str:
    text = " ".join(text.split())
    return text if len(text) <= 160 else text[:159] + TRUNCATION_MARK


def _search_xml_blob(blob: str) -> str | None:
    try:
        from lxml import etree
        root = etree.fromstring(blob.encode("utf-8", "ignore"), etree.XMLParser(recover=True))
    except Exception:
        return None
    if root is None:
        return None
    wanted = {t.lower() for t in OBJECT_XML_TAGS}
    for el in root.iter():
        tag = el.tag.rsplit("}", 1)[-1] if isinstance(el.tag, str) else ""
        if tag.lower() in wanted:
            text = "".join(el.itertext()).strip()
            if text:
                return text
    return None


def _search_collections(stage: Stage) -> str | None:
    wanted = {c.lower() for c in OBJECT_PROPERTY_CANDIDATES}
    fallback = None
    for name, rows in stage.collections.items():
        if not any(h.lower() in name.lower() for h in OBJECT_COLLECTION_HINTS):
            continue
        for row in rows:
            key = next((row[f] for f in COLLECTION_KEY_FIELDS if row.get(f)), None)
            value = next((row[f] for f in COLLECTION_VALUE_FIELDS if row.get(f)), None)
            if not value:
                continue
            if key and key.lower() in wanted and _looks_like_object(value):
                return value
            if "<" in value and fallback is None:
                found = _search_xml_blob(value)
                if found and _looks_like_object(found):
                    fallback = found
    return fallback


def stage_object(stage: Stage) -> str | None:
    """Connector stages keep this in collection rows; older stages in properties."""
    flat = stage.prop(*OBJECT_PROPERTY_CANDIDATES)
    if flat and _looks_like_object(flat):
        return _tidy_object(flat)

    from_collection = _search_collections(stage)
    if from_collection:
        return _tidy_object(from_collection)

    for key, value in stage.properties.items():
        if value and value.lstrip().startswith("<"):
            found = _search_xml_blob(value)
            if found and _looks_like_object(found):
                return _tidy_object(found)
    return None


# =============================================================================
# SECTION D   NARRATIVE
# =============================================================================
# The crew explains parsed facts. It never supplies a value that appears in a
# table, and the report still generates without it.

@dataclass
class Narrative:
    summary: str = ""
    stage_descriptions: dict[str, str] = field(default_factory=dict)
    path_explanations: dict[int, str] = field(default_factory=dict)
    design_findings: list[tuple[str, str]] = field(default_factory=list)
    unverified_identifiers: list[str] = field(default_factory=list)


# The model writes markdown, runs long, and wraps lines. Everything reaching the
# page goes through these first.
_MD_HEADING_RE = re.compile(r"^\s{0,3}#{1,6}\s*", re.M)
_MD_BULLET_RE = re.compile(r"^\s*[-*+]\s+", re.M)
_MD_CODE_RE = re.compile(r"`{1,3}([^`]*)`{1,3}", re.S)
# Underscore markdown is not handled: DataStage names are full of underscores
# and treating them as emphasis rewrites EDW_STG.CUST into EDWSTG.CUST.
_MD_BOLD_RE = re.compile(r"\*\*(.+?)\*\*", re.S)
_MD_ITALIC_RE = re.compile(r"(?<!\*)\*(?!\s|\*)(.+?)(?<!\s)\*(?!\*)", re.S)
_SENTENCE_END_RE = re.compile(r"[.!?](?=\s|$)")


def clean_text(text: Any, markdown: bool = False) -> str:
    """Markdown is stripped only from model-written text: doing it everywhere ate
    the leading # of labels such as '# of sources'."""
    if text is None:
        return ""
    out = str(text)
    if markdown:
        out = _MD_HEADING_RE.sub("", out)
        out = _MD_BULLET_RE.sub("", out)
        out = _MD_CODE_RE.sub(r"\1", out)
        out = _MD_BOLD_RE.sub(r"\1", out)
        out = _MD_ITALIC_RE.sub(r"\1", out)
    out = out.replace(" ", " ")
    out = "".join(c for c in out if c in "\n\t" or ord(c) >= 32)
    out = re.sub(r"[-￰-￿]", "", out)
    lines = [l.strip() for l in out.splitlines() if l.strip()]
    joined = [l + ("." if i < len(lines) - 1 and not l.endswith((".", "!", "?", ":", ";", ",")) else "")
              for i, l in enumerate(lines)]
    return re.sub(r"\s+", " ", " ".join(joined)).strip()


def truncate_at_sentence(text: str, limit: int) -> str:
    if not text or len(text) <= limit:
        return text
    window = text[: limit + 1]
    ends = [m.end() for m in _SENTENCE_END_RE.finditer(window)]
    if ends and ends[-1] >= limit * 0.55:
        return window[: ends[-1]].strip()
    return window.rsplit(" ", 1)[0].rstrip(" ,;:-") + TRUNCATION_MARK


def limit_prose(text: str) -> list[str]:
    if not text:
        return []
    blocks = [clean_text(b, markdown=True)
              for b in re.split(r"\n\s*\n|\n", str(text)) if b.strip()]
    blocks = [b for b in blocks if b][:MAX_PROSE_PARAGRAPHS]
    kept, budget = [], MAX_PROSE_CHARS
    for block in blocks:
        if budget <= 0:
            break
        kept.append(truncate_at_sentence(block, budget) if len(block) > budget else block)
        budget -= len(kept[-1])
    return kept


def _load_config_module():
    import importlib
    for name in LLM_CONFIG_MODULES:
        try:
            return importlib.import_module(name)
        except Exception:
            continue
    return None


def _llm_setting(kind: str, module) -> str | None:
    for attr in LLM_CONFIG_NAMES[kind]:
        if module is not None:
            value = getattr(module, attr, None)
            if isinstance(value, str) and value.strip():
                return value.strip()
        value = os.environ.get(attr)
        if value and value.strip():
            return value.strip()
    return None


def _build_llm():
    module = _load_config_module()
    if module is not None:
        for attr in LLM_INSTANCE_NAMES:
            candidate = getattr(module, attr, None)
            if isinstance(candidate, AzureChatOpenAI):
                log.info("Reusing the AzureChatOpenAI client from %s.%s", module.__name__, attr)
                return candidate

    deployment = _llm_setting("deployment", module)
    if not deployment:
        raise RuntimeError(
            "No Azure deployment name found. Looked for "
            f"{', '.join(LLM_CONFIG_NAMES['deployment'][:4])} in "
            f"{module.__name__ if module else 'the environment'}.")

    kwargs: dict[str, Any] = {
        "azure_deployment": deployment,
        "api_version": _llm_setting("api_version", module) or AZURE_API_VERSION_DEFAULT,
        "temperature": LLM_TEMPERATURE,
        "timeout": LLM_TIMEOUT_SECONDS,
    }
    for key, kind in (("azure_endpoint", "endpoint"), ("api_key", "api_key")):
        value = _llm_setting(kind, module)
        if value:
            kwargs[key] = value
    log.info("Azure client: deployment=%s (from %s)", deployment,
             module.__name__ if module else "environment")
    return AzureChatOpenAI(**kwargs)


GROUNDING_RULES = f"""
Rules that override every other instruction:
- Use ONLY the names in the factsheet. Never invent a stage, link, table, file or
  column name. If a fact is not there, write "not specified in the export".
- Measured business English. Complete sentences, no marketing language, no
  metaphors, no bullet fragments.
- Do not mention the factsheet, the model or this report.
- Plain text inside every JSON string: no markdown, no newlines.
- Respect the length limits. Text over the limit is trimmed, so writing more only
  loses your closing sentence.
- Return valid JSON only, with no fences and no commentary.
"""


def build_factsheet(graph: WorkflowGraph) -> dict[str, Any]:
    """A trimmed view for the model. all_records is excluded: for a real job it
    is tens of thousands of tokens of port-identifier noise."""
    def entry(s: Stage) -> dict[str, Any]:
        return {"name": s.display_name, "id": s.identifier, "type": s.display_type,
                "role": ROLE_LABELS[s.role], "inputs": s.inputs, "outputs": s.outputs,
                "object": stage_object(s), "key_properties": _interesting_properties(s)}

    return {
        "job_name": graph.job.job_name,
        "source_file": graph.job.source_file,
        "sources": [entry(s) for s in graph.sources],
        "references": [entry(s) for s in graph.references],
        "transformations": [entry(s) for s in graph.transformations],
        "targets": [entry(s) for s in graph.targets],
        "links": [{"name": l.display_name, "from": graph.name_of(l.from_stage),
                   "to": graph.name_of(l.to_stage), "kind": l.kind} for l in graph.links],
        "paths": [{"index": i + 1, "stages": p} for i, p in enumerate(graph.paths)],
        "designer_notes": graph.job.annotations[:20],
    }


INTERESTING_PROPERTY_NAMES = (
    "StageType", "TableName", "SelectStatement", "Query", "SQL", "FileName",
    "ReadMethod", "WriteMode", "WriteMethod", "UpdateAction", "Partitioning",
    "LookupType", "Condition", "Constraint", "Derivation", "JoinType",
    "BusinessKey", "SurrogateKey", "SortKey", "FunnelType",
)


def _interesting_properties(stage: Stage, limit: int = 8) -> dict[str, str]:
    out: dict[str, str] = {}
    for key in INTERESTING_PROPERTY_NAMES:
        value = stage.properties.get(key)
        if value:
            out[key] = value[:300]
        if len(out) >= limit:
            break
    return out


def known_identifiers(graph: WorkflowGraph) -> set[str]:
    known = {s.display_name for s in graph.stages}
    known |= {l.display_name for l in graph.links}
    known |= {s.display_type for s in graph.stages}
    known |= {stage_object(s) or "" for s in graph.stages}
    known.add(graph.job.job_name)
    return {k for k in known if k}


def generate_narrative(graph: WorkflowGraph) -> Narrative:
    """Returns an empty Narrative rather than raising: a factual report without
    prose is useful, a failed download is not."""
    if not LLM_ENABLED:
        log.info("LLM disabled; tables only.")
        return Narrative()

    factsheet = build_factsheet(graph)
    raw: dict[str, Any] = {}
    for attempt in range(LLM_MAX_RETRIES + 1):
        try:
            raw = _run_crew(factsheet)
            break
        except Exception as exc:
            log.warning("Narrative attempt %s failed: %s", attempt + 1, exc)
    if not raw:
        log.warning("Narrative unavailable; continuing with tables only.")
        return Narrative()

    narrative = _assemble_narrative(raw)
    narrative.unverified_identifiers = _validate_narrative(narrative, graph)
    if narrative.unverified_identifiers:
        log.warning("Narrative names identifiers absent from the export: %s",
                    ", ".join(narrative.unverified_identifiers))
    return narrative


def _run_crew(factsheet: dict[str, Any]) -> dict[str, Any]:
    llm = _build_llm()
    facts = json.dumps(factsheet, indent=1, ensure_ascii=False)
    allowed = sorted({e["name"] for k in ("sources", "references", "transformations", "targets")
                      for e in factsheet[k]} | {l["name"] for l in factsheet["links"]})
    allowed_block = ", ".join(allowed)

    architect = Agent(
        role="ETL Solution Architect",
        goal="Explain how a workflow is put together, accurately and without embellishment.",
        backstory="You document data integration workflows for engineers who did not "
                  "write the job they are reading about.",
        llm=llm, verbose=False, allow_delegation=False)
    documenter = Agent(
        role="Technical Documentation Specialist",
        goal="Describe individual stages in one or two precise sentences each.",
        backstory="You write reference documentation. You never speculate beyond the "
                  "metadata you are given and keep every entry the same shape.",
        llm=llm, verbose=False, allow_delegation=False)

    summary_task = Task(
        description=(f"{GROUNDING_RULES}\nPermitted names: {allowed_block}\n\n"
                     f"Factsheet:\n{facts}\n\n"
                     f"Write a summary of what this workflow does, at most "
                     f"{MAX_PROSE_CHARS} characters: where records enter, how they are "
                     "combined and transformed, where the flow branches and on what "
                     "condition, and where records end up.\n\n"
                     'Return JSON: {"summary": "text"}.'),
        expected_output='JSON with key "summary".', agent=architect)

    # One call for every stage. A Task per stage re-sends the whole context and
    # turns a 40-stage job into 40 round trips.
    stages_task = Task(
        description=(f"{GROUNDING_RULES}\nPermitted names: {allowed_block}\n\n"
                     f"Factsheet:\n{facts}\n\n"
                     "Describe EVERY stage listed under sources, references, transformations "
                     f"and targets. One or two sentences and at most {MAX_DESCRIPTION_CHARS} "
                     "characters each, stating what the stage does in the flow and anything "
                     "operationally significant in its properties.\n\n"
                     'Return JSON mapping each stage name to its description.'),
        expected_output="JSON mapping every stage name to a description.", agent=documenter)

    paths_task = Task(
        description=(f"{GROUNDING_RULES}\nPermitted names: {allowed_block}\n\n"
                     f"Factsheet:\n{facts}\n\n"
                     "For each path, write one or two sentences and at most "
                     f"{MAX_DESCRIPTION_CHARS} characters on what that route represents. "
                     "Where a path begins at a reference stage, say it represents data "
                     "influence rather than record movement.\n\n"
                     "Then list any design observations worth a reader's attention - an "
                     "inner join that silently drops records, a flow with no connection to "
                     "the rest of the job, an overwrite that loses history. Only what the "
                     "factsheet supports.\n\n"
                     'Return JSON: {"paths": {"1": "text"}, '
                     '"findings": [{"object": "STAGE_NAME", "text": "text"}]}.'),
        expected_output='JSON with keys "paths" and "findings".', agent=architect)

    crew = Crew(agents=[architect, documenter],
                tasks=[summary_task, stages_task, paths_task],
                process=Process.sequential, verbose=False)
    crew.kickoff()

    return {"summary": _json_from_task(summary_task),
            "stages": _json_from_task(stages_task),
            "paths": _json_from_task(paths_task)}


def _task_text(task) -> str:
    """CrewAI has moved this between versions; try the known shapes."""
    out = getattr(task, "output", None)
    if out is None:
        return ""
    for attr in ("raw", "raw_output", "exported_output", "result"):
        value = getattr(out, attr, None)
        if isinstance(value, str) and value.strip():
            return value
    return str(out)


def _json_from_task(task) -> dict[str, Any]:
    return _loads_loose(_task_text(task))


def _loads_loose(text: str) -> dict[str, Any]:
    if not text:
        return {}
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip(), flags=re.M).strip()
    for candidate in (text, text[text.find("{"): text.rfind("}") + 1] if "{" in text else ""):
        try:
            value = json.loads(candidate)
            if isinstance(value, dict):
                return value
        except (json.JSONDecodeError, ValueError):
            continue
    return {}


def _assemble_narrative(raw: dict[str, Any]) -> Narrative:
    paths_block = raw.get("paths", {}) or {}
    explanations: dict[int, str] = {}
    for key, value in (paths_block.get("paths") or {}).items():
        try:
            explanations[int(key)] = clean_text(value, markdown=True)
        except (TypeError, ValueError):
            continue

    findings = []
    for item in paths_block.get("findings") or []:
        if isinstance(item, dict) and item.get("text"):
            findings.append((clean_text(item.get("object", ""), markdown=True),
                             clean_text(item["text"], markdown=True)))

    return Narrative(
        summary=str((raw.get("summary", {}) or {}).get("summary", "") or ""),
        stage_descriptions={str(k): clean_text(v, markdown=True)
                            for k, v in (raw.get("stages", {}) or {}).items()
                            if str(v).strip()},
        path_explanations=explanations,
        design_findings=findings,
    )


def _validate_narrative(narrative: Narrative, graph: WorkflowGraph) -> list[str]:
    """A model asked to describe an ETL job will produce plausible table and
    column names. Anything unrecognised is reported rather than shipped."""
    known = {k.upper() for k in known_identifiers(graph)}
    corpus = " ".join([narrative.summary]
                      + list(narrative.stage_descriptions.values())
                      + list(narrative.path_explanations.values())
                      + [t for _, t in narrative.design_findings])
    candidates = set(re.findall(r"\b[A-Z][A-Z0-9]*(?:_[A-Z0-9]+)+\b", corpus))
    return sorted(c for c in candidates if not any(c.upper() in k for k in known))


# =============================================================================
# SECTION E   RENDERING
# =============================================================================
# HTML and CSS rather than Word, so the printed layout is the same everywhere.

_GLYPHS = {
    "box": dict(tl="┌", tr="┐", bl="└", br="┘", h="─", v="│",
                stem="┬", tee="├", elbow="└", down="▼", right="▶"),
    "ascii": dict(tl="+", tr="+", bl="+", br="+", h="-", v="|",
                  stem="+", tee="+", elbow="+", down="v", right=">"),
}


def _ellipsize(text: str, limit: int) -> str:
    return text if len(text) <= limit else text[: limit - 1] + TRUNCATION_MARK


def _stage_depths(graph: WorkflowGraph) -> dict[str, int]:
    depth: dict[str, int] = {}

    def compute(sid: str, seen: frozenset[str]) -> int:
        if sid in depth:
            return depth[sid]
        prior = [l.from_stage for l in graph.links
                 if l.to_stage == sid and l.from_stage not in seen]
        depth[sid] = 0 if not prior else 1 + max(compute(p, seen | {sid}) for p in prior)
        return depth[sid]

    for stage in graph.stages:
        compute(stage.identifier, frozenset())
    return depth


def build_flowchart(graph: WorkflowGraph, only: Sequence[Stage] | None = None) -> str:
    """Top to bottom: a box per stage, arrows for links.

    Vertical rather than left to right because a portrait page is about eighty
    characters wide, which a chain of connector-length names overruns.
    """
    stages = [s for s in (only if only is not None else graph.stages)
              if s.role != ROLE_UNCLASSIFIED]
    if not stages or not graph.links or len(stages) > MAX_STAGES_FOR_DIAGRAM:
        return ""

    depth = _stage_depths(graph)
    order = sorted(stages, key=lambda s: (depth.get(s.identifier, 0), s.display_name))
    shown = {s.identifier for s in order}
    position = {s.identifier: i for i, s in enumerate(order)}

    g = _GLYPHS.get(DIAGRAM_STYLE, _GLYPHS["box"])
    inner = min(max(len(s.display_name) for s in order) + 2, MAX_DIAGRAM_LINE_CHARS - 26)
    stem = 4
    lines: list[str] = []

    for index, stage in enumerate(order):
        label = _ellipsize(stage.display_name, inner - 2)
        lines.append(f"{g['tl']}{g['h'] * inner}{g['tr']}")
        lines.append(f"{g['v']} {label:<{inner - 2}} {g['v']}  {ROLE_LABELS[stage.role].lower()}")

        outgoing = [l for l in graph.outgoing(stage) if l.to_stage in shown]
        if not outgoing:
            lines.append(f"{g['bl']}{g['h'] * inner}{g['br']}")
            lines.append("")
            continue

        lines.append(f"{g['bl']}{g['h'] * stem}{g['stem']}{g['h'] * (inner - stem - 1)}{g['br']}")
        if len(outgoing) == 1 and position.get(outgoing[0].to_stage) == index + 1:
            lines.append(f"{' ' * stem}{g['v']}  {outgoing[0].display_name}")
            lines.append(f"{' ' * stem}{g['down']}")
        else:
            width = max(len(graph.name_of(l.to_stage)) for l in outgoing)
            for i, link in enumerate(outgoing):
                joint = g["elbow"] if i == len(outgoing) - 1 else g["tee"]
                tail = "  (reference)" if link.is_reference else ""
                lines.append(f"{' ' * stem}{joint}{g['h']}{g['right']} "
                             f"{graph.name_of(link.to_stage):<{width}}   "
                             f"{link.display_name}{tail}")
            lines.append("")

    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines)


def build_main_chain(graph: WorkflowGraph) -> list[Stage]:
    """The longest source-to-target route, for the orientation diagram."""
    if not graph.paths:
        return []
    longest = max(graph.paths, key=len)[:GLANCE_DIAGRAM_STAGES]
    by_name = {s.display_name: s for s in graph.stages}
    return [by_name[n] for n in longest if n in by_name]


def _cell(value: Any, limit: int | None = None, markdown: bool = False) -> dict[str, Any]:
    """A table cell: text plus whether it is a missing-value placeholder."""
    if value is None or value == "":
        return {"text": TEXT_NOT_IDENTIFIED, "missing": True}
    text = clean_text(value, markdown=markdown)
    if not text:
        return {"text": TEXT_NOT_IDENTIFIED, "missing": True}
    if limit:
        text = truncate_at_sentence(text, limit)
    return {"text": text, "missing": False}


def build_context(graph: WorkflowGraph, narrative: Narrative) -> dict[str, Any]:
    """Everything the template needs. All formatting decisions live here."""
    from datetime import datetime

    def described(stage: Stage) -> dict[str, Any]:
        return _cell(narrative.stage_descriptions.get(stage.display_name), MAX_DESCRIPTION_CHARS)

    def stage_row(s: Stage) -> dict[str, Any]:
        return {"id": s.identifier,
                "name": _cell(s.name),
                "type": s.display_type,
                "object": _cell(stage_object(s) if s.role in
                                (ROLE_SOURCE, ROLE_REFERENCE, ROLE_TARGET) else TEXT_NOT_APPLICABLE),
                "inputs": len(s.inputs), "outputs": len(s.outputs)}

    groups = [("Sources", graph.sources), ("References", graph.references),
              ("Transformations", graph.transformations), ("Targets", graph.targets),
              ("Unclassified", graph.data_objects)]

    detail = []
    for stage in graph.transformations + graph.sources + graph.references + graph.targets:
        props = {k: v for k, v in _interesting_properties(stage, limit=6).items()
                 if k != "StageType"}
        description = described(stage)
        if not props and description["missing"]:
            continue
        detail.append({"name": stage.display_name, "type": stage.display_type,
                       "id": stage.identifier, "role": ROLE_LABELS[stage.role],
                       "properties": props, "description": description})

    findings = [{"ref": f.ref, "category": f.category, "obj": _cell(f.obj), "text": f.text}
                for f in graph.findings]
    for obj, text in narrative.design_findings:
        findings.insert(0, {"ref": "", "category": "Design",
                            "obj": _cell(obj or TEXT_NOT_APPLICABLE),
                            "text": truncate_at_sentence(text, MAX_DESCRIPTION_CHARS)})
    for ident in narrative.unverified_identifiers:
        findings.append({"ref": "", "category": "Unverified reference",
                         "obj": _cell(ident),
                         "text": f"The descriptive text names {ident}, which does not appear "
                                 "in the export and could not be verified."})
    for i, item in enumerate(findings, start=1):
        item["ref"] = f"F-{i:02d}"

    return {
        "title": REPORT_TITLE,
        "subtitle": REPORT_SUBTITLE,
        "job_name": graph.job.job_name,
        "source_file": graph.job.source_file,
        "generated": datetime.now().strftime("%d %B %Y"),
        "text_not_identified": TEXT_NOT_IDENTIFIED,
        "mark_generated": MARK_GENERATED_PROSE,
        "summary": limit_prose(narrative.summary),
        "basis": _basis_paragraphs(graph, narrative),
        "facts": [
            ("Stages", len(graph.stages)), ("Links", len(graph.links)),
            ("Sources", len(graph.sources)), ("References", len(graph.references)),
            ("Transformations", len(graph.transformations)), ("Targets", len(graph.targets)),
            ("Paths", graph.total_path_count), ("Findings", len(findings)),
        ],
        "glance_chart": build_flowchart(graph, build_main_chain(graph)),
        "full_chart": build_flowchart(graph),
        "groups": [(label, [stage_row(s) for s in members])
                   for label, members in groups if members],
        "links": [{"name": l.display_name, "from": graph.name_of(l.from_stage),
                   "to": graph.name_of(l.to_stage), "kind": l.kind,
                   "columns": l.column_count if l.column_count is not None else "—"}
                  for l in graph.links],
        "detail": detail,
        "paths": [{"index": i, "route": " → ".join(p), "length": len(p),
                   "explanation": _cell(narrative.path_explanations.get(i),
                                        MAX_DESCRIPTION_CHARS)}
                  for i, p in enumerate(graph.paths, start=1)],
        "findings": findings,
        "sections": _present_sections(graph, narrative, detail, findings),
    }


def _present_sections(graph, narrative, detail, findings) -> list[str]:
    """A section with no rows is left out and the rest renumber."""
    candidates = [
        ("At a Glance", True),
        ("Stage Inventory", bool(graph.stages)),
        ("Links", bool(graph.links)),
        ("Stage Detail", bool(detail)),
        ("Data Paths", bool(graph.paths)),
        ("Findings", bool(findings)),
    ]
    return [n for n, present in candidates if present or not SKIP_EMPTY_SECTIONS]


def _basis_paragraphs(graph: WorkflowGraph, narrative: Narrative) -> list[str]:
    """What the parser could and could not determine. Silent omission is the
    failure mode this exists to prevent."""
    out = [
        f"All {len(graph.stages)} stages and {len(graph.links)} links in this export were "
        f"resolved, and {len(graph.stages) - len(graph.data_objects)} of them were classified "
        "by role. "
        + ("Path enumeration completed within its limits and found no cyclic routes."
           if not (graph.paths_truncated or graph.has_cycles)
           else f"Path enumeration reported {graph.total_path_count} paths in total.")
    ]
    gaps = []
    if graph.job.warnings:
        gaps.append(f"{len(graph.job.warnings)} link(s) name a partner outside this job, "
                    "which indicates a shared container; those are excluded from the "
                    "flowchart and from path enumeration")
    missing = [s for s in graph.sources + graph.targets if not stage_object(s)]
    if missing:
        gaps.append(f"{len(missing)} source or target record no table or file name in this "
                    "export, so their Object entries read as not identified")
    if graph.job.inferred_types:
        gaps.append("record types "
                    + ", ".join(sorted(graph.job.inferred_types))
                    + " were not previously known but carry stage properties, so they are "
                      "documented as stages")
    if gaps:
        out.append("Not everything could be determined: " + "; ".join(gaps) + ".")

    if narrative.stage_descriptions or narrative.summary:
        out.append("Descriptive text is generated and marked with a rule. Every table value "
                   "is read directly from the export.")
    else:
        out.append("Descriptive text was not generated for this report, so descriptions read "
                   "as not identified. Every table value is read directly from the export.")
    return out


PAGE_SIZES = {"A4": A4, "LETTER": LETTER}


def _hex(value: str):
    return colors.HexColor(value)


def _styles() -> dict[str, ParagraphStyle]:
    """Paragraph styles built from the tunables in Section A."""
    base = ParagraphStyle("body", fontName=FONT_SERIF, fontSize=SIZE_BODY,
                          leading=SIZE_BODY * 1.5, textColor=_hex(COLOR_INK),
                          alignment=TA_JUSTIFY, spaceAfter=4)
    return {
        "body": base,
        "eyebrow": ParagraphStyle("eyebrow", fontName=FONT_MONO_BOLD,
                                  fontSize=SIZE_SECTION_NUMBER,
                                  leading=SIZE_SECTION_NUMBER * 1.4,
                                  textColor=_hex(COLOR_CORAL_DEEP), spaceAfter=2),
        "title": ParagraphStyle("title", fontName=FONT_SANS_BOLD,
                                fontSize=SIZE_SECTION_TITLE,
                                leading=SIZE_SECTION_TITLE * 1.2,
                                textColor=_hex(COLOR_CORAL), spaceAfter=3),
        "sub": ParagraphStyle("sub", fontName=FONT_MONO_BOLD,
                              fontSize=SIZE_SECTION_NUMBER,
                              leading=SIZE_SECTION_NUMBER * 1.5,
                              textColor=_hex(COLOR_CORAL_DEEP),
                              spaceBefore=8, spaceAfter=3),
        "cell": ParagraphStyle("cell", fontName=FONT_SERIF, fontSize=SIZE_TABLE,
                               leading=SIZE_TABLE * 1.35, textColor=_hex(COLOR_INK)),
        "cell_right": ParagraphStyle("cell_right", fontName=FONT_MONO,
                                     fontSize=SIZE_TABLE_HEAD,
                                     leading=SIZE_TABLE_HEAD * 1.45, alignment=2,
                                     textColor=_hex(COLOR_INK)),
        "head_right": ParagraphStyle("head_right", fontName=FONT_MONO_BOLD,
                                     fontSize=SIZE_TABLE_HEAD,
                                     leading=SIZE_TABLE_HEAD * 1.3, alignment=2,
                                     textColor=_hex(COLOR_INK_DIM)),
        "cell_mono": ParagraphStyle("cell_mono", fontName=FONT_MONO,
                                    fontSize=SIZE_TABLE_HEAD,
                                    leading=SIZE_TABLE_HEAD * 1.45,
                                    textColor=_hex(COLOR_INK)),
        "cell_missing": ParagraphStyle("cell_missing", fontName=FONT_SERIF_ITALIC,
                                       fontSize=SIZE_TABLE, leading=SIZE_TABLE * 1.35,
                                       textColor=_hex(COLOR_INK_FAINT)),
        "head": ParagraphStyle("head", fontName=FONT_MONO_BOLD, fontSize=SIZE_TABLE_HEAD,
                               leading=SIZE_TABLE_HEAD * 1.3,
                               textColor=_hex(COLOR_INK_DIM)),
        "band": ParagraphStyle("band", fontName=FONT_MONO_BOLD, fontSize=SIZE_TABLE_HEAD,
                               leading=SIZE_TABLE_HEAD * 1.3,
                               textColor=_hex(COLOR_INK_DIM)),
        "caption": ParagraphStyle("caption", fontName=FONT_MONO_BOLD, fontSize=SIZE_CAPTION,
                                  leading=SIZE_CAPTION * 1.4,
                                  textColor=_hex(COLOR_CORAL_DEEP), spaceBefore=3,
                                  spaceAfter=10),
        "panel": ParagraphStyle("panel", parent=base, fontSize=SIZE_TABLE,
                                leading=SIZE_TABLE * 1.5, spaceAfter=4),
        "toc_num": ParagraphStyle("toc_num", fontName=FONT_MONO_BOLD, fontSize=SIZE_CAPTION,
                                  leading=SIZE_BODY * 1.4,
                                  textColor=_hex(COLOR_CORAL_DEEP)),
        "toc_name": ParagraphStyle("toc_name", fontName=FONT_SERIF, fontSize=SIZE_BODY,
                                   leading=SIZE_BODY * 1.4, textColor=_hex(COLOR_INK)),
        "cover_title": ParagraphStyle("cover_title", fontName=FONT_SANS_BOLD, fontSize=22,
                                      leading=26, alignment=1, textColor=_hex(COLOR_INK)),
        "cover_sub": ParagraphStyle("cover_sub", fontName=FONT_SERIF, fontSize=11,
                                    leading=15, alignment=1,
                                    textColor=_hex(COLOR_INK_DIM), spaceBefore=6),
        "cover_field": ParagraphStyle("cover_field", fontName=FONT_MONO, fontSize=SIZE_BODY,
                                      leading=SIZE_BODY * 2, alignment=1,
                                      textColor=_hex(COLOR_INK_DIM)),
        "fact_label": ParagraphStyle("fact_label", fontName=FONT_MONO, fontSize=SIZE_CAPTION,
                                     leading=SIZE_CAPTION * 1.4,
                                     textColor=_hex(COLOR_INK_FAINT)),
        "fact_value": ParagraphStyle("fact_value", fontName=FONT_SANS_BOLD, fontSize=13,
                                     leading=15, textColor=_hex(COLOR_INK)),
        "figcap": ParagraphStyle("figcap", fontName=FONT_MONO_BOLD, fontSize=SIZE_CAPTION,
                                 leading=SIZE_CAPTION * 1.4,
                                 textColor=_hex(COLOR_CORAL_DEEP), spaceBefore=3,
                                 spaceAfter=10),
        "chart": ParagraphStyle("chart", fontName=FONT_MONO, fontSize=SIZE_DIAGRAM,
                                leading=SIZE_DIAGRAM * 1.35, textColor=_hex(COLOR_INK)),
    }


def _label(text: str) -> str:
    """Small uppercase label. ReportLab has no letter-spacing, so the design's
    tracking is dropped rather than faked by inserting spaces between letters -
    that ran the words together."""
    return _escape(text).upper()


class _NumberedCanvas(Canvas):
    """Two passes, so the footer can say 'Page 3 of 9'."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved = []

    def showPage(self):
        self._saved.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._saved)
        for state in self._saved:
            self.__dict__.update(state)
            if self._pageNumber > 1:                 # the cover carries nothing
                self._draw_furniture(total)
            super().showPage()
        super().save()

    def _draw_furniture(self, total: int) -> None:
        width, height = self._pagesize
        side = MARGIN_SIDE_MM * mm
        top = height - MARGIN_TOP_MM * mm
        bottom = MARGIN_BOTTOM_MM * mm

        self.setFont(FONT_MONO, SIZE_RUNNING)
        self.setFillColor(_hex(COLOR_INK_FAINT))
        self.drawString(side, top + 5 * mm, self._header_left)
        self.drawRightString(width - side, top + 5 * mm, self._header_right)
        self.setStrokeColor(_hex(COLOR_RULE))
        self.setLineWidth(0.4)
        self.line(side, top + 3.6 * mm, width - side, top + 3.6 * mm)

        self.line(side, bottom - 4 * mm, width - side, bottom - 4 * mm)
        self.drawString(side, bottom - 7.5 * mm, self._footer_left)
        self.drawRightString(width - side, bottom - 7.5 * mm,
                             f"Page {self._pageNumber} of {total}")


def _rule(width: float, colour: str, thickness: float = 0.7) -> Table:
    """A horizontal line as a flowable."""
    t = Table([[""]], colWidths=[width], rowHeights=[0.1])
    t.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), thickness, _hex(colour)),
                           ("TOPPADDING", (0, 0), (-1, -1), 0),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    return t


def _section_heading(number: int | None, name: str, st, width: float) -> list:
    label = f"Section {number}" if number else "Table of"
    return [KeepTogether([
        Paragraph(_label(label), st["eyebrow"]),
        Paragraph(name, st["title"]),
        _rule(width, COLOR_CORAL_LINE, 0.7),
        Spacer(1, 7),
    ])]


def _generated(paragraphs: Sequence[str], st, width: float) -> list:
    """Model-written text carries a rule; parsed values never do."""
    if not paragraphs:
        return []
    body = [Paragraph(p, st["body"]) for p in paragraphs]
    if not MARK_GENERATED_PROSE:
        return body + [Spacer(1, 4)]
    t = Table([[body]], colWidths=[width])
    t.setStyle(TableStyle([
        ("LINEBEFORE", (0, 0), (0, -1), 1.2, _hex(COLOR_CORAL_LINE)),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return [t, Spacer(1, 8)]


def _panel(paragraphs: Sequence[str], st, width: float) -> list:
    inner = [Paragraph(p, st["panel"]) for p in paragraphs]
    t = Table([[inner]], colWidths=[width])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), _hex(COLOR_PANEL_BG)),
        ("BOX", (0, 0), (-1, -1), 0.6, _hex(COLOR_PANEL_LINE)),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return [t]


def _chart_block(text: str, st, width: float, caption: str) -> list:
    if not text:
        return []
    body = Preformatted(text, st["chart"])
    t = Table([[body]], colWidths=[width])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), _hex(COLOR_KEY_BG)),
        ("BOX", (0, 0), (-1, -1), 0.6, _hex(COLOR_RULE)),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return [t, Paragraph(_label(caption), st["figcap"])]


def _cell_para(value: Any, st, mono: bool = False, right: bool = False):
    """A table cell. Missing values are italic and faint."""
    if isinstance(value, dict):
        text, missing = value["text"], value["missing"]
    else:
        text, missing = str(value), False
    if missing:
        style = st["cell_missing"]
    elif right:
        style = st["cell_right"]
    else:
        style = st["cell_mono"] if mono else st["cell"]
    return Paragraph(_escape(text), style)


def _escape(text: str) -> str:
    return (str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _data_table(headers: Sequence[str], rows: Sequence[Sequence[Any]],
                widths: Sequence[float], st, mono: Sequence[int] = (),
                right: Sequence[int] = (), bands: dict[int, str] | None = None) -> Table:
    """Header repeats across pages; column widths are fixed in points."""
    mono, right = set(mono), set(right)
    bands = bands or {}
    data = [[Paragraph(_label(h), st["head_right"] if i in right else st["head"])
             for i, h in enumerate(headers)]]
    for row in rows:
        data.append([_cell_para(v, st, mono=i in mono, right=i in right)
                     for i, v in enumerate(row)])

    style = [
        ("BACKGROUND", (0, 0), (-1, 0), _hex(COLOR_HEAD_BG)),
        ("BOX", (0, 0), (-1, 0), 0.6, _hex(COLOR_RULE)),
        ("INNERGRID", (0, 0), (-1, 0), 0.6, _hex(COLOR_RULE)),
        ("GRID", (0, 1), (-1, -1), 0.6, _hex(COLOR_RULE_SOFT)),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    for row_index, label in bands.items():
        data[row_index] = [Paragraph(_label(label), st["band"])] + [""] * (len(headers) - 1)
        style += [("SPAN", (0, row_index), (-1, row_index)),
                  ("BACKGROUND", (0, row_index), (-1, row_index), _hex(COLOR_BAND_BG))]

    table = Table(data, colWidths=list(widths), repeatRows=1)
    table.setStyle(TableStyle(style))
    return table


def _facts_table(facts: Sequence[tuple[str, Any]], st, width: float) -> Table:
    per_row = 4
    cell_w = width / per_row
    rows = []
    for i in range(0, len(facts), per_row):
        chunk = list(facts[i:i + per_row])
        chunk += [("", "")] * (per_row - len(chunk))
        rows.append([[Paragraph(_label(label), st["fact_label"]),
                      Paragraph(str(value), st["fact_value"])] if label else ""
                     for label, value in chunk])
    table = Table(rows, colWidths=[cell_w] * per_row)
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, _hex(COLOR_RULE_SOFT)),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    return table


def build_story(context: dict[str, Any], width: float) -> list:
    """The report as a flat list of flowables."""
    st = _styles()
    story: list = []

    # Cover
    story += [Spacer(1, 200),
              Paragraph(context["title"], st["cover_title"]),
              Paragraph(context["subtitle"], st["cover_sub"]),
              Spacer(1, 16)]
    hr = Table([[""]], colWidths=[60])
    hr.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 1.4, _hex(COLOR_CORAL))]))
    hr.hAlign = "CENTER"
    story += [hr, Spacer(1, 16),
              Paragraph(context["job_name"], st["cover_field"]),
              Paragraph(context["source_file"], st["cover_field"]),
              Paragraph(context["generated"], st["cover_field"]),
              PageBreak()]

    sections = context["sections"]

    # Contents and basis
    story += _section_heading(None, "Contents", st, width)
    toc_rows = [[Paragraph(_label(f"Section {i}"), st["toc_num"]),
                 Paragraph(name, st["toc_name"])]
                for i, name in enumerate(sections, start=1)]
    toc = Table(toc_rows, colWidths=[70, width - 70])
    toc.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -1), 0.6, _hex(COLOR_RULE_SOFT)),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story += [toc, Spacer(1, 14),
              Paragraph(_label("Basis of this report"), st["sub"])]
    story += _panel(context["basis"], st, width)

    number = 0
    for name in sections:
        number += 1
        if SECTION_STARTS_NEW_PAGE:
            story.append(PageBreak())
        story += _section_heading(number, name, st, width)
        story += _SECTION_BUILDERS[name](context, st, width)
    return story


def _s_glance(ctx, st, width) -> list:
    out = _generated(ctx["summary"], st, width)
    out += _chart_block(ctx["glance_chart"], st, width,
                        "Figure 1 · Main flow"
                        + (" — full chart in the Links section"
                           if ctx["full_chart"] != ctx["glance_chart"] else ""))
    out.append(_facts_table(ctx["facts"], st, width))
    return out


def _s_inventory(ctx, st, width) -> list:
    rows, bands = [], {}
    for label, members in ctx["groups"]:
        bands[len(rows) + 1] = f"{label} — {len(members)}"
        rows.append([""] * 6)
        for r in members:
            rows.append([r["id"], r["name"], r["type"], r["object"],
                         r["inputs"], r["outputs"]])
    w = [0.10, 0.23, 0.18, 0.31, 0.09, 0.09]
    return [_data_table(["ID", "Technical name", "Type", "Object", "In", "Out"],
                        rows, [width * x for x in w], st, mono=[0, 1, 2, 3],
                        right=[4, 5], bands=bands),
            Paragraph(_label("Table 1 · Stage inventory"), st["caption"])]


def _s_links(ctx, st, width) -> list:
    rows = [[l["name"], l["from"], l["to"], l["kind"], l["columns"]] for l in ctx["links"]]
    w = [0.21, 0.25, 0.25, 0.15, 0.14]
    out = [_data_table(["Link", "From", "To", "Kind", "Columns"],
                       rows, [width * x for x in w], st, mono=[0, 1, 2], right=[4]),
           Paragraph(_label("Table 2 · Link inventory"), st["caption"])]
    if ctx["full_chart"]:
        out.append(Paragraph(_label("Full flowchart"), st["sub"]))
        out += _chart_block(ctx["full_chart"], st, width, "Figure 2 · Every stage and link")
    return out


def _s_detail(ctx, st, width) -> list:
    out = []
    for stage in ctx["detail"]:
        block = [Paragraph(_label(f"{stage['name']} · {stage['type']} · {stage['id']}"),
                           st["sub"])]
        if stage["properties"]:
            rows = [[k, v] for k, v in stage["properties"].items()]
            table = _data_table(["Property", "Value"], rows,
                                [width * 0.32, width * 0.68], st, mono=[1])
            table.setStyle(TableStyle([
                ("BACKGROUND", (0, 1), (0, -1), _hex(COLOR_KEY_BG))]))
            block.append(table)
        if not stage["description"]["missing"]:
            block += _generated([stage["description"]["text"]], st, width)
        else:
            block.append(Spacer(1, 8))
        out.append(KeepTogether(block))
    return out


def _s_paths(ctx, st, width) -> list:
    rows = [[p["index"], p["route"], p["length"], p["explanation"]] for p in ctx["paths"]]
    w = [0.05, 0.36, 0.10, 0.49]
    return [_data_table(["#", "Route", "Stages", "What it represents"],
                        rows, [width * x for x in w], st, mono=[1], right=[0, 2]),
            Paragraph(_label("Table 3 · Data paths"), st["caption"])]


def _s_findings(ctx, st, width) -> list:
    rows = [[f["ref"], f["category"], f["obj"], f["text"]] for f in ctx["findings"]]
    w = [0.09, 0.17, 0.21, 0.53]
    return [_data_table(["Ref", "Category", "Object", "Finding"],
                        rows, [width * x for x in w], st, mono=[0, 2]),
            Paragraph(_label("Table 4 · Findings"), st["caption"])]


_SECTION_BUILDERS = {
    "At a Glance": _s_glance,
    "Stage Inventory": _s_inventory,
    "Links": _s_links,
    "Stage Detail": _s_detail,
    "Data Paths": _s_paths,
    "Findings": _s_findings,
}


def render_pdf(context: dict[str, Any], pdf_path: Path) -> Path:
    page = PAGE_SIZES.get(PAGE_SIZE.upper(), A4)
    width = page[0] - 2 * MARGIN_SIDE_MM * mm

    doc = BaseDocTemplate(
        str(pdf_path), pagesize=page,
        leftMargin=MARGIN_SIDE_MM * mm, rightMargin=MARGIN_SIDE_MM * mm,
        topMargin=MARGIN_TOP_MM * mm, bottomMargin=MARGIN_BOTTOM_MM * mm,
        title=f"{context['title']} — {context['job_name']}", author=context["title"])
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="body",
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="report", frames=[frame])])

    canvas_class = type("_ReportCanvas", (_NumberedCanvas,), {
        "_header_left": context["title"].upper(),
        "_header_right": context["job_name"].upper(),
        "_footer_left": context["source_file"],
    })
    doc.build(build_story(context, width), canvasmaker=canvas_class)

    if not pdf_path.is_file() or pdf_path.stat().st_size == 0:
        raise RuntimeError(f"No PDF was produced at {pdf_path}.")
    return pdf_path


# =============================================================================
# SECTION F   ORCHESTRATION AND CLI
# =============================================================================

_SAFE_NAME_RE = re.compile(r"[^A-Za-z0-9._-]+")


def _safe_name(text: str) -> str:
    return _SAFE_NAME_RE.sub("_", text).strip("_") or "report"


def _unique_stem(stem: str, used: set[str]) -> str:
    """Two uploads can share a filename; without this the second overwrites the
    first and shadows it inside the ZIP."""
    candidate, n = stem, 2
    while candidate.lower() in used:
        candidate, n = f"{stem}_{n}", n + 1
    used.add(candidate.lower())
    return candidate


def build_report(job: ParsedJob, out_dir: Path, stem: str) -> Path:
    graph = analyze(job)
    narrative = generate_narrative(graph)
    pdf_path = render_pdf(build_context(graph, narrative), out_dir / f"{stem}.pdf")
    log.info("Wrote %s", pdf_path.name)
    return pdf_path


def build_reports_for_file(xml_path: Path, out_dir: Path,
                           used_names: set[str] | None = None) -> list[Path]:
    """One PDF per job: a file holding several jobs produces several reports
    rather than silently documenting the first."""
    used_names = used_names if used_names is not None else set()
    jobs = DataStageParser(xml_path).parse_all_jobs()
    stem = _safe_name(xml_path.stem)
    pdfs = []
    for job in jobs:
        if not job.stages:
            log.warning("%s / %s has no stages; skipped.", xml_path.name, job.job_name)
            continue
        base = stem if len(jobs) == 1 else f"{stem}__{_safe_name(job.job_name)}"
        pdfs.append(build_report(job, out_dir, _unique_stem(base, used_names)))
    return pdfs


def generate(inputs: Sequence[str | Path], out_dir: str | Path) -> Path:
    """Main entry point. One report gives a .pdf, more give a .zip."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    pdfs: list[Path] = []
    failures: list[str] = []
    used: set[str] = set()

    for raw in inputs:
        path = Path(raw)
        try:
            pdfs.extend(build_reports_for_file(path, out_dir, used))
        except ParseError as exc:
            failures.append(f"{path.name}: {exc}")
            log.error("%s", exc)
        except Exception as exc:                    # one bad file must not sink the batch
            failures.append(f"{path.name}: {exc}")
            log.exception("Failed to process %s", path.name)

    if not pdfs:
        raise RuntimeError("No reports were generated.\n" + "\n".join(failures))
    if len(pdfs) == 1 and not failures:
        return pdfs[0]

    zip_path = out_dir / f"{_safe_name(Path(str(inputs[0])).stem)}_reports.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for pdf in pdfs:
            archive.write(pdf, arcname=pdf.name)
        if failures:
            archive.writestr("NOT_PROCESSED.txt",
                             "These inputs could not be processed:\n\n" + "\n".join(failures))
    log.info("Wrote %s (%s report(s))", zip_path.name, len(pdfs))
    return zip_path


class ReportStream(io.BytesIO):
    """Both a file object and a path, because the route may treat it as either.
    Flask's send_file accepts it both ways."""

    def __init__(self, path: Path):
        super().__init__(path.read_bytes())
        self.path = path
        self.name = path.name
        self.mimetype = "application/zip" if path.suffix == ".zip" else "application/pdf"

    def __fspath__(self) -> str:
        return str(self.path)

    def __str__(self) -> str:
        return str(self.path)

    def __repr__(self) -> str:
        return f"<ReportStream {self.path.name} {len(self.getbuffer()):,} bytes>"


def _resolve_export_path(file_name: str | Path, project: str | None) -> Path:
    """The route passes a project and a file name, not a path, and the convention
    joining them is not recorded anywhere, so each arrangement is tried."""
    candidate = Path(file_name)
    if candidate.is_file():
        return candidate

    tried = [str(candidate)]
    roots = [Path(r) for r in PROJECT_ROOTS]
    layouts: list[Path] = []
    for root in roots:
        if project:
            layouts.append(root / project / candidate)
        layouts.append(root / candidate)
    if project:
        layouts.append(Path(project) / candidate)

    for layout in layouts:
        tried.append(str(layout))
        if layout.is_file():
            log.info("Resolved %s to %s", file_name, layout)
            return layout

    for root in roots:
        if not root.is_dir():
            continue
        for depth in range(1, PROJECT_FILE_SEARCH_DEPTH + 1):
            for hit in root.glob("/".join(["*"] * depth) + "/" + candidate.name):
                if hit.is_file():
                    log.info("Found %s by searching %s: %s", candidate.name, root, hit)
                    return hit

    raise FileNotFoundError(
        f"Could not find the export '{file_name}'"
        + (f" for project '{project}'" if project else "")
        + ".\nLooked in:\n  " + "\n  ".join(tried)
        + "\nAdd the correct directory to PROJECT_ROOTS in generate_doc.py.")


def document_generate(*args: Any, **kwargs: Any) -> ReportStream:
    """Entry point for the Flask route.

        stream = document_generate(project_name_selection, file_name_selection)
    """
    project = selection = None
    out_dir = kwargs.get("out_dir") or kwargs.get("output_dir") or kwargs.get("dest")

    positional = list(args)
    if len(positional) >= 2:
        project, selection = positional[0], positional[1]
        if len(positional) >= 3 and out_dir is None:
            out_dir = positional[2]
        # A second string argument is ambiguous; an existing directory is an
        # output directory, anything else is the file selection.
        if isinstance(selection, (str, Path)) and Path(selection).is_dir():
            out_dir, selection, project = selection, project, None
    elif positional:
        selection = positional[0]

    for key in ("project_name_selection", "project_name", "project", "project_id"):
        if project is None and key in kwargs:
            project = kwargs[key]
    for key in ("file_name_selection", "file_name", "file_names", "file_path",
                "file_paths", "files", "inputs", "xml_path", "path"):
        if selection is None and key in kwargs:
            selection = kwargs[key]

    if selection is None:
        raise TypeError("document_generate() needs the export file(s) to analyse.")

    selection = [selection] if isinstance(selection, (str, Path)) else list(selection)
    if not selection:
        raise ValueError("document_generate() was given an empty file selection.")

    project = str(project) if project else None
    inputs = [_resolve_export_path(name, project) for name in selection]
    log.info("Generating from %s export(s)%s.", len(inputs),
             f" in project {project}" if project else "")
    return ReportStream(generate(inputs, out_dir or DEFAULT_OUTPUT_DIR))


def main(argv: Sequence[str] | None = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(
        description="Generate a workflow analysis report from DataStage XML exports.")
    parser.add_argument("inputs", nargs="+")
    parser.add_argument("-o", "--out-dir", default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--no-llm", action="store_true", help="Tables only, no descriptions.")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(levelname)s  %(message)s")

    global LLM_ENABLED
    if args.no_llm:
        LLM_ENABLED = False

    try:
        print(generate(args.inputs, args.out_dir))
    except (RuntimeError, FileNotFoundError) as exc:
        log.error("%s", exc)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
