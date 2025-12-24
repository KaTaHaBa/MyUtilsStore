"""
xbrl_parser.py

Single-file parser for EDINET XBRL ZIP submissions.

This module:
- Extracts the XBRL instance document from an EDINET ZIP (expects the instance under "PublicDoc").
- Parses the label linkbase to map concepts to human-readable labels when present.
- Extracts context metadata (periods and explicit dimensions/members) and unit metadata (measures and ISO 4217 currency).
- Emits a pandas DataFrame in a long-form "facts" layout, covering both numeric and non-numeric facts (including XHTML text blocks).
- Provides normalized fields (period_type/start/end, parsed numeric_value, decimals/precision/scale, dimensions, unit measures) for downstream analytics.

Canonical statement construction (e.g., PL/BS/CF mapping) is intentionally separated from parsing and is represented as stubs at the end of this file.
"""

from __future__ import annotations

import glob
import json
import os
import re
import tempfile
import zipfile
from pathlib import Path
from dataclasses import dataclass
from datetime import date
from html import unescape
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
from bs4 import BeautifulSoup


# -----------------------------
# Helpers / Types
# -----------------------------

@dataclass(frozen=True)
class Period:
    """Normalized representation of an XBRL context period."""
    period_type: str  # 'instant' | 'duration' | 'forever' | 'unknown'
    start_date: Optional[date]
    end_date: Optional[date]

    def as_string(self) -> str:
        """Return the human-readable period string used in the output."""
        if self.period_type == "instant" and self.end_date:
            return f"Instant: {self.end_date.isoformat()}"
        if self.period_type == "duration" and self.start_date and self.end_date:
            return f"{self.start_date.isoformat()} - {self.end_date.isoformat()}"
        if self.period_type == "forever":
            return "Forever"
        return "Unknown"


_NUMERIC_RE = re.compile(r"^\s*[-+]?(\d+(\.\d*)?|\.\d+)\s*$")
_TEXTBLOCK_RE = re.compile(r"textblock$", re.IGNORECASE)


def _safe_date(s: str) -> Optional[date]:
    try:
        y, m, d = s.split("-")
        return date(int(y), int(m), int(d))
    except Exception:
        return None


def _coerce_float(s: str) -> Optional[float]:
    """
    Coerce a numeric string to float.
    - This function is intentionally conservative: it does NOT strip commas.
      If commas appear, it returns None (commas are not expected in compliant XBRL facts).
    """
    if not isinstance(s, str):
        return None
    if not _NUMERIC_RE.match(s):
        return None
    try:
        return float(s)
    except Exception:
        return None


def _is_likely_text_block(tag: Any, element_name: Optional[str] = None, label: Optional[str] = None) -> bool:
    """
    Heuristic: text blocks often contain nested XHTML/HTML-like tags in the instance.
    We also check element/label suffixes to avoid missing empty-but-text block facts.
    """
    if element_name and _TEXTBLOCK_RE.search(element_name):
        return True
    if label and _TEXTBLOCK_RE.search(label):
        return True
    try:
        return any(getattr(child, "name", None) for child in tag.children)
    except Exception:
        return False


def _extract_text_block(tag: Any) -> str:
    """
    Extract and compact a text block:
    - Use inner XML content (decode_contents) to capture embedded XHTML.
    - Strip tags to plain text with line compaction and HTML entity cleanup.
    """
    try:
        inner = tag.decode_contents()  # type: ignore[attr-defined]
    except Exception:
        inner = tag.text or ""

    text_content = ""
    if inner:
        try:
            text_content = BeautifulSoup(inner, "lxml").get_text(separator="\n")
        except Exception:
            text_content = str(inner)

    text_content = unescape(text_content or "")
    if "<" in text_content and ">" in text_content:
        text_content = re.sub(r"<[^>]+>", " ", text_content)

    lines = [re.sub(r"\s+", " ", line).strip() for line in str(text_content).splitlines()]
    return "\n".join([line for line in lines if line]).strip()


# -----------------------------
# Parser
# -----------------------------

class XbrlParser:
    """
    Parses EDINET XBRL zip files to extract:
    - Numeric facts (with context/unit metadata)
    - Non-numeric facts / text blocks (compacted plain text)

    The output is a DataFrame suitable for downstream:
    - A "facts" table (long form) is naturally represented by this output.
    - "Canonical" financial statements should be built from this output (outside this parser).
    """

    def __init__(self, zip_file_path: str) -> None:
        self.zip_file_path = zip_file_path
        self.temp_dir: Optional[tempfile.TemporaryDirectory[str]] = None

    def parse(self) -> pd.DataFrame:
        """
        Main execution method.

        Steps:
        1. Extract the ZIP file.
        2. Find the main XBRL instance file under 'PublicDoc'.
        3. Parse the label linkbase file to map tags to labels.
        4. Parse contexts and units from the instance.
        5. Parse facts (numeric + text blocks) with metadata.
        """
        self.temp_dir = tempfile.TemporaryDirectory()
        extract_path = self.temp_dir.name

        try:
            with zipfile.ZipFile(self.zip_file_path, "r") as zf:
                zf.extractall(extract_path)

            xbrl_file = self._find_public_doc_xbrl(extract_path)
            if not xbrl_file:
                raise FileNotFoundError("No .xbrl file found under 'PublicDoc' directory.")

            target_dir = os.path.dirname(xbrl_file)
            label_file = self._find_file(target_dir, "*_lab.xml")

            label_map: Dict[str, str] = {}
            if label_file:
                label_map = self._parse_label_linkbase(label_file)

            data_list = self._parse_instance_file(xbrl_file, label_map)
            return pd.DataFrame(data_list)

        finally:
            if self.temp_dir:
                self.temp_dir.cleanup()

    # -----------------------------
    # File discovery
    # -----------------------------

    def _find_file(self, root_dir: str, pattern: str) -> Optional[str]:
        """Recursively search for the first file matching the pattern."""
        for root, _, files in os.walk(root_dir):
            for file in files:
                if glob.fnmatch.fnmatch(file, pattern):
                    return os.path.join(root, file)
        return None

    def _find_public_doc_xbrl(self, root_dir: str) -> Optional[str]:
        """Find an .xbrl file located under a 'PublicDoc' directory."""
        candidates: List[str] = []
        for root, _, files in os.walk(root_dir):
            for file in files:
                if file.endswith(".xbrl"):
                    candidates.append(os.path.join(root, file))

        # Prefer the instance directly under "PublicDoc"
        for path in candidates:
            if os.path.basename(os.path.dirname(path)) == "PublicDoc":
                return path

        # If no instance is found directly under "PublicDoc", use the first discovered .xbrl file.
        return candidates[0] if candidates else None

    # -----------------------------
    # Label linkbase
    # -----------------------------

    def _parse_label_linkbase(self, label_file_path: str) -> Dict[str, str]:
        """
        Parse label linkbase to map XBRL concept IDs to labels.

        Notes:
        - EDINET label linkbase typically maps loc -> concept, labelArc -> label resource.
        - We use the standard role 'http://www.xbrl.org/2003/role/label' by default.
        """
        with open(label_file_path, "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f, "lxml-xml")

        label_map: Dict[str, str] = {}

        locs = soup.find_all("loc")
        arcs = soup.find_all("labelArc")
        labels = soup.find_all("label")

        loc_label_to_tag: Dict[str, str] = {}
        for loc in locs:
            href = loc.get("xlink:href")
            if href and "#" in href:
                tag_name = href.split("#")[1]
                loc_label = loc.get("xlink:label")
                if loc_label:
                    loc_label_to_tag[loc_label] = tag_name

        res_label_to_text: Dict[str, str] = {}
        for lab in labels:
            if lab.get("xlink:role") == "http://www.xbrl.org/2003/role/label":
                res_label = lab.get("xlink:label")
                if res_label:
                    res_label_to_text[res_label] = lab.text.strip()

        for arc in arcs:
            from_loc = arc.get("xlink:from")
            to_res = arc.get("xlink:to")
            if not from_loc or not to_res:
                continue
            tag_name = loc_label_to_tag.get(from_loc)
            text = res_label_to_text.get(to_res)
            if tag_name and text:
                label_map[tag_name] = text

        return label_map

    # -----------------------------
    # Instance parsing
    # -----------------------------

    def _parse_instance_file(self, xbrl_file_path: str, label_map: Dict[str, str]) -> List[Dict[str, Any]]:
        """
        Parse the XBRL instance file and return rows for DataFrame construction.

        The output includes core columns and additional metadata:
          - period_type, period_start, period_end
          - unit_measures (json string), currency (if iso4217)
          - decimals, precision, scale
          - numeric_value (float when parseable)
          - dimensions (json string): explicit members in scenario/segment
          - is_text_block (bool)
        """
        with open(xbrl_file_path, "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f, "lxml-xml")

        context_details = self._extract_context_details(soup)
        unit_details = self._extract_units(soup)

        # Collect candidate facts under the instance root (avoid linkbase nodes in theory)
        root = soup.find(recursive=False)
        all_tags = root.find_all(recursive=False) if root else soup.find_all()

        # Phase 1: Scan & collect (standard may be determined later)
        detected_standard: Optional[str] = None
        temp_rows: List[Dict[str, Any]] = []

        for tag in all_tags:
            context_ref = tag.get("contextRef")
            if not context_ref:
                continue

            # Namespace logic
            prefix = getattr(tag, "prefix", None)
            tag_name = tag.name
            if not prefix and isinstance(tag_name, str) and ":" in tag_name:
                prefix, element = tag_name.split(":", 1)
            else:
                element = str(tag_name)
                if not prefix:
                    prefix = "unknown"

            full_tag_name = f"{prefix}:{element}" if prefix != "unknown" else element

            # Label lookup
            label = label_map.get(full_tag_name) or label_map.get(element) or element

            # Extract value
            is_text_block = _is_likely_text_block(tag, element_name=element, label=label)
            if is_text_block:
                clean_value = _extract_text_block(tag)
            else:
                clean_value = (tag.text or "").strip()

            if not clean_value:
                continue

            # Detect accounting standard from DEI (keyword-based detection).
            if element == "AccountingStandardsDEI":
                if ("IFRS" in clean_value) or ("International" in clean_value):
                    detected_standard = "IFRS"
                elif ("US" in clean_value) or ("United States" in clean_value) or ("米国" in clean_value):
                    detected_standard = "USGAAP"
                elif ("Japan" in clean_value) or ("日本" in clean_value):
                    detected_standard = "JGAAP"

            # Context details
            ctx = context_details.get(context_ref, {})
            period: Period = ctx.get("period", Period("unknown", None, None))
            period_string = period.as_string()

            # Consolidated status (prefer dimension members; fall back to contextRef string).
            consolidated = self._detect_consolidated(ctx, context_ref)

            # Unit details
            unit_id = tag.get("unitRef")
            unit = unit_details.get(unit_id, {}) if unit_id else {}
            unit_measures = unit.get("measures", [])
            currency = unit.get("currency")

            # Numeric attributes
            decimals = tag.get("decimals")
            precision = tag.get("precision")
            scale = tag.get("scale")  # typically in inline XBRL, but harmless here

            numeric_value = None
            if unit_id:
                numeric_value = _coerce_float(clean_value)

            # Dimensions (scenario/segment explicit members)
            dimensions = ctx.get("dimensions", {})

            temp_rows.append({
                # Core columns
                "Tag": full_tag_name,
                "Element": element,
                "Prefix": prefix,
                "Label": label,
                "Value": clean_value,
                "ContextID": context_ref,
                "Period/Setting": period_string,
                "UnitID": unit_id,
                "Consolidated": consolidated,

                # Additional metadata columns
                "Standard": None,  # filled later
                "period_type": period.period_type,
                "period_start": period.start_date.isoformat() if period.start_date else None,
                "period_end": period.end_date.isoformat() if period.end_date else None,

                "decimals": decimals,
                "precision": precision,
                "scale": scale,

                "numeric_value": numeric_value,
                "unit_measures": json.dumps(unit_measures, ensure_ascii=False),
                "currency": currency,

                "dimensions": json.dumps(dimensions, ensure_ascii=False),
                "is_text_block": bool(is_text_block),
            })

        # Phase 2: Finalize Standard
        if not detected_standard:
            prefixes = [r.get("Prefix", "") or "" for r in temp_rows]
            if any("ifrs" in p.lower() or "jpigp" in p.lower() for p in prefixes):
                detected_standard = "IFRS"
            elif any("us-gaap" in p.lower() for p in prefixes):
                detected_standard = "USGAAP"
            else:
                detected_standard = "JGAAP"

        for row in temp_rows:
            row["Standard"] = detected_standard

        return temp_rows

    # -----------------------------
    # Context / Unit extraction
    # -----------------------------

    def _extract_context_details(self, soup: BeautifulSoup) -> Dict[str, Dict[str, Any]]:
        """
        Extract detailed context information:
          - period (instant/duration/forever)
          - explicit dimensions and members under scenario/segment

        Output shape:
          {
            "C_2024_Consolidated": {
                "period": Period(...),
                "dimensions": { "<dimensionQName>": "<memberQName>", ... }
            },
            ...
          }
        """
        out: Dict[str, Dict[str, Any]] = {}
        for ctx in soup.find_all("context"):
            ctx_id = ctx.get("id")
            if not ctx_id:
                continue

            # Period parsing
            period_node = ctx.find("period")
            period = Period("unknown", None, None)
            if period_node:
                instant = period_node.find("instant")
                if instant and instant.text:
                    period = Period("instant", None, _safe_date(instant.text.strip()))
                else:
                    start = period_node.find("startDate")
                    end = period_node.find("endDate")
                    if start and end and start.text and end.text:
                        period = Period("duration", _safe_date(start.text.strip()), _safe_date(end.text.strip()))
                    else:
                        forever = period_node.find("forever")
                        if forever is not None:
                            period = Period("forever", None, None)

            # Dimension/member extraction
            dimensions: Dict[str, str] = {}
            for container_name in ("scenario", "segment"):
                cont = ctx.find(container_name)
                if not cont:
                    continue
                for em in cont.find_all(re.compile(r".*explicitMember$")):
                    dim = em.get("dimension")
                    mem = (em.text or "").strip()
                    if dim and mem:
                        dimensions[str(dim)] = mem

            out[ctx_id] = {"period": period, "dimensions": dimensions}

        return out

    def _extract_units(self, soup: BeautifulSoup) -> Dict[str, Dict[str, Any]]:
        """
        Extract unit definitions:
          - measures: list[str] of <measure> qnames
          - currency: ISO 4217 code when detectable (iso4217:JPY -> JPY)

        Output shape:
          {
            "JPY": {"measures": ["iso4217:JPY"], "currency": "JPY"},
            "Pure": {"measures": ["xbrli:pure"], "currency": None},
            ...
          }
        """
        out: Dict[str, Dict[str, Any]] = {}
        for unit in soup.find_all("unit"):
            unit_id = unit.get("id")
            if not unit_id:
                continue

            measures = []
            for m in unit.find_all("measure"):
                if m.text:
                    measures.append(m.text.strip())

            currency = None
            for meas in measures:
                # Typical representation: "iso4217:JPY"
                if isinstance(meas, str) and meas.lower().startswith("iso4217:"):
                    currency = meas.split(":", 1)[1].upper()
                    break

            out[unit_id] = {"measures": measures, "currency": currency}

        return out

    # -----------------------------
    # Consolidation logic
    # -----------------------------

    def _detect_consolidated(self, ctx: Dict[str, Any], context_ref: str) -> bool:
        """
        Determine consolidated/non-consolidated status.

        Priority:
        1) Use explicit members in dimensions.
        2) If dimension members are unavailable, use a contextRef string heuristic.

        Notes:
        - EDINET commonly uses "*NonConsolidatedMember" for non-consolidated contexts.
        - This function is intentionally conservative: if unsure, returns True.
        """
        dims: Dict[str, str] = ctx.get("dimensions", {}) if isinstance(ctx, dict) else {}
        members = list(dims.values())

        # Dimension-based checks
        for mem in members:
            if "NonConsolidatedMember" in mem:
                return False
            if "ConsolidatedMember" in mem:
                return True

        # String-based heuristic when dimension members are not available
        if "NonConsolidatedMember" in (context_ref or ""):
            return False

        return True

# -----------------------------
# Downstream analytics (Facts -> Canonical statements)
# -----------------------------

from enum import Enum


class StatementType(str, Enum):
    """Canonical statement types for downstream analytics."""
    PL = "PL"
    BS = "BS"
    CF = "CF"
    OTHER = "OTHER"


@dataclass(frozen=True)
class CanonicalPeriodFilter:
    """Filter configuration to select annual (FY) facts.

    This is intentionally heuristic because context patterns differ across taxonomies and filers.
    """
    prefer_duration: bool = True
    min_duration_days: int = 300  # Treat >= ~10 months as annual
    choose_latest_end_date: bool = True


@dataclass(frozen=True)
class MappingCandidate:
    """A candidate matcher to map Facts to canonical keys."""
    field: str  # "element" | "tag" | "label"
    exact: Optional[str] = None
    regex: Optional[str] = None
    label_regex: Optional[str] = None
    weight: float = 1.0


@dataclass(frozen=True)
class CanonicalRule:
    """Mapping rule for a canonical line item."""
    canonical_key: str
    statement: StatementType
    period_type: Optional[str] = None  # "duration" | "instant" | None
    candidates: Tuple[MappingCandidate, ...] = tuple()


class MappingLoader:
    """Load and merge mapping rules from YAML/JSON/dicts.

    Company-specific overrides should be stored outside core logic and merged last.
    """

    @staticmethod
    def load(path: Path) -> Dict[str, Any]:
        """Load mapping from YAML or JSON file."""
        ...

    @staticmethod
    def merge(base: Dict[str, Any], overlay: Optional[Dict[str, Any]]) -> Dict[str, Any]:
        """Merge mappings with shallow merge on 'items'. Overlay wins."""
        ...

    @staticmethod
    def rules_from_dict(d: Dict[str, Any]) -> List[CanonicalRule]:
        """Convert mapping dict into CanonicalRule list."""
        ...


class FactsAnalytics:
    """Build canonical financial statement tables from normalized facts.

    Expected input:
      - DataFrame produced by XbrlParser.parse() or a CSV exported from that DataFrame.
    """

    @staticmethod
    def attach_identifiers(*args, **kwargs): ...

    @staticmethod
    def filter_annual_facts(*args, **kwargs): ...

    @staticmethod
    def resolve_canonical_long(*args, **kwargs): ...

    @staticmethod
    def canonical_wide(*args, **kwargs): ...
