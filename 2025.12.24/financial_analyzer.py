from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

from columuns_definition_config import ColumnDefinitionConfig


class StatementType(str, Enum):
    """Canonical statement types for downstream analytics."""

    PL = "PL"
    BS = "BS"
    CF = "CF"
    PORTFOLIO = "PORTFOLIO"
    OTHER = "OTHER"


@dataclass(frozen=True)
class CanonicalPeriodFilter:
    """Filter configuration to select annual (FY) facts."""

    prefer_duration: bool = True
    min_duration_days: int = 300


@dataclass(frozen=True)
class MappingCandidate:
    """A candidate matcher to map facts to a canonical key."""

    field: str
    exact: Optional[str] = None
    regex: Optional[str] = None
    weight: float = 1.0


@dataclass(frozen=True)
class CanonicalRule:
    """Mapping rule for a canonical line item."""

    canonical_key: str
    statement: StatementType
    period_type: Optional[str]
    candidates: Tuple[MappingCandidate, ...] = tuple()


@dataclass(frozen=True)
class PortfolioRule:
    """Mapping rule for portfolio position extraction."""

    portfolio_key: str
    period_type: str
    candidates: Tuple[MappingCandidate, ...] = tuple()
    aggregate_mode: Optional[str] = None
    exclude_regex: Tuple[str, ...] = tuple()


class FinancialAnalyzer:
    """Analyze parsed XBRL facts and build canonical financial statements."""

    _ALLOWED_DIMENSION_SUFFIXES = (
        "ConsolidatedOrNonConsolidatedAxis",
        "ConsolidatedOrSeparateAxis",
        "ConsolidatedOrSeparateFinancialStatementsAxis",
    )

    def __init__(
        self,
        facts: pd.DataFrame,
        mapping: Optional[Dict[str, Any]] = None,
        prefer_consolidated: bool = True,
        standard: Optional[str] = None,
        company_name: Optional[str] = None,
    ) -> None:
        self.facts = self._normalize_facts(facts)
        self.prefer_consolidated = prefer_consolidated
        self.standard = standard or self._infer_standard(self.facts)
        self.company_name = company_name

        mapping_dict = mapping or ColumnDefinitionConfig.resolve_mapping(
            standard=self.standard,
            company_name=self.company_name,
        )
        self.rules = self._rules_from_mapping(mapping_dict)

    @classmethod
    def from_csvs(
        cls,
        csv_paths: Iterable[Path],
        mapping: Optional[Dict[str, Any]] = None,
        prefer_consolidated: bool = True,
        standard: Optional[str] = None,
        company_name: Optional[str] = None,
    ) -> "FinancialAnalyzer":
        """Load multiple parsed CSVs and construct an analyzer."""
        dfs = [pd.read_csv(path) for path in csv_paths]
        merged = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
        return cls(
            merged,
            mapping=mapping,
            prefer_consolidated=prefer_consolidated,
            standard=standard,
            company_name=company_name,
        )

    def get_pl_data(self) -> pd.DataFrame:
        """Return a wide PL table for trend analysis."""
        return self.canonical_wide(StatementType.PL)

    def get_bs_data(self) -> pd.DataFrame:
        """Return a wide BS table for structure analysis."""
        df = self.canonical_wide(StatementType.BS)
        return self._reconcile_balance_sheet(df)

    def get_cf_data(self) -> pd.DataFrame:
        """Return a wide CF table for cash flow analysis."""
        return self.canonical_wide(StatementType.CF)

    def get_portfolio_positions(
        self,
        period_filter: Optional[CanonicalPeriodFilter] = None,
        consolidated: Optional[bool] = None,
    ) -> pd.DataFrame:
        """Extract portfolio-related positions by category."""
        period_filter = period_filter or CanonicalPeriodFilter(prefer_duration=False, min_duration_days=0)
        rows: List[Dict[str, Any]] = []
        for rule in self._portfolio_rules():
            matched = self._match_portfolio_rule(rule, period_filter, consolidated)
            if matched.empty:
                continue
            for _, row in matched.iterrows():
                rows.append(
                    {
                        "portfolio_key": rule.portfolio_key,
                        "period_end": row.get("period_end"),
                        "period_start": row.get("period_start"),
                        "period_type": row.get("period_type"),
                        "value": row.get("value"),
                        "currency": row.get("currency"),
                        "consolidated": row.get("Consolidated"),
                        "standard": row.get("Standard"),
                        "tag": row.get("Tag"),
                        "element": row.get("Element"),
                        "label": row.get("Label"),
                        "context_id": row.get("ContextID"),
                        "match_score": row.get("match_score"),
                    }
                )
        return pd.DataFrame(rows)

    def get_portfolio_timeseries(
        self,
        series_defs: Optional[Sequence[Dict[str, Any]]] = None,
        period_filter: Optional[CanonicalPeriodFilter] = None,
        consolidated: Optional[bool] = None,
    ) -> pd.DataFrame:
        """Return a wide portfolio table with period labels as rows."""
        long_df = self.get_portfolio_positions(period_filter=period_filter, consolidated=consolidated)
        if long_df.empty:
            columns = ["period_label"]
            if series_defs:
                columns += [spec.get("column_name", spec.get("portfolio_key", "")) for spec in series_defs]
            return pd.DataFrame(columns=columns)

        long_df = long_df.copy()
        long_df["period_label"] = long_df["period_end"].astype(str)
        wide = (
            long_df.pivot_table(
                index="period_label",
                columns="portfolio_key",
                values="value",
                aggfunc="first",
            )
            .reset_index()
        )

        if "TotalSecurities" in wide.columns:
            total_series = pd.to_numeric(wide["TotalSecurities"], errors="coerce")
        else:
            total_series = None

        if "EquitySecurities" in wide.columns:
            equity_series = pd.to_numeric(wide["EquitySecurities"], errors="coerce")
        else:
            equity_series = None

        if total_series is not None:
            if equity_series is not None:
                equity_adjusted = equity_series.where(total_series.isna() | (equity_series <= total_series), total_series)
                wide["EquitySecurities"] = equity_adjusted
                derived_debt = total_series - equity_adjusted
                derived_debt = derived_debt.where(derived_debt > 0, 0)
                if "DebtSecurities" not in wide.columns or wide["DebtSecurities"].isna().all():
                    wide["DebtSecurities"] = derived_debt
            if "DebtSecurities" not in wide.columns or wide["DebtSecurities"].isna().all():
                wide["DebtSecurities"] = total_series

        if "DerivativeAssets" in wide.columns and "DerivativeLiabilities" in wide.columns:
            wide["DerivativeNet"] = pd.to_numeric(wide["DerivativeAssets"], errors="coerce") - pd.to_numeric(
                wide["DerivativeLiabilities"], errors="coerce"
            )

        if not series_defs:
            return wide

        rename_map: Dict[str, str] = {}
        ordered_cols = ["period_label"]
        for spec in series_defs:
            key = spec.get("portfolio_key")
            col_name = spec.get("column_name", key)
            if key and col_name:
                rename_map[key] = col_name
                ordered_cols.append(col_name)

        wide = wide.rename(columns=rename_map)
        for col in ordered_cols:
            if col not in wide.columns:
                wide[col] = None

        return wide[ordered_cols]

    def build_slide_payload(self) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
        """Build slide-ready data frames and chart mappings."""
        layout = ColumnDefinitionConfig.resolve_layout(
            standard=self.standard,
            company_name=self.company_name,
        )
        pl_series = layout.get("pl_series", [])
        cf_series = layout.get("cf_series", [])
        bs_series = layout.get("bs_series", [])

        df_pl = self.get_pl_data()
        df_bs = self.get_bs_data()
        df_cf = self.get_cf_data()

        pl_trend = self._build_trend_dataframe(df_pl, pl_series)
        cf_trend = self._build_trend_dataframe(df_cf, cf_series)
        bs_snapshot = self._build_snapshot_dataframe(df_bs, bs_series)

        data_store = {
            "pl_trend": pl_trend,
            "cf_trend": cf_trend,
            "bs_snapshot": bs_snapshot,
        }

        pl_chart = layout.get("pl_chart", {})
        cf_chart = layout.get("cf_chart", {})
        bs_chart = layout.get("bs_chart", {})

        pl_bar_traces = self._build_traces(pl_series, pl_chart.get("bar_keys", []))
        pl_line_traces = self._build_traces(pl_series, pl_chart.get("line_keys", []))
        cf_bar_traces = self._build_traces(cf_series, cf_chart.get("bar_keys", []))
        cf_line_traces = self._build_traces(cf_series, cf_chart.get("line_keys", []))

        slides_structure = [
            {
                "slide_title": pl_chart.get("slide_title", "PL"),
                "category": pl_chart.get("category", "combo_bar_line_2axis"),
                "data_source": "pl_trend",
                "data_mapping": {
                    "x_col": pl_chart.get("x_col", "period_label"),
                    "x_label_format": pl_chart.get("x_label_format", "fy"),
                    "unit_scale": pl_chart.get("unit_scale", 1.0),
                    "bar_traces": pl_bar_traces,
                    "line_traces": pl_line_traces,
                },
                "chart_text": pl_chart.get("chart_text", {}),
                "layout_type": "horizontal",
            },
            {
                "slide_title": bs_chart.get("slide_title", "BS"),
                "category": bs_chart.get("category", "balance_sheet"),
                "data_source": "bs_snapshot",
                "data_mapping": bs_chart,
                "chart_text": bs_chart.get("chart_text", {"title": "Balance Sheet Structure", "y1_label": "JPY (bn)"}),
                "layout_type": "horizontal",
            },
            {
                "slide_title": cf_chart.get("slide_title", "CF"),
                "category": cf_chart.get("category", "combo_bar_line_2axis"),
                "data_source": "cf_trend",
                "data_mapping": {
                    "x_col": cf_chart.get("x_col", "period_label"),
                    "x_label_format": cf_chart.get("x_label_format", "fy"),
                    "unit_scale": cf_chart.get("unit_scale", 1.0),
                    "bar_traces": cf_bar_traces,
                    "line_traces": cf_line_traces,
                },
                "chart_text": cf_chart.get("chart_text", {}),
                "layout_type": "horizontal",
            },
        ]

        portfolio_series = layout.get("portfolio_series", [])
        portfolio_chart = layout.get("portfolio_chart", {})
        if portfolio_series and portfolio_chart:
            portfolio_trend = self.get_portfolio_timeseries(series_defs=portfolio_series)
            data_store["portfolio_trend"] = portfolio_trend

            portfolio_traces = self._build_portfolio_traces(
                portfolio_series,
                portfolio_chart.get("series_keys", []),
            )
            slides_structure.append(
                {
                    "slide_title": portfolio_chart.get("slide_title", "Portfolio"),
                    "category": portfolio_chart.get("category", "portfolio_timeseries"),
                    "data_source": "portfolio_trend",
                    "data_mapping": {
                        "x_col": portfolio_chart.get("x_col", "period_label"),
                        "x_label_format": portfolio_chart.get("x_label_format", "fy"),
                        "unit_scale": portfolio_chart.get("unit_scale", 1.0),
                        "series": portfolio_traces,
                        "stacked": portfolio_chart.get("stacked", True),
                        "chart_style": portfolio_chart.get("chart_style", "area"),
                        "bar_width": portfolio_chart.get("bar_width", 0.6),
                        "x_label_rotation": portfolio_chart.get("x_label_rotation", 0),
                        "x_tick_step": portfolio_chart.get("x_tick_step"),
                    },
                    "chart_text": portfolio_chart.get("chart_text", {}),
                    "layout_type": portfolio_chart.get("layout_type", "horizontal"),
                }
            )

        return data_store, slides_structure

    def canonical_wide(self, statement: StatementType) -> pd.DataFrame:
        """Pivot canonical items into a wide table (period_end columns)."""
        long_df = self.resolve_canonical_long(statement)
        if long_df.empty:
            return long_df
        long_df = long_df.copy()
        long_df["period_label"] = long_df["period_start"].astype(str) + "-" + long_df["period_end"].astype(str)

        label_choice = (
            long_df.groupby(["canonical_key", "label"])
            .size()
            .reset_index(name="count")
            .sort_values(["canonical_key", "count"], ascending=[True, False])
            .drop_duplicates(subset=["canonical_key"])
            .set_index("canonical_key")["label"]
        )

        wide = (
            long_df.pivot_table(
                index=["canonical_key", "statement"],
                columns="period_label",
                values="value",
                aggfunc="first",
            )
            .reset_index()
        )
        wide["label"] = wide["canonical_key"].map(label_choice)
        wide = wide[["canonical_key", "label", "statement"] + [c for c in wide.columns if c not in {"canonical_key", "label", "statement"}]]
        return wide

    def resolve_canonical_long(
        self,
        statement: StatementType,
        period_filter: Optional[CanonicalPeriodFilter] = None,
        consolidated: Optional[bool] = None,
    ) -> pd.DataFrame:
        """Resolve canonical facts into a long-form table."""
        period_filter = period_filter or CanonicalPeriodFilter()
        target_rules = [r for r in self.rules if r.statement == statement]

        rows: List[Dict[str, Any]] = []
        for rule in target_rules:
            matched = self._match_rule(rule, period_filter, consolidated)
            if matched.empty:
                continue

            for _, row in matched.iterrows():
                rows.append(
                    {
                        "canonical_key": rule.canonical_key,
                        "statement": rule.statement.value,
                        "period_end": row.get("period_end"),
                        "period_start": row.get("period_start"),
                        "period_type": row.get("period_type"),
                        "value": row.get("value"),
                        "currency": row.get("currency"),
                        "consolidated": row.get("Consolidated"),
                        "standard": row.get("Standard"),
                        "tag": row.get("Tag"),
                        "element": row.get("Element"),
                        "label": row.get("Label"),
                        "context_id": row.get("ContextID"),
                        "match_score": row.get("match_score"),
                    }
                )

        return pd.DataFrame(rows)

    def _match_rule(
        self,
        rule: CanonicalRule,
        period_filter: CanonicalPeriodFilter,
        consolidated: Optional[bool],
    ) -> pd.DataFrame:
        """Return best matched facts for a single canonical rule."""
        df = self._filter_facts(rule.period_type, period_filter, include_dimensioned=False)
        if df.empty:
            return df

        if consolidated is not None:
            df = df[df["Consolidated"] == consolidated]
        elif self.prefer_consolidated and "Consolidated" in df.columns:
            # Keep both, but prefer consolidated in ranking.
            pass

        if df.empty:
            return df

        scores = []
        for _, row in df.iterrows():
            if rule.canonical_key == "TotalEquity":
                element_value = str(row.get("Element", "") or "")
                label_value = str(row.get("Label", "") or "")
                if ("LiabilitiesAndNetAssets" in element_value) or ("LiabilitiesAndNetAssets" in label_value):
                    scores.append(0.0)
                    continue
            score = 0.0
            for cand in rule.candidates:
                if self._candidate_match(row, cand):
                    score += cand.weight
            scores.append(score)

        df = df.copy()
        df["match_score"] = scores
        df = df[df["match_score"] > 0]
        if df.empty:
            return df

        df["value"] = df["numeric_value"].where(df["numeric_value"].notna(), df["Value"])

        df["is_preferred_consolidated"] = False
        if "Consolidated" in df.columns:
            df["is_preferred_consolidated"] = df["Consolidated"] == True

        df = df.sort_values(
            by=["period_end_dt", "is_preferred_consolidated", "match_score", "has_currency", "abs_numeric"],
            ascending=[False, False, False, False, False],
        )

        best_rows = []
        for _, group in df.groupby("period_end"):
            best_rows.append(group.iloc[0])

        return pd.DataFrame(best_rows)

    def _match_portfolio_rule(
        self,
        rule: PortfolioRule,
        period_filter: CanonicalPeriodFilter,
        consolidated: Optional[bool],
    ) -> pd.DataFrame:
        """Return best matched facts for a single portfolio rule."""
        df = self._filter_facts(rule.period_type, period_filter, include_dimensioned=True)
        if df.empty:
            return df

        if rule.exclude_regex:
            combined = (
                df.get("Element", "").astype(str).fillna("")
                + " "
                + df.get("Label", "").astype(str).fillna("")
                + " "
                + df.get("Tag", "").astype(str).fillna("")
            )
            pattern = "|".join(rule.exclude_regex)
            df = df[~combined.str.contains(pattern, case=False, regex=True, na=False)]
            if df.empty:
                return df

        if consolidated is not None:
            df = df[df["Consolidated"] == consolidated]
        elif self.prefer_consolidated and "Consolidated" in df.columns:
            pass

        if df.empty:
            return df

        scores = []
        for _, row in df.iterrows():
            score = 0.0
            for cand in rule.candidates:
                if self._candidate_match(row, cand):
                    score += cand.weight
            scores.append(score)

        df = df.copy()
        df["match_score"] = scores
        df = df[df["match_score"] > 0]
        if df.empty:
            return df

        df["value"] = df["numeric_value"].where(df["numeric_value"].notna(), df["Value"])
        if rule.aggregate_mode == "sum":
            df["value"] = pd.to_numeric(df["value"], errors="coerce")
            grouped_rows = []
            for _, group in df.groupby("period_end"):
                if "Consolidated" in group.columns and (group["Consolidated"] == True).any():
                    group = group[group["Consolidated"] == True]
                total = group["value"].sum(skipna=True)
                if pd.isna(total) or total == 0:
                    continue
                group = group.copy()
                group["value"] = total
                group["numeric_value"] = total
                group["is_preferred_consolidated"] = group["Consolidated"] == True if "Consolidated" in group.columns else False
                group = group.sort_values(
                    by=["period_end_dt", "is_preferred_consolidated", "match_score", "has_currency", "abs_numeric"],
                    ascending=[False, False, False, False, False],
                )
                grouped_rows.append(group.iloc[0])
            return pd.DataFrame(grouped_rows)

        df["is_preferred_consolidated"] = False
        if "Consolidated" in df.columns:
            df["is_preferred_consolidated"] = df["Consolidated"] == True

        df = df.sort_values(
            by=["period_end_dt", "is_preferred_consolidated", "match_score", "has_currency", "abs_numeric"],
            ascending=[False, False, False, False, False],
        )

        best_rows = []
        for _, group in df.groupby("period_end"):
            best_rows.append(group.iloc[0])

        return pd.DataFrame(best_rows)

    def _filter_facts(
        self,
        period_type: Optional[str],
        period_filter: CanonicalPeriodFilter,
        include_dimensioned: bool = False,
    ) -> pd.DataFrame:
        """Apply period type and annual duration filters."""
        df = self.facts

        if not include_dimensioned and "dimension_allowed" in df.columns:
            df = df[df["dimension_allowed"]]

        if period_type:
            df = df[df["period_type"] == period_type]

        if period_filter.prefer_duration:
            duration_mask = df["period_type"] == "duration"
            if duration_mask.any():
                df = df[duration_mask]

        if period_filter.min_duration_days and "duration_days" in df.columns:
            # Apply duration filter only to duration facts; instant facts stay untouched.
            duration_mask = df["period_type"] == "duration"
            df = df[~duration_mask | df["duration_days"].isna() | (df["duration_days"] >= period_filter.min_duration_days)]

        return df

    def _candidate_match(self, row: pd.Series, cand: MappingCandidate) -> bool:
        """Check whether a candidate matches a fact row."""
        field_map = {
            "element": "Element",
            "tag": "Tag",
            "label": "Label",
        }
        field_name = field_map.get(cand.field)
        if not field_name:
            return False

        value = str(row.get(field_name, "") or "")
        if cand.exact and value.lower() == cand.exact.lower():
            return True
        if cand.regex:
            return re.search(cand.regex, value, flags=re.IGNORECASE) is not None
        return False

    def _portfolio_rules(self) -> List[PortfolioRule]:
        """Return portfolio extraction rules with regex-based candidates."""
        return [
            PortfolioRule(
                portfolio_key="EquitySecurities",
                period_type="instant",
                aggregate_mode="sum",
                candidates=(
                    MappingCandidate(field="element", regex=r"BookValueDetailsOf.*EquitySecurities"),
                    MappingCandidate(field="label", regex=r"Book Value.*Equity Securities"),
                ),
                exclude_regex=(
                    r"NumberOfShares",
                    r"NameOfSecurities",
                    r"PurposesOfHolding",
                    r"WhetherIssuer",
                    r"ReasonForIncrease",
                    r"NumberOfNames",
                ),
            ),
            PortfolioRule(
                portfolio_key="DebtSecurities",
                period_type="instant",
                candidates=(
                    MappingCandidate(field="element", regex=r"DebtSecurities|BondSecurities|BondsSecurities"),
                    MappingCandidate(field="label", regex=r"Debt Securities|Bond Investments"),
                ),
            ),
            PortfolioRule(
                portfolio_key="TotalSecurities",
                period_type="instant",
                candidates=(
                    MappingCandidate(field="element", exact="InvestmentSecurities", weight=1.2),
                    MappingCandidate(field="element", exact="SecuritiesAssetsBNK", weight=1.2),
                    MappingCandidate(field="element", regex=r"^AvailableForSaleSecurities$"),
                    MappingCandidate(field="element", regex=r"^HeldToMaturitySecurities$"),
                    MappingCandidate(field="element", regex=r"^TradingSecurities$"),
                    MappingCandidate(field="label", regex=r"^Investment Securities$|^Securities$"),
                ),
                exclude_regex=(
                    r"ValuationDifference",
                    r"GainLoss",
                    r"Unrealized",
                    r"ChangesInFairValue",
                ),
            ),
            PortfolioRule(
                portfolio_key="DerivativeAssets",
                period_type="instant",
                candidates=(
                    MappingCandidate(field="element", regex=r"DerivativeAssets|DerivativesAssets"),
                    MappingCandidate(field="element", regex=r"DerivativeFinancialAssets|DerivativesFinancialAssets"),
                    MappingCandidate(field="label", regex=r"Derivative Assets"),
                ),
            ),
            PortfolioRule(
                portfolio_key="DerivativeLiabilities",
                period_type="instant",
                candidates=(
                    MappingCandidate(field="element", regex=r"DerivativeLiabilities|DerivativesLiabilities"),
                    MappingCandidate(field="element", regex=r"DerivativeFinancialLiabilities|DerivativesFinancialLiabilities"),
                    MappingCandidate(field="label", regex=r"Derivative Liabilities"),
                ),
            ),
        ]

    def _rules_from_mapping(self, mapping: Dict[str, Any]) -> List[CanonicalRule]:
        """Convert a mapping dict into CanonicalRule list."""
        rules = []
        for item in mapping.get("items", []):
            candidates = tuple(
                MappingCandidate(
                    field=c.get("field", "element"),
                    exact=c.get("exact"),
                    regex=c.get("regex"),
                    weight=float(c.get("weight", 1.0)),
                )
                for c in item.get("candidates", [])
            )
            rules.append(
                CanonicalRule(
                    canonical_key=item["canonical_key"],
                    statement=StatementType(item.get("statement", "OTHER")),
                    period_type=item.get("period_type"),
                    candidates=candidates,
                )
            )
        return rules

    def _normalize_facts(self, facts: pd.DataFrame) -> pd.DataFrame:
        """Normalize parsed facts for analytics."""
        df = facts.copy()

        required_cols = [
            "Tag",
            "Element",
            "Label",
            "Value",
            "ContextID",
            "Consolidated",
            "Standard",
            "period_type",
            "period_start",
            "period_end",
            "numeric_value",
            "currency",
            "is_text_block",
        ]
        for col in required_cols:
            if col not in df.columns:
                df[col] = None

        # Normalize consolidated flags to proper booleans for ranking.
        if "Consolidated" in df.columns:
            normalized = df["Consolidated"].astype(str).str.strip().str.lower()
            mapped = normalized.map(
                {
                    "true": True,
                    "false": False,
                    "1": True,
                    "0": False,
                    "yes": True,
                    "no": False,
                }
            )
            df["Consolidated"] = mapped.where(mapped.notna(), df["Consolidated"])

        df["numeric_value"] = pd.to_numeric(df["numeric_value"], errors="coerce")
        value_num = pd.to_numeric(df["Value"], errors="coerce")
        df["numeric_value"] = df["numeric_value"].fillna(value_num)

        if "Period/Setting" in df.columns:
            period_setting = df["Period/Setting"].astype(str)
            duration = period_setting.str.extract(r"(?P<start>\\d{4}-\\d{2}-\\d{2})\\s*-\\s*(?P<end>\\d{4}-\\d{2}-\\d{2})")
            instant = period_setting.str.extract(r"Instant:\\s*(?P<instant>\\d{4}-\\d{2}-\\d{2})")
            df["period_start"] = df["period_start"].fillna(duration["start"])
            df["period_end"] = df["period_end"].fillna(duration["end"])
            df["period_end"] = df["period_end"].fillna(instant["instant"])

        df["period_start_dt"] = pd.to_datetime(df["period_start"], errors="coerce")
        df["period_end_dt"] = pd.to_datetime(df["period_end"], errors="coerce")
        instant_mask = (df["period_type"] == "instant") & df["period_start_dt"].isna() & df["period_end_dt"].notna()
        if instant_mask.any():
            df.loc[instant_mask, "period_start_dt"] = df.loc[instant_mask, "period_end_dt"]
            df.loc[instant_mask, "period_start"] = df.loc[instant_mask, "period_end"]
        df["period_end"] = df["period_end_dt"].dt.date.astype(str)

        df["duration_days"] = (df["period_end_dt"] - df["period_start_dt"]).dt.days

        df["has_currency"] = df["currency"].notna() & (df["currency"].astype(str).str.len() > 0)
        df["abs_numeric"] = df["numeric_value"].abs()

        if "dimensions" in df.columns:
            dim_series = df["dimensions"].astype(str).str.strip()
            dim_mask = dim_series.isna() | dim_series.eq("") | dim_series.eq("{}") | dim_series.eq("nan") | dim_series.eq("None")
            allowed_mask = dim_series.apply(self._is_allowed_dimension)
            df["dimension_allowed"] = dim_mask | allowed_mask
        else:
            df["dimension_allowed"] = True

        df = df[df["is_text_block"] != True].copy()
        text_block_mask = (
            df["Element"].astype(str).str.endswith("TextBlock")
            | df["Tag"].astype(str).str.endswith("TextBlock")
            | df["Label"].astype(str).str.endswith("TextBlock")
        )
        df = df.loc[~text_block_mask]
        df = df[df["period_end_dt"].notna()]

        return df

    def _is_allowed_dimension(self, dim_value: str) -> bool:
        """Allow consolidation-only dimensions while rejecting segment-level splits."""
        if not dim_value or dim_value in {"{}", "nan", "None"}:
            return False
        try:
            parsed = json.loads(dim_value)
        except Exception:
            return False
        if not isinstance(parsed, dict) or not parsed:
            return False
        for key in parsed.keys():
            if not any(str(key).endswith(suffix) for suffix in self._ALLOWED_DIMENSION_SUFFIXES):
                return False
        return True

    def _infer_standard(self, facts: pd.DataFrame) -> Optional[str]:
        """Infer the accounting standard from available facts."""
        if "Standard" not in facts.columns:
            return None

        series = facts["Standard"].dropna()
        if series.empty:
            return None

        return series.mode().iloc[0]

    @staticmethod
    def _build_traces(series_defs: Sequence[Dict[str, Any]], keys: Sequence[str]) -> List[Dict[str, Any]]:
        """Build trace definitions from series configs."""
        color_map = {spec["column_name"]: spec.get("color_key", "navy") for spec in series_defs}
        traces: List[Dict[str, Any]] = []
        for key in keys:
            if key in color_map:
                traces.append({"col": key, "name": key, "color_key": color_map[key]})
        return traces

    @staticmethod
    def _build_portfolio_traces(series_defs: Sequence[Dict[str, Any]], keys: Sequence[str]) -> List[Dict[str, Any]]:
        """Build portfolio trace definitions from series configs."""
        available_keys = set(keys) if keys else {spec.get("column_name") for spec in series_defs}
        traces: List[Dict[str, Any]] = []
        for spec in series_defs:
            col_name = spec.get("column_name")
            if not col_name or col_name not in available_keys:
                continue
            trace = {
                "col": col_name,
                "name": spec.get("display_name", col_name),
                "color_key": spec.get("color_key", "navy"),
                "chart_type": spec.get("chart_type", "area"),
            }
            if "line_width" in spec:
                trace["line_width"] = spec["line_width"]
            if "marker_size" in spec:
                trace["marker_size"] = spec["marker_size"]
            traces.append(trace)
        return traces

    @staticmethod
    def _get_period_columns(df: pd.DataFrame) -> List[str]:
        """Return period columns from canonical wide tables."""
        if df is None or df.empty:
            return []
        reserved = {"canonical_key", "label", "statement"}
        return [col for col in df.columns if col not in reserved]

    @staticmethod
    def _select_latest_period(period_cols: Sequence[str]) -> Optional[str]:
        """Return the latest period label."""
        if not period_cols:
            return None
        parsed = []
        for col in period_cols:
            match = re.match(r"(\\d{4}-\\d{2}-\\d{2})-(\\d{4}-\\d{2}-\\d{2})", str(col))
            if match:
                parsed.append((col, match.group(2)))
        if parsed:
            return sorted(parsed, key=lambda x: x[1])[-1][0]
        return sorted(period_cols)[-1]

    def _build_trend_dataframe(self, df_wide: pd.DataFrame, series_defs: Sequence[Dict[str, Any]]) -> pd.DataFrame:
        """Build a trend table with period labels as rows and metrics as columns."""
        if df_wide is None or df_wide.empty:
            cols = ["period_label"] + [spec["column_name"] for spec in series_defs]
            return pd.DataFrame(columns=cols)
        period_cols = self._get_period_columns(df_wide)
        data: Dict[str, List[Any]] = {"period_label": period_cols}
        for spec in series_defs:
            row = df_wide[df_wide["canonical_key"] == spec["canonical_key"]]
            if row.empty:
                data[spec["column_name"]] = [None] * len(period_cols)
                continue
            row = row.iloc[0]
            data[spec["column_name"]] = [row.get(col) for col in period_cols]
        return pd.DataFrame(data)

    def _build_snapshot_dataframe(
        self,
        df_wide: pd.DataFrame,
        series_defs: Sequence[Dict[str, Any]],
        period_label: Optional[str] = None,
    ) -> pd.DataFrame:
        """Build a single-row snapshot table for a selected period."""
        if df_wide is None or df_wide.empty:
            cols = ["period_label"] + [spec["column_name"] for spec in series_defs]
            return pd.DataFrame(columns=cols)
        period_cols = self._get_period_columns(df_wide)
        target_period = period_label or self._select_latest_period(period_cols)

        if not period_label and period_cols:
            ordered_cols = sorted(period_cols, reverse=True)
            preferred_row = (
                df_wide[df_wide["canonical_key"] == "TotalAssets"]
                if any(spec.get("column_name") == "Total Assets" for spec in series_defs)
                else pd.DataFrame()
            )
            preferred_cols = []
            if not preferred_row.empty:
                for col in ordered_cols:
                    val = pd.to_numeric(preferred_row.iloc[0].get(col), errors="coerce")
                    if pd.notna(val):
                        preferred_cols.append(col)
            if preferred_cols:
                target_period = preferred_cols[0]
            else:
                candidate = None
                for col in ordered_cols:
                    for spec in series_defs:
                        row = df_wide[df_wide["canonical_key"] == spec["canonical_key"]]
                        if row.empty:
                            continue
                        val = pd.to_numeric(row.iloc[0].get(col), errors="coerce")
                        if pd.notna(val):
                            candidate = col
                            break
                    if candidate:
                        break
                if candidate:
                    target_period = candidate

        row_data: Dict[str, Any] = {"period_label": target_period}
        for spec in series_defs:
            row = df_wide[df_wide["canonical_key"] == spec["canonical_key"]]
            if row.empty or not target_period:
                row_data[spec["column_name"]] = None
                continue
            row_data[spec["column_name"]] = row.iloc[0].get(target_period)
        return pd.DataFrame([row_data])

    def _reconcile_balance_sheet(self, df_wide: pd.DataFrame, tolerance: float = 0.03) -> pd.DataFrame:
        """Fill or adjust TotalLiabilities when assets/equity are available but mismatch is large."""
        if df_wide is None or df_wide.empty:
            return df_wide

        df = df_wide.copy()
        period_cols = self._get_period_columns(df)
        if not period_cols:
            return df

        assets_row = df[df["canonical_key"] == "TotalAssets"]
        equity_row = df[df["canonical_key"] == "TotalEquity"]
        liab_row = df[df["canonical_key"] == "TotalLiabilities"]
        if assets_row.empty or equity_row.empty:
            return df

        if liab_row.empty:
            new_row = {"canonical_key": "TotalLiabilities", "label": "Total Liabilities (Derived)", "statement": StatementType.BS.value}
            for col in period_cols:
                new_row[col] = None
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            liab_row = df[df["canonical_key"] == "TotalLiabilities"]

        for col in period_cols:
            ta = pd.to_numeric(assets_row.iloc[0].get(col), errors="coerce")
            te = pd.to_numeric(equity_row.iloc[0].get(col), errors="coerce")
            tl = pd.to_numeric(liab_row.iloc[0].get(col), errors="coerce")
            if pd.isna(ta) or pd.isna(te):
                continue
            derived = ta - te
            if pd.isna(tl):
                df.loc[df["canonical_key"] == "TotalLiabilities", col] = derived
                continue
            if ta != 0 and abs(ta - (tl + te)) / abs(ta) > tolerance:
                df.loc[df["canonical_key"] == "TotalLiabilities", col] = derived

        return df


def load_mapping_with_override(
    override_path: Optional[Path],
    standard: Optional[str] = None,
    company_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Load mapping from the column definition config and merge a JSON override if provided."""
    base = ColumnDefinitionConfig.resolve_mapping(standard=standard, company_name=company_name)
    if not override_path:
        return base
    if not override_path.exists():
        logging.warning("Mapping override not found: %s; using default mapping.", override_path)
        return base

    overlay = ColumnDefinitionConfig.load_json(override_path)
    return ColumnDefinitionConfig.merge(base, overlay)
