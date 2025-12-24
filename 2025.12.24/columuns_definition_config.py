from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, Optional, Sequence, Tuple

from financial_mapping import MappingConfig


@dataclass(frozen=True)
class MappingItemSpec:
    """Declarative spec for a canonical item override."""

    canonical_key: str
    statement: str
    period_type: Optional[str]
    element_exact: Tuple[str, ...] = ()
    element_regex: Tuple[str, ...] = ()
    tag_exact: Tuple[str, ...] = ()
    label_regex: Tuple[str, ...] = ()


def build_mapping_override(items: Sequence[MappingItemSpec]) -> Dict[str, Any]:
    """Build a mapping dict from ordered item specs."""
    mapping_items: list[Dict[str, Any]] = []

    for item in items:
        candidates: list[Dict[str, Any]] = []

        def _append_exact(field: str, values: Iterable[str], base_weight: float) -> None:
            """Append exact-match candidates with descending weights."""
            for idx, value in enumerate(values):
                candidates.append(
                    {
                        "field": field,
                        "exact": value,
                        "weight": round(base_weight - (idx * 0.1), 3),
                    }
                )

        def _append_regex(field: str, values: Iterable[str], base_weight: float) -> None:
            """Append regex-match candidates with descending weights."""
            for idx, value in enumerate(values):
                candidates.append(
                    {
                        "field": field,
                        "regex": value,
                        "weight": round(base_weight - (idx * 0.1), 3),
                    }
                )

        _append_exact("element", item.element_exact, base_weight=4.0)
        _append_exact("tag", item.tag_exact, base_weight=3.0)
        _append_regex("element", item.element_regex, base_weight=2.5)
        _append_regex("label", item.label_regex, base_weight=1.5)

        mapping_items.append(
            {
                "canonical_key": item.canonical_key,
                "statement": item.statement,
                "period_type": item.period_type,
                "candidates": candidates,
            }
        )

    return {"version": 1, "items": mapping_items}


_IFRS_OVERRIDES = build_mapping_override(
    [
        MappingItemSpec(
            canonical_key="Revenue",
            statement="PL",
            period_type="duration",
            element_exact=(
                "RevenueIFRSSummaryOfBusinessResults",
                "RevenueIFRS",
                "NetSalesIFRS",
                "SalesRevenuesIFRS",
                "TotalNetRevenuesIFRS",
                "OperatingRevenuesIFRSKeyFinancialData",
                "OperatingRevenueFromExternalCustomersIFRS",
                "NetSalesSummaryOfBusinessResults",
            ),
            element_regex=(
                r"SalesRevenueNet",
                r"OperatingRevenue",
            ),
            label_regex=(r"\\b(Revenue|Net Sales|Sales Revenue|Operating Revenue)\\b",),
        ),
        MappingItemSpec(
            canonical_key="OperatingIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "OperatingProfitLossIFRSSummaryOfBusinessResults",
                "OperatingProfitLossIFRS",
                "OperatingIncome",
                "OperatingProfit",
            ),
            element_regex=(r"OperatingIncomeSummaryOfBusinessResults",),
            label_regex=(r"Operating (Income|Profit)",),
        ),
        MappingItemSpec(
            canonical_key="PretaxIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "ProfitLossBeforeTaxIFRSSummaryOfBusinessResults",
                "ProfitLossBeforeTaxIFRS",
                "IncomeBeforeIncomeTaxes",
                "IncomeBeforeIncomeTaxesIFRS",
            ),
            element_regex=(r"ProfitLossBeforeTax",),
            label_regex=(r"Before Tax|Pretax",),
        ),
        MappingItemSpec(
            canonical_key="NetIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "ProfitLossAttributableToOwnersOfParentIFRSSummaryOfBusinessResults",
                "ProfitLossAttributableToOwnersOfParentIFRS",
                "ProfitLossIFRS",
                "ProfitLoss",
            ),
            element_regex=(r"NetIncomeLossSummaryOfBusinessResults",),
            label_regex=(r"Net (Income|Profit)",),
        ),
    ]
)


_JGAAP_OVERRIDES = build_mapping_override(
    [
        MappingItemSpec(
            canonical_key="Revenue",
            statement="PL",
            period_type="duration",
            element_exact=(
                "NetSalesSummaryOfBusinessResults",
                "NetSales",
                "OperatingRevenueSummaryOfBusinessResults",
            ),
            element_regex=(r"SalesRevenueNet",),
            label_regex=(r"\\b(Revenue|Net Sales|Sales Revenue|Operating Revenue)\\b",),
        ),
        MappingItemSpec(
            canonical_key="OperatingIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "OperatingIncomeSummaryOfBusinessResults",
                "OperatingIncome",
                "OperatingProfit",
            ),
            element_regex=(r"OperatingIncome",),
            label_regex=(r"Operating (Income|Profit)",),
        ),
        MappingItemSpec(
            canonical_key="PretaxIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "ProfitLossBeforeTaxSummaryOfBusinessResults",
                "ProfitLossBeforeTax",
                "IncomeBeforeIncomeTaxes",
            ),
            element_regex=(r"ProfitLossBeforeTax|IncomeBeforeIncomeTaxes|IncomeBeforeTax",),
            label_regex=(r"Before Tax|Pretax",),
        ),
        MappingItemSpec(
            canonical_key="NetIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "ProfitLossAttributableToOwnersOfParentSummaryOfBusinessResults",
                "ProfitLossAttributableToOwnersOfParent",
                "NetIncomeLossSummaryOfBusinessResults",
                "ProfitLoss",
            ),
            element_regex=(r"NetIncome|NetProfit",),
            label_regex=(r"Net (Income|Profit)",),
        ),
    ]
)


_USGAAP_OVERRIDES = build_mapping_override(
    [
        MappingItemSpec(
            canonical_key="Revenue",
            statement="PL",
            period_type="duration",
            element_exact=(
                "RevenuesUSGAAPSummaryOfBusinessResults",
                "SalesRevenueNet",
            ),
            element_regex=(r"Revenue|NetSales",),
            label_regex=(r"\\b(Revenue|Net Sales|Sales Revenue)\\b",),
        ),
        MappingItemSpec(
            canonical_key="OperatingIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "OperatingIncomeSummaryOfBusinessResults",
                "OperatingIncome",
                "OperatingProfit",
            ),
            element_regex=(r"OperatingIncome",),
            label_regex=(r"Operating (Income|Profit)",),
        ),
        MappingItemSpec(
            canonical_key="PretaxIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "IncomeBeforeIncomeTaxesUSGAAPSummaryOfBusinessResults",
                "IncomeBeforeIncomeTaxes",
            ),
            element_regex=(r"IncomeBeforeTax|ProfitLossBeforeTax",),
            label_regex=(r"Before Tax|Pretax",),
        ),
        MappingItemSpec(
            canonical_key="NetIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "NetIncomeLossAttributableToOwnersOfParentUSGAAPSummaryOfBusinessResults",
                "NetIncomeLossSummaryOfBusinessResults",
                "NetIncomeLoss",
                "NetIncome",
            ),
            element_regex=(r"ProfitLoss",),
            label_regex=(r"Net (Income|Profit)",),
        ),
    ]
)


_MUFG_OVERRIDES = build_mapping_override(
    [
        MappingItemSpec(
            canonical_key="Revenue",
            statement="PL",
            period_type="duration",
            element_exact=(
                "OrdinaryIncomeSummaryOfBusinessResults",
                "OrdinaryIncomeBNK",
            ),
        ),
        MappingItemSpec(
            canonical_key="OperatingIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "OrdinaryIncomeLossSummaryOfBusinessResults",
                "OrdinaryIncome",
            ),
        ),
        MappingItemSpec(
            canonical_key="NetIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "ProfitLossAttributableToOwnersOfParentSummaryOfBusinessResults",
                "ProfitLossAttributableToOwnersOfParent",
                "ProfitLoss",
            ),
        ),
    ]
)


_TOYOTA_OVERRIDES = build_mapping_override(
    [
        MappingItemSpec(
            canonical_key="Revenue",
            statement="PL",
            period_type="duration",
            element_exact=(
                "TotalNetRevenuesIFRS",
                "OperatingRevenuesIFRSKeyFinancialData",
                "SalesRevenuesIFRS",
                "RevenuesUSGAAPSummaryOfBusinessResults",
            ),
        ),
        MappingItemSpec(
            canonical_key="PretaxIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "ProfitLossBeforeTaxIFRSSummaryOfBusinessResults",
                "ProfitLossBeforeTaxIFRS",
                "ProfitLossBeforeTaxUSGAAPSummaryOfBusinessResults",
            ),
        ),
        MappingItemSpec(
            canonical_key="NetIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "ProfitLossAttributableToOwnersOfParentIFRSSummaryOfBusinessResults",
                "ProfitLossAttributableToOwnersOfParentIFRS",
                "NetIncomeLossAttributableToOwnersOfParentUSGAAPSummaryOfBusinessResults",
                "NetIncomeLossSummaryOfBusinessResults",
                "ProfitLoss",
            ),
        ),
    ]
)


_HONDA_OVERRIDES = build_mapping_override(
    [
        MappingItemSpec(
            canonical_key="Revenue",
            statement="PL",
            period_type="duration",
            element_exact=(
                "RevenueIFRSSummaryOfBusinessResults",
                "RevenueIFRS",
                "RevenuesUSGAAPSummaryOfBusinessResults",
            ),
        ),
        MappingItemSpec(
            canonical_key="PretaxIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "ProfitLossBeforeTaxIFRSSummaryOfBusinessResults",
                "ProfitLossBeforeTaxIFRS",
                "ProfitLossBeforeTaxUSGAAPSummaryOfBusinessResults",
            ),
        ),
        MappingItemSpec(
            canonical_key="NetIncome",
            statement="PL",
            period_type="duration",
            element_exact=(
                "ProfitLossAttributableToOwnersOfParentIFRSSummaryOfBusinessResults",
                "ProfitLossAttributableToOwnersOfParentIFRS",
                "NetIncomeLossAttributableToOwnersOfParentUSGAAPSummaryOfBusinessResults",
                "NetIncomeLossSummaryOfBusinessResults",
                "ProfitLoss",
            ),
        ),
    ]
)


_DEFAULT_LAYOUT: Dict[str, Any] = {
    "pl_series": [
        {"canonical_key": "Revenue", "column_name": "Revenue", "color_key": "navy"},
        {"canonical_key": "OperatingIncome", "column_name": "Operating Income", "color_key": "red"},
    ],
    "cf_series": [
        {"canonical_key": "OperatingCashFlow", "column_name": "Operating CF", "color_key": "navy"},
        {"canonical_key": "InvestingCashFlow", "column_name": "Investing CF", "color_key": "cadet_blue"},
        {"canonical_key": "FinancingCashFlow", "column_name": "Financing CF", "color_key": "salmon"},
    ],
    "bs_series": [
        {"canonical_key": "TotalAssets", "column_name": "Total Assets", "color_key": "navy"},
        {"canonical_key": "TotalLiabilities", "column_name": "Total Liabilities", "color_key": "gray_medium"},
        {"canonical_key": "TotalEquity", "column_name": "Total Equity", "color_key": "sky_blue"},
        {"canonical_key": "CashAndCashEquivalents", "column_name": "Cash & Equivalents", "color_key": "mint"},
        {"canonical_key": "AccountsReceivable", "column_name": "Accounts Receivable", "color_key": "sky_blue"},
        {
            "canonical_key": "FinancialServicesReceivablesCurrent",
            "column_name": "Financial Services Receivables (Current)",
            "color_key": "cadet_blue",
        },
        {
            "canonical_key": "FinancialServicesReceivablesNonCurrent",
            "column_name": "Financial Services Receivables (Non-Current)",
            "color_key": "cadet_blue",
        },
        {"canonical_key": "Inventories", "column_name": "Inventories", "color_key": "mustard"},
        {"canonical_key": "PropertyPlantAndEquipment", "column_name": "PPE", "color_key": "navy"},
        {"canonical_key": "RightOfUseAssets", "column_name": "Right-of-Use Assets", "color_key": "teal"},
        {"canonical_key": "IntangibleAssets", "column_name": "Intangibles", "color_key": "lavender"},
        {"canonical_key": "Goodwill", "column_name": "Goodwill", "color_key": "pale_pink"},
        {"canonical_key": "Investments", "column_name": "Investments", "color_key": "gold"},
        {"canonical_key": "OtherFinancialAssets", "column_name": "Other Financial Assets", "color_key": "cream"},
        {
            "canonical_key": "OtherFinancialAssetsCurrent",
            "column_name": "Other Financial Assets (Current)",
            "color_key": "cream",
        },
        {
            "canonical_key": "OtherFinancialAssetsNonCurrent",
            "column_name": "Other Financial Assets (Non-Current)",
            "color_key": "cream",
        },
        {"canonical_key": "AssetsHeldForSale", "column_name": "Assets Held for Sale", "color_key": "mustard"},
        {"canonical_key": "DeferredTaxAssets", "column_name": "Deferred Tax Assets", "color_key": "spring_green"},
        {"canonical_key": "OtherCurrentAssets", "column_name": "Other Current Assets", "color_key": "ice_blue"},
        {"canonical_key": "OtherNonCurrentAssets", "column_name": "Other Non-Current Assets", "color_key": "gray_light"},
        {"canonical_key": "OtherAssets", "column_name": "Other Assets", "color_key": "gray_light"},
        {"canonical_key": "CurrentAssets", "column_name": "Current Assets", "color_key": "mint"},
        {"canonical_key": "NonCurrentAssets", "column_name": "Non-Current Assets", "color_key": "ice_blue"},
        {"canonical_key": "AccountsPayable", "column_name": "Accounts Payable", "color_key": "gray_medium"},
        {"canonical_key": "ShortTermBorrowings", "column_name": "Short-Term Borrowings", "color_key": "dark_slate"},
        {"canonical_key": "LongTermBorrowings", "column_name": "Long-Term Borrowings", "color_key": "dark_slate"},
        {"canonical_key": "BondsPayable", "column_name": "Bonds Payable", "color_key": "coral"},
        {"canonical_key": "LeaseLiabilities", "column_name": "Lease Liabilities", "color_key": "gray_light"},
        {"canonical_key": "Provisions", "column_name": "Provisions", "color_key": "cream"},
        {"canonical_key": "ProvisionsCurrent", "column_name": "Provisions (Current)", "color_key": "cream"},
        {"canonical_key": "ProvisionsNonCurrent", "column_name": "Provisions (Non-Current)", "color_key": "cream"},
        {"canonical_key": "AccruedExpenses", "column_name": "Accrued Expenses", "color_key": "gray_medium"},
        {"canonical_key": "IncomeTaxesPayable", "column_name": "Income Taxes Payable", "color_key": "gray_medium"},
        {"canonical_key": "RetirementBenefitLiabilities", "column_name": "Retirement Benefit Liabilities", "color_key": "gray_medium"},
        {
            "canonical_key": "RetirementBenefitLiabilitiesCurrent",
            "column_name": "Retirement Benefit Liabilities (Current)",
            "color_key": "gray_medium",
        },
        {
            "canonical_key": "RetirementBenefitLiabilitiesNonCurrent",
            "column_name": "Retirement Benefit Liabilities (Non-Current)",
            "color_key": "gray_medium",
        },
        {"canonical_key": "DeferredTaxLiabilities", "column_name": "Deferred Tax Liabilities", "color_key": "teal"},
        {"canonical_key": "OtherFinancialLiabilities", "column_name": "Other Financial Liabilities", "color_key": "cadet_blue"},
        {
            "canonical_key": "OtherFinancialLiabilitiesCurrent",
            "column_name": "Other Financial Liabilities (Current)",
            "color_key": "cadet_blue",
        },
        {
            "canonical_key": "OtherFinancialLiabilitiesNonCurrent",
            "column_name": "Other Financial Liabilities (Non-Current)",
            "color_key": "cadet_blue",
        },
        {"canonical_key": "OtherCurrentLiabilities", "column_name": "Other Current Liabilities", "color_key": "gray_light"},
        {"canonical_key": "OtherNonCurrentLiabilities", "column_name": "Other Non-Current Liabilities", "color_key": "gray_light"},
        {"canonical_key": "LiabilitiesHeldForSale", "column_name": "Liabilities Held for Sale", "color_key": "gray_light"},
        {"canonical_key": "OtherLiabilities", "column_name": "Other Liabilities", "color_key": "gray_light"},
        {"canonical_key": "ShareCapital", "column_name": "Share Capital", "color_key": "midnight_blue"},
        {"canonical_key": "CapitalSurplus", "column_name": "Capital Surplus", "color_key": "cadet_blue"},
        {"canonical_key": "RetainedEarnings", "column_name": "Retained Earnings", "color_key": "spring_green"},
        {"canonical_key": "OtherComponentsOfEquity", "column_name": "Other Components of Equity", "color_key": "gold"},
        {"canonical_key": "TreasuryShares", "column_name": "Treasury Shares", "color_key": "gray_dark"},
        {"canonical_key": "AccumulatedOtherComprehensiveIncome", "column_name": "AOCI", "color_key": "gold"},
        {"canonical_key": "NonControllingInterests", "column_name": "Non-Controlling Interests", "color_key": "gray_light"},
        {"canonical_key": "CurrentLiabilities", "column_name": "Current Liabilities", "color_key": "gray_medium"},
        {"canonical_key": "NonCurrentLiabilities", "column_name": "Non-Current Liabilities", "color_key": "gray_light"},
        {"canonical_key": "Loans", "column_name": "Loans", "color_key": "cadet_blue"},
        {"canonical_key": "Deposits", "column_name": "Deposits", "color_key": "salmon"},
        {"canonical_key": "Securities", "column_name": "Securities", "color_key": "gold"},
        {"canonical_key": "Borrowings", "column_name": "Borrowings", "color_key": "dark_slate"},
        {"canonical_key": "TradingAssets", "column_name": "Trading Assets", "color_key": "mustard"},
        {
            "canonical_key": "ReceivablesUnderResaleAgreements",
            "column_name": "Receivables Under Resale Agreements",
            "color_key": "cadet_blue",
        },
        {
            "canonical_key": "CustomersLiabilitiesForAcceptancesAndGuarantees",
            "column_name": "Customers Liabilities for Acceptances & Guarantees",
            "color_key": "spring_green",
        },
        {"canonical_key": "MonetaryClaimsBought", "column_name": "Monetary Claims Bought", "color_key": "lavender"},
        {
            "canonical_key": "ReceivablesUnderSecuritiesBorrowingTransactions",
            "column_name": "Receivables Under Securities Borrowing Transactions",
            "color_key": "sky_blue",
        },
        {"canonical_key": "RepoLiabilities", "column_name": "Repo Liabilities", "color_key": "coral"},
        {"canonical_key": "TradingLiabilities", "column_name": "Trading Liabilities", "color_key": "coral"},
        {"canonical_key": "NegotiableCertificatesOfDeposit", "column_name": "Negotiable CDs", "color_key": "gold"},
        {"canonical_key": "AcceptancesAndGuarantees", "column_name": "Acceptances & Guarantees", "color_key": "cream"},
        {"canonical_key": "CallMoneyAndBillsSold", "column_name": "Call Money & Bills Sold", "color_key": "coral"},
        {"canonical_key": "TrustAccountBorrowings", "column_name": "Trust Account Borrowings", "color_key": "dark_slate"},
        {"canonical_key": "CommercialPapers", "column_name": "Commercial Papers", "color_key": "mustard"},
        {"canonical_key": "ForeignExchangeLiabilities", "column_name": "FX Liabilities", "color_key": "cadet_blue"},
        {"canonical_key": "SecuritiesLendingPayables", "column_name": "Securities Lending Payables", "color_key": "gray_medium"},
    ],
    "portfolio_series": [
        {"portfolio_key": "EquitySecurities", "column_name": "Equity Securities", "color_key": "navy", "chart_type": "area"},
        {"portfolio_key": "DebtSecurities", "column_name": "Debt Securities", "color_key": "cadet_blue", "chart_type": "area"},
        {"portfolio_key": "DerivativeAssets", "column_name": "Derivatives (Assets)", "color_key": "salmon", "chart_type": "area"},
        {"portfolio_key": "DerivativeLiabilities", "column_name": "Derivatives (Liabilities)", "color_key": "gray_medium", "chart_type": "area"},
        {"portfolio_key": "DerivativeNet", "column_name": "Derivatives (Net)", "color_key": "red", "chart_type": "line"},
    ],
    "pl_chart": {
        "slide_title": "PL",
        "category": "combo_bar_line_2axis",
        "x_col": "period_label",
        "x_label_format": "fy",
        "unit_scale": 1e9,
        "bar_keys": ["Revenue"],
        "line_keys": ["Operating Income"],
        "chart_text": {"title": "Revenue and Operating Income", "y1_label": "JPY (bn)", "y2_label": "JPY (bn)"},
    },
    "cf_chart": {
        "slide_title": "CF",
        "category": "combo_bar_line_2axis",
        "x_col": "period_label",
        "x_label_format": "fy",
        "unit_scale": 1e9,
        "bar_keys": ["Operating CF", "Investing CF", "Financing CF"],
        "line_keys": [],
        "chart_text": {"title": "Cash Flow Summary", "y1_label": "JPY (bn)", "y2_label": ""},
    },
    "bs_chart": {
        "slide_title": "BS",
        "category": "balance_sheet",
        "unit_scale": 1e9,
        "total_assets_col": "Total Assets",
        "stack_preference": "summary",
        "auto_balance_assets": True,
        "auto_balance_liab_equity": True,
        "balance_tolerance": 0.005,
        "left_label": "Assets",
        "right_label": "Liabilities & Equity",
        "show_legend": True,
        "legend_max_items": 5,
        "legend_source": "detail",
        "show_segment_labels": False,
        "show_summary_labels": True,
        "summary_label_max_length": 30,
        "summary_label_position": "inside",
        "segment_label_min_ratio": 0.1,
        "segment_label_font_size": 9,
        "segment_label_color_key": "gray_dark",
        "collapse_equity_on_negative": True,
        "group_label_color_key": "gray_dark",
        "group_label_font_size": 9,
        "group_label_position": "inside",
        "avoid_label_overlap": True,
        "detail_bars": False,
        "is_bank": False,
        "exclusive_groups": [
            {
                "aggregate": "Other Financial Assets",
                "components": ["Other Financial Assets (Current)", "Other Financial Assets (Non-Current)"],
            },
            {
                "aggregate": "Other Financial Liabilities",
                "components": ["Other Financial Liabilities (Current)", "Other Financial Liabilities (Non-Current)"],
            },
            {"aggregate": "Provisions", "components": ["Provisions (Current)", "Provisions (Non-Current)"]},
            {
                "aggregate": "Retirement Benefit Liabilities",
                "components": ["Retirement Benefit Liabilities (Current)", "Retirement Benefit Liabilities (Non-Current)"],
            },
            {
                "aggregate": "Other Assets",
                "components": ["Other Current Assets", "Other Non-Current Assets"],
            },
            {
                "aggregate": "Total Equity",
                "components": [
                    "Share Capital",
                    "Capital Surplus",
                    "Retained Earnings",
                    "Other Components of Equity",
                    "AOCI",
                    "Treasury Shares",
                    "Non-Controlling Interests",
                ],
            },
        ],
        "left_groups": [
            {
                "label": "Current Assets",
                "items": [
                    "Cash & Equivalents",
                    "Accounts Receivable",
                    "Financial Services Receivables (Current)",
                    "Inventories",
                    "Other Financial Assets (Current)",
                    "Other Current Assets",
                    "Assets Held for Sale",
                ],
            },
            {
                "label": "Non-Current Assets",
                "items": [
                    "Financial Services Receivables (Non-Current)",
                    "Investments",
                    "PPE",
                    "Right-of-Use Assets",
                    "Intangibles",
                    "Goodwill",
                    "Other Financial Assets (Non-Current)",
                    "Other Financial Assets",
                    "Deferred Tax Assets",
                    "Other Non-Current Assets",
                    "Other Assets",
                ],
            },
        ],
        "right_groups": [
            {
                "label": "Current Liabilities",
                "items": [
                    "Accounts Payable",
                    "Short-Term Borrowings",
                    "Provisions (Current)",
                    "Accrued Expenses",
                    "Income Taxes Payable",
                    "Retirement Benefit Liabilities (Current)",
                    "Other Financial Liabilities (Current)",
                    "Other Current Liabilities",
                    "Liabilities Held for Sale",
                ],
            },
            {
                "label": "Non-Current Liabilities",
                "items": [
                    "Long-Term Borrowings",
                    "Bonds Payable",
                    "Lease Liabilities",
                    "Provisions (Non-Current)",
                    "Retirement Benefit Liabilities (Non-Current)",
                    "Deferred Tax Liabilities",
                    "Other Financial Liabilities (Non-Current)",
                    "Other Non-Current Liabilities",
                    "Other Liabilities",
                ],
            },
            {
                "label": "Equity",
                "items": [
                    "Share Capital",
                    "Capital Surplus",
                    "Retained Earnings",
                    "Other Components of Equity",
                    "AOCI",
                    "Treasury Shares",
                    "Non-Controlling Interests",
                    "Total Equity",
                ],
            },
        ],
        "left_groups_for_bank": [
            {
                "label": "Loans & Securities",
                "items": [
                    "Loans",
                    "Securities",
                    "Investments",
                    "Monetary Claims Bought",
                ],
            },
            {
                "label": "Trading & Resale",
                "items": [
                    "Trading Assets",
                    "Receivables Under Resale Agreements",
                    "Receivables Under Securities Borrowing Transactions",
                    "Customers Liabilities for Acceptances & Guarantees",
                ],
            },
            {
                "label": "Other Assets",
                "items": [
                    "Cash & Equivalents",
                    "PPE",
                    "Intangibles",
                    "Goodwill",
                    "Deferred Tax Assets",
                    "Other Financial Assets",
                    "Other Assets",
                ],
            },
        ],
        "right_groups_for_bank": [
            {
                "label": "Customer Funding",
                "items": ["Deposits"],
            },
            {
                "label": "Market Funding",
                "items": [
                    "Repo Liabilities",
                    "Borrowings",
                    "Negotiable CDs",
                    "Call Money & Bills Sold",
                    "Trust Account Borrowings",
                    "Commercial Papers",
                ],
            },
            {
                "label": "Other Liabilities",
                "items": [
                    "Trading Liabilities",
                    "Acceptances & Guarantees",
                    "FX Liabilities",
                    "Securities Lending Payables",
                    "Deferred Tax Liabilities",
                    "Bonds Payable",
                    "Other Liabilities",
                ],
            },
            {
                "label": "Equity",
                "items": ["Total Equity"],
            },
        ],
        "left_stack": [
            {"col": "Cash & Equivalents", "name": "Cash & Equivalents", "color_key": "mint"},
            {"col": "Accounts Receivable", "name": "Accounts Receivable", "color_key": "sky_blue"},
            {
                "col": "Financial Services Receivables (Current)",
                "name": "Financial Services Receivables (Current)",
                "color_key": "cadet_blue",
            },
            {"col": "Inventories", "name": "Inventories", "color_key": "mustard"},
            {"col": "Other Financial Assets (Current)", "name": "Other Financial Assets (Current)", "color_key": "cream"},
            {"col": "Other Current Assets", "name": "Other Current Assets", "color_key": "ice_blue"},
            {"col": "Assets Held for Sale", "name": "Assets Held for Sale", "color_key": "mustard"},
            {
                "col": "Financial Services Receivables (Non-Current)",
                "name": "Financial Services Receivables (Non-Current)",
                "color_key": "cadet_blue",
            },
            {"col": "Investments", "name": "Investments", "color_key": "gold"},
            {"col": "PPE", "name": "PPE", "color_key": "navy"},
            {"col": "Right-of-Use Assets", "name": "Right-of-Use Assets", "color_key": "teal"},
            {"col": "Intangibles", "name": "Intangibles", "color_key": "lavender"},
            {"col": "Goodwill", "name": "Goodwill", "color_key": "pale_pink"},
            {"col": "Other Financial Assets", "name": "Other Financial Assets", "color_key": "cream"},
            {
                "col": "Other Financial Assets (Non-Current)",
                "name": "Other Financial Assets (Non-Current)",
                "color_key": "cream",
            },
            {"col": "Deferred Tax Assets", "name": "Deferred Tax Assets", "color_key": "spring_green"},
            {"col": "Other Non-Current Assets", "name": "Other Non-Current Assets", "color_key": "gray_light"},
            {"col": "Other Assets", "name": "Other Assets", "color_key": "gray_light"},
        ],
        "left_stack_for_bank": [
            {"col": "Loans", "name": "Loans", "color_key": "cadet_blue"},
            {"col": "Securities", "name": "Securities", "color_key": "gold"},
            {"col": "Monetary Claims Bought", "name": "Monetary Claims Bought", "color_key": "lavender"},
            {
                "col": "Investments", "name": "Investments", "color_key": "gold"
            },
            {"col": "Trading Assets", "name": "Trading Assets", "color_key": "mustard"},
            {
                "col": "Receivables Under Resale Agreements",
                "name": "Receivables Under Resale Agreements",
                "color_key": "cadet_blue",
            },
            {
                "col": "Receivables Under Securities Borrowing Transactions",
                "name": "Receivables Under Securities Borrowing Transactions",
                "color_key": "sky_blue",
            },
            {
                "col": "Customers Liabilities for Acceptances & Guarantees",
                "name": "Customers Liabilities for Acceptances & Guarantees",
                "color_key": "spring_green",
            },
            {"col": "Cash & Equivalents", "name": "Cash & Equivalents", "color_key": "mint"},
            {"col": "PPE", "name": "PPE", "color_key": "navy"},
            {"col": "Intangibles", "name": "Intangibles", "color_key": "lavender"},
            {"col": "Goodwill", "name": "Goodwill", "color_key": "pale_pink"},
            {"col": "Deferred Tax Assets", "name": "Deferred Tax Assets", "color_key": "spring_green"},
            {"col": "Other Financial Assets", "name": "Other Financial Assets", "color_key": "cream"},
            {"col": "Other Assets", "name": "Other Assets", "color_key": "gray_light"},
        ],
        "left_stack_summary": [
            {"col": "Current Assets", "name": "Current Assets", "color_key": "mint"},
            {"col": "Non-Current Assets", "name": "Non-Current Assets", "color_key": "ice_blue"},
        ],
        "right_stack": [
            {"col": "Accounts Payable", "name": "Accounts Payable", "color_key": "gray_medium"},
            {"col": "Short-Term Borrowings", "name": "Short-Term Borrowings", "color_key": "dark_slate"},
            {"col": "Provisions (Current)", "name": "Provisions (Current)", "color_key": "cream"},
            {"col": "Accrued Expenses", "name": "Accrued Expenses", "color_key": "gray_medium"},
            {"col": "Income Taxes Payable", "name": "Income Taxes Payable", "color_key": "gray_medium"},
            {
                "col": "Retirement Benefit Liabilities (Current)",
                "name": "Retirement Benefit Liabilities (Current)",
                "color_key": "gray_medium",
            },
            {
                "col": "Other Financial Liabilities (Current)",
                "name": "Other Financial Liabilities (Current)",
                "color_key": "cadet_blue",
            },
            {"col": "Other Current Liabilities", "name": "Other Current Liabilities", "color_key": "gray_light"},
            {"col": "Liabilities Held for Sale", "name": "Liabilities Held for Sale", "color_key": "gray_light"},
            {"col": "Long-Term Borrowings", "name": "Long-Term Borrowings", "color_key": "dark_slate"},
            {"col": "Bonds Payable", "name": "Bonds Payable", "color_key": "coral"},
            {"col": "Lease Liabilities", "name": "Lease Liabilities", "color_key": "gray_light"},
            {"col": "Provisions", "name": "Provisions", "color_key": "cream"},
            {"col": "Provisions (Non-Current)", "name": "Provisions (Non-Current)", "color_key": "cream"},
            {
                "col": "Retirement Benefit Liabilities",
                "name": "Retirement Benefit Liabilities",
                "color_key": "gray_medium",
            },
            {
                "col": "Retirement Benefit Liabilities (Non-Current)",
                "name": "Retirement Benefit Liabilities (Non-Current)",
                "color_key": "gray_medium",
            },
            {"col": "Deferred Tax Liabilities", "name": "Deferred Tax Liabilities", "color_key": "teal"},
            {"col": "Other Financial Liabilities", "name": "Other Financial Liabilities", "color_key": "cadet_blue"},
            {
                "col": "Other Financial Liabilities (Non-Current)",
                "name": "Other Financial Liabilities (Non-Current)",
                "color_key": "cadet_blue",
            },
            {"col": "Other Non-Current Liabilities", "name": "Other Non-Current Liabilities", "color_key": "gray_light"},
            {"col": "Other Liabilities", "name": "Other Liabilities", "color_key": "gray_light"},
            {"col": "Share Capital", "name": "Share Capital", "color_key": "midnight_blue"},
            {"col": "Capital Surplus", "name": "Capital Surplus", "color_key": "cadet_blue"},
            {"col": "Retained Earnings", "name": "Retained Earnings", "color_key": "spring_green"},
            {"col": "Other Components of Equity", "name": "Other Components of Equity", "color_key": "gold"},
            {"col": "Treasury Shares", "name": "Treasury Shares", "color_key": "gray_dark"},
            {"col": "AOCI", "name": "AOCI", "color_key": "gold"},
            {"col": "Non-Controlling Interests", "name": "Non-Controlling Interests", "color_key": "gray_light"},
        ],
        "right_stack_for_bank": [
            {"col": "Deposits", "name": "Deposits", "color_key": "salmon"},
            {"col": "Repo Liabilities", "name": "Repo Liabilities", "color_key": "coral"},
            {"col": "Borrowings", "name": "Borrowings", "color_key": "dark_slate"},
            {"col": "Negotiable CDs", "name": "Negotiable CDs", "color_key": "gold"},
            {"col": "Call Money & Bills Sold", "name": "Call Money & Bills Sold", "color_key": "coral"},
            {"col": "Trust Account Borrowings", "name": "Trust Account Borrowings", "color_key": "dark_slate"},
            {"col": "Commercial Papers", "name": "Commercial Papers", "color_key": "mustard"},
            {"col": "Trading Liabilities", "name": "Trading Liabilities", "color_key": "coral"},
            {"col": "Acceptances & Guarantees", "name": "Acceptances & Guarantees", "color_key": "cream"},
            {"col": "FX Liabilities", "name": "FX Liabilities", "color_key": "cadet_blue"},
            {"col": "Securities Lending Payables", "name": "Securities Lending Payables", "color_key": "gray_medium"},
            {"col": "Deferred Tax Liabilities", "name": "Deferred Tax Liabilities", "color_key": "teal"},
            {"col": "Bonds Payable", "name": "Bonds Payable", "color_key": "coral"},
            {"col": "Other Liabilities", "name": "Other Liabilities", "color_key": "gray_light"},
            {"col": "Total Equity", "name": "Total Equity", "color_key": "sky_blue"},
        ],
        "right_stack_summary": [
            {"col": "Current Liabilities", "name": "Current Liabilities", "color_key": "gray_medium"},
            {"col": "Non-Current Liabilities", "name": "Non-Current Liabilities", "color_key": "gray_light"},
            {"col": "Total Equity", "name": "Total Equity", "color_key": "sky_blue"},
        ],
    },
    "portfolio_chart": {
        "slide_title": "Portfolio",
        "category": "portfolio_timeseries",
        "x_col": "period_label",
        "x_label_format": "fy",
        "unit_scale": 1e9,
        "series_keys": ["Equity Securities", "Debt Securities", "Derivatives (Net)"],
        "chart_style": "stacked_bar",
        "bar_width": 0.6,
        "x_label_rotation": 0,
        "stacked": True,
        "chart_text": {"title": "Portfolio Positions (Equity, Debt, Derivatives)", "y1_label": "JPY (bn)"},
        "layout_type": "horizontal",
    },
}


_COMPANY_LAYOUT_OVERRIDES: Dict[str, Dict[str, Any]] = {
    "toyota": {
        "bs_chart": {
            "left_stack_preference": "primary",
            "right_stack_preference": "primary",
            "auto_balance_assets": True,
            "auto_balance_liab_equity": True,
        }
    },
    "ajinomoto": {
        "bs_chart": {
            "left_stack_preference": "primary",
            "right_stack_preference": "primary",
            "auto_balance_assets": True,
            "auto_balance_liab_equity": True,
        }
    },
    "honda": {
        "bs_chart": {
            "left_stack_preference": "primary",
            "right_stack_preference": "primary",
            "auto_balance_assets": True,
            "auto_balance_liab_equity": True,
        }
    },
    "mufg": {
        "bs_chart": {
            "auto_balance_assets": True,
            "auto_balance_liab_equity": True,
            "is_bank": True,
        }
    },
}


class ColumnDefinitionConfig:
    """Resolve canonical mapping definitions by standard and company overrides."""

    _STANDARD_OVERRIDES: Dict[str, Dict[str, Any]] = {
        "IFRS": _IFRS_OVERRIDES,
        "JGAAP": _JGAAP_OVERRIDES,
        "USGAAP": _USGAAP_OVERRIDES,
    }
    _COMPANY_OVERRIDES: Dict[str, Dict[str, Dict[str, Any]]] = {
        "mufg": {"JGAAP": _MUFG_OVERRIDES},
        "toyota": {"IFRS": _TOYOTA_OVERRIDES},
        "honda": {"IFRS": _HONDA_OVERRIDES},
    }

    @staticmethod
    def _normalize_standard(standard: Optional[str]) -> Optional[str]:
        """Normalize standard labels to expected keys."""
        if not standard:
            return None
        normalized = str(standard).strip().upper()
        if normalized in {"JPGAAP", "JAPAN GAAP"}:
            return "JGAAP"
        return normalized

    @staticmethod
    def _normalize_company(company_name: Optional[str]) -> Optional[str]:
        """Normalize company identifiers for lookup keys."""
        if not company_name:
            return None
        return str(company_name).strip().lower()

    @classmethod
    def get_base_mapping(cls, standard: Optional[str]) -> Dict[str, Any]:
        """Return a base mapping for the given accounting standard."""
        _ = cls._normalize_standard(standard)
        return MappingConfig.default_mapping()

    @classmethod
    def get_standard_override(cls, standard: Optional[str]) -> Optional[Dict[str, Any]]:
        """Return standard-specific overrides when defined."""
        key = cls._normalize_standard(standard)
        if not key:
            return None
        return cls._STANDARD_OVERRIDES.get(key)

    @classmethod
    def get_company_override(cls, company_name: Optional[str], standard: Optional[str]) -> Optional[Dict[str, Any]]:
        """Return company-specific overrides for the given standard."""
        company_key = cls._normalize_company(company_name)
        if not company_key:
            return None
        std_key = cls._normalize_standard(standard) or "DEFAULT"
        company_bucket = cls._COMPANY_OVERRIDES.get(company_key, {})
        return company_bucket.get(std_key) or company_bucket.get("DEFAULT")

    @classmethod
    def resolve_mapping(cls, standard: Optional[str] = None, company_name: Optional[str] = None) -> Dict[str, Any]:
        """Resolve the effective mapping by layering standard and company overrides."""
        base = cls.get_base_mapping(standard)
        merged = cls.merge(base, cls.get_standard_override(standard))
        return cls.merge(merged, cls.get_company_override(company_name, standard))

    @classmethod
    def resolve_layout(cls, standard: Optional[str] = None, company_name: Optional[str] = None) -> Dict[str, Any]:
        """Resolve layout definitions for charts and snapshots."""
        _ = cls._normalize_standard(standard)
        base = json.loads(json.dumps(_DEFAULT_LAYOUT))
        company_key = cls._normalize_company(company_name)
        overlay = _COMPANY_LAYOUT_OVERRIDES.get(company_key, {}) if company_key else {}
        return cls._merge_layout(base, overlay)

    @staticmethod
    def load_json(path: Path) -> Dict[str, Any]:
        """Load a JSON mapping file."""
        with path.open("r", encoding="utf-8") as f:
            return json.load(f)

    @staticmethod
    def merge(base: Dict[str, Any], overlay: Optional[Dict[str, Any]]) -> Dict[str, Any]:
        """Merge mapping dictionaries with overlay items winning by canonical_key."""
        return MappingConfig.merge(base, overlay)

    @staticmethod
    def _merge_layout(base: Dict[str, Any], overlay: Dict[str, Any]) -> Dict[str, Any]:
        """Merge layout dictionaries with overlay keys replacing lists."""
        if not overlay:
            return base
        merged = dict(base)
        for key, value in overlay.items():
            if isinstance(value, dict) and isinstance(base.get(key), dict):
                merged[key] = ColumnDefinitionConfig._merge_layout(base.get(key, {}), value)
            else:
                merged[key] = value
        return merged
