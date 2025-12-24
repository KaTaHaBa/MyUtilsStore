from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional


@dataclass(frozen=True)
class CandidateSpec:
    """Typed candidate spec for matching facts."""

    field: str
    exact: Optional[str] = None
    regex: Optional[str] = None
    weight: float = 1.0

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "CandidateSpec":
        return cls(
            field=str(data.get("field", "")),
            exact=data.get("exact"),
            regex=data.get("regex"),
            weight=float(data.get("weight", 1.0) or 1.0),
        )

    def to_dict(self) -> Dict[str, Any]:
        payload: Dict[str, Any] = {"field": self.field, "weight": float(self.weight)}
        if self.exact is not None:
            payload["exact"] = self.exact
        if self.regex is not None:
            payload["regex"] = self.regex
        return payload


@dataclass(frozen=True)
class MappingItem:
    """Typed canonical mapping item."""

    canonical_key: str
    statement: str
    period_type: Optional[str]
    candidates: List[CandidateSpec] = field(default_factory=list)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "MappingItem":
        candidates = [CandidateSpec.from_dict(c) for c in data.get("candidates", []) if isinstance(c, dict)]
        return cls(
            canonical_key=str(data.get("canonical_key", "")),
            statement=str(data.get("statement", "")),
            period_type=data.get("period_type"),
            candidates=candidates,
        )

    def to_dict(self) -> Dict[str, Any]:
        return {
            "canonical_key": self.canonical_key,
            "statement": self.statement,
            "period_type": self.period_type,
            "candidates": [c.to_dict() for c in self.candidates],
        }


class MappingConfig:
    """Provide default mapping rules and utilities to merge overrides."""

    @staticmethod
    def default_mapping() -> Dict[str, Any]:
        """Return the default, standard-agnostic canonical mapping."""
        data = {
            "version": 1,
            "items": [
                # PL
                {
                    "canonical_key": "Revenue",
                    "statement": "PL",
                    "period_type": "duration",
                    "candidates": [
                        {"field": "element", "exact": "NetSalesIFRS", "weight": 4.0},
                        {"field": "element", "exact": "RevenueIFRSSummaryOfBusinessResults", "weight": 4.0},
                        {"field": "element", "exact": "RevenueIFRS", "weight": 3.5},
                        {"field": "element", "exact": "Revenue", "weight": 3.0},
                        {"field": "element", "regex": r"SalesRevenueNet|NetSalesIFRS|NetSales|OperatingRevenue|Revenues|RevenueIFRS", "weight": 2.5},
                        {"field": "element", "exact": "NetSalesSummaryOfBusinessResults", "weight": 1.0},
                        {"field": "element", "regex": r"OrdinaryIncomeSummaryOfBusinessResults|OperatingRevenue\\d*SummaryOfBusinessResults", "weight": 2.2},
                        {"field": "label", "regex": r"\\b(Revenue|Net Sales|Sales Revenue|Operating Revenue|Ordinary Income)\\b", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "GrossProfit",
                    "statement": "PL",
                    "period_type": "duration",
                    "candidates": [
                        {"field": "element", "regex": r"GrossProfit", "weight": 2.5},
                        {"field": "label", "regex": r"Gross Profit", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OperatingIncome",
                    "statement": "PL",
                    "period_type": "duration",
                    "candidates": [
                        {"field": "element", "regex": r"OperatingIncomeSummaryOfBusinessResults|OperatingProfitSummaryOfBusinessResults", "weight": 3.0},
                        {"field": "element", "regex": r"OperatingProfitLossIFRS|OperatingProfitLoss", "weight": 3.0},
                        {"field": "element", "regex": r"OperatingIncome|OperatingProfit", "weight": 2.5},
                        {"field": "element", "regex": r"OrdinaryIncomeSummaryOfBusinessResults", "weight": 1.0},
                        {"field": "label", "regex": r"Operating (Income|Profit)", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OrdinaryIncome",
                    "statement": "PL",
                    "period_type": "duration",
                    "candidates": [
                        {"field": "element", "regex": r"OrdinaryIncomeSummaryOfBusinessResults|OrdinaryIncomeLossSummaryOfBusinessResults", "weight": 3.0},
                        {"field": "element", "regex": r"OrdinaryIncome|OrdinaryProfit|OrdinaryIncomeLoss", "weight": 2.0},
                        {"field": "label", "regex": r"Ordinary (Income|Profit)", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "PretaxIncome",
                    "statement": "PL",
                    "period_type": "duration",
                    "candidates": [
                        {"field": "element", "regex": r"ProfitLossBeforeTax|IncomeBeforeIncomeTaxes|IncomeBeforeTax", "weight": 2.5},
                        {"field": "label", "regex": r"Before Tax|Pretax", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "NetIncome",
                    "statement": "PL",
                    "period_type": "duration",
                    "candidates": [
                        {"field": "element", "regex": r"ProfitLossAttributableToOwnersOfParentSummaryOfBusinessResults", "weight": 2.5},
                        {"field": "element", "regex": r"ProfitLossSummaryOfBusinessResults|NetIncomeLossSummaryOfBusinessResults", "weight": 2.5},
                        {"field": "element", "regex": r"ProfitLoss|NetIncome|NetProfit", "weight": 2.0},
                        {"field": "label", "regex": r"Net (Income|Profit)", "weight": 1.0},
                    ],
                },
                # BS
                {
                    "canonical_key": "TotalAssets",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "exact": "TotalAssetsIFRS", "weight": 3.5},
                        {"field": "element", "exact": "TotalAssetsIFRSSummaryOfBusinessResults", "weight": 3.5},
                        {"field": "element", "exact": "TotalAssetsSummaryOfBusinessResults", "weight": 3.5},
                        {"field": "element", "regex": r"Assets$|TotalAssets", "weight": 3.0},
                        {"field": "label", "regex": r"Total Assets|Assets,? Total", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "TotalLiabilities",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "exact": "LiabilitiesIFRS", "weight": 3.5},
                        {"field": "element", "exact": "Liabilities", "weight": 3.0},
                        {"field": "element", "regex": r"Liabilities$|TotalLiabilities", "weight": 3.0},
                        {"field": "label", "regex": r"Total Liabilities|Liabilities,? Total", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "TotalEquity",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "exact": "EquityIFRS", "weight": 3.5},
                        {"field": "element", "exact": "Equity", "weight": 3.0},
                        {"field": "element", "exact": "EquityAttributableToOwnersOfParentIFRSSummaryOfBusinessResults", "weight": 2.0},
                        {"field": "element", "regex": r"NetAssetsSummaryOfBusinessResults|NetAssets", "weight": 3.0},
                        {"field": "element", "regex": r"Equity$|TotalEquity", "weight": 2.5},
                        {"field": "label", "regex": r"Total Equity|Equity,? Total", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "CashAndCashEquivalents",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"CashAndCashEquivalents|CashAndDeposits|CashAndDueFromBanksAssetsBNK", "weight": 2.5},
                        {"field": "label", "regex": r"Cash and Cash Equivalents", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "AccountsReceivable",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"TradeAndOtherReceivables|AccountsReceivable|NotesAndAccountsReceivable|Receivables(?!RelatedToFinancialServices)(?!FromFinancialServices)", "weight": 2.0},
                        {"field": "label", "regex": r"Receivables|Trade and Other Receivables", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "FinancialServicesReceivablesCurrent",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"ReceivablesRelatedToFinancialServicesCAIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"ReceivablesFromFinancialServicesCAIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"ReceivablesRelatedToFinancialServicesCurrent", "weight": 2.0},
                        {"field": "label", "regex": r"Receivables Related to Financial Services", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "FinancialServicesReceivablesNonCurrent",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"ReceivablesRelatedToFinancialServicesNCAIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"ReceivablesFromFinancialServicesNCAIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"ReceivablesRelatedToFinancialServicesNoncurrent", "weight": 2.0},
                        {"field": "label", "regex": r"Receivables Related to Financial Services", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "Inventories",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"Inventories", "weight": 2.0},
                        {"field": "label", "regex": r"Inventories", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "PropertyPlantAndEquipment",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "exact": "PropertyPlantAndEquipmentIFRS", "weight": 3.0},
                        {"field": "element", "exact": "PropertyPlantAndEquipment", "weight": 3.0},
                        {"field": "element", "regex": r"PropertyPlantAndEquipment(?!AcquisitionCost|AccumulatedDepreciation|Accumulated)", "weight": 2.0},
                        {"field": "label", "regex": r"Property,? Plant and Equipment|PPE", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "IntangibleAssets",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "exact": "IntangibleAssetsIFRS", "weight": 3.0},
                        {"field": "element", "regex": r"IntangibleAssets(?!AcquisitionCost|Accumulated)", "weight": 2.0},
                        {"field": "label", "regex": r"Intangible Assets", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "Goodwill",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"Goodwill", "weight": 2.0},
                        {"field": "label", "regex": r"Goodwill", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "Investments",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"InvestmentsAccountedForUsingEquityMethodIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"CarryingAmountShares.*Investment", "weight": 2.2},
                        {"field": "element", "regex": r"Investments|InvestmentSecurities|AvailableForSaleSecurities|TradingSecurities", "weight": 2.0},
                        {"field": "label", "regex": r"Investments|Investment Securities|Available[- ]for[- ]sale", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OtherFinancialAssets",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"OtherFinancialAssets(?!.*CAIFRS)(?!.*NCAIFRS)", "weight": 2.0},
                        {"field": "label", "regex": r"Other Financial Assets", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OtherFinancialAssetsCurrent",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"OtherFinancialAssetsCAIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"OtherFinancialAssetsCurrent", "weight": 2.0},
                        {"field": "label", "regex": r"Other Financial Assets", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OtherFinancialAssetsNonCurrent",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"OtherFinancialAssetsNCAIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"OtherFinancialAssetsNoncurrent", "weight": 2.0},
                        {"field": "label", "regex": r"Other Financial Assets", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "AssetsHeldForSale",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"^AssetsHeldForSale", "weight": 2.5},
                        {"field": "label", "regex": r"Assets Held for Sale", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "DeferredTaxAssets",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"DeferredTaxAssets|DeferredTaxAsset", "weight": 2.0},
                        {"field": "label", "regex": r"Deferred Tax Assets", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "RightOfUseAssets",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"RightOfUseAssets", "weight": 2.0},
                        {"field": "element", "regex": r"EquipmentOnOperatingLeases(?!.*AcquisitionCost)(?!.*Accumulated).*IFRS", "weight": 2.0},
                        {"field": "label", "regex": r"Right[- ]of[- ]Use Assets", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "CurrentAssets",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"(?<!Non)(?<!Other)CurrentAssets|AssetsCurrent", "weight": 2.5},
                        {"field": "label", "regex": r"\\bCurrent Assets\\b", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OtherCurrentAssets",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"OtherCurrentAssets", "weight": 2.0},
                        {"field": "label", "regex": r"Other Current Assets", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "NonCurrentAssets",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"(?<!Other)NonCurrentAssets|(?<!Other)NoncurrentAssets|AssetsNoncurrent", "weight": 2.5},
                        {"field": "label", "regex": r"\\bNon-?Current Assets\\b|\\bNoncurrent Assets\\b", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OtherNonCurrentAssets",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"OtherNonCurrentAssets|OtherNoncurrentAssets", "weight": 2.0},
                        {"field": "label", "regex": r"Other Non-?Current Assets|Other Noncurrent Assets", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OtherAssets",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"OtherAssets.*BNK|OtherAssets", "weight": 2.0},
                        {"field": "label", "regex": r"Other Assets", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "CurrentLiabilities",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "exact": "TotalCurrentLiabilitiesIFRS", "weight": 3.0},
                        {"field": "element", "regex": r"(?<!Non)(?<!Other)CurrentLiabilities|LiabilitiesCurrent", "weight": 2.5},
                        {"field": "label", "regex": r"\\bCurrent Liabilities\\b", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OtherCurrentLiabilities",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"OtherCurrentLiabilities", "weight": 2.0},
                        {"field": "label", "regex": r"Other Current Liabilities", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "NonCurrentLiabilities",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "exact": "NonCurrentLiabilitiesIFRS", "weight": 3.0},
                        {"field": "element", "exact": "TotalNonCurrentLiabilitiesIFRS", "weight": 3.0},
                        {"field": "element", "regex": r"(?<!Other)NonCurrentLiabilities|(?<!Other)NoncurrentLiabilities|LiabilitiesNoncurrent", "weight": 2.5},
                        {"field": "label", "regex": r"\\bNon-?Current Liabilities\\b|\\bNoncurrent Liabilities\\b", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OtherNonCurrentLiabilities",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"OtherNonCurrentLiabilities|OtherNoncurrentLiabilities", "weight": 2.0},
                        {"field": "label", "regex": r"Other Non-?Current Liabilities|Other Noncurrent Liabilities", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "LiabilitiesHeldForSale",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"LiabilitiesDirectlyAssociatedWithAssetsHeldForSaleIFRS", "weight": 2.5},
                        {"field": "label", "regex": r"Liabilities Directly Associated With Assets Held for Sale", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OtherLiabilities",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"OtherLiabilities.*BNK|OtherLiabilities", "weight": 2.0},
                        {"field": "label", "regex": r"Other Liabilities", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "AccountsPayable",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"AccountsPayable|TradeAndOtherPayables|NotesAndAccountsPayable|Payables", "weight": 2.0},
                        {"field": "label", "regex": r"Payables|Trade and Other Payables", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "ShortTermBorrowings",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"InterestBearingLiabilitiesCLIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"FinancingLiabilitiesCLIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"BorrowingsCLIFRS", "weight": 2.2},
                        {"field": "element", "regex": r"CurrentPortionOfLongTermBorrowingsCLIFRS", "weight": 2.2},
                        {"field": "element", "regex": r"ShortTermBorrowings|ShortTermDebt|CurrentPortionOfLongTermDebt", "weight": 2.0},
                        {"field": "label", "regex": r"Short[- ]Term Borrowings|Short[- ]Term Debt", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "LongTermBorrowings",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"InterestBearingLiabilitiesNCLIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"FinancingLiabilitiesNCLIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"BorrowingsNCLIFRS", "weight": 2.2},
                        {"field": "element", "regex": r"LongTermBorrowings|LongTermDebt|NonCurrentBorrowings", "weight": 2.0},
                        {"field": "label", "regex": r"Long[- ]Term Borrowings|Long[- ]Term Debt", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "BondsPayable",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"BondsPayable", "weight": 2.0},
                        {"field": "label", "regex": r"Bonds Payable", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "LeaseLiabilities",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"LeaseLiabilities", "weight": 2.0},
                        {"field": "label", "regex": r"Lease Liabilities", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "Provisions",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"LiabilitiesForQualityAssuranceCLIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"Provisions(?!CLIFRS)(?!NCLIFRS)", "weight": 2.0},
                        {"field": "label", "regex": r"Provisions", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "ProvisionsCurrent",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"ProvisionsCLIFRS", "weight": 2.5},
                        {"field": "label", "regex": r"Provisions", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "ProvisionsNonCurrent",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"ProvisionsNCLIFRS", "weight": 2.5},
                        {"field": "label", "regex": r"Provisions", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "AccruedExpenses",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"AccruedExpensesCLIFRS|AccruedExpenses", "weight": 2.5},
                        {"field": "label", "regex": r"Accrued Expenses", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "IncomeTaxesPayable",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"IncomeTaxesPayableCLIFRS|IncomeTaxesPayable", "weight": 2.5},
                        {"field": "label", "regex": r"Income Taxes Payable", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "RetirementBenefitLiabilities",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"RetirementBenefitLiabilityNCLIFRS|RetirementBenefitLiability", "weight": 2.5},
                        {"field": "label", "regex": r"Retirement Benefit", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "RetirementBenefitLiabilitiesCurrent",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"EmployeeBenefitsAccrualsCLIFRS", "weight": 2.5},
                        {"field": "label", "regex": r"Employee Benefits|Retirement Benefit", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "RetirementBenefitLiabilitiesNonCurrent",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"EmployeeBenefitsNCLIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"RetirementBenefitLiabilityNCLIFRS", "weight": 2.2},
                        {"field": "label", "regex": r"Employee Benefits|Retirement Benefit", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "DeferredTaxLiabilities",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"DeferredTaxLiabilities|DeferredTaxLiability", "weight": 2.0},
                        {"field": "label", "regex": r"Deferred Tax Liabilities", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OtherFinancialLiabilities",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"OtherFinancialLiabilities(?!.*CLIFRS)(?!.*NCLIFRS)", "weight": 2.0},
                        {"field": "label", "regex": r"Other Financial Liabilities", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OtherFinancialLiabilitiesCurrent",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"OtherFinancialLiabilitiesCLIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"OtherFinancialLiabilitiesCurrent", "weight": 2.0},
                        {"field": "label", "regex": r"Other Financial Liabilities", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OtherFinancialLiabilitiesNonCurrent",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"OtherFinancialLiabilitiesNCLIFRS", "weight": 2.5},
                        {"field": "element", "regex": r"OtherFinancialLiabilitiesNoncurrent", "weight": 2.0},
                        {"field": "label", "regex": r"Other Financial Liabilities", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "ShareCapital",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"ShareCapital|CapitalStock", "weight": 2.0},
                        {"field": "label", "regex": r"Share Capital|Capital Stock", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "CapitalSurplus",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"CapitalSurplus|AdditionalPaidInCapital", "weight": 2.0},
                        {"field": "label", "regex": r"Capital Surplus|Additional Paid[- ]in Capital", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "RetainedEarnings",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"RetainedEarnings|RetainedEarningsAccumulatedDeficit", "weight": 2.0},
                        {"field": "label", "regex": r"Retained Earnings", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "OtherComponentsOfEquity",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"OtherComponentsOfEquityIFRS|OtherComponentsOfEquity", "weight": 2.5},
                        {"field": "label", "regex": r"Other Components of Equity", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "TreasuryShares",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"TreasurySharesIFRS|TreasuryStock|TreasuryShares", "weight": 2.5},
                        {"field": "label", "regex": r"Treasury", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "AccumulatedOtherComprehensiveIncome",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"AccumulatedOtherComprehensiveIncome|AOCI", "weight": 2.0},
                        {"field": "label", "regex": r"Accumulated Other Comprehensive Income|AOCI", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "NonControllingInterests",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"NoncontrollingInterests|NonControllingInterests", "weight": 2.0},
                        {"field": "label", "regex": r"Non[- ]controlling Interests", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "Securities",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"SecuritiesAssetsBNK|Securities|InvestmentSecurities|AvailableForSaleSecurities|TradingSecurities", "weight": 2.0},
                        {"field": "label", "regex": r"Securities|Investment Securities|Available[- ]for[- ]sale", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "Borrowings",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"BorrowedMoneyLiabilitiesBNK|Borrowings|ShortTermBorrowings|LongTermBorrowings", "weight": 2.0},
                        {"field": "label", "regex": r"Borrowings|Short[- ]term Borrowings|Long[- ]term Borrowings", "weight": 1.0},
                    ],
                },
                # CF
                {
                    "canonical_key": "OperatingCashFlow",
                    "statement": "CF",
                    "period_type": "duration",
                    "candidates": [
                        {"field": "element", "regex": r"NetCashProvidedByUsedInOperatingActivities|OperatingActivities", "weight": 2.5},
                        {"field": "label", "regex": r"Operating Activities", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "InvestingCashFlow",
                    "statement": "CF",
                    "period_type": "duration",
                    "candidates": [
                        {"field": "element", "regex": r"NetCashProvidedByUsedInInvestingActivities|InvestingActivities", "weight": 2.5},
                        {"field": "label", "regex": r"Investing Activities", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "FinancingCashFlow",
                    "statement": "CF",
                    "period_type": "duration",
                    "candidates": [
                        {"field": "element", "regex": r"NetCashProvidedByUsedInFinancingActivities|FinancingActivities", "weight": 2.5},
                        {"field": "label", "regex": r"Financing Activities", "weight": 1.0},
                    ],
                },
                # Banking-friendly optional items
                {
                    "canonical_key": "NetInterestIncome",
                    "statement": "PL",
                    "period_type": "duration",
                    "candidates": [
                        {"field": "element", "regex": r"NetInterestIncome|NetInterestRevenue", "weight": 2.0},
                        {"field": "label", "regex": r"Net Interest", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "Loans",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"LoansAndBillsDiscountedAssetsBNK|LoansAndAdvances|Loans", "weight": 2.0},
                        {"field": "label", "regex": r"Loans", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "Deposits",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"Deposits", "weight": 2.0},
                        {"field": "label", "regex": r"Deposits", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "TradingAssets",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"TradingAssetsAssetsBNK|TradingAssets", "weight": 2.0},
                        {"field": "label", "regex": r"Trading Assets", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "ReceivablesUnderResaleAgreements",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"ReceivablesUnderResaleAgreementsAssetsBNK", "weight": 2.0},
                        {"field": "label", "regex": r"Receivables Under Resale Agreements", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "CustomersLiabilitiesForAcceptancesAndGuarantees",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"CustomersLiabilitiesForAcceptancesAndGuaranteesAssetsBNK", "weight": 2.0},
                        {"field": "label", "regex": r"Customers Liabilities for Acceptances and Guarantees", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "MonetaryClaimsBought",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"MonetaryClaimsBoughtAssetsBNK", "weight": 2.0},
                        {"field": "label", "regex": r"Monetary Claims Bought", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "ReceivablesUnderSecuritiesBorrowingTransactions",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"ReceivablesUnderSecuritiesBorrowingTransactionsAssetsBNK", "weight": 2.0},
                        {"field": "label", "regex": r"Receivables Under Securities Borrowing Transactions", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "RepoLiabilities",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"PayablesUnderRepurchaseAgreementsLiabilitiesBNK", "weight": 2.0},
                        {"field": "label", "regex": r"Payables Under Repurchase Agreements", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "TradingLiabilities",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"TradingLiabilitiesLiabilitiesBNK|TradingLiabilities", "weight": 2.0},
                        {"field": "label", "regex": r"Trading Liabilities", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "NegotiableCertificatesOfDeposit",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"NegotiableCertificatesOfDepositLiabilitiesBNK", "weight": 2.0},
                        {"field": "label", "regex": r"Negotiable Certificates of Deposit", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "AcceptancesAndGuarantees",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"AcceptancesAndGuaranteesLiabilitiesBNK", "weight": 2.0},
                        {"field": "label", "regex": r"Acceptances and Guarantees", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "CallMoneyAndBillsSold",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"CallMoneyAndBillsSoldLiabilitiesBNK", "weight": 2.0},
                        {"field": "label", "regex": r"Call Money and Bills Sold", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "TrustAccountBorrowings",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"BorrowedMoneyFromTrustAccountLiabilitiesBNK", "weight": 2.0},
                        {"field": "label", "regex": r"Borrowed Money from Trust Accounts", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "CommercialPapers",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"CommercialPapersLiabilities", "weight": 2.0},
                        {"field": "label", "regex": r"Commercial Papers", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "ForeignExchangeLiabilities",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"ForeignExchangesLiabilitiesBNK", "weight": 2.0},
                        {"field": "label", "regex": r"Foreign Exchange Liabilities", "weight": 1.0},
                    ],
                },
                {
                    "canonical_key": "SecuritiesLendingPayables",
                    "statement": "BS",
                    "period_type": "instant",
                    "candidates": [
                        {"field": "element", "regex": r"PayablesUnderSecuritiesLendingTransactionsLiabilitiesBNK", "weight": 2.0},
                        {"field": "label", "regex": r"Payables Under Securities Lending Transactions", "weight": 1.0},
                    ],
                },
            ],
        }
        return MappingConfig.from_dict(data).to_dict()

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "MappingConfig":
        version = int(data.get("version", 1) or 1)
        items = [MappingItem.from_dict(item) for item in data.get("items", []) if isinstance(item, dict)]
        return cls(version=version, items=items)

    def __init__(self, version: int = 1, items: Optional[List[MappingItem]] = None) -> None:
        self.version = int(version)
        self.items = items or []

    def to_dict(self) -> Dict[str, Any]:
        return {"version": self.version, "items": [item.to_dict() for item in self.items]}

    def merge_over(self, overlay: Optional["MappingConfig"]) -> "MappingConfig":
        if overlay is None:
            return MappingConfig(version=self.version, items=list(self.items))
        base_items = {item.canonical_key: item for item in self.items}
        for item in overlay.items:
            if item.canonical_key:
                base_items[item.canonical_key] = item
        merged_items = list(base_items.values())
        return MappingConfig(version=self.version, items=merged_items)

    @staticmethod
    def load_json(path: Path) -> Dict[str, Any]:
        """Load a JSON mapping file."""
        with path.open("r", encoding="utf-8") as f:
            return json.load(f)

    @staticmethod
    def merge(base: Dict[str, Any], overlay: Optional[Dict[str, Any]]) -> Dict[str, Any]:
        """Merge mapping dictionaries with overlay items winning by canonical_key."""
        base_config = MappingConfig.from_dict(base)
        overlay_config = MappingConfig.from_dict(overlay) if overlay else None
        return base_config.merge_over(overlay_config).to_dict()
