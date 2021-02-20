#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Screening definition
"""

SCREENING = True
ENABLE_DIVIDEND_YIELD = True
TARGET_DIVIDEND_YIELD = 3.75

ENABLE_PAYOUT_RATIO = True
TARGET_PAYOUT_RATIO = 50

ENABLE_PBR = True
TARGET_PBR = 1.5

ENABLE_CAPTIAL_ADEQUACY_RATIO = True
TARGET_CAPTIAL_ADEQUACY_RATIO = 50

ENABLE_OPERATING_PROFIT_RATIO = True
TARGET_OPERATING_PROFIT_RATIO = 10

ENABLE_CONTINUOUS_DIVIDEND_INCREASE = True
TARGET_CONTINUOUS_DIVIDEND_INCREASE = 3

"""
# Purpose
Confirm that it matches the following conditions.
"""

def util_screening(dividend_yield,dividend_payout_ratio,pbr,capital_adequacy_ratio,operating_profit_ratio,continuous_dividend_increase):
    if ENABLE_DIVIDEND_YIELD:
        # Dividend_yield has a high priority and does not allow None.
        if dividend_yield is not None and dividend_yield >= TARGET_DIVIDEND_YIELD: 
            pass
        else:
            return False
    if ENABLE_PAYOUT_RATIO:
        # Since dividend_payout_ratio has a low priority, allow None.
        if dividend_payout_ratio is None or dividend_payout_ratio <= TARGET_PAYOUT_RATIO:
            pass
        else:
            return False
    if ENABLE_PBR:
        # Since pbr has a low priority, allow None.
        if pbr is None or pbr <= TARGET_PBR:
            pass
        else:
            return False
    if ENABLE_CAPTIAL_ADEQUACY_RATIO:
        if capital_adequacy_ratio is None or capital_adequacy_ratio >= TARGET_CAPTIAL_ADEQUACY_RATIO:
            pass
        else:
            return False
    if ENABLE_OPERATING_PROFIT_RATIO:
        # Since operating_profit_ratio has a low priority, allow None.
        if operating_profit_ratio is None or operating_profit_ratio >= TARGET_OPERATING_PROFIT_RATIO:
            pass
        else:
            return False
    if ENABLE_CONTINUOUS_DIVIDEND_INCREASE:
        # Since continuous_dividend_increase has a high priority, does not allow None.
        if continuous_dividend_increase is not None and continuous_dividend_increase >= TARGET_CONTINUOUS_DIVIDEND_INCREASE:
            pass
        else:
            return False
    return True