# -*- coding: utf-8 -*-
"""Regression tests for the human-facing CLI value parsers."""
import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from main import parse_budget_input, parse_duration_input


@pytest.mark.parametrize("raw, expected", [
    ("365", 365),
    ("365 days", 365),
    ("12 months", 360),
    ("1.5 months", 45),
    ("1,000 days", 1000),
])
def test_parse_duration_input(raw, expected):
    assert parse_duration_input(raw) == expected


@pytest.mark.parametrize("raw, expected", [
    ("1000", 1000.0),
    ("₹ 1,000 Cr", 1000.0),
    ("150 lakh", 1.5),
    ("2.5 crore", 2.5),
])
def test_parse_budget_input(raw, expected):
    assert parse_budget_input(raw) == pytest.approx(expected)


def test_invalid_budget_raises():
    with pytest.raises(ValueError):
        parse_budget_input("not a number")
