"""
Unit tests for data preprocessing module.
"""

import pytest
import pandas as pd
import numpy as np
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.data.preprocess import load_data, FEATURE_COLS, TARGET_COL

DATA_PATH = Path(__file__).resolve().parents[1] / "data" / "heart.csv"


class TestLoadData:
    """Tests for the load_data function."""

    def test_returns_dataframe(self):
        df = load_data(str(DATA_PATH))
        assert isinstance(df, pd.DataFrame)

    def test_no_missing_values(self):
        df = load_data(str(DATA_PATH))
        assert df.isnull().sum().sum() == 0

    def test_expected_columns_present(self):
        df = load_data(str(DATA_PATH))
        for col in FEATURE_COLS + [TARGET_COL]:
            assert col in df.columns, f"Missing column: {col}"

    def test_target_is_binary(self):
        df = load_data(str(DATA_PATH))
        unique_vals = set(df[TARGET_COL].unique())
        assert unique_vals.issubset({0, 1, 0.0, 1.0})

    def test_non_empty(self):
        df = load_data(str(DATA_PATH))
        assert len(df) > 0

    def test_feature_count(self):
        load_data(str(DATA_PATH))
        assert len(FEATURE_COLS) == 13

    def test_invalid_path_raises(self):
        with pytest.raises(Exception):
            load_data("nonexistent_file.csv")

    def test_handles_missing_values(self, tmp_path):
        """Create a CSV with missing values and verify they get filled."""
        csv_path = tmp_path / "test_missing.csv"
        data = {col: [1.0, np.nan, 3.0] for col in FEATURE_COLS}
        data[TARGET_COL] = [0, 1, 0]
        pd.DataFrame(data).to_csv(csv_path, index=False)

        df = load_data(str(csv_path))
        assert df.isnull().sum().sum() == 0
