"""
Unit tests for model training and inference modules.
"""

import pytest
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from sklearn.pipeline import Pipeline

from src.models.train import load_data, FEATURE_COLS, cross_val_evaluate, MODELS
from src.models.inference import HeartDiseasePredictor, get_sample_input

BASE_DIR = Path(__file__).resolve().parents[1]
DATA_PATH = BASE_DIR / "data" / "heart.csv"
MODEL_PATH = BASE_DIR / "src" / "models" / "heart_model.pkl"


class TestTraining:
    """Tests for training utilities."""

    def test_models_dict_has_entries(self):
        assert len(MODELS) >= 2

    def test_each_model_has_pipeline_and_grid(self):
        for name, spec in MODELS.items():
            assert "pipeline" in spec, f"{name} missing pipeline"
            assert "param_grid" in spec, f"{name} missing param_grid"
            assert isinstance(spec["pipeline"], Pipeline)

    def test_cross_val_evaluate_returns_metrics(self):
        df = load_data(str(DATA_PATH))
        X = df[FEATURE_COLS]
        y = df['target']
        from sklearn.model_selection import StratifiedKFold
        cv = StratifiedKFold(n_splits=3, shuffle=True, random_state=42)
        pipe = MODELS["Logistic Regression"]["pipeline"]
        summary = cross_val_evaluate(pipe, X, y, cv)
        for metric in ["accuracy", "precision", "recall", "roc_auc"]:
            assert metric in summary
            mean, std = summary[metric]
            assert 0 <= mean <= 1
            assert std >= 0


class TestInference:
    """Tests for the inference module."""

    @pytest.fixture(autouse=True)
    def _check_model_exists(self):
        if not MODEL_PATH.exists():
            pytest.skip("Trained model not found; run train.py first")

    def test_predictor_loads(self):
        predictor = HeartDiseasePredictor(str(MODEL_PATH))
        assert predictor.pipeline is not None

    def test_predict_returns_dict(self):
        predictor = HeartDiseasePredictor(str(MODEL_PATH))
        result = predictor.predict(get_sample_input())
        assert isinstance(result, dict)

    def test_predict_has_required_keys(self):
        predictor = HeartDiseasePredictor(str(MODEL_PATH))
        result = predictor.predict(get_sample_input())
        for key in ["prediction", "prediction_label", "confidence", "risk_level"]:
            assert key in result

    def test_prediction_is_binary(self):
        predictor = HeartDiseasePredictor(str(MODEL_PATH))
        result = predictor.predict(get_sample_input())
        assert result["prediction"] in [0, 1]

    def test_confidence_in_range(self):
        predictor = HeartDiseasePredictor(str(MODEL_PATH))
        result = predictor.predict(get_sample_input())
        assert 0.0 <= result["confidence"] <= 1.0

    def test_risk_level_valid(self):
        predictor = HeartDiseasePredictor(str(MODEL_PATH))
        result = predictor.predict(get_sample_input())
        assert result["risk_level"] in ["Low Risk", "Medium Risk", "High Risk"]

    def test_predict_batch(self):
        predictor = HeartDiseasePredictor(str(MODEL_PATH))
        results = predictor.predict_batch([get_sample_input(), get_sample_input()])
        assert len(results) == 2

    def test_sample_input_has_all_features(self):
        sample = get_sample_input()
        for feat in HeartDiseasePredictor.FEATURE_NAMES:
            assert feat in sample
