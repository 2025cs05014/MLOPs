"""
train.py

Model Selection & Training for Heart Disease Dataset
=====================================================
Trains two classifiers:
  1. Logistic Regression  — a linear, interpretable baseline
  2. Random Forest        — an ensemble method for capturing non-linear patterns

Tuning:
  - GridSearchCV (5-fold stratified CV) is used to find optimal hyperparameters
    for each model independently.

Evaluation:
  - 5-fold cross-validation on the training set (accuracy, precision, recall, ROC-AUC)
  - Hold-out test set metrics + classification report
  - The model with the higher mean CV ROC-AUC is selected and saved.

Usage:
    python train.py
"""

import joblib
import mlflow
import mlflow.sklearn
import pandas as pd
from pathlib import Path

from sklearn.ensemble import RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (classification_report, roc_auc_score,
                             make_scorer, precision_score, recall_score,
                             confusion_matrix, roc_curve)
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from sklearn.model_selection import (GridSearchCV, StratifiedKFold,
                                     cross_validate, train_test_split)
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parents[2]   # Assignment/
DATA_PATH = BASE_DIR / "data" / "heart.csv"
MODEL_PATH = BASE_DIR / "src" / "models" / "heart_model.pkl"

FEATURE_COLS = ['age', 'sex', 'cp', 'trestbps', 'chol', 'fbs',
                'restecg', 'thalach', 'exang', 'oldpeak', 'slope', 'ca', 'thal']


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------
def load_data(filepath=None):
    """Load and clean the heart disease dataset."""
    if filepath is None:
        filepath = DATA_PATH
    df = pd.read_csv(filepath)
    for col in df.columns:
        if df[col].isnull().sum() > 0:
            if df[col].dtype in ['float64', 'int64']:
                df[col] = df[col].fillna(df[col].median())
            else:
                df[col] = df[col].fillna(df[col].mode()[0])
    return df


# ---------------------------------------------------------------------------
# Model definitions with hyperparameter grids
# ---------------------------------------------------------------------------
MODELS = {
    "Logistic Regression": {
        "pipeline": Pipeline([
            ("scaler", StandardScaler()),
            ("classifier", LogisticRegression(max_iter=1000, random_state=42))
        ]),
        "param_grid": {
            "classifier__C": [0.01, 0.1, 1, 10],
            "classifier__solver": ["lbfgs", "liblinear"],
            "classifier__penalty": ["l2"],
        },
    },
    "Random Forest": {
        "pipeline": Pipeline([
            ("scaler", StandardScaler()),
            ("classifier", RandomForestClassifier(random_state=42))
        ]),
        "param_grid": {
            "classifier__n_estimators": [100, 200],
            "classifier__max_depth": [5, 10, None],
            "classifier__min_samples_split": [2, 5],
        },
    },
}


# ---------------------------------------------------------------------------
# Evaluation helpers
# ---------------------------------------------------------------------------
def cross_val_evaluate(pipeline, X, y, cv):
    """Run cross-validation and return mean ± std for key metrics."""
    scoring = {
        "accuracy": "accuracy",
        "precision": make_scorer(precision_score, zero_division=0),
        "recall": make_scorer(recall_score, zero_division=0),
        "roc_auc": "roc_auc",
    }
    results = cross_validate(pipeline, X, y, cv=cv, scoring=scoring)
    summary = {
        metric: (results[f"test_{metric}"].mean(),
                 results[f"test_{metric}"].std())
        for metric in scoring
    }
    return summary


def print_cv_summary(name, summary):
    print(f"\n  Cross-Validation Results ({name})")
    print(f"  {'Metric':<12} {'Mean':>8}  {'Std':>8}")
    print(f"  {'-'*32}")
    for metric, (mean, std) in summary.items():
        print(f"  {metric:<12} {mean:>8.4f}  {std:>8.4f}")


def print_test_metrics(name, y_test, y_pred, y_prob):
    auc = roc_auc_score(y_test, y_prob)
    print(f"\n  Hold-out Test Metrics ({name})")
    print(f"  ROC-AUC : {auc:.4f}")
    print(classification_report(y_test, y_pred,
                                target_names=["No Disease", "Disease"]))
    # indent=2))


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    mlflow.set_tracking_uri(str(BASE_DIR / "mlruns"))
    mlflow.set_experiment("heart-disease-classification")

    df = load_data()
    X = df[FEATURE_COLS]
    y = df['target']

    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42, stratify=y
    )

    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)

    best_name = None
    best_pipeline = None
    best_auc = -1.0
    results_summary = {}

    # -----------------------------------------------------------------------
    # Step 1 – Hyperparameter tuning + cross-validation for each model
    # -----------------------------------------------------------------------
    print("=" * 60)
    print("  MODEL SELECTION & TUNING")
    print("=" * 60)

    for name, spec in MODELS.items():
        print(f"\n[{name}]")
        print("  Tuning hyperparameters via GridSearchCV (5-fold CV) ...")

        with mlflow.start_run(run_name=name):
            grid = GridSearchCV(
                spec["pipeline"],
                spec["param_grid"],
                cv=cv,
                scoring="roc_auc",
                n_jobs=-1,
                refit=True,
            )
            grid.fit(X_train, y_train)

            best_params = {
                k.replace("classifier__", ""): v
                for k, v in grid.best_params_.items()
            }
            print(f"  Best params : {best_params}")
            print(f"  Best CV AUC : {grid.best_score_:.4f}")

            # Evaluate best estimator with full metric suite
            tuned_pipeline = grid.best_estimator_
            cv_summary = cross_val_evaluate(tuned_pipeline, X_train, y_train, cv)
            print_cv_summary(name, cv_summary)

            # --- MLflow: log params, metrics, and model ---
            mlflow.log_params(best_params)
            mlflow.log_param("model_type", name)
            for metric, (mean, std) in cv_summary.items():
                mlflow.log_metric(f"cv_{metric}_mean", mean)
                mlflow.log_metric(f"cv_{metric}_std", std)

            # Test-set metrics
            y_pred = tuned_pipeline.predict(X_test)
            y_prob = tuned_pipeline.predict_proba(X_test)[:, 1]
            test_auc = roc_auc_score(y_test, y_prob)
            test_precision = precision_score(y_test, y_pred, zero_division=0)
            test_recall = recall_score(y_test, y_pred, zero_division=0)
            mlflow.log_metric("test_roc_auc", test_auc)
            mlflow.log_metric("test_precision", test_precision)
            mlflow.log_metric("test_recall", test_recall)

            mlflow.sklearn.log_model(tuned_pipeline, artifact_path="model")

            # --- MLflow: log confusion matrix plot ---
            cm = confusion_matrix(y_test, y_pred)
            fig_cm, ax_cm = plt.subplots(figsize=(5, 4))
            ax_cm.imshow(cm, cmap='Blues')
            for i in range(2):
                for j in range(2):
                    ax_cm.text(j, i, str(cm[i, j]), ha='center', va='center',
                               color='white' if cm[i, j] > cm.max() / 2 else 'black', fontsize=14)
            ax_cm.set_xlabel('Predicted')
            ax_cm.set_ylabel('Actual')
            ax_cm.set_title(f'Confusion Matrix — {name}')
            ax_cm.set_xticks([0, 1])
            ax_cm.set_yticks([0, 1])
            ax_cm.set_xticklabels(['No Disease', 'Disease'])
            ax_cm.set_yticklabels(['No Disease', 'Disease'])
            fig_cm.tight_layout()
            cm_path = str(BASE_DIR / f"screenshots/cm_{name.replace(' ', '_').lower()}.png")
            fig_cm.savefig(cm_path, dpi=100)
            mlflow.log_artifact(cm_path)
            plt.close(fig_cm)

            # --- MLflow: log ROC curve plot ---
            fpr, tpr, _ = roc_curve(y_test, y_prob)
            fig_roc, ax_roc = plt.subplots(figsize=(5, 4))
            ax_roc.plot(fpr, tpr, label=f'{name} (AUC={test_auc:.3f})')
            ax_roc.plot([0, 1], [0, 1], 'k--', alpha=0.5)
            ax_roc.set_xlabel('False Positive Rate')
            ax_roc.set_ylabel('True Positive Rate')
            ax_roc.set_title(f'ROC Curve — {name}')
            ax_roc.legend(loc='lower right')
            fig_roc.tight_layout()
            roc_path = str(BASE_DIR / f"screenshots/roc_{name.replace(' ', '_').lower()}.png")
            fig_roc.savefig(roc_path, dpi=100)
            mlflow.log_artifact(roc_path)
            plt.close(fig_roc)

            results_summary[name] = {
                "pipeline": tuned_pipeline,
                "cv_summary": cv_summary,
                "best_params": best_params,
                "cv_auc": cv_summary["roc_auc"][0],
            }

            if cv_summary["roc_auc"][0] > best_auc:
                best_auc = cv_summary["roc_auc"][0]
                best_name = name
                best_pipeline = tuned_pipeline

    # -----------------------------------------------------------------------
    # Step 2 – Final evaluation on held-out test set
    # -----------------------------------------------------------------------
    print("\n" + "=" * 60)
    print("  FINAL TEST-SET EVALUATION")
    print("=" * 60)

    for name, info in results_summary.items():
        pipe = info["pipeline"]
        y_pred = pipe.predict(X_test)
        y_prob = pipe.predict_proba(X_test)[:, 1]
        print_test_metrics(name, y_test, y_pred, y_prob)

    # -----------------------------------------------------------------------
    # Step 3 – Model selection decision
    # -----------------------------------------------------------------------
    print("\n" + "=" * 60)
    print("  MODEL SELECTION DECISION")
    print("=" * 60)
    for name, info in results_summary.items():
        mean_auc, std_auc = info["cv_summary"]["roc_auc"]
        marker = "  <-- SELECTED" if name == best_name else ""
        print(f"  {name:<22}  CV ROC-AUC = {mean_auc:.4f} ± {std_auc:.4f}{marker}")

    print(f"\n  Winner : {best_name}  (highest mean CV ROC-AUC = {best_auc:.4f})")
    print("  Rationale: model chosen purely on generalisation performance "
          "measured by stratified 5-fold cross-validation ROC-AUC.\n")

    # -----------------------------------------------------------------------
    # Step 4 – Save best model & log as MLflow artifact
    # -----------------------------------------------------------------------
    MODEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    joblib.dump(best_pipeline, MODEL_PATH)
    print(f"  Model saved → {MODEL_PATH}")

    with mlflow.start_run(run_name=f"best-model-{best_name}"):
        mlflow.log_param("selected_model", best_name)
        mlflow.log_metric("best_cv_roc_auc", best_auc)
        mlflow.sklearn.log_model(best_pipeline, artifact_path="best_model")
        mlflow.log_artifact(str(MODEL_PATH))

    print(f"  MLflow tracking URI → {BASE_DIR / 'mlruns'}")
    print("=" * 60)


if __name__ == "__main__":
    main()
