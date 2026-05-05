# Heart Disease Prediction — MLOps Assignment Report

**Student ID:** Ambika Prasad Tripathy(2025cs05014)
**Course:** MLOps — M.Tech Semester 2, BIT Mesra
**GitHub Repository:** [https://github.com/2025cs05014/MLOPs](https://github.com/2025cs05014/MLOPs)

---

## Table of Contents

1. [Introduction](#1-introduction)
2. [Setup & Installation Instructions](#2-setup--installation-instructions)
3. [Data Acquisition & Exploratory Data Analysis](#3-data-acquisition--exploratory-data-analysis)
4. [Feature Engineering & Model Development](#4-feature-engineering--model-development)
5. [Experiment Tracking with MLflow](#5-experiment-tracking-with-mlflow)
6. [Model Packaging & Reproducibility](#6-model-packaging--reproducibility)
7. [CI/CD Pipeline & Automated Testing](#7-cicd-pipeline--automated-testing)
8. [Model Containerization](#8-model-containerization)
9. [Production Deployment](#9-production-deployment)
10. [Monitoring & Logging](#10-monitoring--logging)
11. [Architecture Diagram](#11-architecture-diagram)
12. [Conclusion & Future Work](#12-conclusion--future-work)
13. [References](#13-references)

---

## 1. Introduction

### 1.1 Problem Statement

Heart disease is one of the leading causes of death globally. Early detection and risk assessment can significantly improve patient outcomes. This project builds a complete end-to-end Machine Learning Operations (MLOps) pipeline to predict heart disease risk based on patient health data from the UCI Heart Disease dataset.

The solution encompasses the full ML lifecycle: data acquisition, exploratory analysis, model development, experiment tracking, automated testing, containerization, Kubernetes deployment, and production monitoring — following industry-standard MLOps best practices.

### 1.2 Dataset Description

The Heart Disease UCI dataset contains 303 patient records with 14 features:

- **Demographic:** age, sex
- **Clinical:** chest pain type (cp), resting blood pressure (trestbps), cholesterol (chol), fasting blood sugar (fbs), resting ECG (restecg)
- **Exercise:** maximum heart rate (thalach), exercise-induced angina (exang), ST depression (oldpeak), slope of ST segment (slope)
- **Diagnostic:** number of major vessels (ca), thalassemia (thal)
- **Target:** presence (1) or absence (0) of heart disease

**Data Source:** UCI Machine Learning Repository

### 1.3 Technology Stack

| Component              | Technology                    |
|------------------------|-------------------------------|
| Programming Language   | Python 3.11+                  |
| ML Framework           | Scikit-learn                  |
| API Framework          | FastAPI + Flask               |
| Experiment Tracking    | MLflow                        |
| Testing                | Pytest                        |
| CI/CD                  | GitHub Actions                |
| Containerization       | Docker + Docker Compose       |
| Orchestration          | Kubernetes (Minikube)         |
| Monitoring             | Prometheus + Grafana          |

---

## 2. Setup & Installation Instructions

### 2.1 Prerequisites

- Python 3.10+ (tested on 3.11 and 3.14)
- Docker Desktop (for containerization & Kubernetes)
- minikube + kubectl (for K8s deployment)
- Git

### 2.2 Local Setup

| Step | Command / Action |
|------|-----------------|
| Step 1: Clone the repository | `git clone git@github-bits:2025cs05014/MLOPs.git` <br> `cd Assignment` |
| Step 2: Create and activate virtual environment | `python3 -m venv .venv` <br> `source .venv/bin/activate` |
| Step 3: Install dependencies | `pip install -r requirements.txt` |
| Step 4: Run EDA | `jupyter notebook notebooks/EDA.ipynb` |
| Step 5: Train models with experiment tracking | `python src/models/train.py` |
| Step 6: Run unit tests | `python -m pytest tests/ -v` |
| Step 7: Start API server | `uvicorn src.api.model_app:app --reload` |
| Step 8: View MLflow dashboard | `mlflow ui --backend-store-uri mlruns/` |

---

## 3. Data Acquisition & Exploratory Data Analysis

### 3.1 Data Acquisition

The dataset (`data/heart.csv`) is the UCI Heart Disease dataset containing 303 patient records with 13 input features and a binary target variable.

### 3.2 Data Cleaning

Missing values were handled in the preprocessing pipeline (`src/data/preprocess.py`):
- Numerical columns: imputed with median values
- Categorical columns: imputed with mode values

After cleaning:
- **Total records:** 303
- **Features:** 13 (5 numerical + 8 categorical)
- **Target:** Binary (0 = No Disease, 1 = Disease)
- **Missing values:** 0

### 3.3 Exploratory Data Analysis

EDA was performed in `notebooks/EDA.ipynb`. Key insights:

- **Age Distribution:** Ranges from 29 to 77 years with a mean of ~54 years, slightly right-skewed
- **Gender Distribution:** Majority of patients are male (~68.3% male, ~31.7% female)
- **Chest Pain:** Type 0 is the most common (~47.2% of cases)
- **Blood Pressure:** Mostly between 120–140 mmHg
- **Fasting Blood Sugar:** Most patients below 120 mg/dl (85.1%)
- **Heart Disease Prevalence:** Slightly more than half of patients have heart disease (~54.5%)
- **Key Correlations:** Age, exercise-induced angina, ST depression (oldpeak), and maximum heart rate show strong associations with heart disease

Generated visualizations (saved in `screenshots/`):

| Plot | File |
|------|------|
| Feature Histograms | `screenshots/feature_histograms.png` |
| Box Plots (Outlier Detection) | `screenshots/boxplots.png` |
| Categorical Feature Analysis | `screenshots/categorical_features.png` |
| Correlation Heatmap | `screenshots/correlation_heatmap.png` |
| Class Balance | `screenshots/class_balance.png` |

---

## 4. Feature Engineering & Model Development

### 4.1 Feature Engineering

A preprocessing pipeline was built using Scikit-learn's `StandardScaler` embedded within a `Pipeline` object:

- **All 13 features** are standardized to zero mean and unit variance using `StandardScaler`
- The scaler is fitted only on training data and applied consistently during both training and inference

### 4.2 Data Split

- **Training set:** 80% (242 samples)
- **Test set:** 20% (61 samples)
- **Stratified split** to maintain class balance in both sets
- **Random state:** 42 for reproducibility

### 4.3 Models Trained

Two classification models were trained and compared using `GridSearchCV` with 5-fold stratified cross-validation:

**Model 1: Logistic Regression**
- Hyperparameter grid: C ∈ {0.01, 0.1, 1, 10}, solver ∈ {lbfgs, liblinear}, penalty = l2
- A linear model serving as a strong interpretable baseline

**Model 2: Random Forest Classifier**
- Hyperparameter grid: n_estimators ∈ {100, 200}, max_depth ∈ {5, 10, None}, min_samples_split ∈ {2, 5}
- An ensemble tree-based model capturing non-linear relationships

### 4.4 Model Evaluation Results

Cross-Validation Results (5-Fold Stratified):

| Model               | CV Accuracy | CV Precision | CV Recall | CV ROC-AUC |
|---------------------|-------------|--------------|-----------|------------|
| Logistic Regression | ~0.85       | ~0.84        | ~0.87     | ~0.90      |
| Random Forest       | ~0.83       | ~0.82        | ~0.85     | ~0.88      |

Generated evaluation artifacts:

| Artifact | File |
|----------|------|
| Confusion Matrix — Logistic Regression | `screenshots/cm_logistic_regression.png` |
| Confusion Matrix — Random Forest | `screenshots/cm_random_forest.png` |
| ROC Curve — Logistic Regression | `screenshots/roc_logistic_regression.png` |
| ROC Curve — Random Forest | `screenshots/roc_random_forest.png` |

### 4.5 Best Model Selection

**Logistic Regression** was selected as the best model based on the highest mean CV ROC-AUC score (~0.90). Despite being the simpler model, it outperformed Random Forest on this dataset, likely due to the relatively linear separability of the features after standardization.

**Selection criterion:** Highest mean 5-fold stratified cross-validation ROC-AUC.

---

## 5. Experiment Tracking with MLflow

### 5.1 MLflow Integration

MLflow was integrated into the training pipeline (`src/models/train.py`) to provide comprehensive experiment tracking. All experiments are stored in the `mlruns/` directory.

### 5.2 What is Tracked

For each model run, the following items are logged:

**Parameters:**
- model_type, C, max_iter, solver (for Logistic Regression)
- model_type, n_estimators, max_depth, min_samples_split (for Random Forest)

**Metrics:**
- cv_accuracy_mean, cv_accuracy_std
- cv_precision_mean, cv_precision_std
- cv_recall_mean, cv_recall_std
- cv_roc_auc_mean, cv_roc_auc_std
- test_roc_auc, test_precision, test_recall

**Artifacts:**
- Trained model (MLflow sklearn model format)
- Confusion matrix plot (PNG)
- ROC curve plot (PNG)

### 5.3 MLflow Dashboard

The MLflow UI provides a comprehensive view of all experiments:

```bash
mlflow ui --backend-store-uri mlruns/
# Open http://127.0.0.1:5000
```

### 5.4 Experiment Tracking Summary

The experiment tracking system enables:
- Full reproducibility of any experiment run
- Easy comparison of model performance across runs
- Version control of trained models and their associated metadata
- Audit trail of all training decisions and outcomes

---

## 6. Model Packaging & Reproducibility

### 6.1 Model Artifacts

The final production model is packaged as:
- `src/models/heart_model.pkl` — Trained Logistic Regression pipeline (StandardScaler + classifier)

### 6.2 Inference Pipeline

A clean inference module (`src/models/inference.py`) provides the `HeartDiseasePredictor` class that:
- Loads the trained pipeline from the saved `.pkl` artifact
- Accepts raw patient data as a Python dictionary
- Handles preprocessing (scaling) and prediction in a single call
- Returns prediction (0/1), prediction label, confidence score, and risk level
- Supports both single and batch predictions

### 6.3 Reproducibility

Complete reproducibility is ensured through:
- `requirements.txt` with specified package versions
- Preprocessing pipeline (StandardScaler) saved within the model pipeline
- Random state fixed across all operations (random_state=42)
- Stratified train/test split for consistent evaluation

---

## 7. CI/CD Pipeline & Automated Testing

### 7.1 Unit Tests

24 comprehensive unit tests were written across 3 test files:

**tests/test_preprocess.py (8 tests):**
- Data loading returns DataFrame, not empty, has target column
- No missing values after cleaning, target is binary
- All expected columns present, feature count is 13
- Invalid path raises exception, handles missing values correctly

**tests/test_model.py (11 tests):**
- At least 2 models available, each has pipeline and param_grid
- Cross-validation returns expected metrics (accuracy, precision, recall, roc_auc)
- Predictor loads successfully, predict returns dict with required keys
- Prediction is binary, confidence in [0, 1] range
- Risk level is valid, batch prediction works
- Sample input has all features

**tests/test_api.py (5 tests):**
- Root endpoint returns 200 with message
- Health endpoint returns 200 with "healthy" status
- Predict endpoint returns valid prediction with confidence and risk level
- Missing fields return 422 validation error
- Metrics endpoint is accessible

**Test Results (24/24 passed):**

```
tests/test_api.py::TestAPI::test_root PASSED
tests/test_api.py::TestAPI::test_health PASSED
tests/test_api.py::TestAPI::test_predict_valid PASSED
tests/test_api.py::TestAPI::test_predict_missing_field PASSED
tests/test_api.py::TestAPI::test_metrics_endpoint PASSED
tests/test_model.py::TestTraining::test_models_dict_has_entries PASSED
tests/test_model.py::TestTraining::test_each_model_has_pipeline_and_grid PASSED
tests/test_model.py::TestTraining::test_cross_val_evaluate_returns_metrics PASSED
tests/test_model.py::TestInference::test_predictor_loads PASSED
tests/test_model.py::TestInference::test_predict_returns_dict PASSED
tests/test_model.py::TestInference::test_predict_has_required_keys PASSED
tests/test_model.py::TestInference::test_prediction_is_binary PASSED
tests/test_model.py::TestInference::test_confidence_in_range PASSED
tests/test_model.py::TestInference::test_risk_level_valid PASSED
tests/test_model.py::TestInference::test_predict_batch PASSED
tests/test_model.py::TestInference::test_sample_input_has_all_features PASSED
tests/test_preprocess.py::TestLoadData::test_returns_dataframe PASSED
tests/test_preprocess.py::TestLoadData::test_no_missing_values PASSED
tests/test_preprocess.py::TestLoadData::test_expected_columns_present PASSED
tests/test_preprocess.py::TestLoadData::test_target_is_binary PASSED
tests/test_preprocess.py::TestLoadData::test_non_empty PASSED
tests/test_preprocess.py::TestLoadData::test_feature_count PASSED
tests/test_preprocess.py::TestLoadData::test_invalid_path_raises PASSED
tests/test_preprocess.py::TestLoadData::test_handles_missing_values PASSED

======================== 24 passed, 1 warning in 11.50s ========================
```

### 7.2 GitHub Actions CI/CD Pipeline

A 4-stage automated pipeline (`.github/workflows/ci.yml`) runs on every push/PR to the main branch:

**Stage 1 — Lint:**
- Code quality check using Flake8 (`--max-line-length=120 --ignore=E501,W503,E402`)

**Stage 2 — Test:**
- Run all 24 unit tests with Pytest (`python -m pytest tests/ -v --tb=short`)

**Stage 3 — Train Model:**
- Train models using `src/models/train.py`
- Upload trained model artifact (`heart_model.pkl`)
- Upload MLflow experiment logs

**Stage 4 — Docker Build & Test:**
- Download trained model artifact from Stage 3
- Build Docker image (`heart-disease-api:latest`)
- Start container and smoke-test API endpoints (`/health`, `/predict`)

---

## 8. Model Containerization

### 8.1 Docker Configuration

The application is containerized using Docker with a production-ready Dockerfile:

- **Base image:** `python:3.11-slim` (minimal footprint)
- **Approach:** Install dependencies → Copy code & data → Set environment variables
- **Environment variables:** `MODEL_PATH`, `PYTHONUNBUFFERED`
- **Exposed port:** 8000
- **CMD:** `uvicorn src.api.model_app:app --host 0.0.0.0 --port 8000`

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt
COPY src/ src/
COPY data/ data/
ENV MODEL_PATH=/app/src/models/heart_model.pkl
ENV PYTHONUNBUFFERED=1
EXPOSE 8000
CMD ["uvicorn", "src.api.model_app:app", "--host", "0.0.0.0", "--port", "8000"]
```

### 8.2 API Design (FastAPI)

The prediction API (`src/api/model_app.py`) exposes the following endpoints:

| Endpoint   | Method | Description                              |
|------------|--------|------------------------------------------|
| `/`        | GET    | API information and available endpoints  |
| `/health`  | GET    | Health check with model status           |
| `/predict` | POST   | Heart disease prediction from JSON input |
| `/metrics` | GET    | Prometheus-format monitoring metrics     |
| `/docs`    | GET    | Interactive Swagger UI documentation     |

The `/predict` endpoint:
- Accepts JSON with 13 patient features
- Validates input using Pydantic models with field constraints (ge, le ranges)
- Returns prediction (0/1), label, confidence score, and risk level
- Handles errors gracefully with proper HTTP status codes

**Sample Prediction Response:**

```json
{
    "prediction": 0,
    "prediction_label": "No Heart Disease",
    "confidence": 0.1651,
    "risk_level": "Low Risk",
    "timestamp": "2026-05-03T12:17:18.530247"
}
```

### 8.3 Flask UI

A browser-based Flask UI (`src/api/app.py`) is also provided for interactive predictions via a web form at `http://127.0.0.1:5000`.

---

## 9. Production Deployment

### 9.1 Docker Compose Deployment

The full application stack is deployed using Docker Compose (`docker-compose.yml`) with three services:

| Service            | URL                    | Description               |
|--------------------|------------------------|---------------------------|
| Heart Disease API  | http://localhost:8000   | FastAPI prediction server |
| Prometheus         | http://localhost:9090   | Metrics collection        |
| Grafana            | http://localhost:3000   | Dashboard (admin/admin)   |

```bash
docker compose up -d --build
```

**Verified deployment output:**
- All 3 containers running (heart-disease-api, prometheus, grafana)
- Health check: `{"status": "healthy", "pipeline_loaded": true}`
- Prediction working correctly
- Prometheus target status: `"health": "up"`

### 9.2 Kubernetes Deployment (Minikube)

Production-ready Kubernetes manifests are provided for enterprise deployment:

**deployment.yaml:**
- 2 replicas with rolling update strategy
- Liveness probe: `GET /health` (delay=10s, period=30s)
- Readiness probe: `GET /health` (delay=5s, period=10s)
- Resource requests: 250m CPU, 256Mi memory
- Resource limits: 500m CPU, 512Mi memory

**service.yaml:**
- LoadBalancer service exposing port 80 → container port 8000

**Verified K8s deployment:**
```
NAME                                     READY   STATUS    RESTARTS   AGE
pod/heart-disease-api-7b5bc5b856-4fch6   1/1     Running   0          20s
pod/heart-disease-api-7b5bc5b856-7phtg   1/1     Running   0          20s

NAME                                TYPE           CLUSTER-IP     EXTERNAL-IP   PORT(S)
service/heart-disease-api-service   LoadBalancer   10.105.51.75   <pending>     80:31725/TCP

NAME                                READY   UP-TO-DATE   AVAILABLE   AGE
deployment.apps/heart-disease-api   2/2     2            2           20s
```

---

## 10. Monitoring & Logging

### 10.1 Application Logging

Comprehensive logging is implemented at the API level (`src/api/model_app.py`):
- All requests are logged with timestamp, method, endpoint, and status
- Prediction requests log input data, prediction result, confidence, and latency
- Logs are written to both console (stdout) and file (`api.log`)
- Log format: `%(asctime)s - %(name)s - %(levelname)s - %(message)s`

### 10.2 Prometheus Metrics

The `/metrics` endpoint exposes production-grade metrics:

- **`predictions_total`** — Counter of total predictions, labeled by result (Heart Disease / No Heart Disease)
- **`prediction_latency_seconds`** — Histogram of prediction processing time
- **`http_requests_total`** — Counter of all HTTP requests, labeled by method, endpoint, and status code

Prometheus scrape configuration (`monitoring/prometheus.yml`):
- Scrape interval: 15 seconds
- Target: `heart-disease-api:8000`
- Metrics path: `/metrics`

### 10.3 Grafana Dashboard

A pre-provisioned Grafana dashboard ("Heart Disease API Monitoring") is auto-loaded with panels for:
- Total predictions count
- Predictions rate by result
- Prediction latency (p95)
- HTTP request rate by method/endpoint/status
- Average prediction latency over time

**Grafana access:** http://localhost:3000 (Credentials: admin / admin)

---

## 11. Architecture Diagram

```
┌─────────────┐     ┌──────────────┐     ┌─────────────────┐     ┌──────────────┐
│  Data (CSV)  │────▶│  Preprocess   │────▶│  Train + MLflow  │────▶│  Model (.pkl) │
└─────────────┘     └──────────────┘     └─────────────────┘     └──────┬───────┘
                                                                        │
                    ┌──────────────┐     ┌─────────────────┐           │
                    │   Prometheus  │◀────│  FastAPI + Docker │◀──────────┘
                    │   /metrics    │     │  /predict        │
                    └──────┬───────┘     └────────┬────────┘
                           │                      │
                    ┌──────▼───────┐     ┌────────▼────────┐
                    │   Grafana     │     │  Kubernetes (K8s)│
                    │   Dashboard   │     │  LoadBalancer     │
                    └──────────────┘     └─────────────────┘
```

**Pipeline Flow:**

1. **Data Layer:** `data/heart.csv` → `src/data/preprocess.py` (load, clean, impute)
2. **Training Layer:** `src/models/train.py` (GridSearchCV + MLflow tracking) → `heart_model.pkl`
3. **Inference Layer:** `src/models/inference.py` (HeartDiseasePredictor class)
4. **API Layer:** `src/api/model_app.py` (FastAPI) + `src/api/app.py` (Flask UI)
5. **CI/CD Layer:** `.github/workflows/ci.yml` (Lint → Test → Train → Docker)
6. **Deployment Layer:** Docker Compose (API + Prometheus + Grafana) / Kubernetes (Minikube)
7. **Monitoring Layer:** Prometheus metrics + Grafana dashboards + application logging

---

## 12. Conclusion & Future Work

### 12.1 Summary

This project successfully demonstrates a complete MLOps pipeline for heart disease prediction, covering all aspects of production-grade machine learning:

- **Data Pipeline:** Automated data loading, cleaning, and preprocessing with missing value imputation
- **Model Development:** Trained and compared 2 classification models with GridSearchCV and 5-fold stratified cross-validation, achieving ~0.90 ROC-AUC with Logistic Regression
- **Experiment Tracking:** Full MLflow integration with parameter, metric, and artifact logging across all experimental runs
- **Testing:** 24 comprehensive unit tests covering data processing, model training, inference, and API functionality
- **CI/CD:** Automated 4-stage GitHub Actions pipeline (Lint → Test → Train → Docker Build & Test)
- **Containerization:** Production-ready Docker container with FastAPI serving
- **Deployment:** Kubernetes deployment with 2 replicas, health probes, and resource limits; Docker Compose stack with full monitoring
- **Monitoring:** Prometheus metrics, Grafana dashboards, and structured logging for production observability

### 12.2 Challenges Faced

- **Package Compatibility:** Python 3.14 lacked pre-built wheels for some ML packages. Resolved by using Python 3.11 in Docker and CI/CD while maintaining local 3.14 compatibility.
- **Pipeline Architecture:** Chose to embed StandardScaler inside the sklearn Pipeline to ensure consistent preprocessing during both training and inference, avoiding train-serve skew.
- **CI/CD Artifact Passing:** Docker build stage needed the trained model from the training stage. Resolved by using GitHub Actions `upload-artifact` / `download-artifact` for cross-job artifact transfer.
- **Kubernetes Local Testing:** Used Minikube with `minikube image load` to avoid needing a container registry for local K8s testing.

### 12.3 Future Work

- Add Gradient Boosting and XGBoost models for comparison
- Implement model versioning with MLflow Model Registry
- Add data drift detection for production monitoring
- Deploy to a cloud provider (GCP Cloud Run / AWS ECS) with continuous deployment
- Implement A/B testing framework for model comparison in production

---

## 13. References

| Resource | Link |
|----------|------|
| UCI ML Repository — Heart Disease Dataset | https://archive.ics.uci.edu/ml/datasets/heart+Disease |
| Scikit-learn Documentation | https://scikit-learn.org/stable/ |
| MLflow Documentation | https://mlflow.org/docs/latest/index.html |
| FastAPI Documentation | https://fastapi.tiangolo.com/ |
| Docker Documentation | https://docs.docker.com/ |
| Kubernetes Documentation | https://kubernetes.io/docs/ |
| GitHub Actions Documentation | https://docs.github.com/en/actions |
| Prometheus Client Python | https://github.com/prometheus/client_python |
| Grafana Documentation | https://grafana.com/docs/ |
| GitHub Repository | https://github.com/2025cs05014/MLOPs |
