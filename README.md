# Heart Disease Prediction — MLOps End-to-End Pipeline

An end-to-end ML pipeline for predicting heart disease risk using the UCI Heart Disease dataset, built with modern MLOps practices.

## Prerequisites

- **Python** 3.10+ (tested on 3.11 and 3.14)
- **Docker Desktop** (for containerization & Kubernetes)
- **minikube** + **kubectl** (for K8s deployment)
- **Git**
- macOS / Linux (Windows WSL2 should also work)

## Architecture

```
┌─────────────┐     ┌──────────────┐     ┌─────────────────┐     ┌──────────────┐
│  Data (CSV)  │────▶│  Preprocess   │────▶│  Train + MLflow  │────▶│  Model (.pkl) │
└─────────────┘     └──────────────┘     └─────────────────┘     └──────┬───────┘
                                                                        │
                    ┌──────────────┐     ┌─────────────────┐           │
                    │   Prometheus  │◀────│  FastAPI + Docker │◀──────────┘
                    │   /metrics    │     │  /predict        │
                    └──────────────┘     └────────┬────────┘
                                                  │
                                         ┌────────▼────────┐
                                         │  Kubernetes (K8s)│
                                         │  LoadBalancer     │
                                         └─────────────────┘
```

## Project Structure

```
Assignment/
├── .github/workflows/ci.yml    # CI/CD pipeline (lint, test, train, docker)
├── data/heart.csv               # UCI Heart Disease dataset
├── notebooks/EDA.ipynb          # Exploratory Data Analysis
├── src/
│   ├── data/preprocess.py       # Data loading & cleaning
│   ├── models/
│   │   ├── train.py             # Model training with MLflow tracking
│   │   └── inference.py         # Inference module
│   └── api/
│       ├── app.py               # Flask UI
│       └── model_app.py         # FastAPI REST API + Prometheus
├── tests/                       # Unit tests (pytest)
│   ├── test_preprocess.py
│   ├── test_model.py
│   └── test_api.py
├── k8s/                         # Kubernetes manifests
│   ├── deployment.yaml
│   └── service.yaml
├── Dockerfile
├── requirements.txt
└── README.md
```

## Quick Start (End-to-End)

All commands below assume you are in the **project root directory**.

### 1. Clone & Setup

```bash
git clone <YOUR_REPO_URL>
cd Assignment
python3 -m venv .venv
source .venv/bin/activate      # On Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Explore the Data (EDA)

```bash
jupyter notebook notebooks/EDA.ipynb
```

Run all cells to generate visualizations (saved to `screenshots/`).

### 3. Train the Model

```bash
python src/models/train.py
```

This will:
- Train Logistic Regression and Random Forest with GridSearchCV
- Log all parameters, metrics, and models to **MLflow** (`mlruns/`)
- Save the best model to `src/models/heart_model.pkl`

View experiment tracking:

```bash
mlflow ui --backend-store-uri mlruns/
# Open http://127.0.0.1:5000
```

### 4. Test Inference Locally

```bash
python src/models/inference.py
```

### 5. Run Unit Tests

```bash
python -m pytest tests/ -v
```

Runs 24 tests covering data preprocessing, model training, inference, and API endpoints.

### 6. Start the API Locally

**FastAPI (recommended):**

```bash
uvicorn src.api.model_app:app --reload
# Open http://127.0.0.1:8000/docs for Swagger UI
```

**Flask UI (browser form):**

```bash
python src/api/app.py
# Open http://127.0.0.1:5000
```

### 7. Docker

```bash
docker build -t heart-disease-api .
docker run -p 8000:8000 heart-disease-api
```

Test the container:

```bash
curl http://localhost:8000/health
curl -X POST http://localhost:8000/predict \
  -H "Content-Type: application/json" \
  -d '{"age":55,"sex":1,"cp":2,"trestbps":130,"chol":250,"fbs":0,"restecg":1,"thalach":150,"exang":0,"oldpeak":1.5,"slope":1,"ca":0,"thal":3}'
```

### 8. Kubernetes Deployment (Minikube)

**Prerequisites:**

- Docker Desktop installed and running
- minikube (`brew install minikube` or download from [minikube releases](https://github.com/kubernetes/minikube/releases))
- kubectl (`brew install kubectl`)

**Step 1 — Start Minikube cluster:**

```bash
minikube start --driver=docker
```

**Step 2 — Build Docker image (if not already built):**

```bash
docker build -t heart-disease-api:latest .
```

**Step 3 — Load image into Minikube:**

```bash
minikube image load heart-disease-api:latest
```

**Step 4 — Apply Kubernetes manifests:**

```bash
kubectl apply -f k8s/deployment.yaml
kubectl apply -f k8s/service.yaml
```

**Step 5 — Wait for rollout to complete:**

```bash
kubectl rollout status deployment/heart-disease-api --timeout=90s
```

**Step 6 — Verify pods, service, and deployment:**

```bash
kubectl get pods,svc,deployment
```

Expected output:
```
NAME                                     READY   STATUS    RESTARTS   AGE
pod/heart-disease-api-xxxx-xxxxx         1/1     Running   0          30s
pod/heart-disease-api-xxxx-xxxxx         1/1     Running   0          30s

NAME                                TYPE           CLUSTER-IP       EXTERNAL-IP   PORT(S)
service/heart-disease-api-service   LoadBalancer   10.x.x.x         <pending>     80:3xxxx/TCP

NAME                                READY   UP-TO-DATE   AVAILABLE   AGE
deployment/heart-disease-api        2/2     2            2           30s
```

**Step 7 — Expose service via Minikube tunnel:**

```bash
minikube service heart-disease-api-service --url
```

This outputs a URL like `http://127.0.0.1:XXXXX`. Keep this terminal open.

**Step 8 — Test endpoints:**

```bash
# Health check
curl http://127.0.0.1:<PORT>/health

# Prediction
curl -X POST http://127.0.0.1:<PORT>/predict \
  -H "Content-Type: application/json" \
  -d '{"age":55,"sex":1,"cp":2,"trestbps":130,"chol":250,"fbs":0,"restecg":1,"thalach":150,"exang":0,"oldpeak":1.5,"slope":1,"ca":0,"thal":3}'

# Prometheus metrics
curl http://127.0.0.1:<PORT>/metrics

# Swagger UI
open http://127.0.0.1:<PORT>/docs
```

**Cleanup:**

```bash
kubectl delete -f k8s/service.yaml
kubectl delete -f k8s/deployment.yaml
minikube stop
```

### Kubernetes Manifest Details

- **`k8s/deployment.yaml`** — 2 replicas, liveness/readiness probes on `/health`, resource requests/limits
- **`k8s/service.yaml`** — LoadBalancer Service mapping port 80 → container port 8000

## Model Details

- **Dataset**: UCI Heart Disease (303 samples, 13 features, binary target)
- **Models evaluated**: Logistic Regression, Random Forest
- **Selection criterion**: Highest mean 5-fold CV ROC-AUC
- **Best model**: Logistic Regression (CV ROC-AUC ≈ 0.90)

## API Endpoints

| Endpoint    | Method | Description              |
|-------------|--------|--------------------------|
| `/`         | GET    | API info                 |
| `/health`   | GET    | Health check             |
| `/predict`  | POST   | Predict heart disease    |
| `/metrics`  | GET    | Prometheus metrics       |

### `/predict` Request Body

```json
{
  "age": 55,
  "sex": 1,
  "cp": 2,
  "trestbps": 130,
  "chol": 250,
  "fbs": 0,
  "restecg": 1,
  "thalach": 150,
  "exang": 0,
  "oldpeak": 1.5,
  "slope": 1,
  "ca": 0,
  "thal": 3
}
```

### Sample Response

```json
{
  "prediction": 0,
  "prediction_label": "No Heart Disease",
  "confidence": 0.1651,
  "risk_level": "Low Risk",
  "timestamp": "2026-05-03T11:32:57.869206"
}
```

## CI/CD Pipeline

GitHub Actions workflow (`.github/workflows/ci.yml`) runs on every push/PR:
1. **Lint** — flake8 code quality check
2. **Test** — pytest unit tests
3. **Train** — model training with artifact upload
4. **Docker** — build & smoke-test container

## Monitoring (Prometheus + Grafana)

### Quick Start with Docker Compose

```bash
docker compose up -d --build
```

This starts three services:

| Service | URL | Credentials |
|---------|-----|-------------|
| **Heart Disease API** | http://localhost:8000 | — |
| **Prometheus** | http://localhost:9090 | — |
| **Grafana** | http://localhost:3000 | admin / admin |

### Grafana Dashboard

A pre-provisioned dashboard **"Heart Disease API Monitoring"** is auto-loaded with panels for:
- Total predictions count
- Predictions rate by result
- Prediction latency (p95)
- HTTP request rate by method/endpoint/status
- Average prediction latency over time

### Metrics Tracked

- `predictions_total` — counter by prediction result
- `prediction_latency_seconds` — histogram of inference time
- `http_requests_total` — counter by method, endpoint, status

### Logging

- Structured logs to `api.log` and stdout
- Request/response logging with timestamps

### Cleanup

```bash
docker compose down
```
