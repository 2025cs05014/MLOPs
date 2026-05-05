"""
model_app.py

FastAPI application for Heart Disease prediction.
With Prometheus metrics and logging.

Run:
    uvicorn src.api.model_app:app --reload
or:
    python src/api/model_app.py
Then open:
    http://127.0.0.1:8000/docs
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field
import joblib
import pandas as pd
import logging
import time
from datetime import datetime
from prometheus_client import Counter, Histogram, generate_latest
from fastapi.responses import PlainTextResponse

# ============================================================
# LOGGING SETUP
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('api.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ============================================================
# PROMETHEUS METRICS
# ============================================================
PREDICTION_COUNTER = Counter(
    'predictions_total', 'Total number of predictions', ['result']
)
PREDICTION_LATENCY = Histogram(
    'prediction_latency_seconds', 'Time spent processing prediction'
)
REQUEST_COUNTER = Counter(
    'http_requests_total', 'Total HTTP requests', ['method', 'endpoint', 'status']
)

# ============================================================
# FASTAPI APP
# ============================================================
app = FastAPI(
    title="Heart Disease Prediction API",
    description="ML-powered API to predict heart disease risk based on patient health data.",
    version="1.0.0"
)

# ============================================================
# LOAD PIPELINE
# ============================================================
_DEFAULT_MODEL = os.path.join(os.path.dirname(os.path.dirname(__file__)), "models", "heart_model.pkl")
MODEL_PATH = os.getenv("MODEL_PATH", _DEFAULT_MODEL)

try:
    pipeline = joblib.load(MODEL_PATH)
    logger.info("Pipeline loaded successfully.")
except Exception as e:
    logger.error(f"Error loading pipeline: {e}")
    raise

FEATURE_ORDER = [
    'age', 'sex', 'cp', 'trestbps', 'chol', 'fbs',
    'restecg', 'thalach', 'exang', 'oldpeak', 'slope', 'ca', 'thal'
]


# ============================================================
# REQUEST / RESPONSE SCHEMAS
# ============================================================
class PatientData(BaseModel):
    """Input schema for patient data."""
    age: float = Field(..., ge=0, le=120, description="Age in years")
    sex: int = Field(..., ge=0, le=1, description="Sex (0=female, 1=male)")
    cp: int = Field(..., ge=0, le=3, description="Chest pain type (0-3)")
    trestbps: float = Field(..., ge=0, description="Resting blood pressure (mm Hg)")
    chol: float = Field(..., ge=0, description="Serum cholesterol (mg/dl)")
    fbs: int = Field(..., ge=0, le=1, description="Fasting blood sugar > 120 mg/dl")
    restecg: int = Field(..., ge=0, le=2, description="Resting ECG results (0-2)")
    thalach: float = Field(..., ge=0, description="Maximum heart rate achieved")
    exang: int = Field(..., ge=0, le=1, description="Exercise induced angina")
    oldpeak: float = Field(..., ge=0, description="ST depression induced by exercise")
    slope: int = Field(..., ge=0, le=2, description="Slope of peak exercise ST segment")
    ca: int = Field(..., ge=0, le=4, description="Number of major vessels colored by fluoroscopy")
    thal: int = Field(..., ge=0, le=7, description="Thalassemia (3=normal, 6=fixed defect, 7=reversible defect)")

    class Config:
        json_schema_extra = {
            "example": {
                "age": 55, "sex": 1, "cp": 2, "trestbps": 130, "chol": 250,
                "fbs": 0, "restecg": 1, "thalach": 150, "exang": 0,
                "oldpeak": 1.5, "slope": 1, "ca": 0, "thal": 3
            }
        }


class PredictionResponse(BaseModel):
    """Output schema for predictions."""
    prediction: int
    prediction_label: str
    confidence: float
    risk_level: str
    timestamp: str


# ============================================================
# ENDPOINTS
# ============================================================
@app.get("/")
def root():
    """Root endpoint with API information."""
    REQUEST_COUNTER.labels(method='GET', endpoint='/', status='200').inc()
    return {
        "message": "Heart Disease Prediction API",
        "version": "1.0.0",
        "endpoints": {
            "/predict": "POST - Make a prediction",
            "/health": "GET - Health check",
            "/metrics": "GET - Prometheus metrics"
        }
    }


@app.get("/health")
def health_check():
    """Health check endpoint."""
    REQUEST_COUNTER.labels(method='GET', endpoint='/health', status='200').inc()
    return {
        "status": "healthy",
        "pipeline_loaded": pipeline is not None,
        "timestamp": datetime.now().isoformat()
    }


@app.post("/predict", response_model=PredictionResponse)
def predict(patient: PatientData):
    """
    Predict heart disease risk for a patient.

    Returns:
    - Prediction (0 or 1)
    - Prediction label (No Heart Disease / Heart Disease)
    - Confidence score (0.0 - 1.0)
    - Risk level (Low / Medium / High)
    """
    start_time = time.time()

    try:
        input_data = patient.model_dump()
        df = pd.DataFrame([input_data])[FEATURE_ORDER]

        prediction = int(pipeline.predict(df)[0])
        probability = float(pipeline.predict_proba(df)[0][1])

        if probability < 0.3:
            risk_level = "Low Risk"
        elif probability < 0.6:
            risk_level = "Medium Risk"
        else:
            risk_level = "High Risk"

        prediction_label = "Heart Disease" if prediction == 1 else "No Heart Disease"

        latency = time.time() - start_time
        PREDICTION_LATENCY.observe(latency)
        PREDICTION_COUNTER.labels(result=prediction_label).inc()
        REQUEST_COUNTER.labels(method='POST', endpoint='/predict', status='200').inc()

        logger.info(
            f"Prediction: {prediction_label} | "
            f"Confidence: {probability:.4f} | "
            f"Risk: {risk_level} | "
            f"Latency: {latency:.4f}s | "
            f"Input: {input_data}"
        )

        return PredictionResponse(
            prediction=prediction,
            prediction_label=prediction_label,
            confidence=round(probability, 4),
            risk_level=risk_level,
            timestamp=datetime.now().isoformat()
        )

    except Exception as e:
        REQUEST_COUNTER.labels(method='POST', endpoint='/predict', status='500').inc()
        logger.error(f"Prediction error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/metrics")
def metrics():
    """Prometheus metrics endpoint."""
    return PlainTextResponse(generate_latest(), media_type="text/plain")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
