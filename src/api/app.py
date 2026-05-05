"""
app.py

Local UI for Heart Disease Prediction (Flask)

Run:
    python src/api/app.py
Then open:
    http://127.0.0.1:5000
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from flask import Flask, request, render_template_string
from src.models.inference import HeartDiseasePredictor

app = Flask(__name__)

predictor = HeartDiseasePredictor()

HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>Heart Disease Prediction</title>
    <style>
        body { font-family: Arial; background: #f5f5f5; text-align: center; padding: 40px; }
        h1 { color: #c0392b; }
        input { padding: 8px; margin: 5px; width: 220px; }
        label { display: inline-block; width: 320px; text-align: right; margin-right: 8px; font-weight: bold; }
        .field { margin: 6px 0; }
        button { padding: 10px 25px; background: #c0392b; color: white; border: none; margin-top: 15px; cursor: pointer; font-size: 16px; border-radius: 4px; }
        button:hover { background: #96281b; }
        .result { margin-top: 25px; font-size: 20px; font-weight: bold; padding: 15px; border-radius: 6px; display: inline-block; }
        .disease { background: #fadbd8; color: #c0392b; }
        .healthy { background: #d5f5e3; color: #1e8449; }
        .info { font-size: 14px; color: #555; margin-top: 8px; }
    </style>
</head>
<body>
    <h1>Heart Disease Prediction</h1>
    <p>Enter patient details below to predict heart disease risk.</p>

    <form method="POST">
        <div class="field"><label>Age:</label><input name="age" placeholder="e.g. 55" required></div>
        <div class="field"><label>Sex (0=Female, 1=Male):</label><input name="sex" placeholder="0 or 1" required></div>
        <div class="field"><label>Chest Pain Type (0-3):</label><input name="cp" placeholder="0-3" required></div>
        <div class="field"><label>Resting Blood Pressure (mm Hg):</label><input name="trestbps" placeholder="e.g. 130" required></div>
        <div class="field"><label>Serum Cholesterol (mg/dl):</label><input name="chol" placeholder="e.g. 250" required></div>
        <div class="field"><label>Fasting Blood Sugar > 120 (0/1):</label><input name="fbs" placeholder="0 or 1" required></div>
        <div class="field"><label>Resting ECG Results (0-2):</label><input name="restecg" placeholder="0-2" required></div>
        <div class="field"><label>Max Heart Rate Achieved:</label><input name="thalach" placeholder="e.g. 150" required></div>
        <div class="field"><label>Exercise Induced Angina (0/1):</label><input name="exang" placeholder="0 or 1" required></div>
        <div class="field"><label>ST Depression (oldpeak):</label><input name="oldpeak" placeholder="e.g. 1.5" required></div>
        <div class="field"><label>Slope of ST Segment (0-2):</label><input name="slope" placeholder="0-2" required></div>
        <div class="field"><label>Major Vessels (0-4):</label><input name="ca" placeholder="0-4" required></div>
        <div class="field"><label>Thalassemia (1-3):</label><input name="thal" placeholder="1-3" required></div>
        <br>
        <button type="submit">Predict</button>
    </form>

    {% if result %}
        <div class="result {{ 'disease' if result.prediction == 1 else 'healthy' }}">
            {{ result.prediction_label }}
        </div>
        <div class="info">
            Confidence: {{ "%.2f"|format(result.confidence * 100) }}% &nbsp;|&nbsp;
            Risk Level: {{ result.risk_level }}
        </div>
    {% endif %}
</body>
</html>
"""


@app.route("/", methods=["GET", "POST"])
def home():
    result = None

    if request.method == "POST":
        input_data = {
            'age': float(request.form["age"]),
            'sex': int(request.form["sex"]),
            'cp': int(request.form["cp"]),
            'trestbps': float(request.form["trestbps"]),
            'chol': float(request.form["chol"]),
            'fbs': int(request.form["fbs"]),
            'restecg': int(request.form["restecg"]),
            'thalach': float(request.form["thalach"]),
            'exang': int(request.form["exang"]),
            'oldpeak': float(request.form["oldpeak"]),
            'slope': int(request.form["slope"]),
            'ca': int(request.form["ca"]),
            'thal': int(request.form["thal"]),
        }
        result = predictor.predict(input_data)

    return render_template_string(HTML, result=result)


if __name__ == "__main__":
    app.run(debug=True)
