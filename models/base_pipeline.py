from abc import ABC, abstractmethod
import pandas as pd
import numpy as np
import joblib
import glob
import os

class BaseDiseasePipeline(ABC):
    """Основной класс для всех пайплайнов"""
    
    DISEASE_NAME = None
    DEFAULT_THRESHOLD = 0.5
    
    def __init__(self):
        self.models = {}
        self.model_files = self.discover_model_files()
        self.load_models()
    
    def discover_model_files(self):
        """Discover all .pkl files in the disease directory"""
        model_dir = os.path.join(os.path.dirname(__file__), self.DISEASE_NAME)
        return glob.glob(os.path.join(model_dir, "*.pkl"))
    
    def load_models(self):
        """Load all discovered models"""
        if not self.model_files:
            raise FileNotFoundError(
                f"No model files found in {self.DISEASE_NAME} directory"
            )
        
        for model_file in self.model_files:
            key = os.path.basename(model_file)
            self.models[key] = joblib.load(model_file)
    
    def preprocess_data(self, row, features):
        """Предварительная обработка"""
        X = pd.DataFrame([row[features]], columns=features)
        X = X.replace([np.inf, -np.inf], np.nan).fillna(0).clip(-1e10, 1e10)
        return X.astype(np.float32)
    
    @abstractmethod
    def calculate_risk(self, row):
        """Рассчитываем риски"""
        pass
    
    @staticmethod
    def probability_to_score(prob, threshold):
        prob = min(max(prob, 0), 1)
        if prob < threshold:
            score = 4 * prob / threshold
        else:
            score = 4 + 4 * (prob - threshold) / (1 - threshold)
        return 10- round(score, 0)