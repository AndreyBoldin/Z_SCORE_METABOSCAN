from models.base_pipeline import BaseDiseasePipeline
import os
import glob
import joblib

class ONCOPipeline(BaseDiseasePipeline):
    DISEASE_NAME = "ONCO"
    
    def __init__(self):
        self.onco_threshold = 0.62
        self.liver_threshold = 0.64
        super().__init__()
    
    def load_models(self):
        """Improved model loading with better error handling"""
        model_dir = os.path.dirname(__file__)
        
        # Verify directory exists
        if not os.path.exists(model_dir):
            raise FileNotFoundError(
                f"ONCO model directory not found at: {model_dir}\n"
                f"Current working directory: {os.getcwd()}"
            )
        
        model_files = glob.glob(os.path.join(model_dir, "*.pkl"))
        
        # Debug found files
        print(f"Found model files in {model_dir}: {model_files}")
        
        if len(model_files) < 2:
            raise FileNotFoundError(
                f"Need exactly 2 models (control and liver) in ONCO folder.\n"
                f"Found {len(model_files)} files: {model_files}\n"
                f"Directory contents: {os.listdir(model_dir)}"
            )
        
        # Identify models
        control_model = None
        liver_model = None
        
        for model_file in model_files:
            try:
                model_name = os.path.basename(model_file).lower()
                if 'liver' in model_name:
                    liver_model = joblib.load(model_file)
                    print(f"Loaded liver model: {model_file}")
                else:
                    control_model = joblib.load(model_file)
                    print(f"Loaded control model: {model_file}")
            except Exception as e:
                print(f"Error loading model {model_file}: {str(e)}")
                raise
        
        # Final validation
        if not control_model:
            raise ValueError("Control model not found! Files should NOT contain 'liver' in name")
        if not liver_model:
            raise ValueError("Liver model not found! Need one file with 'liver' in name")
        
        self.models = {
            'control': control_model,
            'liver': liver_model
        }
    
    def calculate_risk(self, row):
        try:
            # First stage - control model
            control_model = self.models['control']
            X_control = self.preprocess_data(row, control_model.feature_names_in_)
            control_proba = control_model.predict_proba(X_control)[0][0]
            
            if control_proba < self.onco_threshold:
                return {
                    "Группа риска": "Оценка пролиферативных процессов",
                    "Риск-скор": self.probability_to_score(control_proba, self.onco_threshold),
                    "Метод оценки": "onco-control модель",
                }
            
            # Second stage - liver model
            liver_model = self.models['liver']
            X_liver = self.preprocess_data(row, liver_model.feature_names_in_)
            liver_proba = 1 - liver_model.predict_proba(X_liver)[0][0]
            
            return {
                "Группа риска": "Оценка пролиферативных процессов",
                "Риск-скор": self.probability_to_score(liver_proba, self.liver_threshold),
                "Метод оценки": "onco-liver модель",
            }
            
        except Exception as e:
            print(f"Prediction error: {str(e)}")
            return {
                "Группа риска": "Оценка пролиферативных процессов",
                "Риск-скор": None,
                "Метод оценки": f"ML модель (ошибка: {str(e)})",
            }