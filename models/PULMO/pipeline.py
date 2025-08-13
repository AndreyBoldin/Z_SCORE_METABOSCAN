from models.base_pipeline import BaseDiseasePipeline

class PULMOPipeline(BaseDiseasePipeline):
    DISEASE_NAME = "PULMO"
    DEFAULT_THRESHOLD = 0.64
    
    def calculate_risk(self, row):
        # Get the first model (alphabetically by filename)
        model_name, model = next(iter(self.models.items()))
        
        X = self.preprocess_data(row, model.feature_names_in_)
        pred_proba = model.predict_proba(X)[0][1]
        
        return {
            "Группа риска": "Состояние дыхательной системы",
            "Риск-скор": self.probability_to_score(pred_proba, self.DEFAULT_THRESHOLD),
            "Метод оценки": "ML модель",
        }