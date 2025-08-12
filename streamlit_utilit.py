import pandas as pd
import numpy as np

def get_color_under_normal_dist(n):
    if n <= 0:
        return '#10b981'
    if n <= 10:
        return '#10b962'  # Light green (similar to 9-10)
    elif n <= 20:
        return '#50c150'  # Light yellow (similar to 7-8)
    elif n <= 30:
        return '#9fd047'  # Light orange (similar to 5-6)
    elif n <= 40:
        return '#feb61d'  # Light orange (similar to 5-6)
    elif n <= 50:
        return '#fe991d'  # Light orange (similar to 5-6)
    elif n <= 60:
        return '#f25708'  # Light orange (similar to 5-6)
    elif n <= 70:
        return "#f23b08"  # Light orange (similar to 5-6)
    elif n <= 80:
        return '#f21e08'  # Light orange (similar to 5-6)
    else:
        return '#c90909'  # Orange-red (similar to 3-4)


def calculate_metabolite_ratios(metabolomic_data):
    """Calculate all metabolite ratios from raw metabolomic data"""
    # Read data
    data = pd.read_excel(metabolomic_data)
    
    # Replace all negative values with 0 in the entire DataFrame
    data = data.map(lambda x: 0 if isinstance(x, (int, float)) and x < 0 else x)
    
    try:
        # Prepare all new columns in a dictionary first
        new_columns = {}
        
        # Acylcarnitines
        new_columns['(C2+C3)/C0'] = (data['C2'] + data['C3']) / data['C0']
        new_columns['CACT Deficiency (NBS)'] = data['C0'] / (data['C16'] + data['C18'])
        new_columns['CPT-1 Deficiency (NBS)'] = (data['C16'] + data['C18']) / data['C0']
        new_columns['CPT-2 Deficiency (NBS)'] = (data['C16'] + data['C18']) / data['C2']
        new_columns['EMA (NBS)'] = data['C4'] / data['C8']
        new_columns['IBD Deficiency (NBS)'] = data['C4'] / data['C2']
        new_columns['IVA (NBS)'] = data['C5'] / data['C2']
        new_columns['LCHAD Deficiency (NBS)'] = data['C16-OH'] / data['C16']
        new_columns['MA (NBS)'] = data['C3'] / data['C2']
        new_columns['MC Deficiency (NBS)'] = data['C16'] / data['C3']
        new_columns['MCAD Deficiency (NBS)'] = data['C8'] / data['C2']
        new_columns['MCKAT Deficiency (NBS)'] = data['C8'] / data['C10']
        new_columns['MMA (NBS)'] = data['C3'] / data['C0']
        new_columns['PA (NBS)'] = data['C3'] / data['C16']
        new_columns['С2/С0'] = data['C2'] / data['C0']
        new_columns['Ratio of Acetylcarnitine to Carnitine'] = data['C2'] / data['C0']
        
        # Calculate sums once to reuse
        sum_AC_OHs = (data['C5-OH'] + data['C14-OH'] + data['C16-1-OH'] + 
                     data['C16-OH'] + data['C18-1-OH'] + data['C18-OH'])
        sum_ACs = (data['C0'] + data['C10'] + data['C10-1'] + data['C10-2'] + 
                  data['C12'] + data['C12-1'] + data['C14'] + data['C14-1'] + 
                  data['C14-2'] + data['C16'] + data['C16-1'] + data['C18'] + 
                  data['C18-1'] + data['C18-2'] + data['C2'] + data['C3'] + 
                  data['C4'] + data['C5'] + data['C5-1'] + data['C5-DC'] + 
                  data['C6'] + data['C6-DC'] + data['C8'] + data['C8-1'])
        
        new_columns['Ratio of AC-OHs to ACs'] = sum_AC_OHs / sum_ACs
        
        СДК = (data['C14'] + data['C14-1'] + data['C14-2'] + data['C14-OH'] + 
               data['C16'] + data['C16-1'] + data['C16-1-OH'] + data['C16-OH'] + 
               data['C18'] + data['C18-1'] + data['C18-1-OH'] + data['C18-2'] + 
               data['C18-OH'])
        ССК = (data['C6'] + data['C6-DC'] + data['C8'] + data['C8-1'] + 
               data['C10'] + data['C10-1'] + data['C10-2'] + data['C12'] + 
               data['C12-1'])
        СКК = (data['C2'] + data['C3'] + data['C4'] + data['C5'] + data['C5-1'] + 
               data['C5-DC'] + data['C5-OH'])
        
        new_columns['СДК'] = СДК
        new_columns['ССК'] = ССК
        new_columns['СКК'] = СКК
        new_columns['Ratio of Medium-Chain to Long-Chain ACs'] = ССК / СДК
        new_columns['Ratio of Short-Chain to Long-Chain ACs'] = СКК / СДК
        new_columns['Ratio of Short-Chain to Medium-Chain ACs'] = СКК / ССК
        new_columns['SBCAD Deficiency (NBS)'] = data['C5'] / data['C0']
        new_columns['SCAD Deficiency (NBS)'] = data['C4'] / data['C3']
        new_columns['Sum of ACs'] = sum_AC_OHs + sum_ACs - data['C0']  # Subtract C0 since it's included in sum_ACs
        new_columns['Sum of ACs + С0'] = sum_AC_OHs + sum_ACs
        new_columns['Sum of ACs/C0'] = (sum_AC_OHs + sum_ACs - data['C0']) / data['C0']
        
        new_columns['Sum of MUFA-ACs'] = (data['C16-1-OH'] + data['C18-1-OH'] + 
                                        data['C10-1'] + data['C12-1'] + 
                                        data['C14-1'] + data['C16-1'] + 
                                        data['C18-1'] + data['C8-1'] + 
                                        data['C5-1'])
        new_columns['Sum of PUFA-ACs'] = data['C10-2'] + data['C14-2'] + data['C18-2']
        new_columns['TFP Deficiency (NBS)'] = data['C16'] / data['C16-OH']
        new_columns['VLCAD Deficiency (NBS)'] = data['C14-1'] / data['C16']
        new_columns['(C6+C8+C10)/C2'] = (data['C6'] + data['C8'] + data['C10']) / data['C2']
        new_columns['2MBG (NBS)'] = data['C5'] / data['C3']
        new_columns['Carnitine Uptake Defect (NBS)'] = (data['C0'] + data['C2'] + data['C3'] + 
                                                       data['C16'] + data['C18'] + 
                                                       data['C18-1']) / data['Citrulline']

        new_columns['C2 / C3'] = data['C2'] / data['C3']
        # NO- and urea cycle
        new_columns['GABR'] = data['Arginine'] / (data['Ornitine'] + data['Citrulline'])
        new_columns['Orn Synthesis'] = data['Ornitine'] / data['Arginine']
        new_columns['AOR'] = data['Arginine'] / data['Ornitine']
        new_columns['ADMA/(Adenosin+Arginine)'] = data['ADMA'] / (data['Adenosin'] + data['Arginine'])
        new_columns["Asymmetrical Arg Methylation"] = data['ADMA'] / data['Arginine']
        new_columns['Symmetrical Arg Methylation'] = data['TotalDMA (SDMA)'] / data['Arginine']
        new_columns['(Arg+HomoArg)/ADMA'] = (data['Arginine'] + data['Homoarginine']) / data['ADMA']
        new_columns['ADMA / NMMA'] = data['ADMA'] / data['NMMA']
        new_columns['NO-Synthase Activity'] = data['Citrulline'] / data['Arginine']
        new_columns['OTC Deficiency (NBS)'] = data['Ornitine'] / data['Citrulline']
        new_columns['Ratio of HArg to ADMA'] = data['Homoarginine'] / data['ADMA']
        new_columns['Ratio of HArg to SDMA'] = data['Homoarginine'] / data['TotalDMA (SDMA)']
        new_columns['Sum of Asym. and Sym. Arg Methylation'] = (data['TotalDMA (SDMA)'] + data['ADMA']) / data['Arginine']
        new_columns['Sum of Dimethylated Arg'] = data['TotalDMA (SDMA)'] + data['ADMA']
        new_columns['Cit Synthesis'] = data['Citrulline'] / data['Ornitine']
        new_columns['CPS Deficiency (NBS)'] = data['Citrulline'] / data['Phenylalanine']
        new_columns['HomoArg Synthesis'] = data['Homoarginine'] / (data['Arginine'] + data['Lysine'])
        new_columns['Ratio of Pro to Cit'] = data['Proline'] / data['Citrulline']

        # Tryptophan metabolism
        new_columns['Kynurenine / Trp'] = data['Kynurenine'] / data['Tryptophan']
        new_columns['Serotonin / Trp'] = data['Serotonin'] / data['Tryptophan']
        new_columns['Trp/(Kyn+QA)'] = data['Tryptophan'] / (data['Kynurenine'] + data['Quinolinic acid'])
        new_columns['Kyn/Quin'] = data['Kynurenine'] / data['Quinolinic acid']
        new_columns['Quin/HIAA'] = data['Quinolinic acid'] / data['HIAA']
        new_columns['Tryptamine / IAA'] = data['Tryptamine'] / data['Indole-3-acetic acid']
        new_columns['Kynurenic acid / Kynurenine'] = data['Kynurenic acid'] / data['Kynurenine']

        # Amino acids
        new_columns['Asn Synthesis'] = data['Asparagine'] / data['Aspartic acid']
        new_columns['Glutamine/Glutamate'] = data['Glutamine'] / data['Glutamic acid']
        new_columns['Gly Synthesis'] = data['Glycine'] / data['Serine']
        new_columns['GSG Index'] = data['Glutamic acid'] / (data['Serine'] + data['Glycine'])
        new_columns['GSG_index'] = data['Glutamic acid'] / (data['Serine'] + data['Glycine'])
        new_columns['Sum of Aromatic AAs'] = data['Phenylalanine'] + data['Tyrosin']
        new_columns['BCAA'] = data['Summ Leu-Ile'] + data['Valine']
        new_columns['BCAA/AAA'] = (data['Valine'] + data['Summ Leu-Ile']) / (data['Phenylalanine'] + data['Tyrosin'])
        new_columns['Alanine / Valine'] = data['Alanine'] / data['Valine']
        new_columns['DLD (NBS)'] = data['Proline'] / data['Phenylalanine']
        new_columns['MTHFR Deficiency (NBS)'] = data['Methionine'] / data['Phenylalanine']
        
        # Calculate sums once for AA ratios
        sum_non_essential = (data['Alanine'] + data['Arginine'] + data['Asparagine'] + 
                           data['Aspartic acid'] + data['Glutamine'] + 
                           data['Glutamic acid'] + data['Glycine'] + data['Proline'] + 
                           data['Serine'] + data['Tyrosin'])
        sum_essential = (data['Histidine'] + data['Summ Leu-Ile'] + data['Lysine'] + 
                        data['Methionine'] + data['Phenylalanine'] + 
                        data['Threonine'] + data['Tryptophan'] + data['Valine'])
        
        new_columns['Ratio of Non-Essential to Essential AAs'] = sum_non_essential / sum_essential
        new_columns['Sum of AAs'] = sum_non_essential + sum_essential
        new_columns['Sum of Essential Aas'] = sum_essential
        new_columns['Sum of Non-Essential AAs'] = sum_non_essential
        new_columns['Sum of Solely Glucogenic AAs'] = (data['Alanine'] + data['Arginine'] + 
                                                     data['Asparagine'] + data['Aspartic acid'] + 
                                                     data['Glutamine'] + data['Glutamic acid'] + 
                                                     data['Glycine'] + data['Histidine'] + 
                                                     data['Methionine'] + data['Proline'] + 
                                                     data['Serine'] + data['Threonine'] + 
                                                     data['Valine'])
        new_columns['Sum of Solely Ketogenic AAs'] = data['Summ Leu-Ile'] + data['Lysine']
        new_columns['Valinemia (NBS)'] = data['Valine'] / data['Phenylalanine']
        new_columns['Carnosine Synthesis'] = data['Carnosine'] / data['Histidine']
        new_columns['Histamine Synthesis'] = data['Histamine'] / data['Histidine']

        # Betaine_choline metabolism
        new_columns['Betaine/choline'] = data['Betaine'] / data['Choline']
        new_columns['Methionine + Taurine'] = data['Methionine'] + data['Taurine']
        new_columns['DMG / Choline'] = data['DMG'] / data['Choline']
        new_columns['TMAO Synthesis'] = data['TMAO'] / (data['Betaine'] + data['C0'] + data['Choline'])
        new_columns['TMAO Synthesis (direct)'] = data['TMAO'] / data['Choline']
        new_columns['Met Oxidation'] = data['Methionine-Sulfoxide'] / data['Methionine']

        # Vitamins
        new_columns['Riboflavin / Pantothenic'] = data['Riboflavin'] / data['Pantothenic']

        # ADDED: Oncology-specific ratios that were missing
        new_columns['Arg/ADMA'] = data['Arginine'] / data['ADMA']
        new_columns['Arg/Orn+Cit'] = data['Arginine'] / (data['Ornitine'] + data['Citrulline'])
        new_columns['Glutamine/Glutamate'] = data['Glutamine'] / data['Glutamic acid']
        new_columns['Pro/Cit'] = data['Proline'] / data['Citrulline']
        new_columns['Kyn/Trp'] = data['Kynurenine'] / data['Tryptophan']
        new_columns['Trp/Kyn'] = data['Tryptophan'] / data['Kynurenine']
        
        # Arthritis
        new_columns['Phe/Tyr'] = data['Phenylalanine'] / data['Tyrosin']
        new_columns['Glycine/Serine'] = data['Glycine'] / data['Serine']
        # Lungs
        new_columns['C4 / C2'] = data['C4'] / data['C2']
        new_columns['Valine / Alanine'] = data['Valine'] / data['Alanine']
        # Liver
        new_columns['C0/(C16+C18)'] = data['C0'] / (data['C16'] + data['C18'])
        new_columns['(Leu+IsL)/(C3+С5+С5-1+C5-DC)'] = (data['Summ Leu-Ile']) / (data['C3'] + data['C5'] + data['C5-1'] + data['C5-DC'])
        new_columns['Val/C4'] = data['Valine'] / data['C4']
        new_columns['(C16+C18)/C2'] = (data['C16'] + data['C18']) / data['C2']
        new_columns['C3 / C0'] = data['C3'] / data['C0']

        # Convert the dictionary to a DataFrame
        new_data = pd.DataFrame(new_columns)

        # Get columns that exist in both DataFrames
        common_cols = data.columns.intersection(new_data.columns)

        # For overlapping columns, fill NaN in original data with new_data values
        for col in common_cols:
            data[col] = data[col].fillna(new_data[col])

        # Get columns that only exist in new_data
        new_cols_only = new_data.columns.difference(data.columns)

        # Add the new columns
        data = pd.concat([data, new_data[new_cols_only]], axis=1)

        # Drop Group column if it exists
        if 'Group' in data.columns:
            data = data.drop('Group', axis=1)

        return data

    except Exception as e:
        print(f"Error calculating metabolite ratios: {str(e)}")
        return None

import numpy as np

def prepare_final_dataframe_old(risk_params_data, metabolomic_data_with_ratios):
    # Load the data
    risk_params = pd.read_excel(risk_params_data)
    metabolic_data = pd.read_excel(metabolomic_data_with_ratios)
    
    # Get values for each marker from metabolomic data
    values_conc = []
    for metabolite in risk_params['Маркер / Соотношение']:
        try:
            value = metabolic_data.loc[0, metabolite]
            # Handle negative and infinite values
            if pd.isna(value) or np.isinf(value):
                values_conc.append(np.nan)
            elif value < 0:
                values_conc.append(0)
            else:
                values_conc.append(value)
        except KeyError:
            values_conc.append(np.nan)  # Handle missing metabolites
    
    risk_params['Patient'] = values_conc
    
    # Drop rows with infinite or NaN values in Patient column
    risk_params = risk_params[~risk_params['Patient'].isin([np.inf, -np.inf]) & 
                  ~risk_params['Patient'].isna()].copy()
    
    subgroup_list=[]
    subgroup_scores=[]
    categories=risk_params['Категория'].unique()
    for category in categories:
        data_category=risk_params[risk_params['Категория']==category]
        metabolite_inputs=[]
        for index, row in data_category.iterrows():
            metabolite_input=0
            weight=data_category.loc[index, 'веса']
            patient_value=data_category.loc[index, 'Patient']
            norm_1=data_category.loc[index, 'norm_1']
            norm_2=data_category.loc[index, 'norm_2']
            risk_1=data_category.loc[index, 'High_risk_1']
            risk_2=data_category.loc[index, 'High_risk_2']
            metab_group=data_category.loc[index, 'Группа_метаб']
            if metab_group==0:
                if norm_1<=patient_value<=norm_2:
                    metabolite_input=0
                elif risk_1<=patient_value<norm_1 or norm_2<patient_value<=risk_2:
                    metabolite_input=1
                else:
                    metabolite_input=2
            elif metab_group==1:
                if patient_value<=norm_2:
                    metabolite_input=0
                elif norm_2<patient_value<=risk_2:
                    metabolite_input=1
                else:
                    metabolite_input=2
            else:
                if norm_1<=patient_value:
                    metabolite_input=0
                elif risk_1<=patient_value<norm_1:
                    metabolite_input=1
                else:
                    metabolite_input=2
            metabolite_inputs.append(metabolite_input*weight)
            max_score=data_category['веса'].sum()*2
            subgroup_score=sum(metabolite_inputs)/max_score*100
        subgroup_scores.append(subgroup_score)
        subgroup_list.append(category)
    
    # in risk_params make column Subgroup_score and for each row where [Категория] is in subgroup_list, [Subgroup_score] is subgroup_scores[subgroup_list.index([Категория])]
    risk_params['Subgroup_score'] = np.nan
    for index, row in risk_params.iterrows():
        if row['Категория'] in subgroup_list:
            risk_params.loc[index, 'Subgroup_score'] = subgroup_scores[subgroup_list.index(row['Категория'])]
    
    # in risk_params make column Subgroup_score and for each row where [Категория] is in subgroup_list, [Subgroup_score] is subgroup_scores[subgroup_list.index([Категория])]
    risk_params['Subgroup_score'] = np.nan
    for index, row in risk_params.iterrows():
        if row['Категория'] in subgroup_list:
            risk_params.loc[index, 'Subgroup_score'] = subgroup_scores[subgroup_list.index(row['Категория'])]
    return risk_params

def prepare_final_dataframe_zscore(risk_params_data, metabolomic_data_with_ratios, ref_data_path):
    """
    Подготавливает итоговый датафрейм с расчетами метаболитов и оценками рисков
    
    Параметры:
        risk_params_data - путь к файлу с параметрами рисков
        metabolomic_data_with_ratios - путь к файлу с метаболическими данными
        ref_data_path - путь к файлу с референсными значениями
        
    Возвращает:
        Датафрейм с рассчитанными значениями и оценками
    """
    # Загрузка данных
    risk_params = pd.read_excel(risk_params_data)
    metabolic_data = pd.read_excel(metabolomic_data_with_ratios)
    
    # Загрузка и подготовка референсных данных
    ref_stats = (
        pd.read_excel(ref_data_path, header=None)
        .pipe(lambda df: df.set_axis(['stat'] + list(df.iloc[0, 1:]), axis=1)
        .drop(0)
        .set_index('stat')
        .apply(lambda x: pd.to_numeric(x.astype(str).str.replace(',', '.'), errors='coerce'))
        ))
    
    # Функция для расчета z-скор
    def calculate_zscore(metabolite, value):
        """Рассчитывает z-score для метаболита"""
        if metabolite not in ref_stats.columns:
            return np.nan
            
        mean = ref_stats.loc['mean', metabolite]
        sd = ref_stats.loc['sd', metabolite]
        
        return round((value - mean)/sd, 2) if pd.notna(sd) and sd > 0 else np.nan
    
    # Обработка метаболитов
    results = []
    for metabolite in risk_params['Маркер / Соотношение']:
        try:
            value = metabolic_data.loc[0, metabolite]
            
            # Обработка некорректных значений
            if pd.isna(value) or np.isinf(value):
                results.append((np.nan, np.nan))
                continue
                
            value = max(0, value)  # Отрицательные значения заменяем на 0
            z_score = calculate_zscore(metabolite, value) if value >= 0 else np.nan
            results.append((value, z_score))
            
        except KeyError:
            results.append((np.nan, np.nan))
    
    # Добавляем результаты в датафрейм
    risk_params = risk_params.assign(
        Patient=[r[0] for r in results],
        Z_score=[r[1] for r in results]
    ).copy()
    
    # Расчет групповых оценок
    def calculate_subgroup_score(group):
        """Рассчитывает оценку для подгруппы"""
        inputs = []
        for _, row in group.iterrows():
            value = abs(row['Z_score'])
            
            risk = 0 if value < 1.54 else \
                    1 if (1.54 <= value <= 1.96) else \
                    2 if value > 1.96 else np.nan
            if risk is np.nan:
                print(row['Маркер / Соотношение'], row['Z_score'])
                        
            inputs.append(risk * row['веса'])
        
        max_score = group['веса'].sum() * 2
        return sum(inputs) / max_score * 100 if max_score > 0 else 0
    
    # Применяем расчет для каждой категории
    subgroup_scores = {
        cat: calculate_subgroup_score(group)
        for cat, group in risk_params.groupby('Категория')
    }
    
    # Добавляем оценки подгрупп
    risk_params['Subgroup_score'] = risk_params['Категория'].map(subgroup_scores)
    
    return risk_params
    
def probability_to_score(prob, threshold):
    prob = min(max(prob, 0), 1)
    if prob < threshold:
        score = 4 * prob / threshold
    else:
        score = 4 + 4 * (prob - threshold) / (1 - threshold)
    return 10- round(score, 0)

from importlib import import_module
def calculate_risks(risk_params_data, metabolic_data_with_ratios):
    """
    Расчет комбинированных рисков с использованием:
    - ML-моделей для определенных групп (Онко, ССЗ, Печень, Легкие, РА)
    - Параметров рисков для остальных групп
    Возвращает DataFrame с колонками: ['Группа риска', 'Риск-скор', 'Метод оценки']
    """
    # Группы, для которых используем только ML модели
    ml_only_groups = {
        "Состояние сердечно-сосудистой системы",
        "Состояние функции печени",
        "Оценка пролиферативных процессов",
    }
    
    metabolic_data_with_ratios = metabolic_data_with_ratios[~metabolic_data_with_ratios.index.duplicated()]
    risk_params_data = risk_params_data[~risk_params_data.index.duplicated()]
    
    # Define which diseases to process
    disease_pipelines = {
    }
    
    results = []
    
    # Process each row
    for idx, row in metabolic_data_with_ratios.iterrows():
        for disease_name, pipeline_path in disease_pipelines.items():
            try:
                # Dynamically import and instantiate the pipeline
                module_path, class_name = pipeline_path.rsplit('.', 1)
                module = import_module(f"models.{module_path}")
                pipeline_class = getattr(module, class_name)
                pipeline = pipeline_class()
                
                # Calculate risk
                result = pipeline.calculate_risk(row)
                results.append(result)
                
            except Exception as e:
                print(f"Error processing {disease_name}: {str(e)}")
                results.append({
                    "Группа риска": disease_name,
                    "Риск-скор": None,
                    "Метод оценки": f"ML модель (ошибка: {str(e)})"
                })
       

    # 2. Process other groups with parameter-based method
    # Filter out ML-only groups
    other_groups = set(risk_params_data['Группа_риска'].unique())
    
    if other_groups:
        # Prepare parameter risk_params_data
        values_conc = []
        for metabolite in risk_params_data['Маркер / Соотношение']:
            try:
                value = row[metabolite]
                if pd.isna(value) or np.isinf(value):
                    values_conc.append(np.nan)
                elif value < 0:
                    values_conc.append(0)
                else:
                    values_conc.append(value)
            except KeyError:
                values_conc.append(np.nan)
        
        risk_params_data['Patient'] = values_conc
        risk_params_data = risk_params_data[~risk_params_data['Patient'].isin([np.inf, -np.inf]) & ~risk_params_data['Patient'].isna()].copy()
        
        # Calculate scores for remaining groups
        for risk_group in other_groups:
            risk_params_data_group = risk_params_data[risk_params_data['Группа_риска'] == risk_group]
            if len(risk_params_data_group) == 0:
                continue
                
            metabolite_inputs = []
            for index, row_group in risk_params_data_group.iterrows():
                metabolite_input = 0
                weight = row_group['веса']
                patient_value = row_group['Patient']
                norm_1 = row_group['norm_1']
                norm_2 = row_group['norm_2']
                risk_1 = row_group['High_risk_1']
                risk_2 = row_group['High_risk_2']
                metab_group = row_group['Группа_метаб']
                
                if metab_group == 0:
                    if norm_1 <= patient_value <= norm_2:
                        metabolite_input = 0
                    elif risk_1 <= patient_value < norm_1 or norm_2 < patient_value <= risk_2:
                        metabolite_input = 1
                    else:
                        metabolite_input = 2
                elif metab_group == 1:
                    if patient_value <= norm_2:
                        metabolite_input = 0
                    elif norm_2 < patient_value <= risk_2:
                        metabolite_input = 1
                    else:
                        metabolite_input = 2
                else:
                    if norm_1 <= patient_value:
                        metabolite_input = 0
                    elif risk_1 <= patient_value < norm_1:
                        metabolite_input = 1
                    else:
                        metabolite_input = 2
                
                metabolite_inputs.append(metabolite_input * weight)
            
            max_score = risk_params_data_group['веса'].sum() * 2
            group_score = 10 - sum(metabolite_inputs) / max_score * 10
            results.append({
                "Группа риска": risk_group,
                "Риск-скор": np.round(group_score, 0),
                "Метод оценки": "Параметры"
            })

    # Create final DataFrame
    result_df = pd.DataFrame(results)
    
    # Sort by risk score (descending) and group name
    result_df.sort_values(['Группа риска'], ascending=True, inplace=True)
    
    return result_df[['Группа риска', 'Риск-скор', 'Метод оценки']].reset_index(drop=True)

