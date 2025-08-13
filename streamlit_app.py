import streamlit as st
import os
import tempfile
import pandas as pd
from datetime import datetime
import logging

from streamlit_utilit import *

def validate_inputs(name, file1):
    """Validate user inputs before processing"""
    if not name.strip():
        st.error("Please enter a valid patient name")
        return False
    if not file1:
        st.error("Please upload metabolomic data file")
        return False
    return True

def display_group_cards(risk_params_df, risk_scores):
    # Group by risk group first
    grouped = risk_params_df.groupby('Группа_риска')
    
    for group_name, group_df in grouped:
        # Get unique categories within this risk group
        categories = group_df[['Категория', 'Subgroup_score']].drop_duplicates()
        
        # Calculate average score for the risk group header
        group_score = risk_scores['Риск-скор'][risk_scores['Группа риска'] == group_name].mean()
        group_color = get_color_under_normal_dist(100 - group_score *10)
        
        # Display risk group header
        st.markdown(f"""
        <div style="
            border-left: 5px solid {group_color};
            padding: 10px;
            margin: 10px 0 5px 0;
            background-color: #f0f2f6;
            border-radius: 5px;
            font-weight: bold;
        ">
            <div style="display: flex; justify-content: space-between;">
                <span>{group_name}</span>
                <span>Балл: {group_score:.0f}/10</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Display each category under this risk group
        for _, row in categories.iterrows():
            score = row['Subgroup_score']
            color = get_color_under_normal_dist(score)
            
            st.markdown(f"""
            <div style="
                border-left: 3px solid {color};
                padding: 8px;
                margin: 2px 0 2px 15px;
                background-color: #f8f9fa;
                border-radius: 3px;
            ">
                <div style="display: flex; justify-content: space-between;">
                    <span>{row['Категория']}</span>
                    <span>{score:.1f}%</span>
                </div>
                <div style="
                    height: 6px;
                    background: #e9ecef;
                    margin-top: 5px;
                    border-radius: 3px;
                ">
                    <div style="
                        width: {100 - score}%;
                        height: 100%;
                        background-color: {color};
                        border-radius: 3px;
                    "></div>
                </div>
            </div>
            """, unsafe_allow_html=True)

def main():
    st.set_page_config(
        page_title="Отчет Metaboscan",
        page_icon="🏥",
        layout="wide",
        initial_sidebar_state="collapsed"
    )

    # Apply smaller font sizes via custom CSS
    st.markdown("""
        <style>
            body {
                font-size: 14px !important;
            }
            .stTextInput input, .stNumberInput input, .stSelectbox select, .stDateInput input {
                font-size: 14px !important;
            }
            .stDataFrame {
                font-size: 14px !important;
            }
            .stButton button {
                font-size: 14px !important;
            }
        </style>
    """, unsafe_allow_html=True)

    # Path to the reference file
    REF_FILE = "Ref.xlsx"
    
    # Create two columns
    col1, col2 = st.columns([1, 1])
    
    with col1:
        with st.form("report_form"):
            st.write("Информация о пациенте")
            
            cols = st.columns(4)
            with cols[0]:
                name = st.text_input("Полное имя (ФИО)", placeholder="Иванов Иван Иванович")
            with cols[1]:
                age = st.number_input("Возраст", min_value=0, max_value=120, value=47)
            with cols[2]:
                gender = st.selectbox("Пол", ("М", "Ж"), index=0)
            with cols[3]:
                date = st.date_input("Дата отчета", datetime.now(), format="DD.MM.YYYY")
            
            st.write("Загрузите данные")
            
            metabolomic_data = st.file_uploader(
                "Метаболомный профиль пациента (Excel)",
                type=["xlsx", "xls"],
                key="metabolomic_data"
            )
            
            submitted = st.form_submit_button("Сформировать отчет", type="primary")
    
    with col2:
        st.write("Редактирование параметров")
        
        if os.path.exists(REF_FILE):
            try:
                # Initialize session state for both original and edited data
                if 'original_ref' not in st.session_state or 'edited_ref' not in st.session_state:
                    xls = pd.ExcelFile(REF_FILE)
                    st.session_state.original_ref = {
                        sheet_name: xls.parse(sheet_name) 
                        for sheet_name in xls.sheet_names
                    }
                    st.session_state.edited_ref = {
                        sheet_name: df.copy() 
                        for sheet_name, df in st.session_state.original_ref.items()
                    }
                
                # Create tabs for each sheet
                tabs = st.tabs(st.session_state.edited_ref.keys())
                
                for tab, (sheet_name, df) in zip(tabs, st.session_state.edited_ref.items()):
                    with tab:
                        edited_df = st.data_editor(
                            df,
                            num_rows="dynamic",
                            use_container_width=True,
                            height=300,
                            key=f"editor_{sheet_name}"
                        )
                        st.session_state.edited_ref[sheet_name] = edited_df
                
                if st.button("Сбросить изменения", key="reset_button"):
                    st.session_state.edited_ref = {
                        sheet_name: df.copy() 
                        for sheet_name, df in st.session_state.original_ref.items()
                    }
                    st.rerun()
                
            except Exception as e:
                st.error(f"Ошибка при загрузке файла параметров: {str(e)}")
        else:
            st.error(f"Файл параметров не найден: {REF_FILE}")
            st.session_state.edited_ref = None
    
    if submitted:
        if validate_inputs(name, metabolomic_data):
            if not os.path.exists(REF_FILE):
                st.error("Reference file not found. Cannot proceed without Ref.xlsx")
                return
                
            if 'edited_ref' not in st.session_state:
                st.error("Reference data not loaded")
                return

            required_sheets = ['Params_metaboscan', 'Ref_stats']
            for sheet in required_sheets:
                if sheet not in st.session_state.edited_ref:
                    st.error(f"Required sheet '{sheet}' not found in reference file")
                    return

            with st.spinner("🔬 Читаем данные и генерируем отчет. Это займет не больше минуты..."):
                with tempfile.TemporaryDirectory() as temp_dir:
                    try:
                        # Save Params_metaboscan sheet as temporary file
                        risk_params_path = os.path.join(temp_dir, "risk_params.xlsx")
                        st.session_state.edited_ref['Params_metaboscan'].to_excel(risk_params_path, index=False)

                        # Save Ref_stats sheet as temporary file
                        ref_stats_path = os.path.join(temp_dir, "Ref_stats.xlsx")
                        st.session_state.edited_ref['Ref_stats'].to_excel(ref_stats_path, index=False)

                        # Process data
                        metabolomic_data_with_ratios = calculate_metabolite_ratios(metabolomic_data)
                        metabolomic_data_with_ratios_path = os.path.join(temp_dir, "metabolomic_data.xlsx")
                        metabolomic_data_with_ratios.to_excel(metabolomic_data_with_ratios_path, index=False)
                        
                        metabolite_data = safe_parse_metabolite_data(metabolomic_data_with_ratios_path)
                        
                        # Check if input file contains multiple patients (more than 1 row after header)
                        df_metabolomic = pd.read_excel(metabolomic_data)
                        multiple_patients = len(df_metabolomic) > 1
                        
                        if multiple_patients:
                            st.info("Обнаружены данные для нескольких пациентов. Показаны результаты для всех пациентов.")
                            st.warning("Для генерации индивидуальных отчетов, пожалуйста, загружайте данные по одному пациенту за раз.")
                            
                            # Get patient identifiers and groups from file
                            patient_ids = df_metabolomic.get('Код', [f"Пациент {i+1}" for i in range(len(df_metabolomic))])
                            patient_groups = df_metabolomic.get('Группа', ["-" for _ in range(len(df_metabolomic))])
                            
                            # Create tabs for each patient
                            tabs = st.tabs([f"Пациент {i+1}" for i in range(len(patient_ids))])
                            
                            for idx, tab in enumerate(tabs):
                                with tab:
                                    with st.spinner(f"Расчет показателей для пациента {idx+1}/{len(patient_ids)}..."):
                                        # Get individual patient data
                                        patient_data = metabolomic_data_with_ratios.iloc[[idx]]
                                        patient_data_path = os.path.join(temp_dir, f"patient_data_{idx}.xlsx")
                                        patient_data.to_excel(patient_data_path, index=False)
                                        
                                        # Calculate risk parameters for this patient only
                                        patient_risk_params_exp = prepare_final_dataframe_zscore(risk_params_path, patient_data_path, ref_stats_path)
                                        patient_risk_params_exp_old = prepare_final_dataframe_old(risk_params_path, patient_data_path)
                                        
                                        
                                        # Calculate risk scores for this patient only
                                        patient_risk_scores = calculate_risks(patient_risk_params_exp, patient_data)
                                        patient_risk_scores_old = calculate_risks(patient_risk_params_exp_old, patient_data)
                                        # Display results
                                        
                                        
                                        st.markdown(f"**Код пациента:** {patient_ids[idx]}")
                                        st.markdown(f"**Группа:** {patient_groups[idx]}")
                                        st.markdown("---")
                                        col2, col3, col4 = st.columns([1 , 1, 1])
                                            
                                        with col2:
                                            with st.expander("Коридоры", expanded=True ):
                                                st.image(plot_metabolite_z_scores(
                                                            group_title= "Метаболизм фенилаланина",
                                                        metabolite_concentrations= {
                                                                "Phenylalanine": metabolite_data["Phenylalanine"],
                                                                "Tyrosin": metabolite_data["Tyrosin"],
                                                                "Summ Leu-Ile": metabolite_data["Summ Leu-Ile"],
                                                                "Valine": metabolite_data["Valine"],
                                                                "BCAA": metabolite_data["BCAA"],
                                                                "BCAA/AAA": metabolite_data["BCAA/AAA"],
                                                                "Phe/Tyr": metabolite_data["Phe/Tyr"],
                                                                "Val/C4": metabolite_data["Val/C4"],
                                                                "(Leu+IsL)/(C3+С5+С5-1+C5-DC)": metabolite_data[
                                                                    "(Leu+IsL)/(C3+С5+С5-1+C5-DC)"
                                                                ],
                                                            },
                                                            ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                                st.image(plot_metabolite_z_scores(
                                                            group_title= "Метаболизм гистидина",
                                                        metabolite_concentrations= {
                                                                "Histidine": metabolite_data["Histidine"],
                                                                "Methylhistidine": metabolite_data["Methylhistidine"],
                                                                "Threonine": metabolite_data["Threonine"],
                                                                "Glycine": metabolite_data["Glycine"],
                                                                "DMG": metabolite_data["DMG"],
                                                                "Serine": metabolite_data["Serine"],
                                                                "Lysine": metabolite_data["Lysine"],
                                                                "Glutamic acid": metabolite_data["Glutamic acid"],
                                                                "Glutamine/Glutamate": metabolite_data["Glutamine"],
                                                                "Glutamine/Glutamate": metabolite_data["Glutamine/Glutamate"],
                                                                "Glycine/Serine": metabolite_data["Glycine/Serine"],
                                                                "GSG Index": metabolite_data["GSG Index"],
                                                                "Carnosine": metabolite_data["Carnosine"],
                                                            },
                                                            ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                                st.image(plot_metabolite_z_scores(
                                                            group_title= "Метаболизм метионина",
                                                        metabolite_concentrations= {
                                                                "Methionine": metabolite_data["Methionine"],
                                                                "Methionine-Sulfoxide": metabolite_data["Methionine-Sulfoxide"],
                                                                "Taurine": metabolite_data["Taurine"],
                                                                "Betaine": metabolite_data["Betaine"],
                                                                "Choline": metabolite_data["Choline"],
                                                                "TMAO": metabolite_data["TMAO"],
                                                                "Betaine/choline": metabolite_data["Betaine/choline"],
                                                                "Methionine + Taurine": metabolite_data["Methionine + Taurine"],
                                                                "Met Oxidation": metabolite_data["Met Oxidation"],
                                                                "TMAO Synthesis": metabolite_data["TMAO Synthesis"],
                                                                "DMG / Choline": metabolite_data["DMG / Choline"],
                                                            },
                                                            ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                                st.image(plot_metabolite_z_scores(
                                                            group_title= "Кинурениновый путь",
                                                        metabolite_concentrations= {
                                                                "Tryptophan": metabolite_data["Tryptophan"],
                                                                "Kynurenine": metabolite_data["Kynurenine"],
                                                                "Antranillic acid": metabolite_data["Antranillic acid"],
                                                                "Quinolinic acid": metabolite_data["Quinolinic acid"],
                                                                "Xanthurenic acid": metabolite_data["Xanthurenic acid"],
                                                                "Kynurenic acid": metabolite_data["Kynurenic acid"],
                                                                "Kyn/Trp": metabolite_data["Kyn/Trp"],
                                                                "Trp/(Kyn+QA)": metabolite_data["Trp/(Kyn+QA)"],
                                                                "Kyn/Quin": metabolite_data["Kyn/Quin"],
                                                            },
                                                            ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                                st.image(plot_metabolite_z_scores(
                                                            group_title= "Серотониновый путь",
                                                        metabolite_concentrations= {
                                                                "Serotonin": metabolite_data["Serotonin"],
                                                                "HIAA": metabolite_data["HIAA"],
                                                                "5-hydroxytryptophan": metabolite_data["5-hydroxytryptophan"],
                                                                "Serotonin / Trp": metabolite_data["Serotonin / Trp"],
                                                            },
                                                            ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                                st.image(plot_metabolite_z_scores(
                                                            group_title= "Индоловый путь",
                                                        metabolite_concentrations= {
                                                                "Indole-3-acetic acid": metabolite_data["Indole-3-acetic acid"],
                                                                "Indole-3-lactic acid": metabolite_data["Indole-3-lactic acid"],
                                                                "Indole-3-carboxaldehyde": metabolite_data[
                                                                    "Indole-3-carboxaldehyde"
                                                                ],
                                                                "Indole-3-propionic acid": metabolite_data[
                                                                    "Indole-3-propionic acid"
                                                                ],
                                                                "Indole-3-butyric": metabolite_data["Indole-3-butyric"],
                                                                "Tryptamine": metabolite_data["Tryptamine"],
                                                                "Tryptamine / IAA": metabolite_data["Tryptamine / IAA"],
                                                            },
                                                            ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                                st.image(plot_metabolite_z_scores(
                                                            group_title= "Метаболизм аргинина",
                                                        metabolite_concentrations= {
                                                                "Proline": metabolite_data["Proline"],
                                                                "Hydroxyproline": metabolite_data["Hydroxyproline"],
                                                                "ADMA": metabolite_data["ADMA"],
                                                                "NMMA": metabolite_data["NMMA"],
                                                                "TotalDMA (SDMA)": metabolite_data["TotalDMA (SDMA)"],
                                                                "Homoarginine": metabolite_data["Homoarginine"],
                                                                "Arginine": metabolite_data["Arginine"],
                                                                "Citrulline": metabolite_data["Citrulline"],
                                                                "Ornitine": metabolite_data["Ornitine"],
                                                                "Asparagine": metabolite_data["Asparagine"],
                                                                "Aspartic acid": metabolite_data["Aspartic acid"],
                                                                "Creatinine": metabolite_data["Creatinine"],
                                                                "Arg/ADMA": metabolite_data["Arg/ADMA"],
                                                                "(Arg+HomoArg)/ADMA": metabolite_data["(Arg+HomoArg)/ADMA"],
                                                                "Arg/Orn+Cit": metabolite_data["Arg/Orn+Cit"],
                                                                "ADMA/(Adenosin+Arginine)": metabolite_data[
                                                                    "ADMA/(Adenosin+Arginine)"
                                                                ],
                                                                "Symmetrical Arg Methylation": metabolite_data[
                                                                    "Symmetrical Arg Methylation"
                                                                ],
                                                                "Sum of Dimethylated Arg": metabolite_data[
                                                                    "Sum of Dimethylated Arg"
                                                                ],
                                                                "Ratio of Pro to Cit": metabolite_data["Ratio of Pro to Cit"],
                                                                "Cit Synthesis": metabolite_data["Cit Synthesis"],
                                                            },
                                                            ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                                st.image(plot_metabolite_z_scores(
                                                            group_title= "Метаболизм ацилкарнитинов (соотношения)",
                                                        metabolite_concentrations= {
                                                                "Alanine": metabolite_data["Alanine"],
                                                                "C0": metabolite_data["C0"],
                                                                "Ratio of AC-OHs to ACs": metabolite_data["Ratio of AC-OHs to ACs"],
                                                                "СДК": metabolite_data["СДК"],
                                                                "ССК": metabolite_data["ССК"],
                                                                "СКК": metabolite_data["СКК"],
                                                                "C0/(C16+C18)": metabolite_data["C0/(C16+C18)"],
                                                                "CPT-2 Deficiency (NBS)": metabolite_data["CPT-2 Deficiency (NBS)"],
                                                                "С2/С0": metabolite_data["С2/С0"],
                                                                "Ratio of Short-Chain to Long-Chain ACs": metabolite_data[
                                                                    "Ratio of Short-Chain to Long-Chain ACs"
                                                                ],
                                                                "Ratio of Medium-Chain to Long-Chain ACs": metabolite_data[
                                                                    "Ratio of Medium-Chain to Long-Chain ACs"
                                                                ],
                                                                "Ratio of Short-Chain to Medium-Chain ACs": metabolite_data[
                                                                    "Ratio of Short-Chain to Medium-Chain ACs"
                                                                ],
                                                                "Sum of ACs": metabolite_data["Sum of ACs"],
                                                                "Sum of ACs + С0": metabolite_data["Sum of ACs + С0"],
                                                                "Sum of ACs/C0": metabolite_data["Sum of ACs/C0"],
                                                            },
                                                            ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                                st.image(plot_metabolite_z_scores(
                                                            group_title= "Короткоцепочечные ацилкарнитины",
                                                        metabolite_concentrations= {
                                                                "C2": metabolite_data["C2"],
                                                                "C3": metabolite_data["C3"],
                                                                "C4": metabolite_data["C4"],
                                                                "C5": metabolite_data["C5"],
                                                                "C5-1": metabolite_data["C5-1"],
                                                                "C5-DC": metabolite_data["C5-DC"],
                                                                "C5-OH": metabolite_data["C5-OH"],
                                                            },
                                                            ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                                st.image(plot_metabolite_z_scores(
                                                            group_title= "Среднецепочечные ацилкарнитины",
                                                        metabolite_concentrations= {
                                                                "C6": metabolite_data["C6"],
                                                                "C6-DC": metabolite_data["C6-DC"],
                                                                "C8": metabolite_data["C8"],
                                                                "C8-1": metabolite_data["C8-1"],
                                                                "C10": metabolite_data["C10"],
                                                                "C10-1": metabolite_data["C10-1"],
                                                                "C10-2": metabolite_data["C10-2"],
                                                                "C12": metabolite_data["C12"],
                                                                "C12-1": metabolite_data["C12-1"],
                                                            },
                                                            ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                                st.image(plot_metabolite_z_scores(
                                                            group_title= "Длинноцепочечные ацилкарнитины",
                                                        metabolite_concentrations= {
                                                                "C14": metabolite_data["C14"],
                                                                "C14-1": metabolite_data["C14-1"],
                                                                "C14-2": metabolite_data["C14-2"],
                                                                "C14-OH": metabolite_data["C14-OH"],
                                                                "C16": metabolite_data["C16"],
                                                                "C16-1": metabolite_data["C16-1"],
                                                                "C16-1-OH": metabolite_data["C16-1-OH"],
                                                                "C16-OH": metabolite_data["C16-OH"],
                                                                "C18": metabolite_data["C18"],
                                                                "C18-1": metabolite_data["C18-1"],
                                                                "C18-1-OH": metabolite_data["C18-1-OH"],
                                                                "C18-2": metabolite_data["C18-2"],
                                                                "C18-OH": metabolite_data["C18-OH"],
                                                            },
                                                            ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                                st.image(plot_metabolite_z_scores(
                                                            group_title= "Другие метаболиты",
                                                        metabolite_concentrations= {
                                                                "Pantothenic": metabolite_data["Pantothenic"],
                                                                "Riboflavin": metabolite_data["Riboflavin"],
                                                                "Melatonin": metabolite_data["Melatonin"],
                                                                "Uridine": metabolite_data["Uridine"],
                                                                "Adenosin": metabolite_data["Adenosin"],
                                                                "Cytidine": metabolite_data["Cytidine"],
                                                                "Cortisol": metabolite_data["Cortisol"],
                                                                "Histamine": metabolite_data["Histamine"],
                                                            },
                                                            ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                                    
                                        with col3:
                                            st.markdown("**Cтарый метод:**")
                                            st.dataframe(patient_risk_scores_old.sort_values(by="Метод оценки", ascending=True), hide_index=True,column_order=('Риск-скор', 'Группа риска', 'Метод оценки'))
                                            with st.expander("Показатели по группам:", expanded=True):
                                                display_group_cards(patient_risk_params_exp_old, patient_risk_scores_old)
                                        
                                        with col4:
                                            # Display individual risk scores
                                            st.markdown("**Z-scores:**")
                                            st.dataframe(patient_risk_scores.sort_values(by="Метод оценки", ascending=True), hide_index=True,column_order=('Риск-скор', 'Группа риска', 'Метод оценки'))
                                            with st.expander("Показатели по группам:", expanded=True):
                                                display_group_cards(patient_risk_params_exp, patient_risk_scores)
                                        
                                        

                        else:  # Single patient case (original behavior)
                            risk_params_exp_zscore = prepare_final_dataframe_zscore(risk_params_path, metabolomic_data_with_ratios_path, ref_stats_path)
                            risk_params_exp_old = prepare_final_dataframe_old(risk_params_path, metabolomic_data_with_ratios_path)
                            risk_params_exp_path = os.path.join(temp_dir, "risk_exp_params.xlsx")
                            risk_params_exp_old_path = os.path.join(temp_dir, "risk_exp_params_old.xlsx")
                            risk_params_exp_zscore.to_excel(risk_params_exp_path, index=False)
                            risk_params_exp_old.to_excel(risk_params_exp_old_path, index=False)
                                
                            risk_scores = calculate_risks(risk_params_exp_zscore, metabolomic_data_with_ratios)
                            risk_scores_path = os.path.join(temp_dir, "risk_scores.xlsx")
                            risk_scores.to_excel(risk_scores_path, index=False)
                            
                            risk_scores_old = calculate_risks(risk_params_exp_old, metabolomic_data_with_ratios)
                            risk_scores_old_path = os.path.join(temp_dir, "risk_scores_old.xlsx")
                            risk_scores_old.to_excel(risk_scores_old_path, index=False)
                            
                            metrics_path = os.path.join(temp_dir, "metrics.xlsx")
                            st.session_state.edited_ref['metrics_ml_models'].to_excel(metrics_path, index=False)
                            st.info("✅ Предварительный просмотр рассчитанных значений!")
                            cols = st.columns(3)
                            with cols[0]:
                                with st.expander("Коридоры", expanded=True ):
                                    st.image(plot_metabolite_z_scores(
                                                group_title= "Метаболизм фенилаланина",
                                               metabolite_concentrations= {
                                                    "Phenylalanine": metabolite_data["Phenylalanine"],
                                                    "Tyrosin": metabolite_data["Tyrosin"],
                                                    "Summ Leu-Ile": metabolite_data["Summ Leu-Ile"],
                                                    "Valine": metabolite_data["Valine"],
                                                    "BCAA": metabolite_data["BCAA"],
                                                    "BCAA/AAA": metabolite_data["BCAA/AAA"],
                                                    "Phe/Tyr": metabolite_data["Phe/Tyr"],
                                                    "Val/C4": metabolite_data["Val/C4"],
                                                    "(Leu+IsL)/(C3+С5+С5-1+C5-DC)": metabolite_data[
                                                        "(Leu+IsL)/(C3+С5+С5-1+C5-DC)"
                                                    ],
                                                },
                                                ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                    st.image(plot_metabolite_z_scores(
                                                group_title= "Метаболизм гистидина",
                                               metabolite_concentrations= {
                                                    "Histidine": metabolite_data["Histidine"],
                                                    "Methylhistidine": metabolite_data["Methylhistidine"],
                                                    "Threonine": metabolite_data["Threonine"],
                                                    "Glycine": metabolite_data["Glycine"],
                                                    "DMG": metabolite_data["DMG"],
                                                    "Serine": metabolite_data["Serine"],
                                                    "Lysine": metabolite_data["Lysine"],
                                                    "Glutamic acid": metabolite_data["Glutamic acid"],
                                                    "Glutamine/Glutamate": metabolite_data["Glutamine"],
                                                    "Glutamine/Glutamate": metabolite_data["Glutamine/Glutamate"],
                                                    "Glycine/Serine": metabolite_data["Glycine/Serine"],
                                                    "GSG Index": metabolite_data["GSG Index"],
                                                    "Carnosine": metabolite_data["Carnosine"],
                                                },
                                                ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                    st.image(plot_metabolite_z_scores(
                                                group_title= "Метаболизм метионина",
                                               metabolite_concentrations= {
                                                    "Methionine": metabolite_data["Methionine"],
                                                    "Methionine-Sulfoxide": metabolite_data["Methionine-Sulfoxide"],
                                                    "Taurine": metabolite_data["Taurine"],
                                                    "Betaine": metabolite_data["Betaine"],
                                                    "Choline": metabolite_data["Choline"],
                                                    "TMAO": metabolite_data["TMAO"],
                                                    "Betaine/choline": metabolite_data["Betaine/choline"],
                                                    "Methionine + Taurine": metabolite_data["Methionine + Taurine"],
                                                    "Met Oxidation": metabolite_data["Met Oxidation"],
                                                    "TMAO Synthesis": metabolite_data["TMAO Synthesis"],
                                                    "DMG / Choline": metabolite_data["DMG / Choline"],
                                                },
                                                ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                    st.image(plot_metabolite_z_scores(
                                                group_title= "Кинурениновый путь",
                                               metabolite_concentrations= {
                                                    "Tryptophan": metabolite_data["Tryptophan"],
                                                    "Kynurenine": metabolite_data["Kynurenine"],
                                                    "Antranillic acid": metabolite_data["Antranillic acid"],
                                                    "Quinolinic acid": metabolite_data["Quinolinic acid"],
                                                    "Xanthurenic acid": metabolite_data["Xanthurenic acid"],
                                                    "Kynurenic acid": metabolite_data["Kynurenic acid"],
                                                    "Kyn/Trp": metabolite_data["Kyn/Trp"],
                                                    "Trp/(Kyn+QA)": metabolite_data["Trp/(Kyn+QA)"],
                                                    "Kyn/Quin": metabolite_data["Kyn/Quin"],
                                                },
                                                ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                    st.image(plot_metabolite_z_scores(
                                                group_title= "Серотониновый путь",
                                               metabolite_concentrations= {
                                                    "Serotonin": metabolite_data["Serotonin"],
                                                    "HIAA": metabolite_data["HIAA"],
                                                    "5-hydroxytryptophan": metabolite_data["5-hydroxytryptophan"],
                                                    "Serotonin / Trp": metabolite_data["Serotonin / Trp"],
                                                },
                                                ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                    st.image(plot_metabolite_z_scores(
                                                group_title= "Индоловый путь",
                                               metabolite_concentrations= {
                                                    "Indole-3-acetic acid": metabolite_data["Indole-3-acetic acid"],
                                                    "Indole-3-lactic acid": metabolite_data["Indole-3-lactic acid"],
                                                    "Indole-3-carboxaldehyde": metabolite_data[
                                                        "Indole-3-carboxaldehyde"
                                                    ],
                                                    "Indole-3-propionic acid": metabolite_data[
                                                        "Indole-3-propionic acid"
                                                    ],
                                                    "Indole-3-butyric": metabolite_data["Indole-3-butyric"],
                                                    "Tryptamine": metabolite_data["Tryptamine"],
                                                    "Tryptamine / IAA": metabolite_data["Tryptamine / IAA"],
                                                },
                                                ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                    st.image(plot_metabolite_z_scores(
                                                group_title= "Метаболизм аргинина",
                                               metabolite_concentrations= {
                                                    "Proline": metabolite_data["Proline"],
                                                    "Hydroxyproline": metabolite_data["Hydroxyproline"],
                                                    "ADMA": metabolite_data["ADMA"],
                                                    "NMMA": metabolite_data["NMMA"],
                                                    "TotalDMA (SDMA)": metabolite_data["TotalDMA (SDMA)"],
                                                    "Homoarginine": metabolite_data["Homoarginine"],
                                                    "Arginine": metabolite_data["Arginine"],
                                                    "Citrulline": metabolite_data["Citrulline"],
                                                    "Ornitine": metabolite_data["Ornitine"],
                                                    "Asparagine": metabolite_data["Asparagine"],
                                                    "Aspartic acid": metabolite_data["Aspartic acid"],
                                                    "Creatinine": metabolite_data["Creatinine"],
                                                    "Arg/ADMA": metabolite_data["Arg/ADMA"],
                                                    "(Arg+HomoArg)/ADMA": metabolite_data["(Arg+HomoArg)/ADMA"],
                                                    "Arg/Orn+Cit": metabolite_data["Arg/Orn+Cit"],
                                                    "ADMA/(Adenosin+Arginine)": metabolite_data[
                                                        "ADMA/(Adenosin+Arginine)"
                                                    ],
                                                    "Symmetrical Arg Methylation": metabolite_data[
                                                        "Symmetrical Arg Methylation"
                                                    ],
                                                    "Sum of Dimethylated Arg": metabolite_data[
                                                        "Sum of Dimethylated Arg"
                                                    ],
                                                    "Ratio of Pro to Cit": metabolite_data["Ratio of Pro to Cit"],
                                                    "Cit Synthesis": metabolite_data["Cit Synthesis"],
                                                },
                                                ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                    st.image(plot_metabolite_z_scores(
                                                group_title= "Метаболизм ацилкарнитинов (соотношения)",
                                               metabolite_concentrations= {
                                                    "Alanine": metabolite_data["Alanine"],
                                                    "C0": metabolite_data["C0"],
                                                    "Ratio of AC-OHs to ACs": metabolite_data["Ratio of AC-OHs to ACs"],
                                                    "СДК": metabolite_data["СДК"],
                                                    "ССК": metabolite_data["ССК"],
                                                    "СКК": metabolite_data["СКК"],
                                                    "C0/(C16+C18)": metabolite_data["C0/(C16+C18)"],
                                                    "CPT-2 Deficiency (NBS)": metabolite_data["CPT-2 Deficiency (NBS)"],
                                                    "С2/С0": metabolite_data["С2/С0"],
                                                    "Ratio of Short-Chain to Long-Chain ACs": metabolite_data[
                                                        "Ratio of Short-Chain to Long-Chain ACs"
                                                    ],
                                                    "Ratio of Medium-Chain to Long-Chain ACs": metabolite_data[
                                                        "Ratio of Medium-Chain to Long-Chain ACs"
                                                    ],
                                                    "Ratio of Short-Chain to Medium-Chain ACs": metabolite_data[
                                                        "Ratio of Short-Chain to Medium-Chain ACs"
                                                    ],
                                                    "Sum of ACs": metabolite_data["Sum of ACs"],
                                                    "Sum of ACs + С0": metabolite_data["Sum of ACs + С0"],
                                                    "Sum of ACs/C0": metabolite_data["Sum of ACs/C0"],
                                                },
                                                ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                    st.image(plot_metabolite_z_scores(
                                                group_title= "Короткоцепочечные ацилкарнитины",
                                               metabolite_concentrations= {
                                                    "C2": metabolite_data["C2"],
                                                    "C3": metabolite_data["C3"],
                                                    "C4": metabolite_data["C4"],
                                                    "C5": metabolite_data["C5"],
                                                    "C5-1": metabolite_data["C5-1"],
                                                    "C5-DC": metabolite_data["C5-DC"],
                                                    "C5-OH": metabolite_data["C5-OH"],
                                                },
                                                ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                    st.image(plot_metabolite_z_scores(
                                                group_title= "Среднецепочечные ацилкарнитины",
                                               metabolite_concentrations= {
                                                    "C6": metabolite_data["C6"],
                                                    "C6-DC": metabolite_data["C6-DC"],
                                                    "C8": metabolite_data["C8"],
                                                    "C8-1": metabolite_data["C8-1"],
                                                    "C10": metabolite_data["C10"],
                                                    "C10-1": metabolite_data["C10-1"],
                                                    "C10-2": metabolite_data["C10-2"],
                                                    "C12": metabolite_data["C12"],
                                                    "C12-1": metabolite_data["C12-1"],
                                                },
                                                ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                    st.image(plot_metabolite_z_scores(
                                                group_title= "Длинноцепочечные ацилкарнитины",
                                               metabolite_concentrations= {
                                                    "C14": metabolite_data["C14"],
                                                    "C14-1": metabolite_data["C14-1"],
                                                    "C14-2": metabolite_data["C14-2"],
                                                    "C14-OH": metabolite_data["C14-OH"],
                                                    "C16": metabolite_data["C16"],
                                                    "C16-1": metabolite_data["C16-1"],
                                                    "C16-1-OH": metabolite_data["C16-1-OH"],
                                                    "C16-OH": metabolite_data["C16-OH"],
                                                    "C18": metabolite_data["C18"],
                                                    "C18-1": metabolite_data["C18-1"],
                                                    "C18-1-OH": metabolite_data["C18-1-OH"],
                                                    "C18-2": metabolite_data["C18-2"],
                                                    "C18-OH": metabolite_data["C18-OH"],
                                                },
                                                ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                                    st.image(plot_metabolite_z_scores(
                                                group_title= "Другие метаболиты",
                                               metabolite_concentrations= {
                                                    "Pantothenic": metabolite_data["Pantothenic"],
                                                    "Riboflavin": metabolite_data["Riboflavin"],
                                                    "Melatonin": metabolite_data["Melatonin"],
                                                    "Uridine": metabolite_data["Uridine"],
                                                    "Adenosin": metabolite_data["Adenosin"],
                                                    "Cytidine": metabolite_data["Cytidine"],
                                                    "Cortisol": metabolite_data["Cortisol"],
                                                    "Histamine": metabolite_data["Histamine"],
                                                },
                                                ref_stats=create_ref_stats_from_excel(ref_stats_path)))
                            with cols[1]:
                                st.header("Старые риски:")
                                st.dataframe(risk_scores_old.sort_values(by="Метод оценки", ascending=True), hide_index=True,column_order=('Риск-скор', 'Группа риска', 'Метод оценки'))
                                with st.expander("Показатели по группам:", expanded=True):
                                    display_group_cards(risk_params_exp_old, risk_scores_old)
                                
                                
                            with cols[2]:
                                st.header("Z-score:")
                                st.dataframe(risk_scores.sort_values(by="Метод оценки", ascending=True), hide_index=True,column_order=('Риск-скор', 'Группа риска', 'Метод оценки'))
                                with st.expander("Показатели по группам:", expanded=True):
                                    display_group_cards(risk_params_exp_zscore, risk_scores)
                                
                    except Exception as e:
                        st.error(f"An error occurred: {str(e)}")
                        logging.error(f"Error in report generation: {str(e)}")

if __name__ == "__main__":
    main()