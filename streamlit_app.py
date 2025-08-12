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
                                        col1, col2, col3 = st.columns([1, 2, 2])
                                        
                                        with col1:
                                            st.markdown(f"**Код пациента:** {patient_ids[idx]}")
                                            st.markdown(f"**Группа:** {patient_groups[idx]}")
                                            st.markdown("---")
                                        
                                        with col2:
                                            st.markdown("**Cтарый метод:**")
                                            with st.expander("Как выглядит в отчете:"):
                                                display_group_cards(patient_risk_params_exp_old, patient_risk_scores_old)
                                            st.dataframe(
                                                patient_risk_scores_old,
                                                use_container_width=True
                                            )
                                            st.dataframe(patient_risk_params_exp_old[['Категория', 'Subgroup_score']])
                                        
                                        with col3:
                                            # Display individual risk scores
                                            st.markdown("**Z-scores:**")
                                            with st.expander("Как выглядит в отчете:"):
                                                display_group_cards(patient_risk_params_exp, patient_risk_scores)
                                            st.dataframe( 
                                                patient_risk_scores,
                                                use_container_width=True
                                            )
                                            st.dataframe(
                                                patient_risk_params_exp,
                                                use_container_width=True
                                            )
                                        

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
                            cols = st.columns(2)
                            with cols[0]:
                                st.header("Старые риски:")
                                with st.expander("Как выглядит в отчете:"):
                                    display_group_cards(risk_params_exp_old, risk_scores_old)
                                st.dataframe(risk_scores_old)
                                st.dataframe(risk_params_exp_old[['Категория', 'Subgroup_score']])
                                
                            with cols[1]:
                                st.header("Z-score:")
                                with st.expander("Как выглядит в отчете:"):
                                    display_group_cards(risk_params_exp_zscore, risk_scores)
                                st.dataframe(risk_scores)
                                st.dataframe(risk_params_exp_zscore[['Категория','Z_score', 'Subgroup_score']])
                                
                    except Exception as e:
                        st.error(f"An error occurred: {str(e)}")
                        logging.error(f"Error in report generation: {str(e)}")

if __name__ == "__main__":
    main()