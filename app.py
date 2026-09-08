import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import requests
import plotly.express as px
import plotly.graph_objects as go

# Настройка страницы
st.set_page_config(
    page_title="Рейтинг ЖД",
    page_icon="🚂",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Стилизация
st.markdown("""
    <style>
    .main-header {
        font-size: 3rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #2c3e50;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-header">🚂 Рейтинг ЖД</p>', unsafe_allow_html=True)

# Боковая панель
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/train.png", width=80)
    st.markdown("## 📋 Настройки")
    
    # Загрузка данных
    @st.cache_data
    def load_data_from_github():
        url = "https://raw.githubusercontent.com/aidarpavl/Reiting_ZHD/main/reiting.xlsx"
        try:
            response = requests.get(url)
            response.raise_for_status()
            df = pd.read_excel(BytesIO(response.content))
            return df
        except Exception as e:
            st.error(f"❌ Ошибка загрузки данных: {e}")
            return None
    
    # Загрузка данных
    with st.spinner("Загрузка данных..."):
        df = load_data_from_github()
    
    if df is not None:
        st.success(f"✅ Загружено {len(df)} записей")
        st.info(f"📊 Колонки: {', '.join(df.columns[:5])}...")
    
    st.markdown("---")
    st.caption("📌 Данные загружены из GitHub репозитория")

# Основной контент
if df is not None:
    # Вкладки
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Данные", "📈 Рейтинг", "📊 Визуализация", "📥 Экспорт"])
    
    with tab1:
        st.markdown('<p class="sub-header">📊 Исходные данные</p>', unsafe_allow_html=True)
        
        # Фильтры
        col1, col2 = st.columns(2)
        with col1:
            search = st.text_input("🔍 Поиск", placeholder="Введите текст для поиска...")
        
        with col2:
            show_rows = st.selectbox("Показать строк", [10, 25, 50, 100, 1000], index=0)
        
        # Фильтрация данных
        df_filtered = df.copy()
        if search:
            mask = df_filtered.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
            df_filtered = df_filtered[mask]
        
        st.dataframe(df_filtered.head(show_rows), use_container_width=True)
        st.caption(f"Показано {min(show_rows, len(df_filtered))} из {len(df_filtered)} записей")
    
    with tab2:
        st.markdown('<p class="sub-header">📈 Расчет рейтинга</p>', unsafe_allow_html=True)
        
        # Копия данных для расчетов
        df_calculated = df.copy()
        
        # Определяем числовые колонки
        numeric_cols = df_calculated.select_dtypes(include=[np.number]).columns.tolist()
        
        # Исключаем ID если есть
        if 'ID' in numeric_cols:
            numeric_cols.remove('ID')
        
        # Параметры расчета
        with st.expander("⚙️ Параметры расчета", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                method = st.selectbox(
                    "Метод нормализации",
                    ["Min-Max", "Z-Score", "Robust"],
                    help="Min-Max: масштабирование от 0 до 1\nZ-Score: стандартизация\nRobust: устойчивый к выбросам"
                )
            
            with col2:
                weight_type = st.selectbox(
                    "Вес показателей",
                    ["Равные веса", "Экспертные веса"],
                    help="Равные веса: все показатели одинаково важны\nЭкспертные веса: можно настроить важность"
                )
        
        # Нормализация
        st.write("**Нормализация данных:**")
        
        for col in numeric_cols:
            if method == "Min-Max":
                if df_calculated[col].max() != df_calculated[col].min():
                    df_calculated[f'{col}_norm'] = (df_calculated[col] - df_calculated[col].min()) / (df_calculated[col].max() - df_calculated[col].min())
                else:
                    df_calculated[f'{col}_norm'] = 0.5
            elif method == "Z-Score":
                df_calculated[f'{col}_norm'] = (df_calculated[col] - df_calculated[col].mean()) / df_calculated[col].std()
            elif method == "Robust":
                q1 = df_calculated[col].quantile(0.25)
                q3 = df_calculated[col].quantile(0.75)
                iqr = q3 - q1
                if iqr != 0:
                    df_calculated[f'{col}_norm'] = (df_calculated[col] - df_calculated[col].median()) / iqr
                else:
                    df_calculated[f'{col}_norm'] = 0.5
        
        # Расчет рейтинга
        norm_cols = [f'{col}_norm' for col in numeric_cols]
        
        if weight_type == "Экспертные веса":
            st.write("**Настройка весов показателей:**")
            weights = {}
            cols_per_row = 3
            for i in range(0, len(numeric_cols), cols_per_row):
                cols = st.columns(cols_per_row)
                for j, col in enumerate(numeric_cols[i:i+cols_per_row]):
                    with cols[j]:
                        weights[col] = st.slider(
                            f"{col}",
                            min_value=0.0,
                            max_value=1.0,
                            value=1.0,
                            step=0.1
                        )
            
            # Нормализация весов
            total_weight = sum(weights.values())
            if total_weight > 0:
                for col in numeric_cols:
                    weights[col] = weights[col] / total_weight
            else:
                weights = {col: 1/len(numeric_cols) for col in numeric_cols}
        else:
            weights = {col: 1/len(numeric_cols) for col in numeric_cols}
        
        # Расчет взвешенного рейтинга
        df_calculated['Рейтинг'] = 0
        for col in numeric_cols:
            df_calculated['Рейтинг'] += df_calculated[f'{col}_norm'] * weights[col]
        
        df_calculated['Рейтинг (%)'] = df_calculated['Рейтинг'] * 100
        
        # Сортировка
        df_sorted = df_calculated.sort_values('Рейтинг', ascending=False).reset_index(drop=True)
        df_sorted['Место'] = df_sorted.index + 1
        
        # Отображение результатов
        col1, col2 = st.columns([3, 1])
        
        with col1:
            st.write("**Результаты расчета:**")
            result_cols = ['Место'] + [col for col in df_sorted.columns if col not in norm_cols and '_norm' not in col]
            st.dataframe(df_sorted[result_cols], use_container_width=True)
        
        with col2:
            st.markdown("**📊 Статистика:**")
            st.metric("🏆 Лидер", df_sorted.iloc[0].get('Наименование', 'N/A') if len(df_sorted) > 0 else 'N/A', 
                     delta=f"Рейтинг {df_sorted.iloc[0]['Рейтинг (%)']:.1f}%" if len(df_sorted) > 0 else 'N/A')
            st.metric("📈 Средний рейтинг", f"{df_sorted['Рейтинг (%)'].mean():.2f}%")
            st.metric("📊 Медиана", f"{df_sorted['Рейтинг (%)'].median():.2f}%")
            st.metric("📉 Минимальный", f"{df_sorted['Рейтинг (%)'].min():.2f}%")
        
        # Топ-10
        st.markdown("**🏆 Топ-10 лучших:**")
        top10 = df_sorted[['Место'] + [col for col in df_sorted.columns if col not in norm_cols and '_norm' not in col]].head(10)
        st.dataframe(top10, use_container_width=True)
    
    with tab3:
        st.markdown('<p class="sub-header">📊 Визуализация данных</p>', unsafe_allow_html=True)
        
        # Выбор типа графика
        chart_type = st.selectbox(
            "Выберите тип визуализации:",
            ["Столбчатая диаграмма", "Круговая диаграмма", "Точечная диаграмма", "Тепловая карта"]
        )
        
        # Выбор колонок для визуализации
        chart_cols = [col for col in df_sorted.columns if col not in ['Место', 'Рейтинг', 'Рейтинг (%)'] + norm_cols]
        chart_cols = [col for col in chart_cols if '_norm' not in col]
        
        if chart_type == "Столбчатая диаграмма":
            col1, col2 = st.columns(2)
            with col1:
                x_axis = st.selectbox("X-ось", chart_cols, index=0)
            with col2:
                y_axis = st.selectbox("Y-ось", ['Рейтинг (%)'] + chart_cols, index=0)
            
            if x_axis and y_axis:
                fig = px.bar(
                    df_sorted.head(20),
                    x=x_axis,
                    y=y_axis,
                    color='Рейтинг (%)' if y_axis != 'Рейтинг (%)' else None,
                    title=f"{y_axis} по {x_axis}",
                    color_continuous_scale='Blues',
                    text_auto=True
                )
                st.plotly_chart(fig, use_container_width=True)
        
        elif chart_type == "Круговая диаграмма":
            col1, col2 = st.columns(2)
            with col1:
                pie_col = st.selectbox("Показатель для круговой диаграммы:", ['Рейтинг (%)'] + chart_cols, index=0)
            with col2:
                num_items = st.slider("Количество элементов:", 3, 20, 10)
            
            if pie_col:
                pie_data = df_sorted.head(num_items).copy()
                fig = px.pie(
                    pie_data,
                    values=pie_col,
                    names=pie_data.index + 1,
                    title=f"Распределение {pie_col}",
                    hover_data=[pie_col]
                )
                st.plotly_chart(fig, use_container_width=True)
        
        elif chart_type == "Точечная диаграмма":
            col1, col2 = st.columns(2)
            with col1:
                x_scatter = st.selectbox("X-ось", chart_cols, index=0)
            with col2:
                y_scatter = st.selectbox("Y-ось", ['Рейтинг (%)'] + chart_cols, index=0)
            
            if x_scatter and y_scatter:
                fig = px.scatter(
                    df_sorted,
                    x=x_scatter,
                    y=y_scatter,
                    color='Рейтинг (%)' if y_scatter != 'Рейтинг (%)' else None,
                    size='Рейтинг (%)' if y_scatter != 'Рейтинг (%)' else None,
                    title=f"Зависимость {y_scatter} от {x_scatter}",
                    color_continuous_scale='Blues',
                    hover_data=['Рейтинг (%)'] + chart_cols[:2]
                )
                st.plotly_chart(fig, use_container_width=True)
        
        elif chart_type == "Тепловая карта":
            # Корреляционная матрица
            corr_matrix = df_sorted[numeric_cols + ['Рейтинг']].corr()
            
            fig = px.imshow(
                corr_matrix,
                text_auto=True,
                color_continuous_scale='RdBu_r',
                title="Корреляционная матрица показателей"
            )
            st.plotly_chart(fig, use_container_width=True)
    
    with tab4:
        st.markdown('<p class="sub-header">📥 Экспорт результатов</p>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Экспорт в Excel:**")
            
            @st.cache_data
            def convert_df_to_excel(df_sorted):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # Основной лист с рейтингом
                    df_sorted.to_excel(writer, sheet_name='Рейтинг', index=False)
                    
                    # Дополнительный лист с нормализованными данными
                    norm_data = df_sorted[[col for col in df_sorted.columns if '_norm' in col or col in ['Место', 'Рейтинг', 'Рейтинг (%)']]]
                    norm_data.to_excel(writer, sheet_name='Нормализованные данные', index=False)
                    
                    # Статистика
                    stats = pd.DataFrame({
                        'Показатель': ['Средний рейтинг', 'Максимальный рейтинг', 'Минимальный рейтинг', 'Количество записей'],
                        'Значение': [
                            df_sorted['Рейтинг (%)'].mean(),
                            df_sorted['Рейтинг (%)'].max(),
                            df_sorted['Рейтинг (%)'].min(),
                            len(df_sorted)
                        ]
                    })
                    stats.to_excel(writer, sheet_name='Статистика', index=False)
                    
                    # Форматирование
                    workbook = writer.book
                    
                    # Форматирование для колонки рейтинга
                    if 'Рейтинг (%)' in df_sorted.columns:
                        worksheet = writer.sheets['Рейтинг']
                        col_num = df_sorted.columns.get_loc('Рейтинг (%)')
                        worksheet.conditional_format(1, col_num, len(df_sorted), col_num, {
                            'type': 'data_bar',
                            'bar_color': '#2E86C1'
                        })
                
                return output.getvalue()
            
            excel_data = convert_df_to_excel(df_sorted)
            
            st.download_button(
                label="📥 Скачать Excel файл",
                data=excel_data,
                file_name="reiting_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col2:
            st.write("**Экспорт в CSV:**")
            
            @st.cache_data
            def convert_df_to_csv(df_sorted):
                return df_sorted.to_csv(index=False).encode('utf-8')
            
            csv_data = convert_df_to_csv(df_sorted)
            
            st.download_button(
                label="📥 Скачать CSV файл",
                data=csv_data,
                file_name="reiting_result.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        # Превью экспорта
        st.markdown("**📋 Превью экспортируемых данных:**")
        st.dataframe(df_sorted.head(10), use_container_width=True)

else:
    st.error("❌ Не удалось загрузить данные. Проверьте доступность файла.")
    st.info("""
    **Возможные причины:**
    - Файл reiting.xlsx отсутствует в репозитории
    - Проблемы с доступом к GitHub
    - Некорректный формат файла
    
    **Решение:**
    Проверьте, что файл существует по ссылке:
    https://raw.githubusercontent.com/aidarpavl/Reiting_ZHD/main/reiting.xlsx
    """)

# Footer
st.markdown("---")
col1, col2, col3 = st.columns(3)
with col1:
    st.caption("📌 Данные загружены из GitHub репозитория")
with col2:
    st.caption("🚂 Приложение создано на Streamlit")
with col3:
    st.caption(f"🔄 Обновлено: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}")