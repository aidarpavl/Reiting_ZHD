import streamlit as st
import pandas as pd
import requests
import base64
from io import BytesIO

# ==================== GITHUB ПАРАМЕТРЛЕРІ ====================
GITHUB_RAW_URL = "https://raw.githubusercontent.com/aidarpavl/reiting/refs/heads/main/reiting1.xlsx"
GITHUB_API_URL = "https://api.github.com/repos/aidarpavl/reiting/contents/reiting1.xlsx"

# ==================== БЕТ КОНФИГУРАЦИЯСЫ ====================
st.set_page_config(
    page_title="Мұғалімдер рейтингі",
    page_icon="📊",
    layout="wide"
)

# ==================== СЕССИЯ КҮЙІН ТАЗАРТУ ====================
# Сессия күйін қауіпсіз инициализациялау
if "df" not in st.session_state:
    st.session_state.df = None
if "github_token" not in st.session_state:
    st.session_state.github_token = None
if "data_loaded" not in st.session_state:
    st.session_state.data_loaded = False

# ==================== БАҒАН АТАУЛАРЫНЫҢ СӘЙКЕСТІГІ ====================
COLUMN_MAPPING = {
    'ФИО': ['фио', 'ф.и.о', 'фамилия', 'фио учителя', 'учитель'],
    'Предмет': ['предмет', 'пән', 'предметы', 'пәні'],
    'Коэф': ['коэф', 'коэффициент', 'коэф.', 'коэфициент', 'k'],
    'Ср.балл': ['ср балл', 'ср.балл', 'средний балл', 'орташа балл', 'средний'],
    'Үштік': ['үштік', 'троек', 'кол-во итоговых троек', 'итоговые тройки', 'тройки'],
    'Внеклас.': ['внеклас', 'внекласная', 'внекл', 'итог внекл'],
    'Метод.': ['метод', 'методическая', 'итог метод'],
    'Админ.рейтинг': ['админ', 'рейтинг', 'итог у администрации'],
}

def find_column_index(headers, target_variants):
    """Баған атауын іздеу (регистрге қарамай)"""
    for i, header in enumerate(headers):
        header_str = str(header).strip().lower()
        for variant in target_variants:
            if variant.lower() in header_str or header_str == variant.lower():
                return i
    return None

def load_from_github():
    """GitHub-тан Excel файлын жүктеу"""
    try:
        with st.spinner("📥 GitHub-тан жүктелуде..."):
            # 1. Файлды жүктеу
            response = requests.get(GITHUB_RAW_URL, timeout=30)
            response.raise_for_status()
            
            # 2. Excel оқу
            excel_data = BytesIO(response.content)
            df_raw = pd.read_excel(excel_data, engine='openpyxl', header=None)
            
            if len(df_raw) < 2:
                st.error("Файлда деректер жоқ!")
                return
            
            # 3. Баған атауларын алу
            headers = df_raw.iloc[0].fillna('').astype(str).tolist()
            
            st.info(f"📋 Табылған бағандар: {', '.join([str(h)[:20] for h in headers[:10]])}...")
            
            # 4. Бағандарды сәйкестендіру
            col_indices = {}
            for target, variants in COLUMN_MAPPING.items():
                idx = find_column_index(headers, variants)
                if idx is not None:
                    col_indices[target] = idx
            
            # Міндетті бағандарды тексеру
            required = ['ФИО', 'Предмет']
            missing = [r for r in required if r not in col_indices]
            if missing:
                st.error(f"Міндетті бағандар табылмады: {missing}")
                st.info("Excel файлының бірінші жолында 'ФИО' және 'Предмет' бағандары болуы керек.")
                return
            
            # 5. Деректерді оқу
            data_rows = []
            for row_idx in range(1, min(len(df_raw), 200)):  # максимум 200 жол
                row = df_raw.iloc[row_idx]
                
                # ФИО тексеру
                fio_val = row[col_indices['ФИО']]
                if pd.isna(fio_val) or str(fio_val).strip() == '':
                    continue
                
                teacher = {
                    'ФИО': str(fio_val).strip(),
                    'Предмет': str(row[col_indices['Предмет']]).strip() if 'Предмет' in col_indices else '',
                }
                
                # Коэф (егер табылса)
                if 'Коэф' in col_indices:
                    try:
                        teacher['Коэф'] = float(row[col_indices['Коэф']]) if pd.notna(row[col_indices['Коэф']]) else 1.0
                    except:
                        teacher['Коэф'] = 1.0
                else:
                    teacher['Коэф'] = 1.0
                
                # Ср.балл
                if 'Ср.балл' in col_indices:
                    try:
                        teacher['Ср.балл'] = float(row[col_indices['Ср.балл']]) if pd.notna(row[col_indices['Ср.балл']]) else 0
                    except:
                        teacher['Ср.балл'] = 0
                else:
                    teacher['Ср.балл'] = 0
                
                # Үштік
                if 'Үштік' in col_indices:
                    try:
                        teacher['Үштік'] = float(row[col_indices['Үштік']]) if pd.notna(row[col_indices['Үштік']]) else 0
                    except:
                        teacher['Үштік'] = 0
                else:
                    teacher['Үштік'] = 0
                
                # Внеклас.
                if 'Внеклас.' in col_indices:
                    try:
                        teacher['Внеклас.'] = float(row[col_indices['Внеклас.']]) if pd.notna(row[col_indices['Внеклас.']]) else 0
                    except:
                        teacher['Внеклас.'] = 0
                else:
                    teacher['Внеклас.'] = 0
                
                # Метод.
                if 'Метод.' in col_indices:
                    try:
                        teacher['Метод.'] = float(row[col_indices['Метод.']]) if pd.notna(row[col_indices['Метод.']]) else 0
                    except:
                        teacher['Метод.'] = 0
                else:
                    teacher['Метод.'] = 0
                
                # Админ.рейтинг
                if 'Админ.рейтинг' in col_indices:
                    try:
                        teacher['Админ.рейтинг'] = float(row[col_indices['Админ.рейтинг']]) if pd.notna(row[col_indices['Админ.рейтинг']]) else 0
                    except:
                        teacher['Админ.рейтинг'] = 0
                else:
                    teacher['Админ.рейтинг'] = 0
                
                # Рейтинг есептеу
                teacher['Результативность'] = (teacher['Ср.балл'] * teacher['Коэф']) - (teacher['Үштік'] * 0.9)
                teacher['Жалпы итог'] = teacher['Результативность'] + teacher['Внеклас.'] + teacher['Метод.'] + teacher['Админ.рейтинг']
                
                # Дөңгелектеу
                teacher['Результативность'] = round(teacher['Результативность'], 1)
                teacher['Жалпы итог'] = round(teacher['Жалпы итог'], 1)
                
                data_rows.append(teacher)
            
            if not data_rows:
                st.error("Деректер табылмады!")
                return
            
            # DataFrame құру
            df = pd.DataFrame(data_rows)
            df.insert(0, '№', range(1, len(df) + 1))
            
            # Бағандарды реттеу
            column_order = ['№', 'ФИО', 'Предмет', 'Коэф', 'Ср.балл', 'Үштік', 
                           'Результативность', 'Внеклас.', 'Метод.', 'Админ.рейтинг', 'Жалпы итог']
            for col in column_order:
                if col not in df.columns:
                    df[col] = 0
            
            st.session_state.df = df[column_order]
            st.session_state.data_loaded = True
            
            st.success(f"✅ {len(df)} мұғалім жүктелді!")
            
            # Кестені кішірейтіп көрсету
            st.subheader("📋 Мұғалімдер тізімі")
            st.dataframe(st.session_state.df, use_container_width=True, height=400)
            
    except requests.exceptions.ConnectionError:
        st.error("❌ Интернет қосылымын тексеріңіз!")
    except Exception as e:
        st.error(f"❌ Жүктеу қатесі: {str(e)}")

def save_to_github_direct():
    """GitHub-қа сақтау"""
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return
    
    token = st.session_state.github_token
    if not token:
        st.error("🔐 GitHub токені енгізілмеген! Сол жақ панельге токеніңізді енгізіңіз.")
        return
    
    try:
        with st.spinner("💾 GitHub-қа сақталуда..."):
            output = BytesIO()
            save_df = st.session_state.df.drop(columns=['№'], errors='ignore')
            save_df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)
            content_b64 = base64.b64encode(output.getvalue()).decode('utf-8')
            
            headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
            get_response = requests.get(GITHUB_API_URL, headers=headers)
            sha = get_response.json().get('sha') if get_response.status_code == 200 else None
            
            data = {"message": "Update reiting1.xlsx", "content": content_b64, "branch": "main"}
            if sha:
                data["sha"] = sha
            
            put_response = requests.put(GITHUB_API_URL, headers=headers, json=data)
            
            if put_response.status_code in [200, 201]:
                st.success("✅ Файл GitHub-қа сәтті сақталды!")
            else:
                st.error(f"❌ GitHub API қатесі: {put_response.status_code}")
                
    except Exception as e:
        st.error(f"❌ Сақтау қатесі: {str(e)}")

def save_to_local():
    """Жергілікті сақтау"""
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return
    
    output = BytesIO()
    save_df = st.session_state.df.drop(columns=['№'], errors='ignore')
    save_df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    
    st.download_button(
        label="📥 Excel файлын жүктеу",
        data=output,
        file_name="reiting1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def calculate_rating():
    """Рейтингті есептеу"""
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return
    
    df = st.session_state.df
    df['Результативность'] = (df['Ср.балл'] * df['Коэф']) - (df['Үштік'] * 0.9)
    df['Жалпы итог'] = df['Результативность'] + df['Внеклас.'] + df['Метод.'] + df['Админ.рейтинг']
    df['Результативность'] = df['Результативность'].round(1)
    df['Жалпы итог'] = df['Жалпы итог'].round(1)
    st.session_state.df = df
    st.success("✅ Рейтинг сәтті есептелді!")
    st.rerun()

def add_teacher():
    """Жаңа мұғалім қосу"""
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return
    
    with st.form("add_teacher_form"):
        st.subheader("➕ Жаңа мұғалім қосу")
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("ФИО *")
            subject = st.text_input("Предмет *")
            coeff = st.number_input("Коэф.", value=1.0, step=0.1)
            avg_score = st.number_input("Ср.балл", value=4.0, step=0.1)
            triplets = st.number_input("Үштік саны", value=0, step=1)
        with col2:
            extra = st.number_input("Внеклас. работа", value=0.0, step=0.5)
            method = st.number_input("Метод. деятельность", value=0.0, step=0.5)
            admin = st.number_input("Админ. рейтинг", value=0.0, step=0.5)
        
        if st.form_submit_button("Қосу"):
            if not name or not subject:
                st.error("ФИО және Предмет толтырылуы керек!")
                return
            
            new_row = {
                '№': len(st.session_state.df) + 1,
                'ФИО': name,
                'Предмет': subject,
                'Коэф': coeff,
                'Ср.балл': avg_score,
                'Үштік': triplets,
                'Результативность': 0,
                'Внеклас.': extra,
                'Метод.': method,
                'Админ.рейтинг': admin,
                'Жалпы итог': 0
            }
            
            new_df = pd.DataFrame([new_row])
            st.session_state.df = pd.concat([st.session_state.df, new_df], ignore_index=True)
            st.session_state.df['№'] = range(1, len(st.session_state.df) + 1)
            calculate_rating()
            st.success(f"✅ {name} қосылды!")
            st.rerun()

def delete_teacher():
    """Мұғалімді өшіру"""
    if st.session_state.df is None or len(st.session_state.df) == 0:
        st.warning("Өшіретін мұғалім жоқ!")
        return
    
    teacher_options = {f"{row['№']}. {row['ФИО']} - {row['Предмет']}": idx 
                       for idx, row in st.session_state.df.iterrows()}
    selected = st.selectbox("Өшіретін мұғалімді таңдаңыз:", list(teacher_options.keys()))
    
    if st.button("🗑 Өшіру"):
        idx = teacher_options[selected]
        teacher_name = st.session_state.df.loc[idx, 'ФИО']
        st.session_state.df = st.session_state.df.drop(idx).reset_index(drop=True)
        st.session_state.df['№'] = range(1, len(st.session_state.df) + 1)
        st.success(f"🗑 {teacher_name} өшірілді!")
        st.rerun()

# ==================== НЕГІЗГІ ИНТЕРФЕЙС ====================

# Сол жақ панель
with st.sidebar:
    st.markdown("### 🔐 GitHub авторизациясы")
    token = st.text_input("GitHub Personal Access Token:", type="password")
    if token:
        st.session_state.github_token = token
        st.success("✅ Токен сақталды!")
    st.markdown("---")
    st.markdown("**Токен алу үшін:**")
    st.markdown("1. [GitHub Settings](https://github.com/settings/tokens)")
    st.markdown("2. Generate new token (classic)")
    st.markdown("3. `repo` рұқсатын қосыңыз")

# Негізгі бет
st.title("📊 Мұғалімдер рейтингі")
st.markdown("---")

# Батырмалар қатары
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    if st.button("📥 GitHub-тан жүктеу", use_container_width=True):
        load_from_github()

with col2:
    if st.button("💾 GitHub-қа сақтау", use_container_width=True):
        save_to_github_direct()

with col3:
    if st.button("📁 Excel жүктеу", use_container_width=True):
        save_to_local()

with col4:
    if st.button("➕ Мұғалім қосу", use_container_width=True):
        add_teacher()

with col5:
    if st.button("🗑 Мұғалім өшіру", use_container_width=True):
        delete_teacher()

st.markdown("---")

# Кестені көрсету
if st.session_state.df is not None and len(st.session_state.df) > 0:
    col1, col2 = st.columns([3, 1])
    with col1:
        st.subheader("📋 Мұғалімдер тізімі")
    with col2:
        if st.button("🔄 Рейтингті есептеу", use_container_width=True):
            calculate_rating()
    
    st.dataframe(st.session_state.df, use_container_width=True, height=400)
    
    # Статистика
    st.subheader("📊 Статистика")
    valid_df = st.session_state.df[st.session_state.df['Жалпы итог'].notna()]
    
    if len(valid_df) > 0:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("👥 Мұғалімдер саны", len(valid_df))
        with c2:
            st.metric("📊 Орташа итог", f"{valid_df['Жалпы итог'].mean():.1f}")
        with c3:
            best = valid_df.loc[valid_df['Жалпы итог'].idxmax()]
            st.metric("🏆 Ең жақсы", f"{best['ФИО'][:15]} ({best['Жалпы итог']:.1f})")
        with c4:
            worst = valid_df.loc[valid_df['Жалпы итог'].idxmin()]
            st.metric("⚠️ Ең төмен", f"{worst['ФИО'][:15]} ({worst['Жалпы итог']:.1f})")
    
    # CSV экспорт
    csv = st.session_state.df.to_csv(index=False).encode('utf-8-sig')
    st.download_button("📥 CSV жүктеу", data=csv, file_name="reiting.csv", mime="text/csv")

else:
    st.info("👆 Бастау үшін **«GitHub-тан жүктеу»** батырмасын басыңыз.")

st.markdown("---")
st.caption("© Мұғалімдер рейтингі | GitHub интеграциясы")