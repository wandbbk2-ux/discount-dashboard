import streamlit as st
import pandas as pd
from typing import Dict, Tuple, Optional
from datetime import datetime, timedelta
import os

# ------------------------------
# 1. Загрузка данных артикулов
# ------------------------------
@st.cache_data
def load_sku_data(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path, engine='openpyxl')
    df.columns = df.columns.str.strip()
    if 'Артикул' in df.columns:
        df.rename(columns={'Артикул': 'sku'}, inplace=True)
    df.dropna(subset=['sku'], inplace=True)
    df['sku'] = df['sku'].astype(str).str.strip()
    return df

# ------------------------------
# 2. Загрузка правил скидок
# ------------------------------
@st.cache_data
def load_discount_rules(file_path: str = "discount_rules.xlsx") -> pd.DataFrame:
    df = pd.read_excel(file_path, engine='openpyxl')
    df.columns = df.columns.str.strip()
    for col in ['level04', 'location', 'type', 'size', 'functional']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    return df

# ------------------------------
# 3. Загрузка и сохранение лога
# ------------------------------
LOG_FILE = "compensation_log.xlsx"

def load_log() -> pd.DataFrame:
    if os.path.exists(LOG_FILE):
        df = pd.read_excel(LOG_FILE, engine='openpyxl')
        if 'Дата и время' in df.columns:
            df['Дата и время'] = pd.to_datetime(df['Дата и время'], errors='coerce')
        return df
    else:
        return pd.DataFrame(columns=[
            'Дата и время', 'Артикул', 'Модель', 'Категория L2', 'Категория L3', 'Категория L4',
            'Цена', 'Локация повреждения', 'Тип повреждения', 'Размер', 'Влияет на функциональность',
            'Мин скидка %', 'Макс скидка %', 'Предложенная скидка %',
            'Маркетплейс', 'Промокод', 'Номер заказа', 'Статус'
        ])

def save_to_log(data: dict):
    df = load_log()
    new_row = pd.DataFrame([data])
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_excel(LOG_FILE, index=False, engine='openpyxl')

# ------------------------------
# 4. Конфигурация локаций и размеров по категориям ИЗ ПРАВИЛ
# ------------------------------
def get_locations_for_category(rules_df: pd.DataFrame, category: str) -> list:
    """Извлекает уникальные локации для категории из файла правил"""
    locations = rules_df[rules_df['level04'] == category]['location'].unique()
    if len(locations) == 0:
        # Если нет специфичных — берём default
        locations = rules_df[rules_df['level04'] == 'default']['location'].unique()
    return sorted(locations.tolist())

def get_sizes_for_category(rules_df: pd.DataFrame, category: str, location: str, damage_type: str) -> list:
    """Извлекает возможные размеры для конкретной комбинации"""
    sizes = rules_df[
        (rules_df['level04'] == category) &
        (rules_df['location'] == location) &
        (rules_df['type'] == damage_type)
    ]['size'].unique()
    
    if len(sizes) == 0:
        # Дефолтные размеры
        sizes = rules_df[rules_df['level04'] == 'default']['size'].unique()
    
    # Сортируем по логике: сначала числовые диапазоны, потом 'Любой'
    def sort_key(s):
        if s == 'Любой':
            return (2, s)
        return (1, s)
    
    return sorted(sizes.tolist(), key=sort_key)

# ------------------------------
# 5. Функция расчёта скидки ПО ПРАВИЛАМ (с учётом размера!)
# ------------------------------
def get_discount_from_rules(
    rules_df: pd.DataFrame,
    level04: str,
    location: str,
    damage_type: str,
    size: str,
    functional: str,
    product_price: float,
    missing_part: Optional[Tuple[str, float]] = None
) -> Tuple[float, float]:
    
    # Особый случай: отсутствует комплектующее
    if damage_type == 'Отсутствует комплектующее' and missing_part:
        percent = (missing_part[1] / product_price) * 100 + 5
        percent = min(percent, 15)
        percent = round(percent / 5) * 5
        return (percent, percent)
    
    # Нормализация размера
    size_normalized = size.strip()
    
    # 1. Ищем точное совпадение
    match = rules_df[
        (rules_df['level04'] == level04) &
        (rules_df['location'] == location) &
        (rules_df['type'] == damage_type) &
        (rules_df['size'] == size_normalized) &
        (rules_df['functional'] == functional)
    ]
    
    if not match.empty:
        row = match.iloc[0]
        return (float(row['min_discount']), float(row['max_discount']))
    
    # 2. Ищем с size = 'Любой' (если не нашли точный размер)
    match_any = rules_df[
        (rules_df['level04'] == level04) &
        (rules_df['location'] == location) &
        (rules_df['type'] == damage_type) &
        (rules_df['size'].str.lower() == 'любой') &
        (rules_df['functional'] == functional)
    ]
    
    if not match_any.empty:
        row = match_any.iloc[0]
        return (float(row['min_discount']), float(row['max_discount']))
    
    # 3. Ищем default правила
    default_match = rules_df[
        (rules_df['level04'] == 'default') &
        (rules_df['location'] == location) &
        (rules_df['type'] == damage_type) &
        (rules_df['functional'] == functional)
    ]
    
    if not default_match.empty:
        default_match = default_match[default_match['size'] == size_normalized]
        if default_match.empty:
            default_match = default_match[default_match['size'] == 'Любой']
        
        if not default_match.empty:
            row = default_match.iloc[0]
            return (float(row['min_discount']), float(row['max_discount']))
    
    # 4. Абсолютный fallback
    return (5, 10)

# ------------------------------
# 6. Динамический вопросник (без жёстких списков)
# ------------------------------
def build_questions_from_rules(rules_df: pd.DataFrame, category: str):
    """Строит вопросы динамически на основе правил"""
    # Получаем уникальные типы повреждений для категории
    damage_types = rules_df[rules_df['level04'] == category]['type'].unique()
    if len(damage_types) == 0:
        damage_types = rules_df[rules_df['level04'] == 'default']['type'].unique()
    
    # Базовые вопросы
    questions_config = {
        'location': {
            'label': 'Где расположено повреждение?',
            'options': get_locations_for_category(rules_df, category)
        },
        'type': {
            'label': 'Тип повреждения?',
            'options': sorted(damage_types.tolist())
        },
        'functional': {
            'label': 'Влияет ли на функциональность?',
            'options': ['Нет', 'Да']
        }
    }
    return questions_config

# ------------------------------
# 7. Загрузка данных
# ------------------------------
DEFAULT_SKU_FILE = "Список товарных групп.xlsx"
RULES_FILE = "discount_rules.xlsx"

df_skus = None
try:
    df_skus = load_sku_data(DEFAULT_SKU_FILE)
    st.sidebar.success(f"✅ Загружены артикулы: {DEFAULT_SKU_FILE}")
except:
    st.sidebar.error("❌ Файл артикулов не найден")
    st.stop()

rules_df = None
try:
    rules_df = load_discount_rules(RULES_FILE)
    st.sidebar.success(f"✅ Загружены правила скидок: {RULES_FILE}")
    st.sidebar.info(f"📋 Всего правил: {len(rules_df)}")
except Exception as e:
    st.sidebar.error(f"❌ Ошибка загрузки discount_rules.xlsx: {e}")
    st.stop()

sku_to_info = {}
for _, row in df_skus.iterrows():
    sku = str(row['sku']).strip()
    sku_to_info[sku] = {
        'level02': row.get('Level 02', ''),
        'level03': row.get('Level 03', ''),
        'level04': row.get('Level 04', ''),
        'model': row.get('Модель', '')
    }

MARKETPLACES_LIST = ["Ozon", "Yandex Market", "Wildberries"]
MARKETPLACES = {
    "Ozon": {"emoji": "💙🌸", "name": "Ozon"},
    "Yandex Market": {"emoji": "🟡", "name": "Yandex Market"},
    "Wildberries": {"emoji": "🟣", "name": "Wildberries"}
}

# ------------------------------
# 8. Интерфейс Streamlit
# ------------------------------
st.set_page_config(page_title="Система компенсаций", layout="wide")
st.title("📦 Определение компенсации за повреждения")

# Боковая панель с отчётами
st.sidebar.header("📊 Отчёты")
log_df = load_log()

if not log_df.empty:
    st.sidebar.subheader("📈 Статистика по площадкам")
    marketplace_counts = log_df['Маркетплейс'].value_counts()
    for mp, count in marketplace_counts.items():
        emoji = MARKETPLACES.get(mp, {}).get("emoji", "🛒")
        st.sidebar.metric(f"{emoji} {mp}", count)
    
    st.sidebar.divider()
    
    # Фильтр по дате
    st.sidebar.subheader("📅 Фильтр по дате")
    period = st.sidebar.selectbox(
        "Выберите период",
        options=["Сегодня", "Вчера", "Последние 7 дней", "Последние 30 дней", "Этот месяц", "Прошлый месяц", "Произвольный"]
    )
    
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    
    if period == "Сегодня":
        start_date = today
        end_date = today + timedelta(days=1) - timedelta(seconds=1)
    elif period == "Вчера":
        start_date = today - timedelta(days=1)
        end_date = today - timedelta(seconds=1)
    elif period == "Последние 7 дней":
        start_date = today - timedelta(days=7)
        end_date = today + timedelta(days=1) - timedelta(seconds=1)
    elif period == "Последние 30 дней":
        start_date = today - timedelta(days=30)
        end_date = today + timedelta(days=1) - timedelta(seconds=1)
    elif period == "Этот месяц":
        start_date = today.replace(day=1)
        end_date = today + timedelta(days=1) - timedelta(seconds=1)
    elif period == "Прошлый месяц":
        first_day_this_month = today.replace(day=1)
        last_day_last_month = first_day_this_month - timedelta(days=1)
        start_date = last_day_last_month.replace(day=1)
        end_date = last_day_last_month
    else:
        col1, col2 = st.sidebar.columns(2)
        with col1:
            start_date = st.date_input("От", today - timedelta(days=30))
        with col2:
            end_date = st.date_input("До", today)
        start_date = datetime.combine(start_date, datetime.min.time())
        end_date = datetime.combine(end_date, datetime.max.time())
    
    marketplace_filter = st.sidebar.selectbox("Маркетплейс", options=["Все"] + MARKETPLACES_LIST)
    
    # Применяем фильтры
    filtered_df = log_df[(log_df['Дата и время'] >= start_date) & (log_df['Дата и время'] <= end_date)]
    if marketplace_filter != "Все":
        filtered_df = filtered_df[filtered_df['Маркетплейс'] == marketplace_filter]
    
    st.sidebar.divider()
    st.sidebar.subheader("📊 Результат фильтрации")
    st.sidebar.metric("Всего записей", len(filtered_df))
    
    if not filtered_df.empty:
        avg_discount = filtered_df['Предложенная скидка %'].mean()
        st.sidebar.metric("Средняя скидка", f"{avg_discount:.0f}%")
        
        csv_data = filtered_df.to_csv(index=False, sep=';', encoding='utf-8-sig')
        st.sidebar.download_button(
            label=f"📥 Скачать отчёт ({len(filtered_df)} записей)",
            data=csv_data,
            file_name=f"compensation_report_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}_{marketplace_filter}.csv",
            mime="text/csv"
        )
    
    st.sidebar.divider()
    st.sidebar.subheader("🕐 Последние действия")
    if not log_df.empty:
        last_5 = log_df.tail(5)
        for _, row in last_5.iterrows():
            date_str = row['Дата и время'].strftime("%d.%m.%Y %H:%M") if hasattr(row['Дата и время'], 'strftime') else str(row['Дата и время'])[:16]
            marketplace = row.get('Маркетплейс', '—')
            emoji = MARKETPLACES.get(marketplace, {}).get("emoji", "🛒")
            st.sidebar.caption(f"{emoji} **{row['Артикул']}** → {row['Предложенная скидка %']}% | {date_str}")

# ------------------------------
# 9. Основная форма
# ------------------------------
sku_input = st.text_input("Введите артикул товара", placeholder="Например: 397997")

if sku_input:
    sku_input = str(sku_input).strip()
    sku_info = sku_to_info.get(sku_input)
    
    if not sku_info:
        st.error("❌ Артикул не найден в базе. Проверьте ввод.")
        st.stop()
    
    st.subheader("📋 Информация о товаре")
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Модель:** {sku_info['model']}")
        st.write(f"**Категория (L2):** {sku_info['level02']}")
        st.write(f"**Категория (L3):** {sku_info['level03']}")
    with col2:
        st.write(f"**Категория (L4):** {sku_info['level04']}")
        price = st.number_input("💰 Цена товара (тенге)", min_value=0.0, value=100000.0, step=1000.0)
    
    level04 = sku_info['level04']
    
    # Строим вопросы динамически из правил
    questions_config = build_questions_from_rules(rules_df, level04)
    
    st.subheader("🔍 Опишите повреждение")
    
    answers = {}
    
    # Локация
    answers['location'] = st.selectbox(
        questions_config['location']['label'],
        questions_config['location']['options']
    )
    
    # Тип повреждения
    answers['type'] = st.selectbox(
        questions_config['type']['label'],
        questions_config['type']['options']
    )
    
    # Размер (только если тип требует размера)
    damage_types_without_size = ['Отсутствует комплектующее', 'Повреждение упаковки']
    if answers['type'] not in damage_types_without_size:
        size_options = get_sizes_for_category(rules_df, level04, answers['location'], answers['type'])
        answers['size'] = st.selectbox(
            "📏 Размер повреждения",
            size_options
        )
    else:
        answers['size'] = 'Любой'
    
    # Влияние на функциональность
    answers['functional'] = st.selectbox(
        questions_config['functional']['label'],
        questions_config['functional']['options']
    )
    
    missing_part = None
    if answers['type'] == 'Отсутствует комплектующее':
        st.info("🔧 Укажите недостающую деталь")
        missing_part_name = st.text_input("Название детали")
        missing_part_price = st.number_input("Рыночная стоимость детали (тенге)", min_value=0.0, value=0.0)
        if missing_part_name and missing_part_price > 0:
            missing_part = (missing_part_name, missing_part_price)
    
    if st.button("💸 Рассчитать скидку", type="primary"):
        if price <= 0:
            st.warning("⚠️ Пожалуйста, укажите корректную цену товара")
        else:
            min_d, max_d = get_discount_from_rules(
                rules_df,
                level04,
                answers['location'],
                answers['type'],
                answers.get('size', 'Любой'),
                answers['functional'],
                price,
                missing_part
            )
            
            mid = int((min_d + max_d) / 2)
            
            st.success(f"**Рекомендуемая скидка:** {int(min_d)}% – {int(max_d)}%")
            st.info(f"**🎯 Стартовая компенсация:** {mid}%")
            
            # Формирование предложения
            if answers['type'] == 'Отсутствует комплектующее' and missing_part:
                proposal = f"Здравствуйте! По вашему обращению: отсутствует комплектующее ({missing_part[0]}). Мы можем предложить компенсацию в размере {mid}% от стоимости товара для приобретения недостающей детали. Если вы готовы оставить товар, сообщите нам."
            else:
                proposal = f"Здравствуйте! По вашему обращению о повреждении ({answers['location']}, {answers['type']}, размер: {answers.get('size', 'не указан')}) мы можем предложить компенсацию в размере {mid}% от стоимости товара. Если вы готовы оставить товар, сообщите нам."
            
            st.code(proposal, language='text')
            
            # Сохраняем в session_state
            st.session_state['calculation'] = {
                'sku': sku_input,
                'model': sku_info['model'],
                'level02': sku_info['level02'],
                'level03': sku_info['level03'],
                'level04': sku_info['level04'],
                'price': price,
                'location': answers['location'],
                'damage_type': answers['type'],
                'size': answers.get('size', 'Любой'),
                'functional': answers['functional'],
                'min_discount': int(min_d),
                'max_discount': int(max_d),
                'proposed_discount': mid,
                'missing_part': missing_part[0] if missing_part else None,
                'missing_part_price': missing_part[1] if missing_part else None
            }
            st.session_state['show_agreement'] = True
    
    # Форма согласия
    if st.session_state.get('show_agreement', False):
        st.divider()
        st.subheader("✅ Фиксация согласия клиента")
        
        with st.form("agreement_form"):
            marketplace_options = list(MARKETPLACES.keys())
            marketplace_labels = [f"{MARKETPLACES[m]['emoji']} {MARKETPLACES[m]['name']}" for m in marketplace_options]
            selected_label = st.selectbox("Маркетплейс", options=marketplace_labels)
            selected_marketplace = selected_label.split(" ")[1] if " " in selected_label else selected_label
            
            promo_code = st.text_input("Промокод (если применимо)", placeholder="Например: DISCOUNT2024")
            order_number = st.text_input("Номер заказа", placeholder="Обязательное поле")
            
            submitted = st.form_submit_button("✅ Подтвердить согласие клиента")
            
            if submitted:
                if not order_number:
                    st.error("Пожалуйста, введите номер заказа")
                else:
                    calc = st.session_state['calculation']
                    log_entry = {
                        'Дата и время': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        'Артикул': calc['sku'],
                        'Модель': calc['model'],
                        'Категория L2': calc['level02'],
                        'Категория L3': calc['level03'],
                        'Категория L4': calc['level04'],
                        'Цена': calc['price'],
                        'Локация повреждения': calc['location'],
                        'Тип повреждения': calc['damage_type'],
                        'Размер': calc['size'],
                        'Влияет на функциональность': calc['functional'],
                        'Мин скидка %': calc['min_discount'],
                        'Макс скидка %': calc['max_discount'],
                        'Предложенная скидка %': calc['proposed_discount'],
                        'Маркетплейс': selected_marketplace,
                        'Промокод': promo_code if promo_code else "",
                        'Номер заказа': order_number,
                        'Статус': 'Согласие получено'
                    }
                    
                    save_to_log(log_entry)
                    st.success("✅ Согласие зафиксировано!")
                    st.success(f"📊 Данные сохранены в файл: {LOG_FILE}")
                    st.balloons()
                    
                    st.session_state['show_agreement'] = False
                    st.info(f"**Итог:** Клиент {order_number} согласился на компенсацию {calc['proposed_discount']}% на маркетплейсе {selected_marketplace}")