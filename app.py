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
    """Загружает Excel-файл с колонками Артикул, Level 02, Level 03, Level 04, Модель"""
    df = pd.read_excel(file_path, engine='openpyxl')
    df.columns = df.columns.str.strip()
    if 'Артикул' in df.columns:
        df.rename(columns={'Артикул': 'sku'}, inplace=True)
    df.dropna(subset=['sku'], inplace=True)
    df['sku'] = df['sku'].astype(str).str.strip()
    return df

# ------------------------------
# 2. Загрузка и сохранение лога компенсаций
# ------------------------------
LOG_FILE = "compensation_log.xlsx"

def load_log() -> pd.DataFrame:
    """Загружает существующий лог или создаёт новый"""
    if os.path.exists(LOG_FILE):
        df = pd.read_excel(LOG_FILE, engine='openpyxl')
        if 'Дата и время' in df.columns:
            df['Дата и время'] = pd.to_datetime(df['Дата и время'], errors='coerce')
        return df
    else:
        df = pd.DataFrame(columns=[
            'Дата и время', 'Артикул', 'Модель', 'Категория L2', 'Категория L3', 'Категория L4',
            'Цена', 'Локация повреждения', 'Тип повреждения', 'Размер', 'Влияет на функциональность',
            'Мин скидка %', 'Макс скидка %', 'Предложенная скидка %', 
            'Маркетплейс', 'Промокод', 'Номер заказа', 'Статус'
        ])
        return df

def save_to_log(data: dict):
    """Сохраняет запись о согласии клиента в Excel"""
    df = load_log()
    new_row = pd.DataFrame([data])
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_excel(LOG_FILE, index=False, engine='openpyxl')

# ------------------------------
# 3. Загрузка файла артикулов
# ------------------------------
DEFAULT_FILE = "Список товарных групп.xlsx"
df_skus = None

try:
    df_skus = load_sku_data(DEFAULT_FILE)
    st.sidebar.success(f"Загружен файл: {DEFAULT_FILE}")
except FileNotFoundError:
    st.sidebar.error("Файл по умолчанию не найден. Пожалуйста, загрузите файл вручную.")
    uploaded_file = st.sidebar.file_uploader("Выберите Excel-файл с артикулами", type=['xlsx'])
    if uploaded_file:
        df_skus = load_sku_data(uploaded_file)
        st.sidebar.success("Файл загружен")
    else:
        st.stop()

if df_skus is None:
    st.stop()

# Создаём словарь для быстрого поиска по артикулу
sku_to_info = {}
for _, row in df_skus.iterrows():
    sku = str(row['sku']).strip()
    sku_to_info[sku] = {
        'level02': row.get('Level 02', ''),
        'level03': row.get('Level 03', ''),
        'level04': row.get('Level 04', ''),
        'model': row.get('Модель', '')
    }

# ------------------------------
# 4. Конфигурация чеклистов
# ------------------------------
QUESTIONS = {
    'location': {
        'label': 'Где расположено повреждение?',
        'options': [
            'Лицевая часть', 'Боковая часть', 'Тыльная часть',
            'Экран', 'Корпус', 'Дверца', 'Клавиатура', 'Рамка',
            'Камера', 'Дисплей', 'Панель управления'
        ]
    },
    'type': {
        'label': 'Тип повреждения?',
        'options': ['Царапина', 'Скол', 'Вмятина', 'Трещина', 'Отсутствует комплектующее', 'Повреждение упаковки']
    },
    'size': {
        'label': 'Размер повреждения (для царапин/сколов/трещин)?',
        'options': ['До 1 см', '1–3 см', '3–5 см', 'Более 5 см', 'Любой']
    },
    'functional': {
        'label': 'Влияет ли на функциональность?',
        'options': ['Да', 'Нет']
    }
}

CHECKLISTS = {
    'Ноутбуки': {
        'questions': ['location', 'type', 'size', 'functional'],
        'location_options': ['Экран', 'Корпус (верхняя крышка)', 'Корпус (нижняя часть)', 'Клавиатура', 'Петли', 'Торцы']
    },
    'Смартфоны': {
        'questions': ['location', 'type', 'size', 'functional'],
        'location_options': ['Экран', 'Задняя крышка', 'Рамка', 'Камера']
    },
    'Планшеты': {
        'questions': ['location', 'type', 'size', 'functional'],
        'location_options': ['Экран', 'Корпус']
    },
    'Телевизоры (LCD TV)': {
        'questions': ['location', 'type', 'size', 'functional'],
        'location_options': ['Экран', 'Корпус', 'Подставка']
    },
    'default': {
        'questions': ['location', 'type', 'size', 'functional'],
        'location_options': QUESTIONS['location']['options']
    }
}

# Список маркетплейсов
MARKETPLACES_LIST = ["Ozon", "Yandex Market", "Wildberries"]

# ------------------------------
# 5. Логотипы маркетплейсов
# ------------------------------
MARKETPLACES = {
    "Ozon": {"emoji": "💙🌸", "name": "Ozon"},
    "Yandex Market": {"emoji": "🟡", "name": "Yandex Market"},
    "Wildberries": {"emoji": "🟣", "name": "Wildberries"}
}

# ------------------------------
# 6. Функция для расчёта скидки с дефолтными правилами
# ------------------------------
def get_discount_range(
    level04: str,
    location: str,
    damage_type: str,
    size: str,
    functional: str,
    product_price: float,
    missing_part: Optional[Tuple[str, float]] = None
) -> Tuple[float, float]:
    """Возвращает (min_discount, max_discount) в процентах"""
    
    if damage_type == 'Отсутствует комплектующее' and missing_part:
        percent = (missing_part[1] / product_price) * 100 + 5
        percent = min(percent, 15)
        percent = round(percent / 5) * 5
        return (percent, percent)
    
    if damage_type == 'Царапина':
        if functional == 'Нет':
            return (5, 10)
        else:
            return (10, 15)
    elif damage_type == 'Скол':
        if functional == 'Нет':
            return (10, 15)
        else:
            return (15, 20)
    elif damage_type == 'Вмятина':
        if functional == 'Нет':
            return (15, 20)
        else:
            return (20, 25)
    elif damage_type == 'Трещина':
        if functional == 'Нет':
            return (25, 30)
        else:
            return (30, 35)
    elif damage_type == 'Повреждение упаковки':
        return (1, 3)
    else:
        return (5, 10)

# ------------------------------
# 7. Функция для фильтрации отчёта
# ------------------------------
def filter_report(df: pd.DataFrame, start_date: datetime, end_date: datetime, marketplace: str) -> pd.DataFrame:
    """Фильтрует отчёт по дате и маркетплейсу"""
    if df.empty:
        return df
    
    mask_date = (df['Дата и время'] >= start_date) & (df['Дата и время'] <= end_date)
    filtered = df[mask_date]
    
    if marketplace != "Все":
        filtered = filtered[filtered['Маркетплейс'] == marketplace]
    
    return filtered

# ------------------------------
# 8. Интерфейс Streamlit
# ------------------------------
st.set_page_config(page_title="Система компенсаций", layout="wide")
st.title("📦 Определение компенсации за повреждения")

# ==================== БОКОВАЯ ПАНЕЛЬ С ОТЧЁТАМИ ====================
st.sidebar.header("📊 Отчёты")

log_df = load_log()

if not log_df.empty:
    # Статистика по маркетплейсам
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
        st.sidebar.info(f"📅 {start_date.strftime('%d.%m.%Y')}")
        
    elif period == "Вчера":
        start_date = today - timedelta(days=1)
        end_date = today - timedelta(seconds=1)
        st.sidebar.info(f"📅 {start_date.strftime('%d.%m.%Y')}")
        
    elif period == "Последние 7 дней":
        start_date = today - timedelta(days=7)
        end_date = today + timedelta(days=1) - timedelta(seconds=1)
        st.sidebar.info(f"📅 {start_date.strftime('%d.%m.%Y')} – {today.strftime('%d.%m.%Y')}")
        
    elif period == "Последние 30 дней":
        start_date = today - timedelta(days=30)
        end_date = today + timedelta(days=1) - timedelta(seconds=1)
        st.sidebar.info(f"📅 {start_date.strftime('%d.%m.%Y')} – {today.strftime('%d.%m.%Y')}")
        
    elif period == "Этот месяц":
        start_date = today.replace(day=1)
        end_date = today + timedelta(days=1) - timedelta(seconds=1)
        st.sidebar.info(f"📅 {start_date.strftime('%d.%m.%Y')} – {today.strftime('%d.%m.%Y')}")
        
    elif period == "Прошлый месяц":
        first_day_this_month = today.replace(day=1)
        last_day_last_month = first_day_this_month - timedelta(days=1)
        start_date = last_day_last_month.replace(day=1)
        end_date = last_day_last_month
        st.sidebar.info(f"📅 {start_date.strftime('%d.%m.%Y')} – {end_date.strftime('%d.%m.%Y')}")
        
    else:
        col1, col2 = st.sidebar.columns(2)
        with col1:
            start_date = st.date_input("От", today - timedelta(days=30))
        with col2:
            end_date = st.date_input("До", today)
        start_date = datetime.combine(start_date, datetime.min.time())
        end_date = datetime.combine(end_date, datetime.max.time())
        st.sidebar.caption(f"📅 {start_date.strftime('%d.%m.%Y')} – {end_date.strftime('%d.%m.%Y')}")
    
    # Фильтр по маркетплейсу
    st.sidebar.subheader("🏪 Фильтр по маркетплейсу")
    marketplace_filter = st.sidebar.selectbox(
        "Маркетплейс",
        options=["Все"] + MARKETPLACES_LIST
    )
    
    # Применяем фильтры
    filtered_df = filter_report(log_df, start_date, end_date, marketplace_filter)
    
    # Показываем статистику
    st.sidebar.divider()
    st.sidebar.subheader("📊 Результат фильтрации")
    st.sidebar.metric("Всего записей", len(filtered_df))
    
    if not filtered_df.empty:
        avg_discount = filtered_df['Предложенная скидка %'].mean()
        st.sidebar.metric("Средняя скидка", f"{avg_discount:.0f}%")
        
        csv_data = filtered_df.to_csv(index=False, sep=';', encoding='utf-8-sig')
        
        date_str = start_date.strftime("%Y%m%d")
        end_str = end_date.strftime("%Y%m%d")
        market_str = marketplace_filter if marketplace_filter != "Все" else "all"
        filename = f"compensation_report_{date_str}_{end_str}_{market_str}.csv"
        
        st.sidebar.download_button(
            label=f"📥 Скачать отчёт ({len(filtered_df)} записей)",
            data=csv_data,
            file_name=filename,
            mime="text/csv",
            help=f"Отчёт за период {start_date.strftime('%d.%m.%Y')} – {end_date.strftime('%d.%m.%Y')} по маркетплейсу {marketplace_filter}"
        )
        
        # Excel версия
        filtered_df.to_excel("temp_filtered.xlsx", index=False, engine='openpyxl')
        with open("temp_filtered.xlsx", "rb") as f:
            st.sidebar.download_button(
                label="📊 Скачать отчёт (Excel)",
                data=f,
                file_name=filename.replace('.csv', '.xlsx'),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        os.remove("temp_filtered.xlsx")
    
    st.sidebar.divider()
    
    # Полный отчёт
    if st.sidebar.button("📋 Полный отчёт (Excel)"):
        log_df.to_excel("compensation_report_full.xlsx", index=False, engine='openpyxl')
        with open("compensation_report_full.xlsx", "rb") as f:
            st.sidebar.download_button(
                label="✅ Полный отчёт",
                data=f,
                file_name="compensation_report_full.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    # Последние 5 записей
    st.sidebar.divider()
    st.sidebar.subheader("🕐 Последние действия")
    if not log_df.empty:
        last_5 = log_df.tail(5)
        for _, row in last_5.iterrows():
            date_str = row['Дата и время'].strftime("%d.%m.%Y %H:%M") if hasattr(row['Дата и время'], 'strftime') else str(row['Дата и время'])[:16]
            marketplace = row.get('Маркетплейс', '—')
            emoji = MARKETPLACES.get(marketplace, {}).get("emoji", "🛒")
            st.sidebar.caption(f"{emoji} **{row['Артикул']}** → {row['Предложенная скидка %']}% | {date_str}")
    
else:
    st.sidebar.info("Пока нет записей о согласиях. После первой фиксации здесь появится отчёт.")

# ==================== ОСНОВНАЯ ЧАСТЬ ====================

sku_input = st.text_input("Введите артикул товара", placeholder="Например: 470591")

if sku_input:
    sku_input = str(sku_input).strip()
    sku_info = sku_to_info.get(sku_input)
    if not sku_info:
        st.error("Артикул не найден в базе. Пожалуйста, проверьте ввод.")
        st.stop()

    st.subheader("Информация о товаре")
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Модель:** {sku_info['model']}")
        st.write(f"**Категория (L2):** {sku_info['level02']}")
        st.write(f"**Категория (L3):** {sku_info['level03']}")
    with col2:
        st.write(f"**Категория (L4):** {sku_info['level04']}")
        price = st.number_input("Цена товара (тенге)", min_value=0.0, value=100000.0, step=1000.0)

    level04 = sku_info['level04']
    checklist = CHECKLISTS.get(level04, CHECKLISTS['default'])
    st.subheader("Опишите повреждение")

    answers = {}
    for q_name in checklist['questions']:
        q_data = QUESTIONS[q_name]
        if q_name == 'location':
            options = checklist.get('location_options', q_data['options'])
            answers[q_name] = st.selectbox(q_data['label'], options)
        else:
            answers[q_name] = st.selectbox(q_data['label'], q_data['options'])
    
    if answers.get('size') == '' or answers.get('size') is None:
        answers['size'] = 'Любой'

    missing_part = None
    if answers['type'] == 'Отсутствует комплектующее':
        st.info("Укажите недостающую деталь")
        missing_part_name = st.text_input("Название детали")
        missing_part_price = st.number_input("Рыночная стоимость детали (тенге)", min_value=0.0, value=0.0)
        if missing_part_name and missing_part_price > 0:
            missing_part = (missing_part_name, missing_part_price)

    # Кнопка расчёта
    if st.button("Рассчитать скидку"):
        if price <= 0:
            st.warning("Пожалуйста, укажите корректную цену товара")
        else:
            min_d, max_d = get_discount_range(
                level04,
                answers['location'],
                answers['type'],
                answers.get('size', 'Любой'),
                answers.get('functional', 'Нет'),
                price,
                missing_part
            )
            
            if min_d == 0 and max_d == 0:
                st.warning("Для данной комбинации повреждений нет настроенных правил.")
            else:
                mid = int((min_d + max_d) / 2)
                st.success(f"**Рекомендуемая скидка:** {int(min_d)}% – {int(max_d)}%")
                st.info(f"**Стартовая компенсация:** {mid}%")

                if answers['type'] == 'Отсутствует комплектующее' and missing_part:
                    proposal = f"Здравствуйте! По вашему обращению: отсутствует комплектующее ({missing_part[0]}). Мы можем предложить компенсацию в размере {mid}% от стоимости товара для приобретения недостающей детали. Если вы готовы оставить товар, сообщите нам."
                else:
                    proposal = f"Здравствуйте! По вашему обращению о повреждении ({answers['location']}, {answers['type']} размер {answers.get('size', '')}) мы можем предложить компенсацию в размере {mid}% от стоимости товара. Если вы готовы оставить товар, сообщите нам."
                
                st.code(proposal, language='text')
                
                # Сохраняем расчёт в session_state
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
                    'functional': answers.get('functional', 'Нет'),
                    'min_discount': int(min_d),
                    'max_discount': int(max_d),
                    'proposed_discount': mid,
                    'missing_part': missing_part[0] if missing_part else None,
                    'missing_part_price': missing_part[1] if missing_part else None
                }
                
                # Показываем форму согласия
                st.session_state['show_agreement'] = True

    # Форма согласия клиента
    if st.session_state.get('show_agreement', False):
        st.divider()
        st.subheader("✅ Фиксация согласия клиента")
        
        with st.form("agreement_form"):
            # Выбор маркетплейса с логотипом
            marketplace_options = list(MARKETPLACES.keys())
            marketplace_labels = [f"{MARKETPLACES[m]['emoji']} {MARKETPLACES[m]['name']}" for m in marketplace_options]
            
            selected_label = st.selectbox(
                "Маркетплейс",
                options=marketplace_labels,
                help="Выберите площадку, на которой была предоставлена скидка"
            )
            # Извлекаем чистое название маркетплейса
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
                    st.success(f"✅ Согласие зафиксировано!")
                    st.success(f"📊 Данные сохранены в файл: {LOG_FILE}")
                    st.balloons()
                    
                    # Сбрасываем флаг
                    st.session_state['show_agreement'] = False
                    
                    # Показываем итоговую информацию
                    st.info(f"**Итог:** Клиент {order_number} согласился на компенсацию {calc['proposed_discount']}% на маркетплейсе {selected_marketplace}")
else:
    st.info("Введите артикул товара, чтобы начать.")
