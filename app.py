import streamlit as st
import pandas as pd
from typing import Dict, Tuple, Optional

# ------------------------------
# 1. Загрузка данных артикулов
# ------------------------------
@st.cache_data
def load_sku_data(file_path: str) -> pd.DataFrame:
    """Загружает Excel-файл с колонками Артикул, Level 02, Level 03, Level 04, Модель"""
    df = pd.read_excel(file_path, engine='openpyxl')
    # Приводим колонки к стандартным названиям
    df.columns = df.columns.str.strip()
    # Если колонка с артикулом называется 'Артикул', переименуем для удобства
    if 'Артикул' in df.columns:
        df.rename(columns={'Артикул': 'sku'}, inplace=True)
    # Убираем возможные пустые строки
    df.dropna(subset=['sku'], inplace=True)
    df['sku'] = df['sku'].astype(str).str.strip()
    return df

# Попытка загрузить файл по умолчанию
DEFAULT_FILE = "Список товарных групп.xlsx"
df_skus = None
if st.sidebar.checkbox("Загрузить свой файл артикулов"):
    uploaded_file = st.sidebar.file_uploader("Выберите Excel-файл", type=['xlsx'])
    if uploaded_file:
        df_skus = load_sku_data(uploaded_file)
        st.sidebar.success("Файл загружен")
    else:
        st.sidebar.info("Пожалуйста, загрузите файл")
else:
    try:
        df_skus = load_sku_data(DEFAULT_FILE)
        st.sidebar.success(f"Загружен файл по умолчанию: {DEFAULT_FILE}")
    except FileNotFoundError:
        st.sidebar.error("Файл по умолчанию не найден. Пожалуйста, загрузите файл вручную.")
        df_skus = None

if df_skus is None:
    st.stop()

# Создаём словарь для быстрого поиска по артикулу
sku_to_info = {}
for _, row in df_skus.iterrows():
    sku = row['sku']
    sku_to_info[sku] = {
        'level02': row.get('Level 02', ''),
        'level03': row.get('Level 03', ''),
        'level04': row.get('Level 04', ''),
        'model': row.get('Модель', '')
    }

# ------------------------------
# 2. Конфигурация чеклистов и правил скидок
# ------------------------------
# Определяем возможные вопросы
QUESTIONS = {
    'location': {
        'label': 'Где расположено повреждение?',
        'options': [
            'Лицевая часть', 'Боковая часть', 'Тыльная часть',
            'Экран', 'Корпус', 'Дверца', 'Клавиатура', 'Рамка'
        ]
    },
    'type': {
        'label': 'Тип повреждения?',
        'options': ['Царапина', 'Скол', 'Вмятина', 'Трещина', 'Отсутствует комплектующее', 'Повреждение упаковки']
    },
    'size': {
        'label': 'Размер повреждения (для царапин/сколов/трещин)?',
        'options': ['До 1 см', '1–3 см', '3–5 см', 'Более 5 см']
    },
    'functional': {
        'label': 'Влияет ли на функциональность?',
        'options': ['Да', 'Нет']
    }
}

# Настройка чеклиста для каждой группы (ключ – Level 04)
CHECKLISTS = {
    'Ноутбуки': {
        'questions': ['location', 'type', 'size', 'functional'],
        'location_options': ['Экран', 'Корпус (верх)', 'Корпус (низ)', 'Клавиатура', 'Петли', 'Торцы']
    },
    'Смартфоны': {
        'questions': ['location', 'type', 'size', 'functional'],
        'location_options': ['Экран', 'Задняя крышка', 'Рамка', 'Камера']
    },
    'Встр. Духовка': {
        'questions': ['location', 'type', 'size', 'functional'],
        'location_options': ['Лицевая часть (стекло/панель)', 'Боковая часть', 'Тыльная часть', 'Внутренняя камера']
    },
    'Холодильники': {
        'questions': ['location', 'type', 'size'],
        'location_options': ['Лицевая дверца', 'Боковая стенка', 'Внутренняя полка', 'Ручка']
    },
    'default': {
        'questions': ['location', 'type', 'size', 'functional'],
        'location_options': QUESTIONS['location']['options']
    }
}

# Правила скидок: ключ – (level04, location, type, size, functional) -> (min, max)
RULES = {
    # Ноутбуки
    ('Ноутбуки', 'Экран', 'Царапина', 'До 1 см', 'Нет'): (5, 8),
    ('Ноутбуки', 'Экран', 'Царапина', '1–3 см', 'Нет'): (8, 12),
    ('Ноутбуки', 'Экран', 'Царапина', '3–5 см', 'Нет'): (12, 18),
    ('Ноутбуки', 'Экран', 'Трещина', 'Любой', 'Нет'): (20, 30),
    ('Ноутбуки', 'Экран', 'Трещина', 'Любой', 'Да'): (40, 50),
    ('Ноутбуки', 'Корпус (верх)', 'Царапина', 'До 3 см', 'Нет'): (3, 5),
    ('Ноутбуки', 'Корпус (верх)', 'Царапина', '3–5 см', 'Нет'): (5, 8),
    ('Ноутбуки', 'Корпус (верх)', 'Вмятина', 'Любой', 'Нет'): (8, 12),
    # Смартфоны
    ('Смартфоны', 'Экран', 'Царапина', 'До 1 см', 'Нет'): (8, 12),
    ('Смартфоны', 'Экран', 'Царапина', '1–3 см', 'Нет'): (12, 18),
    ('Смартфоны', 'Экран', 'Царапина', '3–5 см', 'Нет'): (15, 20),
    ('Смартфоны', 'Экран', 'Трещина', 'Любой', 'Нет'): (25, 35),
    ('Смартфоны', 'Экран', 'Трещина', 'Любой', 'Да'): (40, 50),
    ('Смартфоны', 'Задняя крышка', 'Царапина', 'Любой', 'Нет'): (5, 10),
    ('Смартфоны', 'Рамка', 'Вмятина', 'Любой', 'Нет'): (8, 12),
    # Встраиваемая техника (духовка)
    ('Встр. Духовка', 'Лицевая часть (стекло/панель)', 'Царапина', 'До 2 см', 'Нет'): (3, 5),
    ('Встр. Духовка', 'Лицевая часть (стекло/панель)', 'Царапина', '2–5 см', 'Нет'): (7, 10),
    ('Встр. Духовка', 'Лицевая часть (стекло/панель)', 'Царапина', 'Более 5 см', 'Нет'): (12, 15),
    ('Встр. Духовка', 'Лицевая часть (стекло/панель)', 'Трещина', 'Любой', 'Нет'): (20, 30),
    ('Встр. Духовка', 'Лицевая часть (стекло/панель)', 'Трещина', 'Любой', 'Да'): (40, 50),
    ('Встр. Духовка', 'Боковая часть', 'Царапина', 'Любой', 'Нет'): (2, 4),
    ('Встр. Духовка', 'Тыльная часть', 'Царапина', 'Любой', 'Нет'): (1, 3),
    # Холодильники
    ('Холодильники', 'Лицевая дверца', 'Царапина', 'До 5 см', 'Нет'): (3, 5),
    ('Холодильники', 'Лицевая дверца', 'Царапина', 'Более 5 см', 'Нет'): (5, 8),
    ('Холодильники', 'Боковая стенка', 'Царапина', 'Любой', 'Нет'): (2, 4),
    ('Холодильники', 'Лицевая дверца', 'Вмятина', 'Любой', 'Нет'): (8, 12),
    # Общие правила (для остальных групп)
    ('default', 'Лицевая часть', 'Царапина', 'До 2 см', 'Нет'): (2, 4),
    ('default', 'Лицевая часть', 'Царапина', '2–5 см', 'Нет'): (5, 8),
    ('default', 'Лицевая часть', 'Царапина', 'Более 5 см', 'Нет'): (8, 12),
    ('default', 'Боковая часть', 'Царапина', 'Любой', 'Нет'): (1, 3),
    ('default', 'Тыльная часть', 'Царапина', 'Любой', 'Нет'): (1, 2),
}

# Дополнительная логика для отсутствия комплектующих
def handle_missing_part(part_name: str, part_price: float, product_price: float) -> Tuple[float, float]:
    """Возвращает мин/макс скидку на основе стоимости детали"""
    percent = (part_price / product_price) * 100 + 5
    percent = min(percent, 15)  # не более 15%
    return (percent, percent)

# ------------------------------
# 3. Функция для расчёта скидки
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
    if missing_part:
        return handle_missing_part(*missing_part, product_price)

    # Сначала ищем точное совпадение
    key = (level04, location, damage_type, size, functional)
    if key in RULES:
        return RULES[key]

    # Если не нашли, пробуем с заменой размера на 'Любой'
    key = (level04, location, damage_type, 'Любой', functional)
    if key in RULES:
        return RULES[key]

    # Иначе ищем в default
    key = ('default', location, damage_type, size, functional)
    if key in RULES:
        return RULES[key]

    key = ('default', location, damage_type, 'Любой', functional)
    if key in RULES:
        return RULES[key]

    # Если ничего не подошло, возвращаем 0
    return (0, 0)

# ------------------------------
# 4. Интерфейс Streamlit
# ------------------------------
st.set_page_config(page_title="Система компенсаций", layout="wide")
st.title("📦 Определение компенсации за повреждения")

# Поле ввода артикула
sku_input = st.text_input("Введите артикул товара", placeholder="Например: 500964")

if sku_input:
    sku_info = sku_to_info.get(sku_input)
    if not sku_info:
        st.error("Артикул не найден в базе. Пожалуйста, проверьте ввод.")
        st.stop()

    # Отображаем информацию о товаре
    st.subheader("Информация о товаре")
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Модель:** {sku_info['model']}")
        st.write(f"**Категория (L2):** {sku_info['level02']}")
        st.write(f"**Категория (L3):** {sku_info['level03']}")
    with col2:
        st.write(f"**Категория (L4):** {sku_info['level04']}")
        # В реальной системе можно подтянуть цену из другой таблицы, пока просим вручную
        price = st.number_input("Цена товара (тенге)", min_value=0.0, value=100000.0, step=1000.0)

    # Определяем, какой чеклист использовать
    level04 = sku_info['level04']
    checklist = CHECKLISTS.get(level04, CHECKLISTS['default'])
    st.subheader("Опишите повреждение")

    # Динамические вопросы
    answers = {}
    for q_name in checklist['questions']:
        q_data = QUESTIONS[q_name]
        if q_name == 'location':
            # Используем специфичные для группы варианты
            options = checklist.get('location_options', q_data['options'])
            answers[q_name] = st.selectbox(q_data['label'], options)
        else:
            answers[q_name] = st.selectbox(q_data['label'], q_data['options'])

    # Особый случай: отсутствие комплектующего
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
                st.warning("Для данной комбинации повреждений нет настроенных правил. Обратитесь к администратору.")
            else:
                st.success(f"**Рекомендуемая скидка:** {min_d:.0f}% – {max_d:.0f}%")
                st.info(f"**Стартовая компенсация:** {(min_d+max_d)/2:.0f}%")

                # Кнопка копирования предложения
                proposal = f"Здравствуйте! По вашему обращению о повреждении ({answers['location']}, {answers['type']}) мы можем предложить компенсацию в размере {int((min_d+max_d)/2)}% от стоимости товара. Если вы готовы оставить товар, сообщите нам."
                st.code(proposal, language='text')
                st.button("Скопировать предложение", on_click=lambda: st.write("Для копирования выделите текст выше"))

else:
    st.info("Введите артикул товара, чтобы начать.")