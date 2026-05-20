import streamlit as st
import pandas as pd
import openpyxl
import json
import os
import io
import traceback
import re
from datetime import datetime

st.set_page_config(page_title="SmartOrder Web — АлгаДистрибьюшн", page_icon="📊", layout="wide")

# ==========================================
# СИСТЕМА ЛОГИРОВАНИЯ И ОЧИСТКИ ДАННЫХ
# ==========================================
def log_error(context, exception):
    log_filename = "error_log.txt"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    error_msg = f"[{timestamp}] КРИТИЧЕСКАЯ ОШИБКА в блоке: {context}\nОписание: {str(exception)}\nТрассировка кода:\n{traceback.format_exc()}\n{'='*80}\n\n"
    with open(log_filename, "a", encoding="utf-8") as f:
        f.write(error_msg)

def parse_number(val, num_type=float):
    """Вытаскивает чистые цифры из любой ячейки (спасает от пробелов, долларов и запятых)"""
    if val is None or pd.isna(val): return num_type(0)
    s = str(val).replace(",", ".").replace(" ", "").strip()
    s = re.sub(r'[^\d\.-]', '', s)
    try: return num_type(s) if s else num_type(0)
    except: return num_type(0)

def normalize_strict(name):
    """ЖЕЛЕЗОБЕТОННАЯ нормализация. Оставляет только буквы и цифры.
    Убирает кавычки, пробелы и елочки для точного сравнения баз."""
    if not isinstance(name, str) or pd.isna(name): return ""
    return re.sub(r'[^\w\d]', '', str(name).lower())

# ==========================================
# БЛОК БЕЗОПАСНОСТИ
# ==========================================
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if st.session_state["authenticated"]: return True

    st.title("🔒 Вход в систему Alga Distribution")
    login_input = st.text_input("Логин", key="login_field")
    password_input = st.text_input("Пароль", type="password", key="password_field")
    
    if st.button("Войти"):
        try:
            target_login = st.secrets.get("ALGA_LOGIN", "alga_team")
            target_password = st.secrets.get("ALGA_PASSWORD", "alga2026")
        except:
            target_login = "alga_team"
            target_password = "alga2026"
            
        if login_input == target_login and password_input == target_password:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("❌ Неверный логин или пароль!")
    return False

if not check_password(): st.stop()

# ==========================================
# ОБЛАЧНЫЕ ФУНКЦИИ БАЗЫ ЗНАНИЙ
# ==========================================
@st.cache_data
def load_mapping_cloud():
    if os.path.exists("mapping.csv"):
        try: return pd.read_csv("mapping.csv", sep=";")
        except: pass
    return pd.DataFrame(columns=["Номенклатура АлгаДистрибьюшн факт", "Наименование от производителя"])

def save_new_pair_cloud(factory_name, our_name):
    try:
        mapping_df = load_mapping_cloud()
        new_row = pd.DataFrame([{"Номенклатура АлгаДистрибьюшн факт": our_name.strip(), "Наименование от производителя": factory_name.strip()}])
        mapping_df = pd.concat([mapping_df, new_row], ignore_index=True)
        mapping_df.to_csv("mapping.csv", sep=";", index=False, encoding="utf-8")
        st.cache_data.clear() 
    except Exception as e: log_error("save_new_pair", e)

# ==========================================
# ОБНОВЛЕННЫЙ ПАРСЕР ЦЕН ШАБЛОНА (БЕЗУСЛОВНЫЙ)
# ==========================================
def check_and_cache_template():
    """Собирает ВСЕ строки, наследует любую найденную цену строго сверху вниз"""
    if not os.path.exists("template.xlsx"): return
    try:
        wb = openpyxl.load_workbook("template.xlsx", data_only=True)
        cache_data = {"products": []}

        if "заказ-order" in wb.sheetnames:
            ws = wb["заказ-order"]
            current_price = 0.0

            for row in ws.iter_rows(min_row=5, values_only=True):
                if not row or len(row) < 7: continue
                name = str(row[1]).strip() if row[1] else ""
                
                if not name: continue

                # Проверяем цену в текущей строке
                parsed_price = parse_number(row[6], float)
                
                # Если нашли любую цену > 0 — это новая базовая цена для идущих ниже товаров
                if parsed_price > 0:
                    current_price = parsed_price 
                
                # Если у текущей строки нет цены, берем унаследованную сверху
                item_price = parsed_price if parsed_price > 0 else current_price
                box_size = parse_number(row[2], int)

                cache_data["products"].append({
                    "factory_name": name,
                    "norm_name": normalize_strict(name),
                    "box_size": box_size,
                    "price": item_price
                })

        with open("template_cache.json", "w", encoding="utf-8") as f:
            json.dump(cache_data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        log_error("check_and_cache_template", e)

# ==========================================
# ИНТЕРФЕЙС
# ==========================================
col_t1, col_t2 = st.columns([9, 1])
with col_t1: st.title("📊 Умный шлюз заказов SmartOrder Web")
with col_t2:
    if st.button("🚪 Выйти"):
        st.session_state["authenticated"] = False
        st.rerun()

tab1, tab2 = st.tabs(["📥 Расшифровка входящего заказа", "📤 Набрать новый заказ для завода"])

# ------------------------------------------
# ВКЛАДКА 1: ОБРАБОТКА ПОСТУПИВШЕГО ЗАКАЗА
# ------------------------------------------
with tab1:
    st.header("Загрузка и очистка файла от завода")
    uploaded_file = st.file_uploader("Перетащите сюда Excel или CSV файл заказа", type=["xlsx", "xls", "csv"])
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'): df_raw = pd.read_csv(uploaded_file, header=None, sep=";")
            else: df_raw = pd.read_excel(uploaded_file, header=None)
                
            mapping_df = load_mapping_cloud()
            existing_1c_items = sorted(mapping_df["Номенклатура АлгаДистрибьюшн факт"].dropna().unique().tolist())
            mapping_dict = {normalize_strict(row["Наименование от производителя"]): str(row["Номенклатура АлгаДистрибьюшн факт"]).strip() for _, row in mapping_df.iterrows()}
            
            current_price = 0.0
            processed_rows = []
            unknown_products = []
            
            for idx, row in df_raw.iloc[4:].iterrows():
                if len(row) < 7: continue
                name = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                
                name_lower = name.lower()
                if not name or "всего" in name_lower or "итого" in name_lower or "общая стоимость" in name_lower: 
                    continue

                box_size = parse_number(row.iloc[2] if len(row) > 2 else 0, int)
                boxes = parse_number(row.iloc[4] if len(row) > 4 else 0, int)
                parsed_price = parse_number(row.iloc[6] if len(row) > 6 else None, float)

                if parsed_price > 0:
                    current_price = parsed_price 
                
                # ЖЕСТКИЙ ФИЛЬТР: категория или пустая строка — мимо
                if box_size == 0 or boxes == 0:
                    continue 

                pcs = parse_number(row.iloc[3] if len(row) > 3 else 0, int)
                item_price = parsed_price if parsed_price > 0 else current_price
                    
                factory_norm_name = normalize_strict(name)
                
                if factory_norm_name not in mapping_dict:
                    if not any(p['norm_name'] == factory_norm_name for p in unknown_products):
                        unknown_products.append({"name": name, "norm_name": factory_norm_name, "row_num": idx + 1})
                        
                our_name = mapping_dict.get(factory_norm_name, f"[НЕОПРЕДЕЛЕН] {name}")
                processed_rows.append({
                    "Ваша Номенклатура (1С)": our_name, 
                    "Наименование Завода": name, 
                    "Ящиков": boxes, 
                    "Штук в ящ": box_size, 
                    "Всего штук": pcs, 
                    "Цена ($)": item_price, 
                    "Сумма ($)": pcs * item_price
                })
            
            if unknown_products:
                st.error(f"⚠️ Найдено {len(unknown_products)} новых позиций в файле.")
                for i, prod_info in enumerate(unknown_products[:3]):
                    st.info(f"🏭 В файле заказа: `{prod_info['name']}`")
                    selected_1c_name = st.selectbox("Связать с номенклатурой 1С:", options=["-- Выбрать из существующих в 1С --"] + existing_1c_items, key=f"sel_{i}")
                    manual_1c_name = st.text_input("Или ввести новое имя вручную:", key=f"man_{i}")
                    
                    if st.button("Запомнить связку", key=f"btn_{i}"):
                        final_1c_name = manual_1c_name.strip() if manual_1c_name.strip() else (selected_1c_name if selected_1c_name != "-- Выбрать из существующих в 1С --" else "")
                        if final_1c_name:
                            save_new_pair_cloud(prod_info['name'], final_1c_name)
                            st.success(f"✅ Связка зафиксирована! Обновите страницу.")
                            st.rerun()
            else:
                if processed_rows:
                    df_res = pd.DataFrame(processed_rows)
                    st.markdown("### 📋 Итоги распознавания заказа:")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Распознано SKU", len(df_res))
                    c2.metric("Всего коробок", int(df_res["Ящиков"].sum()))
                    c3.metric("Общая сумма", f'{df_res["Сумма ($)"].sum():,.2f} $')
                    
                    st.dataframe(df_res, use_container_width=True)
                    
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer: 
                        df_res.to_excel(writer, index=False)
                    st.download_button(label="💾 Скачать очищенный Excel для 1С", data=buffer.getvalue(), file_name="Обработанный_Заказ_АлгаWeb.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            log_error("Вкладка1_Парсинг", e)
            st.error("Ошибка обработки файла.")

# ------------------------------------------
# ВКЛАДКА 2: НАБОР НОВОГО ЗАКАЗА ДЛЯ ЗАВОДА
# ------------------------------------------
with tab2:
    st.header("Единый бланк заказа")
    st.write("Скролл зафиксирован. Заполняйте ящики в таблице — корзина ниже сформируется автоматически.")
    
    if st.button("🔄 Сбросить и обновить справочник цен из template.xlsx"):
        with st.spinner("Перечитываем и пересчитываем файл шаблона..."):
            check_and_cache_template()
            st.success("База цен успешно обновлена!")
            st.rerun()

    if not os.path.exists("template_cache.json"):
        with st.spinner("Собираем базу данных..."):
            check_and_cache_template()
            
    if os.path.exists("template_cache.json"):
        try:
            with open("template_cache.json", "r", encoding="utf-8") as f:
                cache = json.load(f)
            
            factory_cache_dict = {p["norm_name"]: {"box_size": p["box_size"], "price": p["price"]} for p in cache["products"]}
                
            mapping_df = load_mapping_cloud()
            mapping_clean = mapping_df.dropna(subset=["Номенклатура АлгаДистрибьюшн факт", "Наименование от производителя"])
            mapping_clean = mapping_clean.drop_duplicates(subset=["Номенклатура АлгаДистрибьюшн факт"])
            mapping_clean = mapping_clean.sort_values(by="Номенклатура АлгаДистрибьюшн факт")
            
            table_rows = []
            for _, row in mapping_clean.iterrows():
                our_name = str(row["Номенклатура АлгаДистрибьюшн факт"]).strip()
                f_name = str(row["Наименование от производителя"]).strip()
                f_name_key = normalize_strict(f_name)
                
                factory_info = factory_cache_dict.get(f_name_key, {"box_size": 0, "price": 0.0})
                
                table_rows.append({
                    "Номенклатура (1С)": our_name,
                    "Товар завода": f_name,
                    "В коробке (шт)": int(factory_info["box_size"]),
                    "Цена ($)": float(factory_info["price"]),
                    "Заказ (Ящиков)": 0
                })
                
            df_form = pd.DataFrame(table_rows)
            
            if df_form.empty:
                st.warning("⚠️ В файле mapping.csv пока нет связей.")
                st.stop()
            
            # ТАБЛИЦА ВВОДА (Чистая, без прыжков)
            edited_df = st.data_editor(
                df_form,
                column_config={
                    "Номенклатура (1С)": st.column_config.TextColumn(disabled=True, width="large"),
                    "Товар завода": st.column_config.TextColumn(disabled=True, width="medium"),
                    "В коробке (шт)": st.column_config.NumberColumn(disabled=True, width="small"),
                    "Цена ($)": st.column_config.NumberColumn(disabled=True, format="$ %.2f", width="small"),
                    "Заказ (Ящиков)": st.column_config.NumberColumn(min_value=0, max_value=5000, step=1, required=True, width="small")
                },
                disabled=["Номенклатура (1С)", "Товар завода", "В коробке (шт)", "Цена ($)"],
                use_container_width=True,
                key="super_stable_editor", 
                hide_index=True
            )

            # МГНОВЕННЫЙ РАСЧЕТ ИТОГОВ КОРЗИНЫ
            edited_df["Заказ (Ящиков)"] = pd.to_numeric(edited_df["Заказ (Ящиков)"]).fillna(0).astype(int)
            edited_df["В коробке (шт)"] = pd.to_numeric(edited_df["В коробке (шт)"]).fillna(0).astype(int)
            edited_df["Цена ($)"] = pd.to_numeric(edited_df["Цена ($)"]).fillna(0.0).astype(float)
            
            edited_df["Итого штук"] = edited_df["Заказ (Ящиков)"] * edited_df["В коробке (шт)"]
            edited_df["Итого сумма ($)"] = edited_df["Итого штук"] * edited_df["Цена ($)"]
            
            active_orders = edited_df[edited_df["Заказ (Ящиков)"] > 0]

            st.markdown("---")
            st.subheader("🛒 Ваша корзина (формируется автоматически):")

            if not active_orders.empty:
                c_m1, c_m2, c_m3 = st.columns(3)
                c_m1.metric("Позиций (SKU)", len(active_orders))
                c_m2.metric("Всего ящиков", int(active_orders["Заказ (Ящиков)"].sum()))
                c_m3.metric("Общая сумма ($)", f'{active_orders["Итого сумма ($)"].sum():,.2f} $')
                
                st.dataframe(
                    active_orders[["Номенклатура (1С)", "Заказ (Ящиков)", "Итого штук", "Цена ($)", "Итого сумма ($)"]], 
                    use_container_width=True, 
                    hide_index=True
                )
                
                if st.button("🚀 Подготовить и скачать файл для завода", type="primary"):
                    if not os.path.exists("template.xlsx"):
                        st.error("Отсутствует оригинальный шаблон template.xlsx")
                    else:
                        with st.spinner("Создаем Excel-файл..."):
                            try:
                                input_data = {normalize_strict(k): v for k, v in zip(active_orders["Товар завода"], active_orders["Заказ (Ящиков)"])}
                                wb = openpyxl.load_workbook("template.xlsx", data_only=False)
                                
                                if "заказ-order" in wb.sheetnames:
                                    ws = wb["заказ-order"]
                                    for row_idx in range(5, ws.max_row + 1):
                                        name_val = ws.cell(row=row_idx, column=2).value
                                        if name_val is None: continue
                                        
                                        f_name_key = normalize_strict(name_val)
                                        if f_name_key in input_data:
                                            boxes_to_order = input_data[f_name_key]
                                            if boxes_to_order > 0:
                                                ws.cell(row=row_idx, column=5).value = boxes_to_order
                                            else:
                                                ws.cell(row=row_idx, column=5).value = None
                                                
                                out_buf = io.BytesIO()
                                wb.save(out_buf)
                                out_buf.seek(0)
                                st.session_state["excel_ready_bytes"] = out_buf.getvalue()
                            except Exception as e:
                                log_error("Генерация_Завода", e)
                                st.error("Ошибка записи данных в Excel.")

                if st.session_state.get("excel_ready_bytes"):
                    st.success("🎉 Файл успешно сгенерирован!")
                    st.download_button(
                        label="📥 СКАЧАТЬ ГОТОВЫЙ ФАЙЛ ДЛЯ ЗАВОДА",
                        data=st.session_state["excel_ready_bytes"],
                        file_name="Сформированный_Заказ_Завод_Web.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.info("Корзина пуста. Начните вводить количество ящиков в таблице выше.")
                if "excel_ready_bytes" in st.session_state:
                    st.session_state["excel_ready_bytes"] = None
            
        except Exception as e:
            log_error("Вкладка2_Рендеринг", e)
            st.error("Произошла ошибка при отрисовке бланка заказа.")
