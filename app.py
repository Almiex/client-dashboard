import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import re
from datetime import datetime

# ==============================================================================
# НАСТРОЙКА СТРАНИЦЫ
# ==============================================================================
st.set_page_config(page_title="Анализ аудитории клиники", layout="wide")

st.markdown("""
    <style>
    .clinic-header-audience {
        font-family: 'Segoe UI', Arial, sans-serif;
        margin-bottom: 25px;
        padding: 15px;
        background: #FFFFFF;
        border-left: 6px solid #9E6B75;
        border-radius: 4px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        border: 1px solid #EAEAEA;
    }
    .clinic-title-audience {
        font-size: 22px;
        font-weight: 700;
        color: #2B2D42;
        letter-spacing: 0.5px;
    }
    .clinic-subtitle-audience {
        font-size: 14px;
        color: #6C757D;
        margin-top: 6px;
        font-weight: 500;
    }
    .kpi-card-audience {
        flex: 1;
        min-width: 180px;
        background: #FFFFFF;
        border-left: 5px solid #9E6B75;
        border-radius: 6px;
        padding: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        border-top: 1px solid #EAEAEA;
        border-right: 1px solid #EAEAEA;
        border-bottom: 1px solid #EAEAEA;
    }
    .kpi-label {
        color: #6C757D;
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
    }
    .kpi-value {
        font-size: 22px;
        font-weight: 700;
        color: #2B2D42;
        margin: 4px 0;
    }
    .kpi-sub {
        font-size: 11px;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🏥 Анализ аудитории пациентов")
st.write("Загрузите Excel-выгрузку по пациентам для построения портрета аудитории.")

uploaded_file = st.file_uploader("Выберите Excel файл (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # =========================================================================
        # 1. ЧТЕНИЕ ФАЙЛА БЕЗ ЗАГОЛОВКОВ
        # =========================================================================
        df_all = pd.read_excel(uploaded_file, sheet_name=0, header=None)

        # =========================================================================
        # 2. ДИНАМИЧЕСКИЙ ПОИСК СТРОКИ С ЗАГОЛОВКАМИ
        # =========================================================================
        header_row_index = None
        for idx, row in df_all.iterrows():
            # ГАРАНТИРОВАННО преобразуем каждую ячейку в строку, NaN → пустая строка
            cells = []
            for val in row.values:
                if pd.notna(val):
                    cells.append(str(val).lower())
                else:
                    cells.append("")
            
            has_id = any('пациент' in s or 'айди' in s or 'id' in s for s in cells)
            has_services_or_money = any('усл' in s or 'кол' in s or 'оплат' in s or 'сумм' in s for s in cells)
            
            if has_id and has_services_or_money:
                header_row_index = idx
                break

        if header_row_index is None:
            st.error("❌ Не удалось автоматически найти строку заголовков таблицы.")
            st.stop()

        # =========================================================================
        # 3. ИЗВЛЕЧЕНИЕ МЕТАДАННЫХ
        # =========================================================================
        clinic_name_str = "КЛИНИКА"
        for col in df_all.columns:
            cells_to_check = [str(col)] + [str(v) for v in df_all[col].values if pd.notna(v)]
            for cell in cells_to_check:
                if 'клиника:' in cell.lower():
                    match = re.search(r'клиника:\s*\d*[\s,]*"?([^"\n;]+)"?', cell, flags=re.IGNORECASE)
                    if match:
                        clinic_name_str = match.group(1).strip().replace('"', '').replace('«', '').replace('»', '').strip()
                        break
            if clinic_name_str != "КЛИНИКА":
                break

        all_dates = []
        for col in df_all.columns:
            text_content = " ".join([str(col)] + [str(v) for v in df_all[col].values if pd.notna(v)])
            all_dates.extend(re.findall(r'\d{2}\.\d{2}\.\d{4}', text_content))

        date_past_str = all_dates[0] if len(all_dates) > 0 else ""
        date_curr_str = all_dates[1] if len(all_dates) > 1 else date_past_str

        # =========================================================================
        # 4. ОЧИСТКА КОЛОНОК И МАППИНГ
        # =========================================================================
        raw_cols = []
        for val in df_all.iloc[header_row_index].values:
            if pd.notna(val):
                s = str(val)
                s = re.sub(r'~00\d*', '', s)  # убираем ~000, ~001, ~non-000 и т.д.
                raw_cols.append(s.strip())
            else:
                raw_cols.append('')

        df_clean = df_all.iloc[header_row_index + 1:].copy()
        df_clean.columns = raw_cols
        df_clean = df_clean.reset_index(drop=True)
        df_clean = df_clean.dropna(how='all')

        col_id = col_gender = col_birth = col_count = col_money = col_city = None

        for col_name in raw_cols:
            c_low = col_name.lower()
            if 'пациент' in c_low or 'айди' in c_low or 'id' in c_low:
                col_id = col_name
            elif 'пол' in c_low:
                col_gender = col_name
            elif 'рожд' in c_low or 'возраст' in c_low or 'дата рожд' in c_low:
                col_birth = col_name
            elif 'усл' in c_low or 'кол' in c_low:
                col_count = col_name
            elif 'оплат' in c_low or 'сумм' in c_low or 'цена' in c_low:
                col_money = col_name
            elif 'город' in c_low:
                col_city = col_name

        if not all([col_id, col_count, col_money]):
            st.error(f"❌ Ошибка маппинга колонок. Найдено: ID={col_id}, Услуги={col_count}, Деньги={col_money}")
            st.info(f"Доступные колонки после очистки: {raw_cols}")
            st.stop()

        # =========================================================================
        # 5. ОБРАБОТКА ДАННЫХ
        # =========================================================================
        df_clean[col_count] = pd.to_numeric(df_clean[col_count], errors='coerce').fillna(0).astype(int)
        df_clean[col_money] = pd.to_numeric(df_clean[col_money], errors='coerce').fillna(0)

        # Возраст — точно как в оригинальном коде
        if col_birth:
            df_clean['Parsed_Birth'] = pd.to_datetime(df_clean[col_birth], errors='coerce', format='mixed')
            current_year = datetime.now().year
            df_clean['Возраст'] = current_year - df_clean['Parsed_Birth'].dt.year
            df_clean['Возраст'] = df_clean['Возраст'].fillna(0).astype(int)
        else:
            df_clean['Возраст'] = 0

        # Пол
        if col_gender:
            df_clean[col_gender] = df_clean[col_gender].fillna('Не указан').astype(str).str.strip()
        else:
            df_clean['Пол'] = 'Не указан'
            col_gender = 'Пол'

        # Город
        if col_city:
            df_clean[col_city] = df_clean[col_city].fillna('Не указано').astype(str).str.strip()
            df_clean[col_city] = df_clean[col_city].replace(['', 'nan', 'None'], 'Не указано')
        else:
            df_clean['Город'] = 'Не указано'
            col_city = 'Город'

        # =========================================================================
        # 6. АГРЕГАЦИЯ ПО ПАЦИЕНТАМ
        # =========================================================================
        df_patients_report = df_clean.groupby(col_id).agg({
            col_gender: 'first',
            'Возраст': 'first',
            col_city: 'first',
            col_count: 'sum',
            col_money: 'sum'
        }).reset_index()
        df_patients_report.columns = ['ID Пациента', 'Пол', 'Возраст', 'Город', 'Количество услуг', 'LTV сумма']

        def get_age_cohort(age):
            if age <= 17: return '0-17 (Дети/Подростки)'
            elif age <= 35: return '18-35 (Молодежь)'
            elif age <= 60: return '36-60 (Взрослые)'
            else: return '61+ (Пожилые)'

        def get_loyalty_segment(count):
            if count == 1: return 'Разовые визиты (1 услуга)'
            elif count <= 5: return 'Постоянные (2-5 услуг)'
            else: return 'Супер-Лояльные (6+ услуг)'

        df_patients_report['Возрастная группа'] = df_patients_report['Возраст'].apply(get_age_cohort)
        df_patients_report['Сегмент лояльности'] = df_patients_report['Количество услуг'].apply(get_loyalty_segment)

        # =========================================================================
        # 7. РАСЧЁТ KPI
        # =========================================================================
        total_unique = len(df_patients_report)
        total_revenue = df_patients_report['LTV сумма'].sum()
        total_services = df_patients_report['Количество услуг'].sum()
        avg_ltv = total_revenue / total_unique if total_unique > 0 else 0
        avg_services = total_services / total_unique if total_unique > 0 else 0
        avg_age = df_patients_report[df_patients_report['Возраст'] > 0]['Возраст'].mean()
        if pd.isna(avg_age):
            avg_age = 0

        gender_counts = df_patients_report['Пол'].value_counts()
        women_pct = (gender_counts.get('жен', 0) / total_unique * 100) if total_unique > 0 else 0
        men_pct = (gender_counts.get('муж', 0) / total_unique * 100) if total_unique > 0 else 0

        # =========================================================================
        # 8. ШАПКА И KPI-КАРТОЧКИ
        # =========================================================================
        st.markdown(f"""
            <div class="clinic-header-audience">
                <div class="clinic-title-audience">🏥 ПОРТРЕТ АУДИТОРИИ ПАЦИЕНТОВ: {clinic_name_str.upper()}</div>
                <div class="clinic-subtitle-audience">
                    🎯 Демография, плотность потребления медицинских услуг и LTV
                    <span style="color:#9E6B75;font-weight:700;margin-left:10px;">(Период: с {date_past_str} по {date_curr_str})</span>
                </div>
            </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
            <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:25px;">
                <div class="kpi-card-audience" style="border-left-color:#9E6B75;">
                    <div class="kpi-label">Уникальных пациентов</div>
                    <div class="kpi-value">{total_unique:,} пац.</div>
                    <div class="kpi-sub" style="color:#9E6B75;">Женщины: {women_pct:.1f}%</div>
                    <div class="kpi-sub" style="color:#005F73;">Мужчины: {men_pct:.1f}%</div>
                </div>
                <div class="kpi-card-audience" style="border-left-color:#005F73;">
                    <div class="kpi-label">Общая выручка базы</div>
                    <div class="kpi-value">{total_revenue:,.0f} ₽</div>
                    <div class="kpi-sub" style="color:#005F73;">Выручка от реализации услуг</div>
                </div>
                <div class="kpi-card-audience" style="border-left-color:#F4A261;">
                    <div class="kpi-label">Выручка на 1 пац. (ARPU)</div>
                    <div class="kpi-value">{avg_ltv:,.0f} ₽</div>
                    <div class="kpi-sub" style="color:#4A4A4A;">Средний доход с пациента</div>
                </div>
                <div class="kpi-card-audience" style="border-left-color:#005F73;">
                    <div class="kpi-label">Услуг на пациента</div>
                    <div class="kpi-value">{avg_services:.1f} усл.</div>
                    <div class="kpi-sub" style="color:#6C757D;">Всего услуг: {total_services:,}</div>
                </div>
                <div class="kpi-card-audience" style="border-left-color:#E9C46A;">
                    <div class="kpi-label">Средний возраст</div>
                    <div class="kpi-value">{avg_age:.1f} лет</div>
                    <div class="kpi-sub" style="color:#6C757D;">Активный фокус</div>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # =========================================================================
        # 9. ГРАФИК 1: ГЕНДЕРНАЯ СТРУКТУРА
        # =========================================================================
        st.subheader("1. Распределение пациентов по Полу")
        df_gender = df_patients_report.groupby('Пол').agg({'ID Пациента': 'count', 'LTV сумма': 'sum'}).reset_index()
        p1 = px.pie(
            df_gender, values='ID Пациента', names='Пол', color='Пол',
            color_discrete_map={'жен': '#9E6B75', 'муж': '#005F73', 'Не указан': '#E0E0E0'}
        )
        p1.update_layout(height=400, template="plotly_white")
        p1.update_traces(hovertemplate="<b>Пол: %{label}</b><br>Количество: %{value} чел.<br>Доля: %{percent}<extra></extra>")
        st.plotly_chart(p1, use_container_width=True)

        # =========================================================================
        # 10. ГРАФИК 2: ДЕМОГРАФИЯ ПО ВОЗРАСТУ
        # =========================================================================
        st.subheader("2. Плотность распределения аудитории по Возрасту")
        df_age_filtered = df_patients_report[df_patients_report['Возраст'] > 0]
        p2 = px.histogram(
            df_age_filtered, x='Возраст', color='Пол', nbins=40, barmode='group',
            color_discrete_map={'жен': '#9E6B75', 'муж': '#005F73', 'Не указан': '#E0E0E0'}
        )
        p2.update_layout(xaxis_title="Возраст (лет)", yaxis_title="Количество человек", template="plotly_white", height=450)
        p2.update_traces(hovertemplate="<b>Пол: %{fullData.name}</b><br>Возраст: %{x} лет<br>Количество: %{y} чел.<extra></extra>")
        st.plotly_chart(p2, use_container_width=True)

        # =========================================================================
        # 11. ГРАФИК 3: СТРУКТУРА ЛОЯЛЬНОСТИ (ВЫРУЧКА)
        # =========================================================================
        st.subheader("3. Категории лояльности клиентов: Вклад сегментов в общую выручку")
        df_loyalty = df_patients_report.groupby('Сегмент лояльности').agg({'ID Пациента': 'count', 'LTV сумма': 'sum'}).reset_index()
        p3 = px.bar(
            df_loyalty, x='LTV сумма', y='Сегмент лояльности', orientation='h',
            color='Сегмент лояльности', color_discrete_sequence=['#F4A261', '#9E6B75', '#005F73']
        )
        p3.update_layout(xaxis_title="Суммарная выручка сегмента (₽)", yaxis_title="", showlegend=False, template="plotly_white", height=350)
        p3.update_traces(texttemplate="%{x:,.0f} ₽", textposition="outside", hovertemplate="<b>%{y}</b><br>Сумма оплат: %{x:,.0f} ₽<extra></extra>")
        st.plotly_chart(p3, use_container_width=True)

        # # =========================================================================
        # # 12. ГРАФИК 4: ЛОЯЛЬНОСТЬ В РАЗРЕЗЕ ПОЛА
        # # =========================================================================
        # st.subheader("4. Распределение категорий лояльности по Полу")
        # df_loyalty_gender = df_patients_report.groupby(['Сегмент лояльности', 'Пол']).agg({'ID Пациента': 'count'}).reset_index()
        # df_loyalty_gender.rename(columns={'ID Пациента': 'Пациенты'}, inplace=True)

        # loyalty_order = ['Разовые визиты (1 услуга)', 'Постоянные (2-5 услуг)', 'Супер-Лояльные (6+ услуг)']
        # df_loyalty_gender['Сегмент лояльности'] = pd.Categorical(df_loyalty_gender['Сегмент лояльности'], categories=loyalty_order, ordered=True)
        # df_loyalty_gender = df_loyalty_gender.sort_values('Сегмент лояльности')

        # p4 = px.bar(
        #     df_loyalty_gender, x='Пациенты', y='Сегмент лояльности', color='Пол', barmode='group', orientation='h',
        #     color_discrete_map={'жен': '#9E6B75', 'муж': '#005F73', 'Не указан': '#E0E0E0'}
        # )
        # p4.update_layout(
        #     xaxis_title="Количество человек", yaxis_title="", template="plotly_white", height=400,
        #     legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        # )
        # p4.update_traces(hovertemplate="<b>Сегмент: %{y}</b><br>Пол: %{fullData.name}<br>Количество: %{x} чел.<extra></extra>")
        # st.plotly_chart(p4, use_container_width=True)

        # =========================================================================
        # 12. ГРАФИК 4: ЛОЯЛЬНОСТЬ В РАЗРЕЗЕ ПОЛА (STACKED + ИТОГИ СПРАВА)
        # =========================================================================
        st.subheader("4. Распределение категорий лояльности по Полу")
        df_loyalty_gender = df_patients_report.groupby(['Сегмент лояльности', 'Пол']).agg({'ID Пациента': 'count'}).reset_index()
        df_loyalty_gender.rename(columns={'ID Пациента': 'Пациенты'}, inplace=True)

        loyalty_order = ['Разовые визиты (1 услуга)', 'Постоянные (2-5 услуг)', 'Супер-Лояльные (6+ услуг)']
        df_loyalty_gender['Сегмент лояльности'] = pd.Categorical(df_loyalty_gender['Сегмент лояльности'], categories=loyalty_order, ordered=True)
        df_loyalty_gender = df_loyalty_gender.sort_values('Сегмент лояльности')

        p4 = px.bar(
            df_loyalty_gender, x='Пациенты', y='Сегмент лояльности', color='Пол', 
            barmode='stack', orientation='h',
            color_discrete_map={'жен': '#9E6B75', 'муж': '#005F73', 'Не указан': '#E0E0E0'}
        )
        p4.update_layout(
            xaxis_title="Количество человек", 
            yaxis_title="", 
            template="plotly_white", 
            height=400,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(r=120)  # ← запас справа для подписей
        )
        p4.update_traces(
            hovertemplate="<b>Сегмент: %{y}</b><br>Пол: %{fullData.name}<br>Количество: %{x} чел.<extra></extra>"
        )

        # Подписи с общим количеством пациентов в каждой когорте (справа от полосы)
        total_per_segment = df_loyalty_gender.groupby('Сегмент лояльности', observed=False)['Пациенты'].sum().reset_index()
        for _, row in total_per_segment.iterrows():
            p4.add_annotation(
                x=row['Пациенты'],
                y=row['Сегмент лояльности'],
                text=f"<b>{row['Пациенты']:,} чел.</b>",
                showarrow=False,
                xanchor='left',
                yanchor='middle',
                xshift=12,
                font=dict(size=13, color='#2B2D42', family='Segoe UI, sans-serif'),
                bgcolor='rgba(255,255,255,0.8)',
                borderpad=3
            )

        st.plotly_chart(p4, use_container_width=True)
        
        # =========================================================================
        # 13. ГРАФИК 5: ЛОЯЛЬНОСТЬ В РАЗРЕЗЕ ВОЗРАСТА
        # =========================================================================
        st.subheader("5. Распределение категорий лояльности по Возрастным группам")
        df_loyalty_age = df_patients_report.groupby(['Сегмент лояльности', 'Возрастная группа']).agg({'ID Пациента': 'count'}).reset_index()
        df_loyalty_age.rename(columns={'ID Пациента': 'Пациенты'}, inplace=True)

        age_order = ['0-17 (Дети/Подростки)', '18-35 (Молодежь)', '36-60 (Взрослые)', '61+ (Пожилые)']
        df_loyalty_age['Возрастная группа'] = pd.Categorical(df_loyalty_age['Возрастная группа'], categories=age_order, ordered=True)
        df_loyalty_age['Сегмент лояльности'] = pd.Categorical(df_loyalty_age['Сегмент лояльности'], categories=loyalty_order, ordered=True)
        df_loyalty_age = df_loyalty_age.sort_values(['Сегмент лояльности', 'Возрастная группа'])

        p5 = px.bar(
            df_loyalty_age, x='Пациенты', y='Сегмент лояльности', color='Возрастная группа', 
            barmode='group', orientation='h',
            color_discrete_sequence=['#F4A261', '#E9C46A', '#9E6B75', '#005F73']
        )
        p5.update_layout(
            xaxis_title="Количество человек", yaxis_title="", template="plotly_white", height=450,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, title="Возраст:")
        )
        p5.update_traces(hovertemplate="<b>Сегмент: %{y}</b><br>Группа: %{fullData.name}<br>Количество: %{x} чел.<extra></extra>")
        st.plotly_chart(p5, use_container_width=True)
        
        # =========================================================================
        # 14. ГРАФИК 6: ТОП-10 ГОРОДОВ (БАР)
        # =========================================================================
        st.subheader("6. Топ-10 городов по количеству пациентов")
        df_city = df_patients_report.copy()
        df_city['Город'] = df_city['Город'].replace({'Не указано': 'нет данных', '': 'нет данных', 'nan': 'нет данных', 'None': 'нет данных'}).fillna('нет данных').astype(str).str.strip()

        df_city_agg = df_city.groupby('Город').agg({'ID Пациента': 'count', 'LTV сумма': 'sum', 'Количество услуг': 'sum'}).reset_index()
        df_city_agg.columns = ['Город', 'Пациенты', 'Выручка', 'Услуги']
        df_city_agg['ARPU'] = (df_city_agg['Выручка'] / df_city_agg['Пациенты'].replace(0, np.nan)).round(0).fillna(0)

        df_real = df_city_agg[df_city_agg['Город'] != 'нет данных'].copy()
        df_top10_bar = df_real.nlargest(10, 'Пациенты').sort_values('Пациенты', ascending=True)

        p6 = px.bar(df_top10_bar, x='Пациенты', y='Город', orientation='h', text='Пациенты', color_discrete_sequence=['#005F73'])
        p6.update_layout(
            xaxis_title="Количество пациентов", yaxis_title="", template="plotly_white", height=450,
            showlegend=False, margin=dict(l=150, r=40, t=80, b=40)
        )
        p6.update_traces(
            texttemplate="%{text} чел.", textposition="outside", cliponaxis=False,
            hovertemplate="<b>%{y}</b><br>Пациентов: %{x} чел.<br>Выручка: %{customdata[0]:,.0f} ₽<br>ARPU: %{customdata[1]:,.0f} ₽<extra></extra>",
            customdata=np.stack((df_top10_bar['Выручка'], df_top10_bar['ARPU']), axis=-1)
        )
        st.plotly_chart(p6, use_container_width=True)
        st.caption("*учитывались только пациенты с указанным городом")

        # =========================================================================
        # 15. ГРАФИК 7: СТРУКТУРА ПО ГОРОДАМ (ПИРОГ)
        # =========================================================================
        st.subheader("7. Структура аудитории по городам (Топ-10)")
        df_no_data = df_city_agg[df_city_agg['Город'] == 'нет данных'].copy()
        df_with_data = df_city_agg[df_city_agg['Город'] != 'нет данных'].copy()
        df_top10_pie = df_with_data.nlargest(10, 'Пациенты').copy()
        cities_top10_set = set(df_top10_pie['Город'])
        df_others = df_with_data[~df_with_data['Город'].isin(cities_top10_set)].copy()

        frames = [df_top10_pie]
        if len(df_others) > 0:
            others_row = pd.DataFrame({
                'Город': ['Другие'],
                'Пациенты': [df_others['Пациенты'].sum()],
                'Выручка': [df_others['Выручка'].sum()],
                'Услуги': [df_others['Услуги'].sum()]
            })
            others_row['ARPU'] = (others_row['Выручка'] / others_row['Пациенты'].replace(0, np.nan)).round(0).fillna(0)
            frames.append(others_row)
        if len(df_no_data) > 0:
            frames.append(df_no_data)

        df_pie_final = pd.concat(frames, ignore_index=True)

        color_map = {'Другие': '#9CA3AF', 'нет данных': '#D1D5DB'}
        p7 = px.pie(
            df_pie_final, values='Пациенты', names='Город', hole=0.4,
            color_discrete_sequence=['#9E6B75', '#005F73', '#F4A261', '#E9C46A', '#2B2D42',
                                     '#6C757D', '#A8DADC', '#457B9D', '#E63946', '#D4A373'],
            color_discrete_map=color_map
        )
        p7.update_layout(
            template="plotly_white", height=520,
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05, font=dict(size=12), title=dict(text="Город", font=dict(size=13))),
            margin=dict(l=40, r=200, t=80, b=40)
        )
        p7.update_traces(
            textinfo='percent+label', textposition='outside',
            pull=[0.04 if city in ['Другие', 'нет данных'] else 0 for city in df_pie_final['Город']],
            hovertemplate="<b>%{label}</b><br>Пациентов: %{value} чел.<br>Доля: %{percent}<br>Выручка: %{customdata[0]:,.0f} ₽<extra></extra>",
            customdata=np.stack((df_pie_final['Выручка'],), axis=-1), textfont=dict(size=11)
        )
        st.plotly_chart(p7, use_container_width=True)
        st.caption('*учитывались все пациенты. Если город не указан — попадает в группу "Нет данных"')

    except Exception as e:
        st.error(f"❌ Ошибка при обработке файла: {e}")
        st.exception(e)
