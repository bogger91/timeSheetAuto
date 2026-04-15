"""
Читает Excel из FineBI и строит двухуровневую сводную таблицу:
  Управление → Отдел | Capacity, ч | Потрачено, ч | % списания

load_pivot()     — возвращает DataFrame со строками трёх типов:
                   row_type = "group"   — строка Управления (суммарная)
                   row_type = "dept"    — строка Отдела (дочерняя)
                   row_type = "total"   — итоговая строка

pivot_to_html()  — HTML-таблица для вставки в тело письма (только отделы + итог)
"""
import pandas as pd
import config


def load_pivot(excel_path: str | None = None) -> pd.DataFrame:
    path = excel_path or config.EXCEL_PATH

    df = pd.read_excel(path, engine="openpyxl")

    # Приводим числовые колонки — FineBI иногда выгружает их как строки с запятой
    for col in (config.COL_CAPACITY, config.COL_SPENT):
        if col not in df.columns:
            raise ValueError(
                f"Колонка '{col}' не найдена в файле. "
                f"Доступные колонки: {list(df.columns)}"
            )
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("\xa0", "", regex=False)
            .str.replace(" ", "", regex=False)
            .str.replace(",", ".", regex=False)
            .pipe(pd.to_numeric, errors="coerce")
            .fillna(0)
        )

    group_col = config.COL_GROUP
    dept_col  = config.COL_DEPT

    if group_col not in df.columns:
        raise ValueError(
            f"Колонка группировки '{group_col}' не найдена. "
            f"Доступные колонки: {list(df.columns)}"
        )

    has_dept = dept_col in df.columns

    rows = []

    if has_dept:
        # Двухуровневая сводка: Управление → Отделы
        # Строки где Отдел == Управление — это начальник управления:
        # его часы входят в итог по управлению, но отдельной строкой не показываются.
        for group_name, group_df in df.groupby(group_col, sort=False):
            dept_pivot = (
                group_df.groupby(dept_col, sort=False)[[config.COL_CAPACITY, config.COL_SPENT]]
                .sum()
                .reset_index()
            )

            # Суммируем все строки включая начальника управления
            group_cap   = dept_pivot[config.COL_CAPACITY].sum()
            group_spent = dept_pivot[config.COL_SPENT].sum()
            group_pct   = round(group_spent / group_cap * 100, 1) if group_cap > 0 else 0.0

            rows.append({
                "row_type":      "group",
                "Управление":    group_name,
                "Подразделение": group_name,
                "Capacity, ч":   group_cap,
                "Потрачено, ч":  group_spent,
                "% списания":    group_pct,
            })

            for _, dr in dept_pivot.iterrows():
                # Пропускаем строку начальника управления (Отдел == Управление)
                if str(dr[dept_col]).strip() == str(group_name).strip():
                    continue
                cap   = dr[config.COL_CAPACITY]
                spent = dr[config.COL_SPENT]
                pct   = round(spent / cap * 100, 1) if cap > 0 else 0.0
                rows.append({
                    "row_type":      "dept",
                    "Управление":    group_name,
                    "Подразделение": dr[dept_col],
                    "Capacity, ч":   cap,
                    "Потрачено, ч":  spent,
                    "% списания":    pct,
                })
    else:
        # Одноуровневая сводка (нет колонки Отдел)
        pivot = (
            df.groupby(group_col, sort=False)[[config.COL_CAPACITY, config.COL_SPENT]]
            .sum()
            .reset_index()
        )
        for _, r in pivot.iterrows():
            cap   = r[config.COL_CAPACITY]
            spent = r[config.COL_SPENT]
            pct   = round(spent / cap * 100, 1) if cap > 0 else 0.0
            rows.append({
                "row_type":    "group",
                "Управление":  r[group_col],
                "Подразделение": r[group_col],
                "Capacity, ч":  cap,
                "Потрачено, ч": spent,
                "% списания":   pct,
            })

    pivot = pd.DataFrame(rows)

    # Итоговая строка
    total_cap   = pivot.loc[pivot["row_type"] != "group", "Capacity, ч"].sum() if has_dept \
                  else pivot["Capacity, ч"].sum()
    total_spent = pivot.loc[pivot["row_type"] != "group", "Потрачено, ч"].sum() if has_dept \
                  else pivot["Потрачено, ч"].sum()
    total_pct   = round(total_spent / total_cap * 100, 1) if total_cap > 0 else 0.0

    total_row = pd.DataFrame([{
        "row_type":    "total",
        "Управление":  "ИТОГО",
        "Подразделение": "ИТОГО",
        "Capacity, ч":  total_cap,
        "Потрачено, ч": total_spent,
        "% списания":   total_pct,
    }])
    pivot = pd.concat([pivot, total_row], ignore_index=True)

    return pivot


def pivot_to_html(pivot: pd.DataFrame) -> str:
    """Возвращает HTML-таблицу для вставки в тело письма."""
    styles = (
        "border-collapse:collapse;font-family:Calibri,Arial,sans-serif;"
        "font-size:11pt;"
    )
    th_style = (
        "background:#1F497D;color:#fff;padding:6px 10px;"
        "border:1px solid #ccc;text-align:left;"
    )
    td_style     = "padding:5px 10px;border:1px solid #ccc;"
    td_num_style = td_style + "text-align:right;"
    total_style     = td_style + "font-weight:bold;background:#f2f2f2;"
    total_num_style = total_style + "text-align:right;"
    group_style     = td_style + "font-weight:bold;background:#e8edf4;"
    group_num_style = group_style + "text-align:right;"
    dept_style      = td_style + "padding-left:24px;"

    rows_html = []
    for _, row in pivot.iterrows():
        rtype = row.get("row_type", "dept")
        pct_val = row["% списания"]

        if rtype == "total":
            td = total_style
            td_n = total_num_style
            pct_cell = f'<td style="{td_n}">{pct_val:.1f}%</td>'
        elif rtype == "group":
            td = group_style
            td_n = group_num_style
            pct_cell = f'<td style="{td_n}">{pct_val:.1f}%</td>'
        else:
            td = dept_style
            td_n = td_num_style
            if pct_val < 80:
                color = "#c00000"
            elif pct_val < 95:
                color = "#9c6500"
            else:
                color = "#375623"
            pct_cell = (
                f'<td style="{td_n}color:{color};font-weight:bold;">'
                f"{pct_val:.1f}%</td>"
            )

        rows_html.append(
            f"<tr>"
            f'<td style="{td}">{row["Подразделение"]}</td>'
            f'<td style="{td_n}">{row["Capacity, ч"]:,.1f}</td>'
            f'<td style="{td_n}">{row["Потрачено, ч"]:,.1f}</td>'
            f"{pct_cell}"
            f"</tr>"
        )

    header = (
        f'<tr>'
        f'<th style="{th_style}">Подразделение</th>'
        f'<th style="{th_style}">Capacity, ч</th>'
        f'<th style="{th_style}">Потрачено, ч</th>'
        f'<th style="{th_style}">% списания</th>'
        f"</tr>"
    )

    return (
        f'<table style="{styles}">'
        f"<thead>{header}</thead>"
        f"<tbody>{''.join(rows_html)}</tbody>"
        f"</table>"
    )
