"""
Читает Excel из FineBI и строит сводную таблицу:
  Управление | Capacity, ч | Потрачено, ч | % списания
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
            .str.replace("\xa0", "", regex=False)   # неразрывный пробел
            .str.replace(" ", "", regex=False)
            .str.replace(",", ".", regex=False)
            .pipe(pd.to_numeric, errors="coerce")
            .fillna(0)
        )

    group_col = config.COL_GROUP
    if group_col not in df.columns:
        raise ValueError(
            f"Колонка группировки '{group_col}' не найдена. "
            f"Доступные колонки: {list(df.columns)}"
        )

    pivot = (
        df.groupby(group_col, sort=False)[[config.COL_CAPACITY, config.COL_SPENT]]
        .sum()
        .reset_index()
    )

    pivot["% списания"] = pivot.apply(
        lambda row: (row[config.COL_SPENT] / row[config.COL_CAPACITY] * 100)
        if row[config.COL_CAPACITY] > 0
        else 0,
        axis=1,
    ).round(1)

    pivot = pivot.rename(
        columns={
            config.COL_GROUP: "Подразделение",
            config.COL_CAPACITY: "Capacity, ч",
            config.COL_SPENT: "Потрачено, ч",
        }
    )

    # Итоговая строка
    total_cap = pivot["Capacity, ч"].sum()
    total_spent = pivot["Потрачено, ч"].sum()
    total_pct = round(total_spent / total_cap * 100, 1) if total_cap > 0 else 0

    total_row = pd.DataFrame(
        [["ИТОГО", total_cap, total_spent, total_pct]],
        columns=pivot.columns,
    )
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
    td_style = "padding:5px 10px;border:1px solid #ccc;"
    td_num_style = td_style + "text-align:right;"
    total_style = td_style + "font-weight:bold;background:#f2f2f2;"
    total_num_style = total_style + "text-align:right;"

    rows_html = []
    for _, row in pivot.iterrows():
        is_total = str(row["Подразделение"]).upper() == "ИТОГО"
        td = total_style if is_total else td_style
        td_n = total_num_style if is_total else td_num_style

        pct_val = row["% списания"]
        # Цветовая индикация: <80% — красный, 80-95% — жёлтый, >=95% — зелёный
        if not is_total:
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
        else:
            pct_cell = f'<td style="{td_n}">{pct_val:.1f}%</td>'

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
