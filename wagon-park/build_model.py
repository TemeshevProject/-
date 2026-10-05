#!/usr/bin/env python3
"""Финансовая модель парка вагонов.

Жёлтые ячейки на листе «Вводы» можно менять: ставка, дни загрузки, цена,
ремонт, налог, курс, размер парка. Выручка, EBITDA, окупаемость, IRR и NPV
пересчитываются формулами.

Абсолютные ставки крытого вагона в открытых источниках не опубликованы —
базовое значение 1 500 руб./сутки помечено как модельное допущение.
"""

from pathlib import Path

import xlsxwriter

OUT = Path(__file__).resolve().parent / "Wagon-Park-Kazakhstan.xlsx"


def main() -> None:
    wb = xlsxwriter.Workbook(OUT)
    n = wb.add_format({"num_format": "#,##0", "font_name": "Calibri", "font_size": 11})
    n1 = wb.add_format({"num_format": "#,##0.0", "font_name": "Calibri", "font_size": 11})
    pct = wb.add_format({"num_format": "0.0%", "font_name": "Calibri", "font_size": 11})
    inp = wb.add_format(
        {
            "num_format": "#,##0.00",
            "bg_color": "#FFF3CD",
            "border": 1,
            "font_name": "Calibri",
            "font_size": 11,
            "locked": False,
        }
    )
    inp_i = wb.add_format(
        {
            "num_format": "#,##0",
            "bg_color": "#FFF3CD",
            "border": 1,
            "font_name": "Calibri",
            "font_size": 11,
        }
    )
    label = wb.add_format({"font_name": "Calibri", "font_size": 11})
    bold = wb.add_format({"font_name": "Calibri", "font_size": 11, "bold": True})
    head = wb.add_format(
        {
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": "white",
            "bg_color": "#0F2B47",
            "text_wrap": True,
            "valign": "vcenter",
        }
    )
    section = wb.add_format(
        {"font_name": "Calibri", "font_size": 14, "bold": True, "font_color": "#0F2B47"}
    )
    note = wb.add_format(
        {"font_name": "Calibri", "font_size": 10, "font_color": "#44566C", "text_wrap": True, "valign": "top"}
    )
    result = wb.add_format(
        {"num_format": "#,##0", "font_name": "Calibri", "font_size": 11, "bg_color": "#E8F3ED"}
    )
    result_pct = wb.add_format(
        {"num_format": "0.0%", "font_name": "Calibri", "font_size": 11, "bg_color": "#E8F3ED"}
    )
    link = wb.add_format({"font_name": "Calibri", "font_size": 10, "font_color": "#16395C"})

    # ---------- Вводы ----------
    ws = wb.add_worksheet("Вводы")
    ws.set_column("A:A", 46)
    ws.set_column("B:B", 18)
    ws.set_column("C:C", 22)
    ws.set_column("D:D", 88)
    ws.write("A1", "Вводы модели — жёлтые ячейки можно менять", section)
    ws.write("A2", "Параметр", head)
    ws.write("B2", "Значение", head)
    ws.write("C2", "Единица", head)
    ws.write("D2", "Что это и откуда", head)

    rows = [
        (3, "fx", 5.35, "тенге за 1 рубль", "Официальный курс НБК на 5 октября 2026"),
        (4, "tax", 0.20, "доля", "Базовая ставка КПН с 2026 года, НК РК"),
        (5, "life", 32, "лет", "Назначенный срок службы современного грузового вагона колеи 1520. Сверять по паспорту модели"),
        (6, "residual", 0.15, "доля цены", "Допущение: остаточная стоимость в конце срока службы"),
        (7, "ins", 0.0035, "доля цены в год", "Допущение: страхование 0,35% стоимости в год"),
        (8, "nbk", 0.1625, "доля", "Базовая ставка НБК с 7 сентября 2026, коридор ±1 п.п."),
        (9, "bank_spread", 0.05, "доля", "Допущение: надбавка банка к базовой ставке. Это не оферта"),
        (10, "loan_share", 0.50, "доля цены", "Доля долга в сравнении «покупка / кредит»"),
        (11, "loan_years", 7, "лет", "Срок кредита в модели"),
        (12, "overhead_small", 30000000, "тенге в год", "Допущение: постоянные расходы оператора при парке до 50 вагонов"),
        (13, "overhead_100", 45000000, "тенге в год", "Допущение: постоянные расходы при парке около 100 вагонов"),
        (14, "overhead_300", 80000000, "тенге в год", "Допущение: постоянные расходы при парке около 300 вагонов"),
        (15, "overhead_500", 120000000, "тенге в год", "Допущение: постоянные расходы при парке около 500 вагонов"),
    ]
    for r, name, val, unit, comment in rows:
        ws.write(r - 1, 0, name, label)
        fmt = inp_i if isinstance(val, int) else inp
        if name in ("tax", "residual", "ins", "nbk", "bank_spread", "loan_share"):
            fmt = wb.add_format(
                {"num_format": "0.00%", "bg_color": "#FFF3CD", "border": 1, "font_name": "Calibri", "font_size": 11}
            )
        ws.write(r - 1, 1, val, fmt)
        ws.write(r - 1, 2, unit, label)
        ws.write(r - 1, 3, comment, note)
        wb.define_name(name, f"=Вводы!$B${r}")

    ws.write("A17", "Цены и содержание одного вагона", section)
    headers_w = [
        (18, "price_grain", 5080000, "Зерновоз, цена нового, руб.", "Обзор «Паровоз Лизинг», июль 2025: 5,08 млн руб. после снижения с 5,90"),
        (19, "price_covered", 5600000, "Крытый 23,5 тс, цена нового, руб.", "Тот же обзор, июль 2025: 5,60 млн. Отдельной цены большекубового с люками на крыше в открытом источнике нет — используется как ориентир"),
        (20, "price_gondola", 5730000, "Полувагон 25 тс, цена нового, руб.", "Тот же обзор, июль 2025: 5,73 млн после снижения с 6,60"),
        (21, "used_coef", 0.50, "Цена б/у к цене нового", "Допущение: открытого индекса вторичного рынка на октябрь 2026 нет"),
        (22, "maint_grain", 175000, "Содержание зерновоза, руб./год", "Сборка: ДР 130 тыс. без запчастей (2025) + допущение по запчастям, КР 400–500 тыс. раз в 12 лет"),
        (23, "maint_covered", 185000, "Содержание крытого, руб./год", "Та же логика, КР крытого 450–600 тыс."),
        (24, "maint_gondola", 320835, "Содержание полувагона, руб./год", "879 руб./сутки × 365. АО «ОЖдПС», динамика 2024→2025: с 708 до 879 руб./сутки"),
        (25, "maint_used_uplift", 0.20, "Надбавка к ремонту б/у", "Допущение: б/у дороже в содержании на 20%"),
        (26, "life_used", 16, "Остаточный срок б/у, лет", "Допущение для вагона примерно середины срока службы"),
    ]
    for r, name, val, title, comment in headers_w:
        ws.write(r - 1, 0, title, label)
        if name in ("used_coef", "maint_used_uplift"):
            fmt = wb.add_format(
                {"num_format": "0%", "bg_color": "#FFF3CD", "border": 1, "font_name": "Calibri", "font_size": 11}
            )
        elif isinstance(val, int):
            fmt = inp_i
        else:
            fmt = inp
        ws.write(r - 1, 1, val, fmt)
        ws.write(r - 1, 3, comment, note)
        wb.define_name(name, f"=Вводы!$B${r}")

    ws.write("A28", "Ставка и дни, когда ставка зарабатывается", section)
    ws.write("A29", "Ставка — руб./сутки. Дни — сколько суток в году вагон находится в коммерческом использовании и приносит ставку. Остальные сутки ставка равна нулю, а содержание платится за весь год.", note)
    rate_rows = [
        (30, "g_rate_c", 1200, "Зерновоз, консервативная ставка"),
        (31, "g_days_c", 180, "Зерновоз, консервативные дни"),
        (32, "g_rate_b", 1700, "Зерновоз, базовая ставка"),
        (33, "g_days_b", 230, "Зерновоз, базовые дни"),
        (34, "g_rate_o", 2500, "Зерновоз, оптимистичная ставка"),
        (35, "g_days_o", 280, "Зерновоз, оптимистичные дни"),
        (36, "c_rate_c", 1100, "Крытый, консервативная ставка"),
        (37, "c_days_c", 170, "Крытый, консервативные дни"),
        (38, "c_rate_b", 1500, "Крытый, базовая ставка — ДОПУЩЕНИЕ"),
        (39, "c_days_b", 220, "Крытый, базовые дни"),
        (40, "c_rate_o", 2000, "Крытый, оптимистичная ставка"),
        (41, "c_days_o", 270, "Крытый, оптимистичные дни"),
        (42, "p_rate_c", 650, "Полувагон, консервативная ставка"),
        (43, "p_days_c", 180, "Полувагон, консервативные дни"),
        (44, "p_rate_b", 900, "Полувагон, базовая ставка"),
        (45, "p_days_b", 240, "Полувагон, базовые дни"),
        (46, "p_rate_o", 1260, "Полувагон, оптимистичная ставка"),
        (47, "p_days_o", 300, "Полувагон, оптимистичные дни"),
    ]
    for r, name, val, title in rate_rows:
        ws.write(r - 1, 0, title, label)
        ws.write(r - 1, 1, val, inp_i)
        wb.define_name(name, f"=Вводы!$B${r}")
    ws.write(
        "D30",
        "База зерновоза 1 700 — ниже июльской аренды 1 910 («Паровоз Лизинг») и ниже августовского спота 2 300–3 500 (ИЦО), потому что долгосрочные контракты на межсезонье не заключаются. Консерватив и оптимум — границы модели, не котировка.",
        note,
    )
    ws.write(
        "D36",
        "Абсолютная ставка крытого вагона в открытых обзорах ИЦО и РЖД-Партнер не раскрыта, опубликовано только направление (снижение на фоне профицита и конкуренции с автотранспортом). 1 500 руб./сутки — модельное допущение между полувагоном и зерновозом. Меняйте ячейку под коммерческое предложение.",
        note,
    )
    ws.write(
        "D42",
        "База 900 руб./сутки — долгосрочная аренда полувагона 23,5 тс в августе 2025, ИЦО (Vgudok). Консерватив 650 — нижняя граница ставки предоставления 650–1 000. Оптимум 1 260 — аренда в июле 2025 до августовского провала.",
        note,
    )

    ws.write("A49", "Размер парка для расчёта портфеля", section)
    qty = [
        (50, "q1_grain", 10, "Портфель 1: зерновозы"),
        (51, "q1_covered", 10, "Портфель 1: крытые"),
        (52, "q1_gondola", 10, "Портфель 1: полувагоны"),
        (53, "q2_grain", 20, "Портфель 2: зерновозы"),
        (54, "q2_covered", 20, "Портфель 2: крытые"),
        (55, "q2_gondola", 10, "Портфель 2: полувагоны"),
    ]
    for r, name, val, title in qty:
        ws.write(r - 1, 0, title, label)
        ws.write(r - 1, 1, val, inp_i)
        wb.define_name(name, f"=Вводы!$B${r}")
    ws.write("D50", "Портфель 1 — нижняя граница запроса: 10 крытых, 10 зерновозов, 10 полувагонов. Портфель 2 — верхняя: 20 / 20 / 10.", note)

    ws.write("A57", "Как считается один вагон", bold)
    ws.write(
        "A58",
        "Выручка = ставка × дни коммерческого использования. Содержание и страхование платятся за год целиком. "
        "EBITDA = выручка − содержание − страхование. Амортизация линейная до остаточной стоимости. "
        "Налог = 20% × max(EBITDA − амортизация; 0). Денежный поток после налога = EBITDA − налог. "
        "Простая окупаемость = цена / EBITDA. IRR считается на всём оставшемся сроке службы, в последнем году возвращается остаточная стоимость. "
        "ROI — годовой денежный поток после налога, делённый на цену. NPV дисконтирует те же потоки по базовой ставке НБК. "
        "Инфраструктурный тариф КТЖ и порожний пробег в этой модели платит грузоотправитель: собственник получает только ставку предоставления. "
        "Если порожний тариф переложен на собственника, результат будет хуже.",
        note,
    )
    ws.set_row(57, 78)

    # ---------- Юнит ----------
    u = wb.add_worksheet("Юнит")
    u.set_column("A:A", 38)
    u.set_column("B:G", 16)
    u.write("A1", "Экономика одного нового вагона", section)
    for col, text in enumerate(["Показатель", "Зерновоз конс.", "Зерновоз база", "Зерновоз опт.", "Крытый конс.", "Крытый база", "Крытый опт."]):
        u.write(2, col, text, head)
    # rates and days row references via names
    # Row 4 price, 5 rate, 6 days, 7 revenue, 8 maint, 9 ins, 10 opex, 11 ebitda, 12 dep, 13 ebit, 14 tax, 15 net, 16 atcf, 17 payback, 18 life irr placeholder calculated in python? 
    # IRR over 32 years is painful in pure Excel without a helper. I'll use a 32-step cashflow block on a hidden area and IRR().
    # Simpler: compute IRR with Excel IRR on a column of cash flows per scenario. That's 6 scenarios * 33 rows. Doable to the right.

    specs = [
        ("B", "price_grain", "g_rate_c", "g_days_c", "maint_grain", "life"),
        ("C", "price_grain", "g_rate_b", "g_days_b", "maint_grain", "life"),
        ("D", "price_grain", "g_rate_o", "g_days_o", "maint_grain", "life"),
        ("E", "price_covered", "c_rate_c", "c_days_c", "maint_covered", "life"),
        ("F", "price_covered", "c_rate_b", "c_days_b", "maint_covered", "life"),
        ("G", "price_covered", "c_rate_o", "c_days_o", "maint_covered", "life"),
    ]
    labels_u = [
        (3, "Цена, руб."),
        (4, "Ставка, руб./сутки"),
        (5, "Дни со ставкой"),
        (6, "Выручка, руб./год"),
        (7, "Содержание, руб./год"),
        (8, "Страхование, руб./год"),
        (9, "OPEX, руб./год"),
        (10, "EBITDA, руб./год"),
        (11, "Амортизация, руб./год"),
        (12, "Прибыль до налога"),
        (13, "КПН"),
        (14, "Чистая прибыль"),
        (15, "Денежный поток после налога"),
        (16, "Окупаемость по EBITDA, лет"),
        (17, "Окупаемость по ден. потоку, лет"),
        (18, "EBITDA-маржа"),
        (19, "IRR на срок службы"),
        (20, "ROI, денежный поток / цена"),
        (21, "NPV при ставке НБК, руб."),
    ]
    for r, text in labels_u:
        u.write(r, 0, text, bold if r in (10, 16, 19) else label)

    for col, price, rate, days, maint, life in specs:
        u.write_formula(3, ord(col) - 65, f"={price}", n)
        u.write_formula(4, ord(col) - 65, f"={rate}", n)
        u.write_formula(5, ord(col) - 65, f"={days}", n)
        u.write_formula(6, ord(col) - 65, f"={col}5*{col}6", result)
        u.write_formula(7, ord(col) - 65, f"={maint}", n)
        u.write_formula(8, ord(col) - 65, f"={col}4*ins", n)
        u.write_formula(9, ord(col) - 65, f"={col}8+{col}9", n)
        u.write_formula(10, ord(col) - 65, f"={col}7-{col}10", result)
        u.write_formula(11, ord(col) - 65, f"={col}4*(1-residual)/{life}", n)
        u.write_formula(12, ord(col) - 65, f"={col}11-{col}12", n)
        u.write_formula(13, ord(col) - 65, f"=MAX({col}13,0)*tax", n)
        u.write_formula(14, ord(col) - 65, f"={col}13-{col}14", n)
        u.write_formula(15, ord(col) - 65, f"={col}11-{col}14", result)
        u.write_formula(16, ord(col) - 65, f'=IF({col}11<=0,"не окупается",{col}4/{col}11)', n1)
        u.write_formula(17, ord(col) - 65, f'=IF({col}16<=0,"не окупается",{col}4/{col}16)', n1)
        u.write_formula(18, ord(col) - 65, f"=IF({col}7=0,NA(),{col}11/{col}7)", pct)

    # IRR block starting at column I, row 4: for each scenario, year 0..life
    # life is 32, rows 4 to 36
    u.write("I2", "Денежные потоки для IRR (год 0 — покупка, последний год + остаток)", note)
    for idx, col in enumerate("BCDEFG"):
        irr_col = 8 + idx  # I=8
        u.write(2, irr_col, f"={col}3", head)
        u.write_formula(3, irr_col, f"=-{col}4", n)
        # years 1..31 at rows 5..35 (excel rows 5-35), year 32 at row 36
        # spreadsheet row index 4 is year 0 (excel row 5). Years 1-31: rows 5..35, year 32 row 36.
        for year in range(1, 32):
            u.write_formula(3 + year, irr_col, f"={col}16", n)
        u.write_formula(3 + 32, irr_col, f"={col}16+{col}4*residual", n)
        u.write_formula(19, ord(col) - 65, f"=IRR({chr(65+irr_col)}4:{chr(65+irr_col)}36)", result_pct)
        u.write_formula(20, ord(col) - 65, f"=IF({col}4=0,NA(),{col}16/{col}4)", result_pct)
        u.write_formula(21, ord(col) - 65, f"=NPV(nbk,{chr(65+irr_col)}5:{chr(65+irr_col)}36)+{chr(65+irr_col)}4", result)

    # Gondola block
    u.write("A22", "Полувагон", section)
    for col, text in enumerate(["Показатель", "Консервативный", "Базовый", "Оптимистичный"]):
        u.write(22, col, text, head)
    gspecs = [
        ("B", "p_rate_c", "p_days_c"),
        ("C", "p_rate_b", "p_days_b"),
        ("D", "p_rate_o", "p_days_o"),
    ]
    glabels = [
        (23, "Цена, руб."),
        (24, "Ставка, руб./сутки"),
        (25, "Дни со ставкой"),
        (26, "Выручка"),
        (27, "Содержание"),
        (28, "Страхование"),
        (29, "OPEX"),
        (30, "EBITDA"),
        (31, "Амортизация"),
        (32, "Прибыль до налога"),
        (33, "КПН"),
        (34, "Чистая прибыль"),
        (35, "Денежный поток после налога"),
        (36, "Окупаемость по EBITDA, лет"),
        (37, "IRR на срок службы"),
        (38, "ROI, денежный поток / цена"),
        (39, "NPV при ставке НБК, руб."),
    ]
    for r, text in glabels:
        u.write(r, 0, text, bold if r in (30, 36, 37) else label)
    for col, rate, days in gspecs:
        c = ord(col) - 65
        u.write_formula(23, c, "=price_gondola", n)
        u.write_formula(24, c, f"={rate}", n)
        u.write_formula(25, c, f"={days}", n)
        u.write_formula(26, c, f"={col}25*{col}26", result)
        u.write_formula(27, c, "=maint_gondola", n)
        u.write_formula(28, c, f"={col}24*ins", n)
        u.write_formula(29, c, f"={col}28+{col}29", n)
        u.write_formula(30, c, f"={col}27-{col}30", result)
        u.write_formula(31, c, f"={col}24*(1-residual)/life", n)
        u.write_formula(32, c, f"={col}31-{col}32", n)
        u.write_formula(33, c, f"=MAX({col}33,0)*tax", n)
        u.write_formula(34, c, f"={col}33-{col}34", n)
        u.write_formula(35, c, f"={col}31-{col}34", result)
        u.write_formula(36, c, f'=IF({col}31<=0,"не окупается",{col}24/{col}31)', n1)
    for idx, col in enumerate("BCD"):
        irr_col = 8 + idx
        u.write_formula(40, irr_col, f"=-{col}24", n)
        for year in range(1, 32):
            u.write_formula(40 + year, irr_col, f"={col}36", n)
        u.write_formula(40 + 32, irr_col, f"={col}36+{col}24*residual", n)
        u.write_formula(37, ord(col) - 65, f"=IRR({chr(65+irr_col)}41:{chr(65+irr_col)}73)", result_pct)
        u.write_formula(38, ord(col) - 65, f"=IF({col}24=0,NA(),{col}36/{col}24)", result_pct)
        u.write_formula(39, ord(col) - 65, f"=NPV(nbk,{chr(65+irr_col)}42:{chr(65+irr_col)}73)+{chr(65+irr_col)}41", result)

    u.write("A41", "Потоки полувагона для IRR и NPV начинаются в столбцах I–K, строки 41–73. Годы после остаточного срока б/у на листе «Барьер» остаются пустыми: Excel их в IRR не включает.", note)

    # ---------- Портфель ----------
    p = wb.add_worksheet("Портфель")
    p.set_column("A:A", 52)
    p.set_column("B:C", 22)
    p.write("A1", "Два парка из запроса, базовый сценарий", section)
    p.write("A2", "EBITDA одного вагона берётся из листа «Юнит»: зерновоз C11, крытый F11, полувагон C31.", note)
    p.write_row(3, 0, ["Показатель", "10 + 10 + 10", "20 + 20 + 10"], head)
    items = [
        (4, "Зерновозы, шт.", "=q1_grain", "=q2_grain"),
        (5, "Крытые, шт.", "=q1_covered", "=q2_covered"),
        (6, "Полувагоны, шт.", "=q1_gondola", "=q2_gondola"),
        (7, "CAPEX, руб.", "=B5*price_grain+B6*price_covered+B7*price_gondola", "=C5*price_grain+C6*price_covered+C7*price_gondola"),
        (8, "CAPEX, млн тенге", "=B8*fx/1000000", "=C8*fx/1000000"),
        (9, "Выручка, руб./год", "=B5*'Юнит'!C7+B6*'Юнит'!F7+B7*'Юнит'!C27", "=C5*'Юнит'!C7+C6*'Юнит'!F7+C7*'Юнит'!C27"),
        (10, "EBITDA вагонов, руб./год", "=B5*'Юнит'!C11+B6*'Юнит'!F11+B7*'Юнит'!C31", "=C5*'Юнит'!C11+C6*'Юнит'!F11+C7*'Юнит'!C31"),
        (11, "Постоянные расходы, руб./год", "=overhead_small/fx", "=overhead_small/fx"),
        (12, "EBITDA после офиса", "=B11-B12", "=C11-C12"),
        (13, "Окупаемость вагонов без офиса, лет", '=IF(B11<=0,"не окупается",B8/B11)', '=IF(C11<=0,"не окупается",C8/C11)'),
        (14, "Окупаемость с офисом, лет", '=IF(B13<=0,"не окупается",B8/B13)', '=IF(C13<=0,"не окупается",C8/C13)'),
    ]
    for r, title, b, c in items:
        p.write(r, 0, title, bold if r in (7, 10, 12, 14) else label)
        fmt = result if r in (7, 10, 12) else n
        p.write_formula(r, 1, b, fmt)
        p.write_formula(r, 2, c, fmt)
    p.write("A16", "Офис заложен на уровне парка до 50 вагонов (ячейка overhead_small). Если диспетчерская уже есть у экспедитора, поставьте её равной нулю и увидите экономику без содержания компании.", note)
    p.set_row(16, 32)

    p.write("A18", "Тот же парк в оптимистичном сценарии", section)
    p.write_row(19, 0, ["Показатель", "10 + 10 + 10", "20 + 20 + 10"], head)
    opt = [
        (20, "EBITDA вагонов, руб./год", "=B5*'Юнит'!D11+B6*'Юнит'!G11+B7*'Юнит'!D31", "=C5*'Юнит'!D11+C6*'Юнит'!G11+C7*'Юнит'!D31"),
        (21, "EBITDA после офиса", "=B21-B12", "=C21-C12"),
        (22, "Окупаемость с офисом, лет", '=IF(B22<=0,"не окупается",B8/B22)', '=IF(C22<=0,"не окупается",C8/C22)'),
    ]
    for r, title, b, c in opt:
        p.write(r, 0, title, label)
        p.write_formula(r, 1, b, result if r != 22 else n1)
        p.write_formula(r, 2, c, result if r != 22 else n1)

    # ---------- Барьер ----------
    b = wb.add_worksheet("Барьер")
    b.set_column("A:A", 62)
    b.set_column("B:D", 18)
    b.write("A1", "Какая ставка нужна, чтобы вагон окупился", section)
    b.write("A2", "Считается при базовом числе дней. Ставка безубыточности покрывает содержание и страхование, но не возвращает цену вагона.", note)
    b.write_row(3, 0, ["Порог", "Зерновоз", "Крытый", "Полувагон"], head)
    # days base: 'Юнит'!C6 grain, F6 covered, C26 gondola
    # opex: Юнит C10, F10, C30
    # price C4, F4, C24
    lines = [
        (4, "Дни базового сценария", "='Юнит'!C6", "='Юнит'!F6", "='Юнит'!C26"),
        (5, "OPEX, руб./год", "='Юнит'!C10", "='Юнит'!F10", "='Юнит'!C30"),
        (6, "Ставка денежной безубыточности, руб./сутки", "=B6/B5", "=C6/C5", "=D6/D5"),
        (7, "Ставка для окупаемости 8 лет", "=('Юнит'!C4/8+B6)/B5", "=('Юнит'!F4/8+C6)/C5", "=('Юнит'!C24/8+D6)/D5"),
        (8, "Ставка для окупаемости 10 лет", "=('Юнит'!C4/10+B6)/B5", "=('Юнит'!F4/10+C6)/C5", "=('Юнит'!C24/10+D6)/D5"),
        (9, "Ставка для окупаемости 12 лет", "=('Юнит'!C4/12+B6)/B5", "=('Юнит'!F4/12+C6)/C5", "=('Юнит'!C24/12+D6)/D5"),
        (10, "Ставка для окупаемости 15 лет", "=('Юнит'!C4/15+B6)/B5", "=('Юнит'!F4/15+C6)/C5", "=('Юнит'!C24/15+D6)/D5"),
        (11, "Рыночный ориентир, руб./сутки", 1910, "не опубликован", 900),
    ]
    for r, title, *vals in lines:
        b.write(r, 0, title, bold if r in (6, 9) else label)
        for i, v in enumerate(vals):
            if isinstance(v, str) and v.startswith("="):
                b.write_formula(r, 1 + i, v, n)
            else:
                b.write(r, 1 + i, v, n if isinstance(v, (int, float)) else label)
    b.write("A13", "Рыночный ориентир зерновоза — аренда июля 2025, 1 910 руб./сутки («Паровоз Лизинг»), не августовский спот. Полувагон — долгосрочная аренда августа 2025, 900 руб./сутки (ИЦО). Крытый ориентир в открытом источнике отсутствует.", note)
    b.set_row(13, 32)
    b.write("A15", "Б/у зерновоз: цена = price_grain × used_coef, содержание × (1+maint_used_uplift), срок = life_used. Поменяйте коэффициент на листе «Вводы».", note)

    b.write_row(16, 0, ["Б/у зерновоз", "База", "Оптимистичный"], head)
    b.write(17, 0, "Цена, руб.", label)
    b.write_formula(17, 1, "=price_grain*used_coef", n)
    b.write_formula(17, 2, "=B18", n)
    b.write(18, 0, "Содержание, руб./год", label)
    b.write_formula(18, 1, "=maint_grain*(1+maint_used_uplift)", n)
    b.write_formula(18, 2, "=B19", n)
    b.write(19, 0, "OPEX", label)
    b.write_formula(19, 1, "=B19+B18*ins", n)
    b.write_formula(19, 2, "=C19+C18*ins", n)
    b.write(20, 0, "Выручка", label)
    b.write_formula(20, 1, "=g_rate_b*g_days_b", n)
    b.write_formula(20, 2, "=g_rate_o*g_days_o", n)
    b.write(21, 0, "EBITDA", label)
    b.write_formula(21, 1, "=B21-B20", result)
    b.write_formula(21, 2, "=C21-C20", result)
    b.write(22, 0, "Окупаемость по EBITDA, лет", label)
    b.write_formula(22, 1, '=IF(B22<=0,"не окупается",B18/B22)', n1)
    b.write_formula(22, 2, '=IF(C22<=0,"не окупается",C18/C22)', n1)
    b.write(23, 0, "Амортизация, руб./год", label)
    b.write_formula(23, 1, "=B18*(1-residual)/life_used", n)
    b.write_formula(23, 2, "=C18*(1-residual)/life_used", n)
    b.write(24, 0, "Прибыль до налога", label)
    b.write_formula(24, 1, "=B22-B24", n)
    b.write_formula(24, 2, "=C22-C24", n)
    b.write(25, 0, "КПН", label)
    b.write_formula(25, 1, "=MAX(B25,0)*tax", n)
    b.write_formula(25, 2, "=MAX(C25,0)*tax", n)
    b.write(26, 0, "Денежный поток после налога", label)
    b.write_formula(26, 1, "=B22-B26", result)
    b.write_formula(26, 2, "=C22-C26", result)
    b.write(27, 0, "ROI, денежный поток / цена", label)
    b.write_formula(27, 1, "=IF(B18=0,NA(),B27/B18)", result_pct)
    b.write_formula(27, 2, "=IF(C18=0,NA(),C27/C18)", result_pct)
    b.write(28, 0, "IRR на остаточный срок", label)
    b.write_formula(28, 1, "=IRR(H40:H72)", result_pct)
    b.write_formula(28, 2, "=IRR(I40:I72)", result_pct)
    b.write(29, 0, "NPV при ставке НБК, руб.", label)
    b.write_formula(29, 1, "=NPV(nbk,H41:H72)+H40", result)
    b.write_formula(29, 2, "=NPV(nbk,I41:I72)+I40", result)
    b.write(31, 0, "Потоки б/у для IRR: год 0 в строке 40, дальше пустая ячейка после life_used.", note)
    b.write(38, 7, "Б/у база", head)
    b.write(38, 8, "Б/у оптимум", head)
    b.write_formula(39, 7, "=-B18", n)
    b.write_formula(39, 8, "=-C18", n)
    for year in range(1, 33):
        b.write_formula(
            39 + year,
            7,
            f'=IF({year}>life_used,"",IF({year}=life_used,B27+B18*residual,B27))',
            n,
        )
        b.write_formula(
            39 + year,
            8,
            f'=IF({year}>life_used,"",IF({year}=life_used,C27+C18*residual,C27))',
            n,
        )

    # ---------- Масштаб ----------
    s = wb.add_worksheet("Масштаб")
    s.set_column("A:A", 42)
    s.set_column("B:F", 16)
    s.write("A1", "Парк 20 / 50 / 100 / 300 / 500 в базовом сценарии", section)
    s.write("A2", "Смешанный парк повторяет пропорцию 40% зерновозы / 40% крытые / 20% полувагоны. Рядом — парк только из зерновозов, это лучший из трёх типов.", note)
    s.write_row(3, 0, ["", 20, 50, 100, 300, 500], head)
    s.write(4, 0, "Постоянные расходы, тенге", label)
    for col, name in enumerate(["overhead_small", "overhead_small", "overhead_100", "overhead_300", "overhead_500"], start=1):
        s.write_formula(4, col, f"={name}", n)
    s.write(5, 0, "Постоянные расходы, руб.", label)
    for col in range(1, 6):
        s.write_formula(5, col, f"={chr(64+col+1)}5/fx", n)
    # per wagon ebitda and price
    s.write(7, 0, "Смешанный парк: CAPEX, руб.", label)
    s.write(8, 0, "Смешанный парк: EBITDA вагонов", label)
    s.write(9, 0, "Смешанный парк: EBITDA после офиса", label)
    s.write(10, 0, "Окупаемость с офисом, лет", label)
    s.write(12, 0, "Только зерновозы: CAPEX, руб.", label)
    s.write(13, 0, "Только зерновозы: EBITDA вагонов", label)
    s.write(14, 0, "Только зерновозы: EBITDA после офиса", label)
    s.write(15, 0, "Окупаемость с офисом, лет", label)
    mix_price = "(0.4*price_grain+0.4*price_covered+0.2*price_gondola)"
    mix_ebit = "(0.4*'Юнит'!C11+0.4*'Юнит'!F11+0.2*'Юнит'!C31)"
    for col, qty_n in enumerate([20, 50, 100, 300, 500], start=1):
        letter = chr(64 + col + 1)
        s.write_formula(7, col, f"={qty_n}*{mix_price}", n)
        s.write_formula(8, col, f"={qty_n}*{mix_ebit}", n)
        s.write_formula(9, col, f"={letter}9-{letter}6", result)
        s.write_formula(10, col, f'=IF({letter}10<=0,"не окупается",{letter}8/{letter}10)', n1)
        s.write_formula(12, col, f"={qty_n}*price_grain", n)
        s.write_formula(13, col, f"={qty_n}*'Юнит'!C11", n)
        s.write_formula(14, col, f"={letter}14-{letter}6", result)
        s.write_formula(15, col, f'=IF({letter}15<=0,"не окупается",{letter}13/{letter}15)', n1)
    s.write("A17", "Масштаб не лечит слабую единицу. Если один вагон не покрывает стоимость денег, пятьсот вагонов увеличивают и выручку, и замороженный капитал.", note)

    # ---------- Финансирование ----------
    f = wb.add_worksheet("Финансирование")
    f.set_column("A:A", 55)
    f.set_column("B:E", 20)
    f.write("A1", "Покупка, кредит, лизинг — на примере одного нового зерновоза, базовый сценарий", section)
    f.write_row(2, 0, ["", "Свои деньги", "Кредит 50%", "Фин. лизинг", "Операционная аренда"], head)
    f.write(3, 0, "Цена / объём финансирования, руб.", label)
    f.write_formula(3, 1, "=price_grain", n)
    f.write_formula(3, 2, "=price_grain*loan_share", n)
    f.write_formula(3, 3, "=price_grain*0.8", n)
    f.write(3, 4, "0, вагон не покупается", label)
    f.write(4, 0, "Собственный капитал на старте, руб.", label)
    f.write_formula(4, 1, "=price_grain", n)
    f.write_formula(4, 2, "=price_grain*(1-loan_share)", n)
    f.write_formula(4, 3, "=price_grain*0.2", n)
    f.write(4, 4, "Депозит 2 месяца ставки", label)
    f.write(5, 0, "Ставка долга", label)
    f.write(5, 1, "—", label)
    f.write_formula(5, 2, "=nbk+bank_spread", pct)
    f.write(5, 3, 0.22, pct)
    f.write(5, 4, "рыночная аренда", label)
    f.write(6, 0, "Платёж в год, руб.", label)
    f.write(6, 1, 0, n)
    f.write_formula(6, 2, "=PMT(B6,loan_years,-C4)", n)  # wait columns: rate is C6? 
    # I put rate in column C row 6 (index 5) which is excel C6. Payment should use C6 and C4.
    # Let me fix: column B is equity, C loan, D lease, E operating. 
    # Row 6 excel is index 5. Loan rate is C6. Loan principal is C4 (excel row 4 = index 3). 
    # I wrote PMT(B6,...) which is wrong. Need to overwrite.
    f.write_formula(6, 2, "=PMT(C6,loan_years,-C4)", n)
    f.write_formula(6, 3, "=PMT(D6,loan_years,-D4)", n)
    f.write_formula(6, 4, "=g_rate_b*365", n)
    f.write(7, 0, "EBITDA базового зерновоза", label)
    f.write_formula(7, 1, "='Юнит'!C11", result)
    f.write_formula(7, 2, "='Юнит'!C11", result)
    f.write_formula(7, 3, "='Юнит'!C11", result)
    f.write(7, 4, "Спреда нет, если сдаёте по той же ставке", label)
    f.write(8, 0, "EBITDA − долговой платёж", label)
    f.write_formula(8, 1, "=B8-B7", result)
    f.write_formula(8, 2, "=C8-C7", result)
    f.write_formula(8, 3, "=D8-D7", result)
    f.write(9, 0, "DSCR (EBITDA / платёж)", label)
    f.write(9, 1, "—", label)
    f.write_formula(9, 2, '=IF(C7=0,"—",C8/C7)', n1)
    f.write_formula(9, 3, '=IF(D7=0,"—",D8/D7)', n1)
    f.write(10, 0, "Денежная доходность на собственный капитал", label)
    f.write_formula(10, 1, "=IF(B5=0,NA(),B9/B5)", result_pct)
    f.write_formula(10, 2, "=IF(C5=0,NA(),C9/C5)", result_pct)
    f.write_formula(10, 3, "=IF(D5=0,NA(),D9/D5)", result_pct)
    f.write(11, 0, "Платёж в месяц, руб.", label)
    f.write(11, 1, 0, n)
    f.write_formula(11, 2, "=PMT(C6/12,loan_years*12,-C4)", n)
    f.write_formula(11, 3, "=PMT(D6/12,loan_years*12,-D4)", n)
    f.write("A13", "Операционная аренда в столбце E — это отказ от покупки: вы платите рыночную ставку и не несёте остаточный риск. При текущем рынке спред пересдачи близок к нулю. Финансовый лизинг в модели: аванс 20%, ставка 22% годовых, срок 7 лет — допущение, не коммерческое предложение лизинговой компании. Кредит: базовый 16,25% + надбавка 5 п.п. Месячный платёж — аннуитет. Доходность на капитал здесь денежная: (EBITDA − долговой платёж) / собственные деньги на старте. Это не бухгалтерский ROE.", note)
    f.set_row(12, 64)

    # ---------- Источники ----------
    src = wb.add_worksheet("Источники")
    src.set_column("A:A", 28)
    src.set_column("B:B", 110)
    src.write("A1", "Источники, на которые опираются жёлтые вводы", section)
    src.write_row(2, 0, ["Блок", "Источник"], head)
    sources = [
        ("Курс", "Национальный Банк Казахстана, официальный курс на 5 октября 2026: 1 RUB = 5,35 KZT; 1 USD = 447,73 KZT."),
        ("Базовая ставка", "НБК, решение от 4 сентября 2026: 16,25% с 7 сентября 2026, коридор ±1 п.п. Следующее плановое решение — 23 октября 2026."),
        ("Налоги", "НК РК с 1 января 2026: базовая ставка КПН 20% (КГД), НДС 16%. Налог на имущество считается от зданий и сооружений; на грузовой вагон в модели не начисляется."),
        ("Перевозки, всего", "БНС РК, «Транспорт в Республике Казахстан», январь–декабрь 2025: 1 163,6 млн т, +7,9%; грузооборот 584,3 млрд т-км, +13,6%."),
        ("Железная дорога, БНС", "Там же: 466,7 млн т, +6,8%; грузооборот 361,3 млрд т-км, +10,2%."),
        ("Автомобильный", "Там же и пересказ TELS/ati.su: 352,7 млн т, +2,8%; 1,4 трлн тенге; грузооборот +32,8%."),
        ("Сеть КТЖ", "КТЖ для Kazinform: 2025 год — 319,97 млн т, в том числе внутри республики 176,91, экспорт 89,23, импорт 20,78, транзит 33,04. Акорда: тарифный грузооборот 289 млрд т-км, +10%; объём +5,5%."),
        ("Зерно", "Пресс-служба КТЖ, январь 2026: 14,3 млн т (+31%), экспорт 11 млн т (+33%). Узбекистан свыше 2 млн т (+41%), Центральная Азия 2,8 млн т (+27%)."),
        ("Уголь", "КТЖ, 11 месяцев 2025: 95,7 млн т (+3,3%), экспорт угля 31,2 млн т (+8%)."),
        ("Руда, 2026", "КТЖ, январь–июль 2026 (пересказ RZD-Partner): железная руда 13,8 млн т (+16,6%), цветная руда 12,5 млн т, нефтепродукты 15,2 млн т, стройматериалы 11,3 млн т."),
        ("ТМТМ", "Министерство транспорта, март 2026: ряд 2019–2025 до 4,119 млн т в 2025 году; около 77 тыс. ДФЭ; ориентир 300 тыс. ДФЭ к 2029 году."),
        ("Парк зерновозов", "Вице-министр транспорта Ж. Тайжанов, Kazinform: 16,3 тыс. зерновозов против 15,1 тыс. годом ранее; для муки задействовано 16,5 тыс. крытых; парк полностью закрывает потребность в перевозке зерна."),
        ("Қазтеміртранс", "Rail-news, интервью и отчёт на 1 октября 2025: 39 711 вагонов, нерабочий парк 2,8%. Закупки полувагонов: 1 000 в 2021–2022 и 3 000 в 2023–2025 у Казахстанской вагоностроительной компании. Оценка менеджмента: на сети 1520 профицит, активного спроса на новые вагоны нет."),
        ("Астық Транс", "Eldala, данные компании на 2025: 3,4 тыс. собственных зерновозов, 2,7 тыс. в аренде, до 1,5 тыс. привлекаются в сезон."),
        ("PTC Holding", "ptc-holding.kz, показатели 2025: парк 14 334 вагона, 15,4 млн т нефтепродуктов, 15% контейнерного грузооборота через Казахстан, 182 437 ДФЭ на терминале Достык."),
        ("KTZ Express", "Годовая отчётность за 2024: выручка 3 249 млн тенге, чистая прибыль 238 млн тенге, экспедирование железнодорожных перевозок 4 193 тыс. т."),
        ("Ставки полувагона", "ИЦО, август 2025, публикация Vgudok от 15 октября 2025: долгосрочная аренда 23,5 тс — 900 руб./сутки, спот — 600, ставка предоставления — 650–1 000. В 2023–2024 ставка была 2 500–3 600 руб./сутки (N.Trans Lab)."),
        ("Ставки зерновоза", "ИЦО, август 2025: краткосрочная аренда 2 300–3 500 руб./сутки, −20% к августу 2024. Новые долгосрочные контракты не заключаются из-за межсезонья с февраля. «Паровоз Лизинг»: средняя аренда снизилась с 4 120 в январе до 1 910 руб./сутки в июле 2025."),
        ("Крытые и платформы", "ИЦО: ставки крытых снижались из-за профицита и перетока на автотранспорт; абсолют в открытом доступе не опубликован. Фиттинговые 80 фут., август 2025: долгосрочные 1 400–2 200 руб./сутки, спот 960–1 300. 40 фут.: 1 100–1 430 и 500–700."),
        ("Цены новых вагонов", "«Паровоз Лизинг», июль 2025: зерновоз 5,08 млн руб., крытый 5,60 млн, полувагон 25 тс 5,73 млн, платформа 80 фут. 5,15 млн. Перекрёстная проверка: средняя цена производителя грузового вагона в РФ в марте 2026 — 5,31 млн руб. (обзор ACG)."),
        ("Ремонт", "РЖД-Партнер / отраслевые оценки 2025: средний деповской ремонт около 130 тыс. руб. без запчастей. Капитальный ремонт полувагона 1,2–1,4 млн руб., крытого 450–600 тыс., зерновоза 400–500 тыс. Суточное содержание полувагона — 879 руб. (АО «ОЖдПС»)."),
        ("Чего в модели нет", "Интервью с грузовладельцами, коммерческие предложения казахстанских операторов, паспорт конкретного вагона, договор примыкания и фактический порожний тариф маршрута. Пока их нет, казахстанская ставка может отличаться от российского бенчмарка 1520."),
    ]
    for i, (a, btext) in enumerate(sources):
        src.write(3 + i, 0, a, bold)
        src.write(3 + i, 1, btext, note)
        src.set_row(3 + i, 36)

    readme = wb.add_worksheet("Как пользоваться")
    readme.set_column("A:A", 120)
    readme.write("A1", "Модель парка вагонов — Казахстан, октябрь 2026", section)
    readme.write("A3", "1. Меняйте только жёлтые ячейки на листе «Вводы».", label)
    readme.write("A4", "2. Листы «Юнит», «Портфель», «Барьер», «Масштаб» и «Финансирование» пересчитают выручку, расходы, EBITDA, чистую прибыль, денежный поток, ROI, IRR, NPV и окупаемость.", label)
    readme.write("A5", "3. Ставка крытого вагона — допущение. Подставьте ставку из коммерческого предложения, как только она появится.", label)
    readme.write("A6", "4. Модель считает доход владельца вагона, а не маржу трейдера на зерне, угле или руде.", label)
    readme.write("A7", "5. Тариф инфраструктуры и порожний рейс в базе платит клиент. Иная договорная схема ухудшает результат.", label)
    readme.write("A8", "6. Это расчёт для решения, а не оценка имущества и не оферта.", label)
    readme.write("A10", "Презентация с выводами лежит рядом: Wagon-Park-Kazakhstan.pdf", bold)

    wb.close()
    print(OUT)


if __name__ == "__main__":
    main()
