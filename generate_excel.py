#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta

# ---------- Данные из фото + Регламент спортдиректора ----------
SEASONS = [
    {
        "season": "2012-2013",
        "format": "7×7 • 2 круга",
        "leagues": [
            {
                "name": "2012-2013 • 7×7 — 2 КРУГА",
                "format": "7×7",
                "double": True,
                "stage": "2 круга — каждый с каждым дома/в гостях",
                "start": "2012-09-01",
                "time": "18:00",
                "interval": 7,
                "stadium": "Главная арена",
                "teams": [
                    ("Школа 68", "Шкреба"),
                    ("Крылья-школа №7", "Уразаков"),
                    ("Смена-Кошелев", "Майоров"),
                    ("Школа 64/92", "Тетерин"),
                    ("Школа 139", "Гурин"),
                    ("Профики", "Беляев"),
                    ("Смена-Ведиси", "Куликов"),
                    ("Комета", "Женя"),
                    ("Эверест", "Пузырев"),
                    ("Молот", "Цыганов"),
                    ("Союз", "Цыганов"),
                    ("Школа 1", "Галимуллин"),
                    ("Галатасарай", "Лысиков"),
                    ("Коршуны", "Анисимов"),
                    ("Коршуны 2", "Анисимов"),
                    ("Коршуны 3", "Анисимов"),
                ]
            }
        ]
    },
    {
        "season": "2014-2015",
        "format": "7×7",
        "leagues": [
            {
                "name": "2014-2015 ПРЕМЬЕР-ЛИГА 7×7 — 2 КРУГА",
                "format": "7×7",
                "double": True,
                "stage": "2 круга",
                "start": "2014-09-06",
                "time": "18:00",
                "interval": 7,
                "stadium": "Главная арена",
                "teams": [
                    ("Крылья-школа №7", "Уразаков"),
                    ("Смена-Кошелев", "Майоров"),
                    ("Эверест 1", "Пузырев"),
                    ("Профики", "Беляев"),
                    ("КС 26", "Кузьмичев"),
                    ("Школа 1", "Галимуллин"),
                    ("Молния", "Касьянов"),
                    ("Орлы", "Анисимов"),
                ]
            },
            {
                "name": "2014-2015 ВЫСШАЯ ЛИГА 7×7 — ГРУППЫ + ЗОЛОТО/СЕРЕБРО",
                "format": "7×7",
                "double": False,
                "stage": "Группы 3×6 → Золотая (1-3 места) 1 круг / Серебряная (4-6) 1 круг",
                "start": "2014-09-07",
                "time": "17:00",
                "interval": 7,
                "stadium": "Поле №2",
                "groups": 3,
                "group_names": ["Группа A", "Группа B", "Группа C"],
                "teams": [
                    ("68 школа", "Шкреба"),
                    ("Школа 64/92", "Тетерин"),
                    ("Школа 16", "Смолкин"),
                    ("Комета", "Женя"),
                    ("Школа 35", "Колюжный"),
                    ("Орлы 2", "Анисимов"),
                    ("Школа 175/96", "Лысиков"),
                    ("Смена КС (К)", "Якимов"),
                    ("КС 26", "Кузьмичев"),
                    ("Молот", "Цыганов"),
                    ("Олимп", "Седин"),
                    ("Орлы 3", "Анисимов"),
                    ("Крылья дубль-школа №7", "Уразаков"),
                    ("Школа 139", "Гурин"),
                    ("Смена-Ведиси", "Куликов"),
                    ("Галатасарай", "Лысиков"),
                    ("Орлы 4", "Анисимов"),
                    ("Орлы 5", "Анисимов"),
                ],
                "second_stage": {
                    "gold": {"name": "ЗОЛОТАЯ ЛИГА — 1-3 места групп (9 команд) — 1 круг", "spots": ["1A","2A","3A","1B","2B","3B","1C","2C","3C"]},
                    "silver": {"name": "СЕРЕБРЯНАЯ ЛИГА — 4-6 места групп (9 команд) — 1 круг", "spots": ["4A","5A","6A","4B","5B","6B","4C","5C","6C"]},
                },
                "second_start": "2014-11-02",
            },
        ]
    },
    {
        "season": "2016-2017",
        "format": "7×7 / 5×5",
        "leagues": [
            {
                "name": "2016-2017 ПРЕМЬЕР-ЛИГА 7×7 — 2 КРУГА",
                "format": "7×7",
                "double": True,
                "stage": "2 круга",
                "start": "2016-09-03",
                "time": "18:00",
                "interval": 7,
                "stadium": "Главная арена",
                "teams": [
                    ("68 школа", "Шкреба"),
                    ("Школа 175/96", "Лысиков"),
                    ("Крылья-Школа №7", "Уразаков"),
                    ("Смена-Кошелев", "Майоров"),
                    ("Профики", "Беляев"),
                    ("школа 16", "Смолкин"),
                    ("КС 26", "Кузьмичев"),
                    ("Комета", "Женя"),
                    ("Школа 35", "Колюжный"),
                    ("Красная фурия", "Касьянов"),
                    ("Огонь", "Анисимов"),
                    ("ФШМ", "Винокуров"),
                ]
            },
            {
                "name": "2016-2017 ВЫСШАЯ ЛИГА 5×5 — 5 ГРУПП ×4 + ЗОЛОТО/СЕРЕБРО",
                "format": "5×5",
                "double": False,
                "stage": "Группы 5×4 → Золотая (1-2 места) 1 круг / Серебряная (3-4) 1 круг",
                "start": "2016-09-04",
                "time": "16:00",
                "interval": 7,
                "stadium": "Манеж 5×5",
                "groups": 5,
                "group_names": ["Группа A", "Группа B", "Группа C", "Группа D", "Группа E"],
                "teams": [
                    ("Эверест 2", "Пузырев"),
                    ("Профики-2", "Беляев"),
                    ("Молот", "Цыганов"),
                    ("Огонь 2", "Анисимов"),
                    ("Школа 64/92", "Тетерин"),
                    ("школа 16", "Смолкин"),
                    ("Школа 1", "Галимуллин"),
                    ("Огонь 3", "Анисимов"),
                    ("Школа 139", "Гурин"),
                    ("КС 26", "Кузьмичев"),
                    ("Школа 1/2", "Владимиров"),
                    ("Огонь 4", "Анисимов"),
                    ("Профики", "Беляев"),
                    ("Комета", "Женя"),
                    ("Барсы", "Седин"),
                    ("Огонь 5", "Анисимов"),
                    ("Красная фурия 2", "Касьянов"),
                    ("Олимп 2", "Щадин"),
                    ("Дружба", "Майоров"),
                    ("Эверест 3 (резерв)", "Пузырев"),
                ],
                "second_stage": {
                    "gold": {"name": "ЗОЛОТАЯ ЛИГА — 1-2 места групп (10 команд) — 1 круг", "spots": ["1A","2A","1B","2B","1C","2C","1D","2D","1E","2E"]},
                    "silver": {"name": "СЕРЕБРЯНАЯ ЛИГА — 3-4 места групп (10 команд) — 1 круг", "spots": ["3A","4A","3B","4B","3C","4C","3D","4D","3E","4E"]},
                },
                "second_start": "2016-10-23",
            },
        ]
    },
    {
        "season": "2018-2019",
        "format": "5×5 • 2 круга",
        "leagues": [
            {
                "name": "2018-2019 ПРЕМЬЕР-ЛИГА 5×5 — 2 КРУГА",
                "format": "5×5",
                "double": True,
                "stage": "2 круга",
                "start": "2018-09-08",
                "time": "18:00",
                "interval": 7,
                "stadium": "Манеж 5×5",
                "teams": [
                    ("68 школа", "Шкреба"),
                    ("Спутник-1", "Агарков"),
                    ("Смена-Кошелев", "Майоров"),
                    ("Комета", "Женя"),
                    ("Молот", "Цыганов"),
                    ("Акулы 42/11", "Щадин"),
                    ("Красные опасные", "Анисимов"),
                    ("Профики", "Беляев"),
                ]
            },
            {
                "name": "2018-2019 ВЫСШАЯ ЛИГА 5×5 — 2 КРУГА",
                "format": "5×5",
                "double": True,
                "stage": "2 круга",
                "start": "2018-09-09",
                "time": "17:00",
                "interval": 7,
                "stadium": "Поле 5×5 №2",
                "teams": [
                    ("Акулы 42/11", "Щадин"),
                    ("Школа 64/92", "Тетерин"),
                    ("Профики", "Беляев"),
                    ("Школа 16", "Смолкин"),
                    ("КС 26", "Кузьмичев"),
                    ("Школа 1", "Галимуллин"),
                    ("Школа 35", "Колюжный"),
                    ("Дружба", "Майоров"),
                    ("Красные опасные 2", "Анисимов"),
                    ("Красные опасные 3", "Анисимов"),
                ]
            },
        ]
    },
    {
        "season": "2020-2021",
        "format": "5×5",
        "leagues": [
            {
                "name": "2020-2021 — 1 КРУГ + ЗОЛОТО/СЕРЕБРО",
                "format": "5×5",
                "double": False,
                "stage": "1 круг → Золотая (1-6 места) 1 круг / Серебряная (7-13) 1 круг",
                "start": "2020-09-05",
                "time": "17:00",
                "interval": 7,
                "stadium": "Главная арена",
                "teams": [
                    ("68 школа", "Шкреба"),
                    ("Спутник", "Агарков"),
                    ("Киты 42/11", "Щадин"),
                    ("Профики", "Беляев"),
                    ("школа 76", "Глинин"),
                    ("Молот (153/157 школа)", "Вукалов"),
                    ("Школа 1", "Галимуллин"),
                    ("Барселона", "Ботов"),
                    ("Крепыши", "Анисимов"),
                    ("Крепыши 2", "Анисимов"),
                    ("Крепыши 3", "Анисимов"),
                    ("ЯнгФорс", "Сусляев"),
                    ("Профики 2", "Беляев"),
                ],
                "second_stage": {
                    "gold": {"name": "ЗОЛОТАЯ ЛИГА — 1-6 места первого этапа (6 команд) — 1 круг", "spots": ["1 место","2 место","3 место","4 место","5 место","6 место"]},
                    "silver": {"name": "СЕРЕБРЯНАЯ ЛИГА — 7-13 места первого этапа (7 команд) — 1 круг", "spots": ["7 место","8 место","9 место","10 место","11 место","12 место","13 место"]},
                },
                "second_start": "2020-11-28",
            }
        ]
    },
]

# --- helpers ---
def round_robin(team_names, double=False):
    teams = list(team_names)
    is_odd = len(teams) % 2 == 1
    if is_odd:
        teams.append("— выходной —")
    n = len(teams)
    rounds = n - 1
    half = n // 2
    arr = list(teams)
    schedule = []
    gid = 1
    for r in range(rounds):
        for i in range(half):
            home = arr[i]
            away = arr[n-1-i]
            if home == "— выходной —" or away == "— выходной —":
                continue
            is_swap = (r % 2 == 1) and (i == 0)
            if is_swap:
                home, away = away, home
            schedule.append({"id": gid, "round": r+1, "home": home, "away": away})
            gid += 1
        fixed = arr[0]
        rest = arr[1:]
        rest = [rest[-1]] + rest[:-1]
        arr = [fixed] + rest
    if double:
        second = []
        for m in schedule:
            second.append({"id": gid, "round": m["round"] + rounds, "home": m["away"], "away": m["home"]})
            gid += 1
        schedule = schedule + second
    return schedule

# Styles
HEADER_FILL = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
HEADER_FONT = Font(name="Calibri", color="FFFFFF", bold=True, size=10)
SUB_HEADER_FILL = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
SUB_HEADER_FONT = Font(name="Calibri", color="334155", bold=True, size=9)
TITLE_FILL = PatternFill(start_color="16A34A", end_color="16A34A", fill_type="solid")
TITLE_FONT = Font(name="Calibri", color="FFFFFF", bold=True, size=13)
ALT_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
thin = Side(style="thin", color="CBD5E1")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def split_groups(teams, n_groups):
    groups = []
    chunk = len(teams) // n_groups
    rem = len(teams) % n_groups
    idx = 0
    for i in range(n_groups):
        size = chunk + (1 if i < rem else 0)
        groups.append(teams[idx:idx+size])
        idx += size
    return groups

def create_summary_sheet(wb):
    ws = wb.active
    ws.title = "ОГЛАВЛЕНИЕ"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.print_title_rows = "1:1"
    ws.merge_cells("A1:F1")
    c = ws["A1"]
    c.value = "МИХАИЛ • РАСПИСАНИЕ МАТЧЕЙ • РЕГЛАМЕНТ СПОРТДИРЕКТОРА"
    c.font = Font(name="Calibri", size=14, bold=True, color="0F172A")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A2:F2")
    c = ws["A2"]
    c.value = "2012/13 — 2 круга | 2014/15 Премьер 2 круга / Высшая группы + Золото/Серебро | 2016/17 Премьер 2 круга / Высшая 5×4 + Золото/Серебро | 2018/19 — 2 круга | 2020/21 — 1 круг + Золото(1-6)/Серебро(7-13)"
    c.font = Font(name="Calibri", size=7, color="64748B", italic=True)
    c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 24
    headers = ["Сезон", "Лига", "Формат / Регламент", "Команд", "Матчей всего", "В т.ч. 2 этап"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[4].height = 30
    row = 5
    for s in SEASONS:
        for lg in s["leagues"]:
            # count
            total = 0
            second = 0
            if "groups" in lg and "second_stage" in lg:
                groups = split_groups(lg["teams"], lg["groups"])
                group_matches = sum(len(round_robin([t[0] for t in g], lg["double"])) for g in groups)
                gold = len(round_robin(lg["second_stage"]["gold"]["spots"], False))
                silver = len(round_robin(lg["second_stage"]["silver"]["spots"], False))
                total = group_matches + gold + silver
                second = gold + silver
            elif "second_stage" in lg and "groups" not in lg:
                # 2020-2021
                first = len(round_robin([t[0] for t in lg["teams"]], lg["double"]))
                gold = len(round_robin(lg["second_stage"]["gold"]["spots"], False))
                silver = len(round_robin(lg["second_stage"]["silver"]["spots"], False))
                total = first + gold + silver
                second = gold + silver
            else:
                total = len(round_robin([t[0] for t in lg["teams"]], lg["double"]))
            stage_txt = lg.get("stage","")
            ws.cell(row=row, column=1, value=s["season"]).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=row, column=2, value=lg["name"]).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.cell(row=row, column=3, value=stage_txt).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.cell(row=row, column=3).font = Font(size=7, color="334155")
            ws.cell(row=row, column=4, value=len(lg["teams"])).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=row, column=5, value=total).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=row, column=5).font = Font(bold=True, color="16A34A")
            ws.cell(row=row, column=6, value=second if second else "—").alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=row, column=6).font = Font(size=9, color="6B7280")
            is_alt = row % 2 == 0
            for c in range(1,7):
                cell = ws.cell(row=row, column=c)
                cell.border = BORDER
                if is_alt:
                    cell.fill = ALT_FILL
                cell.font = Font(name="Calibri", size=9, color="0F172A") if c!=5 else Font(name="Calibri", size=9, bold=True, color="16A34A")
                if c==2:
                    cell.font = Font(name="Calibri", size=8, bold=True)
            if "Золото" in stage_txt or "ЗОЛОТО" in lg["name"]:
                ws.cell(row=row, column=2).fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
            ws.row_dimensions[row].height = 28
            row += 1
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 36
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 12
    ws.freeze_panes = "A5"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:4"

def create_league_sheet(wb, league, season_name):
    name = league["name"]
    # sanitize for Excel
    invalid = '\\/?*[]:'
    for ch in invalid:
        name = name.replace(ch, '-')
    sheet_name = name[:31]
    base = sheet_name
    idx = 2
    while sheet_name in wb.sheetnames:
        suffix = f" ({idx})"
        sheet_name = (base[:31-len(suffix)] + suffix)
        idx += 1
    # также чистим base для дублей
    for ch in invalid:
        sheet_name = sheet_name.replace(ch, '-')
    ws = wb.create_sheet(title=sheet_name)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4
    ws.print_title_rows = "1:10"
    is_grouped = "groups" in league
    has_second = "second_stage" in league
    # Title
    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = name.upper()
    c.font = TITLE_FONT
    c.fill = TITLE_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26
    ws.merge_cells("A2:H2")
    c = ws["A2"]
    double_txt = "ДВОЙНОЙ КРУГ" if league["double"] else "ОДИН КРУГ"
    extra = league.get("stage","")
    if is_grouped:
        c.value = f"Сезон {season_name}  •  {league['format']}  •  {extra}  •  Старт: {league['start']}  •  {league['time']}  •  {league['stadium']}  •  {len(league['teams'])} команд"
    else:
        c.value = f"Сезон {season_name}  •  {league['format']}  •  {double_txt}  •  {extra}  •  Старт: {league['start']}  •  {league['time']}  •  {len(league['teams'])} команд"
    c.font = Font(name="Calibri", size=7, color="FFFFFF", bold=True)
    c.fill = PatternFill(start_color="15803D", end_color="15803D", fill_type="solid")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 28
    ws.merge_cells("A3:H3")
    c = ws["A3"]
    if is_grouped and has_second:
        c.value = "РЕГЛАМЕНТ СПОРТДИРЕКТОРА: Групповой этап → Золотая и Серебряная лиги по итогам групп. Расписание второго этапа — заглушки по местам (1A,2A...), после групп впиши реальные команды."
    elif has_second:
        c.value = "РЕГЛАМЕНТ: 1 круг предварительного этапа → Золотая (1-6) и Серебряная (7-13) лиги — 1 круг."
    else:
        c.value = f"Формула: каждый с каждым ({double_txt})  •  Всего матчей: {len(round_robin([t[0] for t in league['teams']], league['double']))}  •  Сгенерировано автоматически"
    c.font = Font(name="Calibri", size=8, color="475569", italic=True)
    c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[3].height = 16
    # Teams header
    row = 5
    ws.merge_cells(f"A{row}:H{row}")
    c = ws.cell(row=row, column=1, value="1. КОМАНДЫ И ТРЕНЕРЫ")
    c.font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 20
    row += 1
    team_headers = ["№", "Команда", "Тренер", "Взнос", "Заявка"]
    for col, h in enumerate(team_headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.fill = SUB_HEADER_FILL
        cell.font = SUB_HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = 18
    row += 1
    for idx, (team, trainer) in enumerate(league["teams"], 1):
        ws.cell(row=row, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row, column=1).font = Font(bold=True, color="0F172A")
        ws.cell(row=row, column=1).border = BORDER
        c = ws.cell(row=row, column=2, value=team)
        c.font = Font(bold=True, size=10)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = BORDER
        c = ws.cell(row=row, column=3, value=trainer)
        c.font = Font(size=10, color="475569")
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = BORDER
        ws.cell(row=row, column=4, value="☐").alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row, column=4).border = BORDER
        ws.cell(row=row, column=5, value="☐").alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row, column=5).border = BORDER
        if idx % 2 == 0:
            for col in range(1, 6):
                ws.cell(row=row, column=col).fill = ALT_FILL
        ws.row_dimensions[row].height = 16
        row += 1
    # Schedule
    if is_grouped:
        groups = split_groups(league["teams"], league["groups"])
        group_names = league.get("group_names", [f"Группа {chr(65+i)}" for i in range(len(groups))])
        base_date = datetime.strptime(league["start"] + " " + league["time"], "%Y-%m-%d %H:%M")
        interval = league["interval"]
        row += 1
        ws.merge_cells(f"A{row}:H{row}")
        c = ws.cell(row=row, column=1, value=f"2. ГРУППОВОЙ ЭТАП — {league['groups']} ГРУПП ПО {len(groups[0])}")
        c.font = Font(bold=True, size=11, color="FFFFFF")
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 20
        for gi, group in enumerate(groups):
            gname = group_names[gi]
            row += 1
            ws.merge_cells(f"A{row}:H{row}")
            c = ws.cell(row=row, column=1, value=f"{gname.upper()} — {', '.join([t[0] for t in group])}")
            c.font = Font(bold=True, size=10, color="FFFFFF")
            colors = ["0F172A","1D4ED8","065F46","7C2D12","6D28D9"]
            c.fill = PatternFill(start_color=colors[gi%len(colors)], end_color=colors[gi%len(colors)], fill_type="solid")
            c.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[row].height = 18
            row += 1
            for col_idx, h in enumerate(["№","Команда","Тренер","Взнос","Заявка"], 1):
                cell = ws.cell(row=row, column=col_idx, value=h)
                cell.fill = SUB_HEADER_FILL
                cell.font = SUB_HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = BORDER
            ws.row_dimensions[row].height = 16
            for idx, (team, trainer) in enumerate(group, 1):
                row += 1
                ws.cell(row=row, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(row=row, column=1).font = Font(bold=True, size=9)
                ws.cell(row=row, column=1).border = BORDER
                c = ws.cell(row=row, column=2, value=team)
                c.font = Font(bold=True, size=9)
                c.border = BORDER
                c = ws.cell(row=row, column=3, value=trainer)
                c.font = Font(size=9, color="475569")
                c.border = BORDER
                ws.cell(row=row, column=4, value="☐").border = BORDER
                ws.cell(row=row, column=5, value="☐").border = BORDER
                if idx %2==0:
                    for cc in range(1,6):
                        ws.cell(row=row, column=cc).fill = ALT_FILL
                ws.row_dimensions[row].height = 14
            row += 1
            ws.merge_cells(f"A{row}:H{row}")
            c = ws.cell(row=row, column=1, value=f"Расписание {gname} — каждый с каждым (1 круг)")
            c.font = Font(bold=True, size=9, color="334155")
            c.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border = BORDER
            ws.row_dimensions[row].height = 16
            row += 1
            for col_idx, h in enumerate(["Тур","Дата","Время","Хозяева","—","Гости","Стадион","Счёт"], 1):
                cell = ws.cell(row=row, column=col_idx, value=h)
                cell.fill = SUB_HEADER_FILL
                cell.font = SUB_HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = BORDER
            ws.row_dimensions[row].height = 16
            g_names = [t[0] for t in group]
            sched = round_robin(g_names, league["double"])
            for m in sched:
                d = base_date + timedelta(days=(m["round"]-1)*interval)
                m["date"] = d.strftime("%d.%m.%Y")
                m["time_val"] = league["time"]
                m["stadium"] = league["stadium"] + f" ({gname})"
            for m in sorted(sched, key=lambda x: (x["round"], x["id"])):
                row += 1
                ws.cell(row=row, column=1, value=m["round"]).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(row=row, column=1).font = Font(bold=True, size=9)
                ws.cell(row=row, column=1).border = BORDER
                ws.cell(row=row, column=2, value=m["date"]).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(row=row, column=2).border = BORDER
                ws.cell(row=row, column=3, value=m["time_val"]).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(row=row, column=3).border = BORDER
                c = ws.cell(row=row, column=4, value=m["home"])
                c.font = Font(bold=True, size=9)
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.border = BORDER
                ws.cell(row=row, column=5, value="—").alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(row=row, column=5).border = BORDER
                c = ws.cell(row=row, column=6, value=m["away"])
                c.font = Font(bold=True, size=9)
                c.alignment = Alignment(horizontal="left", vertical="center")
                c.border = BORDER
                ws.cell(row=row, column=7, value=m["stadium"]).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(row=row, column=7).font = Font(size=7, color="475569")
                ws.cell(row=row, column=7).border = BORDER
                ws.cell(row=row, column=8, value=" : ").alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(row=row, column=8).font = Font(bold=True, size=10)
                ws.cell(row=row, column=8).border = BORDER
                if row %2==0:
                    for cc in range(1,9):
                        if ws.cell(row=row, column=cc).fill.fill_type is None:
                            ws.cell(row=row, column=cc).fill = ALT_FILL
                ws.row_dimensions[row].height = 14
            row += 1
        # Second stage for grouped
        if has_second:
            base2 = datetime.strptime(league.get("second_start", league["start"]) + " " + league["time"], "%Y-%m-%d %H:%M")
            for stage_key, title in [("gold","ЗОЛОТАЯ ЛИГА"), ("silver","СЕРЕБРЯНАЯ ЛИГА")]:
                spots = league["second_stage"][stage_key]["spots"]
                st_name = league["second_stage"][stage_key]["name"]
                row += 2
                ws.merge_cells(f"A{row}:H{row}")
                c = ws.cell(row=row, column=1, value=f"3. {title} — {st_name}")
                c.font = Font(bold=True, size=11, color="FFFFFF")
                c.fill = PatternFill(start_color="B45309" if stage_key=="gold" else "475569", end_color="B45309" if stage_key=="gold" else "475569", fill_type="solid")
                c.alignment = Alignment(horizontal="center", vertical="center")
                ws.row_dimensions[row].height = 22
                row += 1
                ws.merge_cells(f"A{row}:H{row}")
                c = ws.cell(row=row, column=1, value=f"Состав {title.lower()}: {', '.join(spots)} — места определяются после групп. Впиши реальные команды вместо 1A,2A...")
                c.font = Font(size=8, italic=True, color="475569")
                c.alignment = Alignment(horizontal="center", wrap_text=True)
                ws.row_dimensions[row].height = 18
                row += 1
                for col_idx, h in enumerate(["Тур","Дата","Время","Хозяева","—","Гости","Стадион","Счёт"], 1):
                    cell = ws.cell(row=row, column=col_idx, value=h)
                    cell.fill = SUB_HEADER_FILL
                    cell.font = SUB_HEADER_FONT
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = BORDER
                ws.row_dimensions[row].height = 16
                sched2 = round_robin(spots, False)
                for m in sched2:
                    d = base2 + timedelta(days=(m["round"]-1)*7)
                    m["date"] = d.strftime("%d.%m.%Y")
                    m["time_val"] = league["time"]
                    m["stadium"] = "Главная арена (финал)"
                for m in sorted(sched2, key=lambda x: (x["round"], x["id"])):
                    row += 1
                    ws.cell(row=row, column=1, value=m["round"]).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(row=row, column=1).font = Font(bold=True, size=9)
                    ws.cell(row=row, column=1).border = BORDER
                    # highlight place-based names
                    ws.cell(row=row, column=2, value=m["date"]).border = BORDER
                    ws.cell(row=row, column=3, value=m["time_val"]).border = BORDER
                    c = ws.cell(row=row, column=4, value=m["home"])
                    c.font = Font(bold=True, size=9, color="92400E" if stage_key=="gold" else "334155")
                    c.alignment = Alignment(horizontal="right", vertical="center")
                    c.border = BORDER
                    c.fill = PatternFill(start_color="FFFBEB" if stage_key=="gold" else "F8FAFC", end_color="FFFBEB" if stage_key=="gold" else "F8FAFC", fill_type="solid")
                    ws.cell(row=row, column=5, value="—").border = BORDER
                    c = ws.cell(row=row, column=6, value=m["away"])
                    c.font = Font(bold=True, size=9, color="92400E" if stage_key=="gold" else "334155")
                    c.alignment = Alignment(horizontal="left", vertical="center")
                    c.border = BORDER
                    c.fill = PatternFill(start_color="FFFBEB" if stage_key=="gold" else "F8FAFC", end_color="FFFBEB" if stage_key=="gold" else "F8FAFC", fill_type="solid")
                    ws.cell(row=row, column=7, value=m["stadium"]).border = BORDER
                    ws.cell(row=row, column=8, value=" : ").border = BORDER
                    ws.row_dimensions[row].height = 14
                base2 += timedelta(days=len(set([m["round"] for m in sched2]))*7 + 7)
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 9
        ws.column_dimensions["D"].width = 22
        ws.column_dimensions["E"].width = 4
        ws.column_dimensions["F"].width = 22
        ws.column_dimensions["G"].width = 18
        ws.column_dimensions["H"].width = 9
        ws.freeze_panes = "A8"
        ws.print_area = f"A1:H{row}"
        return ws
    elif has_second and not is_grouped:
        # 2020-2021 single group + gold/silver
        names = [t[0] for t in league["teams"]]
        schedule = round_robin(names, league["double"])
        base_date = datetime.strptime(league["start"] + " " + league["time"], "%Y-%m-%d %H:%M")
        interval = league["interval"]
        for m in schedule:
            d = base_date + timedelta(days=(m["round"]-1)*interval)
            m["date"] = d.strftime("%d.%m.%Y")
            m["time_val"] = league["time"]
            m["stadium"] = league["stadium"]
        row += 1
        ws.merge_cells(f"A{row}:H{row}")
        c = ws.cell(row=row, column=1, value=f"2. ПРЕДВАРИТЕЛЬНЫЙ ЭТАП — 1 КРУГ ({len(names)} команд, {len(schedule)} матчей)")
        c.font = Font(bold=True, size=11, color="FFFFFF")
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 20
        row += 1
        for col, h in enumerate(["Тур","Дата","Время","Хозяева","","Гости","Стадион","Счёт"], 1):
            cell = ws.cell(row=row, column=col, value=h if h else "—")
            cell.fill = SUB_HEADER_FILL
            cell.font = SUB_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER
        ws.row_dimensions[row].height = 16
        for m in sorted(schedule, key=lambda x: (x["round"], x["id"])):
            row += 1
            ws.cell(row=row, column=1, value=m["round"]).border = BORDER
            ws.cell(row=row, column=1).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=row, column=2, value=m["date"]).border = BORDER
            ws.cell(row=row, column=3, value=m["time_val"]).border = BORDER
            c = ws.cell(row=row, column=4, value=m["home"])
            c.font = Font(bold=True, size=9)
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.border = BORDER
            ws.cell(row=row, column=5, value="—").border = BORDER
            c = ws.cell(row=row, column=6, value=m["away"])
            c.font = Font(bold=True, size=9)
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border = BORDER
            ws.cell(row=row, column=7, value=m["stadium"]).border = BORDER
            ws.cell(row=row, column=8, value=" : ").border = BORDER
            ws.row_dimensions[row].height = 14
        # second stage
        base2 = datetime.strptime(league.get("second_start", league["start"]) + " " + league["time"], "%Y-%m-%d %H:%M")
        for stage_key in ["gold","silver"]:
            spots = league["second_stage"][stage_key]["spots"]
            st_name = league["second_stage"][stage_key]["name"]
            row += 2
            ws.merge_cells(f"A{row}:H{row}")
            c = ws.cell(row=row, column=1, value=f"{'3.' if stage_key=='gold' else '4.'} {st_name.upper()}")
            c.font = Font(bold=True, size=11, color="FFFFFF")
            c.fill = PatternFill(start_color="B45309" if stage_key=="gold" else "475569", end_color="B45309" if stage_key=="gold" else "475569", fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[row].height = 22
            row += 1
            for col, h in enumerate(["Тур","Дата","Время","Хозяева","","Гости","Стадион","Счёт"], 1):
                cell = ws.cell(row=row, column=col, value=h if h else "—")
                cell.fill = SUB_HEADER_FILL
                cell.font = SUB_HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = BORDER
            sched2 = round_robin(spots, False)
            for m in sched2:
                d = base2 + timedelta(days=(m["round"]-1)*7)
                m["date"] = d.strftime("%d.%m.%Y")
                m["time_val"] = league["time"]
                m["stadium"] = "Главная арена (финал)"
            for m in sorted(sched2, key=lambda x: (x["round"], x["id"])):
                row += 1
                ws.cell(row=row, column=1, value=m["round"]).border = BORDER
                ws.cell(row=row, column=1).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(row=row, column=2, value=m["date"]).border = BORDER
                ws.cell(row=row, column=3, value=m["time_val"]).border = BORDER
                c = ws.cell(row=row, column=4, value=m["home"])
                c.font = Font(bold=True, size=9, color="92400E" if stage_key=="gold" else "334155")
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.border = BORDER
                c.fill = PatternFill(start_color="FFFBEB" if stage_key=="gold" else "F8FAFC", end_color="FFFBEB" if stage_key=="gold" else "F8FAFC", fill_type="solid")
                ws.cell(row=row, column=5, value="—").border = BORDER
                c = ws.cell(row=row, column=6, value=m["away"])
                c.font = Font(bold=True, size=9, color="92400E" if stage_key=="gold" else "334155")
                c.alignment = Alignment(horizontal="left", vertical="center")
                c.border = BORDER
                c.fill = PatternFill(start_color="FFFBEB" if stage_key=="gold" else "F8FAFC", end_color="FFFBEB" if stage_key=="gold" else "F8FAFC", fill_type="solid")
                ws.cell(row=row, column=7, value=m["stadium"]).border = BORDER
                ws.cell(row=row, column=8, value=" : ").border = BORDER
                ws.row_dimensions[row].height = 14
            base2 += timedelta(days=len(set([m["round"] for m in sched2]))*7 + 7)
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 9
        ws.column_dimensions["D"].width = 22
        ws.column_dimensions["E"].width = 4
        ws.column_dimensions["F"].width = 22
        ws.column_dimensions["G"].width = 18
        ws.column_dimensions["H"].width = 9
        ws.freeze_panes = "A8"
        ws.print_area = f"A1:H{row}"
        return ws
    else:
        names = [t[0] for t in league["teams"]]
        schedule = round_robin(names, league["double"])
        base_date = datetime.strptime(league["start"] + " " + league["time"], "%Y-%m-%d %H:%M")
        interval = league["interval"]
        for m in schedule:
            d = base_date + timedelta(days=(m["round"]-1)*interval)
            m["date"] = d.strftime("%d.%m.%Y")
            m["time_val"] = league["time"]
            m["stadium"] = league["stadium"]
        row += 1
        ws.merge_cells(f"A{row}:H{row}")
        c = ws.cell(row=row, column=1, value="2. РАСПИСАНИЕ МАТЧЕЙ — " + ("ДВОЙНОЙ КРУГ" if league["double"] else "КРУГОВАЯ СИСТЕМА"))
        c.font = Font(bold=True, size=11, color="FFFFFF")
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 20
        row += 1
        for col, h in enumerate(["Тур","Дата","Время","Хозяева","","Гости","Стадион","Счёт"], 1):
            cell = ws.cell(row=row, column=col, value=h if h else "—")
            cell.fill = SUB_HEADER_FILL
            cell.font = SUB_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER
        header_row = row
        ws.row_dimensions[row].height = 20
        row += 1
        current_round = None
        for m in sorted(schedule, key=lambda x: (x["round"], x["id"])):
            if m["round"] != current_round:
                ws.merge_cells(f"A{row}:H{row}")
                c = ws.cell(row=row, column=1, value=f"ТУР {m['round']}  —  {m['date']}")
                c.font = Font(bold=True, size=9, color="0F172A")
                c.fill = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
                c.alignment = Alignment(horizontal="left", vertical="center")
                c.border = BORDER
                for col in range(1, 9):
                    ws.cell(row=row, column=col).fill = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
                    ws.cell(row=row, column=col).border = BORDER
                ws.row_dimensions[row].height = 18
                row += 1
                current_round = m["round"]
            ws.cell(row=row, column=1, value=m["round"]).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=row, column=1).font = Font(bold=True, size=9)
            ws.cell(row=row, column=1).border = BORDER
            ws.cell(row=row, column=2, value=m["date"]).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=row, column=2).border = BORDER
            ws.cell(row=row, column=3, value=m["time_val"]).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=row, column=3).border = BORDER
            c = ws.cell(row=row, column=4, value=m["home"])
            c.font = Font(bold=True, size=10)
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.border = BORDER
            ws.cell(row=row, column=5, value="—").alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=row, column=5).border = BORDER
            c = ws.cell(row=row, column=6, value=m["away"])
            c.font = Font(bold=True, size=10)
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border = BORDER
            ws.cell(row=row, column=7, value=m["stadium"]).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=row, column=7).border = BORDER
            ws.cell(row=row, column=8, value=" : ").alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=row, column=8).border = BORDER
            if row %2==0:
                for col in range(1, 9):
                    if ws.cell(row=row, column=col).fill.fill_type is None:
                        ws.cell(row=row, column=col).fill = ALT_FILL
            ws.row_dimensions[row].height = 15
            row += 1
        row += 1
        ws.merge_cells(f"A{row}:H{row}")
        c = ws.cell(row=row, column=1, value="* Счёт вписывай в столбец «Счёт» (напр. 2:1).")
        c.font = Font(size=8, italic=True, color="64748B")
        c.alignment = Alignment(horizontal="left", wrap_text=True)
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 9
        ws.column_dimensions["D"].width = 22
        ws.column_dimensions["E"].width = 4
        ws.column_dimensions["F"].width = 22
        ws.column_dimensions["G"].width = 16
        ws.column_dimensions["H"].width = 9
        ws.freeze_panes = "A8"
        ws.print_area = f"A1:H{row}"
        ws.auto_filter.ref = f"A{header_row}:H{row-2}"
        return ws

def main():
    wb = openpyxl.Workbook()
    create_summary_sheet(wb)
    for s in SEASONS:
        for lg in s["leagues"]:
            create_league_sheet(wb, lg, s["season"])
    wb.active = 0
    out_main = "Mikhail_5_Sezonov_Raspisanie.xlsx"
    wb.save(out_main)
    print(f"Saved {out_main} with {len(wb.sheetnames)} sheets: {wb.sheetnames}")
    for s in SEASONS:
        wb2 = openpyxl.Workbook()
        wb2.remove(wb2.active)
        for lg in s["leagues"]:
            create_league_sheet(wb2, lg, s["season"])
        fname = f"Mikhail_{s['season'].replace('-','_')}_{s['leagues'][0]['format'].replace('×','x')}.xlsx"
        fname = fname.replace("/", "_").replace(" ", "_")
        wb2.save(fname)
        print(f"Saved {fname}")

if __name__ == "__main__":
    main()
