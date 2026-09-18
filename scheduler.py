#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Спортдиректор — расписание с учётом полей, времени и тренеров (без параллелей)
Даты и поля как дал клиент
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, timedelta
from collections import defaultdict

# ---------- Конфиг времени ----------
# 2*25 +5 +5 =60 для всех кроме 2020-2021 где 2*20+5+15=60 — всё равно 60 мин слот
SLOT_DURATION = 60  # минут на игру

# Даты как дал клиент: дата, время начала-конца, поля
# Формат: (дата_str, start, end, fields_dict)
# fields_dict: {"7*7": n, "5*5": n, "4*4": n} — количество полей
RAW_DATES = [
    ("26.08", "17:30", "20:30", {"7*7":2, "5*5":3, "4*4":0}),
    ("27.08", "10:00", "13:00", {"7*7":2, "5*5":2, "4*4":3}),
    ("03.10", "17:30", "20:30", {"7*7":2, "5*5":3, "4*4":0}),
    ("04.10", "10:00", "13:00", {"7*7":2, "5*5":3, "4*4":0}),
    ("10.10", "17:30", "20:30", {"7*7":2, "5*5":3, "4*4":0}),
    ("11.10", "10:00", "13:00", {"7*7":2, "5*5":2, "4*4":3}),
    ("17.10", "17:30", "20:30", {"7*7":2, "5*5":3, "4*4":0}),
    ("18.10", "10:00", "13:00", {"7*7":2, "5*5":3, "4*4":0}),
    ("24.10", "17:30", "20:30", {"7*7":2, "5*5":3, "4*4":0}),
    ("25.10", "10:00", "13:00", {"7*7":2, "5*5":2, "4*4":3}),
    ("31.10", "17:00", "21:00", {"7*7":2, "5*5":3, "4*4":0}),
    ("07.11", "17:00", "21:00", {"7*7":2, "5*5":2, "4*4":3}),
    ("14.11", "17:00", "21:00", {"7*7":2, "5*5":3, "4*4":0}),
    ("21.11", "17:00", "21:00", {"7*7":2, "5*5":3, "4*4":0}),
    ("28.11", "14:00", "20:00", {"7*7":2, "5*5":2, "4*4":3}),
    ("05.12", "17:00", "21:00", {"7*7":2, "5*5":3, "4*4":0}),
    ("12.12", "17:00", "21:00", {"7*7":2, "5*5":2, "4*4":3}),
    ("19.12", "17:00", "21:00", {"7*7":2, "5*5":3, "4*4":0}),
    ("26.12", "17:00", "21:00", {"7*7":2, "5*5":2, "4*4":3}),
]

# Предполагаем год 2025 для дат (или 2026) — берём 2025
YEAR = 2025

def parse_time(s):
    h,m = map(int, s.split(":"))
    return h*60+m

def generate_slots():
    slots = []  # list of (date_str, date_obj, start_time_str, field_id, field_type)
    for date_str, start, end, fields in RAW_DATES:
        d,m = map(int, date_str.split("."))
        # даты августа — 2025, октября-декабря — 2025
        # если месяц 8 — август 2025, иначе октябрь-декабрь 2025
        y = YEAR
        # 26.08 и 27.08 — август
        # остальные — октябрь-декабрь
        try:
            date_obj = datetime(y, m, d)
        except:
            date_obj = datetime(YEAR, m, d)
        start_min = parse_time(start)
        end_min = parse_time(end)
        # количество слотов = (end-start)/60
        n_slots = (end_min - start_min)//SLOT_DURATION
        for slot_idx in range(n_slots):
            slot_start_min = start_min + slot_idx*SLOT_DURATION
            hh = slot_start_min//60
            mm = slot_start_min%60
            time_str = f"{hh:02d}:{mm:02d}"
            # поля
            for ftype, cnt in fields.items():
                if cnt==0: continue
                for fid in range(1, cnt+1):
                    field_id = f"{ftype} #{fid}" if cnt>1 else f"{ftype}"
                    if ftype=="7*7":
                        field_id = f"Поле 7×7 #{fid}" if cnt>1 else "Поле 7×7"
                    elif ftype=="5*5":
                        field_id = f"Поле 5×5 #{fid}"
                    elif ftype=="4*4":
                        field_id = f"Поле 4×4 #{fid}"
                    slots.append({
                        "date_str": date_str,
                        "date_obj": date_obj,
                        "time": time_str,
                        "field": field_id,
                        "field_type": ftype,
                        "slot_key": f"{date_str} {time_str}",
                    })
    return slots

# ---------- Команды и матчи ----------
# Берём из прошлого регламента, но для расписания на эти даты — делаем упор на групповые этапы + предварительные
# Для демо — берём ВСЕ матчи, но посчитаем вместимость

def round_robin(names, double=False):
    teams=list(names)
    if len(teams)%2==1:
        teams.append("— выходной —")
    n=len(teams)
    rounds=n-1
    half=n//2
    arr=list(teams)
    sched=[]
    gid=1
    for r in range(rounds):
        for i in range(half):
            home=arr[i]
            away=arr[n-1-i]
            if home=="— выходной —" or away=="— выходной —":
                continue
            is_swap=(r%2==1) and (i==0)
            if is_swap:
                home,away=away,home
            sched.append({"round":r+1, "home":home, "away":away})
            gid+=1
        fixed=arr[0]
        rest=arr[1:]
        rest=[rest[-1]]+rest[:-1]
        arr=[fixed]+rest
    if double:
        second=[]
        for m in sched:
            second.append({"round":m["round"]+rounds, "home":m["away"], "away":m["home"]})
        sched=sched+second
    return sched

# Данные лиг — берём как в generate_excel, но для расписания с тренерами
LEAGUES = [
    {"name":"2012-2013 7×7 2 круга","format":"7*7","teams":[("Школа 68","Шкреба"),("Крылья-школа №7","Уразаков"),("Смена-Кошелев","Майоров"),("Школа 64/92","Тетерин"),("Школа 139","Гурин"),("Профики","Беляев"),("Смена-Ведиси","Куликов"),("Комета","Женя"),("Эверест","Пузырев"),("Молот","Цыганов"),("Союз","Цыганов"),("Школа 1","Галимуллин"),("Галатасарай","Лысиков"),("Коршуны","Анисимов"),("Коршуны 2","Анисимов"),("Коршуны 3","Анисимов")], "double":True},
    {"name":"2014-2015 Премьер 7×7 2кр","format":"7*7","teams":[("Крылья-школа №7","Уразаков"),("Смена-Кошелев","Майоров"),("Эверест 1","Пузырев"),("Профики","Беляев"),("КС 26","Кузьмичев"),("Школа 1","Галимуллин"),("Молния","Касьянов"),("Орлы","Анисимов")], "double":True},
    {"name":"2014-2015 Высшая 3×6 группы","format":"7*7","groups":3, "teams":[("68 школа","Шкреба"),("Школа 64/92","Тетерин"),("Школа 16","Смолкин"),("Комета","Женя"),("Школа 35","Колюжный"),("Орлы 2","Анисимов"),("Школа 175/96","Лысиков"),("Смена КС (К)","Якимов"),("КС 26","Кузьмичев"),("Молот","Цыганов"),("Олимп","Седин"),("Орлы 3","Анисимов"),("Крылья дубль-школа №7","Уразаков"),("Школа 139","Гурин"),("Смена-Ведиси","Куликов"),("Галатасарай","Лысиков"),("Орлы 4","Анисимов"),("Орлы 5","Анисимов")], "double":False, "is_group":True},
    {"name":"2016-2017 Премьер 7×7 2кр","format":"7*7","teams":[("68 школа","Шкреба"),("Школа 175/96","Лысиков"),("Крылья-Школа №7","Уразаков"),("Смена-Кошелев","Майоров"),("Профики","Беляев"),("школа 16","Смолкин"),("КС 26","Кузьмичев"),("Комета","Женя"),("Школа 35","Колюжный"),("Красная фурия","Касьянов"),("Огонь","Анисимов"),("ФШМ","Винокуров")], "double":True},
    {"name":"2016-2017 Высшая 5×5 5×4 группы","format":"5*5","groups":5, "teams":[("Эверест 2","Пузырев"),("Профики-2","Беляев"),("Молот","Цыганов"),("Огонь 2","Анисимов"),("Школа 64/92","Тетерин"),("школа 16","Смолкин"),("Школа 1","Галимуллин"),("Огонь 3","Анисимов"),("Школа 139","Гурин"),("КС 26","Кузьмичев"),("Школа 1/2","Владимиров"),("Огонь 4","Анисимов"),("Профики","Беляев"),("Комета","Женя"),("Барсы","Седин"),("Огонь 5","Анисимов"),("Красная фурия 2","Касьянов"),("Олимп 2","Щадин"),("Дружба","Майоров"),("Эверест 3","Пузырев")], "double":False, "is_group":True},
    {"name":"2018-2019 Премьер 5×5 2кр","format":"5*5","teams":[("68 школа","Шкреба"),("Спутник-1","Агарков"),("Смена-Кошелев","Майоров"),("Комета","Женя"),("Молот","Цыганов"),("Акулы 42/11","Щадин"),("Красные опасные","Анисимов"),("Профики","Беляев")], "double":True},
    {"name":"2018-2019 Высшая 5×5 2кр","format":"5*5","teams":[("Акулы 42/11","Щадин"),("Школа 64/92","Тетерин"),("Профики","Беляев"),("Школа 16","Смолкин"),("КС 26","Кузьмичев"),("Школа 1","Галимуллин"),("Школа 35","Колюжный"),("Дружба","Майоров"),("Красные опасные 2","Анисимов"),("Красные опасные 3","Анисимов")], "double":True},
    {"name":"2020-2021 13 ком 1круг","format":"5*5","teams":[("68 школа","Шкреба"),("Спутник","Агарков"),("Киты 42/11","Щадин"),("Профики","Беляев"),("школа 76","Глинин"),("Молот (153/157)","Вукалов"),("Школа 1","Галимуллин"),("Барселона","Ботов"),("Крепыши","Анисимов"),("Крепыши 2","Анисимов"),("Крепыши 3","Анисимов"),("ЯнгФорс","Сусляев"),("Профики 2","Беляев")], "double":False, "is_group_2020":True},
]

def build_all_matches():
    all_matches=[]
    for lg in LEAGUES:
        fmt=lg["format"]
        if lg.get("is_group"):
            # группы
            n_groups=lg["groups"]
            teams=lg["teams"]
            chunk=len(teams)//n_groups
            groups=[teams[i*chunk:(i+1)*chunk] for i in range(n_groups)]
            group_names=[f"Группа {chr(65+i)}" for i in range(n_groups)]
            for gi,g in enumerate(groups):
                names=[t[0] for t in g]
                sched=round_robin(names, lg["double"])
                for m in sched:
                    # найдём тренеров
                    home_train = next((t[1] for t in g if t[0]==m["home"]), "?")
                    away_train = next((t[1] for t in g if t[0]==m["away"]), "?")
                    # для поиска тренера по имени в целом лиге
                    # но для группы достаточно
                    all_matches.append({
                        "league":lg["name"],
                        "format":fmt,
                        "group":group_names[gi],
                        "home":m["home"],
                        "away":m["away"],
                        "home_train":home_train,
                        "away_train":away_train,
                        "round":m["round"],
                    })
        else:
            names=[t[0] for t in lg["teams"]]
            sched=round_robin(names, lg["double"])
            for m in sched:
                home_train=next((t[1] for t in lg["teams"] if t[0]==m["home"]), "?")
                away_train=next((t[1] for t in lg["teams"] if t[0]==m["away"]), "?")
                all_matches.append({
                    "league":lg["name"],
                    "format":fmt,
                    "group":"",
                    "home":m["home"],
                    "away":m["away"],
                    "home_train":home_train,
                    "away_train":away_train,
                    "round":m["round"],
                })
    return all_matches

def schedule_matches(all_matches, slots):
    # Сортируем матчи по "тяжести" тренера — у кого много команд, того первым ставим (Анисимов)
    from collections import Counter
    trainer_counts=Counter()
    for m in all_matches:
        trainer_counts[m["home_train"]]+=1
        trainer_counts[m["away_train"]]+=1
    # вес матча = сумма counts тренеров
    def weight(m):
        return trainer_counts[m["home_train"]] + trainer_counts[m["away_train"]]
    all_matches_sorted=sorted(all_matches, key=weight, reverse=True)

    # Слоты сортируем по дате/времени
    slots_sorted=sorted(slots, key=lambda s: (s["date_obj"], s["time"], s["field"]))

    # Для проверки параллели тренера: busy[slot_key][trainer]=True
    busy=defaultdict(set)  # slot_key -> set(trainers)
    # также busy per field? поле занято одним матчем на слот, но у нас каждый слот — уже конкретное поле+время, так что просто проверяем field занятость? Мы уже итерируемся по слотам, каждый слот — уникальное поле+время, так что field не дублируется.
    # Но нужно проверять, что тренер не занят в это же время на другом поле (same slot_key)
    scheduled=[]
    unscheduled=[]

    # Группируем слоты по времени: slot_key = date+time
    # Для каждого матча ищем первый подходящий слот где формат поля совпадает и оба тренера свободны в это время

    # Подготовим доступность: для каждого slot_key — какие поля есть

    # Создадим индекс слотов по времени
    # Уже есть slots_sorted

    for m in all_matches_sorted:
        placed=False
        needed_format=m["format"]  # "7*7" or "5*5"
        # 4*4 поля могут использоваться для 5*5? пока нет, только точное совпадение, но можно разрешить 5*5 на 4*4? Пока строго
        for s in slots_sorted:
            if s["field_type"] != needed_format:
                # разрешим 5*5 на 4*4 как запасной? пока нет
                # Но если не хватает 5*5, можно использовать 4*4 для 5*5 (с натяжкой)
                # Сделаем fallback: если 5*5 матч и слот 4*4 — считаем возможным, но пометим
                if not (needed_format=="5*5" and s["field_type"]=="4*4"):
                    continue
            # проверяем, слот ещё не занят (каждый слот — одно поле, значит если слот уже в scheduled, он занят)
            # Проверим, что слот не использован
            if any(sc["date"]==s["date_str"] and sc["time"]==s["time"] and sc["field"]==s["field"] for sc in scheduled):
                continue
            # тренер свободен в это время?
            slot_key=s["date_str"]+" "+s["time"]
            if m["home_train"] in busy[slot_key] or m["away_train"] in busy[slot_key]:
                continue
            # также если тренер ведёт обе команды? тогда и home и away один тренер — но это возможно? тогда ок, но всё равно один слот
            # ставим
            scheduled.append({
                "date":s["date_str"],
                "date_obj":s["date_obj"],
                "time":s["time"],
                "field":s["field"],
                "field_type":s["field_type"],
                "league":m["league"],
                "format":m["format"],
                "group":m["group"],
                "home":m["home"],
                "away":m["away"],
                "home_train":m["home_train"],
                "away_train":m["away_train"],
                "round":m["round"],
                "slot_key":slot_key,
                "is_4x4_fallback": (needed_format=="5*5" and s["field_type"]=="4*4"),
            })
            busy[slot_key].add(m["home_train"])
            busy[slot_key].add(m["away_train"])
            placed=True
            break
        if not placed:
            unscheduled.append(m)

    return scheduled, unscheduled, busy

def main():
    slots=generate_slots()
    print(f"Всего слотов: {len(slots)} (7×7:{len([s for s in slots if s['field_type']=='7*7'])}, 5×5:{len([s for s in slots if s['field_type']=='5*5'])}, 4×4:{len([s for s in slots if s['field_type']=='4*4'])})")
    all_matches=build_all_matches()
    print(f"Всего матчей к расписанию (все лиги, только первый этап/группы): {len(all_matches)}")
    # Посчитаем по форматам
    from collections import Counter
    fmt_cnt=Counter(m["format"] for m in all_matches)
    print(f"По форматам: {fmt_cnt}")
    # Легенда тренеров с множеством команд
    trainer_teams=defaultdict(set)
    for lg in LEAGUES:
        for name,tr in lg["teams"]:
            trainer_teams[tr].add(name)
    print("Тренеры с >1 командой:")
    for tr,teams in sorted(trainer_teams.items(), key=lambda x: len(x[1]), reverse=True):
        if len(teams)>1:
            print(f"  {tr}: {len(teams)} — {', '.join(list(teams)[:5])}")

    scheduled, unscheduled, busy = schedule_matches(all_matches, slots)
    print(f"Запланировано: {len(scheduled)}, не влезло: {len(unscheduled)}")
    if unscheduled:
        print(f"Пример невместившихся (5):")
        for m in unscheduled[:5]:
            print(f"  {m['league']} {m['group']} {m['home']}({m['home_train']}) - {m['away']}({m['away_train']})")

    # Генерим Excel
    wb=openpyxl.Workbook()
    ws=wb.active
    ws.title="Расписание"
    # стили
    HEADER_FILL=PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    HEADER_FONT=Font(color="FFFFFF", bold=True, size=9)
    SUB_FILL=PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    ALT_FILL=PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    GOLD_FILL=PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
    thin=Side(style="thin", color="CBD5E1")
    BORDER=Border(left=thin,right=thin,top=thin,bottom=thin)

    # Заголовок
    ws.merge_cells("A1:K1")
    ws["A1"].value="РАСПИСАНИЕ ПО ДАТАМ — С УЧЁТОМ ПОЛЕЙ И ТРЕНЕРОВ (БЕЗ ПАРАЛЛЕЛЕЙ)"
    ws["A1"].font=Font(bold=True, size=12, color="0F172A")
    ws["A1"].alignment=Alignment(horizontal="center")
    ws.merge_cells("A2:K2")
    ws["A2"].value="Слот 60 мин: 2×25+5 перерыв+5 между играми (для 2020-2021 2×20+5+15) — заложен 1 час. Тренер не может быть в двух местах одновременно."
    ws["A2"].font=Font(size=8, italic=True, color="64748B")
    ws["A2"].alignment=Alignment(horizontal="center")

    headers=["Дата","День","Время","Поле","Формат","Лига","Группа","Хозяева","Тренер","Гости","Тренер"]
    for c,h in enumerate(headers,1):
        cell=ws.cell(row=4, column=c, value=h)
        cell.fill=HEADER_FILL
        cell.font=HEADER_FONT
        cell.alignment=Alignment(horizontal="center", vertical="center")
        cell.border=BORDER
    ws.row_dimensions[4].height=22

    # сортировка по дате/времени
    scheduled_sorted=sorted(scheduled, key=lambda x: (x["date_obj"], x["time"], x["field"]))

    # Проверка параллелей тренеров после генерации
    # Проверка параллелей тренеров после генерации
    check=defaultdict(list)
    for sc in scheduled_sorted:
        key=sc["date"]+" "+sc["time"]
        check[key].extend([sc["home_train"], sc["away_train"]])
    conflicts=[]
    for k,trainers in check.items():
        seen=set()
        dups=set()
        for t in trainers:
            if t in seen and t!="—":
                dups.add(t)
            seen.add(t)
        if dups:
            conflicts.append((k,dups))

    row=5
    for sc in scheduled_sorted:
        # день недели
        try:
            dow=sc["date_obj"].strftime("%a")
            # русские дни
            days={"Mon":"Пн","Tue":"Вт","Wed":"Ср","Thu":"Чт","Fri":"Пт","Sat":"Сб","Sun":"Вс"}
            dow=days.get(dow,dow)
        except:
            dow=""
        vals=[sc["date"], dow, sc["time"], sc["field"], sc["format"], sc["league"], sc["group"], sc["home"], sc["home_train"], sc["away"], sc["away_train"]]
        for c,v in enumerate(vals,1):
            cell=ws.cell(row=row, column=c, value=v)
            cell.font=Font(size=8)
            cell.alignment=Alignment(horizontal="center", vertical="center")
            cell.border=BORDER
            if c in (8,10):
                cell.font=Font(size=8, bold=True)
            if c in (9,11):
                cell.font=Font(size=7, color="475569")
            if sc.get("is_4x4_fallback"):
                cell.fill=PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        # подсветка золотого времени?
        if "2020" in sc["league"]:
            for c in range(1,12):
                if ws.cell(row=row, column=c).fill.fill_type is None:
                    ws.cell(row=row, column=c).fill=GOLD_FILL
        ws.row_dimensions[row].height=16
        row+=1

    # ширины
    widths=[8,6,7,16,7,28,10,18,12,18,12]
    for i,w in enumerate(widths,1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    ws.freeze_panes="A5"
    ws.auto_filter.ref=f"A4:K{row-1}"
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.orientation="landscape"
    ws.page_setup.paperSize=ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth=1
    ws.page_setup.fitToHeight=0

    # Лист 2 — сводка
    ws2=wb.create_sheet("Сводка")
    ws2["A1"].value="Сводка спортдиректора"
    ws2["A1"].font=Font(bold=True, size=12)
    ws2["A3"].value=f"Всего слотов: {len(slots)}"
    ws2["A4"].value=f"Запланировано матчей: {len(scheduled)}"
    ws2["A5"].value=f"Не влезло: {len(unscheduled)}"
    ws2["A6"].value=f"Конфликтов тренеров параллельно: {len(conflicts)} (должно быть 0)"
    if conflicts:
        ws2["A7"].value="Конфликты:"
        for i,(k,dups) in enumerate(conflicts[:10],8):
            ws2.cell(row=i, column=1, value=f"{k}: {', '.join(dups)}")
    else:
        ws2["A7"].value="Параллелей у тренеров нет — проверка пройдена ✓"
        ws2["A7"].font=Font(color="16A34A", bold=True)
    ws2["A10"].value="Использование полей:"
    # по датам
    from collections import Counter
    cnt_date=Counter(s["date"] for s in scheduled_sorted)
    r=11
    for date in sorted(set(s["date"] for s in scheduled_sorted), key=lambda d: datetime.strptime(d+".2025","%d.%m.%Y")):
        ws2.cell(row=r, column=1, value=date)
        ws2.cell(row=r, column=2, value=cnt_date[date])
        # сколько 7*7 и 5*5
        c7=len([s for s in scheduled_sorted if s["date"]==date and s["format"]=="7*7"])
        c5=len([s for s in scheduled_sorted if s["date"]==date and s["format"]=="5*5"])
        ws2.cell(row=r, column=3, value=f"7×7:{c7} 5×5:{c5}")
        r+=1
    ws2["A25"].value="Лиги:"
    r=26
    cnt_league=Counter(s["league"] for s in scheduled_sorted)
    for lg,cnt in cnt_league.items():
        ws2.cell(row=r, column=1, value=lg)
        ws2.cell(row=r, column=2, value=cnt)
        r+=1
    if unscheduled:
        ws2.cell(row=r+2, column=1, value="Невместившиеся матчи (нужны доп. даты/поля):")
        ws2.cell(row=r+2, column=1).font=Font(bold=True, color="DC2626")
        for i,m in enumerate(unscheduled[:20], r+3):
            ws2.cell(row=i, column=1, value=f"{m['league']} {m['group']} {m['home']}-{m['away']} ({m['format']})")

    wb.save("Raspisanie_Po_Datam_S_Polyami.xlsx")
    print("Saved Raspisanie_Po_Datam_S_Polyami.xlsx")

    # Также сохраним отдельный файл только с групповыми/предварительными (без 2 кругов двойных) — чтобы влезло
    # Для этого фильтруем только is_group или 2020
    # Но уже сделали полный — посмотрим, сколько не влезло двойных

if __name__=="__main__":
    main()
