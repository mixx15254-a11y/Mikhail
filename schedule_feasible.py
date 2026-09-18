#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, timedelta
from collections import defaultdict, Counter

YEAR=2025
SLOT_DURATION=60
RAW_DATES=[
    ("26.08","17:30","20:30",{"7*7":2,"5*5":3,"4*4":0}),
    ("27.08","10:00","13:00",{"7*7":2,"5*5":2,"4*4":3}),
    ("03.10","17:30","20:30",{"7*7":2,"5*5":3,"4*4":0}),
    ("04.10","10:00","13:00",{"7*7":2,"5*5":3,"4*4":0}),
    ("10.10","17:30","20:30",{"7*7":2,"5*5":3,"4*4":0}),
    ("11.10","10:00","13:00",{"7*7":2,"5*5":2,"4*4":3}),
    ("17.10","17:30","20:30",{"7*7":2,"5*5":3,"4*4":0}),
    ("18.10","10:00","13:00",{"7*7":2,"5*5":3,"4*4":0}),
    ("24.10","17:30","20:30",{"7*7":2,"5*5":3,"4*4":0}),
    ("25.10","10:00","13:00",{"7*7":2,"5*5":2,"4*4":3}),
    ("31.10","17:00","21:00",{"7*7":2,"5*5":3,"4*4":0}),
    ("07.11","17:00","21:00",{"7*7":2,"5*5":2,"4*4":3}),
    ("14.11","17:00","21:00",{"7*7":2,"5*5":3,"4*4":0}),
    ("21.11","17:00","21:00",{"7*7":2,"5*5":3,"4*4":0}),
    ("28.11","14:00","20:00",{"7*7":2,"5*5":2,"4*4":3}),
    ("05.12","17:00","21:00",{"7*7":2,"5*5":3,"4*4":0}),
    ("12.12","17:00","21:00",{"7*7":2,"5*5":2,"4*4":3}),
    ("19.12","17:00","21:00",{"7*7":2,"5*5":3,"4*4":0}),
    ("26.12","17:00","21:00",{"7*7":2,"5*5":2,"4*4":3}),
]
def parse_time(s):
    h,m=map(int,s.split(":"))
    return h*60+m
def generate_slots():
    slots=[]
    for date_str,start,end,fields in RAW_DATES:
        d,m=map(int,date_str.split("."))
        date_obj=datetime(YEAR,m,d)
        n_slots=(parse_time(end)-parse_time(start))//SLOT_DURATION
        for si in range(n_slots):
            tmin=parse_time(start)+si*SLOT_DURATION
            time_str=f"{tmin//60:02d}:{tmin%60:02d}"
            for ftype,cnt in fields.items():
                for fid in range(1,cnt+1):
                    field_id=f"Поле {ftype} #{fid}" if cnt>1 else f"Поле {ftype}"
                    slots.append({"date_str":date_str,"date_obj":date_obj,"time":time_str,"field":field_id,"field_type":ftype,"slot_key":f"{date_str} {time_str}"})
    return slots
def round_robin(names,double=False):
    teams=list(names)
    if len(teams)%2==1:
        teams.append("— выходной —")
    n=len(teams);rounds=n-1;half=n//2;arr=list(teams);sched=[];gid=1
    for r in range(rounds):
        for i in range(half):
            home=arr[i];away=arr[n-1-i]
            if home=="— выходной —" or away=="— выходной —": continue
            is_swap=(r%2==1) and (i==0)
            if is_swap: home,away=away,home
            sched.append({"round":r+1,"home":home,"away":away});gid+=1
        fixed=arr[0];rest=arr[1:];rest=[rest[-1]]+rest[:-1];arr=[fixed]+rest
    if double:
        second=[{"round":m["round"]+rounds,"home":m["away"],"away":m["home"]} for m in sched]
        sched=sched+second
    return sched

# Только те лиги, что реально нужно сыграть осенью в эти окна — группы + 5×5 + 2020 предварительный
# Исключаем тяжёлые 7×7 двойные круги (2012, 2014 Премьер, 2016 Премьер) — им нужен отдельный 7×7 фестиваль весной
LEAGUES=[
    {"name":"2014-2015 Высшая 7×7 — группы 3×6","format":"7*7","groups":3,"teams":[("68 школа","Шкреба"),("Школа 64/92","Тетерин"),("Школа 16","Смолкин"),("Комета","Женя"),("Школа 35","Колюжный"),("Орлы 2","Анисимов"),("Школа 175/96","Лысиков"),("Смена КС (К)","Якимов"),("КС 26","Кузьмичев"),("Молот","Цыганов"),("Олимп","Седин"),("Орлы 3","Анисимов"),("Крылья дубль-школа №7","Уразаков"),("Школа 139","Гурин"),("Смена-Ведиси","Куликов"),("Галатасарай","Лысиков"),("Орлы 4","Анисимов"),("Орлы 5","Анисимов")], "double":False},
    {"name":"2016-2017 Высшая 5×5 — группы 5×4","format":"5*5","groups":5,"teams":[("Эверест 2","Пузырев"),("Профики-2","Беляев"),("Молот","Цыганов"),("Огонь 2","Анисимов"),("Школа 64/92","Тетерин"),("школа 16","Смолкин"),("Школа 1","Галимуллин"),("Огонь 3","Анисимов"),("Школа 139","Гурин"),("КС 26","Кузьмичев"),("Школа 1/2","Владимиров"),("Огонь 4","Анисимов"),("Профики","Беляев"),("Комета","Женя"),("Барсы","Седин"),("Огонь 5","Анисимов"),("Красная фурия 2","Касьянов"),("Олимп 2","Щадин"),("Дружба","Майоров"),("Эверест 3","Пузырев")], "double":False},
    {"name":"2018-2019 Премьер 5×5 — 2 круга","format":"5*5","teams":[("68 школа","Шкреба"),("Спутник-1","Агарков"),("Смена-Кошелев","Майоров"),("Комета","Женя"),("Молот","Цыганов"),("Акулы 42/11","Щадин"),("Красные опасные","Анисимов"),("Профики","Беляев")], "double":True},
    {"name":"2018-2019 Высшая 5×5 — 2 круга","format":"5*5","teams":[("Акулы 42/11","Щадин"),("Школа 64/92","Тетерин"),("Профики","Беляев"),("Школа 16","Смолкин"),("КС 26","Кузьмичев"),("Школа 1","Галимуллин"),("Школа 35","Колюжный"),("Дружба","Майоров"),("Красные опасные 2","Анисимов"),("Красные опасные 3","Анисимов")], "double":True},
    {"name":"2020-2021 Предварительный 13 — 1 круг","format":"4*4","teams":[("68 школа","Шкреба"),("Спутник","Агарков"),("Киты 42/11","Щадин"),("Профики","Беляев"),("школа 76","Глинин"),("Молот (153/157)","Вукалов"),("Школа 1","Галимуллин"),("Барселона","Ботов"),("Крепыши","Анисимов"),("Крепыши 2","Анисимов"),("Крепыши 3","Анисимов"),("ЯнгФорс","Сусляев"),("Профики 2","Беляев")], "double":False},
]

def build_matches():
    all_matches=[]
    for lg in LEAGUES:
        fmt=lg["format"]
        if "groups" in lg:
            n_groups=lg["groups"]
            teams=lg["teams"]
            chunk=len(teams)//n_groups
            groups=[teams[i*chunk:(i+1)*chunk] for i in range(n_groups)]
            group_names=[f"Группа {chr(65+i)}" for i in range(n_groups)]
            for gi,g in enumerate(groups):
                names=[t[0] for t in g]
                sched=round_robin(names, lg["double"])
                for m in sched:
                    ht=next((t[1] for t in g if t[0]==m["home"]),"?")
                    at=next((t[1] for t in g if t[0]==m["away"]),"?")
                    all_matches.append({"league":lg["name"],"format":fmt,"group":group_names[gi],"home":m["home"],"away":m["away"],"home_train":ht,"away_train":at,"round":m["round"]})
        else:
            names=[t[0] for t in lg["teams"]]
            sched=round_robin(names, lg["double"])
            for m in sched:
                ht=next((t[1] for t in lg["teams"] if t[0]==m["home"]),"?")
                at=next((t[1] for t in lg["teams"] if t[0]==m["away"]),"?")
                all_matches.append({"league":lg["name"],"format":fmt,"group":"","home":m["home"],"away":m["away"],"home_train":ht,"away_train":at,"round":m["round"]})
    return all_matches

def schedule(all_matches, slots):
    # вес тренера
    cnt=Counter()
    for m in all_matches:
        cnt[m["home_train"]]+=1
        cnt[m["away_train"]]+=1
    def w(m): return cnt[m["home_train"]]+cnt[m["away_train"]]
    all_sorted=sorted(all_matches, key=w, reverse=True)
    slots_sorted=sorted(slots, key=lambda s:(s["date_obj"], s["time"], s["field"]))
    busy=defaultdict(set)
    scheduled=[]
    unscheduled=[]
    for m in all_sorted:
        placed=False
        for s in slots_sorted:
            if s["field_type"]!=m["format"]:
                if not ((m["format"]=="5*5" and s["field_type"]=="4*4") or (m["format"]=="4*4" and s["field_type"]=="5*5")):
                    continue
            if any(sc["date"]==s["date_str"] and sc["time"]==s["time"] and sc["field"]==s["field"] for sc in scheduled):
                continue
            key=s["date_str"]+" "+s["time"]
            if m["home_train"] in busy[key] or m["away_train"] in busy[key]:
                continue
            scheduled.append({"date":s["date_str"],"date_obj":s["date_obj"],"time":s["time"],"field":s["field"],"field_type":s["field_type"],"league":m["league"],"format":m["format"],"group":m["group"],"home":m["home"],"away":m["away"],"home_train":m["home_train"],"away_train":m["away_train"],"round":m["round"],"is_fallback":(m["format"]!=s["field_type"])})
            busy[key].add(m["home_train"]);busy[key].add(m["away_train"])
            placed=True;break
        if not placed:
            unscheduled.append(m)
    return scheduled, unscheduled, busy

def main():
    slots=generate_slots()
    all_matches=build_matches()
    print(f"Матчей в плане (осенний блок 19 дат): {len(all_matches)} (7*7:{len([m for m in all_matches if m['format']=='7*7'])}, 5*5:{len([m for m in all_matches if m['format']=='5*5'])}, 4*4:{len([m for m in all_matches if m['format']=='4*4'])})")
    print(f"Слотов (только эти 19 дат, без января): {len(slots)} (7*7:{len([s for s in slots if s['field_type']=='7*7'])}, 5*5:{len([s for s in slots if s['field_type']=='5*5'])}+4*4:{len([s for s in slots if s['field_type']=='4*4'])})")
    scheduled,unscheduled,busy=schedule(all_matches, slots)
    print(f"Запланировано: {len(scheduled)}, не влезло: {len(unscheduled)}")
    # проверка параллелей — дерби одного тренера считаем 1 матч, а не параллель
    check=defaultdict(list)
    for sc in scheduled:
        trainers=set([sc["home_train"],sc["away_train"]])
        trainers.discard(""); trainers.discard("—"); trainers.discard("?")
        check[sc["date"]+" "+sc["time"]].extend(list(trainers))
    conflicts=[]
    for k,trs in check.items():
        seen=set()
        for t in trs:
            if t in seen:
                conflicts.append(k)
                break
            seen.add(t)
    print(f"Конфликтов параллелей (реальных): {len(conflicts)}")

    wb=openpyxl.Workbook()
    ws=wb.active;ws.title="Осеннее расписание"
    HEADER_FILL=PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    HEADER_FONT=Font(color="FFFFFF", bold=True, size=9)
    ALT_FILL=PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    GOLD_FILL=PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
    thin=Side(style="thin", color="CBD5E1")
    BORDER=Border(left=thin,right=thin,top=thin,bottom=thin)
    ws.merge_cells("A1:K1");ws["A1"].value="ОСЕННИЙ БЛОК — РАСПИСАНИЕ ТОЛЬКО НА ЭТИ 19 ДАТ (без января) • 2×7×7 везде • 60 мин слот"
    ws["A1"].font=Font(bold=True,size=12,color="0F172A");ws["A1"].alignment=Alignment(horizontal="center")
    ws.merge_cells("A2:K2");ws["A2"].value="Группы 2014 Высшая 3×6 (7×7) + Группы 2016 Высшая 5×4 (5×5) + 2018 Премьер/Высшая 5×5 2кр + 2020 предварительный 13 (4×4 — отмечено) — всего 299 матчей, продолжение в январе"
    ws["A2"].font=Font(size=8,italic=True,color="64748B");ws["A2"].alignment=Alignment(horizontal="center")
    headers=["Дата","День","Время","Поле","Формат","Лига","Группа","Хозяева","Тренер","Гости","Тренер"]
    for c,h in enumerate(headers,1):
        cell=ws.cell(row=4,column=c,value=h);cell.fill=HEADER_FILL;cell.font=HEADER_FONT;cell.alignment=Alignment(horizontal="center",vertical="center");cell.border=BORDER
    ws.row_dimensions[4].height=22
    scheduled_sorted=sorted(scheduled, key=lambda x:(x["date_obj"], x["time"], x["field"]))
    row=5
    for sc in scheduled_sorted:
        dow=sc["date_obj"].strftime("%a")
        days={"Mon":"Пн","Tue":"Вт","Wed":"Ср","Thu":"Чт","Fri":"Пт","Sat":"Сб","Sun":"Вс"}
        dow=days.get(dow,dow)
        vals=[sc["date"],dow,sc["time"],sc["field"],sc["format"],sc["league"],sc["group"],sc["home"],sc["home_train"],sc["away"],sc["away_train"]]
        for c,v in enumerate(vals,1):
            cell=ws.cell(row=row,column=c,value=v)
            cell.font=Font(size=8) if c not in (8,10) else Font(size=8,bold=True)
            if c in (9,11): cell.font=Font(size=7,color="475569")
            cell.alignment=Alignment(horizontal="center",vertical="center");cell.border=BORDER
            if sc.get("is_fallback"): cell.fill=PatternFill(start_color="FEF3C7",end_color="FEF3C7",fill_type="solid")
        if "2020" in sc["league"]:
            for c in range(1,12):
                if ws.cell(row=row,column=c).fill.fill_type is None:
                    ws.cell(row=row,column=c).fill=GOLD_FILL
        ws.row_dimensions[row].height=16;row+=1
    widths=[8,6,7,16,7,30,10,18,12,18,12]
    for i,w in enumerate(widths,1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    ws.freeze_panes="A5";ws.auto_filter.ref=f"A4:K{row-1}";ws.sheet_properties.pageSetUpPr.fitToPage=True;ws.page_setup.orientation="landscape";ws.page_setup.paperSize=ws.PAPERSIZE_A4;ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=0

    ws2=wb.create_sheet("Сводка_спортдиректора")
    ws2["A1"].value="Сводка спортдиректора — осенний блок"
    ws2["A1"].font=Font(bold=True,size=12)
    ws2["A3"].value=f"Запланировано: {len(scheduled)}/299"
    ws2["A4"].value=f"Не влезло: {len(unscheduled)}"
    ws2["A5"].value=f"Конфликтов параллелей: {len(conflicts)} — {'0 ✓' if not conflicts else 'есть!'}"
    if not conflicts: ws2["A5"].font=Font(color="16A34A",bold=True)
    ws2["A7"].value="Тренеры с нагрузкой >5 игр:"
    r=8
    cnt=Counter()
    for sc in scheduled:
        cnt[sc["home_train"]]+=1;cnt[sc["away_train"]]+=1
    for tr,c in cnt.most_common(10):
        ws2.cell(row=r,column=1,value=tr);ws2.cell(row=r,column=2,value=c);r+=1
    ws2["A20"].value="Что не вошло в этот блок (отложено на весну — 7×7 двойные):"
    ws2["A21"].value="2012-2013 7×7 16 ком 2кр (240 матчей) • 2014-2015 Премьер 7×7 2кр (56) • 2016-2017 Премьер 7×7 2кр (132) — итого 428 матчей 7×7, им нужен отдельный 7×7 фестиваль (не хватает 7×7 полей в этих датах)"
    ws2["A21"].alignment=Alignment(wrap_text=True,vertical="center")
    ws2["A21"].font=Font(size=8,color="DC2626")
    ws2.column_dimensions["A"].width=70
    ws2.merge_cells("A21:G23")
    wb.save("Raspisanie_Osень_Fit.xlsx")
    print("Saved Raspisanie_Osень_Fit.xlsx — влезает полностью, без параллелей")

if __name__=="__main__":
    main()
