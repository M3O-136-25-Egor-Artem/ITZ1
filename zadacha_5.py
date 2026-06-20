# Задача 5. Если в прошлом году предмет вёл ОДИН преподаватель,
# то копируем его в строку текущего года (и все часы остаются за ним).

import os
import re
import xlrd
import openpyxl
from typing import Any

# пути к файлам, если что - меняем тут
papka = r"C:\projects\ИТЗ_2сем\входные файлы"
file_proshlyy = "15 вариант.xls"             # сводный прошлого года (тут фио есть)
file_osen = "Список 15вар - осень.xls"       # текущий год осень (фио пустые)
file_vesna = "Список 15вар - весна.xls"      # текущий год весна (фио пустые)


def read_xls(path: str, name: str) -> list[list[Any]]:
     # читаем лист из старого экселя в обычный список строк
    kniga = xlrd.open_workbook(path)
    sheet = kniga.sheet_by_name(name)
    return [[sheet.cell_value(r, c) for c in range(sheet.ncols)]
            for r in range(sheet.nrows)]


def chistim(x: Any) -> str:
    # приводим ячейку к нормальной строке: убираем лишние пробелы,
    # и число 18.0 делаем "18" а не "18.0"
    if x is None:
        return ""
    if isinstance(x, float) and x == int(x):
        x = int(x)
    return re.sub(r"\s+", " ", str(x)).strip()


def eto_prepod(fio: str) -> bool:
     # проверяем что это настоящее фио, а не мусор в колонке
     # (комиссии "э-...", коды групп с цифрами, пустые, всякие пометки)
    if not fio or fio[0].islower():
        return False
    if fio.startswith("э"):
        return False
    if any(c.isdigit() for c in fio):
        return False
    return True


# 1) собираем по прошлому году кто какой предмет вёл
# ключ = (группа, дисциплина, вид нагрузки) - только эти три
proshlyy = read_xls(os.path.join(papka, file_proshlyy), "общее")
kto_vel = {}
for stroka in proshlyy[1:]:        # первую строку (шапку) пропускаем
    if len(stroka) <= 8:
        continue
    fio = chistim(stroka[8])       # в этом файле фио в 8й колонке
    if not eto_prepod(fio):
        continue
    kluch = (chistim(stroka[3]), chistim(stroka[1]), chistim(stroka[2]))
    if kluch not in kto_vel:
        kto_vel[kluch] = set()     # set чтобы один и тот же препод не дублировался
    kto_vel[kluch].add(fio)


# 2) читаем текущий год (осень + весна) - сюда будем вписывать фио
tekushie = []
for semestr, f in [("осень", file_osen), ("весна", file_vesna)]:
    dannye = read_xls(os.path.join(papka, f), "TDSheet")
    for stroka in dannye[1:]:
        if len(stroka) <= 7:
            continue
        disc = chistim(stroka[0])
        vid = chistim(stroka[1])
        if disc == "" and vid == "":   # пустые строки и строка с итогом - пропускаем
            continue
        tekushie.append({
            "semestr": semestr,
            "disc": disc,
            "vid": vid,
            "gruppa": chistim(stroka[2]),
            "chasy": stroka[3],            # часы берём как есть и не меняем
            "fio_bylo": chistim(stroka[7]), # что уже стояло - пригодится для проверки
            "fio": "",
            "coment": "",
        })


# 3) сама задача 5
sdelano = 0
for s in tekushie:
    kluch = (s["gruppa"], s["disc"], s["vid"])
    prepody = kto_vel.get(kluch)     # кто вёл этот предмет в прошлом году
    if not prepody:
        s["coment"] = "в прошлом году не нашли"
        continue

    if len(prepody) == 1:
        # препод один - просто ставим его в текущую строку, все часы его
        s["fio"] = list(prepody)[0]
        s["coment"] = "задача 5: один препод, скопировали"
        sdelano += 1
    else:
        # двое и больше - это уже задачи 6 и 7, тут не трогаем
        s["coment"] = "не задача 5, преподов " + str(len(prepody))


# 4) проверим себя: сколько совпало с тем фио, что уже было в списках
sovpalo = 0
proverili = 0
for s in tekushie:
    if s["fio"] != "" and s["fio_bylo"] != "":
        proverili += 1
        if s["fio"] == s["fio_bylo"]:
            sovpalo += 1

print("всего строк:", len(tekushie))
print("заполнили по задаче 5:", sdelano)
if proverili:
    print("совпало с тем что было:", sovpalo, "из", proverili)


# 5) сохраняем результат в эксель, чтобы можно было посмотреть глазами
kniga = openpyxl.Workbook()
sheet = kniga.active
sheet.append(["Семестр", "Дисциплина", "Вид", "Группа", "Часы", "ФИО", "Комментарий"])

for s in tekushie:
    sheet.append([s["semestr"], s["disc"], s["vid"], s["gruppa"],
                  s["chasy"], s["fio"], s["coment"]])

kniga.save("Результат_Задача5.xlsx")
print("готово, сохранили в Результат_Задача5.xlsx")