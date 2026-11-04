#!/usr/bin/env python3
"""BPMN 2.0 по проекту регламента запросов на разработку от 12.10.2026."""

from pathlib import Path
from xml.sax.saxutils import escape

OUT = Path(__file__).resolve().parent / "zapros-na-razrabotku.bpmn"

# top-left x, y, w, h. События 36, шлюзы 50, задачи шире.
NODES = {
    "Start_need": ("Потребность в разработке", "start", 120, 150, 36, 36),
    "Start_b2b": ("Из процесса B2B", "message", 120, 250, 36, 36),
    "Gw_incident": ("Инцидент в продуктиве?", "xor", 230, 168, 50, 50),
    "End_support": ("В поддержку", "end", 237, 280, 36, 36),
    "Task_card": ("Карточка в Simbase", "task", 340, 155, 180, 76),
    "Task_fill": ("Дозаполнить поля", "task", 620, 155, 170, 76),
    "Timer_fill": ("5 рабочих дней", "timer", 754, 213, 36, 36),
    "End_autoclose": ("Автозакрытие", "end", 850, 175, 36, 36),
    "Task_accept": ("Принять результат", "task", 3600, 155, 180, 76),
    "Gw_accept": ("Принято?", "xor", 3840, 168, 50, 50),
    "Task_drb": ("Справка ДРБ до 15-го числа", "task", 1180, 400, 220, 80),
    "Task_qualify": ("Квалификация запроса", "task", 600, 620, 200, 80),
    "Gw_qual": ("Итог квалификации", "xor", 860, 635, 50, 50),
    "Task_pool": ("Квартальный пул", "task", 980, 620, 170, 76),
    "Gw_split": ("", "and", 1480, 633, 50, 50),
    "Task_value": ("Ценность, экономика и балл", "task", 1600, 620, 210, 76),
    "Gw_join": ("", "and", 1880, 633, 50, 50),
    "Task_reco": ("Материалы для CPO", "task", 2000, 620, 180, 76),
    "Gw_cpo": ("Решение CPO", "xor", 2240, 633, 50, 50),
    "Gw_level": ("Уровень утверждения", "xor", 2380, 633, 50, 50),
    "Task_cpo": ("Утвердить в пороге CPO", "task", 2520, 620, 200, 76),
    "Task_jira": ("Эпик Jira с номером запроса", "task", 2920, 620, 220, 76),
    "Timer_post": ("3–6 мес. после релиза", "timer", 3920, 640, 36, 36),
    "Task_post": ("Пост-анализ", "task", 4040, 620, 160, 76),
    "End_closed": ("Закрыт", "end", 4260, 640, 36, 36),
    "Gw_typec": ("Тип C?", "xor", 900, 820, 50, 50),
    "Task_fast_b": ("Оценить сделку, 5 р. дней", "task", 1040, 760, 210, 72),
    "Task_fast_c": ("Согласовать срок обязательства", "task", 1040, 870, 230, 72),
    "Timer_postpone": ("До даты пересмотра", "timer", 2247, 820, 36, 36),
    "End_reject_qual": ("Отклонён", "end", 867, 900, 36, 36),
    "End_reject_cpo": ("Отклонён", "end", 2247, 900, 36, 36),
    "Task_cost": ("Оценить стоимость и срок", "task", 1600, 1080, 210, 76),
    "Task_build": ("Разработка, статус раз в месяц", "task", 2920, 1080, 240, 76),
    "Task_fin": ("Проверить финансирование", "task", 1600, 1260, 220, 76),
    "Task_board": ("Продуктовый совет", "task", 2520, 1440, 190, 76),
    "Gw_board": ("", "xor", 2780, 1453, 50, 50),
    "End_reject_board": ("Отклонён", "end", 2787, 1540, 36, 36),
    "Task_ik": ("Инвестиционный комитет", "task", 2520, 1640, 220, 76),
    "Gw_ik": ("", "xor", 2800, 1653, 50, 50),
    "End_reject_ik": ("Отклонён", "end", 2807, 1740, 36, 36),
    "Task_order": ("Приказ CEO о запуске", "task", 2920, 1820, 200, 76),
}

LANES = [
    ("Lane_init", "Инициатор", 320, [
        "Start_need", "Start_b2b", "Gw_incident", "End_support",
        "Task_card", "Task_fill", "Timer_fill", "End_autoclose",
        "Task_accept", "Gw_accept",
    ]),
    ("Lane_drb", "ДРБ", 180, ["Task_drb"]),
    ("Lane_po", "Product Office", 520, [
        "Task_qualify", "Gw_qual", "End_reject_qual", "Task_pool",
        "Gw_typec", "Task_fast_b", "Task_fast_c", "Gw_split", "Task_value",
        "Gw_join", "Task_reco", "Gw_cpo", "Timer_postpone", "End_reject_cpo",
        "Gw_level", "Task_cpo", "Task_jira", "Timer_post", "Task_post", "End_closed",
    ]),
    ("Lane_st", "Sergek Technologies", 180, ["Task_cost", "Task_build"]),
    ("Lane_fin", "Финансовый департамент", 160, ["Task_fin"]),
    ("Lane_board", "Продуктовый совет", 180, ["Task_board", "Gw_board", "End_reject_board"]),
    ("Lane_ik", "Инвестиционный комитет", 180, ["Task_ik", "Gw_ik", "End_reject_ik"]),
    ("Lane_ceo", "CEO", 160, ["Task_order"]),
]

# (source, target, name or "")
FLOWS = [
    ("Start_need", "Gw_incident", ""),
    ("Start_b2b", "Task_card", ""),
    ("Gw_incident", "End_support", "да"),
    ("Gw_incident", "Task_card", "нет"),
    ("Task_card", "Task_qualify", ""),
    ("Task_qualify", "Gw_qual", ""),
    ("Gw_qual", "Task_fill", "не хватает полей"),
    ("Task_fill", "Task_qualify", "дозаполнено"),
    ("Timer_fill", "End_autoclose", "срок вышел"),
    ("Gw_qual", "End_reject_qual", "дубль или стоп-фактор"),
    ("Gw_qual", "Task_pool", "A, D, E"),
    ("Gw_qual", "Gw_typec", "B или C"),
    ("Task_pool", "Task_drb", ""),
    ("Task_drb", "Gw_split", ""),
    ("Gw_split", "Task_value", ""),
    ("Gw_split", "Task_cost", ""),
    ("Gw_split", "Task_fin", ""),
    ("Task_value", "Gw_join", ""),
    ("Task_cost", "Gw_join", ""),
    ("Task_fin", "Gw_join", ""),
    ("Gw_join", "Task_reco", ""),
    ("Gw_typec", "Task_fast_c", "обязательство"),
    ("Gw_typec", "Task_fast_b", "сделка"),
    ("Task_fast_c", "Task_reco", ""),
    ("Task_fast_b", "Task_reco", ""),
    ("Task_reco", "Gw_cpo", ""),
    ("Gw_cpo", "Task_fill", "нужны данные"),
    ("Gw_cpo", "End_reject_cpo", "отклонить"),
    ("Gw_cpo", "Timer_postpone", "отложить"),
    ("Timer_postpone", "Task_pool", ""),
    ("Gw_cpo", "Gw_level", "на утверждение"),
    ("Gw_level", "Task_cpo", "порог CPO"),
    ("Gw_level", "Task_board", "Продсовет"),
    ("Gw_level", "Task_ik", "ИК"),
    ("Task_cpo", "Task_jira", ""),
    ("Task_board", "Gw_board", ""),
    ("Gw_board", "Task_jira", "утверждён"),
    ("Gw_board", "End_reject_board", "отклонён"),
    ("Task_ik", "Gw_ik", ""),
    ("Gw_ik", "Task_order", "утверждён"),
    ("Gw_ik", "End_reject_ik", "отклонён"),
    ("Task_order", "Task_jira", ""),
    ("Task_jira", "Task_build", ""),
    ("Task_build", "Task_accept", ""),
    ("Task_accept", "Gw_accept", ""),
    ("Gw_accept", "Task_build", "замечания"),
    ("Gw_accept", "Timer_post", "принято"),
    ("Timer_post", "Task_post", ""),
    ("Task_post", "End_closed", ""),
]

DOCS = {
    "Task_card": "Статус «Черновик», затем «Зарегистрирован». Письмо, мессенджер и устное поручение не заменяют карточку. Обязательные поля блоков 0–3: что, зачем, как, источник финансирования, владелец со стороны бизнеса.",
    "Task_qualify": "Intake-координатор Product Office, 2 рабочих дня. Проверяет полноту и дубли, присваивает тип A–E и трек, назначает PM. Статус «Зарегистрирован».",
    "Task_fill": "Статус «Нужны доп. данные». Инициатор дополняет карточку. Если за 5 рабочих дней поля не закрыты, карточка закрывается автоматически.",
    "Task_pool": "Типы A, D и E ждут квартальный цикл. Резерв 5–10% ёмкости держится под ускоренный трек.",
    "Task_drb": "До 15-го числа последнего месяца квартала: заказчики, регионы, бюджет, приоритет ДРБ, ссылки на карточки Simbase. Ответ Product Office по каждому пункту — не позже 2 недель.",
    "Task_value": "PM. Сначала стоп-факторы, затем балл 0–3 по шести критериям. P1 от 2,2, P2 1,5–2,2, P3 ниже 1,5 уходит в бэклог с датой пересмотра.",
    "Task_cost": "Sergek Technologies, R/A оценки. Полная стоимость: ФОТ × срок, оборудование, подрядчики, лицензии, сертификация. Срок анализа квартального трека — 10 рабочих дней от справки.",
    "Task_fin": "Финансовый департамент проверяет источник финансирования и экономику на горизонте 3 лет.",
    "Task_fast_b": "Тип B, ускоренный трек, 5 рабочих дней. PM и Sergek Technologies. Подтверждение: КП, ТС, тендер или гарантийное письмо.",
    "Task_fast_c": "Тип C. Решается срок исполнения, а не целесообразность. Подтверждение: договор, НПА или предписание. Трек ускоренный, утверждение всё равно по порогу.",
    "Task_cpo": "В проекте регламента рекомендован вариант C: до 15 млн тг и до 3 месяцев. Суммы не утверждены.",
    "Task_board": "Промежуточный уровень между CPO и ИК. В проекте — 15–75 млн тг или срок до 6 месяцев. Создание совета ещё не решено.",
    "Task_ik": "Выше порога и всегда для особых категорий: новый продукт, hardware/АПК/Edge, инвестиционная разработка без заказчика больше 20 млн тг, срок больше 6 месяцев. Для ускоренного трека — внеочередное или заочное заседание.",
    "Task_order": "Для уровня ИК: согласование Project office, ДУП и бухгалтерии, подпись CEO. Как в процессе продаж B2B.",
    "Task_jira": "PM, 2 рабочих дня. Номер запроса Simbase обязателен в эпике. Задачу без номера команда не берёт.",
    "Task_build": "Разработка силами Sergek Technologies. PM обновляет статус в Simbase раз в месяц. Перерасход больше 20% или сдвиг срока больше 30% возвращает запрос на тот же уровень утверждения.",
    "Task_accept": "Инициатор, 5 рабочих дней, статус «Выполнен».",
    "Task_post": "PM через 3–6 месяцев после релиза: факт выручки и использования против плана. Затем статус «Закрыт».",
    "Gw_level": "Порог в схеме показан по рекомендованному варианту C. Вариант ещё не выбран.",
}


def lane_bounds():
    y = 56
    bounds = {}
    for lid, name, height, _ in LANES:
        bounds[lid] = (70, y, 4360, height)
        y += height
    return bounds, y - 56


def center(nid):
    _, _, x, y, w, h = NODES[nid]
    return x + w / 2, y + h / 2


def edge_points(src, tgt):
    sx, sy = center(src)
    tx, ty = center(tgt)
    # Возвраты идут по внешней стороне, чтобы не резать основной ряд.
    if src == "Start_b2b" and tgt == "Task_card":
        return [(sx, sy), (sx, 330), (520, 330), (520, ty), (tx, ty)]
    if tgt == "Task_qualify" and src == "Task_fill":
        return [(sx, sy), (sx, 80), (700, 80), (700, ty)]
    if tgt == "Task_fill" and src == "Gw_cpo":
        return [(sx, sy), (sx, 100), (tx, 100), (tx, ty)]
    if tgt == "Task_pool" and src == "Timer_postpone":
        return [(sx, sy), (sx, 980), (1065, 980), (1065, ty)]
    if tgt == "Task_build" and src == "Gw_accept":
        return [(sx, sy), (sx, 120), (3040, 120), (3040, ty)]
    if abs(sy - ty) < 30:
        return [(sx, sy), (tx, ty)]
    mid = (sx + tx) / 2
    return [(sx, sy), (mid, sy), (mid, ty), (tx, ty)]


def shape_xml(nid, kind):
    name, _, x, y, w, h = NODES[nid]
    label = ""
    if kind in {"xor", "and", "start", "message", "end", "timer"} and name:
        label = f"""
        <bpmndi:BPMNLabel>
          <dc:Bounds x="{x - 40}" y="{y + h + 4}" width="{w + 80}" height="32"/>
        </bpmndi:BPMNLabel>"""
    return f"""      <bpmndi:BPMNShape id="Shape_{nid}" bpmnElement="{nid}">
        <dc:Bounds x="{x}" y="{y}" width="{w}" height="{h}"/>{label}
      </bpmndi:BPMNShape>"""


def node_xml(nid):
    name, kind, *_ = NODES[nid]
    label = escape(name)
    doc = DOCS.get(nid)
    doc_xml = ""
    if doc:
        doc_xml = f"<bpmn:documentation>{escape(doc)}</bpmn:documentation>"
    if kind == "start":
        return f'<bpmn:startEvent id="{nid}" name="{label}"/>'
    if kind == "message":
        return (
            f'<bpmn:startEvent id="{nid}" name="{label}">'
            f'<bpmn:messageEventDefinition id="Msg_{nid}"/>'
            f"</bpmn:startEvent>"
        )
    if kind == "end":
        return f'<bpmn:endEvent id="{nid}" name="{label}"/>'
    if kind == "timer" and nid == "Timer_fill":
        return (
            f'<bpmn:boundaryEvent id="{nid}" name="{label}" attachedToRef="Task_fill" cancelActivity="true">'
            f'<bpmn:timerEventDefinition id="Def_{nid}"/>'
            f"</bpmn:boundaryEvent>"
        )
    if kind == "timer":
        return (
            f'<bpmn:intermediateCatchEvent id="{nid}" name="{label}">'
            f'<bpmn:timerEventDefinition id="Def_{nid}"/>'
            f"</bpmn:intermediateCatchEvent>"
        )
    if kind == "xor":
        return f'<bpmn:exclusiveGateway id="{nid}" name="{label}"/>'
    if kind == "and":
        return f'<bpmn:parallelGateway id="{nid}" name="{label}"/>'
    return f'<bpmn:userTask id="{nid}" name="{label}">{doc_xml}</bpmn:userTask>'


def build() -> str:
    lane_xml = []
    for lid, name, _, refs in LANES:
        refs_xml = "".join(f"<bpmn:flowNodeRef>{n}</bpmn:flowNodeRef>" for n in refs)
        lane_xml.append(f'<bpmn:lane id="{lid}" name="{escape(name)}">{refs_xml}</bpmn:lane>')
    nodes = "\n    ".join(node_xml(n) for n in NODES)
    flows = []
    edges = []
    for i, (src, tgt, label) in enumerate(FLOWS, start=1):
        fid = f"Flow_{i}"
        name = f' name="{escape(label)}"' if label else ""
        flows.append(f'<bpmn:sequenceFlow id="{fid}" sourceRef="{src}" targetRef="{tgt}"{name}/>')
        points = edge_points(src, tgt)
        wps = "\n        ".join(
            f'<di:waypoint x="{int(x)}" y="{int(y)}"/>' for x, y in points
        )
        edges.append(
            f'      <bpmndi:BPMNEdge id="Edge_{fid}" bpmnElement="{fid}">\n        {wps}\n      </bpmndi:BPMNEdge>'
        )
    bounds, height = lane_bounds()
    lane_shapes = []
    for lid, name, h, _ in LANES:
        x, y, w, hh = bounds[lid]
        lane_shapes.append(
            f"""      <bpmndi:BPMNShape id="Shape_{lid}" bpmnElement="{lid}" isHorizontal="true">
        <dc:Bounds x="{x}" y="{y}" width="{w}" height="{hh}"/>
      </bpmndi:BPMNShape>"""
        )
    shapes = "\n".join(shape_xml(n, NODES[n][1]) for n in NODES)
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<bpmn:definitions xmlns:bpmn="http://www.omg.org/spec/BPMN/20100524/MODEL" xmlns:bpmndi="http://www.omg.org/spec/BPMN/20100524/DI" xmlns:dc="http://www.omg.org/spec/DD/20100524/DC" xmlns:di="http://www.omg.org/spec/DD/20100524/DI" id="Definitions_DevRequest" targetNamespace="http://sergek.group/product-office" exporter="Product Office">
  <bpmn:collaboration id="Collaboration_DevRequest">
    <bpmn:participant id="Participant_DevRequest" name="Запрос на разработку продукта" processRef="Process_DevRequest"/>
  </bpmn:collaboration>
  <bpmn:process id="Process_DevRequest" name="Запрос на разработку" isExecutable="false">
    <bpmn:laneSet id="LaneSet_DevRequest">
      {''.join(lane_xml)}
    </bpmn:laneSet>
    {nodes}
    {''.join(flows)}
    <bpmn:textAnnotation id="Note_stop">
      <bpmn:text>Стоп-факторы: продукт в категории «Закрытие»; нет источника финансирования; нет владельца со стороны бизнеса; функция уже есть или закрывается конфигурацией, партнёром, интеграцией.</bpmn:text>
    </bpmn:textAnnotation>
    <bpmn:association id="Assoc_stop" sourceRef="Task_qualify" targetRef="Note_stop"/>
    <bpmn:textAnnotation id="Note_threshold">
      <bpmn:text>Пороги на схеме — рекомендованный вариант C, он ещё не утверждён: CPO до 15 млн тг и 3 мес.; Продсовет 15–75 млн или до 6 мес.; ИК выше 75 млн и особые категории.</bpmn:text>
    </bpmn:textAnnotation>
    <bpmn:association id="Assoc_threshold" sourceRef="Gw_level" targetRef="Note_threshold"/>
  </bpmn:process>
  <bpmndi:BPMNDiagram id="Diagram_DevRequest">
    <bpmndi:BPMNPlane id="Plane_DevRequest" bpmnElement="Collaboration_DevRequest">
      <bpmndi:BPMNShape id="Shape_Participant" bpmnElement="Participant_DevRequest" isHorizontal="true">
        <dc:Bounds x="40" y="56" width="4390" height="{height}"/>
      </bpmndi:BPMNShape>
      {''.join(lane_shapes)}
      {shapes}
      <bpmndi:BPMNShape id="Shape_Note_stop" bpmnElement="Note_stop">
        <dc:Bounds x="1280" y="960" width="300" height="90"/>
      </bpmndi:BPMNShape>
      <bpmndi:BPMNShape id="Shape_Note_threshold" bpmnElement="Note_threshold">
        <dc:Bounds x="2520" y="820" width="300" height="100"/>
      </bpmndi:BPMNShape>
      <bpmndi:BPMNEdge id="Edge_Assoc_stop" bpmnElement="Assoc_stop">
        <di:waypoint x="760" y="700"/>
        <di:waypoint x="1280" y="1000"/>
      </bpmndi:BPMNEdge>
      <bpmndi:BPMNEdge id="Edge_Assoc_threshold" bpmnElement="Assoc_threshold">
        <di:waypoint x="2405" y="683"/>
        <di:waypoint x="2600" y="820"/>
      </bpmndi:BPMNEdge>
      {''.join(edges)}
    </bpmndi:BPMNPlane>
  </bpmndi:BPMNDiagram>
</bpmn:definitions>
"""
    return xml


def main() -> None:
    OUT.write_text(build(), encoding="utf-8")
    print(OUT, OUT.stat().st_size)


if __name__ == "__main__":
    main()
