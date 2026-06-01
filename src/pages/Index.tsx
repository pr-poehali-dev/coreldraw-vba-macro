import { useState } from "react";
import Icon from "@/components/ui/icon";

// ============================================================
// МАКРОС 1 — Контур вокруг группы
// ============================================================
const vbaCodeOutline = `' ============================================
' Макрос CorelDRAW: Контур вокруг группы
' Версия: 1.1 — ввод отступа через диалог
' ============================================

' Добавьте этот тип в НАЧАЛО модуля (до всех Sub/Function):
' Type BoundingBox
'     x As Double
'     y As Double
'     w As Double
'     h As Double
' End Type


' --- ТОЧКА ВХОДА ---
' Назначьте эту процедуру на контекстное меню CorelDRAW
Sub DrawGroupOutline()
    If Application.Documents.Count = 0 Then
        MsgBox "Нет открытого документа.", vbExclamation, "Ошибка"
        Exit Sub
    End If

    ' Application.ActiveSelection — объект Shape всего выделения (CorelDRAW 26)
    Dim sel As Shape
    Set sel = Application.ActiveSelection

    If sel Is Nothing Then
        MsgBox "Выберите объект или группу.", vbExclamation, "Ошибка"
        Exit Sub
    End If

    If sel.SizeWidth = 0 And sel.SizeHeight = 0 Then
        MsgBox "Выберите объект или группу.", vbExclamation, "Ошибка"
        Exit Sub
    End If

    ' Запрашиваем отступ у пользователя (по умолчанию 20 мм)
    Dim inputVal As String
    inputVal = InputBox("Введите суммарный отступ контура в мм:" & vbCrLf & _
                        "(контур станет шире и выше на это значение)", _
                        "Отступ контура", "20")

    If inputVal = "" Then Exit Sub

    If Not IsNumeric(inputVal) Then
        MsgBox "Введите корректное числовое значение.", vbExclamation, "Ошибка"
        Exit Sub
    End If

    Dim totalOffset As Double
    totalOffset = CDbl(inputVal)

    If totalOffset <= 0 Then
        MsgBox "Отступ должен быть больше нуля.", vbExclamation, "Ошибка"
        Exit Sub
    End If

    Dim halfOffset As Double
    halfOffset = totalOffset / 2

    ' Координаты bbox выделения
    Dim bX As Double: bX = sel.LeftX
    Dim bY As Double: bY = sel.BottomY
    Dim bW As Double: bW = sel.SizeWidth
    Dim bH As Double: bH = sel.SizeHeight

    Dim doc As Document
    Set doc = Application.ActiveDocument

    ' Создаём прямоугольник-контур
    Dim rW As Double: rW = bW + halfOffset * 2
    Dim rH As Double: rH = bH + halfOffset * 2
    Dim rect As Shape
    Set rect = doc.ActiveLayer.CreateRectangle2(0, 0, rW, rH)

    rect.Fill.ApplyNoFill

    Dim outlineColor As New Color
    outlineColor.CMYKAssign 0, 0, 0, 100

    rect.Outline.SetProperties _
        Width:=0.2, _
        style:=1, _
        color:=outlineColor, _
        arrow1:=Application.ArrowHeads(1), _
        arrow2:=Application.ArrowHeads(1), _
        linecaps:=cdrButtLineCaps, _
        linejoins:=cdrMiterLineJoins, _
        nib_angle:=0, _
        stretch:=100, _
        behind_fill:=False, _
        scale_on_resize:=False

    ' Позиционируем по центру выделения
    rect.SetPosition bX - halfOffset, bY - halfOffset

    MsgBox "Контур создан! Отступ: " & totalOffset & " мм (" & halfOffset & " мм с каждой стороны).", _
           vbInformation, "Готово"
End Sub


' --- ФУНКЦИЯ 1: Расчёт границ группы ---
Function GetGroupBounds(grp As Shape) As BoundingBox
    Dim bb As BoundingBox
    bb.x = grp.LeftX
    bb.y = grp.BottomY
    bb.w = grp.SizeWidth
    bb.h = grp.SizeHeight
    GetGroupBounds = bb
End Function


' --- ФУНКЦИЯ 2: Создание прямоугольника-контура ---
' halfOffset — отступ с каждой стороны в мм
Function CreateOutlineRect(doc As Document, bounds As BoundingBox, halfOffset As Double) As Shape
    Const STROKE_MM As Double = 0.2

    Dim rW As Double: rW = bounds.w + halfOffset * 2
    Dim rH As Double: rH = bounds.h + halfOffset * 2

    Dim rect As Shape
    Set rect = doc.ActiveLayer.CreateRectangle2(0, 0, rW, rH)

    rect.Fill.ApplyNoFill

    Dim outlineColor As New Color
    outlineColor.CMYKAssign 0, 0, 0, 100

    rect.Outline.SetProperties _
        Width:=STROKE_MM, _
        style:=1, _
        color:=outlineColor, _
        arrow1:=Application.ArrowHeads(1), _
        arrow2:=Application.ArrowHeads(1), _
        linecaps:=cdrButtLineCaps, _
        linejoins:=cdrMiterLineJoins, _
        nib_angle:=0, _
        stretch:=100, _
        behind_fill:=False, _
        scale_on_resize:=False

    Set CreateOutlineRect = rect
End Function


' --- ФУНКЦИЯ 3: Позиционирование контура ---
' halfOffset — отступ с каждой стороны в мм
Sub PositionOutline(outline As Shape, bounds As BoundingBox, halfOffset As Double)
    Dim cx As Double: cx = bounds.x + bounds.w / 2
    Dim cy As Double: cy = bounds.y + bounds.h / 2

    Dim newX As Double: newX = cx - (bounds.w / 2 + halfOffset)
    Dim newY As Double: newY = cy - (bounds.h / 2 + halfOffset)

    outline.SetPosition newX, newY
End Sub`;

// ============================================================
// МАКРОС 2 — Угловые точки (метки реза)
// ============================================================
const vbaCodeDots = `' ============================================
' Макрос CorelDRAW: Угловые точки (метки реза)
' CorelDRAW 26.1 — VBA
' Версия: 1.0
' ============================================
' Рисует 4 квадратные точки 10×10 мм по углам выделения.
' Отступ от края выделения — 10 мм наружу.
' Заливка: CMYK 0/0/0/100. Без абриса.
' Работает с одним объектом, несколькими или группой.
' ============================================


Sub DrawCornerDots()
    If Application.Documents.Count = 0 Then
        MsgBox "Нет открытого документа.", vbExclamation, "Ошибка"
        Exit Sub
    End If

    ' Application.ActiveSelection — суммарный bbox всего выделения (CorelDRAW 26)
    Dim sel As Shape
    Set sel = Application.ActiveSelection

    If sel Is Nothing Then
        MsgBox "Выберите объект, объекты или группу.", vbExclamation, "Ошибка"
        Exit Sub
    End If

    If sel.SizeWidth = 0 And sel.SizeHeight = 0 Then
        MsgBox "Выберите объект, объекты или группу.", vbExclamation, "Ошибка"
        Exit Sub
    End If

    ' Суммарный bounding box выделения
    Dim minX As Double: minX = sel.LeftX
    Dim minY As Double: minY = sel.BottomY
    Dim maxX As Double: maxX = sel.LeftX + sel.SizeWidth
    Dim maxY As Double: maxY = sel.BottomY + sel.SizeHeight

    ' --- Параметры точек ---
    Const DOT_SIZE As Double = 10   ' 10 мм — размер точки
    Const OFFSET   As Double = 10   ' 10 мм — отступ от края выделения

    ' Координаты четырёх углов (левый нижний угол каждого квадрата):
    '   Левый нижний   — (minX - OFFSET - DOT_SIZE,  minY - OFFSET - DOT_SIZE)
    '   Правый нижний  — (maxX + OFFSET,              minY - OFFSET - DOT_SIZE)
    '   Левый верхний  — (minX - OFFSET - DOT_SIZE,  maxY + OFFSET)
    '   Правый верхний — (maxX + OFFSET,              maxY + OFFSET)

    Dim corners(3, 1) As Double
    corners(0, 0) = minX - OFFSET - DOT_SIZE : corners(0, 1) = minY - OFFSET - DOT_SIZE
    corners(1, 0) = maxX + OFFSET             : corners(1, 1) = minY - OFFSET - DOT_SIZE
    corners(2, 0) = minX - OFFSET - DOT_SIZE : corners(2, 1) = maxY + OFFSET
    corners(3, 0) = maxX + OFFSET             : corners(3, 1) = maxY + OFFSET

    ' --- Цвет заливки: CMYK 0/0/0/100 ---
    Dim fillColor As New Color
    fillColor.CMYKAssign 0, 0, 0, 100

    ' --- Рисуем 4 точки ---
    Dim dot As Shape
    Dim k As Integer
    For k = 0 To 3
        Set dot = Application.ActiveDocument.ActiveLayer.CreateRectangle2( _
            corners(k, 0), corners(k, 1), DOT_SIZE, DOT_SIZE)

        ' Заливка — чёрная CMYK
        dot.Fill.ApplyUniform fillColor

        ' Без абриса
        dot.Outline.Remove
    Next k

    MsgBox "Готово! 4 угловые точки 10×10 мм созданы с отступом 10 мм.", _
           vbInformation, "Угловые точки"
End Sub`;

// ============================================================
// Данные для вкладок
// ============================================================

const stepsOutline = [
  { num: "01", title: "Открыть редактор VBA", desc: "Сервис → Макросы → Редактор Visual Basic (или Alt+F11)" },
  { num: "02", title: "Создать модуль", desc: "В дереве проекта: ПКМ на проекте → Вставить → Модуль. Назвать GroupOutline." },
  { num: "03", title: "Добавить тип BoundingBox", desc: "Скопировать блок Type BoundingBox (в комментарии вверху) в начало модуля и раскомментировать." },
  { num: "04", title: "Вставить код макроса", desc: "Скопировать весь остальной код в тот же модуль ниже объявления типа." },
  { num: "05", title: "Подключить к меню", desc: "Сервис → Настройка → Команды → найти DrawGroupOutline → перетащить в нужное меню." },
];

const stepsDots = [
  { num: "01", title: "Открыть редактор VBA", desc: "Сервис → Макросы → Редактор Visual Basic (или Alt+F11)" },
  { num: "02", title: "Создать модуль", desc: "ПКМ на проекте → Вставить → Модуль. Назвать CornerDots." },
  { num: "03", title: "Вставить код макроса", desc: "Скопировать весь код в модуль CornerDots." },
  { num: "04", title: "Запустить", desc: "Выделить объект(ы) или группу в CorelDRAW, затем Сервис → Макросы → DrawCornerDots → Выполнить." },
  { num: "05", title: "Подключить к меню", desc: "Сервис → Настройка → Команды → найти DrawCornerDots → перетащить в нужное меню или панель." },
];

const paramsOutline = [
  { label: "Отступ контура", value: "Задаётся в диалоге при запуске (по умолч. 20 мм)" },
  { label: "Толщина обводки", value: "0.2 мм" },
  { label: "Цвет обводки", value: "CMYK 0 / 0 / 0 / 100" },
  { label: "Заливка", value: "Прозрачная (нет заливки)" },
  { label: "Привязка к группе", value: "Нет — независимый объект" },
  { label: "Вызов", value: "Диалоговое меню CorelDRAW" },
];

const paramsDots = [
  { label: "Размер точки", value: "10 × 10 мм" },
  { label: "Отступ от края", value: "10 мм наружу от выделения" },
  { label: "Цветовой режим", value: "CMYK" },
  { label: "Заливка", value: "CMYK 0 / 0 / 0 / 100 (чёрная)" },
  { label: "Абрис", value: "Отсутствует" },
  { label: "Работает с", value: "Один объект, несколько, группа" },
];

const functionsOutline = [
  { fn: "DrawGroupOutline()", role: "Точка входа", desc: "Проверяет документ и выделение, запрашивает отступ через InputBox.", color: "text-[#7c9cbf]" },
  { fn: "GetGroupBounds()", role: "Расчёт размеров", desc: "Считывает координаты и габариты выбранной группы.", color: "text-[#9cbf7c]" },
  { fn: "CreateOutlineRect()", role: "Создание контура", desc: "Рисует прямоугольник с заданным отступом, обводка CMYK 0/0/0/100.", color: "text-[#bf9c7c]" },
  { fn: "PositionOutline()", role: "Позиционирование", desc: "Центрирует контур относительно исходной группы.", color: "text-[#bf7cbf]" },
];

const functionsDots = [
  { fn: "DrawCornerDots()", role: "Единая точка входа", desc: "Получает bbox через Application.ActiveSelection — работает с объектом, несколькими и группой. Рисует 4 точки по углам.", color: "text-[#7c9cbf]" },
  { fn: "ActiveSelection", role: "Расчёт bbox", desc: "Суммарный bounding box всего выделения — LeftX, BottomY, SizeWidth, SizeHeight без цикла по объектам.", color: "text-[#9cbf7c]" },
  { fn: "corners()", role: "Координаты углов", desc: "Массив 4×2 с координатами левого нижнего угла каждого квадрата-точки.", color: "text-[#bf9c7c]" },
  { fn: "Цикл For k", role: "Создание точек", desc: "Рисует 4 прямоугольника 10×10 мм, задаёт заливку CMYK и убирает абрис.", color: "text-[#bf7cbf]" },
];

const KEYWORDS = ["Sub","End Sub","End Function","Function","Type","End Type","Dim","Set","As","New","If","Then","Exit","Const","True","False","MsgBox","Application","Nothing","For","Next","To","Each","In","Do","Loop","While","Wend","Not","And","Or","ElseIf","Else","End If","Integer","Double","String","Boolean","Long"];

function ColoredLine({ line }: { line: string }) {
  if (line.trim().startsWith("'")) {
    return <span className="text-[#4a4a4a]">{line}{"\n"}</span>;
  }
  const regex = new RegExp(`(\\b(?:${KEYWORDS.map(k => k.replace(/ /g, "\\s+")).join("|")})\\b)`, "g");
  const parts = line.split(regex);
  return (
    <span>
      {parts.map((part, j) =>
        KEYWORDS.includes(part.replace(/\s+/g, " "))
          ? <span key={j} className="text-[#7c9cbf]">{part}</span>
          : <span key={j}>{part}</span>
      )}
      {"\n"}
    </span>
  );
}

type Tab = "outline" | "dots";

export default function Index() {
  const [activeTab, setActiveTab] = useState<Tab>("dots");
  const [copied, setCopied] = useState(false);

  const isOutline = activeTab === "outline";
  const currentCode = isOutline ? vbaCodeOutline : vbaCodeDots;
  const currentParams = isOutline ? paramsOutline : paramsDots;
  const currentSteps = isOutline ? stepsOutline : stepsDots;
  const currentFunctions = isOutline ? functionsOutline : functionsDots;

  const handleCopy = () => {
    navigator.clipboard.writeText(currentCode);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="min-h-screen bg-[#0f0f0f] text-[#e8e8e8] font-mono">
      {/* Header */}
      <header className="border-b border-[#1e1e1e] px-8 py-5 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="w-6 h-6 border border-[#444] flex items-center justify-center">
            <div className="w-2.5 h-2.5 border border-[#888]" />
          </div>
          <span className="text-[11px] tracking-[0.2em] uppercase text-[#555]">
            CorelDRAW · VBA Macro
          </span>
        </div>
        <span className="text-[10px] text-[#333] tracking-widest">v 2.0</span>
      </header>

      <main className="max-w-4xl mx-auto px-8 py-12">

        {/* Tabs */}
        <div className="flex gap-px mb-12 border-b border-[#1a1a1a]">
          <button
            onClick={() => { setActiveTab("dots"); setCopied(false); }}
            className={`px-6 py-3 text-[11px] tracking-[0.15em] uppercase transition-all duration-150 border-b-2 -mb-px ${
              activeTab === "dots"
                ? "border-[#888] text-[#ddd]"
                : "border-transparent text-[#3a3a3a] hover:text-[#666]"
            }`}
          >
            Угловые точки
          </button>
          <button
            onClick={() => { setActiveTab("outline"); setCopied(false); }}
            className={`px-6 py-3 text-[11px] tracking-[0.15em] uppercase transition-all duration-150 border-b-2 -mb-px ${
              activeTab === "outline"
                ? "border-[#888] text-[#ddd]"
                : "border-transparent text-[#3a3a3a] hover:text-[#666]"
            }`}
          >
            Контур вокруг группы
          </button>
        </div>

        {/* Hero */}
        <div className="mb-14">
          <div className="text-[10px] tracking-[0.3em] uppercase text-[#444] mb-4">
            Инструмент автоматизации · CorelDRAW 26.1
          </div>
          {activeTab === "dots" ? (
            <>
              <h1 className="text-3xl font-bold tracking-tight text-white mb-3 leading-snug">
                Угловые точки<br />
                <span className="text-[#777]">метки реза по выделению</span>
              </h1>
              <p className="text-[#555] text-sm leading-relaxed max-w-lg">
                Рисует 4 квадратных точки 10×10 мм по углам выделения с отступом 10 мм наружу.
                Заливка CMYK 0/0/0/100, без абриса. Работает с одним объектом, несколькими и группой.
              </p>
            </>
          ) : (
            <>
              <h1 className="text-3xl font-bold tracking-tight text-white mb-3 leading-snug">
                Контур вокруг<br />
                <span className="text-[#777]">выбранной группы</span>
              </h1>
              <p className="text-[#555] text-sm leading-relaxed max-w-lg">
                Рисует прямоугольный контур вокруг выбранной группы.
                Отступ задаётся в диалоге при запуске. Четыре независимые функции — расчёт, создание, позиционирование.
              </p>
            </>
          )}
        </div>

        {/* Params */}
        <section className="mb-14">
          <div className="text-[10px] tracking-[0.25em] uppercase text-[#3a3a3a] mb-4">
            Параметры
          </div>
          <div className="grid grid-cols-2 md:grid-cols-3 gap-px bg-[#1a1a1a]">
            {currentParams.map((p) => (
              <div key={p.label} className="bg-[#0f0f0f] px-4 py-4">
                <div className="text-[9px] tracking-widest uppercase text-[#3a3a3a] mb-1">
                  {p.label}
                </div>
                <div className="text-sm text-[#bbb]">{p.value}</div>
              </div>
            ))}
          </div>
        </section>

        {/* Code */}
        <section className="mb-14">
          <div className="flex items-center justify-between mb-4">
            <div className="text-[10px] tracking-[0.25em] uppercase text-[#3a3a3a]">
              Код макроса
            </div>
            <button
              onClick={handleCopy}
              className="flex items-center gap-2 text-[11px] px-4 py-2 border border-[#222] text-[#555] hover:text-[#bbb] hover:border-[#444] transition-all duration-150"
            >
              {copied ? (
                <><Icon name="Check" size={11} />Скопировано</>
              ) : (
                <><Icon name="Copy" size={11} />Копировать</>
              )}
            </button>
          </div>

          <div className="bg-[#080808] border border-[#1a1a1a]">
            <div className="flex items-center gap-2 px-5 py-3 border-b border-[#161616]">
              <div className="w-2.5 h-2.5 rounded-full bg-[#3a3a3a]" />
              <div className="w-2.5 h-2.5 rounded-full bg-[#3a3a3a]" />
              <div className="w-2.5 h-2.5 rounded-full bg-[#3a3a3a]" />
              <span className="ml-2 text-[9px] text-[#2e2e2e] tracking-widest">
                {activeTab === "dots" ? "CornerDots.bas" : "GroupOutline.bas"}
              </span>
            </div>
            <pre className="p-6 text-[11px] leading-[1.9] overflow-x-auto">
              <code>
                {currentCode.split("\n").map((line, i) => (
                  <ColoredLine key={i} line={line} />
                ))}
              </code>
            </pre>
          </div>
        </section>

        {/* Steps */}
        <section className="mb-14">
          <div className="text-[10px] tracking-[0.25em] uppercase text-[#3a3a3a] mb-4">
            Установка
          </div>
          <div className="space-y-px">
            {currentSteps.map((s) => (
              <div
                key={s.num}
                className="flex gap-6 bg-[#0a0a0a] border border-[#161616] px-5 py-4 hover:border-[#252525] transition-colors"
              >
                <span className="text-[#222] text-xl font-bold tabular-nums leading-none mt-0.5 flex-shrink-0 w-7">
                  {s.num}
                </span>
                <div>
                  <div className="text-sm text-[#ccc] mb-1">{s.title}</div>
                  <div className="text-xs text-[#4a4a4a] leading-relaxed">{s.desc}</div>
                </div>
              </div>
            ))}
          </div>
        </section>

        {/* Architecture */}
        <section className="mb-14">
          <div className="text-[10px] tracking-[0.25em] uppercase text-[#3a3a3a] mb-4">
            Архитектура модуля
          </div>
          <div className="space-y-px">
            {currentFunctions.map((f) => (
              <div key={f.fn} className="bg-[#0a0a0a] border border-[#161616] px-5 py-4 flex gap-5">
                <div className="flex-shrink-0 w-44">
                  <div className={`text-xs font-bold ${f.color}`}>{f.fn}</div>
                  <div className="text-[9px] tracking-widest uppercase text-[#333] mt-0.5">{f.role}</div>
                </div>
                <div className="text-xs text-[#4a4a4a] leading-relaxed">{f.desc}</div>
              </div>
            ))}
          </div>
        </section>

        {/* Warning */}
        {isOutline && (
          <div className="border border-[#2a1f00] bg-[#0c0800] px-5 py-4 flex gap-3">
            <Icon name="TriangleAlert" size={14} className="text-[#6a4e00] flex-shrink-0 mt-0.5" />
            <div className="text-xs text-[#6a4e00] leading-relaxed">
              <strong className="text-[#9a7200]">Важно:</strong> Тип{" "}
              <code className="bg-[#140e00] px-1 text-[#9a7200]">BoundingBox</code>{" "}
              должен быть объявлен в самом начале модуля — до всех{" "}
              <code className="bg-[#140e00] px-1">Sub</code> и{" "}
              <code className="bg-[#140e00] px-1">Function</code>. CorelDRAW требует объявления пользовательских типов перед их использованием.
            </div>
          </div>
        )}

        {activeTab === "dots" && (
          <div className="border border-[#1a2a1a] bg-[#080c08] px-5 py-4 flex gap-3">
            <Icon name="Info" size={14} className="text-[#3a6a3a] flex-shrink-0 mt-0.5" />
            <div className="text-xs text-[#3a6a3a] leading-relaxed">
              Точки создаются как отдельные объекты на активном слое. После создания их можно выделить и сгруппировать с макетом вручную.
            </div>
          </div>
        )}

      </main>

      <footer className="border-t border-[#141414] px-8 py-5 mt-8 flex items-center justify-between">
        <span className="text-[9px] tracking-widest uppercase text-[#262626]">
          CorelDRAW 26.1 · VBA Macros
        </span>
        <span className="text-[9px] text-[#1e1e1e]">2026</span>
      </footer>
    </div>
  );
}