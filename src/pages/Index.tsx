import { useState } from "react";
import Icon from "@/components/ui/icon";

const vbaCode = `' ============================================
' Макрос CorelDRAW: Контур вокруг группы
' Версия: 1.0
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
    Dim doc As Document
    Dim grp As Shape

    If Application.Documents.Count = 0 Then
        MsgBox "Нет открытого документа.", vbExclamation, "Ошибка"
        Exit Sub
    End If

    Set doc = Application.ActiveDocument

    If doc.ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Выберите группу объектов.", vbExclamation, "Ошибка"
        Exit Sub
    End If

    Set grp = doc.ActiveSelection.Shapes(1)

    Dim bounds As BoundingBox
    bounds = GetGroupBounds(grp)

    Dim outline As Shape
    Set outline = CreateOutlineRect(doc, bounds)

    PositionOutline outline, bounds

    MsgBox "Контур успешно создан!", vbInformation, "Готово"
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
Function CreateOutlineRect(doc As Document, bounds As BoundingBox) As Shape
    Const OFFSET As Double = 20
    Const STROKE_MM As Double = 0.2

    Dim rW As Double: rW = bounds.w + OFFSET * 2
    Dim rH As Double: rH = bounds.h + OFFSET * 2

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
Sub PositionOutline(outline As Shape, bounds As BoundingBox)
    Const OFFSET As Double = 20

    Dim cx As Double: cx = bounds.x + bounds.w / 2
    Dim cy As Double: cy = bounds.y + bounds.h / 2

    Dim newX As Double: newX = cx - (bounds.w / 2 + OFFSET)
    Dim newY As Double: newY = cy - (bounds.h / 2 + OFFSET)

    outline.SetPosition newX, newY
End Sub`;

const steps = [
  {
    num: "01",
    title: "Открыть редактор VBA",
    desc: 'Сервис → Макросы → Редактор Visual Basic (или Alt+F11)',
  },
  {
    num: "02",
    title: "Создать модуль",
    desc: "В дереве проекта: ПКМ на проекте → Вставить → Модуль. Назвать GroupOutline.",
  },
  {
    num: "03",
    title: "Добавить тип BoundingBox",
    desc: "Скопировать блок Type BoundingBox (в комментарии вверху кода) в самое начало модуля и раскомментировать.",
  },
  {
    num: "04",
    title: "Вставить код макроса",
    desc: "Скопировать весь остальной код в тот же модуль ниже объявления типа.",
  },
  {
    num: "05",
    title: "Подключить к контекстному меню",
    desc: "Сервис → Настройка → Команды → найти DrawGroupOutline → перетащить в контекстное меню.",
  },
];

const params = [
  { label: "Отступ контура", value: "20 мм с каждой стороны" },
  { label: "Толщина обводки", value: "0.2 мм" },
  { label: "Цвет обводки", value: "CMYK 0 / 0 / 0 / 100" },
  { label: "Заливка", value: "Прозрачная (нет заливки)" },
  { label: "Привязка к группе", value: "Нет — независимый объект" },
  { label: "Вызов", value: "Контекстное меню CorelDRAW" },
];

const functions = [
  {
    fn: "DrawGroupOutline()",
    role: "Точка входа",
    desc: "Проверяет наличие документа и выделения, запускает цепочку вызовов.",
    color: "text-[#7c9cbf]",
  },
  {
    fn: "GetGroupBounds()",
    role: "Расчёт размеров",
    desc: "Считывает координаты и габариты выбранной группы, возвращает структуру BoundingBox.",
    color: "text-[#9cbf7c]",
  },
  {
    fn: "CreateOutlineRect()",
    role: "Создание контура",
    desc: "Рисует прямоугольник +20 мм с каждой стороны, задаёт обводку CMYK 0/0/0/100 и прозрачную заливку.",
    color: "text-[#bf9c7c]",
  },
  {
    fn: "PositionOutline()",
    role: "Позиционирование",
    desc: "Вычисляет центр исходной группы и смещает контур так, чтобы отступ был ровно 20 мм с каждой стороны.",
    color: "text-[#bf7cbf]",
  },
];

const KEYWORDS = ["Sub","End Sub","End Function","Function","Type","End Type","Dim","Set","As","New","If","Then","Exit","Const","True","False","MsgBox","Application","Nothing"];

function ColoredLine({ line }: { line: string }) {
  if (line.trim().startsWith("'")) {
    return <span className="text-[#4a4a4a]">{line}{"\n"}</span>;
  }
  const parts = line.split(/(\b(?:Sub|End Sub|End Function|Function|Type|End Type|Dim|Set|As|New|If|Then|Exit|Const|True|False|MsgBox|Application|Nothing)\b)/g);
  return (
    <span>
      {parts.map((part, j) =>
        KEYWORDS.includes(part)
          ? <span key={j} className="text-[#7c9cbf]">{part}</span>
          : <span key={j}>{part}</span>
      )}
      {"\n"}
    </span>
  );
}

export default function Index() {
  const [copied, setCopied] = useState(false);

  const handleCopy = () => {
    navigator.clipboard.writeText(vbaCode);
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
        <span className="text-[10px] text-[#333] tracking-widest">v 1.0</span>
      </header>

      <main className="max-w-4xl mx-auto px-8 py-12">
        {/* Hero */}
        <div className="mb-14">
          <div className="text-[10px] tracking-[0.3em] uppercase text-[#444] mb-4">
            Инструмент автоматизации
          </div>
          <h1 className="text-3xl font-bold tracking-tight text-white mb-3 leading-snug">
            Контур вокруг<br />
            <span className="text-[#777]">выбранной группы</span>
          </h1>
          <p className="text-[#555] text-sm leading-relaxed max-w-lg">
            Макрос рисует прямоугольный контур с отступом 20 мм вокруг
            любой выбранной группы. Четыре независимые функции —
            расчёт, создание, позиционирование.
          </p>
        </div>

        {/* Params */}
        <section className="mb-14">
          <div className="text-[10px] tracking-[0.25em] uppercase text-[#3a3a3a] mb-4">
            Параметры
          </div>
          <div className="grid grid-cols-2 md:grid-cols-3 gap-px bg-[#1a1a1a]">
            {params.map((p) => (
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
                <>
                  <Icon name="Check" size={11} />
                  Скопировано
                </>
              ) : (
                <>
                  <Icon name="Copy" size={11} />
                  Копировать
                </>
              )}
            </button>
          </div>

          <div className="bg-[#080808] border border-[#1a1a1a]">
            <div className="flex items-center gap-2 px-5 py-3 border-b border-[#161616]">
              <div className="w-2.5 h-2.5 rounded-full bg-[#3a3a3a]" />
              <div className="w-2.5 h-2.5 rounded-full bg-[#3a3a3a]" />
              <div className="w-2.5 h-2.5 rounded-full bg-[#3a3a3a]" />
              <span className="ml-2 text-[9px] text-[#2e2e2e] tracking-widest">
                GroupOutline.bas
              </span>
            </div>
            <pre className="p-6 text-[11px] leading-[1.9] overflow-x-auto">
              <code>
                {vbaCode.split("\n").map((line, i) => (
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
            {steps.map((s) => (
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
            {functions.map((f) => (
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
      </main>

      <footer className="border-t border-[#141414] px-8 py-5 mt-8 flex items-center justify-between">
        <span className="text-[9px] tracking-widest uppercase text-[#262626]">
          GroupOutline.bas · CorelDRAW VBA
        </span>
        <span className="text-[9px] text-[#1e1e1e]">2026</span>
      </footer>
    </div>
  );
}
