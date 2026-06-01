# PMO Dashboard · Ежедневное обновление через Cowork

## Контекст

Я — Сергей Иванов, Директор PMO ООО «Парк Сказка» (Москва, Крылатское).

**Рабочая папка:** `C:\Users\sivanov\Desktop\Repo-for-claude-ppm-status`

В ней:
- `index.html` — публичный дашборд портфеля проектов (опубликован на GitHub Pages)
- `PMO_Portfolio_Master_v7.xlsx` — мастер-файл портфеля с актуальными данными
- `prompt-daily-update.md` — этот файл (моя инструкция)
- `backups\` — папка с резервными копиями HTML

Дашборд использует встроенные данные в JS-блок `<script>const DATA = window.DATA = { ... }</script>`. Твоя задача — пересчитать DATA из xlsx и заменить в HTML, **не трогая ничего другого**.

---

## Задача дня

Обнови блок `DATA` в `index.html` свежими данными из `PMO_Portfolio_Master_v7.xlsx`. Затем закоммить и запушь изменения.

### Шаг 1 · Backup

```powershell
if (-not (Test-Path "backups")) { mkdir backups }
Copy-Item index.html "backups\index-$(Get-Date -Format 'yyyy-MM-dd-HHmm').html"
```

### Шаг 2 · Прочитай xlsx

⚠ **Важно про v7-файл:** все 5 ключевых листов теперь оформлены как **Excel-таблицы** (`ListObject`), а не просто диапазоны. Это значит:
- Формулы используют структурированные ссылки `[@[Имя_колонки]]` и **сами растягиваются** при добавлении новых строк
- При добавлении проекта в `01_Проекты` колонки `% выполнения`, `Откл. дн.`, `SPI`, `CPI`, `Статус RAG` рассчитываются автоматически
- При добавлении риска в `02_Риски` авто-рассчитываются `Score (P×I)`, `EMV`, `Дней до триг.`, `EMV active`
- При чтении через pandas — формулы могут отдавать NaN если файл создан недавно и **не был открыт в Excel**. Открой xlsx в Excel один раз → формулы пересчитаются → сохрани

**Список таблиц в v7:**

| Лист | Имя таблицы | Что внутри |
|------|-------------|------------|
| `01_Проекты` | `tbl_Projects` | 30 проектов (5 авто-формул) |
| `02_Риски` | `tbl_Risks` | реестр рисков (4 авто-формулы) |
| `04_PM` | `tbl_PM` | команда руководителей |
| `07_Описания` | `tbl_Descriptions` | описания и стадии |
| `13_Каталог_рисков` | `tbl_RiskCatalog` | каталог типовых рисков |

```python
import pandas as pd
xlsx = 'PMO_Portfolio_Master_v7.xlsx'
df_p  = pd.read_excel(xlsx, sheet_name='01_Проекты')
df_r  = pd.read_excel(xlsx, sheet_name='02_Риски')
df_pm = pd.read_excel(xlsx, sheet_name='04_PM')
df_d  = pd.read_excel(xlsx, sheet_name='07_Описания')
```

Если в колонке `% выполнения` или `SPI` все значения NaN — попроси меня открыть xlsx в Excel один раз для пересчёта.

Колонки `01_Проекты`: `ID, Название проекта, PM ФИО, PM ID, Тип, Приоритет, Статус, Бюджет, EV, AC, PV, % выполнения, Дата старт, Дата baseline, Дата прогноз, Откл. дн., BSC-цель, Статус RAG, SPI, CPI`.

### Шаг 3 · Пересчитай агрегаты

```
portfolio.total       = всего строк
portfolio.active      = со статусом «Активный»
portfolio.suspended   = «Приостановлен»
portfolio.backlog     = «Backlog»
portfolio.budgetMln   = sum(Бюджет) активных, round 1 знак
portfolio.avgPct      = avg(% выполнения) активных, round до целого
portfolio.rag.{ahead,ontime,risky,late,noData}
                      = подсчёт активных по «Статус RAG»
```

### Шаг 4 · EVM-блок

```
health.bac = sum(Бюджет) активных
health.ev  = sum(EV)  активных
health.ac  = sum(AC)  активных
health.pv  = sum(PV)  активных
health.spi = round(ev / pv, 2)
health.cpi = round(ev / ac, 2)
health.eac = round(bac / cpi, 1)
health.cv  = round(ev - ac, 1)
```

Если у части проектов EV есть, а PV нет — портфельный SPI/CPI искажён. Тогда:
```
health.dataIssue = "У X проектов EV без PV — портфельный SPI/CPI искажён"
```

### Шаг 5 · DATA.managers[]

Для каждого PM из `04_PM`, у которого есть хотя бы один проект:

```javascript
{
  id: "shalushkin"|"ermolaeva"|"kiselev"|"ivanov",
  name: "Юрий Шалушкин",
  role: "Старший PM",
  rank: 1|2|3|4,
  projects: [
    {
      id: <ID>, name: <Название>, priority: "P1"|"P2"|"P3"|"P4",
      type: <Тип>, bsc: <BSC>, status: "Активный"|"Приостановлен",
      launch: "DD.MM.YYYY", baseline: "DD.MM.YYYY"|null, forecast: "DD.MM.YYYY"|null,
      budget: 299.989, ac: 268.827, pct: 89.6, spi: 1.0, cpi: 1.0,
      rag: "ontime"|"ahead"|"risky"|"late"|"no-data"
    },
    ...
  ]
}
```

**В managers не включай проекты со статусом Backlog** — они идут в `DATA.backlog`.

### Шаг 6 · DATA.backlog[]

Все проекты со статусом «Backlog»:
```javascript
{ id, name, priority: "P4", type, bsc, pm: <Фамилия>, budget, note }
```

### Шаг 7 · DATA.risksMatrices[]

Для каждого активного проекта собрать матрицу 5×5 рисков:
```javascript
{
  id, name, pm: <Фамилия>,
  mx: [[0,0,0,0,0],...],     // 5×5, mx[5-P][I-1] = счётчик
  identified, active, burned, realized,
  res: <сумма EMV активных рисков>,
  isBacklog: false,
  isSuspended: (true для приостановленных)
}
```

### Шаг 8 · DATA.topRisks[]

Топ-5 рисков по score = P × I, отсортированных desc:
```javascript
{ score: 20, proj: <название>, lbl: <описание>, emv: <число> }
```

### Шаг 9 · DATA.timeline[]

Ближайшие 10 запусков активных проектов (по baseline в будущем):
```javascript
{ name, pm, date: "DD.MM", daysLeft }
```

### Шаг 10 · DATA.meta

```javascript
{
  updatedAt: "DD.MM.YYYY HH:MM",
  source: "PMO_Portfolio_Master_v7.xlsx",
  version: "v2.9"
}
```

---

## Шаг 11 · Замена блока DATA в index.html

**Найди в файле:**
```javascript
const DATA = window.DATA = {
  meta: { ... },
  health: { ... },
  portfolio: { ... },
  managers: [ ... ],
  backlog: [ ... ],
  risksMatrices: [ ... ],
  topRisks: [ ... ],
  timeline: [ ... ],
};
```

**Замени только содержимое от `{` до `};`** — больше ничего.

### НЕ ТРОГАТЬ:

- CSS между `<style>` и `</style>`
- HTML (sidebar, секции, hero, footer)
- Все функции после DATA: `renderPortfolioKpis`, `renderHealth`, `renderRag`, `renderPhases`, `renderTimeline`, `renderPmSections`, `renderRiskMatrices`, `renderTopRisks`, `renderReserves`, `renderTable`, `renderFunnel`, `renderInitiatives`, `renderFilterChips`, `renderBacklogFull`, `bindFilters`, `switchView`, `handleHashRoute`, `getPhase`, `parseRuDate`, `esc`
- `PM_PHOTOS` — base64-фото PM (~52 КБ)
- `PM_EMOJI` — фолбэк-эмодзи
- `DEMO_INITIATIVES` — demo-инициативы ветки 2

### Проверка валидности

После замены проверь что `const DATA` — валидный JS-объект:
- ключи без кавычек: `id:` а не `"id":`
- строковые значения в двойных кавычках: `"Активный"`
- числа без кавычек: `299.989`
- `null` без кавычек
- запятые между полями
- скобки `{}` и `[]` сбалансированы

Если что-то не сходится — восстанови из backup и сообщи.

---

## Шаг 12 · Git commit и push

⚠ **Известный баг Cowork на Windows** (issue anthropics/claude-code #55206): git операции из bash-sandbox на mounted host folder могут падать с «unable to unlink .git/index.lock».

Сначала попробуй обычным путём:

```powershell
cd C:\Users\sivanov\Desktop\Repo-for-claude-ppm-status
git status
git add index.html
git diff --cached --stat
git commit -m "Daily update $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
git push origin main
```

**Если падает на unlink/lockfile** — не пытайся чистить `.git/` командами из sandbox. Просто выйди из git и в финальном отчёте укажи: «git операции нужно выполнить вручную в PowerShell». Я сделаю сам:

```powershell
# Что я сделаю вручную если упало:
cd C:\Users\sivanov\Desktop\Repo-for-claude-ppm-status
Remove-Item .git\index.lock -ErrorAction SilentlyContinue
git add index.html
git commit -m "Daily update"
git push
```

---

## Шаг 13 · Финальный отчёт

```
✓ Обновление выполнено: 26.05.2026 09:15
Источник: PMO_Portfolio_Master_v7.xlsx → index.html
Backup: backups\index-2026-05-26-0915.html

Изменения vs вчера:
  · Активных проектов:    17 → 17
  · Бюджет активных:      2 080.4 → 2 095.1 млн ₽ (+14.7)
  · Средний %:            36% → 39% (+3 п.п.)
  · SPI портфеля:         3.57 → 3.42
  · CPI портфеля:         3.55 → 3.40
  · RAG:                  12 ontime, 2 risky, 2 late, 1 no-data
  · Новых рисков:         2 (в проектах #4, #11)

Git:
  ✓ Commit a1b2c3d опубликован
  ✓ Push в origin/main
  → через 1-2 мин обновится https://<логин>.github.io/<репо>/

[или: ⚠ git упал — нужен ручной push в PowerShell]
```

Если данные не изменились — `push` НЕ делай, просто сообщи: «Изменений нет».

---

## Безопасность

1. **Никогда не меняй** структуру HTML, CSS, JS-функции — только `DATA`.
2. **Никогда не удаляй** `PM_PHOTOS`, `PM_EMOJI`, `DEMO_INITIATIVES`.
3. **Всегда делай backup** перед изменением.
4. **Если xlsx не читается или данные пустые** — НЕ коммить, сообщи.
5. **Не меняй** хардкод-тексты в hero-секции (миссия PMO, методология).
6. **Не выкладывай в коммит** xlsx-файл (он `.gitignore`-нут, но всё равно проверяй).
