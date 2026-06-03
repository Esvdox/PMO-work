# Промпт для Cowork — копируй и вставляй в чат

## Минимальный (рутинный апдейт)

```
Workspace: C:\Users\sivanov\Desktop\Repo-for-claude-ppm-status

Прочитай prompt-daily-update.md и выполни все 13 шагов:
обнови блок DATA в index.html из PMO_Portfolio_Master_v7.xlsx,
сделай git add + commit + push в origin/main.

После завершения дай отчёт: что изменилось vs предыдущий коммит,
хэш нового коммита, ссылку на публичный дашборд.

Если git операции упадут с unable to unlink .git/index.lock —
сообщи об этом, не пытайся починить вручную, я сделаю push сам в PowerShell.
```

---

## Расширенный (с указанием что именно поменял)

```
Workspace: C:\Users\sivanov\Desktop\Repo-for-claude-ppm-status

Я обновил PMO_Portfolio_Master_v7.xlsx:
- [опиши коротко что изменил — например: "% выполнения по проектам Шалушкина,
   2 новых риска в проекте #6, статус #14 → Приостановлен"]

Выполни prompt-daily-update.md:
1. Сделай backup index.html → backups\
2. Пересчитай DATA из xlsx, замени блок в index.html
3. Покажи мне дифф (что изменилось в DATA vs текущая версия)
4. После моего ОК — git add + commit + push в origin/main

Commit message: "Daily update <дата>: <краткое описание моих изменений>"
```

---

## Только апдейт HTML, без push (для проверки)

```
Workspace: C:\Users\sivanov\Desktop\Repo-for-claude-ppm-status

Обнови DATA в index.html из PMO_Portfolio_Master_v7.xlsx согласно
prompt-daily-update.md. Сделай backup, но НЕ коммить и НЕ пуш —
я хочу открыть index.html в браузере и проверить визуально.

Дашборд я открою локально (двойной клик на index.html).
После моего ОК — отдельной командой попрошу закоммитить.
```

После проверки:

```
Всё ок. Закоммить с сообщением "Daily update <дата>" и запушь в origin/main.
```

---

## Если упало на git (workaround под Windows-баг)

Cowork сообщит что не может выполнить git operations. Тогда:

```
Понял, git с твоей стороны не работает. Покажи финальное состояние index.html
(размер файла, дата изменения) — я сам сделаю push в PowerShell.
```

Дальше в PowerShell:

```powershell
cd C:\Users\sivanov\Desktop\Repo-for-claude-ppm-status
Remove-Item .git\index.lock -ErrorAction SilentlyContinue
git add index.html
git commit -m "Daily update"
git push origin main
```

---

## Откат если Cowork сломал HTML

```
Workspace: C:\Users\sivanov\Desktop\Repo-for-claude-ppm-status

Посмотри в папке backups\ — найди самый свежий backup index-*.html.
Скопируй его обратно в корень как index.html (перезаписав текущий).

После этого:
git add index.html
git commit -m "Rollback: restore from backup"
git push origin main
```

---

## Подсказки

**Перед запуском Cowork:**
1. Открой PMO_Portfolio_Master_v7.xlsx в Excel, нажми Ctrl+S, закрой
   (это нужно чтобы формулы пересчитались и Cowork прочёл реальные значения)
2. Убедись что Excel не открыт (иначе xlsx залочен)

**После завершения Cowork:**
1. Открой https://github.com/Esvdox/PMO-work/commits/main — увидишь новый коммит
2. Через 1-2 минуты обнови https://esvdox.github.io/PMO-work/ в инкогнито
3. Проверь что данные изменились
