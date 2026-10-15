# Презентации в стиле ВШМ СПбГУ для Claude

Скилл **`gsom-presentation`**: пишете Claude **«Создай презентацию в стиле ВШМ про …»** и получаете
готовый `.pptx` по брендбуку Высшей школы менеджмента СПбГУ — фирменные цвета, шрифт Inter,
логотип, карточки, таблицы, диаграммы, таймлайны.

Слайды собираются из **официального шаблона презентации ВШМ** (110 образцов), поэтому
дизайн совпадает с оригиналом. Claude подбирает слайды, заполняет их вашим текстом, данными
и фото, рендерит превью и сам исправляет огрехи.

![Пример](examples/demo/contact-sheet.png)

## Установка

### Claude Code — терминал, десктоп-приложение, VS Code
В терминале:
```bash
claude plugin marketplace add ardubrovskij-dotcom/gsom-design
claude plugin install gsom-design@gsom-design
```
Или внутри сессии Claude Code: `/plugin marketplace add ardubrovskij-dotcom/gsom-design`,
затем `/plugin install gsom-design@gsom-design`. После установки перезапустите сессию.

Один раз поставьте зависимости:
```bash
pip install python-pptx pillow lxml
```
Рекомендуется также LibreOffice и poppler: с ними Claude видит превью слайдов и проверяет их
(macOS: `brew install --cask libreoffice && brew install poppler`,
Ubuntu: `sudo apt install libreoffice poppler-utils`).

### claude.ai — веб и приложения
1. Скачайте готовый архив: **[gsom-presentation.zip](https://github.com/ardubrovskij-dotcom/gsom-design/raw/HEAD/dist/gsom-presentation.zip)** (38 МБ).
2. В claude.ai откройте **Customize → Skills → «+» → Create skill → Upload a skill** и загрузите архив.
3. Включите **Code execution and file creation**: на Free, Pro и Max — в Settings → Capabilities;
   на Team и Enterprise это делает владелец организации.

> Лимит размера загрузки в справке claude.ai не указан. Если архив не принимается,
> используйте Claude Code.

### Другие способы
- Клонировать репозиторий и запустить `claude` в его папке: скилл подключится как проектный.
- Скопировать `skills/gsom-presentation` в `~/.claude/skills/`.

**Подробная инструкция по работе со скиллом — [GUIDE.md](GUIDE.md).**

## Примеры запросов
- «Создай презентацию в стиле ВШМ: итоги приёмной кампании 2026, данные в приложенной таблице»
- «Сделай 10 слайдов для дня открытых дверей магистратуры ВШМ»
- «Переоформи эту презентацию в фирменный стиль ВШМ» (с приложенным файлом)
- «Make a GSOM-style deck in English about our Executive MBA partnership»

## Что внутри

```
skills/gsom-presentation/
  SKILL.md                  # инструкция для Claude: порядок работы и правила
  reference/brandbook.md    # выжимка брендбука: цвета, шрифт, логотип, элементы, фото
  reference/catalog.md      # каталог из 110 слайдов шаблона
  reference/spec.md         # формат описания презентации (spec.json)
  scripts/gsom_deck.py      # движок: list / fields / build / render / check
  scripts/catalog.py        # ключи и описания слайдов
  assets/template.pptx      # официальный шаблон ВШМ (подготовленный, см. ниже)
  assets/catalog/*.jpg      # превью слайдов с номерами редактируемых фигур
  assets/logos/             # все версии логотипа RU/EN (svg + png)
  assets/elements/          # фирменный паттерн и контейнеры для фото (svg)
  assets/fonts/             # Inter (SIL OFL 1.1)
tools/
  build_templates.py        # исходный шаблон из брендбука -> assets/template.pptx
  build_catalog.py          # превью каталога и reference/catalog.md
  package_skill.sh          # пересобрать dist/gsom-presentation.zip
examples/demo/              # пример spec.json, готовой .pptx и превью
dist/gsom-presentation.zip  # готовый архив скилла для claude.ai (tools/package_skill.sh)
```

Ручная сборка без Claude:
```bash
S=skills/gsom-presentation/scripts
python3 $S/gsom_deck.py list                       # каталог слайдов
python3 $S/gsom_deck.py fields cards-3             # поля слайда
python3 $S/gsom_deck.py build examples/demo/spec.json demo.pptx
python3 $S/gsom_deck.py render demo.pptx preview/  # PNG-превью
```

## Источник и отличия от исходного шаблона

Материалы взяты из папки «Брендбук ВШМ СПбГУ» на Google Drive: `00_Brandbook_GSOM_SPbU.pdf`,
«Шаблон презентации рус.pptx», логотипы, шрифт, паттерн, контейнеры.
`tools/build_templates.py` готовит шаблон так:

| Что | Было в шаблоне | Стало | Основание |
|---|---|---|---|
| Красный | `#BE1E2D` в теме | `#C31632` | брендбук, с. 23 |
| Шрифты | Inter, а также Euclid Circular A/B, Roboto и Arial в отдельных фигурах | Inter везде | брендбук, с. 34: Inter — для печати и документов, Euclid — только для сайта |
| Логотипы | svg с цветами `#BE1E2D` / `#51626F` | официальные файлы из папки «Логотип» | брендбук, с. 18–19 |
| Слайды | 129 | 110: убраны служебный слайд «Резюме», 2 пустых слайда, 9 разделителей-оглавлений самого шаблона и 7 слайдов с флагами | — |
| Размер | 50 МБ | 31 МБ: картинки ужаты до 1800 px по длинной стороне | — |

Английская версия собирается из того же шаблона: при `"lang": "en"` логотипы заменяются
на английские из брендбука.

Вопросы по фирменному стилю — GSOMmarketing@gsom.spbu.ru (контакт из брендбука).
