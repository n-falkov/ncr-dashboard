# NCR Pilot Dashboard — Auto-Deploy Setup

## Что получается

Ты кладёшь новый GR-файл в папку → запускаешь одну команду → через 30 сек дашборд обновлён на `ncr-pilot-bm.netlify.app`.

---

## Шаг 1: Создай GitHub репозиторий

1. Зайди на https://github.com/new
2. Название: `ncr-pilot-dashboard` (можно приватный)
3. Создай repo

Затем в терминале:

```bash
# Перейди в папку с проектом
cd ~/Projects/ncr-pilot-dashboard

# Скопируй туда файлы из этого пакета:
# - update_dashboard.py
# - template.html

# Инициализируй git
git init
git remote add origin https://github.com/ТВОЙ_USERNAME/ncr-pilot-dashboard.git

# Создай папку для данных
mkdir data

# Первый коммит
git add .
git commit -m "Initial setup"
git push -u origin main
```

---

## Шаг 2: Подключи Netlify

1. Зайди на https://app.netlify.com
2. **Add new site → Import an existing project**
3. Выбери GitHub → `ncr-pilot-dashboard`
4. Build settings:
   - **Build command:** (оставь пустым)
   - **Publish directory:** `.`
5. Deploy site
6. В **Site settings → Domain management** поменяй название на `ncr-pilot-bm.netlify.app` (или любое другое)

Готово! Каждый `git push` теперь автоматически обновляет сайт.

---

## Шаг 3: Размести файлы данных

Положи в папку `data/`:

```
data/
├── GR_weekly.xlsx          # General Report (недельный)
├── mgk_ath.xlsx            # Territory mapping
└── new_stores.csv          # NCR new stores (опционально)
```

> **Совет:** Можешь назвать файлы как угодно — скрипт найдёт любой `General_Report*.xlsx` если `GR_weekly.xlsx` не найден.

---

## Шаг 4: Установи зависимости (один раз)

```bash
pip3 install pandas openpyxl
```

---

## Шаг 5: Обновляй дашборд

Каждый раз когда получаешь новый GR-файл:

```bash
# 1. Положи новый файл в data/
cp ~/Downloads/General_Report_truncated-79.xlsx data/GR_weekly.xlsx

# 2. Запусти скрипт
python3 update_dashboard.py

# Всё! Через 30 сек открой ncr-pilot-bm.netlify.app
```

Скрипт автоматически:
- Парсит GR файл
- Матчит магазины к ATS через территориальную карту
- Генерирует index.html
- Коммитит и пушит в GitHub
- Netlify деплоит

---

## Быстрая команда (alias)

Добавь в `~/.zshrc` или `~/.bashrc`:

```bash
alias dashboard="cd ~/Projects/ncr-pilot-dashboard && python3 update_dashboard.py"
```

Теперь в любом терминале просто:

```bash
dashboard
```

---

## Структура проекта

```
ncr-pilot-dashboard/
├── update_dashboard.py    # Главный скрипт (парсинг + деплой)
├── template.html          # HTML-шаблон дашборда с маркерами /*__DATA__*/
├── index.html             # Сгенерированный дашборд (auto, не редактировать)
├── data/                  # Папка с данными (в .gitignore)
│   ├── GR_weekly.xlsx
│   ├── mgk_ath.xlsx
│   └── new_stores.csv
├── .gitignore
└── README.md
```

---

## Troubleshooting

**"No GR weekly file found"** → Положи файл в `data/` или переименуй в `GR_weekly.xlsx`

**"Territory file not found"** → Скопируй `mgk_ath.xlsx` в `data/`. Без него матчинг не работает.

**"Git error"** → Проверь что `git remote -v` показывает правильный GitHub URL и у тебя есть доступ (SSH key или token).

**Данные не обновились на сайте** → Подожди 1-2 минуты, Netlify деплоит автоматически. Проверь на https://app.netlify.com/sites/ncr-pilot-bm/deploys
