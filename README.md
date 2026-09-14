# 📺 YouTube Channel Scraper

Скрипт автоматически собирает все видео с любого YouTube-канала и сохраняет результат в красивый Excel-файл (.xlsx).

---

## ✨ Что умеет

- Собирает **все видео канала** или только последние N штук
- Сохраняет в **.xlsx** — открывается двойным кликом, кириллица без кракозябр
- **Умная обработка ошибок** — если файл открыт в Excel или куки устарели, скрипт объяснит что делать и повторит
- Автоматически находит файл куков на рабочем столе
- **Работает везде** — Windows, macOS, Linux (кроссплатформа)
- **CLI аргументы** — можно автоматизировать в скриптах
- **Открывает папку** автоматически на всех ОС

---

## 📋 Столбцы в Excel

| Столбец | Описание |
|---|---|
| Заголовок | Название видео |
| Дата | Дата публикации (ГГГГ-ММ-ДД) |
| Просмотры | Количество просмотров |
| Ссылка | Прямая ссылка на видео |
| Описание | Первые 300 символов описания |

---

## 🚀 Установка и запуск

### 1. Установи Python

Скачай с [python.org](https://www.python.org/downloads/) — при установке поставь галочку **"Add Python to PATH"**.

### 2. Скачай скрипт

```bash
git clone https://github.com/viktorkanuta-a11y/youtube-channel-scraper.git
cd youtube-channel-scraper
pip install -r requirements.txt
```

Или просто скачай `youtube_scraper.py` напрямую (зависимости установятся автоматически при первом запуске).

### 3. Получи cookies.txt

YouTube требует авторизацию для сбора данных:

1. Открой **Chrome** и зайди на [youtube.com](https://youtube.com) (войди в аккаунт)
2. Установи расширение **[Get cookies.txt LOCALLY](https://chromewebstore.google.com/detail/get-cookiestxt-locally/cclelndahbckbenkjhflpdbgdldlbecc)**
3. Нажми иконку расширения на youtube.com → **Export**
4. Сохрани файл на **Рабочий стол** — скрипт найдёт его автоматически

### 4. Запусти

#### **Интерактивный режим (как раньше):**
```bash
python youtube_scraper.py
```

Скрипт спросит:
- Ссылку на канал
- Сколько видео собрать (Enter = все)
- Подтверждение файла куков

#### **С аргументами командной строки (новое!):**

```bash
# Все видео из канала
python youtube_scraper.py --url https://www.youtube.com/@mkbhd

# Последние 50 видео
python youtube_scraper.py --url https://www.youtube.com/@mkbhd --limit 50

# Сохранить в конкретный файл
python youtube_scraper.py --url https://www.youtube.com/@mkbhd --output my_analysis.xlsx

# Указать путь к куки файлу
python youtube_scraper.py --url https://www.youtube.com/@mkbhd --cookies /path/to/cookies.txt

# Комбинированный запуск (максимум информации)
python youtube_scraper.py --url https://www.youtube.com/@mkbhd --limit 100 --output reports/mkbhd_2026.xlsx --cookies ~/Desktop/cookies.txt
```

**Справка по аргументам:**
```bash
python youtube_scraper.py --help
```

---

## 💡 Примеры использования

### Пример 1: Простой анализ канала
```bash
python youtube_scraper.py --url https://www.youtube.com/@mkbhd
```
Результат: файл `@mkbhd.xlsx` в текущей папке

### Пример 2: Анализ последних 50 видео
```bash
python youtube_scraper.py --url https://www.youtube.com/@mkbhd --limit 50
```

### Пример 3: Автоматизация (для скриптов)
```bash
python youtube_scraper.py \
  --url https://www.youtube.com/@mkbhd \
  --limit 100 \
  --output reports/mkbhd.xlsx \
  --cookies ~/Desktop/cookies.txt
```

### Пример 4: Анализ русского канала
```bash
python youtube_scraper.py --url https://www.youtube.com/@viktoriya_orlinskaya
```
Результат: кириллица без ошибок, файл открывается автоматически

---

## 📦 Зависимости

Устанавливаются **автоматически** при первом запуске:

- [yt-dlp](https://github.com/yt-dlp/yt-dlp) — загрузка метаданных YouTube
- [openpyxl](https://openpyxl.readthedocs.io/) — создание Excel-файлов

**Или вручную:**

```bash
pip install -r requirements.txt
```

**Или штучно:**
```bash
pip install yt-dlp openpyxl
```

---

## 🛠 Частые вопросы

**Скрипт пишет "YouTube не принимает куки"**
→ Куки устарели. Экспортируй их заново через расширение и сохрани на рабочий стол.

**Файл не сохраняется**
→ Закрой Excel — скрипт попросит об этом сам и повторит попытку.

**Видео пропускаются**
→ Некоторые видео могут быть удалены или ограничены по региону — скрипт пропускает их и продолжает.

**Какая разница между интерактивным и CLI аргументами?**
```
Интерактивный режим (без флагов):
├─ Вводишь данные в консоль
├─ Медленнее
└─ Удобнее для одноразовых запусков

CLI аргументы (--url, --limit):
├─ Передаёшь всё в команде
├─ Быстрее
└─ Удобнее для автоматизации/скриптов
```

**На каких ОС работает?**
- ✅ Windows (explorer открывает папку)
- ✅ macOS (open -R открывает папку)
- ✅ Linux (xdg-open открывает папку)

---

## 🔄 История обновлений

### v1.1.0 (13 сентября 2026)
- ✨ **CLI аргументы** — `--url`, `--limit`, `--output`, `--cookies`
- 🌍 **Кроссплатформа** — работает на Windows/Mac/Linux
- 🚀 **Улучшена обработка ошибок** — проверка пустых результатов, валидация ID видео
- 📁 **Открытие папок** — работает на всех ОС
- 📦 **requirements.txt** — простая установка зависимостей

### v1.0.0 (4 июня 2026)
- 🎬 **Первый релиз** — сбор видео в Excel
- 🎨 **Красивое форматирование** — заголовки, чередование строк
- 🛡️ **Умная обработка ошибок** — куки, занятый файл
- 🔍 **Автопоиск куков** — ищет на рабочем столе

---

## 📄 Лицензия

MIT — используй как хочешь, даже для коммерческих проектов

---

## 🤝 Помощь в развитии

Хочешь улучшить проект?
1. Форк репозитория
2. Создай ветку: `git checkout -b feature/твоя-фича`
3. Коммитай изменения: `git commit -m 'Добавил классную фичу'`
4. Пуш: `git push origin feature/твоя-фича`
5. Создай Pull Request

---

**Вопросы?** Открой Issue в репозитории!
