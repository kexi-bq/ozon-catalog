# Catalog 500

Git-ready локальная витрина на `Streamlit`.

Внутри:
- `500` товаров
- только `не метражные`
- только товары с фото
- данные уже зафиксированы в `data/catalog_500.xlsx`

## Запуск локально

```powershell
pip install -r requirements.txt
python -m streamlit run app.py --server.port 8512
```

или просто:

```bat
run.bat
```

Открыть:

`http://localhost:8512`

## Что лежит в папке

- `app.py` — сайт
- `data/catalog_500.xlsx` — данные каталога
- `images/` — локальные картинки для части товаров
- `requirements.txt` — зависимости

## Как залить на GitHub

Если у тебя ещё нет репозитория:

```powershell
cd "C:\Users\User\Desktop\общая прога для парсинга\ozon_demo_git_ready"
git init
git add .
git commit -m "catalog 500 site"
git branch -M main
git remote add origin https://github.com/USERNAME/REPO.git
git push -u origin main
```

## Если хочешь запустить на Streamlit Community Cloud

Нужны:
- репозиторий на GitHub
- корневой файл: `app.py`
- зависимости: `requirements.txt`

После пуша:
1. Открываешь Streamlit Community Cloud
2. New app
3. Выбираешь репозиторий
4. Main file path: `app.py`
5. Deploy

## Важно

Эта папка уже отдельная и подготовлена под git.
Лучше заливать именно `ozon_demo_git_ready`, а не весь большой рабочий проект.
