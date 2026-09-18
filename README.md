# VZHDO Executive Discipline Dashboard v7.2

Готовый проект для GitHub Pages (`/VZHDO/`).

## Обновление структуры
Замените только `public/structure.xlsx`, сохраняя 7 обязательных названий столбцов. Код менять не требуется.

## Локальная проверка сборки
```bash
npm ci
npm run build
```
Результат будет в `dist/`.

## GitHub Pages
1. Загрузите содержимое проекта в репозиторий `VZHDO` в ветку `main`.
2. В GitHub откройте **Settings → Pages**.
3. В **Build and deployment → Source** выберите **GitHub Actions**.
4. Workflow `.github/workflows/deploy.yml` соберёт и опубликует сайт.

Рабочий URL для репозитория пользователя `boler887-art`:
`https://boler887-art.github.io/VZHDO/`
