# Конвертер презентаций в шаблон УГМУ

Сайт на GitHub Pages: перетаскиваете .pptx, получаете ZIP с презентациями в оформлении УГМУ и report.csv.
Вся обработка идёт в браузере (Pyodide), файлы никуда не отправляются.

## Публикация (один раз)
1. На github.com: **New repository**, имя например `ugmu-converter`, Public, создать.
2. **Add file -> Upload files**: перетащите ВСЕ файлы из этой папки, включая скрытый `.nojekyll`
   (если скрытый файл не виден, создайте пустой файл `.nojekyll` через Add file -> Create new file). Commit.
3. **Settings -> Pages -> Build and deployment**: Source = *Deploy from a branch*, Branch = `main`, папка `/ (root)`. Save.
4. Через 1-2 минуты сайт будет по адресу `https://ВАШ_ЛОГИН.github.io/ugmu-converter/`.

## Обновление шаблона
Замените файл `template.pptx` в репозитории (Upload files, затем Commit). Структура шаблона должна остаться:
титульный слайд, слайд раздела и слайд со шапкой (название, линия, номер).

## Файлы
- `index.html` - страница
- `convert_to_ugmu.py` - сам конвертер (тот же скрипт запускается и локально из командной строки)
- `template.pptx` - шаблон УГМУ
