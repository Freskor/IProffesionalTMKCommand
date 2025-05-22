document.addEventListener('DOMContentLoaded', () => {
    const startButton = document.getElementById('start-button');
    const createNewButton = document.getElementById('create-new-button');
    const openExistingButton = document.getElementById('open-existing-button');
    const generateTableButton = document.getElementById('generate-table-button'); // Было generate-table-button
    const voiceInputButton = document.getElementById('voice-input-button');

    const rowsInput = document.getElementById('rows-input');
    const colsInput = document.getElementById('cols-input');
    const voiceStatus = document.getElementById('voice-status');

    const views = {
        initial: document.getElementById('initial-view'),
        actionSelection: document.getElementById('action-selection-view'),
        dimensionInput: document.getElementById('dimension-input-view')
    };
    // URL вашего бэкенд-сервиса
    // ЗАМЕНИТЕ НА ВАШ РЕАЛЬНЫЙ URL БЭКЕНДА
    const BACKEND_URL = 'http://localhost:5000'; // Пример для локальной разработки

    function showView(viewName) {
        Object.values(views).forEach(view => view.classList.remove('active'));
        if (views[viewName]) {
            views[viewName].classList.add('active');
        }
    }

    // --- Обработчики событий ---
    startButton.addEventListener('click', () => {
        showView('actionSelection');
    });

    createNewButton.addEventListener('click', () => {
        showView('dimensionInput');
    });
    
    openExistingButton.addEventListener('click', async () => {
        voiceStatus.textContent = 'Загрузка списка файлов...';
        try {
            const response = await fetch(`${BACKEND_URL}/api/files`); // Предполагаем, что есть такой эндпоинт
            if (!response.ok) {
                throw new Error(`Ошибка сервера: ${response.status} ${response.statusText}`);
            }
            const files = await response.json();

            if (files && files.length > 0) {
                // Здесь нужно будет отобразить список файлов.
                // Для примера просто выведем в alert и консоль.
                // В реальном приложении вы бы создали элементы списка (ul, li).
                alert(`Найденные файлы:\n${files.join('\n')}\n\n(Логика отображения и открытия конкретного файла не реализована в этом примере)`);
                console.log("Доступные файлы:", files);
                voiceStatus.textContent = `Загружено ${files.length} файлов.`;

                // TODO: Реализовать интерфейс для выбора файла из списка
                // и его открытия/скачивания по запросу к бэкенду, например:
                // `${BACKEND_URL}/api/files/НАЗВАНИЕ_ФАЙЛА`
            } else {
                alert('На сервере нет доступных файлов.');
                voiceStatus.textContent = 'Файлы не найдены.';
            }
        } catch (error) {
            console.error('Ошибка при получении списка файлов:', error);
            alert(`Не удалось загрузить список файлов: ${error.message}`);
            voiceStatus.textContent = 'Ошибка загрузки списка файлов.';
        }
    });

    generateTableButton.addEventListener('click', async () => {
        const rows = parseInt(rowsInput.value);
        const cols = parseInt(colsInput.value);
        
        if (isNaN(rows) || isNaN(cols) || rows < 1 || cols < 1) {
            alert('Пожалуйста, введите корректные размеры (целые числа больше 0).');
            return;
        }

        voiceStatus.textContent = 'Создание таблицы на сервере...';
        generateTableButton.disabled = true; // Блокируем кнопку на время запроса

        try {
            const response = await fetch(`${BACKEND_URL}/api/tables`, { // Эндпоинт для создания таблицы
                
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ rows, cols }),
            });


            if (!response.ok) {
                // Попробуем прочитать ошибку от сервера, если она в JSON
                let errorData;
                try {
                    errorData = await response.json();
                } catch (e) {
                    // Ошибка не в JSON, используем statusText
                }
                const errorMessage = errorData?.error || response.statusText || `HTTP error ${response.status}`;
                throw new Error(`Ошибка сервера: ${errorMessage}`);
            }

            // Предполагаем, что бэкенд возвращает JSON с информацией о файле,
            // включая URL для скачивания или имя файла.
            // Пример ответа от бэкенда:
            // { "message": "Table created", "filename": "table_xyz.xlsx", "downloadUrl": "/api/files/table_xyz.xlsx" }
            // или
            // { "message": "Table created", "fileUrl": "http://localhost:5000/static/tables/table_xyz.xlsx" }

            const result = await response.json();

            if (result.downloadUrl || result.fileUrl) {
                const downloadUrl = result.downloadUrl ? `${BACKEND_URL}${result.downloadUrl}` : result.fileUrl;
                voiceStatus.textContent = `Таблица "${result.filename || 'table.xlsx'}" создана. Начинается скачивание...`;

                // Инициируем скачивание файла
                // Способ 1: Прямое открытие URL (браузер сам решит, скачать или попытаться открыть)
                // window.open(downloadUrl, '_blank'); // Откроет в новой вкладке, если браузер не начнет скачивание сразу

                // Способ 2: Создание временной ссылки и клик по ней (более надежно для скачивания)
                const link = document.createElement('a');
                link.href = downloadUrl;
                // Имя файла при скачивании можно задать здесь, если бэкенд его явно не указывает через Content-Disposition
                link.setAttribute('download', result.filename || 'generated_table.xlsx');
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link); // Удаляем временную ссылку

                alert(`Файл "${result.filename || 'table.xlsx'}" должен начать скачиваться.`);
            } else {
                // Если бэкенд вернул только сырые данные файла (менее предпочтительно для .xlsx)
                // Этот блок более актуален, если бы бэкенд вернул, например, CSV как текст
                // или если бы мы хотели обработать Blob напрямую.
                // Для .xlsx лучше, чтобы бэкенд сразу дал ссылку на скачивание.

                // const blob = await response.blob(); // Если бэкенд возвращает сам файл как Blob
                // const downloadUrl = window.URL.createObjectURL(blob);
                // const link = document.createElement('a');
                // link.href = downloadUrl;
                // link.setAttribute('download', 'generated_table.xlsx'); // Укажите имя файла
                // document.body.appendChild(link);
                // link.click();
                // document.body.removeChild(link);
                // window.URL.revokeObjectURL(downloadUrl);
                // voiceStatus.textContent = 'Таблица создана и скачана.';
                // alert('Файл таблицы должен начать скачиваться.');
                console.warn('Бэкенд не вернул downloadUrl или fileUrl. Не удалось инициировать скачивание стандартным образом.');
                voiceStatus.textContent = 'Таблица создана, но URL для скачивания не получен.';
            }

        } catch (error) {
            console.error('Ошибка при создании таблицы:', error);
            alert(`Не удалось создать таблицу: ${error.message}`);
            voiceStatus.textContent = `Ошибка: ${error.message}`;
        } finally {
            generateTableButton.disabled = false; // Разблокируем кнопку
        }
    });
    

    // --- Голосовой ввод (остается как был, если не требует изменений) ---
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (SpeechRecognition) {
        const recognition = new SpeechRecognition();
        recognition.lang = 'ru-RU';
        recognition.interimResults = false;
        recognition.maxAlternatives = 1;

        voiceInputButton.addEventListener('click', () => {
            voiceStatus.textContent = 'Говорите...';
            try {
                recognition.start();
            } catch (e) {
                console.error("Ошибка запуска распознавания:", e);
                voiceStatus.textContent = 'Ошибка микрофона. Убедитесь, что доступ разрешен.';
            }
        });

        recognition.onresult = (event) => {
            const speechResult = event.results[0][0].transcript.toLowerCase();
            voiceStatus.textContent = `Распознано: ${speechResult}`;
            console.log('Confidence: ' + event.results[0][0].confidence);

            const match = speechResult.match(/(\d+|один|два|три|четыре|пять|шесть|семь|восемь|девять|десять)\s*(?:на|x|х)\s*(\d+|один|два|три|четыре|пять|шесть|семь|восемь|девять|десять)/);

            if (match) {
                const wordToNum = {
                    'один': 1, 'два': 2, 'три': 3, 'четыре': 4, 'пять': 5,
                    'шесть': 6, 'семь': 7, 'восемь': 8, 'девять': 9, 'десять': 10
                };

                let num1Text = match[1];
                let num2Text = match[2];

                let num1 = parseInt(num1Text);
                if (isNaN(num1) && wordToNum[num1Text.toLowerCase()]) num1 = wordToNum[num1Text.toLowerCase()];


                let num2 = parseInt(num2Text);
                if (isNaN(num2) && wordToNum[num2Text.toLowerCase()]) num2 = wordToNum[num2Text.toLowerCase()];


                if (num1 && num2) {
                    rowsInput.value = num1;
                    colsInput.value = num2;
                    voiceStatus.textContent = `Установлены размеры: ${num1} X ${num2}`;
                } else {
                    voiceStatus.textContent = 'Не удалось распознать числа. Попробуйте "5 на 3".';
                }
            } else {
                voiceStatus.textContent = 'Не удалось распознать формат. Попробуйте "Число на Число".';
            }
        };

        recognition.onspeechend = () => {
            recognition.stop();
        };

        recognition.onerror = (event) => {
            if (event.error === 'no-speech') {
                voiceStatus.textContent = 'Речь не распознана. Попробуйте снова.';
            } else if (event.error === 'audio-capture') {
                voiceStatus.textContent = 'Ошибка захвата аудио. Проверьте микрофон.';
            } else if (event.error === 'not-allowed') {
                voiceStatus.textContent = 'Доступ к микрофону запрещен.';
            } else {
                voiceStatus.textContent = `Ошибка: ${event.error}`;
            }
        };
    } else {
        voiceInputButton.style.display = 'none';
        voiceStatus.textContent = 'Голосовой ввод не поддерживается вашим браузером.';
    }

    // Удаляем старую функцию generateAndOpenTable, так как таблица генерируется на бэке
    // function generateAndOpenTable(rows, cols) { ... }

    // Начальное состояние
    showView('initial');
});
