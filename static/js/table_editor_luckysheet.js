document.addEventListener('DOMContentLoaded', () => {
    const voiceCommandButton = document.getElementById('voice-command-button');
    const voiceStatusEditor = document.getElementById('voice-status-editor');
    const saveTableButton = document.getElementById('save-table-button');
    const manualCommandInput = document.getElementById('manual-command-input');
    const executeManualCommandButton = document.getElementById('execute-manual-command-button');
    
    const BACKEND_URL = '';                         
    let isTableDirty = false;             
    let recognitionEditor;          
    let isRecognizingEditor = false;          
    let currentSaveOperation = null;             

    let voiceControlActive = false;                   
    let recognizedTextBuffer = "";                  
    let commandTriggers = ["записать", "ячейка", "перейти", "столбец", "вверх", "вниз", "влево", "вправо", "отмена", "назад", "повторить", "вперед", "очистить", "поиск", "рассчитать сумму", "сохранить", "охранить"];
    let triggerDetected = false;
    let COLUMN_HEADER_TRIGGERS = [];
    let lastSpeechTime = Date.now();
    const SPEECH_TIMEOUT_MS = 2500;                   
    let speechTimeoutId = null;

    const luckysheetBaseOptions = {
        container: 'luckysheet',
        lang: 'en',
        title: "MySheet",                      
        allowEdit: true,
        showtoolbar: false,                
        showinfobar: false,       
        showsheetbar: false,
        showstatisticBar: true,
        sheetFormulaBar: false,
        enableAddRow: true,
        enableAddCol: true,
        data: [],             
        hook: {
            updated: function (operate) {                         
                console.log('Luckysheet Hook: updated. Operate object:', operate);
                const nonModifyingOps = ['scroll', 'zoom', 'resize', 'showtoolbar', 'showsheetbar', 'sheetactivate', 'hover', 'focus', 'selection'];
                if (operate && operate.type && nonModifyingOps.includes(operate.type)) {
                    console.log(`isTableDirty НЕ установлен, операция: ${operate.type}`);
                    return;                   
                }

                if (!isTableDirty) {
                    isTableDirty = true;
                    console.log(`isTableDirty установлен в true хуком 'updated'. Тип операции: ${operate ? operate.type : 'N/A'}`);
                    if (saveTableButton) saveTableButton.classList.add('dirty');       
                }
            },
            cellUpdated: function (r, c, oldVal, newVal, isRefresh) {                
                console.log(`Hook: cellUpdated r:${r}, c:${c}, oldVal:${JSON.stringify(oldVal)}, newVal:${JSON.stringify(newVal)}, isRefresh:${isRefresh}`);
                if (!isRefresh && JSON.stringify(oldVal) !== JSON.stringify(newVal)) {                
                    if (!isTableDirty) {
                        isTableDirty = true;
                        console.log("isTableDirty установлен в true хуком 'cellUpdated'");
                        if (saveTableButton) saveTableButton.classList.add('dirty');
                    }
                }
            },
            setCellValue: function (r, c, value) {                
                console.log(`Hook: setCellValue r:${r}, c:${c}, value:`, value);
                if (!isTableDirty) {
                    isTableDirty = true;
                    console.log("isTableDirty установлен в true хуком 'setCellValue'");
                    if (saveTableButton) saveTableButton.classList.add('dirty');
                }
            },
            sheetCreateafter: function (newSheet) { isTableDirty = true; console.log("Sheet created, isTableDirty=true"); if (saveTableButton) saveTableButton.classList.add('dirty'); },
            sheetDeleteBefore: function (sheet) { isTableDirty = true; console.log("Sheet to be deleted, isTableDirty=true"); if (saveTableButton) saveTableButton.classList.add('dirty'); },
            sheetCopyAfter: function (newSheet) { isTableDirty = true; console.log("Sheet copied, isTableDirty=true"); if (saveTableButton) saveTableButton.classList.add('dirty'); },
            sheetEditNameAfter: function (newSheetName, oldSheetName) { isTableDirty = true; console.log("Sheet renamed, isTableDirty=true"); if (saveTableButton) saveTableButton.classList.add('dirty'); },
            commandExecuted: function (command) {
                console.log("Hook: commandExecuted", command);
                let commandName = '';
                if (typeof command === 'string') {
                    commandName = command;
                } else if (command && command.type) {
                    commandName = command.type;
                } else if (command && command.action) {                      
                    commandName = command.action;
                }


                const nonModifyingCommands = [
                    'scroll', 'rangeSelect', 'hover', 'focus', 'search', 'filter',
                    'undo', 'redo',                               
                    'sheetactivate', 'zoomChange', ' colaboración', 'history', 'comment',    
                    'pivotTable', 'chart', 'screenshot'
                ];

                const alwaysModifyingCommands = [
                    'setCellValue', 'clearValue', 'deleteRange', 'insertRow', 'deleteRow',
                    'insertColumn', 'deleteColumn', 'mergeCell', 'cancelMerge', 'sort',
                    'setRangeFormat',          
                ];

                if (commandName && alwaysModifyingCommands.includes(commandName)) {
                    if (!isTableDirty) {
                        isTableDirty = true;
                        console.log(`isTableDirty установлен в true из-за команды (alwaysModifying): ${commandName}`);
                        if (saveTableButton) saveTableButton.classList.add('dirty');
                    }
                } else if (commandName && !nonModifyingCommands.includes(commandName.toLowerCase())) {
                } else {
                    console.log(`isTableDirty НЕ установлен, команда (nonModifying или неясная): ${commandName}`);
                }
            }
        }
    };
    let isLuckysheetReady = false;

    function updateColumnHeadersAsTriggers() {
        if (!isLuckysheetReady || !luckysheet || typeof luckysheet.getCellValue !== 'function') {
            console.warn("updateColumnHeadersAsTriggers: Luckysheet не готов или getCellValue недоступен.");
            return;
        }

        COLUMN_HEADER_TRIGGERS = [];       
        const currentSheetObject = luckysheet.getSheet();
        if (!currentSheetObject || typeof currentSheetObject.index === 'undefined') {
            console.warn("updateColumnHeadersAsTriggers: Не удалось получить текущий лист.");
            return;
        }
        const currentSheetIndexString = currentSheetObject.index;
        const allSheetFiles = luckysheet.getLuckysheetfile();
        if (!allSheetFiles) return;
        const currentSheetFile = allSheetFiles.find(sheet => sheet.index === currentSheetIndexString);
        if (!currentSheetFile) return;

        const defaultColsToScan = (luckysheet.defaultConfig ? luckysheet.defaultConfig.columnlen : 26);
        console.log(`Обновление триггеров-заголовков для листа: ${currentSheetFile.name}. Сканирование до ${currentSheetFile.column || defaultColsToScan} столбцов.`);

        for (let c = 0; c < (currentSheetFile.column || defaultColsToScan); c++) {
            const headerVal = luckysheet.getCellValue(0, c, { sheetIndex: currentSheetIndexString, type: 'm' });
            if (headerVal && String(headerVal) !== "") {
                const triggerText = String(headerVal).trim().toLowerCase();
                if (!commandTriggers.includes(triggerText) && !COLUMN_HEADER_TRIGGERS.includes(triggerText)) {                   
                    COLUMN_HEADER_TRIGGERS.push(triggerText);
                }
            }
        }

        console.log("Обновленные триггеры-заголовки:", COLUMN_HEADER_TRIGGERS);
    }

    async function initializeLuckysheet() {
        if (!currentTableFilename) {
            console.error("Имя файла не определено! Невозможно инициализировать таблицу.");
            if (voiceStatusEditor) voiceStatusEditor.textContent = "Ошибка: Имя файла не задано.";
            return;
        }
        try {
            if (voiceStatusEditor) voiceStatusEditor.textContent = 'Загрузка данных таблицы...';
            const response = await fetch(`${BACKEND_URL}/api/table-data-luckysheet/${currentTableFilename}`);

            if (!response.ok) {
                throw new Error(`Ошибка сервера при загрузке данных: ${response.status} ${response.statusText}`);
            }
            const result = await response.json();
            let optionsToUse;

            if (result && result.data && Array.isArray(result.data) && result.data.length > 0) {
                optionsToUse = { ...luckysheetBaseOptions, data: result.data, title: currentTableFilename };
                if (voiceStatusEditor) voiceStatusEditor.textContent = `Таблица "${currentTableFilename}" загружена.`;
            } else {
                console.warn("Данные с сервера пусты или некорректны, создается пустая таблица.");
                const emptySheetData = [{
                    "name": "Sheet1", "celldata": [], "order": 0, "index": "0", "status": 1,
                    "row": luckysheetBaseOptions.defaultRow || 84,
                    "column": luckysheetBaseOptions.defaultCol || 26,
                }];
                optionsToUse = { ...luckysheetBaseOptions, data: emptySheetData, title: currentTableFilename };
                if (voiceStatusEditor) voiceStatusEditor.textContent = `Таблица "${currentTableFilename}" загружена (пустая).`;
            }

            if (typeof luckysheet === 'undefined' || !luckysheet.create) {
                console.error("Объект luckysheet или метод luckysheet.create не определен!");
                if (voiceStatusEditor) voiceStatusEditor.textContent = "Ошибка: Библиотека Luckysheet не загружена.";
                return;
            }
            luckysheet.create(optionsToUse);
            isLuckysheetReady = true;          
            console.log("Luckysheet готов к работе.");
            isTableDirty = false;
            if (saveTableButton) saveTableButton.classList.remove('dirty');
            requestMicrophonePermissionAndInitSpeech(); 
            updateColumnHeadersAsTriggers();
        } catch (error) {
            console.error('Критическая ошибка при инициализации Luckysheet:', error);
            if (voiceStatusEditor) voiceStatusEditor.textContent = `Ошибка инициализации: ${error.message}.`;
        }
    }

    async function saveTableDataToServer(isAutoSave = false) {
        if (currentSaveOperation) {
            console.log("Сохранение уже выполняется. Новая попытка отложена.");
            if (!isAutoSave && voiceStatusEditor) voiceStatusEditor.textContent = "Сохранение уже идет...";
            return currentSaveOperation;             
        }

        if (typeof luckysheet === 'undefined' || !luckysheet.getAllSheets) {
            if (voiceStatusEditor) voiceStatusEditor.textContent = "Таблица не готова к сохранению.";
            console.error("saveTableDataToServer: Luckysheet API недоступно.");
            return Promise.reject(new Error("Luckysheet API недоступно."));
        }

        if (!isTableDirty && !isAutoSave) {                      
            if (voiceStatusEditor) voiceStatusEditor.textContent = "Нет изменений для сохранения.";
            console.log("saveTableDataToServer: Нет изменений для ручного сохранения.");
            setTimeout(() => {
                if (voiceStatusEditor && voiceStatusEditor.textContent === "Нет изменений для сохранения.") {
                    voiceStatusEditor.textContent = "Готов к командам...";
                }
            }, 3000);
            return Promise.resolve({ message: "Нет изменений." });
        }
        if (!isTableDirty && isAutoSave) {
            console.log("Автосохранение: нет изменений.");
            return Promise.resolve({ message: "Нет изменений для автосохранения." });
        }


        if (voiceStatusEditor) voiceStatusEditor.textContent = isAutoSave ? 'Автосохранение...' : 'Сохранение данных...';
        if (saveTableButton) saveTableButton.disabled = true;

        const allSheetData = luckysheet.getAllSheets();
        console.log("Данные для отправки на сервер (allSheetData):", JSON.stringify(allSheetData).substring(0, 500) + "...");          

        if (!allSheetData || !Array.isArray(allSheetData) || allSheetData.length === 0) {
            console.error("КРИТИЧЕСКАЯ ОШИБКА: luckysheet.getAllSheets() вернул некорректные данные.");
            if (voiceStatusEditor) voiceStatusEditor.textContent = "Ошибка: Не удалось получить данные таблицы.";
            if (saveTableButton) saveTableButton.disabled = false;
            return Promise.reject(new Error("Не удалось получить данные таблицы."));
        }
        if (allSheetData[0] && typeof allSheetData[0].celldata === 'undefined') {
            console.warn("Предупреждение: Первый лист в allSheetData не содержит 'celldata'. Это может быть нормально для пустого листа.");
        }


        currentSaveOperation = fetch(`${BACKEND_URL}/api/save-table-data-luckysheet/${currentTableFilename}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ data: allSheetData })
        })
            .then(async response => {
                const responseText = await response.text();                
                console.log("Ответ сервера (raw):", responseText);

                if (!response.ok) {
                    let errorData;
                    try { errorData = JSON.parse(responseText); } catch (e) { errorData = { error: `Ошибка сервера (не JSON): ${response.status} ${responseText}` }; }
                    const errorMessage = errorData.error || `Ошибка сервера: ${response.status} ${response.statusText}`;
                    console.error("Ошибка ответа сервера при сохранении:", errorMessage, "Статус:", response.status);
                    throw new Error(errorMessage);                   
                }
                return JSON.parse(responseText);             
            })
            .then(result => {
                if (voiceStatusEditor) voiceStatusEditor.textContent = result.message || (isAutoSave ? "Автосохранение успешно." : "Данные успешно сохранены.");
                isTableDirty = false;                   
                if (saveTableButton) saveTableButton.classList.remove('dirty');
                console.log("Данные успешно сохранены, isTableDirty сброшен.");
                return result;                
            })
            .catch(error => {
                console.error('Ошибка сохранения данных (fetch или обработка ответа):', error);
                if (voiceStatusEditor) voiceStatusEditor.textContent = `Ошибка сохранения: ${error.message}`;
                throw error;                      
            })
            .finally(() => {
                if (saveTableButton) saveTableButton.disabled = false;
                currentSaveOperation = null;             
                if (voiceStatusEditor && (voiceStatusEditor.textContent.includes("Автосохранение...") || voiceStatusEditor.textContent.includes("Сохранение данных..."))) {
                    if (!isTableDirty) voiceStatusEditor.textContent = "Готов к командам...";
                }
            });
        return currentSaveOperation;
    }

    async function requestMicrophonePermissionAndInitSpeech() {
        if (navigator.mediaDevices && navigator.mediaDevices.getUserMedia) {
            try {
                await navigator.mediaDevices.getUserMedia({ audio: true });
                if (voiceStatusEditor) voiceStatusEditor.textContent = "Микрофон доступен. Голосовое управление готово (но не активно).";
                initializeSpeechRecognitionEngine();       
                if (voiceCommandButton) voiceCommandButton.disabled = false;                
            } catch (err) {
                console.error("Ошибка доступа к микрофону:", err);
                if (voiceStatusEditor) voiceStatusEditor.textContent = "Доступ к микрофону запрещен. Голосовое управление не работает.";
                if (voiceCommandButton) voiceCommandButton.disabled = true;
            }
        } else {
            if (voiceStatusEditor) voiceStatusEditor.textContent = "API микрофона не поддерживается. Голосовое управление не работает.";
            if (voiceCommandButton) voiceCommandButton.disabled = true;
        }
    }

    function initializeSpeechRecognitionEngine() {
        const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
        if (!SpeechRecognition) {
            if (voiceStatusEditor) voiceStatusEditor.textContent = 'Голосовой ввод не поддерживается.';
            if (voiceCommandButton) voiceCommandButton.style.display = 'none';
            return;
        }

        recognitionEditor = new SpeechRecognition();
        recognitionEditor.lang = 'ru-RU';
        recognitionEditor.continuous = true;        
        recognitionEditor.interimResults = true;          

        recognitionEditor.onstart = () => {
            isRecognizingEditor = true;       
            console.log("SpeechRecognition engine started.");
        };

        recognitionEditor.onresult = (event) => {
            lastSpeechTime = Date.now();
            let interimTranscript = '';
            let finalTranscriptSegment = '';

            for (let i = event.resultIndex; i < event.results.length; ++i) {
                if (event.results[i].isFinal) {
                    finalTranscriptSegment += event.results[i][0].transcript;
                } else {
                    interimTranscript += event.results[i][0].transcript;
                }
            }

            if (voiceStatusEditor) {
                if (interimTranscript) {
                    voiceStatusEditor.textContent = `Слышу: ${recognizedTextBuffer} ${interimTranscript}...`;
                } else if (finalTranscriptSegment) {
                    voiceStatusEditor.textContent = `Распознан сегмент: ${finalTranscriptSegment}`;
                }
            }


            if (finalTranscriptSegment) {
                recognizedTextBuffer += finalTranscriptSegment.toLowerCase().trim().replace(/[.,;:!?]$/, '') + " ";          
                console.log("Recognized segment (appended to buffer):", finalTranscriptSegment);
                console.log("Current buffer:", recognizedTextBuffer);

                if (!triggerDetected) {
                    for (const trigger of commandTriggers ) {
                        if (recognizedTextBuffer.includes(trigger + " ")) {                
                            triggerDetected = true;
                            if (voiceStatusEditor) voiceStatusEditor.textContent = `Обнаружен триггер: "${trigger}". Говорите команду...`;
                            console.log(`Trigger detected: "${trigger}". Buffer: "${recognizedTextBuffer}"`);
                            const triggerStartIndex = recognizedTextBuffer.lastIndexOf(trigger + " ");
                            recognizedTextBuffer = recognizedTextBuffer.substring(triggerStartIndex);
                            break;
                        }
                    }
                    for (const trigger of COLUMN_HEADER_TRIGGERS) {
                        if (recognizedTextBuffer.includes(trigger + " ")) {                
                            triggerDetected = true;
                            if (voiceStatusEditor) voiceStatusEditor.textContent = `Обнаружен триггер: "${trigger}". Говорите команду...`;
                            console.log(`Trigger detected: "${trigger}". Buffer: "${recognizedTextBuffer}"`);
                            const triggerStartIndex = recognizedTextBuffer.lastIndexOf(trigger + " ");
                            recognizedTextBuffer = recognizedTextBuffer.substring(triggerStartIndex);
                            break;
                        }
                    }
                }

                if (triggerDetected) {
                    clearTimeout(speechTimeoutId);
                    speechTimeoutId = setTimeout(() => {
                        console.log("Speech pause detected after trigger. Processing command from buffer:", recognizedTextBuffer);
                        if (voiceStatusEditor) voiceStatusEditor.textContent = `Обработка: "${recognizedTextBuffer.trim()}"`;
                        processCommand(recognizedTextBuffer.trim());
                        recognizedTextBuffer = "";       
                        triggerDetected = false;         
                    }, SPEECH_TIMEOUT_MS);
                }
            }
            if (!triggerDetected && recognizedTextBuffer.length > 150) {
                recognizedTextBuffer = recognizedTextBuffer.substring(recognizedTextBuffer.length - 100);
                console.log("Buffer trimmed (no trigger):", recognizedTextBuffer);
            }
        };

        recognitionEditor.onend = () => {
            isRecognizingEditor = false;
            console.log("SpeechRecognition engine stopped.");
            if (voiceControlActive) {
                console.log("Re-starting SpeechRecognition due to unexpected stop while voice control is active.");
                try {
                    if (recognitionEditor) recognitionEditor.start();
                } catch (e) {
                    console.error("Error re-starting SpeechRecognition:", e);
                }
            } else {
                if (voiceStatusEditor && voiceStatusEditor.textContent.startsWith('Слышу')) {
                    voiceStatusEditor.textContent = 'Голосовое управление выключено.';
                }
            }
        };

        recognitionEditor.onerror = (event) => {
            console.error("Ошибка SpeechRecognition:", event.error, event.message);
            let errorMessage = `Ошибка голоса: ${event.error}`;
            if (event.error === 'no-speech') errorMessage = 'Речь не распознана.';                
            if (event.error === 'not-allowed') {
                errorMessage = 'Доступ к микрофону запрещен.';
                voiceControlActive = false;             
                if (voiceCommandButton) {
                    voiceCommandButton.textContent = '🎤 Голос Выкл.';
                    voiceCommandButton.classList.remove('active');
                }
            }
            if (event.error === 'audio-capture') errorMessage = 'Ошибка захвата аудио.';

            if (voiceStatusEditor) voiceStatusEditor.textContent = errorMessage;

            if (event.error !== 'no-speech') {             
                isRecognizingEditor = false;
            }
        };
    }

    function toggleVoiceControl(forceState) {
        if (typeof forceState === 'boolean') {
            voiceControlActive = forceState;
        } else {
            voiceControlActive = !voiceControlActive;
        }

        if (voiceControlActive) {
            if (!recognitionEditor) {
                console.warn("Движок распознавания не инициализирован. Попытка инициализации...");
                requestMicrophonePermissionAndInitSpeech().then(() => {             
                    if (recognitionEditor && !isRecognizingEditor) {
                        try { recognitionEditor.start(); } catch (e) { console.error("Ошибка старта распознавания:", e); voiceControlActive = false; }
                    }
                });
                if (!recognitionEditor) {                   
                    voiceControlActive = false;          
                    if (voiceStatusEditor) voiceStatusEditor.textContent = "Не удалось инициализировать голосовое управление.";
                    return;
                }
            } else if (!isRecognizingEditor) {
                try {
                    recognitionEditor.start();
                } catch (e) {
                    console.error("Ошибка старта распознавания (уже инициализирован):", e);
                    voiceControlActive = false;
                }
            }
            if (voiceCommandButton) {
                voiceCommandButton.textContent = '🎤 Голос Вкл.';
                voiceCommandButton.classList.add('active');       
            }
            if (voiceStatusEditor) voiceStatusEditor.textContent = 'Голосовое управление включено. Говорите...';
            recognizedTextBuffer = "";             
            triggerDetected = false;
            lastSpeechTime = Date.now();


        } else {
            if (recognitionEditor && isRecognizingEditor) {
                recognitionEditor.stop();                            
            }
            clearTimeout(speechTimeoutId);
            triggerDetected = false;
            recognizedTextBuffer = "";
            if (voiceCommandButton) {
                voiceCommandButton.textContent = '🎤 Голос Выкл.';
                voiceCommandButton.classList.remove('active');
            }
            if (voiceStatusEditor) voiceStatusEditor.textContent = 'Голосовое управление выключено.';
        }
        console.log("Voice control active:", voiceControlActive);
    }


    if (voiceCommandButton) {
        voiceCommandButton.textContent = '🎤 Голос Выкл.';       
        voiceCommandButton.addEventListener('click', () => toggleVoiceControl());
    } else {
        console.warn("Кнопка voiceCommandButton не найдена");
    }

    async function autoSetDateTimeForNewRow(targetRow, targetCol, currentSheetIndex) {
        if (!isLuckysheetReady) return;

        const DATE_COLUMN_NAMES = ["дата", "date", "дата создания", "даты"];             
        let dateColumnIndex = -1;

        const allSheetFiles = luckysheet.getLuckysheetfile();
        if (!allSheetFiles || !allSheetFiles.length) return;

        const sheetFile = allSheetFiles.find(sheet => String(sheet.index) === String(currentSheetIndex));
        if (!sheetFile) return;

        const maxColsToCheck = sheetFile.column || (luckysheet.defaultConfig ? luckysheet.defaultConfig.columnlen : 26);
        for (let c = 0; c < maxColsToCheck; c++) {
            const headerValue = luckysheet.getCellValue(0, c, { sheetIndex: currentSheetIndex, type: 'm' });
            if (headerValue && DATE_COLUMN_NAMES.includes(String(headerValue).toLowerCase())) {
                dateColumnIndex = c;
                break;
            }
        }

        if (dateColumnIndex === -1) {
            console.log("Столбец 'Дата' не найден, автоматическая вставка даты не выполнена.");
            return;             
        }

        const dateCellCurrentValue = luckysheet.getCellValue(targetRow, dateColumnIndex, { sheetIndex: currentSheetIndex });

        if (dateCellCurrentValue === null || String(dateCellCurrentValue).trim() === "") {
            const now = new Date();
            const formattedDateTime = `${String(now.getFullYear()).padStart(1, '0')}.${String(now.getMonth()+1)}.${String(now.getDate()).padStart(1, '0')}`;

            console.log(`Автоматическая вставка даты в строку ${targetRow}, столбец 'Дата' (${dateColumnIndex}): ${formattedDateTime}`);
            luckysheet.setCellValue(targetRow, dateColumnIndex, formattedDateTime, { sheetIndex: currentSheetIndex });
        }
    }

    async function autoSetTimeForNewRow(targetRow, targetCol, currentSheetIndex) {
        if (!isLuckysheetReady) return;

        const TIME_COLUMN_NAMES = ["время", "time"];             
        let timeColumnIndex = -1;

        const allSheetFiles = luckysheet.getLuckysheetfile();
        if (!allSheetFiles || !allSheetFiles.length) return;

        const sheetFile = allSheetFiles.find(sheet => String(sheet.index) === String(currentSheetIndex));
        if (!sheetFile) return;

        const maxColsToCheck = sheetFile.column || (luckysheet.defaultConfig ? luckysheet.defaultConfig.columnlen : 26);
        for (let c = 0; c < maxColsToCheck; c++) {
            const headerValue = luckysheet.getCellValue(0, c, { sheetIndex: currentSheetIndex, type: 'm' });
            if (headerValue && TIME_COLUMN_NAMES.includes(String(headerValue).toLowerCase())) {
                timeColumnIndex = c;
                break;
            }   
        }

        if (timeColumnIndex === -1) {
            console.log("Столбец 'Время' не найден, автоматическая вставка даты не выполнена.");
            return;             
        }

        const timeCellCurrentValue = luckysheet.getCellValue(targetRow, timeColumnIndex, { sheetIndex: currentSheetIndex });

        if (timeCellCurrentValue === null || String(timeCellCurrentValue).trim() === "") {
            const now = new Date();
            const formattedTime = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}`;

            console.log(`Автоматическая вставка даты в строку ${targetRow}, столбец 'Дата' (${timeColumnIndex}): ${formattedTime}`);
            luckysheet.setCellValue(targetRow, timeColumnIndex, formattedTime, { sheetIndex: currentSheetIndex });
        }
    }

    async function processCommand(commandText) {
        if (!isLuckysheetReady) {       
            if (voiceStatusEditor) voiceStatusEditor.textContent = "Таблица еще не готова. Пожалуйста, подождите.";
            console.warn("processCommand вызван до полной готовности Luckysheet.");
            return;
        }
        const commandLower = commandText.toLowerCase().trim();
        if (voiceStatusEditor) voiceStatusEditor.textContent = `Обработка: "${commandText}"...`;

        if (typeof luckysheet === 'undefined'
            ) {

            if (voiceStatusEditor) voiceStatusEditor.textContent = "Таблица не готова для команд (API Luckysheet не доступно).";
            console.error("processCommand: luckysheet.functionTranslate или его методы не доступны.");
            if (typeof luckysheet !== 'undefined' && luckysheet.sheetmanage) {
                console.log("luckysheet.sheetmanage существует, но getGridSettings отсутствует или functionTranslate не найден.");
                console.log("Содержимое luckysheet.sheetmanage:", luckysheet.sheetmanage);
            } else if (typeof luckysheet !== 'undefined') {
                console.log("luckysheet существует, но luckysheet.sheetmanage или luckysheet.functionTranslate отсутствуют.");
                console.log("Содержимое luckysheet:", luckysheet);
            }
            isLuckysheetReady = false;                      
            return;
        }

        if (voiceStatusEditor) voiceStatusEditor.textContent = `Обработка: "${commandText}"...`;

        const cellToRowColFunc = luckysheet.cellToRowCol;
        const rowColToCellFunc = luckysheet.rowColToCell;
        const currentSheetObj = luckysheet.getSheet(); 
        if (!currentSheetObj || typeof currentSheetObj.index === 'undefined') {
            if (voiceStatusEditor) voiceStatusEditor.textContent = "Не удалось определить активный лист.";
            console.error("processCommand: Не удалось получить currentSheet или его index.");
            return;                      
        }
        const currentSheetIndex = currentSheetObj.index;

        let lastSelection = luckysheet.getRange();
        let activeRow = lastSelection && lastSelection.length > 0 ? lastSelection[0].row[0] : 0;
        let activeCol = lastSelection && lastSelection.length > 0 ? lastSelection[0].column[0] : 0;

        const currentSheetObject = luckysheet.getSheet();          
        if (!currentSheetObject || typeof currentSheetObject.index === 'undefined') {
            if (voiceStatusEditor) voiceStatusEditor.textContent = "Не удалось определить активный лист для навигации.";
            console.error("processCommand (навигация): Не удалось получить currentSheetObject или его index.");
            return;
        }
        const currentSheetIndexString = currentSheetObject.index;                      

        const allSheetFiles = luckysheet.getLuckysheetfile();
        if (!allSheetFiles || !Array.isArray(allSheetFiles)) {
            if (voiceStatusEditor) voiceStatusEditor.textContent = "Ошибка: не удалось получить данные книги.";
            console.error("processCommand (навигация): luckysheet.getLuckysheetfile() не вернул массив.");
            return;
        }

        const currentSheetFile = allSheetFiles.find(sheet => sheet.index === currentSheetIndexString);

        if (!currentSheetFile) {
            if (voiceStatusEditor) voiceStatusEditor.textContent = "Ошибка: не найдены данные для активного листа.";
            console.error(`processCommand (навигация): Не удалось найти лист с index="${currentSheetIndexString}" в luckysheet.getLuckysheetfile().`);
            console.log("Доступные листы:", allSheetFiles.map(s => ({ name: s.name, index: s.index })));
            return;
        }

        if (commandLower === "вверх" || commandLower === "вверх.") {
            if (activeRow > 0) activeRow--;
            luckysheet.setRangeShow({ row: [activeRow, activeRow], column: [activeCol, activeCol] });
            voiceStatusEditor.textContent = `Переход к ${rowColToCell(activeRow, activeCol)}.`;

        }
        if (commandLower === "вниз" || commandLower === "вниз.") {
            const totalRows = currentSheetFile.row || (luckysheet.defaultConfig ? luckysheet.defaultConfig.rowlen : 84);
            let newActiveRow = activeRow;
            if (activeRow < totalRows - 1) newActiveRow = activeRow + 1;
            else newActiveRow = activeRow;                                  

            luckysheet.setRangeShow({ row: [newActiveRow, newActiveRow], column: [activeCol, activeCol] });
            voiceStatusEditor.textContent = `Переход к ${rowColToCellFunc}.`;

            if (newActiveRow > activeRow && newActiveRow > 0) {
                await autoSetDateTimeForNewRow(newActiveRow, activeCol, currentSheetIndex);
                await autoSetTimeForNewRow(newActiveRow, activeCol, currentSheetIndex);
            }
            return;
        }
        if (commandLower === "влево" || commandLower === "влево.") {
            if (activeCol > 0) activeCol--;
            luckysheet.setRangeShow({ row: [activeRow, activeRow], column: [activeCol, activeCol] });
            voiceStatusEditor.textContent = `Переход к ${rowColToCell(activeRow, activeCol)}.`;
            return;
        }
        if (commandLower === "вправо" || commandLower === "вправо.") {
            const totalCols = currentSheetFile.column;                   
            if (typeof totalCols !== 'number' || totalCols <= 0) {
                console.warn(`Команда "вправо": не удалось определить totalCols для листа ${currentSheetFile.name}. Используется активный столбец + 10.`);
                if (activeCol < activeCol + 10) activeCol++;
            } else if (activeCol < totalCols - 1) {
                activeCol++;
            }
            luckysheet.setRangeShow({ row: [activeRow, activeRow], column: [activeCol, activeCol] });
            voiceStatusEditor.textContent = `Переход к ${rowColToCell(activeRow, activeCol)}.`;
            return;
        }

        if (commandLower === "отмена" || commandLower === "назад") {
            if (luckysheet.undo) {                
                luckysheet.undo();
                if (voiceStatusEditor) voiceStatusEditor.textContent = "Действие отменено.";
            } else {
                if (voiceStatusEditor) voiceStatusEditor.textContent = "Функция отмены недоступна.";
                console.warn("luckysheet.undo is not available.");
            }
            return;
        }
        if (commandLower === "повторить" || commandLower === "вперед") {
            if (luckysheet.redo) {
                luckysheet.redo();
                if (voiceStatusEditor) voiceStatusEditor.textContent = "Действие повторено.";
            } else {
                if (voiceStatusEditor) voiceStatusEditor.textContent = "Функция повтора недоступна.";
                console.warn("luckysheet.redo is not available.");
            }
            return;
        }

        const setHeaderMatch = commandLower.match(/^столбец\s+([a-z]+|\d+)\s+(.+)$/);
        if (setHeaderMatch) {
            const colRef = setHeaderMatch[1];
            let headerText = setHeaderMatch[2];
                headerText = headerText.trim().replace(/[.,;:!?]$/, '').trim();
            let col_idx = -1;    

            if (isNaN(parseInt(colRef))) {
                const addr = cellToRowCol(colRef + "1");          
                if (addr && typeof addr.c === 'number') {
                    col_idx = addr.c;
                }
            } else {
                col_idx = parseInt(colRef) - 1;                   
            }

            if (col_idx >= 0) {
                luckysheet.setCellValue(0, col_idx, headerText, { sheetIndex: currentSheetIndex });          
                voiceStatusEditor.textContent = `Заголовок '${headerText}' установлен для столбца ${colRef.toUpperCase()}.`;
                updateColumnHeadersAsTriggers();
            } else {
                voiceStatusEditor.textContent = `Некорректный столбец для заголовка: ${colRef}.`;
            }
            return;
        }


        const goToCellByHeaderMatch = commandLower.match(/^([а-яa-zё\s]+)$/);       
        const simpleNavAndActionCommands = ["вверх", "вниз", "влево", "вправо", "назад", "отмена", "повторить", "вперед", "сохранить", "сохранить таблицу", "охранить.", "записать", "поиск", "рассчитать", "очистить"];

        let isLikelyHeaderCommand = goToCellByHeaderMatch &&
            !simpleNavAndActionCommands.some(cmd => commandLower.startsWith(cmd)) &&                      
            commandLower.split(/\s+/).length <= 3;


        if (isLikelyHeaderCommand) {
            const potentialHeaderName = goToCellByHeaderMatch[1].trim();
            let foundColIdx = -1;    

            const allSheetFiles = luckysheet.getLuckysheetfile();
            let sheetFile = null;
            for (let i = 0; i < allSheetFiles.length; i++) {
                if (allSheetFiles[i].index === currentSheetIndex) {
                    sheetFile = allSheetFiles[i];
                    break;
                }
            }

            if (!sheetFile) {
                console.error("Не удалось получить данные текущего листа для поиска заголовка.");
                return;
            }

            const defaultColsToScan = (luckysheet.defaultConfig ? luckysheet.defaultConfig.columnlen : 26)
            for (let c = 0; c < (sheetFile.column || luckysheet.defaultColNum || 26); c++) {
                const cellValue = luckysheet.getCellValue(0, c, { sheetIndex: currentSheetIndex, type: 'm' });
                if (cellValue && cellValue.toLowerCase() === potentialHeaderName.toLowerCase()) {
                    foundColIdx = c;
                    break;
                }
            }

            if (foundColIdx !== -1) {
                let targetRow = -1;
                let lastNonEmptyRowInCol = 0;                               
                const maxRowsToCheck = sheetFile.row || luckysheet.defaultRowNum || 84;

                for (let r = 1; r < maxRowsToCheck; r++) {                
                    const cellValue = luckysheet.getCellValue(r, foundColIdx, { sheetIndex: currentSheetIndex });
                    if (cellValue !== null && String(cellValue).trim() !== "") {
                        lastNonEmptyRowInCol = r;             
                    }
                }

                targetRow = lastNonEmptyRowInCol + 1;

                if (targetRow >= maxRowsToCheck) {
                    console.log(`Целевая строка ${targetRow} может быть за пределами текущего количества строк (${maxRowsToCheck}).`);
                }
                if (targetRow > 0) {                

                    await autoSetDateTimeForNewRow(targetRow, foundColIdx, currentSheetIndexString);
                    await autoSetTimeForNewRow(targetRow, foundColIdx, currentSheetIndexString);
                }

                luckysheet.setRangeShow({ row: [targetRow, targetRow], column: [foundColIdx, foundColIdx] });
                luckysheet.scroll({ targetRow: targetRow, targetCol: foundColIdx });
                voiceStatusEditor.textContent = `Переход к последней свободной ячейке (${rowColToCell(targetRow, foundColIdx)}) в столбце "${potentialHeaderName}".`;
                return;

            } else {
                console.log(`Столбец с заголовком "${potentialHeaderName}" не найден, команда будет обработана дальше.`);
            }
        }

        const writeMatch = commandLower.match(/^записать\s+(.+?)(?:\s+в\s+([a-z]+\d+))?$/);
        if (writeMatch) {
            const value = writeMatch[1];
            let cellAddress = writeMatch[2];
            let r_idx, c_idx;

            if (cellAddress) {
                const addr = cellToRowCol(cellAddress.toUpperCase());
                if (addr && typeof addr.r === 'number' && typeof addr.c === 'number') { r_idx = addr.r; c_idx = addr.c; }
                else { if (voiceStatusEditor) voiceStatusEditor.textContent = `Неверный формат адреса: ${cellAddress}`; return; }
            } else {
                lastSelection = luckysheet.getRange();
                if (lastSelection && lastSelection.length > 0) { r_idx = lastSelection[0].row[0]; c_idx = lastSelection[0].column[0]; }
                else { if (voiceStatusEditor) voiceStatusEditor.textContent = "Нет активной ячейки для записи."; return; }
            }
            luckysheet.setCellValue(r_idx, c_idx, value.startsWith('=') ? value : value);
            if (voiceStatusEditor) voiceStatusEditor.textContent = `Записано "${value}" в ${rowColToCell(r_idx, c_idx)}.`;
            return;
        }

        const searchMatch = commandLower.match(/^поиск\s+(.+)$/);
        if (searchMatch) {
            let searchTerm = searchMatch[1].trim();
            searchTerm = searchTerm.replace(/[.,;:!?]$/, '').trim();

            if (!searchTerm) {
                voiceStatusEditor.textContent = "Укажите текст для поиска.";
                return;
            }
            let found = false;
            if (currentSheetFile && (currentSheetFile.data || currentSheetFile.celldata)) {
                const celldataToSearch = currentSheetFile.celldata || [].concat(...currentSheetFile.data.map((row, r) => row.map((cell, c) => ({ r, c, v: cell }))));

                const sortedCelldata = [...celldataToSearch].sort((a, b) => {
                    if (a.r !== b.r) return a.r - b.r;
                    return a.c - b.c;
                });

                for (const cell of sortedCelldata) {
                    if (!cell || typeof cell.r !== 'number' || typeof cell.c !== 'number') continue;          

                    const cellValueObj = cell.v;
                    let cellText = "";
                    if (cellValueObj) {                      
                        if (typeof cellValueObj.m === 'string') cellText = cellValueObj.m;
                        else if (typeof cellValueObj.v !== 'undefined' && cellValueObj.v !== null) cellText = String(cellValueObj.v);
                        else if (typeof cellValueObj === 'string' || typeof cellValueObj === 'number') cellText = String(cellValueObj);
                    }

                    if (cellText.toLowerCase().includes(searchTerm.toLowerCase())) {
                        luckysheet.setRangeShow({ row: [cell.r, cell.r], column: [cell.c, cell.c] });
                        luckysheet.scroll({ targetRow: cell.r, targetCol: cell.c });
                        voiceStatusEditor.textContent = `Найдено "${searchTerm}" в ${rowColToCell(cell.r, cell.c)}.`;
                        found = true;
                        break;
                    }
                }
            }
            if (!found) {
                voiceStatusEditor.textContent = `Текст "${searchTerm}" не найден на текущем листе.`;
            }
            return;
        }


        const sumMatch = commandLower.match(/^рассчитать\s+сумму\s+([a-zа-яё\s\d]+)$/i);                
        if (sumMatch) {
            const headerNameToFind = sumMatch[1].trim().replace(/[.,;:!?]$/, '').trim().toLowerCase();
            let target_col_idx = -1;    

            if (!currentSheetFile) {
                console.error("Ошибка в 'рассчитать сумму': currentSheetFile не определен.");
                if (voiceStatusEditor) voiceStatusEditor.textContent = "Ошибка: данные активного листа недоступны.";
                return;
            }

            console.log(`Сумма: поиск столбца по заголовку '${headerNameToFind}'`);
            const defaultColsToScan = (luckysheet.defaultConfig ? luckysheet.defaultConfig.columnlen : 26);
            for (let c = 0; c < (currentSheetFile.column || defaultColsToScan); c++) {
                const headerVal = luckysheet.getCellValue(0, c, { sheetIndex: currentSheetIndexString, type: 'm' });          
                if (headerVal && headerVal.trim().toLowerCase() === headerNameToFind) {
                    target_col_idx = c;
                    console.log(`Сумма: столбец найден по заголовку '${headerVal}' -> индекс ${target_col_idx}`);
                    break;
                }
            }

            if (target_col_idx === -1) {
                voiceStatusEditor.textContent = `Столбец с заголовком "${sumMatch[1].trim()}" для суммирования не найден.`;
                return;
            }

            let calculatedSum = 0;
            let hasNumbersToSum = false;
            const totalRowsInSheet = currentSheetFile.row || (luckysheet.defaultConfig ? luckysheet.defaultConfig.rowlen : 84);

            for (let r = 1; r < totalRowsInSheet; r++) {                
                const cellRawValue = luckysheet.getCellValue(r, target_col_idx, { sheetIndex: currentSheetIndexString });
                if (cellRawValue !== null && String(cellRawValue).trim() !== "") {
                    const numValue = Number(cellRawValue);
                    if (!isNaN(numValue)) {
                        calculatedSum += numValue;
                        hasNumbersToSum = true;
                    }
                }
            }
            console.log(`Сумма: результат суммирования для столбца ${target_col_idx} = ${calculatedSum}. Были ли числа: ${hasNumbersToSum}`);

            let first_empty_row_idx = -1;
            for (let r = 1; r < totalRowsInSheet; r++) {          
                const cellRawValue = luckysheet.getCellValue(r, target_col_idx, { sheetIndex: currentSheetIndexString });
                if (cellRawValue === null || String(cellRawValue).trim() === "") {
                    first_empty_row_idx = r;
                    break;             
                }
            }

            if (first_empty_row_idx === -1) {
                first_empty_row_idx = totalRowsInSheet;                
            }
            console.log(`Сумма: первая пустая ячейка для вставки суммы найдена в строке с индексом ${first_empty_row_idx}`);

            if (first_empty_row_idx >= totalRowsInSheet) {
                console.log(`Сумма: нужно добавить строку для вставки суммы на позицию ${first_empty_row_idx}. Текущих строк: ${totalRowsInSheet}`);
                try {
                    luckysheet.insertRow(first_empty_row_idx, 1);
                    console.log(`Сумма: добавлена 1 строка на позицию ${first_empty_row_idx}`);
                } catch (e) {
                    console.error("Ошибка при попытке вставить строку для суммы:", e);
                    voiceStatusEditor.textContent = "Ошибка при добавлении строки для вставки суммы.";
                    return;
                }
            }

            console.log(`Сумма: вставка значения ${calculatedSum} в ячейку (${first_empty_row_idx}, ${target_col_idx})`);
            luckysheet.setCellValue(first_empty_row_idx, target_col_idx, calculatedSum, { sheetIndex: currentSheetIndexString });

            const targetCellAddress = rowColToCell(first_empty_row_idx, target_col_idx);
            voiceStatusEditor.textContent = `Сумма ${calculatedSum} вставлена в ${targetCellAddress} (столбец "${headerNameToFind}").`;
            isTableDirty = true;
            if (saveTableButton) saveTableButton.classList.add('dirty');

            luckysheet.setRangeShow({ row: [first_empty_row_idx, first_empty_row_idx], column: [target_col_idx, target_col_idx] });
            luckysheet.scroll({ targetRow: first_empty_row_idx, targetCol: target_col_idx });
            return;
        }

        if (commandLower.startsWith("сохранить") || commandLower === "охранить.") {
            saveTableDataToServer(false).catch(err => console.warn("Ручное сохранение не удалось:", err));
            return;
        }

        if (voiceStatusEditor) voiceStatusEditor.textContent = `Команда "${commandText}" не распознана или не реализована.`;
    }
    function executeManualCommand() {                
        if (manualCommandInput && typeof manualCommandInput.value === 'string') {
            const commandText = manualCommandInput.value.trim();
            if (commandText) {
                processCommand(commandText);
                manualCommandInput.value = '';
            }
        }
    }

    if (executeManualCommandButton) executeManualCommandButton.addEventListener('click', executeManualCommand);
    if (manualCommandInput) manualCommandInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); executeManualCommand(); } });
    if (saveTableButton) {
        saveTableButton.addEventListener('click', () => {
            console.log("Кнопка 'Сохранить таблицу' нажата.");
            saveTableDataToServer(false)             
                .then(result => console.log("Результат ручного сохранения:", result))
                .catch(err => console.warn("Ручное сохранение не удалось из-за ошибки:", err));
        });
    }


    const AUTOSAVE_INTERVAL = 10000;                      
    setInterval(async () => {
        if (isTableDirty && typeof luckysheet !== 'undefined' && luckysheet.getAllSheets && currentTableFilename) {
            console.log(`Проверка для автосохранения (isTableDirty: ${isTableDirty})...`);
            saveTableDataToServer(true)             
                .then(result => {
                    if (result && result.message !== "Нет изменений для автосохранения.") {
                        console.log("Автосохранение успешно:", result.message);
                    }
                })
                .catch(err => console.error("Ошибка автосохранения:", err));
        }
    }, AUTOSAVE_INTERVAL);

    window.addEventListener('beforeunload', (event) => {
        if (isTableDirty) {
            event.preventDefault();
            event.returnValue = 'У вас есть несохраненные изменения. Вы уверены, что хотите уйти?';
        }
    });

    if (typeof currentTableFilename !== 'undefined' && currentTableFilename) {
        initializeLuckysheet();
    } else {
        console.error("Переменная currentTableFilename не определена в HTML перед загрузкой скрипта table_editor_luckysheet.js");
        if (voiceStatusEditor) voiceStatusEditor.textContent = "Ошибка: Не удалось определить имя файла таблицы.";
        if (saveTableButton) saveTableButton.disabled = true;
        if (voiceCommandButton) voiceCommandButton.disabled = true;
    }
});