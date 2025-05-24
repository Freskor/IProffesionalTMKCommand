document.addEventListener('DOMContentLoaded', () => {
    const startButton = document.getElementById('start-button');
    const createNewButton = document.getElementById('create-new-button');
    const openExistingButton = document.getElementById('open-existing-button');
    const creationStatus = document.getElementById('creation-status');             
    const fileListContainer = document.getElementById('file-list-container');


    const views = {
        initial: document.getElementById('initial-view'),
        actionSelection: document.getElementById('action-selection-view'),
        status: document.getElementById('status-view')       
    };

    const BACKEND_URL = '';

    function showView(viewName) {
        Object.values(views).forEach(view => view.classList.remove('active'));
        if (views[viewName]) {
            views[viewName].classList.add('active');
        }
    }

    startButton.addEventListener('click', () => {
        showView('actionSelection');
    });

    createNewButton.addEventListener('click', async () => {
        showView('status');
        creationStatus.textContent = 'Создание нового файла таблицы...';
        createNewButton.disabled = true;
        openExistingButton.disabled = true;

        try {
            const response = await fetch(`${BACKEND_URL}/api/tables`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({}),                   
            });

            if (!response.ok) {
                let errorData;
                try { errorData = await response.json(); } catch (e) { }
                const errorMessage = errorData?.error || response.statusText || `HTTP error ${response.status}`;
                throw new Error(`Ошибка сервера: ${errorMessage}`);
            }

            const result = await response.json();
            creationStatus.textContent = `Таблица "${result.filename}" создана. Перенаправление...`;
            window.location.href = result.editor_url;             

        } catch (error) {
            console.error('Ошибка при создании таблицы:', error);
            alert(`Не удалось создать таблицу: ${error.message}`);
            creationStatus.textContent = `Ошибка: ${error.message}`;
            showView('actionSelection');             
        } finally {
            createNewButton.disabled = false;
            openExistingButton.disabled = false;
        }
    });

    openExistingButton.addEventListener('click', async () => {
        showView('status');                            
        creationStatus.textContent = 'Загрузка списка файлов...';
        fileListContainer.innerHTML = '';          

        try {
            const response = await fetch(`${BACKEND_URL}/api/files`);
            if (!response.ok) {
                throw new Error(`Ошибка сервера: ${response.status} ${response.statusText}`);
            }
            const files = await response.json();

            if (files && files.length > 0) {
                let html = '<h4>Выберите файл для открытия:</h4><ul>';
                files.forEach(file => {
                    html += `<li><a href="/table-editor/${file}">${file}</a> ( <a href="/download/${file}" download>скачать</a> )</li>`;
                });
                html += '</ul>';
                fileListContainer.innerHTML = html;
                creationStatus.textContent = `Загружено ${files.length} файлов.`;
                document.getElementById('action-selection-view').appendChild(fileListContainer);          
                showView('actionSelection');


            } else {
                fileListContainer.innerHTML = '<p>На сервере нет доступных файлов.</p>';
                creationStatus.textContent = 'Файлы не найдены.';
                document.getElementById('action-selection-view').appendChild(fileListContainer);
                showView('actionSelection');
            }
        } catch (error) {
            console.error('Ошибка при получении списка файлов:', error);
            fileListContainer.innerHTML = `<p>Не удалось загрузить список файлов: ${error.message}</p>`;
            creationStatus.textContent = 'Ошибка загрузки списка файлов.';
            document.getElementById('action-selection-view').appendChild(fileListContainer);
            showView('actionSelection');
        }
    });

    showView('initial');
});