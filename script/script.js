$(document).ready(function() {
    let excelData = [];  // Массив для хранения строк данных

    // Добавим в таблицу заголовки
    const headers = ["УИД Контрагент", "Контрагент.Код", "Контрагент", "УИД Конечный получатель", "Конечный получатель код",
     "Конечный получатель", "УИД Продуктовое направление", "Продуктовое направление", "УИД Номенклатурная группа", "Номенклатурная группа",
      "УИД Номенклатура", "Номенклатура.Артикул", "Номенклатура", "Номенклатура.Архивная", "УИД Единица измерения", "Единица измерения", "Цена с НДС за единицу измерения в валюте", "Валюта"];
    excelData.push(headers);
    let dropdownGroupIndex = 1; // Индекс для создания уникальных идентификаторов для каждой группы полей

    $('.select2').select2();
    $('#namePage').show(); // Показать начальную страницу
    let userNameGlobal = ''; // Глобальная переменная для хранения имени пользователя
    $('#startBtn').on('click', startProcess); // Отработка кнопки Старт
    $('#userName').on('keydown', function(e) { // Отработка нажатия Enter
        if (e.keyCode === 13) {
            startProcess();
        }
    });

    // Отработка кнопки Старт
    function startProcess() {
        let userName = $('#userName').val().charAt(0).toUpperCase() + $('#userName').val().slice(1).toLowerCase(); // Можно написать любым регистром
        if (!userName) {
            alert("Введите ваше имя");
            return;
        }
        userNameGlobal = userName; // Присваиваем значение глобальной переменной

        if (agents[userName]) { // Имя найдено в базе
            let products = agents[userName]; // Все контрагенты
            // Добавляем только один начальный набор полей при нажатии кнопки "Старт"
            $('#dropdownGroupWrapper').empty();  // Удаляем все предыдущие элементы, если они есть
            addDropdownGroup(products, dropdownGroupIndex); // Добавление группы полей

            $('#namePage').hide(); // Скрываем первую страницу
            $('#dropdownPage').show(); // Показываем вторую страницу
        } else {
            alert("Имя не найдено в базе");
        }
    }

    // Функция добавления группы полей по индексу
    function addDropdownGroup(products, index) {
        $('#dropdownGroupWrapper').append(`
            <div class="dropdown-container" id="dropdownGroup${index}" style="display: flex; align-items: center; gap: 15px;">
                <div class="dropdown-element" style="margin-right: 10px;">
                    <select class="select2" id="dropdown1-${index}">
                        <option value="">Контрагент</option>
                    </select>
                </div>
                <div class="dropdown-element" style="margin-right: 10px;">
                    <select class="select2" id="dropdown3-${index}">
                        <option value="">Конечник</option>
                    </select>
                </div>
                <div class="dropdown-element" style="margin-right: 10px;">
                    <select class="select2" id="dropdownCategory-${index}">
                        <option value="">Группа</option>
                    </select>
                </div>
                <div class="dropdown-element" style="margin-right: 10px;">
                    <select class="select2" id="dropdownGood-${index}" disabled multiple>
                        <option value="all">Выбрать все</option>
                        <option value="">Номенклатура</option>
                    </select>
                </div>
                <button class="add-btn" id="addBtn-${index}" style="padding: 10px 20px; font-family: 'DIN Pro', sans-serif; color: white; background-color: #5C4B9A; border: none; cursor: pointer;">
                    Добавить
                </button>
                <img class="success-icon" id="successIcon-${index}" src="img/success.png" alt="Success">
            </div>
        `);

        // Заполнение первого, второго и третьего выпадающих списков
        updateDropdown1Options(products, index);
        updateDropdown3Options(products, index);
        updateDropdownCategoryOptions(index);

        // Реинициализация select2 для новых элементов
        $(`#dropdown1-${index}, #dropdown3-${index}, #dropdownCategory-${index}, #dropdownGood-${index}`).select2();

        // Назначение события для каждой кнопки "Добавить"
        $(`#addBtn-${index}`).on('click', function() {
            handleAddButton(index);
        });
    }

    // Первый список
    function updateDropdown1Options(products, index) {
        products.forEach(function(product) {
            $(`#dropdown1-${index}`).append('<option value="' + product[0] + '">' + product[0] + '</option>');
        });
    }

    // Второй список
    function updateDropdown3Options(products, index) {
        products.forEach(function(product) {
            $(`#dropdown3-${index}`).append('<option value="' + product[0] + '">' + product[0] + '</option>');
        });
    }

    // Третий список
    function updateDropdownCategoryOptions(index) {
        Object.keys(goods).forEach(function(category) { // Заполняем категории ключами словаря goods
            $(`#dropdownCategory-${index}`).append('<option value="' + category + '">' + category + '</option>');
        });

        // Заполнение четвертого списка в зависимости от категории
        $(`#dropdownCategory-${index}`).on('change', function() {
            let selectedCategory = $(this).val(); 
            if (selectedCategory) {
                $(`#dropdownGood-${index}`).prop('disabled', false); // Открываем ячейку
                updateDropdownGoodOptions(selectedCategory, index); // Заполняем соответствующим списком
            } else {
                // Если категория не выбрана, то пустая и заблокированная
                $(`#dropdownGood-${index}`).prop('disabled', true).empty().append('<option value="">Номенклатура</option>');
            }
        });
    }

    // Четвертый список
    function updateDropdownGoodOptions(category, index) {
        $(`#dropdownGood-${index}`).empty().append('<option value="all">Выбрать все</option>'); // Возможность выбрать все номенклатуры

        // Заполняем номенклатурами по соответствующей группе
        goods[category].forEach(function(good) {
            let encodedGood = encodeURIComponent(good[0]); // Чтобы не было проблем с кавычками в названии, иначе не работало
            $(`#dropdownGood-${index}`).append('<option value="' + encodedGood + '">' + good[0] + '</option>');
        });

        $(`#dropdownGood-${index}`).select2();

        // Выбор номенклатуры
        $(`#dropdownGood-${index}`).on('change', function() {
            let selectedOptions = $(this).val();
            if (selectedOptions.includes('all')) { // Если выбрали все, то не отображаем весь список номенклатур
                $(`#dropdownGood-${index}`).val(['all']);
            }
        });
    }

    // Отработка каждой кнопки Добавить
    function handleAddButton(index) {
        let product1Index = $(`#dropdown1-${index} option:selected`).val(); // Контрагент
        let product2Index = $(`#dropdown3-${index} option:selected`).val(); // Конечник
        let category = $(`#dropdownCategory-${index} option:selected`).text(); // Категория номенклатур
        let selectedGoods = $(`#dropdownGood-${index}`).val(); // Номенклатуры

        // Пропусков быть не может
        if (!product1Index || !product2Index || !category || !selectedGoods || selectedGoods.length === 0) {
            alert("Пожалуйста, выберите данные для всех полей.");
            return;
        }

        // Находим в словаре пользователя контрагента и конечника
        let agent1 = agents[userNameGlobal].find(arr => arr[0] === product1Index)
        let agent2 = agents[userNameGlobal].find(arr => arr[0] === product2Index)

        // Если выбрали все, проходимся в словаре по ключу категории и берем все номенклатуры со свойствами в верном порядке
        if (selectedGoods.includes('all')) {
            goods[category].forEach(function(good) {
                let rowData = [
                    agent1[2], agent1[1], product1Index,
                    agent2[2], agent2[1], product2Index,
                    good[1], good[1], good[2], good[3], good[4], good[5], good[0], good[6], good[7], good[8], good[9], good[10]
                ];
                excelData.push(rowData);
            });
        } else { // Иначе для каждой выбранной номенклатуры находим ее в словаре и берем со свойствами в верном порядке
            selectedGoods.forEach(function(encodedGoodIndex) {
                let goodIndex = decodeURIComponent(encodedGoodIndex);
                let good = goods[category].find(arr => arr[0] === goodIndex);
                let rowData = [
                    agent1[2], agent1[1], product1Index,
                    agent2[2], agent2[1], product2Index,
                    good[1], good[1], good[2], good[3], good[4], good[5], good[0], good[6], good[7], good[8], good[9], good[10]
                ];
                excelData.push(rowData);
            });
        }

        // Показать галочку после успешного добавления данных
        $(`#successIcon-${index}`).show();
        dropdownGroupIndex++; // Обновляем Index после успешного нажатия кнопки Далее
        addDropdownGroup(agents[userNameGlobal], dropdownGroupIndex); // Новые ячейки для выбора
    }

    // Отработка кнопки Скачать
    $('#downloadBtn').on('click', function() {
        if (excelData.length === 0) {
        alert("Нет данных для скачивания");
        return;
    }

    // С помощью библиотеки XLSX создаем таблицу и добавляем в нее данные
    let wb = XLSX.utils.book_new();
    let ws = XLSX.utils.aoa_to_sheet(excelData);
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, `${userNameGlobal}.xlsx`);  // Называем таблицу фамилией продавца
    });
});