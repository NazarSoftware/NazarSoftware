// Получаем элементы
const fileDropArea = document.getElementById("fileDropArea");
const fileInput = document.getElementById("fileInput");
const fileAttributesContainer = document.getElementById(
  "fileAttributesContainer"
);
const form = document.getElementById("applicationForm");

// Нажатие на область для выбора файла
fileDropArea.addEventListener("click", function () {
  fileInput.click();
});

// Предотвращаем поведение по умолчанию для перетаскивания файлов
["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
  fileDropArea.addEventListener(eventName, preventDefaults, false);
});

function preventDefaults(e) {
  e.preventDefault();
  e.stopPropagation();
}

// Обработка события перетаскивания файлов
fileDropArea.addEventListener("drop", handleDrop, false);

function handleDrop(e) {
  const files = e.dataTransfer.files;
  handleFiles(files);
}

// Обработка выбора файлов через input
fileInput.addEventListener("change", function () {
  const files = fileInput.files;
  handleFiles(files);
});

function handleFiles(files) {
  Array.from(files).forEach((file) => {
    // Добавляем атрибуты только для DXF-файлов
    if (file.name.endsWith(".dxf")) {
      createFileAttributes(file);
    }
  });
}

function createFileAttributes(file) {
  // Создание контейнера для атрибутов каждого файла
  const fileAttributesDiv = document.createElement("div");
  fileAttributesDiv.classList.add("file-attributes");

  // Название файла
  const fileName = document.createElement("h3");
  fileName.textContent = file.name;
  fileAttributesDiv.appendChild(fileName);

  // Поле "Материал"
  const materialLabel = document.createElement("label");
  materialLabel.textContent = "Материал:";
  const materialSelect = document.createElement("select");
  materialSelect.innerHTML = `
        <option value="steel">Сталь</option>
        <option value="aluminium">Алюминий</option>
        <option value="stainless_steel">Нержавеющая сталь</option>
        
    `;
  fileAttributesDiv.appendChild(materialLabel);
  fileAttributesDiv.appendChild(materialSelect);

  // Поле "Толщина"
  const thicknessLabel = document.createElement("label");
  thicknessLabel.textContent = "Толщина:";
  const thicknessSelect = document.createElement("select");
  thicknessSelect.innerHTML = `
        <option value="0.7">0.7 мм</option>
        <option value="1.0">1.0 мм</option>
        <option value="1.5">1.5 мм</option>
        <option value="2.0">2.0 мм</option>
        <option value="3.0">3.0 мм</option>
        <option value="4.0">4.0 мм</option>
        <option value="5.0">5.0 мм</option>
        <option value="6.0">6.0 мм</option>
        <option value="8.0">8.0 мм</option>
        <option value="10.0">10.0 мм</option>
        <option value="12.0">12.0 мм</option>
        <option value="14.0">14.0 мм</option>
    `;
  fileAttributesDiv.appendChild(thicknessLabel);
  fileAttributesDiv.appendChild(thicknessSelect);

  // Поле "Количество"
  const quantityLabel = document.createElement("label");
  quantityLabel.textContent = "Количество:";
  const quantityInput = document.createElement("input");
  quantityInput.type = "number";
  quantityInput.min = "1";
  quantityInput.value = "1";
  fileAttributesDiv.appendChild(quantityLabel);
  fileAttributesDiv.appendChild(quantityInput);

  // Добавляем созданные атрибуты в контейнер
  fileAttributesContainer.appendChild(fileAttributesDiv);
}

// Отправка формы
form.addEventListener("submit", function (e) {
  e.preventDefault();

  // Собираем данные для генерации Excel файла
  const client = document.getElementById("client").value;
  const name = document.getElementById("name").value;
  const phone = document.getElementById("phone").value;
  const email = document.getElementById("email").value;

  // Создаем массив для хранения данных
  const data = [];

  // Добавляем первую строку с общими данными только в первой строке
  data.push([
    "Заказчик",
    "ФИО",
    "Телефон",
    "Почта",
    "Имя DXF",
    "Материал",
    "Толщина",
    "Количество",
  ]);
  data.push([client, name, phone, email, "", "", "", ""]); // Заполним только первую строку общими данными

  // Добавляем строки для каждого DXF файла и его атрибутов начиная с 10 строки
  const fileAttributesDivs = document.querySelectorAll(".file-attributes");
  fileAttributesDivs.forEach((div) => {
    const fileName = div.querySelector("h3").textContent;
    const material = div.querySelector("select:nth-of-type(1)").value;
    const thickness = div.querySelector("select:nth-of-type(2)").value;
    const quantity = div.querySelector("input").value;

    data.push(["", "", "", "", fileName, material, thickness, quantity]);
  });

  // Создаем файл Excel и загружаем его
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, "Заявка");

  // Устанавливаем имя файла как в поле "Заказчик"
  const fileName = `${client}.xlsx`;
  XLSX.writeFile(wb, fileName);
});
