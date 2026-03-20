const startDatePicker = document.getElementById('startDate');
const endDatePicker = document.getElementById('endDate');

document.getElementById("fileInput").addEventListener("change", async function() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    const filespan = document.getElementById('filespan');
    const fileButton = document.getElementById('fileButton');
    const loader = document.getElementById('loader');
    const loaderspan = document.getElementById('loaderspan');
    const comments = document.getElementById('comments');
    const reloadButton = document.getElementById('reloadButton');

    filespan.innerText = file.name;
    fileButton.setAttribute("disabled", "true");

    loader.style.display = "block";
    loaderspan.style.display = "block";

    if (!file) {
        alert('Please select a file.');
        return;
    }

    const reader = new FileReader();

    reader.onload = async function(e) {
        document.getElementById('startDate').removeAttribute("disabled");
        var data = e.target.result;
        var workbook = XLSX.read(data, {type: 'binary'});
        var sheetName = workbook.SheetNames[0];
        var sheet = workbook.Sheets[sheetName];
        var html = XLSX.utils.sheet_to_html(sheet);
        comments.innerHTML = html;

        filterRows();

        loader.style.display = "none";
        loaderspan.style.display = "none";
        reloadButton.style.display = "block";
    };
    reader.readAsBinaryString(file);
});

async function filterRows() {
    const table = document.querySelector('.comments table');
    const rows = table.querySelectorAll('tr');
    const datePattern = /^\d{2}\.\d{2}\.\d{4}$/;

    const cellsToRemove = [];
    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        let dateFound = false;
        cells.forEach(cell => {
            if (datePattern.test(cell.textContent)) {
                dateFound = true;
            }
            if (cell.innerText.length < 1) {
                cellsToRemove.push(cell);
            }
        });
        if (!dateFound) {
            row.remove();
        }
    });
    cellsToRemove.forEach(cell => cell.remove());
}

let startDate;
let endDate;

document.getElementById('startDate').addEventListener("change", function(e){
    this.setAttribute("disabled", true);
    document.getElementById('endDate').removeAttribute("disabled");
    startDate = parseDate(startDatePicker.value);

    document.getElementById('endDate').min = startDatePicker.value;
    document.getElementById('endDate').max = (()=>{let d=new Date(startDatePicker.value); d.setDate(d.getDate()+21); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`})();


    const table = document.querySelector('.comments table');
    const rows = table.querySelectorAll('tr');
    const datePattern = /^\d{2}\.\d{2}\.\d{4}$/;

    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        let dateFound = false;
        cells.forEach(cell => {
            if (datePattern.test(cell.textContent)) {
                dateFound = true;
                const compareDate = parseDate2(cell.textContent);
                if (compareDate < startDate) {
                    if (startDate.getMonth() === 11) {
                        if (compareDate.getMonth() !== 0) {
                            row.remove();
                        }
                    }
                    else {
                        row.remove();
                    }
                }
            }
        });
    });
    sortTableByDate()
});
document.getElementById('endDate').addEventListener("change", function(e){
    this.setAttribute("disabled", true);
    endDate = parseDate(endDatePicker.value);

    const table = document.querySelector('.comments table');
    const rows = table.querySelectorAll('tr');
    const datePattern = /^\d{2}\.\d{2}\.\d{4}$/;

    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        let dateFound = false;
        cells.forEach(cell => {
            if (datePattern.test(cell.textContent)) {
                dateFound = true;
                const compareDate = parseDate2(cell.textContent);
                if (startDate.getMonth() === 11 && endDate.getMonth() === 11) {
                    if (compareDate.getMonth() === 0 || compareDate > endDate) {
                        row.remove();
                    }
                }
                else if (startDate.getMonth() === 11 && endDate.getMonth() === 0) {
                    if (compareDate > endDate && compareDate.getMonth() === 0) {
                        row.remove();
                    }
                }
                else if (compareDate > endDate) {
                    row.remove();
                }
            }
        });
    });
    groupAndLogByDate();
});

function sortTableByDate() {
    const table = document.querySelector('table');
    const rows = Array.from(table.querySelectorAll('tr'));
    
    rows.sort((rowA, rowB) => {
        let dateA, dateB;
        
        if (startDate.getMonth() === 11) {
            // New Year exception: parse with year adjustment
            dateA = findAndParseDate2(rowA);
            dateB = findAndParseDate2(rowB);
            
            // Adjust: if month is December, use previous year (2000)
            if (dateA.getMonth() === 11) {
                dateA.setFullYear(2000);
            }
            if (dateB.getMonth() === 11) {
                dateB.setFullYear(2000);
            }
        }
        else {
            dateA = findAndParseDate(rowA);
            dateB = findAndParseDate(rowB);
        }
        
        return dateA - dateB;
    });
    
    rows.forEach(row => table.appendChild(row));
}

function findAndParseDate(row) {
    for (let cell of row.cells) {
        const dateMatch = cell.textContent.match(/(\d{2})\.(\d{2})\.\d{4}/);
        if (dateMatch) {
            const [_, day, month] = dateMatch;
            return new Date(2000, month - 1, day);
        }
    }
    return new Date(0);
}
function findAndParseDate2(row) {
    for (let cell of row.cells) {
        const dateMatch = cell.textContent.match(/(\d{2})\.(\d{2})\.\d{4}/);
        if (dateMatch) {
            const [_, day, month] = dateMatch;
            return new Date(2001, month - 1, day);
        }
    }
    return new Date(0);
}

async function groupAndLogByDate() {
    const rows = document.querySelectorAll('table tr');
    const dateNamesArr = [];

    rows.forEach(row => {
        const firstTd = row.querySelector('td:last-child').innerText.trim().substring(0, 5);
        const secondTd = row.querySelector('td:nth-child(2)').innerText.trim();
        const fifthTd = `${row.querySelector('td:nth-child(5)').innerHTML.trim()} ${row.querySelector('td:nth-child(6)').innerHTML.trim()}`;

        const dateName = {
            date: firstTd,
            name: secondTd,
            city: determineCity(fifthTd)
        };
        dateNamesArr.push(dateName);
    });
    result = Object.groupBy(dateNamesArr, ({ date }) => date);
    const dates = Object.keys(result);
    const firstDate = dates.length > 0 ? dates[0] : undefined;
    const lastDate = dates.length > 0 ? dates[dates.length - 1] : undefined;
    let output = '';

    comments.innerHTML = `
        <div class="viberDiv">
            <button class="downloadviberPic">
                <span class="imageIcon">↓</span>
                <span class="imageText"> Завантажити картинку для Viber</span>
            </button>
            <button class="downloadviberPic" id="sendToTelegram">
                <span class="sendToTgIcon">➤</span>
                <span class="imageText"> Надіслати текст і картинку</span>
            </button>
        </div>
        <br><br><br>
        <div class="forViberImage">
            <img src="birthdayImg.png" />
            <span class="inForViberImageText">${firstDate} – ${lastDate}</span>
        </div>
        <br><br><br>
    `;

    for (const [date, names] of Object.entries(result)) {
        output += `🎉 <b>*${date}*</b><br>`;
        names.forEach(item => {
            output += `${item.name}, ${item.city}<br><br>`;
        });
    }
    comments.innerHTML += output;

    comments.innerHTML += `
        <div class="imagesDiv">
            <button class="downloadImages">
                <span class="imageIcon">↓</span>
                <span class="imageText"> Завантажити картинки</span>
            </button>
            <button class="downloadImages" id="sendToGmail">
                <span class="sendToTgIcon">➤</span>
                <span class="imageText"> Надіслати картинки</span>
            </button>
        </div>
        <br><br><br>
        <div class="imagesContainer" id="imagesContainer"></div>
    `;

    const container = document.getElementById("imagesContainer");

    for (const { date, name } of dateNamesArr) {
    const div = document.createElement("div");
    div.className = "forImage";
    div.innerHTML += `
        <img class="personalImg" src="birthdayPersonal.png" />
        <div class="dateDiv">
        <span class="dateInImage">${date}</span>
        </div>
    `;

    const vocative = await toVocative(name.split(" ")[1]);
    div.innerHTML += `
        <span class="nameInImage">${vocative}, вітаємо</span>
    `;

    container.appendChild(div);
    }

comments.innerHTML += `
        <div class=topDiv>
            <button class="sendDiv" id="sendRequestButton">
                <span class="infoScale">Надіслати</span>
                <span class="sendSign">➤</span>
            </button>
            <button class="downloadDiv" onclick="captureDivAsImage()">
                <span class="infoScale">Завантажити</span>
                <span class="downloadSign">↓</span>
            </button>
            <div class="sliderDiv">
                <span class="infoScale">Наблизити / віддалити</span>
                <input type="range" id="scaleSlider" min="0.1" max="1" step="0.1" value="0.4" />
            <div>
        </div>    
    `;
const outerDiv = document.createElement('div');
outerDiv.className = 'outerMailDiv';
comments.appendChild(outerDiv);
outerDiv.innerHTML = `<img src="birthdayMail.png"/>`

const scaleSlider = document.getElementById('scaleSlider');
scaleSlider.addEventListener('input', (e) => {
  const scaleValue = e.target.value;
  outerDiv.style.transform = `scale(${scaleValue})`;
});


let id = 1;
let spacingId = 1;

Object.entries(result).forEach(([date, names]) => {
  
    let div = document.createElement('div');
    div.className = 'dateBox';
    div.id = `dateBox${id}`;
    
    if (id === 1) {
        div.style.left = `${140}px`;
        div.style.top = `${465}px`;
    }
    if (id === 2) {
        div.style.left = `${790}px`;
        div.style.top = `${535}px`;
    }
    if (id === 3) {
        div.style.left = `${75}px`;
        div.style.top = `${document.getElementById('dateBox1').getBoundingClientRect().height/0.4 + 465 + 190}px`;
    }
    if (id === 4) {
        div.style.left = `${750}px`;
        div.style.top = `${document.getElementById('dateBox2').getBoundingClientRect().height/0.4 + 535 + 190}px`;
    }
    if (id === 5) {
        div.style.left = `${110}px`;
        div.style.top = `${document.getElementById('dateBox1').getBoundingClientRect().height/0.4 + 465 + 190 + document.getElementById('dateBox3').getBoundingClientRect().height/0.4 + 190}px`;
    }



    div.innerHTML = `
        <div class="dateStick" id="dateStick${id}">
            <img src="dateStickImg.png" />
            <span class="dateSpan">
                <b>${date}</b>
            </span>
        </div>
        <br>
    `;
    
    names.forEach(item => {
        div.innerHTML += `
            <span class="nameSpan" id="nameSpan${spacingId}" contenteditable="true">
                <b class="employeeInfo">${item.name.split(' ').slice(0, 2).join(' ')}<br>${item.name.split(' ').slice(2, 3)}, ${item.city}</b>
                <div class="spacingHoverDiv" contenteditable="false">
                    <div class="outerDivForPadding">
                        <span class="infoSpan">Інтервал букв.</span>
                        <div class="spacingDiv" id="spacingDiv${spacingId}">
                            <span class="arrowSpan">←</span>
                            <span class="paddingSpan" id="spacingValue${spacingId}">0px</span>
                            <span class="arrowSpan">→</span>
                        </div>
                    </div>
                </div>
            </span>
            <br><br>
        `;
        
        spacingId += 1;
    });

    const hoverDiv = document.createElement('div');
    hoverDiv.className = 'hoverOnDateBox';
    div.appendChild(hoverDiv);
    outerDiv.appendChild(div);

    const spacingDivs = div.querySelectorAll('.spacingDiv');
    spacingDivs.forEach((spacingDiv) => {
        spacingDiv.addEventListener('mousedown', (event) => {
            const parentSpan = spacingDiv.closest('.nameSpan');
            if (!parentSpan) return;

            let currentSpacing = parseFloat(window.getComputedStyle(parentSpan).letterSpacing) || 0;
            let startX = event.clientX;

            const mouseMoveHandler = (moveEvent) => {
                const deltaX = moveEvent.clientX - startX;
                startX = moveEvent.clientX;

                // Adjust spacing dynamically based on mouse movement
                currentSpacing += deltaX * 0.1; // Sensitivity adjustment
                parentSpan.style.letterSpacing = `${currentSpacing}px`;
                spacingDiv.style.letterSpacing = `0px`;
                spacingDiv.closest('.outerDivForPadding').querySelector('.infoSpan').style.letterSpacing = `0px`;

                // Update the display value
                const spacingValueElement = spacingDiv.querySelector('.paddingSpan');
                if (spacingValueElement) {
                    spacingValueElement.textContent = `${currentSpacing.toFixed(1)}px`;
                }
            };

            const mouseUpHandler = () => {
                document.removeEventListener('mousemove', mouseMoveHandler);
                document.removeEventListener('mouseup', mouseUpHandler);
            };

            document.addEventListener('mousemove', mouseMoveHandler);
            document.addEventListener('mouseup', mouseUpHandler);
        });
    });

    const dateStickElement = document.querySelector(`#dateStick${id}`);
    if (!dateStickElement) {
        console.error(`Element with id dateStick${id} not found`);
        return;
    }

    const autoPaddingElements = div.querySelectorAll(".employeeInfo");
    let arrayOfWidth = [];
    
    autoPaddingElements.forEach((element) => {
        arrayOfWidth.push(element.offsetWidth);
    });

    autoPaddingValue = (500 - Math.max(...arrayOfWidth)) / 2;
    div.style.paddingLeft = `${autoPaddingValue}px`;
    div.querySelector('.dateStick').style.marginLeft = `-${autoPaddingValue}px`;



    setTimeout(() => {
        const initialPaddingLeft = parseInt(window.getComputedStyle(div).paddingLeft, 10) || 0;
        const initialSize = parseInt(window.getComputedStyle(div).width, 10) || 0;
        hoverDiv.innerHTML = `
            <div class="outerDivForPadding">
                <span class="infoSpan">Відступ</span>
                <div class="paddingDiv" id="paddingDiv${id}">
                    <span class="arrowSpan">←</span>
                    <span class="paddingSpan" id="paddingValue${id}">${initialPaddingLeft}px</span>
                    <span class="arrowSpan">→</span>
                </div>
            </div>
            <div class="outerDivForPadding">
                <span class="infoSpan">Перетягти</span>
                <div class="dragDiv" id="dragDiv${id}">
                    <span class="dotsSpan">⋯</span>
                </div>
            </div>
            <div class="outerDivForPadding">
                <span class="infoSpan">Ширина</span>
                <div class="sizeDiv" id="sizeDiv${id}">
                    <span class="arrowSpan">←</span>
                    <span class="paddingSpan" id="sizeValue${id}">${initialSize}px</span>
                    <span class="arrowSpan">→</span>
                </div>
            </div>
        `;

        let isDraggingPadding = false;
        let isDraggingSize = false;
        let startX = 0;

        // Padding Adjustment
        const paddingDiv = hoverDiv.querySelector(`#paddingDiv${id}`);
        paddingDiv.addEventListener('mousedown', (event) => {
            isDraggingPadding = true;
            startX = event.clientX;
        });

        // Width Adjustment
        const sizeDiv = hoverDiv.querySelector(`#sizeDiv${id}`);
        sizeDiv.addEventListener('mousedown', (event) => {
            isDraggingSize = true;
            startX = event.clientX;
        });

        document.addEventListener('mousemove', (event) => {
            if (isDraggingPadding) {
                const deltaX = event.clientX - startX;
                startX = event.clientX;
                const currentPadding = parseInt(window.getComputedStyle(div).paddingLeft, 10) || 0;
                const newPadding = Math.max(currentPadding + deltaX * 5, 0);
                div.style.paddingLeft = `${newPadding}px`;
                dateStickElement.style.marginLeft = `${-newPadding}px`;
                const paddingValueElement = hoverDiv.querySelector(`#paddingValue${id}`);
                if (paddingValueElement) {
                    paddingValueElement.innerText = `${newPadding}px`;
                }
            }

            if (isDraggingSize) {
                const deltaX = event.clientX - startX;
                startX = event.clientX;
                const currentWidth = parseInt(window.getComputedStyle(div).width, 10) || 0;
                const newWidth = Math.max(currentWidth + deltaX * 5, 50); // Minimum width: 50px
                div.style.width = `${newWidth}px`;
                const sizeValueElement = hoverDiv.querySelector(`#sizeValue${id}`);
                if (sizeValueElement) {
                    sizeValueElement.innerText = `${newWidth}px`;
                }
            }
        });

        document.addEventListener('mouseup', () => {
            isDraggingPadding = false;
            isDraggingSize = false;
        });
    }, 0);

    id += 1;
});

// Implement drag-and-drop functionality using the dragDiv to move the boxes within the outerDiv
outerDiv.addEventListener('mousedown', (e) => {
    const dragElement = e.target.closest('.dragDiv'); // Detect the dragDiv
    if (!dragElement) return;

    const dateBox = dragElement.closest('.dateBox'); // Target the associated dateBox
    if (!dateBox) return;

    // Account for transform scaling
    const scale = 0.3;

    // Calculate offsets for the drag event
    const outerRect = outerDiv.getBoundingClientRect();
    const dateBoxRect = dateBox.getBoundingClientRect();
    let initialX = e.clientX;
    let initialY = e.clientY;

    // Adjust for scaled offset
    const offsetX = (initialX - dateBoxRect.left) / scale;
    const offsetY = (initialY - dateBoxRect.top) / scale;

    const moveElement = (moveEvent) => {
        // Calculate new position adjusted for scaling
        const deltaX = (moveEvent.clientX - initialX) / scale;
        const deltaY = (moveEvent.clientY - initialY) / scale;

        const left = dateBox.offsetLeft + deltaX;
        const top = dateBox.offsetTop + deltaY;

        // Constrain the movement to the bounds of the outerDiv
        const constrainedLeft = Math.max(
            0,
            Math.min(left, (outerRect.width - dateBoxRect.width) / 0.3)
        );
        const constrainedTop = Math.max(
            0,
            Math.min(top, (outerRect.height - dateBoxRect.height) / 0.3)
        );

        // Apply the constrained position
        dateBox.style.left = `${constrainedLeft}px`;
        dateBox.style.top = `${constrainedTop}px`;

        // Update initial positions for smooth dragging
        initialX = moveEvent.clientX;
        initialY = moveEvent.clientY;
    };

    const stopMoving = () => {
        document.removeEventListener('mousemove', moveElement);
        document.removeEventListener('mouseup', stopMoving);
    };

    document.addEventListener('mousemove', moveElement);
    document.addEventListener('mouseup', stopMoving);
});

document.querySelector('.downloadviberPic').addEventListener("click", () => {
    const tempCanvas = document.createElement("canvas");
    tempCanvas.width = 1080;
    tempCanvas.height = 1080;
    const ctx = tempCanvas.getContext("2d");
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.src = "birthdayImg.png";
    img.onload = function() {
        ctx.drawImage(img, 0, 0, 1080, 1080);
        ctx.font = "29pt Montserrat";
        ctx.fillStyle = "white";
        ctx.fillText(`${firstDate} – ${lastDate}`, 112, 575);
        const downloadLink = document.createElement("a");
        downloadLink.href = tempCanvas.toDataURL("image/png");
        downloadLink.download = "birthdayImage.png";
        downloadLink.click();
    }
})

document.querySelector('.downloadImages').addEventListener("click", async () => {
    const zoomElement = document.querySelector('.imagesContainer');
    zoomElement.style.zoom = "1"; // Set to original size before capturing

    const imagesArray = document.querySelectorAll('.forImage');

    for (const image of imagesArray) {
        await html2canvas(image).then((canvas) => {
            const dataUrl = canvas.toDataURL('image/png');
            const link = document.createElement('a');
            link.href = dataUrl;
            link.download = 'happyBirthday.png';
            link.click();
        });
    }

    zoomElement.style.zoom = ".3"; // Restore zoom after all images are captured
})

document.getElementById('sendToTelegram').addEventListener("click", () => {
    const sendMessageToTelegram = (chatId, message) => {
        const token = '8078380194:AAGVbHR9PHyRyfv9U2PoDgwQ8GCjR_RDZ5Y';
        const url = `https://api.telegram.org/bot${token}/sendMessage`;
    
        fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                chat_id: chatId,
                text: message,
                parse_mode: 'HTML'  // Указываем, что будем использовать HTML форматирование
            })
        })
        .then(response => response.json())
        .then(data => console.log(data))  // Логируем ответ от API
        .catch(error => console.error('Ошибка:', error));  // Логируем ошибку
    };
    
    // Пример использования:
    const chatId = 473127028;  // ID чата
    const message = output.replaceAll('<br>', '\n');;
    
    sendMessageToTelegram(chatId, message);

    const sendImageToTelegram = (chatId, imageData) => {
        const token = '8078380194:AAGVbHR9PHyRyfv9U2PoDgwQ8GCjR_RDZ5Y';
        const url = `https://api.telegram.org/bot${token}/sendDocument`;

        const formData = new FormData();
        formData.append('chat_id', chatId);
        formData.append('document', dataURItoBlob(imageData), 'birthdayImage.png');  // Преобразуем base64 в файл

        fetch(url, {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => console.log(data))  // Логируем ответ от API
        .catch(error => console.error('Ошибка:', error));  // Логируем ошибку
    };

    // Преобразуем base64 строку в Blob
    function dataURItoBlob(dataURI) {
        const byteString = atob(dataURI.split(',')[1]);
        const mimeString = dataURI.split(',')[0].split(':')[1].split(';')[0];
        const arrayBuffer = new ArrayBuffer(byteString.length);
        const uintArray = new Uint8Array(arrayBuffer);

        for (let i = 0; i < byteString.length; i++) {
            uintArray[i] = byteString.charCodeAt(i);
        }

        return new Blob([uintArray], { type: mimeString });
    }

    // Создаем картинку на холсте
    const tempCanvas = document.createElement("canvas");
    tempCanvas.width = 1080;
    tempCanvas.height = 1080;
    const ctx = tempCanvas.getContext("2d");
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.src = "birthdayImg.png";

    img.onload = function() {
        ctx.drawImage(img, 0, 0, 1080, 1080);
        ctx.font = "29pt Montserrat";
        ctx.fillStyle = "white";
        ctx.fillText(`${firstDate} – ${lastDate}`, 112, 575);

        // Получаем изображение как base64
        const imageData = tempCanvas.toDataURL("image/png");

        // Отправляем изображение
        const chatId = 473127028;  // ID чата
        sendImageToTelegram(chatId, imageData);
    };
});

document.getElementById('sendToGmail').addEventListener("click", () => {
    
    const zoomElement = document.querySelector('.imagesContainer');
    zoomElement.style.zoom = "1";
    const divs = document.querySelectorAll('.forImage');
    
    // Создаем массив промисов для всех операций html2canvas
    const capturePromises = Array.from(divs).map(div => 
        html2canvas(div).then(canvas => canvas.toDataURL('image/png'))
    );
    
    // Ждем завершения всех промисов перед отправкой
    Promise.all(capturePromises)
        .then(images => {
            const emailData = {
                email: "A.Val.Smirnov@metinvestholding.com",
                subject: `Листівки-привітання для колег із Днем народження ${parseDate3(startDatePicker.value)}-${parseDate3(endDatePicker.value)}`,
                body: `<p>${output}p>`,
                images: images
            };
            
            return fetch("https://script.google.com/macros/s/AKfycbzLk3RVMd2ZCRJ1Uu1e3A8KeglNKouSq4ocHIOBPUDN3-GMUEfrA1AjElGqrIYnqLo_Pw/exec", {
                method: 'POST',
                mode: 'no-cors',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(emailData)
            });
        })
        .then(() => {
            console.log("Запрос отправлен");
            alert("Письмо отправлено!");
        })
        .catch(error => {
            console.error("Ошибка:", error);
            alert("Ошибка при отправке письма!");
        });

        zoomElement.style.zoom = ".3";
});
    document.getElementById('sendRequestButton').addEventListener("click", () => {
        const scaleSlider = document.getElementById('scaleSlider');
        const currentValue = scaleSlider.value;
        scaleSlider.value = "1";
        scaleSlider.dispatchEvent(new Event('input'));

        html2canvas(document.querySelector('.outerMailDiv')).then((canvas) => {
            const dataUrl = canvas.toDataURL('image/png');
            const emailData = {
                email: "A.Val.Smirnov@metinvestholding.com",
                subject: `Добірка Днів народження колег ${parseDate3(startDatePicker.value)}-${parseDate3(endDatePicker.value)}`,
                body: `<img src="${dataUrl}" />`
            };
            fetch("https://script.google.com/macros/s/AKfycbxHhvgJytBN-DIGeQ9tsfOnC1afaz6v-X_o_EPHOFYdnzYx2J8EeG75PP3Rd7g7EVoUVA/exec", {
                method: 'POST',
                mode: 'no-cors',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(emailData)
            })
            .then(() => {
                console.log("Запрос отправлен");
                alert("Письмо отправлено!");
            })
            .catch(error => {
                console.error("Ошибка:", error);
                alert("Ошибка при отправке письма!");
            });
        });
        scaleSlider.value = currentValue;
        scaleSlider.dispatchEvent(new Event('input'));
    })
}


function captureDivAsImage() {
    const scaleSlider = document.getElementById('scaleSlider');
    const currentValue = scaleSlider.value;
    scaleSlider.value = "1";
    scaleSlider.dispatchEvent(new Event('input'));
    html2canvas(document.querySelector('.outerMailDiv')).then((canvas) => {
        const dataUrl = canvas.toDataURL('image/png');
        const link = document.createElement('a');
        link.href = dataUrl;
        link.download = 'birthdays.png';
        link.click();
    });
    scaleSlider.value = currentValue;
    scaleSlider.dispatchEvent(new Event('input'));
}

function determineCity(cityString) {
    const cities = [
        "Дніпро", "Одеса", "Миколаїв", "Кременчук", "Вінниця", "Бровари", 
        "Кривий Ріг", "Харків", "Святопетрівське", "Львів", "Тернопіль", 
        "Хмельницький", "Брошнів-Осада", "Київ"
    ];
    
    for (let city of cities) {
        if (cityString.includes(city)) {
            return city;
        }
    }
    
    return "Київ";
}

function parseDate(dateString) {
    const [year, month, day] = dateString.split('-');
    return new Date(2000, month - 1, day);
}
function parseDate2(dateString) {
    const [day, month, year] = dateString.split('.');
    return new Date(2000, month - 1, day);
}
function parseDate3(dateString) {
    const [year, month, day] = dateString.split('-');
    const date = new Date(year, month - 1, day);

    const dayFormatted = String(date.getDate()).padStart(2, '0');
    const monthFormatted = String(date.getMonth() + 1).padStart(2, '0');

    return `${dayFormatted}.${monthFormatted}`;
}
async function toVocative(name) {
  const input = { givenName: name };
  input.gender = await shevchenko.detectGender(input) || 'masculine';
  const result = await shevchenko.inVocative(input);
  return result.givenName;
}