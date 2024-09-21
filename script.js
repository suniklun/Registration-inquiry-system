document.addEventListener('DOMContentLoaded', loadRecords);

document.getElementById('registrationForm').addEventListener('submit', function (event) {
    event.preventDefault();

    const name = document.getElementById('name').value;
    const dateInput = document.getElementById('date').value;
    const date = new Date(dateInput);
    const timeSlot = document.getElementById('timeSlot').value;
    const hospitalSelect = document.getElementById('hospital');
    let hospital;

    if (hospitalSelect.value === '其他醫院') {
        hospital = document.getElementById('otherHospital').value;
        if (!hospital || !/^[\u4e00-\u9fa5]+$/.test(hospital)) {
            alert("請填寫其他醫院名稱（中文）");
            return;
        }
    } else {
        hospital = hospitalSelect.value;
    }

    const department = document.getElementById('department').value;
    const doctor = document.getElementById('doctor').value;

    if (!/^[\u4e00-\u9fa5]+$/.test(department)) {
        alert("請填寫科別（中文）");
        return;
    }

    if (!/^[\u4e00-\u9fa5]+$/.test(doctor)) {
        alert("請填寫醫師名稱（中文）");
        return;
    }

    const visitNumber = document.getElementById('visitNumber').value;
    const weekDay = ['日', '一', '二', '三', '四', '五', '六'][date.getDay()];

    const record = { name, date: `${dateInput} (${weekDay})`, dateValue: dateInput, timeSlot, hospital, department, doctor, visitNumber };

    saveRecord(record);
    loadRecords(); // 重新載入資料以顯示更新後的表格

    this.reset();
    document.getElementById('otherHospital').style.display = 'none';
});

document.getElementById('hospital').addEventListener('change', function () {
    const otherHospitalInput = document.getElementById('otherHospital');
    otherHospitalInput.style.display = this.value === '其他醫院' ? 'block' : 'none';
    otherHospitalInput.value = '';
});

document.getElementById('deleteAll').addEventListener('click', function () {
    localStorage.removeItem('records');
    loadRecords();
});

document.getElementById('exportExcel').addEventListener('click', function () {
    const records = JSON.parse(localStorage.getItem('records')) || [];

    const worksheetData = records.map(record => [
        record.name,
        record.date,
        record.timeSlot,
        record.hospital,
        record.department,
        record.doctor,
        record.visitNumber
    ]);

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([["姓名", "日期", "時段", "醫院", "科別", "醫師", "看診號"], ...worksheetData]);

    XLSX.utils.book_append_sheet(workbook, worksheet, "掛號查詢");
    XLSX.writeFile(workbook, '掛號查詢.xlsx');
});

function saveRecord(record) {
    let records = JSON.parse(localStorage.getItem('records')) || [];
    records.push(record);
    // 依據看診日期排序
    records.sort((a, b) => new Date(a.dateValue) - new Date(b.dateValue));
    localStorage.setItem('records', JSON.stringify(records));
}

function loadRecords() {
    const tableBody = document.querySelector('#recordsTable tbody');
    const records = JSON.parse(localStorage.getItem('records')) || [];

    tableBody.innerHTML = '';
    // 依據看診日期排序
    records.forEach(record => {
        addRecordToTable(record);
    });
}

function addRecordToTable(record) {
    const tableBody = document.querySelector('#recordsTable tbody');
    const row = tableBody.insertRow();

    row.insertCell(0).innerText = record.name;
    row.insertCell(1).innerText = record.date;
    row.insertCell(2).innerText = record.timeSlot;
    row.insertCell(3).innerText = record.hospital;
    row.insertCell(4).innerText = record.department;
    row.insertCell(5).innerText = record.doctor;
    row.insertCell(6).innerText = record.visitNumber;
    row.insertCell(7).innerHTML = `
        <button onclick="editRow(this)">編輯</button>
        <button onclick="deleteRow(this)">刪除</button>
    `;
}

function editRow(button) {
    const row = button.parentNode.parentNode;
    const name = row.cells[0].innerText;
    const date = row.cells[1].innerText.split(' (')[0];
    const timeSlot = row.cells[2].innerText;
    const hospital = row.cells[3].innerText;
    const department = row.cells[4].innerText;
    const doctor = row.cells[5].innerText;
    const visitNumber = row.cells[6].innerText;

    document.getElementById('name').value = name;
    document.getElementById('date').value = date;
    document.getElementById('timeSlot').value = timeSlot;
    document.getElementById('hospital').value = hospital === '其他醫院' ? '其他醫院' : hospital;
    document.getElementById('otherHospital').value = hospital === '其他醫院' ? hospital : '';
    document.getElementById('department').value = department;
    document.getElementById('doctor').value = doctor;
    document.getElementById('visitNumber').value = visitNumber;

    deleteRow(button); // 刪除原有記錄
}

function deleteRow(button) {
    const row = button.parentNode.parentNode;
    const name = row.cells[0].innerText;

    let records = JSON.parse(localStorage.getItem('records')) || [];
    records = records.filter(record => record.name !== name);
    localStorage.setItem('records', JSON.stringify(records));

    row.parentNode.removeChild(row);
}
