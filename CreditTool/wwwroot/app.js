const rateTableBody = document.querySelector('#rate-table tbody');
const scheduleTableBody = document.querySelector('#schedule-table tbody');
const importForm = document.getElementById('import-form');
const importStatus = document.getElementById('import-status');
const actionStatus = document.getElementById('action-status');
const totalInterestElement = document.getElementById('total-interest');

function addRateRow(rate) {
    const row = document.createElement('tr');
    row.innerHTML = `
        <td><input type="date" class="date-from" value="${rate?.dateFrom ?? ''}" required></td>
        <td><input type="date" class="date-to" value="${rate?.dateTo ?? ''}" required></td>
        <td><input type="number" step="0.01" class="rate-value" value="${rate?.rate ?? 0}" required></td>
        <td><button type="button" class="secondary remove-rate">Usuń</button></td>
    `;
    row.querySelector('.remove-rate').addEventListener('click', () => {
        row.remove();
    });
    rateTableBody.appendChild(row);
}

document.getElementById('add-rate').addEventListener('click', () => addRateRow());

function readParametersFromForm() {
    return {
        netValue: parseFloat(document.getElementById('net-value').value) || 0,
        marginRate: parseFloat(document.getElementById('margin-rate').value) || 0,
        paymentFrequency: document.getElementById('payment-frequency').value,
        paymentDay: document.getElementById('payment-day').value,
        creditStartDate: document.getElementById('credit-start').value,
        creditEndDate: document.getElementById('credit-end').value,
        dayCountBasis: document.getElementById('day-count').value,
        roundingMode: document.getElementById('rounding-mode').value,
        roundingDecimals: parseInt(document.getElementById('rounding-decimals').value || '2', 10),
        processingFeeRate: parseFloat(document.getElementById('processing-fee').value) || 0,
        bulletRepayment: document.getElementById('bullet-repayment').value === 'true'
    };
}

function setParametersToForm(parameters) {
    document.getElementById('net-value').value = parameters.netValue ?? '';
    document.getElementById('margin-rate').value = parameters.marginRate ?? 0;
    document.getElementById('payment-frequency').value = parameters.paymentFrequency ?? 'Monthly';
    document.getElementById('payment-day').value = parameters.paymentDay ?? 'LastOfMonth';
    document.getElementById('credit-start').value = parameters.creditStartDate?.substring(0, 10) ?? '';
    document.getElementById('credit-end').value = parameters.creditEndDate?.substring(0, 10) ?? '';
    document.getElementById('day-count').value = parameters.dayCountBasis ?? 'Actual365';
    document.getElementById('rounding-mode').value = parameters.roundingMode ?? 'Bankers';
    document.getElementById('rounding-decimals').value = parameters.roundingDecimals ?? 2;
    document.getElementById('processing-fee').value = parameters.processingFeeRate ?? 0;
    document.getElementById('bullet-repayment').value = parameters.bulletRepayment ? 'true' : 'false';
}

function readRatesFromTable() {
    const rows = rateTableBody.querySelectorAll('tr');
    return Array.from(rows).map(row => ({
        dateFrom: row.querySelector('.date-from').value,
        dateTo: row.querySelector('.date-to').value,
        rate: parseFloat(row.querySelector('.rate-value').value) || 0
    }));
}

function populateRateTable(rates) {
    rateTableBody.innerHTML = '';
    (rates ?? []).forEach(rate => addRateRow(rate));
    if (rateTableBody.children.length === 0) {
        addRateRow();
    }
}

function updateTotalInterest(totalInterest) {
    totalInterestElement.textContent = (totalInterest ?? 0).toFixed(2);
}

function displaySchedule(schedule, totalInterest) {
    scheduleTableBody.innerHTML = '';
    (schedule ?? []).forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item.paymentDate?.substring(0, 10)}</td>
            <td>${item.daysInPeriod}</td>
            <td>${item.interestRate.toFixed(4)}%</td>
            <td>${item.interestAmount.toFixed(2)}</td>
            <td>${item.principalPayment.toFixed(2)}</td>
            <td>${item.totalPayment.toFixed(2)}</td>
            <td>${item.remainingPrincipal.toFixed(2)}</td>
        `;
        scheduleTableBody.appendChild(row);
    });

    const calculatedTotal = totalInterest ?? (schedule ?? []).reduce((sum, item) => sum + (item.interestAmount ?? 0), 0);
    updateTotalInterest(calculatedTotal);
}

function buildPayload() {
    return {
        parameters: readParametersFromForm(),
        rates: readRatesFromTable()
    };
}

importForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    importStatus.textContent = '';
    const fileInput = document.getElementById('import-file');
    if (!fileInput.files.length) {
        importStatus.textContent = 'Wybierz plik Excel do importu.';
        importStatus.className = 'status error';
        return;
    }

    const formData = new FormData();
    formData.append('file', fileInput.files[0]);

    try {
        const response = await fetch('/api/import', { method: 'POST', body: formData });
        if (!response.ok) {
            throw new Error(await response.text());
        }
        const data = await response.json();
        setParametersToForm(data.parameters);
        populateRateTable(data.rates);
        importStatus.textContent = 'Parametry zostały zaimportowane.';
        importStatus.className = 'status success';
    } catch (error) {
        importStatus.textContent = `Import nieudany: ${error.message}`;
        importStatus.className = 'status error';
    }
});

document.getElementById('calculate').addEventListener('click', async () => {
    actionStatus.textContent = 'Trwa obliczanie...';
    actionStatus.className = 'status';
    try {
        const payload = buildPayload();
        const response = await fetch('/api/calculate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        if (!response.ok) {
            throw new Error(await response.text());
        }
        const result = await response.json();
        displaySchedule(result.schedule, result.totalInterest);
        actionStatus.textContent = 'Harmonogram został obliczony.';
        actionStatus.className = 'status success';
    } catch (error) {
        actionStatus.textContent = `Błąd obliczeń: ${error.message}`;
        actionStatus.className = 'status error';
        updateTotalInterest(0);
    }
});

document.getElementById('export').addEventListener('click', async () => {
    actionStatus.textContent = 'Trwa eksport...';
    actionStatus.className = 'status';
    try {
        const payload = buildPayload();
        const format = document.getElementById('export-format').value;
        const response = await fetch(`/api/export?format=${encodeURIComponent(format)}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        if (!response.ok) {
            throw new Error(await response.text());
        }
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = format === 'word' ? 'harmonogram.docx' : 'harmonogram.xlsx';
        link.click();
        window.URL.revokeObjectURL(url);
        actionStatus.textContent = 'Eksport zakończony powodzeniem.';
        actionStatus.className = 'status success';
    } catch (error) {
        actionStatus.textContent = `Eksport nieudany: ${error.message}`;
        actionStatus.className = 'status error';
    }
});

// Initialize with one blank rate row and set default dates
populateRateTable();
const today = new Date().toISOString().substring(0, 10);
const inSixMonths = new Date();
inSixMonths.setMonth(inSixMonths.getMonth() + 6);
document.getElementById('credit-start').value = today;
document.getElementById('credit-end').value = inSixMonths.toISOString().substring(0, 10);
