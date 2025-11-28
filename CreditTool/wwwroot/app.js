const rateTableBody = document.querySelector('#rate-table tbody');
const scheduleTableBody = document.querySelector('#schedule-table tbody');
const importForm = document.getElementById('import-form');
const importStatus = document.getElementById('import-status');
const actionStatus = document.getElementById('action-status');
const exportLogButton = document.getElementById('export-log');
const totalInterestElement = document.getElementById('total-interest');
const aprElement = document.getElementById('apr');
const interestApplicationSelect = document.getElementById('interest-application');
const creditStartInput = document.getElementById('credit-start');
const creditEndInput = document.getElementById('credit-end');

const DAY_IN_MS = 24 * 60 * 60 * 1000;

function parseDateInput(value) {
    if (!value) {
        return null;
    }
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? null : date;
}

function formatDateInput(date) {
    return date ? date.toISOString().substring(0, 10) : '';
}

function addDays(date, days) {
    const copy = new Date(date);
    copy.setDate(copy.getDate() + days);
    return copy;
}

function getCreditDates() {
    return {
        startDate: parseDateInput(creditStartInput.value),
        endDate: parseDateInput(creditEndInput.value)
    };
}

function buildPaymentDates(start, end, frequency, paymentDay) {
    const dates = [];
    let current = start;

    const buildMonthDate = (from, monthsToAdd) => {
        const tentative = new Date(from);
        tentative.setMonth(tentative.getMonth() + monthsToAdd);
        const year = tentative.getFullYear();
        const month = tentative.getMonth();

        switch (paymentDay) {
            case 'FirstOfMonth':
                return new Date(year, month, 1);
            case 'TenthOfMonth':
                return new Date(year, month, 10);
            case 'LastOfMonth':
            default:
                return new Date(year, month + 1, 0);
        }
    };

    while (current < end) {
        let next;
        switch (frequency) {
            case 'Daily':
                next = addDays(current, 1);
                break;
            case 'Quarterly':
                next = buildMonthDate(current, 3);
                break;
            case 'Monthly':
            default:
                next = buildMonthDate(current, 1);
                break;
        }

        if (next > end) {
            next = new Date(end);
        }

        dates.push(next);
        current = next;
    }

    return dates;
}

function getLastPaymentPeriodStart(startDate, endDate) {
    const frequency = document.getElementById('payment-frequency').value;
    const paymentDay = document.getElementById('payment-day').value;

    if (!startDate || !endDate || startDate >= endDate) {
        return null;
    }

    const paymentDates = buildPaymentDates(startDate, endDate, frequency, paymentDay);
    if (!paymentDates.length) {
        return null;
    }

    if (paymentDates.length === 1) {
        return startDate;
    }

    return paymentDates[paymentDates.length - 2];
}

function enforceInterestApplicationAvailability() {
    const frequency = document.getElementById('payment-frequency').value;
    const isLongPeriod = frequency === 'Monthly' || frequency === 'Quarterly';

    Array.from(interestApplicationSelect.options).forEach(option => {
        option.disabled = !isLongPeriod && option.value !== 'DailyAccrual';
    });

    if (!isLongPeriod && interestApplicationSelect.value !== 'DailyAccrual') {
        interestApplicationSelect.value = 'DailyAccrual';
    }
}

function addRateRow(rate) {
    const dateFromValue = rate?.dateFrom
        ? rate.dateFrom.split('T')[0]
        : '';

    const dateToValue = rate?.dateTo
        ? rate.dateTo.split('T')[0]
        : '';

    const rateValue = rate?.rate ?? '';

    const row = document.createElement('tr');
    row.innerHTML = `
        <td><input type="date" class="date-from" value="${dateFromValue}" required></td>
        <td><input type="date" class="date-to" value="${dateToValue}" required></td>
        <td><input type="number" step="0.01" class="rate-value" value="${rateValue}" required></td>
        <td><button type="button" class="secondary remove-rate">Usuń</button></td>
    `;

    row.querySelector('.remove-rate').addEventListener('click', () => {
        row.remove();
        alignRateBoundariesToCreditDates();
    });

    rateTableBody.appendChild(row);
}

function alignRateBoundariesToCreditDates() {
    const rows = rateTableBody.querySelectorAll('tr');
    if (!rows.length) {
        return;
    }

    const { startDate, endDate } = getCreditDates();
    if (!startDate || !endDate) {
        return;
    }

    rows[0].querySelector('.date-from').value = formatDateInput(startDate);
    rows[rows.length - 1].querySelector('.date-to').value = formatDateInput(endDate);
}

function handleAddRateRow() {
    const rows = rateTableBody.querySelectorAll('tr');
    const { startDate, endDate } = getCreditDates();
    const lastPeriodStart = getLastPaymentPeriodStart(startDate, endDate);

    if (!rows.length) {
        addRateRow({
            dateFrom: formatDateInput(startDate),
            dateTo: formatDateInput(endDate)
        });
        return;
    }

    const lastRow = rows[rows.length - 1];

    if (endDate && lastPeriodStart && lastPeriodStart < endDate) {
        lastRow.querySelector('.date-to').value = formatDateInput(lastPeriodStart);

        const newStart = addDays(lastPeriodStart, 1);
        addRateRow({
            dateFrom: formatDateInput(newStart),
            dateTo: formatDateInput(endDate)
        });
    } else {
        addRateRow();
    }

    alignRateBoundariesToCreditDates();
}

document.getElementById('add-rate').addEventListener('click', handleAddRateRow);
document.getElementById('payment-frequency').addEventListener('change', enforceInterestApplicationAvailability);
enforceInterestApplicationAvailability();

function readParametersFromForm() {
    return {
        netValue: parseFloat(document.getElementById('net-value').value) || 0,
        marginRate: parseFloat(document.getElementById('margin-rate').value) || 0,
        paymentFrequency: document.getElementById('payment-frequency').value,
        paymentDay: document.getElementById('payment-day').value,
        creditStartDate: document.getElementById('credit-start').value,
        creditEndDate: document.getElementById('credit-end').value,
        dayCountBasis: document.getElementById('day-count').value,
        interestRateApplication: interestApplicationSelect.value,
        roundingMode: document.getElementById('rounding-mode').value,
        roundingDecimals: parseInt(document.getElementById('rounding-decimals').value || '4', 10),
        processingFeeRate: parseFloat(document.getElementById('processing-fee').value) || 0,
        processingFeeAmount: parseFloat(document.getElementById('processing-fee-amount').value) || 0,
        paymentType: document.getElementById('payment-type').value
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
    interestApplicationSelect.value = parameters.interestRateApplication ?? 'DailyAccrual';
    document.getElementById('rounding-mode').value = parameters.roundingMode ?? 'Bankers';
    document.getElementById('rounding-decimals').value = parameters.roundingDecimals ?? 4;
    document.getElementById('processing-fee').value = parameters.processingFeeRate ?? 0;
    document.getElementById('processing-fee-amount').value = parameters.processingFeeAmount ?? 0;
    document.getElementById('payment-type').value = parameters.paymentType ?? 'DecreasingInstallments';

    enforceInterestApplicationAvailability();
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

    alignRateBoundariesToCreditDates();
}

function validateRatePeriods(rates) {
    const { startDate, endDate } = getCreditDates();
    if (!startDate || !endDate) {
        throw new Error('Podaj datę rozpoczęcia i zakończenia kredytu.');
    }

    if (!rates.length) {
        throw new Error('Dodaj co najmniej jeden okres stopy procentowej.');
    }

    const normalized = rates.map((rate, index) => {
        const dateFrom = parseDateInput(rate.dateFrom);
        const dateTo = parseDateInput(rate.dateTo);

        if (!dateFrom || !dateTo) {
            throw new Error(`Wiersz ${index + 1}: podaj daty początku i końca okresu.`);
        }

        if (dateFrom >= dateTo) {
            throw new Error(`Wiersz ${index + 1}: data "Od" musi być wcześniejsza niż data "Do".`);
        }

        return { ...rate, dateFrom, dateTo };
    });

    const sorted = normalized.sort((a, b) => a.dateFrom - b.dateFrom);

    if (sorted[0].dateFrom.getTime() !== startDate.getTime()) {
        throw new Error('Pierwszy okres stopy musi zaczynać się w dniu uruchomienia kredytu.');
    }

    if (sorted[sorted.length - 1].dateTo.getTime() !== endDate.getTime()) {
        throw new Error('Ostatni okres stopy musi kończyć się w dniu zakończenia kredytu.');
    }

    for (let i = 1; i < sorted.length; i++) {
        const prev = sorted[i - 1];
        const current = sorted[i];
        const expectedStart = addDays(prev.dateTo, 1).getTime();

        if (current.dateFrom.getTime() < expectedStart) {
            throw new Error(`Okres ${i + 1} nakłada się na poprzedni.`);
        }

        if (current.dateFrom.getTime() !== expectedStart) {
            throw new Error(`Pomiędzy okresem ${i} i ${i + 1} występuje przerwa.`);
        }
    }
}

function updateTotalInterest(totalInterest) {
    totalInterestElement.textContent = (totalInterest ?? 0).toFixed(2);
}

function updateApr(apr) {
    aprElement.textContent = `${(apr ?? 0).toFixed(4)}%`;
}

function displaySchedule(schedule, totalInterest, annualPercentageRate) {
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
    updateApr(annualPercentageRate ?? 0);
}

function buildPayload() {
    return {
        parameters: readParametersFromForm(),
        rates: readRatesFromTable()
    };
}

function buildValidatedPayload() {
    const payload = buildPayload();
    validateRatePeriods(payload.rates);
    return payload;
}

function getAntiforgeryToken() {
    const match = document.cookie.match(/XSRF-TOKEN=([^;]+)/);
    return match ? decodeURIComponent(match[1]) : '';
}

function buildAntiforgeryHeaders(baseHeaders = {}) {
    const token = getAntiforgeryToken();
    if (token) {
        return { ...baseHeaders, 'RequestVerificationToken': token };
    }
    return baseHeaders;
}

importForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    importStatus.textContent = '';
    const fileInput = document.getElementById('import-file');
    if (!fileInput.files.length) {
        importStatus.textContent = 'Wybierz plik Excel (.xlsx), OpenDocument (.ods) lub JSON do importu.';
        importStatus.className = 'status error';
        return;
    }

    const formData = new FormData();
    formData.append('file', fileInput.files[0]);

    try {
        const response = await fetch('/api/import', { method: 'POST', body: formData, headers: buildAntiforgeryHeaders() });
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

creditStartInput.addEventListener('change', alignRateBoundariesToCreditDates);
creditEndInput.addEventListener('change', alignRateBoundariesToCreditDates);

document.getElementById('calculate').addEventListener('click', async () => {
    actionStatus.textContent = 'Trwa obliczanie...';
    actionStatus.className = 'status';
    try {
        const payload = buildValidatedPayload();
        const response = await fetch('/api/calculate', {
            method: 'POST',
            headers: buildAntiforgeryHeaders({ 'Content-Type': 'application/json' }),
            body: JSON.stringify(payload)
        });
        if (!response.ok) {
            throw new Error(await response.text());
        }
        const result = await response.json();
        displaySchedule(result.schedule, result.totalInterest, result.annualPercentageRate);
        actionStatus.textContent = 'Harmonogram został obliczony.';
        actionStatus.className = 'status success';
    } catch (error) {
        actionStatus.textContent = `Błąd obliczeń: ${error.message}`;
        actionStatus.className = 'status error';
        updateTotalInterest(0);
        updateApr(0);
    }
});

document.getElementById('export').addEventListener('click', async () => {
    actionStatus.textContent = 'Trwa eksport...';
    actionStatus.className = 'status';

    try {
        const payload = buildValidatedPayload();
        const format = document.getElementById('export-format').value;

        const response = await fetch(`/api/export?format=${encodeURIComponent(format)}`, {
            method: 'POST',
            headers: buildAntiforgeryHeaders({ 'Content-Type': 'application/json' }),
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            throw new Error(await response.text());
        }

        // --- NEW: try to read filename from Content-Disposition ---
        const contentDisposition = response.headers.get('content-disposition');
        let fileName;

        if (contentDisposition) {
            // handle filename="..." and filename*=UTF-8''...
            const fileNameStarMatch = contentDisposition.match(/filename\*\s*=\s*UTF-8''([^;]+)/i);
            const fileNameMatch = contentDisposition.match(/filename\s*=\s*"([^"]+)"/i)
                || contentDisposition.match(/filename\s*=\s*([^;]+)/i);

            if (fileNameStarMatch && fileNameStarMatch[1]) {
                fileName = decodeURIComponent(fileNameStarMatch[1]);
            } else if (fileNameMatch && fileNameMatch[1]) {
                fileName = fileNameMatch[1].trim();
            }
        }

        // Fallback if header missing / unparsable
        if (!fileName) {
            const extension =
                format === 'ods' ? 'ods' :
                    format === 'json' ? 'json' :
                        'xlsx';
            fileName = `harmonogram.${extension}`;
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName; // <-- now uses backend name or fallback
        link.click();
        window.URL.revokeObjectURL(url);

        actionStatus.textContent = 'Eksport zakończony powodzeniem.';
        actionStatus.className = 'status success';
    } catch (error) {
        actionStatus.textContent = `Eksport nieudany: ${error.message}`;
        actionStatus.className = 'status error';
    }
});

exportLogButton.addEventListener('click', async () => {
    actionStatus.textContent = 'Trwa przygotowywanie logu obliczeń...';
    actionStatus.className = 'status';
    try {
        const payload = buildValidatedPayload();
        const response = await fetch('/api/export-log', {
            method: 'POST',
            headers: buildAntiforgeryHeaders({ 'Content-Type': 'application/json' }),
            body: JSON.stringify(payload)
        });
        if (!response.ok) {
            throw new Error(await response.text());
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'harmonogram-log.txt';
        link.click();
        window.URL.revokeObjectURL(url);

        actionStatus.textContent = 'Log obliczeń został pobrany.';
        actionStatus.className = 'status success';
    } catch (error) {
        actionStatus.textContent = `Nie udało się pobrać logu: ${error.message}`;
        actionStatus.className = 'status error';
    }
});

// Initialize with default dates and one rate row aligned to them
const today = new Date().toISOString().substring(0, 10);
const inSixMonths = new Date();
inSixMonths.setMonth(inSixMonths.getMonth() + 6);
creditStartInput.value = today;
creditEndInput.value = inSixMonths.toISOString().substring(0, 10);

populateRateTable();
alignRateBoundariesToCreditDates();
