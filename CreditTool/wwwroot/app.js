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

function getFirstPaymentPeriodEnd(startDate, endDate) {
    const frequency = document.getElementById('payment-frequency').value;
    const paymentDay = document.getElementById('payment-day').value;

    if (!startDate || !endDate || startDate >= endDate) {
        return null;
    }

    const paymentDates = buildPaymentDates(startDate, endDate, frequency, paymentDay);
    if (!paymentDates.length) {
        return null;
    }

    return paymentDates[0];
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

// Lock/unlock functionality
function toggleLock(row, field, button) {
    const rows = Array.from(rateTableBody.querySelectorAll('tr'));
    const rowIndex = rows.indexOf(row);
    const dataField = field === 'start' ? 'startLocked' : 'endLocked';
    const isLocked = row.dataset[dataField] === 'true';
    const newLockState = !isLocked ? 'true' : 'false';

    // Lock/unlock this field
    row.dataset[dataField] = newLockState;
    button.textContent = !isLocked ? '🔒' : '🔓';

    // Also lock/unlock the adjacent boundary
    if (field === 'start' && rowIndex > 0) {
        // Lock previous row's end
        const prevRow = rows[rowIndex - 1];
        prevRow.dataset.endLocked = newLockState;
        const prevEndButton = prevRow.querySelector('.lock-end');
        if (prevEndButton) {
            prevEndButton.textContent = !isLocked ? '🔒' : '🔓';
        }
    } else if (field === 'end' && rowIndex < rows.length - 1) {
        // Lock next row's start
        const nextRow = rows[rowIndex + 1];
        nextRow.dataset.startLocked = newLockState;
        const nextStartButton = nextRow.querySelector('.lock-start');
        if (nextStartButton) {
            nextStartButton.textContent = !isLocked ? '🔒' : '🔓';
        }
    }
}

function isDateLocked(row, field) {
    if (!row) return false;
    const dataField = field === 'start' ? 'startLocked' : 'endLocked';
    return row.dataset[dataField] === 'true';
}

function isBoundaryLocked(row, field) {
    // Check if a boundary is locked (either this row's field or the adjacent row's corresponding field)
    if (!row) return false;

    const rows = Array.from(rateTableBody.querySelectorAll('tr'));
    const rowIndex = rows.indexOf(row);

    if (field === 'start') {
        // Check if this row's start is locked OR previous row's end is locked
        if (isDateLocked(row, 'start')) return true;
        if (rowIndex > 0 && isDateLocked(rows[rowIndex - 1], 'end')) return true;
    } else if (field === 'end') {
        // Check if this row's end is locked OR next row's start is locked
        if (isDateLocked(row, 'end')) return true;
        if (rowIndex < rows.length - 1 && isDateLocked(rows[rowIndex + 1], 'start')) return true;
    }

    return false;
}

// Helper to calculate days between dates
function daysBetween(date1, date2) {
    const oneDay = 24 * 60 * 60 * 1000;
    return Math.round((date2 - date1) / oneDay);
}

// Redistribute dates equally in an unlocked segment
function redistributeDatesInSegment(rows, startDate, endDate) {
    if (rows.length === 0) return;

    const totalDays = daysBetween(startDate, endDate) + 1;
    const daysPerRow = Math.floor(totalDays / rows.length);
    const remainder = totalDays % rows.length;

    let currentDate = new Date(startDate);

    rows.forEach((row, index) => {
        const rowDays = daysPerRow + (index < remainder ? 1 : 0);
        const rowStart = new Date(currentDate);
        const rowEnd = addDays(currentDate, rowDays - 1);

        // Ensure last row ends exactly at endDate
        if (index === rows.length - 1) {
            row.querySelector('.date-to').value = formatDateInput(endDate);
        } else {
            row.querySelector('.date-to').value = formatDateInput(rowEnd);
        }

        row.querySelector('.date-from').value = formatDateInput(rowStart);
        currentDate = addDays(rowEnd, 1);
    });
}

// Cascade changes forward through unlocked rows
function cascadeForward(startRow) {
    const rows = Array.from(rateTableBody.querySelectorAll('tr'));
    const startIndex = rows.indexOf(startRow);

    if (startIndex === -1 || startIndex >= rows.length - 1) return;

    // Find the next locked end boundary
    let endIndex = startIndex + 1;
    while (endIndex < rows.length && !isBoundaryLocked(rows[endIndex], 'end')) {
        endIndex++;
    }

    // Get the segment to redistribute
    const segmentRows = rows.slice(startIndex + 1, endIndex + 1);
    if (segmentRows.length === 0) return;

    // Calculate segment boundaries
    const segmentStart = parseDateInput(startRow.querySelector('.date-to').value);
    const segmentStartPlusOne = addDays(segmentStart, 1);

    let segmentEnd;
    if (endIndex < rows.length) {
        segmentEnd = parseDateInput(rows[endIndex].querySelector('.date-to').value);
    } else {
        const { endDate } = getCreditDates();
        segmentEnd = endDate;
    }

    // Redistribute dates in the segment
    redistributeDatesInSegment(segmentRows, segmentStartPlusOne, segmentEnd);
}

// Cascade changes backward through unlocked rows
function cascadeBackward(startRow) {
    const rows = Array.from(rateTableBody.querySelectorAll('tr'));
    const startIndex = rows.indexOf(startRow);

    if (startIndex === -1 || startIndex === 0) return;

    // Find the previous locked start boundary
    let beginIndex = startIndex - 1;
    while (beginIndex >= 0 && !isBoundaryLocked(rows[beginIndex], 'start')) {
        beginIndex--;
    }

    // Get the segment to redistribute
    const segmentRows = rows.slice(beginIndex, startIndex);
    if (segmentRows.length === 0) return;

    // Calculate segment boundaries
    let segmentStart;
    if (beginIndex >= 0) {
        segmentStart = parseDateInput(rows[beginIndex].querySelector('.date-from').value);
    } else {
        const { startDate } = getCreditDates();
        segmentStart = startDate;
    }

    const segmentEnd = parseDateInput(startRow.querySelector('.date-from').value);
    const segmentEndMinusOne = addDays(segmentEnd, -1);

    // Redistribute dates in the segment
    redistributeDatesInSegment(segmentRows, segmentStart, segmentEndMinusOne);
}

// Handle start date changes in auto-continuity mode
function handleStartDateChange(row) {
    const rows = Array.from(rateTableBody.querySelectorAll('tr'));
    const rowIndex = rows.indexOf(row);
    const newStartDate = parseDateInput(row.querySelector('.date-from').value);
    const { startDate: creditStart, endDate: creditEnd } = getCreditDates();

    if (!newStartDate) return;

    // If this is the first row and there's a gap at the start, insert a new row
    if (rowIndex === 0 && newStartDate > creditStart) {
        const gapEnd = addDays(newStartDate, -1);
        addRateRow({
            dateFrom: formatDateInput(creditStart),
            dateTo: formatDateInput(gapEnd)
        }, true);
        return;
    }

    // If there's a previous row, adjust it to maintain continuity
    if (rowIndex > 0) {
        const prevRow = rows[rowIndex - 1];
        const newPrevEnd = addDays(newStartDate, -1);
        prevRow.querySelector('.date-to').value = formatDateInput(newPrevEnd);

        // Only cascade if the boundary is not locked
        if (!isBoundaryLocked(prevRow, 'end')) {
            cascadeBackward(prevRow);
        }
    }
}

// Handle end date changes in auto-continuity mode
function handleEndDateChange(row) {
    const rows = Array.from(rateTableBody.querySelectorAll('tr'));
    const rowIndex = rows.indexOf(row);
    const newEndDate = parseDateInput(row.querySelector('.date-to').value);
    const { startDate: creditStart, endDate: creditEnd } = getCreditDates();

    if (!newEndDate) return;

    // If this is the last row and there's a gap at the end, append a new row
    if (rowIndex === rows.length - 1 && newEndDate < creditEnd) {
        const gapStart = addDays(newEndDate, 1);
        addRateRow({
            dateFrom: formatDateInput(gapStart),
            dateTo: formatDateInput(creditEnd)
        }, false);
        return;
    }

    // If there's a next row, adjust it to maintain continuity
    if (rowIndex < rows.length - 1) {
        const nextRow = rows[rowIndex + 1];
        const newNextStart = addDays(newEndDate, 1);
        nextRow.querySelector('.date-from').value = formatDateInput(newNextStart);

        // Only cascade if the boundary is not locked
        if (!isBoundaryLocked(row, 'end')) {
            cascadeForward(row);
        }
    }
}

function addRateRow(rate, prepend = false) {
    const dateFromValue = rate?.dateFrom
        ? rate.dateFrom.split('T')[0]
        : '';

    const dateToValue = rate?.dateTo
        ? rate.dateTo.split('T')[0]
        : '';

    const rateValue = rate?.rate ?? '';

    const row = document.createElement('tr');
    row.dataset.startLocked = 'false';
    row.dataset.endLocked = 'false';

    const isAutoContinuity = document.getElementById('auto-continuity').checked;
    const lockCellDisplay = isAutoContinuity ? '' : 'none';

    row.innerHTML = `
        <td class="lock-cell" style="display: ${lockCellDisplay};"><button type="button" class="lock-btn lock-start" title="Zablokuj datę początkową">🔓</button></td>
        <td><input type="date" class="date-from" value="${dateFromValue}" required></td>
        <td class="lock-cell" style="display: ${lockCellDisplay};"><button type="button" class="lock-btn lock-end" title="Zablokuj datę końcową">🔓</button></td>
        <td><input type="date" class="date-to" value="${dateToValue}" required></td>
        <td><input type="number" step="0.01" class="rate-value" value="${rateValue}" required></td>
        <td><button type="button" class="secondary remove-rate">Usuń</button></td>
    `;

    // Lock button handlers
    row.querySelector('.lock-start').addEventListener('click', (e) => {
        toggleLock(row, 'start', e.target);
    });

    row.querySelector('.lock-end').addEventListener('click', (e) => {
        toggleLock(row, 'end', e.target);
    });

    row.querySelector('.remove-rate').addEventListener('click', () => {
        row.remove();
        alignRateBoundariesToCreditDates();
    });

    // Date change listeners for auto-continuity
    row.querySelector('.date-from').addEventListener('change', () => {
        if (document.getElementById('auto-continuity').checked) {
            handleStartDateChange(row);
        }
    });

    row.querySelector('.date-to').addEventListener('change', () => {
        if (document.getElementById('auto-continuity').checked) {
            handleEndDateChange(row);
        }
    });

    if (prepend && rateTableBody.firstChild) {
        rateTableBody.insertBefore(row, rateTableBody.firstChild);
    } else {
        rateTableBody.appendChild(row);
    }
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
    const rows = Array.from(rateTableBody.querySelectorAll('tr'));
    const { startDate, endDate } = getCreditDates();
    const autoContinuity = document.getElementById('auto-continuity').checked;

    // If no rows exist, add first row spanning full credit period
    if (!rows.length) {
        addRateRow({
            dateFrom: formatDateInput(startDate),
            dateTo: formatDateInput(endDate)
        });
        return;
    }

    // Without auto-continuity, just add an empty row at start
    if (!autoContinuity) {
        addRateRow({}, true);
        return;
    }

    // With auto-continuity, use smart algorithm
    const firstPeriodEnd = getFirstPaymentPeriodEnd(startDate, endDate);
    const paymentDates = buildPaymentDates(startDate, endDate,
        document.getElementById('payment-frequency').value,
        document.getElementById('payment-day').value);
    const numPaymentPeriods = paymentDates.length;

    // Determine new row duration
    let newRowEnd;

    if (rows.length < numPaymentPeriods && firstPeriodEnd) {
        // Use first payment period if we have fewer rows than periods
        newRowEnd = firstPeriodEnd;
    } else {
        // Find the first unlocked segment and divide it
        const firstRow = rows[0];

        // Find the end of the first unlocked segment
        let segmentEndIndex = 0;
        while (segmentEndIndex < rows.length && !isBoundaryLocked(rows[segmentEndIndex], 'end')) {
            segmentEndIndex++;
        }

        const segmentRows = rows.slice(0, segmentEndIndex + 1);
        const segmentEnd = segmentEndIndex < rows.length
            ? parseDateInput(rows[segmentEndIndex].querySelector('.date-to').value)
            : endDate;

        // Divide the unlocked segment equally among existing rows + new row
        const totalDays = daysBetween(startDate, segmentEnd) + 1;
        const daysForNewRow = Math.floor(totalDays / (segmentRows.length + 1));

        newRowEnd = addDays(startDate, Math.max(0, daysForNewRow - 1));
    }

    // Add the new row at start
    addRateRow({
        dateFrom: formatDateInput(startDate),
        dateTo: formatDateInput(newRowEnd)
    }, true);

    // Redistribute the remaining unlocked rows
    const updatedRows = Array.from(rateTableBody.querySelectorAll('tr'));
    const newRowEndPlusOne = addDays(newRowEnd, 1);

    // Find the first segment of unlocked rows after the new row
    let segmentEndIndex = 1;
    while (segmentEndIndex < updatedRows.length && !isBoundaryLocked(updatedRows[segmentEndIndex], 'end')) {
        segmentEndIndex++;
    }

    if (segmentEndIndex > 1) {
        const segmentRows = updatedRows.slice(1, segmentEndIndex + 1);
        const segmentEnd = segmentEndIndex < updatedRows.length
            ? parseDateInput(updatedRows[segmentEndIndex].querySelector('.date-to').value)
            : endDate;

        redistributeDatesInSegment(segmentRows, newRowEndPlusOne, segmentEnd);
    }

    alignRateBoundariesToCreditDates();
}

document.getElementById('add-rate').addEventListener('click', handleAddRateRow);
document.getElementById('payment-frequency').addEventListener('change', enforceInterestApplicationAvailability);
document.getElementById('auto-continuity').addEventListener('change', toggleLockColumnsVisibility);
enforceInterestApplicationAvailability();
toggleLockColumnsVisibility();

function toggleLockColumnsVisibility() {
    const isChecked = document.getElementById('auto-continuity').checked;
    const display = isChecked ? '' : 'none';

    // Toggle header columns
    document.querySelectorAll('#rate-table th.lock-column').forEach(th => {
        th.style.display = display;
    });

    // Toggle lock cells in all rows
    document.querySelectorAll('#rate-table td.lock-cell').forEach(td => {
        td.style.display = display;
    });
}

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

// Add after the schedule table
function displaySchedule(schedule, totalInterest, annualPercentageRate, warnings, targetPayment, actualFinalPayment) {
    scheduleTableBody.innerHTML = '';
    (schedule ?? []).forEach((item, index) => {
        const row = document.createElement('tr');

        // Add warning class if needed
        if (item.warnings && item.warnings !== 0) {
            row.classList.add('warning-row');
        }

        let paymentDisplay = item.totalPayment.toFixed(2);
        if (item.isFinalPaymentAdjusted) {
            paymentDisplay += ' *';
        }

        row.innerHTML = `
            <td>${item.paymentDate?.substring(0, 10)}</td>
            <td>${item.daysInPeriod}</td>
            <td>${item.interestRate.toFixed(4)}%${item.nominalRate ? ` (nom: ${item.nominalRate.toFixed(4)}%)` : ''}</td>
            <td>${item.interestAmount.toFixed(2)}</td>
            <td>${item.principalPayment.toFixed(2)}</td>
            <td>${paymentDisplay}</td>
            <td>${item.remainingPrincipal.toFixed(2)}</td>
        `;
        scheduleTableBody.appendChild(row);
    });

    const calculatedTotal = totalInterest ?? (schedule ?? []).reduce((sum, item) => sum + (item.interestAmount ?? 0), 0);
    updateTotalInterest(calculatedTotal);
    updateApr(annualPercentageRate ?? 0);

    // Display warnings
    const warningSection = document.getElementById('warnings-section');
    if (warnings && warnings.length > 0) {
        warningSection.innerHTML = '<h3>Ostrzeżenia</h3><ul>' +
            warnings.map(w => `<li>${w}</li>`).join('') + '</ul>';
        warningSection.style.display = 'block';
    } else {
        warningSection.style.display = 'none';
    }

    // Display payment info for equal installments
    const paymentInfo = document.getElementById('payment-info');
    if (targetPayment && actualFinalPayment) {
        const diff = Math.abs(actualFinalPayment - targetPayment);
        if (diff >= 0.01) {
            paymentInfo.innerHTML = `<p><strong>Rata docelowa:</strong> ${targetPayment.toFixed(2)} | <strong>Ostatnia rata:</strong> ${actualFinalPayment.toFixed(2)} (dostosowana)</p>`;
            paymentInfo.style.display = 'block';
        } else {
            paymentInfo.style.display = 'none';
        }
    } else {
        paymentInfo.style.display = 'none';
    }
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
        displaySchedule(
            result.schedule,
            result.totalInterest,
            result.annualPercentageRate,
            result.warnings,
            result.targetLevelPayment,
            result.actualFinalPayment
        );
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
