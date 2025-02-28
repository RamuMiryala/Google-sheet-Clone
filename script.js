let selectedCells = new Set();
let isDragging = false;
let startCell = null;
let endCell = null;

// Convert column index to letters (A, B, C, ...)
function getRowLabel(index) {
    let label = "";
    while (index >= 0) {
        label = String.fromCharCode((index % 26) + 65) + label;
        index = Math.floor(index / 26) - 1;
    }
    return label;
}

// Convert row index to numbers (1, 2, 3, ...)
function getColumnLabel(index) {
    return (index + 1).toString();
}

// Initialize spreadsheet grid
window.onload = function initializeGrid() {
    const spreadsheet = document.getElementById('spreadsheet');

    // Create header row
    const headerRow = document.createElement('div');
    headerRow.classList.add('header-row');
    headerRow.innerHTML = `<div class="cell header"></div>`; // Empty top-left cell

    for (let col = 0; col < 30; col++) {
        headerRow.innerHTML += `<div class="cell header">${getColumnLabel(col)}</div>`;
    }
    spreadsheet.appendChild(headerRow);

    // Create rows with row numbers and cells
    for (let row = 0; row < 26; row++) {
        const newRow = document.createElement('div');
        newRow.classList.add('row');
        newRow.innerHTML = `<div class="cell header">${getRowLabel(row)}</div>`;

        for (let col = 0; col < 30; col++) {
            const cellId = `${getColumnLabel(col)}${getRowLabel(row)}`;
            newRow.innerHTML += `<div class="cell" contenteditable="true" data-cell="${cellId}"></div>`;
        }
        spreadsheet.appendChild(newRow);
    }

    attachCellListeners();
};

// Attach event listeners to cells
function attachCellListeners() {
    document.querySelectorAll('.cell[contenteditable="true"]').forEach(cell => {
        cell.addEventListener('mousedown', (e) => {
            isDragging = true;
            startCell = cell;
            selectedCells.clear();
            selectedCells.add(cell);
            updateSelectionHighlight();
        });

        cell.addEventListener('mouseup', () => {
            isDragging = false;
            updateFormulaBarWithRange();
        });

        cell.addEventListener('mouseover', (e) => {
            if (isDragging) {
                endCell = cell;
                highlightSelection(startCell, endCell);
            }
        });

        cell.addEventListener('click', (e) => {
            if (e.shiftKey) {
                selectedCells.add(cell);
            } else {
                selectedCells.clear();
                selectedCells.add(cell);
            }
            updateSelectionHighlight();
        });
    });

    document.addEventListener('mouseup', () => {
        isDragging = false;
        updateFormulaBarWithRange();
    });
}

// Highlight selected cells
function updateSelectionHighlight() {
    document.querySelectorAll('.cell').forEach(cell => cell.classList.remove('selected'));
    selectedCells.forEach(cell => cell.classList.add('selected'));
}

// Highlight cells while dragging
function highlightSelection(start, end) {
    if (!start || !end) return;

    const startCellPosition = start.getAttribute('data-cell');
    const endCellPosition = end.getAttribute('data-cell');

    const startCol = startCellPosition.match(/[A-Z]+/)[0];
    const startRow = parseInt(startCellPosition.match(/\d+/)[0]);
    const endCol = endCellPosition.match(/[A-Z]+/)[0];
    const endRow = parseInt(endCellPosition.match(/\d+/)[0]);

    const startColIndex = getColumnIndex(startCol);
    const endColIndex = getColumnIndex(endCol);

    selectedCells.clear();

    document.querySelectorAll('.cell[contenteditable="true"]').forEach(cell => {
        const cellId = cell.getAttribute('data-cell');
        const col = getColumnIndex(cellId.match(/[A-Z]+/)[0]);
        const row = parseInt(cellId.match(/\d+/)[0]);

        if (
            row >= Math.min(startRow, endRow) &&
            row <= Math.max(startRow, endRow) &&
            col >= Math.min(startColIndex, endColIndex) &&
            col <= Math.max(startColIndex, endColIndex)
        ) {
            selectedCells.add(cell);
        }
    });

    updateSelectionHighlight();
}

// Get column index from label (A -> 0, B -> 1, ..., Z -> 25, AA -> 26, etc.)
function getColumnIndex(label) {
    let index = 0;
    for (let i = 0; i < label.length; i++) {
        index = index * 26 + (label.charCodeAt(i) - 65 + 1);
    }
    return index - 1;
}

// Update formula bar with selected range
function updateFormulaBarWithRange() {
    if (selectedCells.size > 0) {
        const cellsArray = Array.from(selectedCells);
        const startCellId = cellsArray[0].getAttribute('data-cell');
        const endCellId = cellsArray[cellsArray.length - 1].getAttribute('data-cell');
        document.getElementById('formulaBar').value = `${startCellId}:${endCellId}`;
    }
}

// Perform arithmetic operations (SUM, AVERAGE)
function performArithmetic(operation) {
    let values = [];
    selectedCells.forEach(cell => {
        const cellValue = parseFloat(cell.textContent);
        if (!isNaN(cellValue)) {
            values.push(cellValue);
        }
    });

    if (values.length === 0) {
        alert("No numeric values found in the selected range.");
        return;
    }

    let result;
    if (operation === 'sum') {
        result = values.reduce((acc, val) => acc + val, 0);
    } else if (operation === 'average') {
        result = values.reduce((acc, val) => acc + val, 0) / values.length;
    } else {
        alert("Invalid operation.");
        return;
    }

    document.getElementById('formulaBar').value = result;
}
// Add a new column
function addColumn() {
    const spreadsheet = document.getElementById('spreadsheet');
    const colCount = spreadsheet.querySelector('.header-row').children.length - 1;
    const newColumnLabel = getColumnLabel(colCount);

    const newHeader = document.createElement('div');
    newHeader.classList.add('cell', 'header');
    newHeader.textContent = newColumnLabel;
    spreadsheet.querySelector('.header-row').appendChild(newHeader);

    document.querySelectorAll('.row').forEach((row, index) => {
        const newCell = document.createElement('div');
        newCell.classList.add('cell');
        newCell.setAttribute('contenteditable', 'true');
        newCell.setAttribute('data-cell', `${newColumnLabel}${getRowLabel(index)}`);
        newCell.setAttribute('draggable', 'true');
        row.appendChild(newCell);
    });

    attachCellListeners();
}
// Add a new row
function addRow() {
    const spreadsheet = document.getElementById('spreadsheet');
    const rowCount = spreadsheet.querySelectorAll('.row').length;
    const newRow = document.createElement('div');
    newRow.classList.add('row');
    newRow.innerHTML = `<div class="cell header">${getRowLabel(rowCount)}</div>`;

    const colCount = spreadsheet.querySelector('.header-row').children.length - 1;
    for (let col = 0; col < colCount; col++) {
        const cellId = `${getColumnLabel(col)}${getRowLabel(rowCount)}`;
        newRow.innerHTML += `<div class="cell" contenteditable="true" data-cell="${cellId}" draggable="true"></div>`;
    }

    spreadsheet.appendChild(newRow);
    attachCellListeners();
}
// Apply text styles (Bold, Italic, Underline, Standard)
function applyStyle(style) {
    selectedCells.forEach(cell => {
        if (style === 'bold') {
            cell.style.fontWeight = cell.style.fontWeight === 'bold' ? 'normal' : 'bold';
        } else if (style === 'italic') {
            cell.style.fontStyle = cell.style.fontStyle === 'italic' ? 'normal' : 'italic';
        } else if (style === 'underline') {
            cell.style.textDecoration = cell.style.textDecoration === 'underline' ? 'none' : 'underline';
        } else if (style === 'standard') {
            cell.style.fontWeight = 'normal';
            cell.style.fontStyle = 'normal';
            cell.style.textDecoration = 'none';
            cell.style.color = 'black';
        }
    });
}

// Apply text color
function applyColor(color) {
    selectedCells.forEach(cell => {
        cell.style.color = color;
    });
}
// Text transformations (Uppercase, Lowercase, Capitalize)
function applyTextTransform(transform) {
    selectedCells.forEach(cell => {
        if (transform === 'uppercase') {
            cell.textContent = cell.textContent.toUpperCase();
        } else if (transform === 'lowercase') {
            cell.textContent = cell.textContent.toLowerCase();
        } else if (transform === 'capitalize') {
            cell.textContent = cell.textContent.replace(/\b\w/g, char => char.toUpperCase());
        }
    });
}