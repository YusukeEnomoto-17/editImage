document.addEventListener('DOMContentLoaded', () => {
    const colsInput = document.getElementById('cols');
    const rowsInput = document.getElementById('rows');
    const addImageBtn = document.getElementById('add-image-btn');
    const downloadPdfBtn = document.getElementById('download-pdf-btn');
    const paletteImagesDiv = document.getElementById('palette-images');
    const canvasGrid = document.getElementById('canvas-grid');

    let uploadedImages = []; // { id: 'img-1', src: 'data:image/jpeg;base64,...' }
    let placedImages = {}; // { cellId: { imageId: 'img-1', colSpan: 1, rowSpan: 1, startCell: 'cell-0-0' } }
    let selectedCells = []; // 複数セル選択用

    // --- グリッドの描画と更新 ---
    function renderGrid() {
        const cols = parseInt(colsInput.value);
        const rows = parseInt(rowsInput.value);

        if (isNaN(cols) || isNaN(rows) || cols < 1 || rows < 1) {
            alert('列数と行数は1以上の数値を入力してください。');
            return;
        }

        canvasGrid.innerHTML = '';
        canvasGrid.style.gridTemplateColumns = `repeat(${cols}, 1fr)`;
        canvasGrid.style.gridTemplateRows = `repeat(${rows}, 1fr)`;

        for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
                const cell = document.createElement('div');
                cell.classList.add('grid-cell');
                cell.dataset.row = r;
                cell.dataset.col = c;
                cell.id = `cell-${r}-${c}`;
                canvasGrid.appendChild(cell);
            }
        }
        // グリッド再描画後、配置済みの画像を再配置
        repositionAllPlacedImages();
    }

    function repositionAllPlacedImages() {
        // 一度すべての画像を削除
        document.querySelectorAll('.grid-image-container').forEach(el => el.remove());

        // placedImages オブジェクトを元に再配置
        for (const cellId in placedImages) {
            const placement = placedImages[cellId];
            const startCell = document.getElementById(placement.startCell);
            if (startCell) {
                const imgData = uploadedImages.find(img => img.id === placement.imageId);
                if (imgData) {
                    placeImageOnCanvas(imgData.src, placement.imageId, startCell, placement.colSpan, placement.rowSpan);
                }
            }
        }
    }

    // --- イベントリスナー ---
    colsInput.addEventListener('change', renderGrid);
    rowsInput.addEventListener('change', renderGrid);
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            clearSelectedCells();
        }
    });

    addImageBtn.addEventListener('click', () => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        input.multiple = true;
        input.onchange = (e) => {
            Array.from(e.target.files).forEach(file => {
                const reader = new FileReader();
                reader.onload = (event) => {
                    const imgId = `img-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
                    uploadedImages.push({ id: imgId, src: event.target.result });
                    addThumbnailToPalette(imgId, event.target.result);
                    updateGridSettingsForNewImage();
                };
                reader.readAsDataURL(file);
            });
        };
        input.click();
    });

    downloadPdfBtn.addEventListener('click', downloadPdf);

    // --- 画像パレットへのサムネイル追加 ---
    function addThumbnailToPalette(id, src) {
        const img = document.createElement('img');
        img.src = src;
        img.classList.add('palette-thumbnail');
        img.dataset.imageId = id;
        img.draggable = true;
        paletteImagesDiv.appendChild(img);
    }

    // --- グリッド設定の自動調整 ---
    function updateGridSettingsForNewImage() {
        if (uploadedImages.length === 0) return;

        let numImages = uploadedImages.length;
        let bestCols = 1;
        let bestRows = numImages;

        // 縦横比がA4に近いグリッドを探す (例: 3x3, 4x3, 4x4 など)
        // A4の縦横比は約 1:1.41
        let minRatioDiff = Infinity;

        for (let c = 1; c <= numImages; c++) {
            let r = Math.ceil(numImages / c);
            let currentRatio = (c / r); // グリッドの縦横比
            let a4Ratio = 210 / 297; // A4の縦横比

            let diff = Math.abs(currentRatio - a4Ratio);

            if (diff < minRatioDiff) {
                minRatioDiff = diff;
                bestCols = c;
                bestRows = r;
            }
        }
        colsInput.value = bestCols;
        rowsInput.value = bestRows;
        renderGrid(); // グリッドを更新
    }

    // --- ドラッグ＆ドロップ処理 ---
    let draggedImageId = null;

    paletteImagesDiv.addEventListener('dragstart', (e) => {
        if (e.target.classList.contains('palette-thumbnail')) {
            draggedImageId = e.target.dataset.imageId;
            e.dataTransfer.setData('text/plain', draggedImageId);
            e.target.classList.add('dragging');
        }
    });

    paletteImagesDiv.addEventListener('dragend', (e) => {
        if (e.target.classList.contains('palette-thumbnail')) {
            e.target.classList.remove('dragging');
        }
        draggedImageId = null;
    });

    canvasGrid.addEventListener('dragover', (e) => {
        e.preventDefault(); // ドロップを許可
        e.dataTransfer.dropEffect = 'copy';
    });

    canvasGrid.addEventListener('drop', (e) => {
        e.preventDefault();
        const targetCell = e.target.closest('.grid-cell');
        if (!targetCell) return;

        const imageId = e.dataTransfer.getData('text/plain');
        const imgData = uploadedImages.find(img => img.id === imageId);

        if (!imgData) return;

        let startCellId = targetCell.id;
        let colSpan = 1;
        let rowSpan = 1;

        if (selectedCells.length > 0) {
            // 複数セル選択時
            const { minRow, minCol, maxRow, maxCol } = getSelectedCellBounds();
            colSpan = maxCol - minCol + 1;
            rowSpan = maxRow - minRow + 1;
            startCellId = `cell-${minRow}-${minCol}`;

            // 選択範囲内のセルに既に画像があるかチェック
            let conflict = false;
            for (let r = minRow; r <= maxRow; r++) {
                for (let c = minCol; c <= maxCol; c++) {
                    const cellId = `cell-${r}-${c}`;
                    if (placedImages[cellId] && placedImages[cellId].startCell !== startCellId) {
                        conflict = true;
                        break;
                    }
                }
                if (conflict) break;
            }

            if (conflict) {
                alert('選択範囲内に既に別の画像が配置されています。');
                clearSelectedCells();
                return;
            }

            // 既存の画像を削除 (同じ領域に配置する場合)
            for (let r = minRow; r <= maxRow; r++) {
                for (let c = minCol; c <= maxCol; c++) {
                    const cellId = `cell-${r}-${c}`;
                    if (placedImages[cellId] && placedImages[cellId].startCell === startCellId) {
                        delete placedImages[cellId];
                    }
                }
            }

        } else {
            // 単一セル選択時
            // 既存の画像を削除
            if (placedImages[startCellId]) {
                const existingPlacement = placedImages[startCellId];
                const existingStartCellId = existingPlacement.startCell;
                const existingColSpan = existingPlacement.colSpan;
                const existingRowSpan = existingPlacement.rowSpan;

                // 既存の画像を占有していたセルをクリア
                const existingStartCell = document.getElementById(existingStartCellId);
                const existingStartRow = parseInt(existingStartCell.dataset.row);
                const existingStartCol = parseInt(existingStartCell.dataset.col);

                for (let r = existingStartRow; r < existingStartRow + existingRowSpan; r++) {
                    for (let c = existingStartCol; c < existingStartCol + existingColSpan; c++) {
                        delete placedImages[`cell-${r}-${c}`];
                    }
                }
            }
        }

        // 新しい画像を配置
        placeImageOnCanvas(imgData.src, imageId, document.getElementById(startCellId), colSpan, rowSpan);

        // placedImages オブジェクトを更新
        const startCell = document.getElementById(startCellId);
        const startRow = parseInt(startCell.dataset.row);
        const startCol = parseInt(startCell.dataset.col);

        for (let r = startRow; r < startRow + rowSpan; r++) {
            for (let c = startCol; c < startCol + colSpan; c++) {
                placedImages[`cell-${r}-${c}`] = {
                    imageId: imageId,
                    colSpan: colSpan,
                    rowSpan: rowSpan,
                    startCell: startCellId // 複数セル結合の場合、開始セルを記録
                };
            }
        }
        clearSelectedCells();
    });

    // キャンバス上の画像を移動
    canvasGrid.addEventListener('dragstart', (e) => {
        const imageContainer = e.target.closest('.grid-image-container');
        if (imageContainer) {
            const imageId = imageContainer.dataset.imageId;
            const startCellId = imageContainer.dataset.startCell; // 配置元の開始セルID
            e.dataTransfer.setData('text/plain', imageId);
            e.dataTransfer.setData('text/startCellId', startCellId); // 配置元の情報を渡す
            e.target.classList.add('dragging');
            draggedImageId = imageId; // ドラッグ中の画像IDを保持
        }
    });

    canvasGrid.addEventListener('dragend', (e) => {
        if (e.target.classList.contains('grid-image-container')) {
            e.target.classList.remove('dragging');
        }
        draggedImageId = null;
    });

    canvasGrid.addEventListener('drop', (e) => {
        e.preventDefault();
        const targetCell = e.target.closest('.grid-cell');
        if (!targetCell) return;

        const imageId = e.dataTransfer.getData('text/plain');
        const originalStartCellId = e.dataTransfer.getData('text/startCellId'); // 移動元の開始セルID

        if (!imageId || !originalStartCellId) return; // パレットからのドロップか、キャンバスからの移動か判別

        const imgData = uploadedImages.find(img => img.id === imageId);
        if (!imgData) return;

        // 移動元の情報を取得
        const originalPlacement = placedImages[originalStartCellId];
        if (!originalPlacement || originalPlacement.imageId !== imageId) {
            // 複数セル結合された画像の一部をドラッグした場合、originalStartCellIdは結合の開始セルではない可能性がある
            // その場合は、draggedImageIdを元にplacedImagesを検索して、実際の開始セルを見つける
            let actualOriginalStartCellId = null;
            for (const cellId in placedImages) {
                if (placedImages[cellId].imageId === imageId) {
                    actualOriginalStartCellId = placedImages[cellId].startCell;
                    break;
                }
            }
            if (!actualOriginalStartCellId) return; // 見つからなければエラー
            originalStartCellId = actualOriginalStartCellId;
        }

        const originalColSpan = originalPlacement.colSpan;
        const originalRowSpan = originalPlacement.rowSpan;
        const originalStartCell = document.getElementById(originalStartCellId);
        const originalStartRow = parseInt(originalStartCell.dataset.row);
        const originalStartCol = parseInt(originalStartCell.dataset.col);

        // 移動先の情報を決定
        let newStartCellId = targetCell.id;
        let newColSpan = originalColSpan;
        let newRowSpan = originalRowSpan;

        if (selectedCells.length > 0) {
            // 複数セル選択中にドロップされた場合
            const { minRow, minCol, maxRow, maxCol } = getSelectedCellBounds();
            newColSpan = maxCol - minCol + 1;
            newRowSpan = maxRow - minRow + 1;
            newStartCellId = `cell-${minRow}-${minCol}`;

            // 選択範囲が元の画像のサイズと異なる場合、警告
            if (newColSpan !== originalColSpan || newRowSpan !== originalRowSpan) {
                alert('移動先の選択範囲が元の画像のサイズと異なります。元のサイズで配置されます。');
                newColSpan = originalColSpan;
                newRowSpan = originalRowSpan;
                newStartCellId = targetCell.id; // 選択範囲ではなく、ドロップされたセルを基準にする
            }
        }

        const newStartCell = document.getElementById(newStartCellId);
        const newStartRow = parseInt(newStartCell.dataset.row);
        const newStartCol = parseInt(newStartCell.dataset.col);

        // 移動先の範囲をチェック
        const cols = parseInt(colsInput.value);
        const rows = parseInt(rowsInput.value);

        if (newStartCol + newColSpan > cols || newStartRow + newRowSpan > rows) {
            alert('画像がキャンバスの範囲外にはみ出します。');
            clearSelectedCells();
            return;
        }

        // 移動先の範囲に既に別の画像があるかチェック
        let conflict = false;
        for (let r = newStartRow; r < newStartRow + newRowSpan; r++) {
            for (let c = newStartCol; c < newStartCol + newColSpan; c++) {
                const cellId = `cell-${r}-${c}`;
                if (placedImages[cellId] && placedImages[cellId].startCell !== originalStartCellId) {
                    conflict = true;
                    break;
                }
            }
            if (conflict) break;
        }

        if (conflict) {
            alert('移動先の範囲に既に別の画像が配置されています。');
            clearSelectedCells();
            return;
        }

        // 移動元の画像を placedImages から削除
        for (let r = originalStartRow; r < originalStartRow + originalRowSpan; r++) {
            for (let c = originalStartCol; c < originalStartCol + originalColSpan; c++) {
                delete placedImages[`cell-${r}-${c}`];
            }
        }
        // キャンバスから元の画像要素を削除
        const oldImageContainer = document.querySelector(`.grid-image-container[data-image-id="${imageId}"][data-start-cell="${originalStartCellId}"]`);
        if (oldImageContainer) {
            oldImageContainer.remove();
        }

        // 新しい位置に画像を配置
        placeImageOnCanvas(imgData.src, imageId, newStartCell, newColSpan, newRowSpan);

        // placedImages オブジェクトを更新
        for (let r = newStartRow; r < newStartRow + newRowSpan; r++) {
            for (let c = newStartCol; c < newStartCol + newColSpan; c++) {
                placedImages[`cell-${r}-${c}`] = {
                    imageId: imageId,
                    colSpan: newColSpan,
                    rowSpan: newRowSpan,
                    startCell: newStartCellId
                };
            }
        }
        clearSelectedCells();
    });

    // --- キャンバスへの画像配置ロジック ---
    function placeImageOnCanvas(src, imageId, targetCell, colSpan = 1, rowSpan = 1) {
        // 既存の画像を削除 (同じimageIdで同じstartCellのものが存在する場合)
        const existingContainer = document.querySelector(`.grid-image-container[data-image-id="${imageId}"][data-start-cell="${targetCell.id}"]`);
        if (existingContainer) {
            existingContainer.remove();
        }

        const imageContainer = document.createElement('div');
        imageContainer.classList.add('grid-image-container');
        imageContainer.dataset.imageId = imageId;
        imageContainer.dataset.startCell = targetCell.id; // 結合されたセルの開始セルIDを記録
        imageContainer.draggable = true; // 配置済み画像もドラッグ可能に

        const img = document.createElement('img');
        img.src = src;
        img.classList.add('grid-image');
        imageContainer.appendChild(img);

        const removeBtn = document.createElement('button');
        removeBtn.classList.add('remove-image-btn');
        removeBtn.innerHTML = '&times;';
        removeBtn.title = '画像を削除';
        removeBtn.addEventListener('click', (e) => {
            e.stopPropagation(); // 親要素へのイベント伝播を防ぐ
            removeImageFromCanvas(imageId, targetCell.id);
        });
        imageContainer.appendChild(removeBtn);

        // グリッドのCSSプロパティを設定
        imageContainer.style.gridColumnStart = parseInt(targetCell.dataset.col) + 1;
        imageContainer.style.gridColumnEnd = parseInt(targetCell.dataset.col) + 1 + colSpan;
        imageContainer.style.gridRowStart = parseInt(targetCell.dataset.row) + 1;
        imageContainer.style.gridRowEnd = parseInt(targetCell.dataset.row) + 1 + rowSpan;

        canvasGrid.appendChild(imageContainer);
    }

    // --- キャンバスからの画像削除 ---
    function removeImageFromCanvas(imageId, startCellId) {
        const imageContainer = document.querySelector(`.grid-image-container[data-image-id="${imageId}"][data-start-cell="${startCellId}"]`);
        if (imageContainer) {
            imageContainer.remove();

            // placedImages からも削除
            const startCell = document.getElementById(startCellId);
            const startRow = parseInt(startCell.dataset.row);
            const startCol = parseInt(startCell.dataset.col);

            // 削除する画像の colSpan と rowSpan を取得
            let colSpanToRemove = 1;
            let rowSpanToRemove = 1;
            for (const cellId in placedImages) {
                if (placedImages[cellId].imageId === imageId && placedImages[cellId].startCell === startCellId) {
                    colSpanToRemove = placedImages[cellId].colSpan;
                    rowSpanToRemove = placedImages[cellId].rowSpan;
                    break;
                }
            }

            for (let r = startRow; r < startRow + rowSpanToRemove; r++) {
                for (let c = startCol; c < startCol + colSpanToRemove; c++) {
                    delete placedImages[`cell-${r}-${c}`];
                }
            }
        }
    }

    // --- セル選択機能 ---
    canvasGrid.addEventListener('click', (e) => {
        const cell = e.target.closest('.grid-cell');
        if (!cell) return;

        // 既に画像が配置されているセルは選択できない
        if (placedImages[cell.id]) {
            alert('このセルには既に画像が配置されています。');
            clearSelectedCells();
            return;
        }

        if (selectedCells.includes(cell.id)) {
            // 既に選択されている場合は解除
            selectedCells = selectedCells.filter(id => id !== cell.id);
            cell.classList.remove('highlighted');
        } else {
            // 新しく選択
            selectedCells.push(cell.id);
            cell.classList.add('highlighted');
        }

        // 選択されたセルが長方形を形成しているかチェック
        if (selectedCells.length > 1) {
            if (!isRectangularSelection()) {
                alert('選択されたセルは長方形を形成していません。');
                clearSelectedCells();
            }
        }
    });

    function clearSelectedCells() {
        selectedCells.forEach(id => {
            const cell = document.getElementById(id);
            if (cell) cell.classList.remove('highlighted');
        });
        selectedCells = [];
    }

    function getSelectedCellBounds() {
        if (selectedCells.length === 0) return null;

        let minRow = Infinity, maxRow = -Infinity;
        let minCol = Infinity, maxCol = -Infinity;

        selectedCells.forEach(id => {
            const cell = document.getElementById(id);
            const row = parseInt(cell.dataset.row);
            const col = parseInt(cell.dataset.col);
            minRow = Math.min(minRow, row);
            maxRow = Math.max(maxRow, row);
            minCol = Math.min(minCol, col);
            maxCol = Math.max(maxCol, col);
        });
        return { minRow, maxRow, minCol, maxCol };
    }

    function isRectangularSelection() {
        if (selectedCells.length <= 1) return true;

        const { minRow, maxRow, minCol, maxCol } = getSelectedCellBounds();
        const expectedCellsCount = (maxRow - minRow + 1) * (maxCol - minCol + 1);

        if (selectedCells.length !== expectedCellsCount) {
            return false;
        }

        // すべての期待されるセルが実際に選択されているか確認
        for (let r = minRow; r <= maxRow; r++) {
            for (let c = minCol; c <= maxCol; c++) {
                if (!selectedCells.includes(`cell-${r}-${c}`)) {
                    return false;
                }
            }
        }
        return true;
    }

    // --- PDF生成機能 ---
    async function downloadPdf() {
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF({
            unit: 'mm',
            format: 'a4',
            orientation: 'portrait'
        });

        const canvasWidth = canvasGrid.offsetWidth;
        const canvasHeight = canvasGrid.offsetHeight;
        const cols = parseInt(colsInput.value);
        const rows = parseInt(rowsInput.value);

        const cellWidthPx = canvasWidth / cols;
        const cellHeightPx = canvasHeight / rows;

        // A4サイズ (210mm x 297mm) に合わせるためのスケールファクター
        const scaleX = 210 / canvasWidth;
        const scaleY = 297 / canvasHeight;

        // 配置されている画像を収集
        const imagesToRender = []; // { imageId, startCellId, colSpan, rowSpan }
        const processedImageContainers = new Set(); // 重