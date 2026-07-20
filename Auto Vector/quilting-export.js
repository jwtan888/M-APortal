/* Template-driven browser quilting export based on the supplied 372 workbook. */
(function () {
  const LINES_PER_PAGE = 21;

  function clean(value) {
    return String(value || '').replace(/\s+/g, ' ').trim();
  }

  function clone(value) {
    return value == null ? value : JSON.parse(JSON.stringify(value));
  }

  function base64ToBuffer(base64) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
    return bytes.buffer;
  }

  function safeSheetName(value) {
    return clean(value).replace(/[\\/*?:\[\]]/g, ' ').slice(0, 22) || 'PART';
  }

  function safeFileName(value) {
    return clean(value).replace(/[^A-Za-z0-9 _-]+/g, ' ').replace(/\s+/g, ' ').trim();
  }

  async function loadTemplateWorkbook() {
    if (typeof ExcelJS === 'undefined') throw new Error('ExcelJS is not loaded.');
    if (!window.QUILTING_TEMPLATE_XLSX_BASE64) throw new Error('The quilting template is not loaded.');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(base64ToBuffer(window.QUILTING_TEMPLATE_XLSX_BASE64));
    return workbook;
  }

  function cloneTemplateSheet(workbook, source, name) {
    const target = workbook.addWorksheet(name, {
      properties: clone(source.properties),
      pageSetup: clone(source.pageSetup),
      views: clone(source.views),
      state: source.state
    });
    target.headerFooter = clone(source.headerFooter);
    for (let column = 1; column <= 7; column++) {
      const sourceColumn = source.getColumn(column);
      const targetColumn = target.getColumn(column);
      targetColumn.width = sourceColumn.width;
      targetColumn.hidden = sourceColumn.hidden;
      targetColumn.outlineLevel = sourceColumn.outlineLevel;
      targetColumn.style = clone(sourceColumn.style);
    }
    source.eachRow({ includeEmpty: true }, (sourceRow, rowNumber) => {
      const targetRow = target.getRow(rowNumber);
      targetRow.height = sourceRow.height;
      targetRow.hidden = sourceRow.hidden;
      targetRow.outlineLevel = sourceRow.outlineLevel;
      for (let column = 1; column <= 7; column++) {
        const sourceCell = sourceRow.getCell(column);
        const targetCell = targetRow.getCell(column);
        targetCell.value = clone(sourceCell.value);
        targetCell.style = clone(sourceCell.style);
        targetCell.numFmt = sourceCell.numFmt;
        targetCell.dataValidation = clone(sourceCell.dataValidation);
      }
    });
    (source.model.merges || []).forEach(range => target.mergeCells(range));
    return target;
  }

  function replaceSheetImage(workbook, worksheet, dataUrl, range) {
    worksheet._media = [];
    if (!dataUrl) return;
    const imageId = workbook.addImage({ base64: dataUrl, extension: 'png' });
    worksheet.addImage(imageId, range);
  }

  function writeCover(workbook, worksheet, payload, totalPanels) {
    worksheet.getCell('D1').value = clean(payload.stage) || 'P2';
    worksheet.getCell('G1').value = String(totalPanels);
    worksheet.getCell('D3').value = clean(payload.technician);
    if (payload.completedDate) {
      const [year, month, day] = payload.completedDate.split('-').map(Number);
      worksheet.getCell('G3').value = new Date(Date.UTC(year, month - 1, day, 12));
    } else {
      worksheet.getCell('G3').value = new Date();
    }
    worksheet.getCell('G3').numFmt = '[$-14409]d mmmm, yyyy;@';
    replaceSheetImage(workbook, worksheet, payload.coverImage, {
      tl: { col: 2, row: 5.08 }, br: { col: 7, row: 6 }, editAs: 'oneCell'
    });
  }

  function panelLabel(firstPanel, quantity, totalPanels) {
    const panels = Array.from({ length: quantity }, (_, index) => firstPanel + index).join(',');
    return `${panels}/${totalPanels}`;
  }

  function writePartSheet(workbook, worksheet, payload, part, lines, firstPanel, totalPanels) {
    worksheet.getCell('C1').value = clean(payload.style);
    worksheet.getCell('C3').value = `${clean(part.partName)}  x${part.quantity}`;
    worksheet.getCell('G3').value = panelLabel(firstPanel, part.quantity, totalPanels);
    worksheet.getCell('G3').numFmt = '@';
    worksheet.getCell('G10').value = Number(payload.rpm) || 2200;

    for (let index = 0; index < LINES_PER_PAGE; index++) {
      const row = 11 + index;
      const line = lines[index];
      worksheet.getCell(`B${row}`).value = index + 1;
      worksheet.getCell(`C${row}`).value = line ? Number(line.lengthCm || 0) : null;
      worksheet.getCell(`D${row}`).value = line ? Number(line.transitionCm || 0) : null;
      worksheet.getCell(`F${row}`).value = line ? (clean(line.stitchType) || clean(payload.stitchType) || 'STAY STITCH') : null;
    }

    worksheet.getCell('B32').value = 'Total';
    worksheet.getCell('C32').value = { formula: 'SUM(C11:C31)' };
    worksheet.getCell('D32').value = { formula: 'SUM(D11:D31)' };
    worksheet.getCell('B33').value = '%';
    worksheet.getCell('C33').value = { formula: 'IFERROR(C32/(C32+D32),0)' };
    worksheet.getCell('D33').value = { formula: 'IFERROR(D32/(D32+C32),0)' };
    worksheet.getCell('C33').numFmt = '0%';
    worksheet.getCell('D33').numFmt = '0%';
    worksheet.getCell('B34').value = 'EST. TIME (sec)';
    worksheet.getCell('C34').value = { formula: '(10*(C32/2.54)/G10)*60' };
    worksheet.getCell('D34').value = { formula: `(10*(D32/2.54)/${Number(payload.transitionRpm) || 2800})*60` };
    worksheet.getCell('B36').value = 'EST. TIME (min)';
    worksheet.getCell('C36').value = { formula: '((C34+D34)/60)*110%' };
    worksheet.getCell('C36').numFmt = '0.0000';
    worksheet.getCell('D36').value = '(+ 10% tol) GSD';

    replaceSheetImage(workbook, worksheet, part.previewImage, {
      tl: { col: 2, row: 5.08 }, br: { col: 7, row: 6 }, editAs: 'oneCell'
    });
  }

  window.generateQuiltingExcelBrowser = async function (payload, helpers) {
    const parts = (payload.parts || []).map(part => ({
      partName: clean(part.partName),
      quantity: Math.max(1, Math.floor(Number(part.quantity) || 1)),
      previewImage: part.previewImage || '',
      lines: (part.lines || []).map((line, index) => ({
        lineNumber: index + 1,
        lengthCm: Math.max(0, Number(line.lengthCm) || 0),
        transitionCm: Math.max(0, Number(line.transitionCm) || 0),
        stitchType: clean(line.stitchType) || clean(payload.stitchType) || 'STAY STITCH'
      }))
    })).filter(part => part.partName && part.lines.length);
    if (!parts.length) throw new Error('No quilting rows are available.');
    if (!(Number(payload.rpm) > 0) || !(Number(payload.transitionRpm) > 0)) throw new Error('Quilting and transition RPM must be greater than zero.');

    const workbook = await loadTemplateWorkbook();
    const cover = workbook.getWorksheet('Pg1-Cover Page');
    const sourcePart = workbook.getWorksheet('Pg 3-RFLN') || workbook.worksheets[1];
    if (!cover || !sourcePart) throw new Error('The quilting template is missing its cover or part sheet.');
    const totalPanels = parts.reduce((sum, part) => sum + part.quantity, 0);
    writeCover(workbook, cover, payload, totalPanels);

    let sheetNumber = 2;
    let firstPanel = 1;
    const outputSheets = [];
    parts.forEach(part => {
      const chunks = [];
      for (let index = 0; index < part.lines.length; index += LINES_PER_PAGE) chunks.push(part.lines.slice(index, index + LINES_PER_PAGE));
      chunks.forEach((lines, pageIndex) => {
        const suffix = chunks.length > 1 ? `-${pageIndex + 1}` : '';
        const name = `Pg ${sheetNumber}-${safeSheetName(part.partName)}${suffix}`.slice(0, 31);
        const worksheet = cloneTemplateSheet(workbook, sourcePart, name);
        writePartSheet(workbook, worksheet, payload, part, lines, firstPanel, totalPanels);
        outputSheets.push(worksheet);
        sheetNumber++;
      });
      firstPanel += part.quantity;
    });

    workbook.worksheets.slice().filter(sheet => sheet !== cover && !outputSheets.includes(sheet)).forEach(sheet => workbook.removeWorksheet(sheet.id));
    workbook.creator = 'Auto Vector Measure';
    workbook.modified = new Date();
    workbook.calcProperties = workbook.calcProperties || {};
    workbook.calcProperties.fullCalcOnLoad = true;
    workbook.calcProperties.forceFullCalc = true;

    const buffer = await workbook.xlsx.writeBuffer();
    const date = payload.completedDate ? payload.completedDate.replace(/-/g, '') : '';
    const name = [safeFileName(payload.style), safeFileName(payload.stage), date].filter(Boolean).join('-') || 'Quilting-Consumption';
    helpers.downloadBlob(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${name}.xlsx`);
    return buffer;
  };
})();
