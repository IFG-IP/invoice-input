(function () {
  const DIGITAL_TEXT_THRESHOLD = 20;
  const PDF_RENDER_SCALE = 3;
  const PDF_IMAGE_PAGE_LIMIT = 12;
  const PDF_TEXT_PAGE_LIMIT = 25;
  const PDF_TEXT_PROMPT_LIMIT = 14000;
  const PDF_RENDER_MAX_EDGE = 2200;
  const PDF_RENDER_JPEG_QUALITY = 0.88;
  const PDFJS_VERSION = "3.11.174";
  const PDFJS_CDN_BASE = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VERSION}`;
  const PDFJS_DIST_CDN_BASE = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}`;
  const PDFJS_WORKER_URL = `${PDFJS_CDN_BASE}/pdf.worker.min.js`;
  const PDFJS_CMAP_URL = `${PDFJS_DIST_CDN_BASE}/cmaps/`;
  const PDFJS_STANDARD_FONT_DATA_URL = `${PDFJS_DIST_CDN_BASE}/standard_fonts/`;
  const API_BASE_URL = normalizeApiBaseUrl(window.APP_CONFIG && window.APP_CONFIG.apiBaseUrl);
  const DEFAULT_TEMPLATE_URL = apiUrl("/api/template/extraction-fields");
  const IMAGE_TYPES = new Set(["jpg", "jpeg", "png", "heic", "heif"]);
  const EXTRACTION_FIELDS = [
    ["date", "日付"],
    ["vendor", "取引先"],
    ["total_amount", "合計金額"],
    ["tax_amount", "税額"],
    ["invoice_number", "請求書番号"],
    ["payment_due_date", "支払期日"],
    ["currency", "通貨"],
  ];
  let activeExtractionFields = EXTRACTION_FIELDS.map(function (field) {
    return {
      key: field[0],
      label: field[1],
      source: "default",
    };
  });
  let activeExtractionSchema = buildExtractionSchema(activeExtractionFields);

  const els = {
    aiResultRows: document.getElementById("aiResultRows"),
    aiStatus: document.getElementById("aiStatus"),
    backToUploadBottomButton: document.getElementById("backToUploadBottomButton"),
    backToUploadButton: document.getElementById("backToUploadButton"),
    clearButton: document.getElementById("clearButton"),
    copyReviewJsonButton: document.getElementById("copyReviewJsonButton"),
    dropZone: document.getElementById("dropZone"),
    excelTemplateStatus: document.getElementById("excelTemplateStatus"),
    exportSkyberryButton: document.getElementById("exportSkyberryButton"),
    expectedJsonInput: document.getElementById("expectedJsonInput"),
    fileInput: document.getElementById("fileInput"),
    folderInput: document.getElementById("folderInput"),
    imageCount: document.getElementById("imageCount"),
    jpegQuality: document.getElementById("jpegQuality"),
    jpegQualityValue: document.getElementById("jpegQualityValue"),
    libraryStatus: document.getElementById("libraryStatus"),
    maxEdge: document.getElementById("maxEdge"),
    maxEdgeValue: document.getElementById("maxEdgeValue"),
    ocrCount: document.getElementById("ocrCount"),
    pdfCount: document.getElementById("pdfCount"),
    progressBar: document.getElementById("progressBar"),
    progressCount: document.getElementById("progressCount"),
    progressCurrent: document.getElementById("progressCurrent"),
    progressPanel: document.getElementById("progressPanel"),
    progressTitle: document.getElementById("progressTitle"),
    resultRows: document.getElementById("resultRows"),
    reviewEditStatus: document.getElementById("reviewEditStatus"),
    reviewFileName: document.getElementById("reviewFileName"),
    reviewForm: document.getElementById("reviewForm"),
    reviewImage: document.getElementById("reviewImage"),
    reviewImageEmpty: document.getElementById("reviewImageEmpty"),
    reviewPages: document.getElementById("reviewPages"),
    reviewList: document.getElementById("reviewList"),
    reviewPage: document.getElementById("reviewPage"),
    reviewRouteLabel: document.getElementById("reviewRouteLabel"),
    modelInput: document.getElementById("modelInput"),
    runAiButton: document.getElementById("runAiButton"),
    saveReviewButton: document.getElementById("saveReviewButton"),
    selectedFileCount: document.getElementById("selectedFileCount"),
    selectedFileList: document.getElementById("selectedFileList"),
    selectedFileSummary: document.getElementById("selectedFileSummary"),
    totalCount: document.getElementById("totalCount"),
    uploadPage: document.getElementById("uploadPage"),
  };

  const state = {
    image: 0,
    currentReviewId: null,
    extractions: new Map(),
    items: [],
    isExtracting: false,
    isReadingFiles: false,
    nextSelectedFileId: 1,
    nextItemId: 1,
    ocr: 0,
    selectedFiles: [],
    pdf: 0,
    total: 0,
  };

  if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
    setLibraryStatus("PDF.js 使用可能", "ready");
  } else {
    setLibraryStatus("PDF.js 未読込", "warn");
  }

  els.clearButton.addEventListener("click", resetResults);

  els.selectedFileList.addEventListener("click", function (event) {
    if (!(event.target instanceof Element)) {
      return;
    }

    const button = event.target.closest("[data-remove-selected-file]");
    if (!button) {
      return;
    }

    event.preventDefault();
    removeSelectedFile(Number(button.dataset.removeSelectedFile));
  });

  els.runAiButton.addEventListener("click", function () {
    goToReviewWithExtraction();
  });

  els.backToUploadButton.addEventListener("click", function () {
    showUploadPage();
  });

  els.backToUploadBottomButton.addEventListener("click", function () {
    showUploadPage();
  });

  els.saveReviewButton.addEventListener("click", function () {
    saveReviewEdits();
  });

  els.copyReviewJsonButton.addEventListener("click", function () {
    copyReviewJson();
  });

  els.exportSkyberryButton.addEventListener("click", function () {
    exportSkyberryExcel();
  });

  loadDefaultExcelTemplate();
  loadApiDefaults();

  els.fileInput.addEventListener("change", function (event) {
    handleFiles(Array.from(event.target.files || []));
    event.target.value = "";
  });

  els.folderInput.addEventListener("change", function (event) {
    handleFiles(Array.from(event.target.files || []));
    event.target.value = "";
  });

  els.dropZone.addEventListener("click", function () {
    els.fileInput.click();
  });

  els.dropZone.addEventListener("keydown", function (event) {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      els.fileInput.click();
    }
  });

  els.maxEdge.addEventListener("input", function () {
    els.maxEdgeValue.value = `${els.maxEdge.value}px`;
  });

  els.jpegQuality.addEventListener("input", function () {
    els.jpegQualityValue.value = Number(els.jpegQuality.value).toFixed(2);
  });

  ["dragenter", "dragover"].forEach(function (eventName) {
    els.dropZone.addEventListener(eventName, function (event) {
      event.preventDefault();
      els.dropZone.classList.add("is-dragover");
    });
  });

  ["dragleave", "drop"].forEach(function (eventName) {
    els.dropZone.addEventListener(eventName, function (event) {
      event.preventDefault();
      els.dropZone.classList.remove("is-dragover");
    });
  });

  els.dropZone.addEventListener("drop", function (event) {
    collectDroppedFiles(event.dataTransfer).then(handleFiles);
  });

  async function handleFiles(files) {
    const targetFiles = files.filter(isSupportedFile);

    sendClientLog("files.selected", {
      selected: files.length,
      accepted: targetFiles.length,
      ignored: files.length - targetFiles.length,
    });

    if (targetFiles.length === 0) {
      if (state.selectedFiles.length === 0) {
        renderEmpty("対応ファイルが見つかりませんでした。");
      }
      return;
    }

    const selectedEntries = targetFiles.map(function (file) {
      const entry = {
        id: state.nextSelectedFileId,
        itemId: null,
        file,
        removed: false,
        row: null,
        status: "読込待ち",
        statusKey: "pending",
      };
      state.nextSelectedFileId += 1;
      return entry;
    });
    state.selectedFiles.push(...selectedEntries);
    recalculateSummary();
    renderSummary();
    renderSelectedFiles();
    state.isReadingFiles = true;
    updateNextButtonState();
    updateProgress({
      title: "ファイル読込中",
      current: 0,
      total: selectedEntries.length,
      currentText: "読込を開始しています。",
    });
    removeEmptyRow();

    for (let index = 0; index < selectedEntries.length; index += 1) {
      const entry = selectedEntries[index];
      if (entry.removed) {
        continue;
      }

      const file = entry.file;
      entry.status = "読込中";
      entry.statusKey = "reading";
      renderSelectedFiles();
      removeEmptyRow();
      const row = appendPendingRow(file);
      entry.row = row;
      updateProgress({
        title: "ファイル読込中",
        current: index,
        total: selectedEntries.length,
        currentText: displayName(file),
      });
      try {
        const result = await analyzeFile(file);
        if (entry.removed) {
          removeRow(row);
          recalculateSummary();
          renderSummary();
          continue;
        }

        result.id = state.nextItemId;
        state.nextItemId += 1;
        state.items.push(result);
        entry.itemId = result.id;
        entry.status = result.previewUrl ? "読込完了" : "読込完了（要確認）";
        entry.statusKey = result.previewUrl ? "ready" : "attention";
        updateRow(row, result);
        recalculateSummary();
        renderSelectedFiles();
        renderReviewList();
        logFileAnalysis(result);
      } catch (error) {
        state.ocr += 1;
        if (entry.removed) {
          removeRow(row);
          recalculateSummary();
          renderSummary();
          continue;
        }

        const failedResult = {
          id: state.nextItemId,
          file,
          route: "ocr",
          routeLabel: "要確認",
          note: error.message || "解析に失敗しました。",
          textLength: 0,
          textSample: "",
          previewUrl: "",
          normalizedBytes: 0,
        };
        state.nextItemId += 1;
        state.items.push(failedResult);
        entry.itemId = failedResult.id;
        entry.status = "読込失敗";
        entry.statusKey = "failed";
        updateRow(row, failedResult);
        recalculateSummary();
        renderSelectedFiles();
        renderReviewList();
        logFileAnalysis(failedResult, error);
      }
      updateProgress({
        title: "ファイル読込中",
        current: index + 1,
        total: selectedEntries.length,
        currentText: displayName(file),
      });
      renderSummary();
    }
    state.isReadingFiles = false;
    updateNextButtonState();
    recalculateSummary();
    renderSummary();
    renderSelectedFiles();
    ensureResultRowsNotEmpty();

    const remainingCount = selectedEntries.filter(function (entry) {
      return !entry.removed;
    }).length;
    updateProgress({
      title: "ファイル読込完了",
      current: remainingCount,
      total: remainingCount,
      currentText: remainingCount === 0
        ? "選択したファイルは削除されました。"
        : `${remainingCount}件の読込が完了しました。`,
    });
  }

  function resetResults() {
    revokeItemObjectUrls(state.items);
    state.total = 0;
    state.image = 0;
    state.currentReviewId = null;
    state.extractions = new Map();
    state.items = [];
    state.isExtracting = false;
    state.isReadingFiles = false;
    state.nextSelectedFileId = 1;
    state.nextItemId = 1;
    state.pdf = 0;
    state.ocr = 0;
    state.selectedFiles = [];
    els.fileInput.value = "";
    els.folderInput.value = "";
    renderSummary();
    renderSelectedFiles();
    hideProgress();
    updateNextButtonState();
    renderEmpty("ファイルを選択するとここに判定結果が表示されます。");
    renderAiEmpty("AI抽出結果がここに表示されます。");
    clearReviewPanel();
    setAiStatus("書類をアップロードしてから次へ進んでください。", "");
    showUploadPage();
    sendClientLog("files.cleared", {});
  }

  function renderEmpty(message) {
    els.resultRows.innerHTML = `<tr class="empty-row"><td colspan="5">${escapeHtml(message)}</td></tr>`;
  }

  function removeEmptyRow() {
    const emptyRow = els.resultRows.querySelector(".empty-row");
    if (emptyRow) {
      emptyRow.remove();
    }
  }

  function renderAiEmpty(message) {
    els.aiResultRows.innerHTML = `<tr class="empty-row"><td colspan="4">${escapeHtml(message)}</td></tr>`;
  }

  function ensureResultRowsNotEmpty() {
    if (els.resultRows.children.length === 0) {
      renderEmpty("ファイルを選択するとここに判定結果が表示されます。");
    }
  }

  function ensureAiRowsNotEmpty() {
    if (els.aiResultRows.children.length === 0) {
      renderAiEmpty("AI抽出結果がここに表示されます。");
    }
  }

  function removeRow(row) {
    if (row && row.parentElement) {
      row.remove();
    }
  }

  function removeSelectedFile(entryId) {
    if (state.isExtracting) {
      return;
    }

    const entry = state.selectedFiles.find(function (candidate) {
      return candidate.id === entryId;
    });
    if (!entry) {
      return;
    }

    entry.removed = true;
    state.selectedFiles = state.selectedFiles.filter(function (candidate) {
      return candidate.id !== entryId;
    });

    removeRow(entry.row);
    if (entry.itemId !== null) {
      removeItemArtifacts(entry.itemId);
    }

    recalculateSummary();
    renderSummary();
    renderSelectedFiles();
    renderReviewList();
    ensureResultRowsNotEmpty();
    ensureAiRowsNotEmpty();
    updateNextButtonState();

    if (state.selectedFiles.length === 0) {
      if (!state.isReadingFiles) {
        hideProgress();
      }
      setAiStatus("書類をアップロードしてから次へ進んでください。", "");
    }

    sendClientLog("file.removed", {
      file: displayName(entry.file),
      itemId: entry.itemId,
    });
  }

  function removeItemArtifacts(itemId) {
    const item = state.items.find(function (candidate) {
      return candidate.id === itemId;
    });
    revokeItemObjectUrl(item);
    state.items = state.items.filter(function (item) {
      return item.id !== itemId;
    });
    state.extractions.delete(itemId);

    els.resultRows.querySelectorAll(`tr[data-item-id="${itemId}"]`).forEach(removeRow);
    els.aiResultRows.querySelectorAll(`tr[data-item-id="${itemId}"]`).forEach(removeRow);

    if (state.currentReviewId === itemId) {
      clearReviewPanel();
    }
  }

  function revokeItemObjectUrls(items) {
    (items || []).forEach(revokeItemObjectUrl);
  }

  function revokeItemObjectUrl(item) {
    if (item && item.originalUrl) {
      URL.revokeObjectURL(item.originalUrl);
      item.originalUrl = "";
    }
  }

  function appendPendingRow(file) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>
        <span class="file-name">${escapeHtml(displayName(file))}</span>
        <span class="file-meta">${formatBytes(file.size)}</span>
      </td>
      <td><span class="spinner"></span>解析中</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
    `;
    els.resultRows.appendChild(tr);
    return tr;
  }

  async function analyzeFile(file) {
    const ext = extensionOf(file.name);
    if (ext === "pdf") {
      return analyzePdf(file);
    }

    if (IMAGE_TYPES.has(ext)) {
      return analyzeImage(file, "画像ルート", "image");
    }

    return {
      file,
      route: "ocr",
      routeLabel: "未対応",
      note: "対象外の拡張子です。",
      textLength: 0,
      textSample: "",
      previewUrl: "",
      originalUrl: "",
      pageCount: 0,
      pagePreviewCount: 0,
      normalizedBytes: 0,
    };
  }

  async function analyzePdf(file) {
    if (!window.pdfjsLib) {
      throw new Error("PDF.js を読み込めないためPDF解析をスキップしました。");
    }

    const bytes = new Uint8Array(await file.arrayBuffer());
    const originalUrl = URL.createObjectURL(file);
    const loadingTask = pdfjsLib.getDocument({
      data: bytes,
      disableFontFace: false,
      useSystemFonts: true,
      cMapUrl: PDFJS_CMAP_URL,
      cMapPacked: true,
      standardFontDataUrl: PDFJS_STANDARD_FONT_DATA_URL,
    });
    const pdf = await loadingTask.promise;
    const pageCount = pdf.numPages;
    const text = await extractPdfText(pdf);
    const normalizedText = normalizeText(text);
    const rawTextLength = pdfExtractedTextLength(normalizedText);
    const textIsUsable = rawTextLength >= DIGITAL_TEXT_THRESHOLD && isUsablePdfText(normalizedText);
    const extractedTextLength = textIsUsable ? rawTextLength : 0;
    const textSample = textIsUsable ? normalizedText : "";
    const pagePreviewUrls = await renderPdfPagesToImages(pdf, PDF_IMAGE_PAGE_LIMIT);
    const previewUrl = pagePreviewUrls[0] || "";
    const pagePreviewCount = pagePreviewUrls.length;

    if (extractedTextLength >= DIGITAL_TEXT_THRESHOLD) {
      state.pdf += 1;
      await pdf.destroy();
      return {
        file,
        route: "pdf",
        routeLabel: "デジタルPDF",
        note: pdfPageNote(pageCount, pagePreviewCount, "画像とPDF内部テキストを利用できます。"),
        textLength: extractedTextLength,
        textSample,
        previewUrl,
        originalUrl,
        pagePreviewUrls,
        pageCount,
        pagePreviewCount,
        normalizedBytes: 0,
      };
    }

    state.ocr += 1;
    state.image += 1;
    const normalizedBytes = await estimateDataUrlBytes(pagePreviewUrls);
    await pdf.destroy();
    return {
      file,
      route: "ocr",
      routeLabel: "画像PDF",
      note: pdfPageNote(pageCount, pagePreviewCount, "画像化結果を利用できます。"),
      textLength: extractedTextLength,
      textSample,
      previewUrl,
      originalUrl,
      pagePreviewUrls,
      pageCount,
      pagePreviewCount,
      normalizedBytes,
    };
  }

  function pdfPageNote(pageCount, pagePreviewCount, suffix) {
    if (!pageCount) {
      return `${pagePreviewCount}ページ分の${suffix}`;
    }

    if (pagePreviewCount >= pageCount) {
      return `全${pageCount}ページ分の${suffix}`;
    }

    return `全${pageCount}ページ中、先頭${pagePreviewCount}ページ分の${suffix}`;
  }

  async function extractPdfText(pdf) {
    const chunks = [];
    const pageLimit = Math.min(pdf.numPages, PDF_TEXT_PAGE_LIMIT);

    for (let pageNumber = 1; pageNumber <= pageLimit; pageNumber += 1) {
      let page = null;
      try {
        page = await pdf.getPage(pageNumber);
        const textContent = await page.getTextContent();
        const pageText = pdfTextContentToLines(textContent.items);
        chunks.push(`--- PDF ${pageNumber}ページ目 ---\n${pageText}`);
      } catch (error) {
        chunks.push(`--- PDF ${pageNumber}ページ目 ---\n[PDF内部テキスト取得失敗]`);
      } finally {
        if (page) {
          page.cleanup();
        }
      }
    }

    return chunks.join("\n");
  }

  function pdfExtractedTextLength(text) {
    return stripPdfTextMetadata(text).replace(/\s+/g, "").length;
  }

  function isUsablePdfText(text) {
    const compact = stripPdfTextMetadata(text).replace(/\s+/g, "");
    if (!compact) {
      return false;
    }

    const cidTextLength = ((compact.match(/\(cid:\d+\)/g) || []).join("")).length;
    const replacementLength = (compact.match(/[�□■]/g) || []).length;
    return (cidTextLength + replacementLength) / compact.length < 0.25;
  }

  function stripPdfTextMetadata(text) {
    return String(text || "")
      .replace(/--- PDF \d+ページ目 ---/g, "")
      .replace(/\[PDF内部テキスト取得失敗\]/g, "")
      .trim();
  }

  function pdfTextContentToLines(items) {
    const rows = [];
    (items || []).forEach(function (item) {
      const text = String(item.str || "").trim();
      if (!text) {
        return;
      }

      const transform = item.transform || [];
      const x = Number(transform[4] || 0);
      const y = Number(transform[5] || 0);
      let row = rows.find(function (candidate) {
        return Math.abs(candidate.y - y) <= 3;
      });
      if (!row) {
        row = { y, parts: [] };
        rows.push(row);
      }
      row.parts.push({ x, text });
    });

    return rows
      .sort(function (left, right) {
        return right.y - left.y;
      })
      .map(function (row) {
        return row.parts
          .sort(function (left, right) {
            return left.x - right.x;
          })
          .map(function (part) {
            return part.text;
          })
          .join(" ")
          .replace(/\s+/g, " ")
          .trim();
      })
      .filter(Boolean)
      .join("\n");
  }

  async function renderPdfPagesToImages(pdf, maxPages) {
    const urls = [];
    const pageLimit = Math.min(pdf.numPages, maxPages);
    for (let pageNumber = 1; pageNumber <= pageLimit; pageNumber += 1) {
      urls.push(await renderPdfPageToImage(pdf, pageNumber));
    }
    return urls;
  }

  async function renderPdfPageToImage(pdf, pageNumber) {
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale: PDF_RENDER_SCALE });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d", { alpha: true });
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);

    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    const renderOptions = {
      canvasContext: context,
      viewport,
      background: "#ffffff",
      intent: "display",
      annotationMode: (window.pdfjsLib && window.pdfjsLib.AnnotationMode)
        ? window.pdfjsLib.AnnotationMode.DISABLE
        : 0,
      renderInteractiveForms: false,
    };

    await page.render(renderOptions).promise;
    page.cleanup();
    return resizeCanvasToJpeg(canvas, PDF_RENDER_MAX_EDGE, PDF_RENDER_JPEG_QUALITY);
  }

  async function analyzeImage(file, routeLabel, route) {
    const imageFile = await normalizeHeicIfNeeded(file);
    const previewUrl = await resizeImageFile(imageFile);

    if (route === "image") {
      state.image += 1;
    }

    return {
      file,
      route,
      routeLabel,
      note: "Canvasでリサイズ・JPEG軽量化済みです。",
      textLength: 0,
      textSample: "",
      previewUrl,
      originalUrl: "",
      pagePreviewUrls: [previewUrl],
      pageCount: 1,
      pagePreviewCount: 1,
      normalizedBytes: await estimateImageBytes(previewUrl),
    };
  }

  async function normalizeHeicIfNeeded(file) {
    const ext = extensionOf(file.name);
    if (!["heic", "heif"].includes(ext)) {
      return file;
    }

    if (!window.heic2any) {
      throw new Error("HEIC変換ライブラリを読み込めませんでした。");
    }

    const blob = await heic2any({ blob: file, toType: "image/jpeg", quality: 0.86 });
    return new File([blob], file.name.replace(/\.(heic|heif)$/i, ".jpg"), { type: "image/jpeg" });
  }

  async function resizeImageFile(file) {
    const bitmap = await createImageBitmap(file);
    const canvas = document.createElement("canvas");
    const maxEdge = Number(els.maxEdge.value);
    const ratio = Math.min(1, maxEdge / Math.max(bitmap.width, bitmap.height));
    canvas.width = Math.max(1, Math.round(bitmap.width * ratio));
    canvas.height = Math.max(1, Math.round(bitmap.height * ratio));
    const context = canvas.getContext("2d", { alpha: false });

    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
    bitmap.close();

    return resizeCanvasToJpeg(canvas);
  }

  function resizeCanvasToJpeg(canvas, maxEdgeValue, qualityValue) {
    const maxEdge = Number(maxEdgeValue || els.maxEdge.value);
    const ratio = Math.min(1, maxEdge / Math.max(canvas.width, canvas.height));
    const output = document.createElement("canvas");
    output.width = Math.max(1, Math.round(canvas.width * ratio));
    output.height = Math.max(1, Math.round(canvas.height * ratio));
    const outputContext = output.getContext("2d", { alpha: false });

    outputContext.fillStyle = "#ffffff";
    outputContext.fillRect(0, 0, output.width, output.height);
    outputContext.drawImage(canvas, 0, 0, output.width, output.height);

    return output.toDataURL("image/jpeg", Number(qualityValue || els.jpegQuality.value));
  }

  async function estimateImageBytes(dataUrl) {
    if (!dataUrl) {
      return 0;
    }
    const response = await fetch(dataUrl);
    const blob = await response.blob();
    return blob.size;
  }

  async function estimateDataUrlBytes(dataUrls) {
    const sizes = await Promise.all((dataUrls || []).map(estimateImageBytes));
    return sizes.reduce(function (total, size) {
      return total + size;
    }, 0);
  }

  function updateRow(row, result) {
    row.dataset.itemId = String(result.id);
    row.classList.add("selectable-row");
    row.addEventListener("click", function () {
      selectReviewItem(result.id);
    });
    const preview = result.previewUrl
      ? `<img class="preview" src="${result.previewUrl}" alt="${escapeHtml(displayName(result.file))}のプレビュー" />`
      : "-";
    const sample = result.textSample
      ? `<span class="text-sample">${escapeHtml(result.textSample.slice(0, 1200))}</span>`
      : "";
    const sizeNote = result.normalizedBytes
      ? `<span class="route-note">軽量化後 ${formatBytes(result.normalizedBytes)}</span>`
      : "";

    row.innerHTML = `
      <td>
        <span class="file-name">${escapeHtml(displayName(result.file))}</span>
        <span class="file-meta">${formatBytes(result.file.size)}</span>
      </td>
      <td><span class="badge ${result.route}">${escapeHtml(result.routeLabel)}</span></td>
      <td>${result.textLength}文字${sample}</td>
      <td>
        <span>${escapeHtml(result.note)}</span>
        ${sizeNote}
      </td>
      <td>${preview}</td>
    `;
  }

  function logFileAnalysis(result, error) {
    sendClientLog("file.analysis.completed", {
      file: displayName(result.file),
      size: result.file.size,
      route: result.route,
      routeLabel: result.routeLabel,
      textLength: result.textLength,
      note: result.note,
      normalizedBytes: result.normalizedBytes,
      textSample: result.textSample ? result.textSample.slice(0, 240) : "",
      error: error ? error.message || String(error) : null,
    });
  }

  async function goToReviewWithExtraction() {
    const completed = await runAiExtraction();
    if (completed) {
      showReviewPage();
    }
  }

  async function runAiExtraction() {
    const model = els.modelInput.value.trim() || "gemini-2.5-flash";
    const candidates = state.items.filter(function (item) {
      return Boolean(item.previewUrl);
    });

    if (candidates.length === 0) {
      setAiStatus("先に請求書・領収書をアップロードしてください。", "error");
      sendClientLog("ai.skipped", { reason: "no_preview" });
      return false;
    }

    let expectedValues = {};
    try {
      expectedValues = parseExpectedJson();
      await ensureApiProxyReady();
    } catch (error) {
      setAiStatus(error.message, "error");
      return false;
    }

    state.isExtracting = true;
    updateNextButtonState();
    els.runAiButton.textContent = "処理中...";
    els.aiResultRows.innerHTML = "";
    setAiStatus(`${candidates.length}件を処理中です。`, "");
    updateProgress({
      title: "AI抽出中",
      current: 0,
      total: candidates.length,
      currentText: "抽出を開始しています。",
    });
    renderSelectedFiles();
    sendClientLog("ai.batch.started", {
      count: candidates.length,
      model,
    });

    let compared = 0;
    let matched = 0;
    let succeeded = 0;
    let lastErrorMessage = "";

    for (let index = 0; index < candidates.length; index += 1) {
      const item = candidates[index];
      const row = appendAiPendingRow(item);
      updateProgress({
        title: "AI抽出中",
        current: index,
        total: candidates.length,
        currentText: displayName(item.file),
      });
      try {
        const extraction = await extractFieldsWithGemini(item, model);
        const expected = findExpectedForItem(expectedValues, item);
        const comparison = compareExtraction(extraction, expected);
        state.extractions.set(item.id, extraction);
        succeeded += 1;
        updateSelectedFileStatus(item.id, "AI抽出済み", "extracted");
        renderReviewList();
        compared += comparison.compared;
        matched += comparison.matched;
        updateAiRow(row, item, extraction, comparison);
        sendClientLog("ai.extraction.completed", {
          file: displayName(item.file),
          route: item.routeLabel,
          extraction,
          comparison,
        });
        if (state.currentReviewId === null) {
          selectReviewItem(item.id);
        }
      } catch (error) {
        lastErrorMessage = error.message || "AI抽出に失敗しました。";
        updateAiErrorRow(row, item, error);
        updateSelectedFileStatus(item.id, "AI失敗", "failed");
        renderReviewList();
        sendClientLog("ai.extraction.failed", {
          file: displayName(item.file),
          route: item.routeLabel,
          error: lastErrorMessage,
        });
      }
      updateProgress({
        title: "AI抽出中",
        current: index + 1,
        total: candidates.length,
        currentText: displayName(item.file),
      });
    }

    state.isExtracting = false;
    updateNextButtonState();
    els.runAiButton.textContent = "次へ";
    renderSelectedFiles();

    if (succeeded === 0) {
      const failureMessage = lastErrorMessage
        ? `抽出できた書類がありませんでした。最後のエラー: ${lastErrorMessage}`
        : "抽出できた書類がありませんでした。";
      setAiStatus(failureMessage, "error");
      updateProgress({
        title: "AI抽出失敗",
        current: candidates.length,
        total: candidates.length,
        currentText: failureMessage,
      });
      return false;
    }

    if (compared > 0) {
      const score = Math.round((matched / compared) * 100);
      setAiStatus(`処理が完了しました。一致率 ${score}% (${matched}/${compared})`, "ready");
      sendClientLog("ai.batch.completed", { compared, matched, score });
    } else {
      setAiStatus("処理が完了しました。目視確認で内容を確認してください。", "ready");
      sendClientLog("ai.batch.completed", { compared, matched });
    }
    updateProgress({
      title: "AI抽出完了",
      current: candidates.length,
      total: candidates.length,
      currentText: `${succeeded}件の抽出が完了しました。`,
    });
    return true;
  }

  async function extractFieldsWithGemini(item, model) {
    const response = await fetch(apiUrl(`/api/gemini/generate?model=${encodeURIComponent(model)}`), {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        contents: [
          {
            role: "user",
            parts: geminiPartsForItem(item),
          },
        ],
        generationConfig: {
          responseMimeType: "application/json",
          responseSchema: activeExtractionSchema,
          maxOutputTokens: 2400,
          temperature: 0,
          topK: 1,
          topP: 0.1,
          candidateCount: 1,
        },
      }),
    });

    const payload = await response.json().catch(function () {
      return {};
    });

    if (!response.ok) {
      const message = payload.error && payload.error.message
        ? payload.error.message
        : `Gemini API error: ${response.status}`;
      throw new Error(message);
    }

    const text = extractResponseText(payload);
    return normalizeExtractionResult(JSON.parse(text), item);
  }

  function geminiPartsForItem(item) {
    return [
      {
        text: buildExtractionPrompt(item),
      },
      ...imageDataPartsForItem(item),
    ];
  }

  function imageDataPartsForItem(item) {
    const urls = (Array.isArray(item.pagePreviewUrls) && item.pagePreviewUrls.length > 0)
      ? item.pagePreviewUrls
      : [item.previewUrl];

    return urls.filter(Boolean).map(dataUrlToGeminiInlineData);
  }

  function normalizeExtractionResult(value, item) {
    const extraction = value && typeof value === "object" && !Array.isArray(value)
      ? { ...value }
      : {};

    activeExtractionFields.forEach(function (field) {
      const key = fieldKey(field);
      if (!Object.prototype.hasOwnProperty.call(extraction, key)) {
        extraction[key] = null;
      }
      if (keyMatches(key, "vendor")) {
        extraction[key] = normalizeVendorName(extraction[key], item);
      }
      if (keyMatches(key, "receipt_method")) {
        extraction[key] = normalizeReceiptMethod(extraction[key], item);
      }
      if (keyMatches(key, "payment_method")) {
        extraction[key] = normalizePaymentMethod(extraction[key], extraction, item);
      }
      if (keyMatches(key, "bank_account_type")) {
        extraction[key] = normalizeBankAccountType(extraction[key]);
      }
      if (keyMatches(key, "bank_account_number")) {
        extraction[key] = normalizeBankAccountNumber(extraction[key]);
      }
      if (keyMatches(key, "bank_account_kana")) {
        extraction[key] = normalizeBankAccountKana(extraction[key]);
      }
      if (key === "row_no") {
        extraction[key] = null;
      }
    });

    applyDocumentTypeHints(extraction, item);
    fillExplicitStructuredFields(extraction, item);
    cleanPaymentContentFields(extraction, item);
    fillExplicitPaymentContentFields(extraction, item);
    fixVendorFromTextSample(extraction, item);
    fillSummaryFields(extraction);
    return extraction;
  }

  function normalizeVendorName(value, item) {
    if (isBlank(value)) {
      return null;
    }

    const addresseeSideLabelPattern = /^(宛先|請求先|請求者名?|納品先|お客様名|利用者名|支払者|支払い元|買い手|請求先会社|支払元会社)\s*(?:[：:]|\s)\s*/;
    const issuerSideLabelPattern = /^(請求元|請求者|発行者|発行元|請求会社|請求元会社|発行会社|請求書発行元)\s*(?:[：:]|\s)\s*/;
    const source = String(value).trim();
    if (addresseeSideLabelPattern.test(source)) {
      return null;
    }
    if (hasAddresseeHonorific(source)) {
      return null;
    }

    const labelPattern = /^(支払先名?|支払先|お支払先|振込先|振込先名|振込先名義|受取人名?|受取人名義|口座名義|領収者|店舗名|店名|加盟店名?|取引先|仕入先|販売者|事業者名|請求元|発行者|発行元|請求会社|請求元会社|発行会社|請求書発行元)\s*[：:]\s*/;
    const normalized = String(value)
      .replace(/[\r\n\t]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .replace(issuerSideLabelPattern, "")
      .replace(labelPattern, "")
      .trim();

    return normalized || null;
  }

  function hasAddresseeHonorific(value) {
    return /\s*(御中|様|殿)\s*$/.test(String(value).trim());
  }

  function fixVendorFromTextSample(extraction, item) {
    const vendor = firstValueByBaseKey(extraction, "vendor");
    if (!item || isBlank(item.textSample)) {
      return;
    }

    const text = String(item.textSample);
    const honorificCandidates = extractHonorificCompanyCandidates(text);
    const explicitPayee = extractLabeledPayeeCompanyCandidate(text, item);
    if (isBlank(vendor)) {
      const directDebitReplacement = isDirectDebitDocument(item)
        ? extractPayeeCompanyCandidate(text, honorificCandidates, item)
        : null;
      const replacement = explicitPayee || directDebitReplacement;
      if (isBlank(replacement)) {
        return;
      }
      Object.keys(extraction).forEach(function (key) {
        if (keyMatches(key, "vendor")) {
          extraction[key] = replacement;
        }
      });
      return;
    }

    if (!isBlank(explicitPayee) && !sameEntityText(explicitPayee, vendor)) {
      Object.keys(extraction).forEach(function (key) {
        if (keyMatches(key, "vendor")) {
          extraction[key] = explicitPayee;
        }
      });
      return;
    }

    const vendorCameFromAddressee = honorificCandidates.some(function (candidate) {
      return sameEntityText(candidate, vendor);
    });
    if (!vendorCameFromAddressee) {
      return;
    }

    const replacement = extractPayeeCompanyCandidate(text, honorificCandidates, item);
    Object.keys(extraction).forEach(function (key) {
      if (keyMatches(key, "vendor")) {
        extraction[key] = replacement || null;
      }
    });
  }

  function extractHonorificCompanyCandidates(text) {
    const candidates = [];
    const pattern = /([^\s　,，、。:：]{2,40}(?:株式会社|有限会社|合同会社|合名会社|合資会社)|(?:株式会社|有限会社|合同会社|合名会社|合資会社)[^\s　,，、。:：]{2,40})\s*(?:御中|様|殿)/g;
    let match;
    while ((match = pattern.exec(text)) !== null) {
      candidates.push(match[1]);
    }
    return candidates;
  }

  function extractPayeeCompanyCandidate(text, excludedCandidates, item) {
    const labeled = extractLabeledPayeeCompanyCandidate(text, item);
    if (!isBlank(labeled) && !isExcludedCompanyCandidate(labeled, excludedCandidates)) {
      return labeled;
    }

    const pattern = /([^\s　,，、。:：]{2,40}(?:株式会社|有限会社|合同会社|合名会社|合資会社)|(?:株式会社|有限会社|合同会社|合名会社|合資会社)[^\s　,，、。:：]{2,40})/g;
    let match;
    while ((match = pattern.exec(text)) !== null) {
      const candidate = match[1];
      const afterCandidate = text.slice(match.index + match[0].length);
      if (/^\s*(?:御中|様|殿)/.test(afterCandidate)) {
        continue;
      }
      if (!isExcludedCompanyCandidate(candidate, excludedCandidates)) {
        return candidate;
      }
    }

    return null;
  }

  function extractLabeledPayeeCompanyCandidate(text, item) {
    const payeeLabels = "支払先名?|お支払先|振込先名義|振込先名|振込先|お振込先|受取人名義|受取人名|受取人|口座名義";
    const payee = extractCompanyAfterLabels(text, payeeLabels);
    if (!isBlank(payee)) {
      return payee;
    }

    const fallbackLabels = isDirectDebitDocument(item)
      ? "領収者|販売者|事業者名|店舗名|店名|加盟店名?|発行者|発行元|請求元"
      : "領収者|販売者|事業者名|店舗名|店名|加盟店名?";
    return extractCompanyAfterLabels(text, fallbackLabels);
  }

  function extractCompanyAfterLabels(text, labelPattern) {
    const pattern = new RegExp(`(?:${labelPattern})\\s*[：:\\s]\\s*([^\\s　,，、。:：]{2,40}(?:株式会社|有限会社|合同会社|合名会社|合資会社)|(?:株式会社|有限会社|合同会社|合名会社|合資会社)[^\\s　,，、。:：]{2,40})`);
    const match = String(text || "").match(pattern);
    return match ? match[1] : null;
  }

  function isExcludedCompanyCandidate(candidate, excludedCandidates) {
    return excludedCandidates.some(function (excluded) {
      return sameEntityText(candidate, excluded);
    });
  }

  function normalizeReceiptMethod(value, item) {
    const source = String(value || "").trim();
    const compact = normalizeHeader(source);

    if (/^3$|紙契約書/.test(compact)) {
      return "3";
    }
    if (/^4$|電子契約書/.test(compact)) {
      return "4";
    }

    return defaultReceiptMethod(item);
  }

  function defaultReceiptMethod(item) {
    if (!item) {
      return null;
    }

    if (item.route === "pdf") {
      return "2";
    }

    if (item.route === "image" || item.routeLabel === "画像PDF") {
      return "1";
    }

    return null;
  }

  function normalizePaymentMethod(value, extraction, item) {
    const source = String(value || "").trim();
    const compact = normalizeHeader(source);

    if (/^5$|スポット/.test(compact)) {
      return "5";
    }
    if (/^6$|クレジット.*自動|自動.*クレジット|カード.*自動|継続.*クレジット/.test(compact)) {
      return "6";
    }
    if (/^7$|クレジット.*即日|即日.*クレジット|カード.*即日|カード決済|クレジットカード決済/.test(compact)) {
      return "7";
    }
    if (/^4$|海外送金|海外|外国送金|国際送金|SWIFT|WIRE|IBAN/i.test(compact)) {
      return "4";
    }
    if (/^3$|口座引落|口座振替|自動引落|引落|引き落|口振/.test(compact)) {
      return "3";
    }
    if (/^2$|納付書|払込票|振込用紙|支払書/.test(compact)) {
      return "2";
    }
    if (/^1$|銀行振込|振込|お振込|銀行送金/.test(compact)) {
      return "1";
    }

    return defaultPaymentMethod(extraction, item);
  }

  function defaultPaymentMethod(extraction, item) {
    if (isDirectDebitDocument(item)) {
      return "3";
    }

    if (!extraction || typeof extraction !== "object") {
      return null;
    }

    const explicitValue = firstValueByBaseKey(extraction, "payment_method");
    if (!isBlank(explicitValue)) {
      const explicitMethod = normalizePaymentMethod(explicitValue, null, item);
      if (!isBlank(explicitMethod)) {
        return explicitMethod;
      }
    }

    if (!isBlank(extraction.payment_card_name) || !isBlank(extraction.credit_card_application_no)) {
      return "6";
    }

    if (hasBankTransferDetails(extraction)) {
      return hasForeignCurrencyAmount(extraction) ? "4" : "1";
    }

    return null;
  }

  function applyDocumentTypeHints(extraction, item) {
    if (!isDirectDebitDocument(item)) {
      return;
    }

    Object.keys(extraction).forEach(function (key) {
      if (keyMatches(key, "payment_method")) {
        extraction[key] = "3";
      }
    });
  }

  function fillExplicitStructuredFields(extraction, item) {
    if (!item || isBlank(item.textSample)) {
      return;
    }

    const text = String(item.textSample);
    setBlankExtractionValue(extraction, "invoice_number", extractExplicitInvoiceNumber(text));
    setBlankExtractionValue(extraction, "payment_due_date", extractExplicitPaymentDueDate(text));
    setBlankExtractionValue(extraction, "bank_account_type", extractExplicitBankAccountType(text), normalizeBankAccountType);
    setBlankExtractionValue(extraction, "bank_account_number", extractExplicitBankAccountNumber(text), normalizeBankAccountNumber);
    setBlankExtractionValue(extraction, "bank_account_kana", extractExplicitBankAccountKana(text), normalizeBankAccountKana);

    const explicitMethod = extractExplicitPaymentMethod(text, extraction, item) ||
      normalizePaymentMethod(null, extraction, item);
    setBlankExtractionValue(extraction, "payment_method", explicitMethod, function (value) {
      return normalizePaymentMethod(value, extraction, item);
    });
  }

  function setBlankExtractionValue(extraction, baseKey, value, normalizer) {
    if (isBlank(value)) {
      return;
    }

    Object.keys(extraction).forEach(function (key) {
      if (!keyMatches(key, baseKey) || !isBlank(extraction[key])) {
        return;
      }
      const normalized = typeof normalizer === "function" ? normalizer(value) : value;
      extraction[key] = isBlank(normalized) ? null : normalized;
    });
  }

  function extractExplicitInvoiceNumber(text) {
    const labelPattern = [
      "ご請求書\\s*(?:No\\.?|NO|№|番号|ID)",
      "御請求書\\s*(?:No\\.?|NO|№|番号|ID)",
      "請求書\\s*(?:No\\.?|NO|番号)",
      "請求書\\s*(?:№|#|ID)",
      "ご請求\\s*(?:No\\.?|NO|№|番号|ID)",
      "御請求\\s*(?:No\\.?|NO|№|番号|ID)",
      "請求\\s*(?:No\\.?|NO|番号)",
      "請求\\s*(?:№|#|ID)",
      "領収書\\s*(?:No\\.?|NO|番号)",
      "領収書\\s*(?:№|#|ID)",
      "伝票\\s*(?:No\\.?|NO|№)",
      "伝票番号",
      "管理番号",
      "発行番号",
      "発行\\s*(?:No\\.?|NO|№)",
      "Invoice\\s*(?:No\\.?|Number|#|ID)",
      "Receipt\\s*(?:No\\.?|Number)",
      "Document\\s*(?:No\\.?|Number)",
    ].join("|");
    const candidates = extractLabeledCandidates(text, labelPattern, { allowNextLine: true });
    for (const candidate of candidates) {
      const cleaned = cleanInvoiceNumberCandidate(candidate);
      if (!isBlank(cleaned)) {
        return cleaned;
      }
    }
    return extractInvoiceNumberNearTitle(text);
  }

  function cleanInvoiceNumberCandidate(value) {
    const source = stripCandidateTail(
      toHalfWidthAscii(value),
      /(発行日|請求日|日付|支払|金額|合計|登録番号|事業者番号|適格請求書発行事業者|TEL|電話|銀行|支店|口座|〒)/
    ).replace(/^[：:\-－=\s]+/, "").trim();
    const match = source.match(/[A-Za-z0-9?][A-Za-z0-9?._/-]{1,50}/);
    if (!match) {
      return null;
    }
    const candidate = match[0].replace(/[.,、。]+$/, "");
    if (isLikelyNonInvoiceNumber(candidate, source)) {
      return null;
    }
    return candidate;
  }

  function extractInvoiceNumberNearTitle(text) {
    const lines = textLines(text).slice(0, 100);
    for (let index = 0; index < lines.length; index += 1) {
      const line = toHalfWidthAscii(lines[index]);
      const context = [
        lines[index - 2] || "",
        lines[index - 1] || "",
        lines[index],
        lines[index + 1] || "",
      ].join(" ");

      if (!/(請求書|御請求書|ご請求書|請求|領収書|INVOICE|Invoice|Receipt)/i.test(context)) {
        continue;
      }

      const labeled = line.match(/(?:No\.?|NO|№|#|番号)\s*[：:#\-－]?\s*([A-Za-z0-9?][A-Za-z0-9?._/-]{1,50})/i);
      if (labeled) {
        const cleaned = cleanInvoiceNumberCandidate(labeled[1]);
        if (!isBlank(cleaned)) {
          return cleaned;
        }
      }

      const compactTitle = line.match(/(?:請求書|御請求書|ご請求書|INVOICE|Invoice)[^\w?]{0,8}([A-Za-z0-9?][A-Za-z0-9?._/-]{2,50})/);
      if (compactTitle) {
        const cleaned = cleanInvoiceNumberCandidate(compactTitle[1]);
        if (!isBlank(cleaned)) {
          return cleaned;
        }
      }
    }

    return null;
  }

  function isLikelyNonInvoiceNumber(candidate, source) {
    const value = String(candidate || "");
    const context = String(source || "");
    if (/^\d{10,}$/.test(value) && /(TEL|電話|FAX|登録番号|事業者番号|〒)/i.test(context)) {
      return true;
    }
    if (/^T\d{13}$/i.test(value)) {
      return true;
    }
    if (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$/.test(value)) {
      return true;
    }
    if (/^\d{1,3}(?:,\d{3})+$/.test(value)) {
      return true;
    }
    return false;
  }

  function extractExplicitPaymentDueDate(text) {
    const labelPattern = [
      "支払期日",
      "支払期限",
      "支払日",
      "お支払期限",
      "お支払い期限",
      "お支払期日",
      "振込期限",
      "入金期日",
      "振替日",
      "口座振替日",
      "口座振替予定日",
      "引落日",
      "引き落とし日",
      "引落予定日",
      "Due\\s*Date",
      "Payment\\s*(?:Due|Deadline|Due\\s*Date)",
    ].join("|");
    const candidates = extractLabeledCandidates(text, labelPattern, { allowNextLine: true });
    for (const candidate of candidates) {
      const normalized = normalizeFlexibleDateCandidate(candidate);
      if (!isBlank(normalized)) {
        return normalized;
      }
    }
    return null;
  }

  function normalizeFlexibleDateCandidate(value) {
    const source = toHalfWidthAscii(value).replace(/\s+/g, " ").trim();
    const eraMatch = source.match(/(令和|平成|昭和|R|H|S)\s*(元|\d{1,2})\s*年?\s*[\/.-]?\s*(\d{1,2})\s*(?:月|[\/.-])\s*(\d{1,2})\s*日?/i);
    if (eraMatch) {
      const era = eraMatch[1].toUpperCase();
      const eraYear = eraMatch[2] === "元" ? 1 : Number(eraMatch[2]);
      const baseYear = era === "令和" || era === "R" ? 2018 : era === "平成" || era === "H" ? 1988 : 1925;
      return formatDateParts(baseYear + eraYear, eraMatch[3], eraMatch[4]);
    }

    const westernMatch = source.match(/(\d{4})\s*(?:年|[\/.-])\s*(\d{1,2})\s*(?:月|[\/.-])\s*(\d{1,2})\s*日?/);
    if (westernMatch) {
      return formatDateParts(westernMatch[1], westernMatch[2], westernMatch[3]);
    }

    return null;
  }

  function formatDateParts(year, month, day) {
    const y = String(year).padStart(4, "0");
    const m = String(month).padStart(2, "0");
    const d = String(day).padStart(2, "0");
    if (Number(m) < 1 || Number(m) > 12 || Number(d) < 1 || Number(d) > 31) {
      return null;
    }
    return `${y}-${m}-${d}`;
  }

  function extractExplicitBankAccountType(text) {
    const labelPattern = "口座種別|預金種目|預金種別|科目|種別|Account\\s*Type";
    const candidates = extractLabeledCandidates(text, labelPattern, { allowNextLine: true });
    for (const candidate of candidates) {
      const normalized = normalizeBankAccountType(candidate);
      if (!isBlank(normalized)) {
        return normalized;
      }
    }

    const bankSection = extractBankInfoSection(text);
    const match = normalizeHeader(bankSection).match(/外貨普通|外貨定期|普通預金|当座預金|定期預金|貯蓄預金|別段預金|普通|当座|貯蓄|定期|別段|その他/);
    return match ? normalizeBankAccountType(match[0]) : null;
  }

  function extractExplicitBankAccountNumber(text) {
    const labelPattern = "口座番号|口座\\s*(?:No\\.?|NO\\.?)|Account\\s*(?:Number|No\\.?|#)|A/C\\s*No\\.?";
    const candidates = extractLabeledCandidates(text, labelPattern, { allowNextLine: true });
    for (const candidate of candidates) {
      const cleaned = normalizeBankAccountNumber(candidate);
      if (!isBlank(cleaned)) {
        return cleaned;
      }
    }

    const bankSection = extractBankInfoSection(text);
    const match = toHalfWidthAscii(bankSection).match(/(?:普通|当座|定期|別段|口座)\s*(?:預金)?\s*(?:番号)?\s*([0-9]{5,12})/);
    return match ? match[1] : null;
  }

  function cleanBankAccountNumberCandidate(value) {
    const source = stripCandidateTail(
      toHalfWidthAscii(value),
      /(口座名義|名義|銀行|支店|支払|請求|金額|TEL|電話|FAX|摘要|備考)/
    );
    const exact = source.match(/(?:^|[^\d])(\d{4,12})(?!\d)/);
    if (exact) {
      return exact[1];
    }

    const uncertain = source.match(/(?:^|[^0-9?？＊*xX])([0-9?？＊*xX][0-9?？＊*xX\s\-－]{3,18})(?![0-9?？＊*xX])/);
    if (!uncertain) {
      return null;
    }

    const candidate = uncertain[1]
      .replace(/[？＊*xX]/g, "?")
      .replace(/[\s\-－]/g, "");
    const digitCount = (candidate.match(/\d/g) || []).length;
    return digitCount >= 3 && /^[0-9?]{4,12}$/.test(candidate) ? candidate : null;
  }

  function extractExplicitBankAccountKana(text) {
    const labelPattern = [
      "口座名義フリガナ",
      "口座名義フリ仮名",
      "口座名義カナ",
      "口座名義ﾌﾘｶﾞﾅ",
      "口座名義ﾌﾘ仮名",
      "口座名義ｶﾅ",
      "名義フリガナ",
      "名義フリ仮名",
      "名義カナ",
      "名義ﾌﾘｶﾞﾅ",
      "名義ﾌﾘ仮名",
      "名義ｶﾅ",
      "受取人名カナ",
      "受取人名義カナ",
      "受取人名ｶﾅ",
      "受取人名義ｶﾅ",
      "振込先名義カナ",
      "振込先名義ｶﾅ",
      "カナ名義",
      "ｶﾅ名義",
      "フリガナ",
      "ﾌﾘｶﾞﾅ",
    ].join("|");
    const candidates = extractLabeledCandidates(text, labelPattern, { allowNextLine: true });
    for (const candidate of candidates) {
      const cleaned = cleanBankKanaCandidate(candidate);
      if (!isBlank(cleaned)) {
        return cleaned;
      }
    }

    const bankSection = extractBankInfoSection(text);
    const lines = textLines(bankSection);
    for (const line of lines) {
      const cleaned = cleanBankKanaCandidate(line);
      if (!isBlank(cleaned) && !/(銀行|支店|口座番号|普通|当座|貯蓄|定期|別段)/.test(cleaned)) {
        return cleaned;
      }
    }
    return null;
  }

  function cleanBankKanaCandidate(value) {
    const source = stripCandidateTail(
      String(value || ""),
      /(口座番号|口座種別|預金種目|銀行名|銀行|支店名|支店|支払|請求|金額|TEL|電話|FAX|摘要|備考)/
    )
      .replace(/^(口座名義フリガナ|口座名義フリ仮名|口座名義カナ|口座名義ﾌﾘｶﾞﾅ|口座名義ｶﾅ|名義フリガナ|名義カナ|名義ﾌﾘｶﾞﾅ|名義ｶﾅ|受取人名カナ|受取人名ｶﾅ|振込先名義カナ|振込先名義ｶﾅ|フリガナ|ﾌﾘｶﾞﾅ|カナ名義|ｶﾅ名義)\s*[：:\-－=]?\s*/, "")
      .replace(/^[：:\-－=\s]+/, "")
      .replace(/[、。，,]+$/, "")
      .trim();

    if (!/[ァ-ヶーｦ-ﾟｰ]/.test(source)) {
      return null;
    }
    return source;
  }

  function extractExplicitPaymentMethod(text, extraction, item) {
    const labelPattern = "支払方法|お支払方法|支払い方法|決済方法|決済種別|Payment\\s*Method";
    const candidates = extractLabeledCandidates(text, labelPattern, { allowNextLine: true });
    for (const candidate of candidates) {
      const normalized = normalizePaymentMethod(candidate, extraction, item);
      if (!isBlank(normalized)) {
        return normalized;
      }
    }

    const compact = normalizeHeader(text);
    if (/口座振替請求書兼領収書|口座振替請求書|口座振替領収書|口座引落|口座振替|自動引落|引落|引き落/.test(compact)) {
      return "3";
    }
    if (/クレジット.*自動|自動.*クレジット|カード.*自動|継続.*クレジット/.test(compact)) {
      return "6";
    }
    if (/クレジット.*即日|即日.*クレジット|カード.*即日|カード決済|クレジットカード決済/.test(compact)) {
      return "7";
    }
    if (/海外送金|外国送金|国際送金|SWIFT|WIRE|IBAN/i.test(compact)) {
      return "4";
    }
    if (/納付書|払込票|振込用紙|支払書/.test(compact)) {
      return "2";
    }
    if (/振込先|お振込先|銀行名|金融機関名|支店名|口座番号|口座名義|普通|当座|貯蓄/.test(compact)) {
      return "1";
    }
    return null;
  }

  function extractLabeledCandidates(text, labelPattern, options) {
    const settings = options || {};
    const lines = textLines(text);
    const pattern = new RegExp(`(?:^|[\\s　])(?:${labelPattern})\\s*(?:[：:\\-－=]|\\s{1,}|(?=[A-Za-z0-9ァ-ヶｦ-ﾟ令平昭]))\\s*(.*)$`, "i");
    const candidates = [];

    lines.forEach(function (line, index) {
      const normalizedLine = toHalfWidthAscii(line).replace(/\u00a0/g, " ").trim();
      const match = normalizedLine.match(pattern);
      if (!match) {
        return;
      }

      const value = (match[1] || "").trim();
      if (value) {
        candidates.push(value);
        return;
      }

      if (settings.allowNextLine && lines[index + 1]) {
        candidates.push(lines[index + 1].trim());
      }
    });

    return candidates;
  }

  function textLines(text) {
    return normalizeText(text)
      .split(/\n+/)
      .map(function (line) {
        return line.trim();
      })
      .filter(Boolean);
  }

  function extractBankInfoSection(text) {
    const lines = textLines(text);
    const selected = [];
    lines.forEach(function (line, index) {
      if (!/(振込先|お振込先|銀行名|金融機関|支店名|口座|預金|受取人|名義)/.test(line)) {
        return;
      }
      for (let offset = -1; offset <= 2; offset += 1) {
        const candidate = lines[index + offset];
        if (candidate && !selected.includes(candidate)) {
          selected.push(candidate);
        }
      }
    });
    return selected.join("\n");
  }

  function stripCandidateTail(value, stopPattern) {
    const source = String(value || "").trim();
    const index = source.search(stopPattern);
    return index > 0 ? source.slice(0, index).trim() : source;
  }

  function toHalfWidthAscii(value) {
    return String(value || "")
      .replace(/[！-～]/g, function (char) {
        return String.fromCharCode(char.charCodeAt(0) - 0xFEE0);
      })
      .replace(/　/g, " ");
  }

  function isDirectDebitDocument(item) {
    if (!item) {
      return false;
    }

    const source = normalizeHeader([
      item.textSample || "",
      item.note || "",
    ].join(" "));

    return /口座振替請求書兼領収書|口座振替請求書|口座振替領収書|口座引落請求書|口座引落領収書|自動引落請求書|自動引落領収書/.test(source);
  }

  function firstValueByBaseKey(source, baseKey) {
    const key = Object.keys(source).find(function (candidate) {
      return keyMatches(candidate, baseKey) && !isBlank(source[candidate]);
    });
    return key ? source[key] : null;
  }

  function hasBankTransferDetails(extraction) {
    return [
      "bank_name",
      "bank_branch_name",
      "bank_account_type",
      "bank_account_number",
      "bank_account_kana",
    ].some(function (baseKey) {
      return !isBlank(firstValueByBaseKey(extraction, baseKey));
    });
  }

  function hasForeignCurrencyAmount(extraction) {
    return Object.keys(extraction).some(function (key) {
      return /_(usd|other)$/.test(key) && !isBlank(extraction[key]);
    });
  }

  function normalizeBankAccountType(value) {
    const source = String(value || "").trim();
    if (!source) {
      return null;
    }

    const compact = normalizeHeader(source);
    if (/^4$|外貨普通|外貨普通預金/.test(compact)) {
      return "4";
    }
    if (/^5$|外貨定期|外貨定期預金/.test(compact)) {
      return "5";
    }
    if (/^1$|普通|普通預金/.test(compact)) {
      return "1";
    }
    if (/^2$|定期|定期預金/.test(compact)) {
      return "2";
    }
    if (/^3$|当座|当座預金/.test(compact)) {
      return "3";
    }
    if (/^6$|別段|別段預金/.test(compact)) {
      return "6";
    }
    if (/^7$|その他|他|貯蓄|貯蓄預金/.test(compact)) {
      return "7";
    }

    return null;
  }

  function normalizeBankAccountNumber(value) {
    if (isBlank(value)) {
      return null;
    }

    const source = toHalfWidthAscii(value).replace(/\s+/g, " ").trim();
    const labeledNumber = cleanBankAccountNumberCandidate(source);
    if (!isBlank(labeledNumber)) {
      return labeledNumber;
    }

    const digits = source.replace(/[^\d]/g, "");
    if (digits.length >= 4 && digits.length <= 12) {
      return digits;
    }

    const uncertain = source
      .replace(/[？＊*Xx]/g, "?")
      .replace(/[\s\-－]/g, "");
    const digitCount = (uncertain.match(/\d/g) || []).length;
    if (digitCount >= 3 && /^[0-9?]{4,12}$/.test(uncertain)) {
      return uncertain;
    }

    const compact = source.replace(/\s+/g, "").toUpperCase();
    if (/^IBAN[:：]?[A-Z]{2}\d{2}[A-Z0-9]{8,}$/.test(compact) || /^[A-Z]{2}\d{2}[A-Z0-9]{8,}$/.test(compact)) {
      return compact.replace(/^IBAN[:：]?/, "");
    }

    return null;
  }

  function normalizeBankAccountKana(value) {
    if (isBlank(value)) {
      return null;
    }

    const text = String(value)
      .replace(/[\r\n\t]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    const normalized = text
      .replace(/([ァ-ヶーｦ-ﾟｰ])\s*[AＡ](?![A-Za-zＡ-Ｚａ-ｚ])/g, function (match, previousKana) {
        return `${previousKana}${bankAccountLongAForContext(previousKana)}`;
      })
      .replace(/([ァ-ヶーｦ-ﾟｰ])\s*[IＩlｌ1１][\-－ーｰ](?![A-Za-zＡ-Ｚａ-ｚ])/g, function (match, previousKana) {
        return `${previousKana}${bankAccountLongAForContext(previousKana)}`;
      });

    return normalized || null;
  }

  function bankAccountLongAForContext(previousKana) {
    return /[ｦ-ﾟｰ]/.test(previousKana) ? "ｴｰ" : "エー";
  }

  function fillSummaryFields(extraction) {
    const fallback = firstNonBlankValue([
      firstValueByBaseKey(extraction, "payment_description"),
      firstValueByBaseKey(extraction, "notes"),
    ]);

    if (isBlank(fallback)) {
      return;
    }

    Object.keys(extraction).forEach(function (key) {
      if (keyMatches(key, "summary") && isBlank(extraction[key])) {
        extraction[key] = fallback;
      }
    });
  }

  function fillExplicitPaymentContentFields(extraction, item) {
    const explicitValue = extractExplicitPaymentContent(item);
    if (isBlank(explicitValue) || isLikelyCounterpartyOnly(explicitValue, extraction)) {
      return;
    }

    Object.keys(extraction).forEach(function (key) {
      if ((keyMatches(key, "payment_description") || keyMatches(key, "summary")) && isBlank(extraction[key])) {
        extraction[key] = explicitValue;
      }
    });
  }

  function extractExplicitPaymentContent(item) {
    if (!item || isBlank(item.textSample)) {
      return null;
    }

    const lines = String(item.textSample)
      .split(/\n+/)
      .map(function (line) {
        return line.replace(/\s+/g, " ").trim();
      })
      .filter(Boolean);

    for (let index = 0; index < lines.length; index += 1) {
      const fromLine = explicitPaymentContentFromLine(lines[index]);
      const candidate = normalizeExplicitPaymentContentCandidate(fromLine || nextLinePaymentContent(lines, index));
      if (!isBlank(candidate)) {
        return candidate;
      }
    }

    return null;
  }

  function explicitPaymentContentFromLine(line) {
    const label = "(?:支払内容|摘要|件名|但し書き|但書|請求内容|ご請求内容|利用内容|ご利用内容|品名|内容)";
    const pattern = new RegExp(`${label}\\s*(?:[：:]|\\s{2,}|　+)\\s*(.+)$`);
    const match = String(line).match(pattern);
    if (!match) {
      return null;
    }

    return match[1];
  }

  function nextLinePaymentContent(lines, index) {
    const line = lines[index] || "";
    if (!/(支払内容|摘要|件名|但し書き|但書|請求内容|ご請求内容|利用内容|ご利用内容|品名|内容)\s*$/.test(line)) {
      return null;
    }

    return lines[index + 1] || null;
  }

  function normalizeExplicitPaymentContentCandidate(value) {
    if (isBlank(value)) {
      return null;
    }

    const text = String(value)
      .replace(/^(支払内容|摘要|件名|但し書き|但書|請求内容|ご請求内容|利用内容|ご利用内容|品名|内容)\s*[：:]?\s*/, "")
      .replace(/\s+/g, " ")
      .trim();

    if (!text || isLikelyNonContentValue(text)) {
      return null;
    }

    return text;
  }

  function isLikelyNonContentValue(text) {
    return /^[-－ー―]+$/.test(text) ||
      /^[¥￥]?\s*-?\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*円?$/.test(text) ||
      /^\d{4}[年\/.-]\d{1,2}[月\/.-]\d{1,2}日?$/.test(text) ||
      /^(金額|税額|消費税|合計|小計|単価|数量|備考|請求額|ご請求額|ご請求金額|ご請求分|ご利用明細|請求明細|内訳)$/i.test(text);
  }

  function firstNonBlankValue(values) {
    return values.find(function (value) {
      return !isBlank(value);
    }) || null;
  }

  function cleanPaymentContentFields(extraction, item) {
    Object.keys(extraction).forEach(function (key) {
      if (keyMatches(key, "payment_description") || keyMatches(key, "summary")) {
        extraction[key] = normalizePaymentContent(extraction[key], extraction, item);
      }
    });
  }

  function normalizePaymentContent(value, extraction, item) {
    if (isBlank(value)) {
      return null;
    }

    const text = String(value).replace(/\s+/g, " ").trim();
    if (isLikelyCounterpartyOnly(text, extraction)) {
      return null;
    }
    if (isLikelyGeneratedPaymentSummary(text, item)) {
      return null;
    }
    if (isLikelyBillingBreakdownLine(text, item)) {
      return null;
    }

    return text;
  }

  function isLikelyGeneratedPaymentSummary(text, item) {
    if (!item || isBlank(item.textSample)) {
      return false;
    }

    const source = normalizeHeader(item.textSample);
    const value = normalizeHeader(text);
    if (!value || source.includes(value)) {
      return false;
    }

    return /^(IP電話|電話|通信|インターネット|ネット|クラウド|システム|ソフト|SaaS|サービス).*(利用料|使用料|料金|代金)$|^(サービス利用料|利用料金|ご利用料金|月額利用料|月額料金)$/i.test(text);
  }

  function isLikelyBillingBreakdownLine(text, item) {
    if (!item || isBlank(item.textSample)) {
      return false;
    }

    const value = normalizeHeader(text);
    if (!value || hasExplicitContentLabelForValue(item.textSample, text)) {
      return false;
    }

    const section = extractBillingBreakdownSection(item.textSample);
    if (!section || !normalizeHeader(section).includes(value)) {
      return false;
    }

    return countAmountLikeValues(section) >= 2;
  }

  function hasExplicitContentLabelForValue(source, value) {
    const escaped = escapeRegExp(String(value).replace(/\s+/g, " ").trim());
    const pattern = new RegExp(`(?:件名|摘要|但し書き|但書|支払内容|請求内容|ご請求内容|利用内容|ご利用内容|品名|内容)\\s*[：:\\s]\\s*${escaped}`);
    return pattern.test(String(source).replace(/\s+/g, " "));
  }

  function extractBillingBreakdownSection(source) {
    const text = String(source);
    const startMatch = text.match(/ご請求分|ご請求内訳|ご利用明細|請求明細|明細内訳|内訳明細/);
    if (!startMatch) {
      return "";
    }

    const start = startMatch.index || 0;
    const tail = text.slice(start, start + 5000);
    const nextMatch = tail.slice(startMatch[0].length).match(/お支払い|お支払|ご請求金額|合計|総合計|振替日|引落日|領収|消費税|備考|お問い合わせ|発行日|請求書番号/);
    if (!nextMatch) {
      return tail;
    }

    return tail.slice(0, startMatch[0].length + (nextMatch.index || 0));
  }

  function countAmountLikeValues(source) {
    const matches = String(source).match(/[¥￥]?\s*-?\d{1,3}(?:,\d{3})+(?:\.\d+)?|-?\d+\s*円/g);
    return matches ? matches.length : 0;
  }

  function isLikelyCounterpartyOnly(text, extraction) {
    const vendor = firstValueByBaseKey(extraction, "vendor");
    if (!isBlank(vendor) && sameEntityText(text, vendor)) {
      return true;
    }

    if (hasPaymentContentCue(text)) {
      return false;
    }

    return hasLegalEntityCue(text);
  }

  function sameEntityText(left, right) {
    const leftComparable = comparableEntityText(left);
    const rightComparable = comparableEntityText(right);
    return Boolean(leftComparable && rightComparable && leftComparable === rightComparable);
  }

  function comparableEntityText(value) {
    return normalizeHeader(value)
      .toLowerCase()
      .replace(/株式会社|有限会社|合同会社|合名会社|合資会社|一般社団法人|公益社団法人|一般財団法人|公益財団法人|学校法人|医療法人|社会福祉法人|宗教法人|特定非営利活動法人|npo法人|㈱|株|incorporated|corporation|company|limited|inc|ltd|llc|co/g, "")
      .replace(/[-ーｰ－.,，、。]/g, "")
      .trim();
  }

  function hasLegalEntityCue(value) {
    return /株式会社|有限会社|合同会社|合名会社|合資会社|一般社団法人|公益社団法人|一般財団法人|公益財団法人|学校法人|医療法人|社会福祉法人|宗教法人|特定非営利活動法人|NPO法人|㈱|（株）|\(株\)|Inc\.?|Co\.?,?\s*Ltd\.?|Ltd\.?|LLC|Corporation|Company/i.test(String(value));
  }

  function hasPaymentContentCue(value) {
    return /利用料|使用料|サービス|システム|ソフト|ライセンス|サブスク|保守|管理費|手数料|料金|代金|商品|品代|購入|発注|請求内容|但し書き|摘要|明細|内訳|件名|月分|年会費|会費|広告|送料|運賃|交通費|宿泊|家賃|賃料|備品|消耗品|印刷|制作|製作|開発|委託|業務|コンサル|研修|修理|工事|レンタル|リース|決済|支払/i.test(String(value));
  }

  async function ensureApiProxyReady() {
    let response;
    try {
      response = await fetch(apiUrl("/api/health"), { cache: "no-store" });
    } catch (error) {
      throw new Error("APIに接続できません。Amplifyではconfig.jsのapiBaseUrl、ローカルではopen-local.cmdの起動を確認してください。");
    }

    if (!response.ok) {
      throw new Error("Gemini APIプロキシが見つかりません。AmplifyではAPI GatewayのURL設定とルート、ローカルではopen-local.cmdの起動を確認してください。");
    }
  }

  async function loadApiDefaults() {
    try {
      const response = await fetch(apiUrl("/api/health"), { cache: "no-store" });
      if (!response.ok) {
        return;
      }

      const payload = await response.json();
      if (payload.defaultModel) {
        els.modelInput.value = payload.defaultModel;
      }
    } catch (error) {
      // Keep file-only checks usable when the page is opened without the local server.
    }
  }

  function dataUrlToGeminiInlineData(dataUrl) {
    const match = String(dataUrl).match(/^data:([^;]+);base64,(.+)$/);
    if (!match) {
      throw new Error("画像データをGemini API用に変換できませんでした。");
    }

    return {
      inlineData: {
        mimeType: match[1],
        data: match[2],
      },
    };
  }

  async function loadDefaultExcelTemplate() {
    if (!window.XLSX) {
      setTemplateStatus("Excel読込ライブラリを読み込めませんでした。", "error");
      return;
    }

    try {
      const response = await fetch(DEFAULT_TEMPLATE_URL, { cache: "no-store" });
      if (!response.ok) {
        throw new Error(`既定テンプレートを取得できませんでした (${response.status})`);
      }

      const workbook = await readTemplateWorkbook(response);
      const templateSource = response.headers.get("X-Template-Source");
      const templateName = response.headers.get("X-Template-Name");
      const source = templateSource === "local-workbook"
        ? `ローカルExcel (${decodeURIComponent(templateName || "")})`
        : templateSource === "sharepoint"
          ? "SharePointテンプレート"
          : "ローカル既定テンプレート";
      applyTemplateWorkbook(workbook, source);
    } catch (error) {
      setTemplateStatus(error.message || "既定テンプレートの読込に失敗しました。", "error");
      sendClientLog("template.failed", {
        error: error.message || "既定テンプレートの読込に失敗しました。",
      });
    }
  }

  async function readTemplateWorkbook(response, options) {
    const contentType = String(response.headers.get("Content-Type") || "").toLowerCase();
    if (contentType.includes("text/csv")) {
      const text = await response.text();
      return XLSX.read(stripUtf8Bom(text), {
        type: "string",
        cellDates: Boolean(options && options.cellDates),
      });
    }

    const data = await response.arrayBuffer();
    return XLSX.read(data, options || {});
  }

  function stripUtf8Bom(value) {
    return String(value || "").replace(/^\uFEFF/, "");
  }

  function applyTemplateWorkbook(workbook, label) {
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      blankrows: false,
      defval: "",
    });
    const headerRow = (rows[0] || []).map(function (value) {
      return String(value).replace(/\s+/g, " ").trim();
    });
    const headerEntries = headerRow
      .map(function (label, index) {
        return { label, columnIndex: index };
      })
      .filter(function (entry) {
        return Boolean(entry.label);
      });

    if (headerEntries.length === 0) {
      throw new Error("Excelの1行目に項目名が見つかりませんでした。");
    }

    activeExtractionFields = buildFieldsFromHeaders(headerEntries);
    activeExtractionSchema = buildExtractionSchema(activeExtractionFields);
    const preview = activeExtractionFields.slice(0, 6).map(fieldLabel).join(", ");
    const suffix = activeExtractionFields.length > 6 ? ", ..." : "";
    setTemplateStatus("Excelフォーマット読込済み", "ready");
    sendClientLog("template.loaded", {
      source: label,
      fieldCount: activeExtractionFields.length,
      preview: `${preview}${suffix}`,
    });
  }

  function buildFieldsFromHeaders(headers) {
    const usedKeys = new Set();
    return headers.map(function (entry, index) {
      const header = typeof entry === "string" ? entry : entry.label;
      const columnIndex = typeof entry === "string" ? index : entry.columnIndex;
      const key = uniqueFieldKey(fieldKeyFromHeader(header, columnIndex), usedKeys);
      return {
        key,
        label: header,
        columnIndex,
        source: "excel",
      };
    });
  }

  function fieldKeyFromHeader(header, index) {
    const normalized = normalizeHeader(header);
    const knownMap = [
      [/^No$/i, "row_no"],
      [/^受領方法/, "receipt_method"],
      [/^会社コード/, "payer_company_code"],
      [/^支払先名$/, "vendor"],
      [/^取引先コード$/, "vendor_code"],
      [/^稟議No$/i, "approval_no"],
      [/^￥通貨支払金額税抜$/, "subtotal_amount_jpy"],
      [/^\$通貨支払金額税抜$/, "subtotal_amount_usd"],
      [/^その他通貨支払金額税抜$/, "subtotal_amount_other"],
      [/^支払内容$/, "payment_description"],
      [/^支払方法/, "payment_method"],
      [/^支払(?:日|期日|期限)/, "payment_due_date"],
      [/^お支払い?(?:期限|期日)/, "payment_due_date"],
      [/^振込期限|^入金期日|^引落|^引き落とし|^振替日|^口座振替日/, "payment_due_date"],
      [/^支払カード名$/, "payment_card_name"],
      [/^クレジットカード登録申請No$/i, "credit_card_application_no"],
      [/^利用月$/, "usage_month"],
      [/^契約有無/, "contract_exists"],
      [/^自動更新/, "auto_renewal"],
      [/^契約期間開始日$/, "contract_start_date"],
      [/^契約期間終了日$/, "contract_end_date"],
      [/^経理CD$/i, "accounting_code"],
      [/^摘要$/, "summary"],
      [/^￥通貨合計支払金額$/, "total_amount_jpy"],
      [/^\$通貨合計支払金額$/, "total_amount_usd"],
      [/^その他通貨合計支払金額$/, "total_amount_other"],
      [/^銀行名$/, "bank_name"],
      [/^支店名$/, "bank_branch_name"],
      [/^(口座種別|口座種類|預金種目|預金種別|預金科目|普通|当座|貯蓄)/, "bank_account_type"],
      [/^(口座番号|口座No|AccountNumber|AccountNo|ACNo|IBAN)/i, "bank_account_number"],
      [/^(口座名義(?:フリガナ|フリ仮名|カナ|ﾌﾘｶﾞﾅ|ﾌﾘ仮名|ｶﾅ)?|名義(?:フリガナ|フリ仮名|カナ|ﾌﾘｶﾞﾅ|ﾌﾘ仮名|ｶﾅ)|受取人名(?:義)?(?:フリガナ|カナ|ﾌﾘｶﾞﾅ|ｶﾅ)|振込先名義(?:フリガナ|カナ|ﾌﾘｶﾞﾅ|ｶﾅ)?)/, "bank_account_kana"],
      [/^請求書発行日$/, "invoice_issue_date"],
      [/^伝票種類$/, "voucher_type"],
      [/^伝票日付$/, "voucher_date"],
      [/^仕入先CD社員CD$/i, "supplier_employee_code"],
      [/^支払先区分$/, "payee_type"],
      [/^サブ口座$/, "sub_account"],
      [/^債務科目$/, "liability_account"],
      [/^補助$/, "sub_accounting_code"],
      [/^明細科目$/, "detail_account"],
      [/^税区$/, "tax_category"],
      [/^相手先コード$/, "counterparty_code"],
      [/^イニシャルストック$/, "initial_stock"],
      [/^請求書コピー1$/, "invoice_copy_1"],
      [/^請求書コピー2$/, "invoice_copy_2"],
      [/^請求書原本1$/, "invoice_original_1"],
      [/^請求書原本2$/, "invoice_original_2"],
      [/^明細内訳1$/, "detail_attachment_1"],
      [/^明細内訳2$/, "detail_attachment_2"],
      [/^契約書コピー1$/, "contract_copy_1"],
      [/^契約書コピー2$/, "contract_copy_2"],
      [/^その他1$/, "other_attachment_1"],
      [/^その他2$/, "other_attachment_2"],
      [/^その他の明細$/, "other_detail"],
      [/^備考$/, "notes"],
      [/^(日付|発行日|請求日|領収日|取引日|利用日)$/, "date"],
      [/^(支払先|支払先名|取引先|会社名|店舗名|店名|仕入先|請求元)$/, "vendor"],
      [/^(合計金額|合計|総額|金額|請求金額|税込金額|領収金額)$/, "total_amount"],
      [/^(小計|税抜金額|本体金額)$/, "subtotal_amount"],
      [/^(税額|消費税|消費税額|内税|外税)$/, "tax_amount"],
      [/^(請求書(?:No|NO|番号)|請求(?:No|NO|番号)|領収書(?:No|NO|番号)|伝票番号|管理番号|発行番号|番号)$/i, "invoice_number"],
      [/^(支払日|支払期日|支払期限|振込期限|期限|入金期日|引落日|振替日|口座振替日)$/, "payment_due_date"],
      [/^(通貨|通貨コード|currency)$/i, "currency"],
    ];

    const matched = knownMap.find(function (entry) {
      return entry[0].test(normalized);
    });

    return matched ? matched[1] : `field_${index + 1}`;
  }

  function normalizeHeader(header) {
    return String(header)
      .replace(/[０-９]/g, function (value) {
        return String.fromCharCode(value.charCodeAt(0) - 0xfee0);
      })
      .replace(/[Ａ-Ｚａ-ｚ]/g, function (value) {
        return String.fromCharCode(value.charCodeAt(0) - 0xfee0);
      })
      .replace(/＄/g, "$")
      .replace(/[ 　\r\n\t]/g, "")
      .replace(/[()（）［］\[\]【】]/g, "")
      .replace(/[：:・･\/／]/g, "")
      .trim();
  }

  function uniqueFieldKey(baseKey, usedKeys) {
    let key = baseKey;
    let count = 2;
    while (usedKeys.has(key)) {
      key = `${baseKey}_${count}`;
      count += 1;
    }
    usedKeys.add(key);
    return key;
  }

  function buildExtractionSchema(fields) {
    const properties = {};
    const required = [];

    fields.forEach(function (field) {
      const key = fieldKey(field);
      properties[key] = schemaForField(key, fieldLabel(field));
      required.push(key);
    });

    return {
      type: "OBJECT",
      properties,
      required,
      propertyOrdering: required,
    };
  }

  function schemaForField(key, label) {
    const schema = {
      type: "STRING",
      nullable: true,
      description: fieldDescription(key, label),
    };

    if (isAmountField(key, label)) {
      schema.type = "NUMBER";
    }

    return schema;
  }

  function fieldDescription(key, label) {
    const rules = [
      `Excel列「${label}」。領収書・請求書上の項目名がExcel列名と一致しなくても、意味が同じなら抽出する。`,
      `この列の意味: ${fieldMeaning(key, label)}`,
      `領収書・請求書での表記ゆれ候補: ${fieldAliases(key, label).join("、") || "なし"}`,
      "ただし、意味が近いだけで別項目の場合は入れない。",
      "読み取れない、判断できない、または書類に存在しない場合は必ずnull。",
      "似ている別項目の値で代用しない。",
    ];

    if (keyMatches(key, "vendor")) {
      rules.push(...vendorRules());
    }

    if (key === "row_no") {
      rules.push(...rowNoRules());
    }

    if (keyMatches(key, "receipt_method")) {
      rules.push(...receiptMethodRules());
    }

    if (keyMatches(key, "payment_method")) {
      rules.push(...paymentMethodRules());
    }

    if (keyMatches(key, "bank_account_type")) {
      rules.push(...bankAccountTypeRules());
    }

    if (keyMatches(key, "bank_account_kana")) {
      rules.push(...bankAccountKanaRules());
    }

    if (keyMatches(key, "payment_description")) {
      rules.push(...paymentDescriptionRules());
    }

    if (keyMatches(key, "summary")) {
      rules.push(...summaryRules());
    }

    if (isAmountField(key, label)) {
      rules.push(...amountRules(key, label));
      if (label.includes("￥") || label.includes("円")) {
        rules.push("JPY/円の金額列。外貨金額しかない場合はnull。");
      }
      if (label.includes("$") || label.includes("＄")) {
        rules.push("USD/ドルの金額列。円の金額しかない場合はnull。");
      }
      if (label.includes("その他通貨")) {
        rules.push("円・ドル以外の通貨金額列。円またはドルしかない場合はnull。");
      }
    }

    if (isDateField(key, label)) {
      rules.push("日付は可能ならYYYY-MM-DD形式。請求書発行日、支払日、伝票日付、契約期間を混同しない。");
    }

    if (isOptionField(label)) {
      rules.push("列名に選択肢番号がある場合、書類上の根拠が明確なときだけ対応する番号を返す。推測ならnull。");
    }

    if (isInternalField(key, label)) {
      rules.push("これは社内申請・会計システム用項目。書類に同じ項目名と値が明記されていない限りnull。");
      rules.push("請求書番号、会社名、金額など別項目の値を入れない。");
    }

    if (isAttachmentField(key, label)) {
      rules.push("これは添付ファイル管理欄。書類画像から値を推測せず、明記がない限りnull。");
    }

    return rules.join(" ");
  }

  function amountRules(key, label) {
    const rules = [
      "金額は通貨記号、円記号、カンマを除いた数値だけで返す。",
      "複数の金額候補がある場合は、まず書類上のラベルを確認し、列の意味に一致する金額だけを入れる。",
      "数量、単価、税率、電話番号、口座番号、伝票番号、日付を金額として扱わない。",
    ];

    if (key.startsWith("subtotal_amount") || /税抜|小計|本体/.test(label)) {
      rules.push("この列は税抜金額。'税抜'、'小計'、'本体価格'、'課税対象額'、'支払金額(税抜)' に対応する金額を入れる。");
      rules.push("'税込'、'合計'、'総合計'、'お支払金額'、'ご請求額' しか読めない場合は、この税抜列はnull。");
      rules.push("消費税額そのものは入れない。");
    }

    if (key.startsWith("total_amount") || /合計|総額|請求金額|領収金額|税込/.test(label)) {
      rules.push("この列は最終支払額。'合計'、'税込合計'、'総合計'、'お支払金額'、'ご請求額'、'領収金額' に対応する金額を入れる。");
      rules.push("口座振替請求書兼領収書では、'口座振替額'、'振替金額'、'引落金額'、'請求金額'、'領収金額' に対応する金額を入れる。");
      rules.push("税抜金額や小計だけが読める場合は、この合計列に流用しない。");
      rules.push("税額を足して税込合計を計算しない。書類に合計額が明記されている場合だけ入れる。");
    }

    if (key.startsWith("tax_amount") || /税額|消費税|内税|外税/.test(label)) {
      rules.push("この列は消費税額。'消費税'、'税額'、'内消費税'、'外税' に対応する金額だけを入れる。");
      rules.push("税込合計や税抜金額を入れない。");
    }

    if (key.endsWith("_jpy") || label.includes("￥") || label.includes("円")) {
      rules.push("円/JPY列。円記号、'円'、'JPY' の金額を優先する。ドルやその他通貨の金額はnull。");
    }

    if (key.endsWith("_usd") || label.includes("$") || label.includes("＄")) {
      rules.push("ドル/USD列。'$'、'USD'、'ドル' の金額だけを入れる。円の金額はnull。");
    }

    if (key.endsWith("_other") || label.includes("その他通貨")) {
      rules.push("その他通貨列。JPY/USD以外の通貨が明記されている場合だけ入れる。円やドルしかない場合はnull。");
    }

    return rules;
  }

  function vendorRules() {
    return [
      "支払先名は、実際に支払いを受け取る相手。最優先で、振込先、受取人名、受取人名義、口座名義、支払先として明記された名称を採用する。",
      "請求者名、請求者、宛先、請求先、納品先、お客様名、利用者名、買い手、支払者、支払い元会社は支払先名ではないため採用しない。",
      "振込先がない請求書・領収書では、御中・様・殿が付いていない請求元、発行者、発行元、領収者、店舗名を支払先候補にする。",
      "御中、様、殿が付いている会社名は宛先・請求先側なので支払先名にしない。敬称だけを削って支払先名として返すことも禁止。",
      "同じ書類に 'A社 御中' と 'B社' がある場合、A社は宛先側、B社が支払先候補。必ず御中などが付いていない会社名を優先する。",
      "候補が御中・様・殿付きの会社名しかない場合、支払先名はnull。",
      "銀行振込先しか明記されていない場合は、口座名義または受取人名を支払先名として扱ってよい。ただし銀行名・支店名・口座番号は入れない。",
      "口座振替請求書兼領収書、口座振替請求書、口座振替領収書では、振込先欄がなくても抽出を止めない。御中などが付いていない発行元・請求元・領収者を支払先候補にする。",
      "同じ会社名が複数表記されている場合は、御中などが付いていない正式表記を優先する。",
    ];
  }

  function rowNoRules() {
    return [
      "これはExcel取込用の行管理No。請求書番号、領収書番号、伝票番号、書類上のNo.とは別項目。",
      "AI抽出では必ずnull。画面で手入力された場合だけ出力する。",
    ];
  }

  function receiptMethodRules() {
    return [
      "受領方法は支払方法ではない。銀行振込、口座引落、クレジットなどを入れない。",
      "同じファイル種別なら必ず同じ番号にする。AIの見た目判断ではなく、事前判定を優先する。",
      "選択肢番号だけを返す。デジタルPDFなら2、画像または画像PDFなら1。",
      "紙契約書と明記がある場合だけ3、電子契約書と明記がある場合だけ4。",
    ];
  }

  function paymentMethodRules() {
    return [
      "支払方法は選択肢番号だけを返す。1=銀行振込、2=納付書、3=口座引落、4=海外送金、5=銀行振込(スポット)、6=クレジット自動引落、7=クレジット即日決済。",
      "銀行名、支店名、口座番号、口座名義、振込先が明記されている場合は原則1。海外送金、SWIFT、IBAN、外貨送金が明記されている場合は4。",
      "納付書、払込票、振込用紙が明記されている場合は2。口座振替、口座引落、自動引落が明記されている場合は3。",
      "書類タイトルが口座振替請求書兼領収書、口座振替請求書、口座振替領収書の場合は3。",
      "支払カード名やクレジットカード登録申請Noがあり、自動/継続のカード支払なら6。即日決済、カード決済、クレジットカード決済なら7。",
      "受領方法や領収書の形式と混同しない。",
      "不鮮明でも、支払方法を示すアンカー語（振込先、口座番号、口座名義、普通、当座、貯蓄、口座振替、納付書、クレジット）が見える場合は候補を出す。完全に読めない場合だけnull。",
    ];
  }

  function bankAccountTypeRules() {
    return [
      "口座種別は選択肢番号だけを返す。1=普通、2=定期、3=当座、4=外貨普通、5=外貨定期、6=別段、7=その他。",
      "アンカー語として普通、当座、貯蓄を優先して確認する。普通預金は1、当座預金は3、貯蓄預金は選択肢にないため7。漢字や語句ではなく番号のみを返す。",
      "判別できない場合はnull。銀行名、支店名、口座番号を口座種別に入れない。",
      "文字が一部不鮮明でも、普通/当座/貯蓄のどれかが読める場合は候補番号を返す。",
    ];
  }

  function bankAccountKanaRules() {
    return [
      "口座名義フリガナは銀行振込先に記載されたカナ名義をそのまま抽出する。",
      "半角カナ、全角カナ、英数字、スペース、ハイフン、括弧、記号は書類の表記どおり保持する。不要な変換、補完、整形をしない。",
      "全角カタカナと半角カタカナをアンカーとして見る。口座名義フリガナ、口座名義カナ、名義カナ、受取人名カナ、振込先名義カナ、ﾌﾘｶﾞﾅ、ｶﾅ の周辺値を優先する。",
      "株式会社などの法人略語も省略しない。例: 'ｶ) ﾙ-ﾄｴ-' と記載されていれば、先頭の 'ｶ)' を含めて 'ｶ) ﾙ-ﾄｴ-' と返す。",
      "カナのエーを英字A、I、I-に置き換えない。例: 'カ)ルートエー' と読める場合は 'カ)ルートA' や 'カ)ルートI-' ではなく 'カ)ルートエー' と返す。",
      "支払先名や会社名からカナを推測しない。口座名義フリガナが完全に読めない場合だけnull。一部不鮮明な場合は読める範囲の候補を返す。",
    ];
  }

  function paymentDescriptionRules() {
    return [
      "支払内容は品名、件名、但し書き、利用内容、請求内容、明細の主な内容を入れる。",
      "原則として書類上の文言をそのまま返す。抽象化、要約、言い換え、一般名への丸めは禁止。",
      "件名、摘要、但し書き、請求内容、ご利用内容として明記された文言を優先する。",
      "支払内容、摘要、件名、但し書き、請求内容、ご利用内容、品名、内容として書類に値が記入されている場合は必ず抽出する。",
      "ご請求分、ご請求内訳、ご利用明細、請求明細、内訳明細の表の中に複数行がある場合、その中の1行だけを支払内容として採用しない。ただし支払内容・摘要・件名などの明記欄に書かれた値は採用する。",
      "複数明細の請求で、支払内容・摘要・件名・但し書き・請求内容など全体を表す記入欄が読めない場合だけnull。",
      "'IP電話サービス利用料'、'サービス利用料'、'利用料金' のような分類名をAIが作って返さない。書類にその語句がそのまま印字され、かつそれ以上具体的な明細名がない場合だけ採用する。",
      "会社名、店舗名、支払先名、請求元名だけを支払内容に入れない。会社名しか読めない場合はnull。",
      "支払先名と同じ値、または株式会社/有限会社/合同会社などの法人名だけの値はnull。",
      "金額、日付、請求書番号、口座番号だけを支払内容に入れない。",
    ];
  }

  function summaryRules() {
    return [
      "摘要は会計取込用の支払内容メモ。品名、件名、但し書き、利用内容、請求内容、明細の主な内容を入れる。",
      "支払内容と同じく、書類上の文言を原文優先で返す。抽象化、要約、言い換え、一般名への丸めは禁止。",
      "支払内容、摘要、件名、但し書き、請求内容、ご利用内容、品名、内容として書類に値が記入されている場合は必ず抽出する。",
      "ご請求分、ご請求内訳、ご利用明細、請求明細、内訳明細の表の中に複数行がある場合、その中の1行だけを摘要として採用しない。ただし支払内容・摘要・件名などの明記欄に書かれた値は採用する。",
      "複数明細の請求で、支払内容・摘要・件名・但し書き・請求内容など全体を表す記入欄が読めない場合だけnull。",
      "'IP電話サービス利用料'、'サービス利用料'、'利用料金' のような分類名をAIが作って返さない。書類にその語句がそのまま印字され、かつそれ以上具体的な明細名がない場合だけ採用する。",
      "Excelに支払内容列もある場合、摘要は支払内容と同じ値でよい。支払内容が読めるのに摘要だけnullにしない。",
      "会社名、店舗名、支払先名、請求元名だけを摘要に入れない。会社名しか読めない場合はnull。",
      "金額、日付、請求書番号、口座番号だけを摘要に入れない。内容が本当に読めない場合だけnull。",
    ];
  }

  function keyMatches(key, baseKey) {
    if (key === baseKey || new RegExp(`^${escapeRegExp(baseKey)}_\\d+$`).test(key)) {
      return true;
    }

    const field = activeExtractionFields.find(function (candidate) {
      return fieldKey(candidate) === key;
    });
    if (!field) {
      return false;
    }

    const semanticKey = fieldKeyFromHeader(fieldLabel(field), fieldColumnIndex(field));
    return semanticKey === baseKey || new RegExp(`^${escapeRegExp(baseKey)}_\\d+$`).test(semanticKey);
  }

  function escapeRegExp(value) {
    return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  function fieldMeaning(key, label) {
    if (key === "row_no") {
      return "Excel取込用の行管理No。書類上の請求書番号、領収書番号、伝票番号、No.とは別。AIでは抽出しない。";
    }
    if (keyMatches(key, "vendor")) {
      return "支払先名。実際に支払いを受け取る相手。振込先、受取人名、受取人名義、口座名義、支払先を優先する。請求者名・請求者・宛先・請求先は除外する。振込先がない請求書・領収書では、御中・様・殿が付いていない請求元・発行者・領収者を支払先候補にする。御中・様・殿が付いた宛先側の会社名は除外する。";
    }
    if (key.includes("vendor_code") || key.includes("counterparty_code") || key.includes("supplier_employee_code")) {
      return "社内で管理する取引先コード・仕入先コード。書類に明記がなければnull。";
    }
    if (key.startsWith("subtotal_amount")) {
      return "税抜の支払金額・小計・本体価格。税込合計とは区別する。";
    }
    if (key.startsWith("total_amount")) {
      return "税込または最終的に支払う合計金額・請求金額・領収金額。税抜金額とは区別する。";
    }
    if (key.startsWith("tax_amount")) {
      return "消費税額・税額。合計金額や税抜金額とは区別する。";
    }
    if (keyMatches(key, "payment_description") || keyMatches(key, "summary") || keyMatches(key, "notes")) {
      if (keyMatches(key, "summary")) {
        return "摘要。会計取込用の短い支払内容メモ。支払内容列と同じ値でよいが、会社名だけは入れない。";
      }
      return "支払内容・備考。品名、但し書き、利用内容、請求内容が該当する。会社名だけは入れない。";
    }
    if (keyMatches(key, "receipt_method")) {
      return "請求書・領収書の受領方法。1=紙、2=データ/電子、3=その他(紙契約書)、4=その他(電子契約書)。支払方法とは別。";
    }
    if (keyMatches(key, "payment_method")) {
      return "支払方法。1=銀行振込、2=納付書、3=口座引落、4=海外送金、5=銀行振込(スポット)、6=クレジット自動引落、7=クレジット即日決済。";
    }
    if (keyMatches(key, "payment_due_date")) {
      return "支払日・支払期日・振込期限。請求書発行日とは区別する。";
    }
    if (keyMatches(key, "invoice_issue_date") || keyMatches(key, "date")) {
      return "請求書・領収書の発行日、請求日、領収日、取引日。支払期日とは区別する。";
    }
    if (keyMatches(key, "invoice_number")) {
      return "請求書番号・請求書No・請求No・領収書番号・伝票番号・No.。書類右上やタイトル付近に単に 'No.' とだけ表示される場合も請求書番号として扱う。登録番号(T+13桁)、電話番号、郵便番号、日付、金額とは区別する。";
    }
    if (keyMatches(key, "bank_account_type")) {
      return "銀行振込先の口座種別。1=普通、2=定期、3=当座、4=外貨普通、5=外貨定期、6=別段、7=その他。";
    }
    if (keyMatches(key, "bank_account_kana")) {
      return "銀行振込先の口座名義フリガナ。半角カナ、法人略語、記号、スペースを含めて書類どおり保持する。";
    }
    if (key.startsWith("bank_")) {
      return "銀行振込先の銀行名、支店名、口座種別、口座番号、口座名義。";
    }
    if (isInternalField(key, label)) {
      return "社内申請・会計システム用の管理項目。書類に明記がなければnull。";
    }
    if (isAttachmentField(key, label)) {
      return "添付ファイル管理欄。書類上の値ではないため、明記がなければnull。";
    }
    return "Excel列名の意味に対応する値。項目名の完全一致ではなく、意味で判断する。";
  }

  function fieldAliases(key, label) {
    const aliases = new Set([label]);
    const add = function (values) {
      values.forEach(function (value) {
        aliases.add(value);
      });
    };

    if (keyMatches(key, "vendor")) {
      add(["支払先", "支払先名", "お支払先", "振込先", "振込先名", "振込先名義", "受取人", "受取人名", "受取人名義", "口座名義", "領収者", "店舗名", "店名", "加盟店", "取引先", "仕入先", "Payee", "Remit to", "Remittance", "Beneficiary", "Beneficiary Name", "Account Name", "Account Holder", "Vendor", "Supplier", "Seller", "From"]);
    }
    if (key.startsWith("subtotal_amount")) {
      add(["税抜金額", "小計", "本体価格", "本体金額", "税抜合計", "支払金額(税抜)", "金額(税抜)", "Subtotal", "Sub Total", "Amount before tax", "Taxable amount"]);
    }
    if (key.startsWith("total_amount")) {
      add(["合計", "合計金額", "総合計", "税込金額", "請求金額", "領収金額", "お支払金額", "ご請求額", "支払額", "税込合計", "口座振替額", "振替金額", "引落金額", "引落額", "Total", "Grand Total", "Amount Due", "Balance Due", "Invoice Total", "Total Amount", "Amount Payable", "Total Due"]);
    }
    if (key.startsWith("tax_amount")) {
      add(["消費税", "消費税額", "税額", "内消費税", "内税", "外税", "10%対象税額", "Tax", "VAT", "Sales Tax", "Consumption Tax", "Tax Amount"]);
    }
    if (keyMatches(key, "payment_description") || keyMatches(key, "summary") || keyMatches(key, "notes")) {
      add(["支払内容", "摘要", "備考", "但し書き", "品名", "内容", "利用内容", "利用料金", "ご利用料金", "明細", "内訳", "件名", "請求内容", "Description", "Item", "Items", "Product", "Service", "Details", "Memo", "Notes", "Subject", "Line item"]);
    }
    if (keyMatches(key, "payment_method")) {
      add(["支払方法", "決済方法", "お支払方法", "支払い方法", "決済種別", "銀行振込", "振込先", "お振込先", "納付書", "払込票", "口座引落", "口座振替", "口座振替請求書兼領収書", "口座振替請求書", "口座振替領収書", "海外送金", "SWIFT", "IBAN", "クレジット", "クレジットカード", "カード決済", "Payment Method", "Bank Transfer", "Wire Transfer", "Credit Card", "ACH", "Direct Debit"]);
    }
    if (keyMatches(key, "receipt_method")) {
      add(["受領方法", "受取方法", "受領区分", "紙", "データ", "電子", "原本"]);
    }
    if (keyMatches(key, "payment_due_date")) {
      add(["支払日", "支払期日", "支払期限", "振込期限", "入金期日", "お支払期限", "お支払い期限", "引落日", "引落予定日", "振替日", "口座振替日", "口座振替予定日", "Due Date", "Payment Due", "Payment Due Date", "Payment Deadline"]);
    }
    if (keyMatches(key, "invoice_issue_date") || keyMatches(key, "date")) {
      add(["発行日", "請求日", "領収日", "取引日", "利用日", "日付", "年月日", "発行年月日", "Invoice Date", "Issue Date", "Date", "Billing Date", "Receipt Date", "Transaction Date"]);
    }
    if (keyMatches(key, "invoice_number")) {
      add(["請求書番号", "請求書No", "請求書NO", "請求書№", "請求書ID", "ご請求書No", "御請求書No", "請求No", "請求NO", "請求№", "請求ID", "請求番号", "ご請求番号", "御請求番号", "領収書番号", "領収書No", "領収書№", "伝票番号", "伝票No", "管理番号", "発行番号", "発行No", "No.", "No", "№", "番号", "Invoice No", "Invoice Number", "Invoice #", "Invoice ID", "Receipt No", "Document No"]);
    }
    if (keyMatches(key, "currency") || key.includes("_jpy") || key.includes("_usd") || key.includes("_other")) {
      add(["通貨", "通貨コード", "JPY", "円", "USD", "ドル", "$", "外貨"]);
    }
    if (key === "bank_name") {
      add(["銀行名", "金融機関名", "振込先銀行", "お振込先", "Bank", "Bank Name", "Financial Institution"]);
    }
    if (key === "bank_branch_name") {
      add(["支店名", "支店", "支店コード", "店番", "Branch", "Branch Name", "Branch Code"]);
    }
    if (keyMatches(key, "bank_account_type")) {
      add(["口座種別", "預金種目", "科目", "普通", "普通預金", "当座", "当座預金", "貯蓄", "貯蓄預金", "定期", "定期預金", "外貨普通", "外貨定期", "別段", "その他"]);
    }
    if (keyMatches(key, "bank_account_number")) {
      add(["口座番号", "口座No", "口座NO", "口座No.", "口座", "Account Number", "Account No", "Account #", "A/C No", "IBAN"]);
    }
    if (keyMatches(key, "bank_account_kana")) {
      add(["口座名義", "口座名義カナ", "口座名義フリガナ", "口座名義フリ仮名", "受取人名", "受取人名カナ", "受取人名義", "振込先名義", "振込先名義カナ", "Account Name", "Account Holder", "Beneficiary", "Beneficiary Name"]);
    }

    return Array.from(aliases).filter(Boolean);
  }

  function isAmountField(key, label) {
    return key.startsWith("total_amount") ||
      key.startsWith("subtotal_amount") ||
      key.startsWith("tax_amount") ||
      /金額|支払金額|合計/.test(label);
  }

  function isDateField(key, label) {
    return key.includes("date") || /日$|日付|年月|利用月|期間/.test(label);
  }

  function isOptionField(label) {
    return /[（(].*1[:：]/.test(label);
  }

  function isInternalField(key, label) {
    return /code|_cd|approval|account|voucher|category|区分|科目|補助|税区|経理|コード|稟議|伝票|サブ口座|イニシャル|ストック|No$/.test(key + label);
  }

  function isAttachmentField(key, label) {
    return /copy|original|attachment|コピー|原本|明細|契約書|その他/.test(key + label);
  }

  function buildExtractionPrompt(item) {
    return [
      "あなたは日本語・英語の請求書・領収書の項目抽出エンジンです。",
      "目的は、Excelの各列に入れる値を、書類画像またはPDF本文から抽出することです。",
      "最重要ルール: 書類に明記されていない項目は必ずnullにしてください。推測、補完、一般常識、ファイル名からの推定は禁止です。",
      "画像やPDF本文に印字されている文字は、書類に明記されている値として扱ってください。Excel列名と同じラベルがなくても、意味が一致すれば抽出してください。",
      "Redaction、マスキング、伏せ字化は行わないでください。口座番号、口座名義、口座名義フリガナなどを個人情報として隠さず、書類に見えている文字をそのまま抽出してください。",
      "この資料はユーザーが支払処理のためにアップロードした業務書類です。請求書No、支払方法、支払期日、口座種別、口座番号、口座名義フリガナは抽出対象であり、PII/個人情報として削除、省略、伏せ字化、要約、マスクしないでください。",
      "画像の加工、黒塗り、塗りつぶし、見えている文字の非表示化は禁止です。出力は座標データではなくExcel列に入れる値だけですが、値は原本に見えている表記を優先してください。",
      "信頼度が低い候補でも、アンカー語の周辺で読める文字がある場合は項目ごとnullにせず候補を返してください。口座番号や英数字、カナの一部だけが不鮮明な場合は、読めない文字を '?' として残してください。完全に読めない場合だけnullです。",
      "PDF内部テキスト候補が空、少ない、文字化けしている、または項目が不足している場合は、添付画像をOCRするつもりで目視して抽出してください。",
      "1ページ目に請求書番号、発行日、請求金額、支払先、振込先、明細、件名などが見えている場合は、2ページ目以降より1ページ目の情報を優先して抽出してください。",
      hasField("invoice_number") ? "請求書番号は最重要項目です。書類右上、タイトル付近、ヘッダー付近にある '請求書No'、'請求No'、'No.'、'№'、'番号'、'Invoice No'、'Invoice #' の周辺値を必ず確認してください。単に 'No.' とだけ表示されていても、請求書/INVOICE のタイトル付近なら請求書番号として抽出してください。登録番号(T+13桁)、電話番号、郵便番号、日付、金額は請求書番号にしないでください。" : "",
      "口座振替請求書兼領収書、口座振替請求書、口座振替領収書は振込先欄がないことがあります。その場合も抽出を止めず、支払方法は3、金額は口座振替額・振替金額・引落金額・請求金額・領収金額から、支払日は振替日・引落日から拾ってください。",
      "社内申請用のコード、稟議No、経理CD、科目、補助、税区、コピー/原本/添付管理欄は、書類上に同じ項目名と値が明記されていない限りnullです。",
      "列名が似ていても別項目へ値を流用しないでください。例: 請求書発行日と支払日は別、税抜金額と合計金額は別、支払先名と取引先コードは別です。",
      "Excel列名と書類上の項目名は一致しない前提です。列ごとの意味と表記ゆれ候補を見て、意味が一致する値だけを抽出してください。",
      ["invoice_number", "payment_method", "payment_due_date", "bank_account_type", "bank_account_number", "bank_account_kana"].some(hasField) ? "請求書No、支払方法、支払期日、口座種別、口座番号、口座名義フリガナは重要項目です。PDF本文や画像内の振込先欄、口座情報欄、支払条件欄に印字されている場合は、ページ位置にかかわらず必ず拾ってください。候補の信頼度が低くても、アンカー語の周辺に読める値がある場合は項目ごと削除せず候補を返してください。完全に見えない場合だけnullです。" : "",
      ["bank_account_type", "bank_account_number", "bank_account_kana"].some(hasField) ? "口座情報のアンカー語: 普通、当座、貯蓄、口座種別、預金種目、口座番号、口座名義、口座名義フリガナ、口座名義カナ、受取人名カナ、振込先名義カナ、ﾌﾘｶﾞﾅ、ｶﾅ。これらの前後左右にある値を抽出対象に含めてください。" : "",
      hasField("vendor") ? "支払先名の判定手順: 1) 会社名候補を全て見る。2) 御中・様・殿が付く候補、請求者名、請求者、宛先、請求先、支払者側の候補を除外する。3) 残った候補から振込先・受取人名・口座名義・支払先を優先し、振込先がなければ御中などが付いていない請求元・発行者・発行元・領収者を選ぶ。4) 御中付き候補の敬称だけを削って返すことは禁止。迷う場合はnull。" : "",
      hasField("bank_account_kana") ? "口座名義フリガナは表記どおり保持してください。例: 'ｶ) ﾙ-ﾄｴ-' は 'ｶ)' を省略せず、そのまま返してください。'カ)ルートエー' を 'カ)ルートA' や 'カ)ルートI-' にしないでください。" : "",
      hasField("payment_description") || hasField("summary") ? "支払内容・摘要は、書類に印字された支払内容、摘要、件名、但し書き、請求内容、ご利用内容、品名、内容を原文優先で抽出してください。これらの記入欄に値がある場合は必ず抽出してください。AIが内容を要約した分類名を作ることは禁止です。ご請求分・ご請求内訳・ご利用明細・請求明細の表に複数行がある場合、その中の1行だけを代表として返さないでください。ただし支払内容・摘要・件名などの明記欄に書かれた値は採用してください。" : "",
      "金額の重要ルール: 税抜、小計、消費税、税込合計、支払総額を混同しないでください。計算で補完せず、書類に明記された金額だけを返してください。",
      "金額候補が複数ある場合は、書類上のラベルとExcel列の意味が一致するものだけを採用してください。",
      "日付は判別できる場合 YYYY-MM-DD に正規化してください。",
      "金額は数値だけにし、円記号やカンマは含めないでください。",
      `抽出項目: ${activeExtractionFields.map(function (field) {
        return `${fieldKey(field)}=${fieldLabel(field)}`;
      }).join(", ")}`,
      "列ごとの注意:",
      ...activeExtractionFields.map(function (field) {
        return `- ${fieldKey(field)} (${fieldLabel(field)}): ${fieldDescription(fieldKey(field), fieldLabel(field))}`;
      }),
      `ファイル名: ${displayName(item.file)}`,
      `事前判定: ${item.routeLabel}`,
      Array.isArray(item.pagePreviewUrls) && item.pagePreviewUrls.length > 1
        ? `添付画像: PDF ${item.pagePreviewUrls.length}/${item.pageCount || item.pagePreviewUrls.length}ページ分。添付されたページを順番にすべて確認し、1ページ目以外に明細・合計・口座振替額・支払日・摘要がある場合も抽出する。`
        : "",
      item.pageCount && item.pagePreviewUrls && item.pagePreviewUrls.length < item.pageCount
        ? `注意: 画像添付は先頭${item.pagePreviewUrls.length}ページまで。ただしPDF内部テキスト候補は最大${PDF_TEXT_PAGE_LIMIT}ページ分を含むため、後続ページの文字情報も確認する。`
        : "",
      item.textSample ? `PDF内部テキスト候補（ページ区切りあり）: ${item.textSample.slice(0, PDF_TEXT_PROMPT_LIMIT)}` : "",
    ].filter(Boolean).join("\n");
  }

  function extractResponseText(payload) {
    if (typeof payload.output_text === "string" && payload.output_text.trim()) {
      return payload.output_text;
    }

    const geminiText = [];
    (payload.candidates || []).forEach(function (candidate) {
      ((candidate.content && candidate.content.parts) || []).forEach(function (part) {
        if (typeof part.text === "string") {
          geminiText.push(part.text);
        }
      });
    });

    if (geminiText.length > 0) {
      return geminiText.join("");
    }

    const chunks = [];
    (payload.output || []).forEach(function (outputItem) {
      (outputItem.content || []).forEach(function (contentItem) {
        if (typeof contentItem.text === "string") {
          chunks.push(contentItem.text);
        }
      });
    });

    if (chunks.length > 0) {
      return chunks.join("");
    }

    throw new Error("AI応答からJSONテキストを取得できませんでした。");
  }

  function parseExpectedJson() {
    const source = els.expectedJsonInput.value.trim();
    if (!source) {
      return {};
    }

    try {
      return JSON.parse(source);
    } catch (error) {
      throw new Error("期待値JSONの形式が正しくありません。");
    }
  }

  function findExpectedForItem(expectedValues, item) {
    const fullName = displayName(item.file);
    const fileName = item.file.name;
    const baseName = fileName.replace(/\.[^.]+$/, "");
    return expectedValues[fullName] || expectedValues[fileName] || expectedValues[baseName] || null;
  }

  function compareExtraction(extraction, expected) {
    if (!expected || typeof expected !== "object") {
      return { compared: 0, matched: 0, details: [] };
    }

    const details = [];
    let compared = 0;
    let matched = 0;

    activeExtractionFields.forEach(function (fieldItem) {
      const field = fieldKey(fieldItem);
      const label = fieldLabel(fieldItem);
      const expectedValue = expectedValueForField(expected, fieldItem);
      if (isBlank(expectedValue)) {
        return;
      }

      compared += 1;
      const ok = normalizeForCompare(field, extraction[field]) === normalizeForCompare(field, expectedValue);
      if (ok) {
        matched += 1;
      }
      details.push({
        field,
        label,
        ok,
        actual: extraction[field],
        expected: expectedValue,
      });
    });

    return { compared, matched, details };
  }

  function normalizeForCompare(field, value) {
    if (isBlank(value)) {
      return "";
    }

    if (field.startsWith("total_amount") || field.startsWith("subtotal_amount") || field.startsWith("tax_amount")) {
      const numeric = Number(String(value).replace(/[^\d.-]/g, ""));
      return Number.isFinite(numeric) ? String(Math.round(numeric)) : "";
    }

    if (["date", "payment_due_date"].includes(field)) {
      return normalizeDate(value);
    }

    if (field === "currency") {
      return String(value).trim().toUpperCase();
    }

    return String(value).replace(/\s+/g, "").toLowerCase();
  }

  function normalizeDate(value) {
    const source = String(value).trim();
    const match = source.match(/(\d{4})[年\/.-]\s*(\d{1,2})[月\/.-]\s*(\d{1,2})/);
    if (!match) {
      return source.replace(/\s+/g, "");
    }

    const year = match[1];
    const month = match[2].padStart(2, "0");
    const day = match[3].padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function appendAiPendingRow(item) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>
        <span class="file-name">${escapeHtml(displayName(item.file))}</span>
        <span class="file-meta">${escapeHtml(item.routeLabel)}</span>
      </td>
      <td><span class="spinner"></span>AI抽出中</td>
      <td>-</td>
      <td>-</td>
    `;
    els.aiResultRows.appendChild(tr);
    return tr;
  }

  function updateAiRow(row, item, extraction, comparison) {
    row.dataset.itemId = String(item.id);
    row.classList.add("selectable-row");
    row.addEventListener("click", function () {
      selectReviewItem(item.id);
    });
    row.innerHTML = `
      <td>
        <span class="file-name">${escapeHtml(displayName(item.file))}</span>
        <span class="file-meta">${escapeHtml(item.routeLabel)}</span>
      </td>
      <td>${renderFieldList(extraction)}</td>
      <td>${renderComparison(comparison)}</td>
      <td><pre class="json-block">${escapeHtml(JSON.stringify(extraction, null, 2))}</pre></td>
    `;
  }

  function updateAiErrorRow(row, item, error) {
    row.dataset.itemId = String(item.id);
    row.innerHTML = `
      <td>
        <span class="file-name">${escapeHtml(displayName(item.file))}</span>
        <span class="file-meta">${escapeHtml(item.routeLabel)}</span>
      </td>
      <td><span class="badge ocr">失敗</span></td>
      <td>-</td>
      <td><pre class="json-block">${escapeHtml(error.message || "AI抽出に失敗しました。")}</pre></td>
    `;
  }

  function renderFieldList(extraction) {
    const rows = activeExtractionFields.map(function (field) {
      const key = fieldKey(field);
      const label = fieldLabel(field);
      return `<div><dt>${escapeHtml(label)}</dt><dd>${escapeHtml(formatFieldValue(extraction[key]))}</dd></div>`;
    }).join("");

    return `<dl class="field-list">${rows}</dl>`;
  }

  function renderComparison(comparison) {
    if (comparison.compared === 0) {
      return '<span class="route-note">期待値なし</span>';
    }

    const score = Math.round((comparison.matched / comparison.compared) * 100);
    const scoreClass = score >= 80 ? "score" : "score warn";
    const details = comparison.details.map(function (detail) {
      const mark = detail.ok ? "OK" : "NG";
      return `<span>${mark} ${escapeHtml(detail.label)}: ${escapeHtml(formatFieldValue(detail.actual))} / ${escapeHtml(formatFieldValue(detail.expected))}</span>`;
    }).join("");

    return `<span class="${scoreClass}">${score}% (${comparison.matched}/${comparison.compared})</span><div class="match-list">${details}</div>`;
  }

  function selectReviewItem(itemId) {
    const item = state.items.find(function (candidate) {
      return candidate.id === itemId;
    });
    if (!item) {
      return;
    }

    state.currentReviewId = itemId;
    markSelectedRows(itemId);
    renderReviewList();
    renderReviewPanel(item, state.extractions.get(itemId) || {});
  }

  function markSelectedRows(itemId) {
    els.resultRows.querySelectorAll("tr[data-item-id]").forEach(function (row) {
      row.classList.toggle("is-selected", row.dataset.itemId === String(itemId));
    });
    els.aiResultRows.querySelectorAll("tr[data-item-id]").forEach(function (row) {
      row.classList.toggle("is-selected", row.dataset.itemId === String(itemId));
    });
  }

  function renderReviewList() {
    const reviewableItems = state.items.filter(function (item) {
      return Boolean(item.previewUrl);
    });

    if (reviewableItems.length === 0) {
      els.reviewList.innerHTML = '<span class="review-list-empty">書類をアップロードすると、ここに確認対象が並びます。</span>';
      return;
    }

    els.reviewList.innerHTML = reviewableItems.map(function (item) {
      const isSelected = item.id === state.currentReviewId;
      const status = state.extractions.has(item.id) ? "抽出済み" : "未抽出";
      return `
        <button type="button" class="review-list-item ${isSelected ? "is-active" : ""}" data-review-id="${item.id}">
          <span>${escapeHtml(displayName(item.file))}</span>
          <small>${escapeHtml(status)}</small>
        </button>
      `;
    }).join("");

    els.reviewList.querySelectorAll("[data-review-id]").forEach(function (button) {
      button.addEventListener("click", function () {
        selectReviewItem(Number(button.dataset.reviewId));
      });
    });
  }

  function renderReviewPanel(item, extraction) {
    const pageUrls = pagePreviewUrlsForItem(item);
    els.reviewFileName.textContent = displayName(item.file);
    els.reviewRouteLabel.textContent = reviewRouteText(item, pageUrls.length);
    els.reviewImage.src = pageUrls[0] || "";
    const viewerContent = renderReviewPages(pageUrls);
    els.reviewImage.parentElement.classList.toggle("has-image", Boolean(viewerContent));
    if (els.reviewPages) {
      els.reviewPages.innerHTML = viewerContent;
    }
    els.reviewEditStatus.textContent = state.extractions.has(item.id) ? "編集できます" : "AI抽出前";

    els.reviewForm.innerHTML = activeExtractionFields.map(function (field) {
      const key = fieldKey(field);
      const label = fieldLabel(field);
      const value = extraction[key] === null || extraction[key] === undefined ? "" : String(extraction[key]);
      const isLong = /内容|摘要|備考|明細|detail|description|notes|summary/.test(key + label);
      const input = isLong
        ? `<textarea data-field-key="${escapeHtml(key)}">${escapeHtml(value)}</textarea>`
        : `<input data-field-key="${escapeHtml(key)}" type="${isAmountField(key, label) ? "number" : "text"}" value="${escapeHtml(value)}" />`;
      return `
        <div class="review-field">
          <label for="review-${escapeHtml(key)}">${escapeHtml(label)}</label>
          ${input}
          <span class="review-field-meta">${escapeHtml(key)}</span>
        </div>
      `;
    }).join("");
  }

  function pagePreviewUrlsForItem(item) {
    const urls = Array.isArray(item.pagePreviewUrls) && item.pagePreviewUrls.length > 0
      ? item.pagePreviewUrls
      : [item.previewUrl];
    return urls.filter(Boolean);
  }

  function reviewRouteText(item, renderedPageCount) {
    const pageCount = item.pageCount || renderedPageCount;
    if (pageCount > 1) {
      const pageText = renderedPageCount >= pageCount
        ? `全${pageCount}ページをAI送信済み`
        : `${pageCount}ページ中、先頭${renderedPageCount}ページをAI送信済み`;
      return `${pageText}。内容を確認して、必要に応じて修正してください。`;
    }

    return "内容を確認して、必要に応じて修正してください。";
  }

  function renderReviewPages(pageUrls) {
    if (pageUrls.length === 0) {
      return "";
    }

    return pageUrls.map(function (url, index) {
      return `
        <figure class="viewer-page">
          <figcaption>${index + 1} / ${pageUrls.length}</figcaption>
          <img src="${escapeHtml(url)}" alt="ページ ${index + 1}" />
        </figure>
      `;
    }).join("");
  }

  function saveReviewEdits() {
    if (state.currentReviewId === null) {
      els.reviewEditStatus.textContent = "書類を選択してください";
      return;
    }

    const item = state.items.find(function (candidate) {
      return candidate.id === state.currentReviewId;
    });
    if (!item) {
      return;
    }

    const updated = {};
    activeExtractionFields.forEach(function (field) {
      const key = fieldKey(field);
      const label = fieldLabel(field);
      const input = els.reviewForm.querySelector(`[data-field-key="${cssEscape(key)}"]`);
      const rawValue = input ? input.value.trim() : "";
      if (!rawValue) {
        updated[key] = null;
      } else if (isAmountField(key, label)) {
        const numeric = Number(rawValue.replace(/[^\d.-]/g, ""));
        updated[key] = Number.isFinite(numeric) ? numeric : null;
      } else if (keyMatches(key, "receipt_method")) {
        updated[key] = normalizeReceiptMethod(rawValue, item);
      } else if (keyMatches(key, "payment_method")) {
        updated[key] = rawValue;
      } else {
        updated[key] = rawValue;
      }
    });

    activeExtractionFields.forEach(function (field) {
      const key = fieldKey(field);
      if (keyMatches(key, "receipt_method")) {
        updated[key] = normalizeReceiptMethod(updated[key], item);
      }
      if (keyMatches(key, "payment_method")) {
        updated[key] = normalizePaymentMethod(updated[key], updated, item);
      }
      if (keyMatches(key, "bank_account_type")) {
        updated[key] = normalizeBankAccountType(updated[key]);
      }
      if (keyMatches(key, "bank_account_number")) {
        updated[key] = normalizeBankAccountNumber(updated[key]);
      }
      if (keyMatches(key, "bank_account_kana")) {
        updated[key] = normalizeBankAccountKana(updated[key]);
      }
    });
    cleanPaymentContentFields(updated, item);
    fillExplicitPaymentContentFields(updated, item);
    fillSummaryFields(updated);

    state.extractions.set(state.currentReviewId, updated);
    els.reviewEditStatus.textContent = "反映しました";
    renderReviewList();
    refreshAiRowForItem(item, updated);
  }

  function refreshAiRowForItem(item, extraction) {
    const row = els.aiResultRows.querySelector(`tr[data-item-id="${item.id}"]`);
    if (!row) {
      return;
    }

    row.innerHTML = `
      <td>
        <span class="file-name">${escapeHtml(displayName(item.file))}</span>
        <span class="file-meta">${escapeHtml(item.routeLabel)}</span>
      </td>
      <td>${renderFieldList(extraction)}</td>
      <td><span class="route-note">手修正済み</span></td>
      <td><pre class="json-block">${escapeHtml(JSON.stringify(extraction, null, 2))}</pre></td>
    `;
    row.classList.add("selectable-row", "is-selected");
    row.addEventListener("click", function () {
      selectReviewItem(item.id);
    });
  }

  async function copyReviewJson() {
    if (state.currentReviewId === null || !state.extractions.has(state.currentReviewId)) {
      els.reviewEditStatus.textContent = "コピーする抽出データがありません";
      return;
    }

    const json = JSON.stringify(state.extractions.get(state.currentReviewId), null, 2);
    try {
      await navigator.clipboard.writeText(json);
      els.reviewEditStatus.textContent = "JSONをコピーしました";
    } catch (error) {
      els.reviewEditStatus.textContent = "コピーに失敗しました";
    }
  }

  async function exportSkyberryExcel() {
    if (!window.XLSX) {
      els.reviewEditStatus.textContent = "Excel出力ライブラリを読み込めません";
      return;
    }

    saveReviewEditsIfSelected();

    const rows = state.items
      .filter(function (item) {
        return state.extractions.has(item.id);
      })
      .map(function (item) {
        return {
          item,
          extraction: state.extractions.get(item.id),
        };
      });

    if (rows.length === 0) {
      els.reviewEditStatus.textContent = "出力する抽出データがありません";
      return;
    }

    els.exportSkyberryButton.disabled = true;
    els.reviewEditStatus.textContent = "Excel出力中...";

    try {
      const response = await fetch(DEFAULT_TEMPLATE_URL, { cache: "no-store" });
      if (!response.ok) {
        throw new Error(`テンプレートExcelを取得できませんでした (${response.status})`);
      }

      const workbook = await readTemplateWorkbook(response, { cellDates: true });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const headerRow = readSheetHeader(sheet);
      const originalRange = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");

      rows.forEach(function (row, rowIndex) {
        const excelRowNumber = rowIndex + 2;
        activeExtractionFields.forEach(function (field) {
          const key = fieldKey(field);
          const label = fieldLabel(field);
          const columnIndex = findHeaderColumn(headerRow, label, key, field);
          if (columnIndex < 0) {
            return;
          }

          const value = valueForSkyberryCell(row.extraction, key, rowIndex);
          if (value === undefined || value === null || value === "") {
            return;
          }

          const cellAddress = XLSX.utils.encode_cell({ r: excelRowNumber - 1, c: columnIndex });
          sheet[cellAddress] = cellForValue(value);
        });
      });

      const maxColumn = Math.max(headerRow.length - 1, originalRange.e.c, 0);
      const maxRow = Math.max(rows.length + 1, originalRange.e.r + 1);
      sheet["!ref"] = XLSX.utils.encode_range({ s: originalRange.s, e: { r: maxRow - 1, c: maxColumn } });

      const outputName = `skyberry-import-${formatDateForFileName(new Date())}.xlsx`;
      XLSX.writeFile(workbook, outputName);
      els.reviewEditStatus.textContent = `${rows.length}件をExcel出力しました`;
      sendClientLog("skyberry.excel.exported", {
        fileName: outputName,
        rows: rows.length,
        sheetName,
      });
    } catch (error) {
      els.reviewEditStatus.textContent = error.message || "Excel出力に失敗しました";
      sendClientLog("skyberry.excel.failed", {
        error: error.message || "Excel出力に失敗しました",
      });
    } finally {
      els.exportSkyberryButton.disabled = false;
    }
  }

  function saveReviewEditsIfSelected() {
    if (state.currentReviewId !== null && els.reviewForm.children.length > 0) {
      saveReviewEdits();
    }
  }

  function readSheetHeader(sheet) {
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    const headers = [];
    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: columnIndex });
      const cell = sheet[cellAddress];
      headers[columnIndex] = cell && cell.v !== undefined ? String(cell.v).trim() : "";
    }
    return headers;
  }

  function findHeaderColumn(headerRow, label, key, field) {
    const columnIndex = fieldColumnIndex(field);
    if (columnIndex >= 0 && normalizeHeader(headerRow[columnIndex] || "") === normalizeHeader(label)) {
      return columnIndex;
    }

    const normalizedLabel = normalizeHeader(label);
    const exactIndex = headerRow.findIndex(function (header) {
      return normalizeHeader(header) === normalizedLabel;
    });
    if (exactIndex >= 0) {
      return exactIndex;
    }

    const matchedField = activeExtractionFields.find(function (candidate) {
      return fieldKey(candidate) === key;
    });
    if (!matchedField) {
      return -1;
    }

    const aliases = fieldAliases(key, fieldLabel(matchedField)).map(normalizeHeader);
    return headerRow.findIndex(function (header) {
      return aliases.includes(normalizeHeader(header));
    });
  }

  function valueForSkyberryCell(extraction, key, rowIndex) {
    return extraction[key];
  }

  function cellForValue(value) {
    if (typeof value === "number") {
      return { t: "n", v: value };
    }

    const text = String(value);
    const numeric = Number(text.replace(/[^\d.-]/g, ""));
    if (text && /^-?[\d,]+(\.\d+)?$/.test(text) && Number.isFinite(numeric)) {
      return { t: "n", v: numeric };
    }

    return { t: "s", v: text };
  }

  function formatDateForFileName(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    const hour = String(date.getHours()).padStart(2, "0");
    const minute = String(date.getMinutes()).padStart(2, "0");
    return `${year}${month}${day}-${hour}${minute}`;
  }

  function clearReviewPanel() {
    state.currentReviewId = null;
    els.reviewFileName.textContent = "書類未選択";
    els.reviewRouteLabel.textContent = "書類をアップロードして次へ進んでください。";
    els.reviewImage.removeAttribute("src");
    els.reviewImage.parentElement.classList.remove("has-image");
    if (els.reviewPages) {
      els.reviewPages.innerHTML = "";
    }
    els.reviewForm.innerHTML = "";
    els.reviewEditStatus.textContent = "未選択";
    renderReviewList();
  }

  function showReviewPage() {
    const firstExtractedItem = state.items.find(function (item) {
      return state.extractions.has(item.id);
    });
    if (firstExtractedItem && !state.extractions.has(state.currentReviewId)) {
      selectReviewItem(firstExtractedItem.id);
    }

    els.uploadPage.classList.add("is-hidden");
    els.reviewPage.classList.remove("is-hidden");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function showUploadPage() {
    els.reviewPage.classList.add("is-hidden");
    els.uploadPage.classList.remove("is-hidden");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function cssEscape(value) {
    if (window.CSS && typeof window.CSS.escape === "function") {
      return window.CSS.escape(value);
    }
    return String(value).replace(/"/g, '\\"');
  }

  function labelForField(field) {
    const found = activeExtractionFields.find(function (item) {
      return fieldKey(item) === field;
    });
    return found ? fieldLabel(found) : field;
  }

  function expectedValueForField(expected, field) {
    const key = fieldKey(field);
    const label = fieldLabel(field);
    if (Object.prototype.hasOwnProperty.call(expected, key)) {
      return expected[key];
    }
    if (Object.prototype.hasOwnProperty.call(expected, label)) {
      return expected[label];
    }
    return undefined;
  }

  function fieldKey(field) {
    return Array.isArray(field) ? field[0] : field.key;
  }

  function fieldLabel(field) {
    return Array.isArray(field) ? field[1] : field.label;
  }

  function fieldColumnIndex(field) {
    if (!field || Array.isArray(field) || typeof field.columnIndex !== "number") {
      return -1;
    }
    return field.columnIndex;
  }

  function hasField(baseKey) {
    return activeExtractionFields.some(function (field) {
      return keyMatches(fieldKey(field), baseKey);
    });
  }

  function formatFieldValue(value) {
    if (isBlank(value)) {
      return "-";
    }
    return String(value);
  }

  function isBlank(value) {
    return value === null || value === undefined || String(value).trim() === "";
  }

  function setAiStatus(message, mode) {
    els.aiStatus.textContent = message;
    els.aiStatus.className = mode ? `ai-status ${mode}` : "ai-status";
  }

  function setTemplateStatus(message, mode) {
    els.excelTemplateStatus.textContent = message;
    els.excelTemplateStatus.className = mode ? `template-status ${mode}` : "template-status";
  }

  function updateProgress(options) {
    const total = Math.max(0, Number(options.total) || 0);
    const current = Math.max(0, Math.min(total, Number(options.current) || 0));
    const percent = total === 0 ? 0 : Math.round((current / total) * 100);

    els.progressPanel.classList.remove("is-hidden");
    els.progressTitle.textContent = options.title || "処理中";
    els.progressCount.textContent = `${current} / ${total}`;
    els.progressCurrent.textContent = options.currentText || "";
    els.progressBar.style.width = `${percent}%`;
  }

  function hideProgress() {
    els.progressPanel.classList.add("is-hidden");
    els.progressTitle.textContent = "処理中";
    els.progressCount.textContent = "0 / 0";
    els.progressCurrent.textContent = "待機中";
    els.progressBar.style.width = "0%";
  }

  function updateNextButtonState() {
    els.runAiButton.disabled = state.isReadingFiles || state.isExtracting;
    if (state.isReadingFiles) {
      els.runAiButton.textContent = "読込中...";
      return;
    }
    if (state.isExtracting) {
      els.runAiButton.textContent = "処理中...";
      return;
    }
    els.runAiButton.textContent = "次へ";
  }

  function sendClientLog(eventName, detail) {
    fetch(apiUrl("/api/log"), {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        event: eventName,
        at: new Date().toISOString(),
        detail: detail || {},
      }),
    }).catch(function () {
      // ログ送信は画面操作を止めない。
    });
  }

  function renderSummary() {
    els.totalCount.textContent = String(state.total);
    els.imageCount.textContent = String(state.image);
    els.pdfCount.textContent = String(state.pdf);
    els.ocrCount.textContent = String(state.ocr);
  }

  function recalculateSummary() {
    state.total = state.selectedFiles.length;
    state.image = 0;
    state.pdf = 0;
    state.ocr = 0;

    state.items.forEach(function (item) {
      if (item.route === "pdf") {
        state.pdf += 1;
        return;
      }

      if (item.route === "image") {
        state.image += 1;
        return;
      }

      if (item.route === "ocr") {
        state.ocr += 1;
        if (item.previewUrl && item.routeLabel === "画像PDF") {
          state.image += 1;
        }
      }
    });
  }

  function renderSelectedFiles() {
    els.selectedFileCount.textContent = `${state.selectedFiles.length}件`;
    if (state.selectedFiles.length === 0) {
      els.selectedFileSummary.textContent = "未選択";
      els.selectedFileList.innerHTML = '<li class="selected-file-empty">まだ選択されていません。</li>';
      return;
    }

    const loadedCount = state.selectedFiles.filter(function (entry) {
      return entry.itemId !== null;
    }).length;
    const remainingCount = state.selectedFiles.length - loadedCount;
    els.selectedFileSummary.textContent = remainingCount > 0
      ? `${loadedCount}件読込済み / ${remainingCount}件処理中`
      : `${loadedCount}件すべて読込済み`;

    els.selectedFileList.innerHTML = state.selectedFiles.map(function (entry) {
      const name = displayName(entry.file);
      const statusKey = entry.statusKey || "pending";
      const disabled = state.isExtracting ? "disabled" : "";
      return `
        <li class="selected-file-item is-${statusKey}">
          <div class="selected-file-detail">
            <strong title="${escapeHtml(name)}">${escapeHtml(name)}</strong>
            <div class="selected-file-meta-row">
              <span>${formatBytes(entry.file.size)}</span>
              <span class="selected-file-status ${statusKey}">${escapeHtml(entry.status)}</span>
            </div>
          </div>
          <button
            type="button"
            class="selected-file-remove"
            data-remove-selected-file="${entry.id}"
            aria-label="${escapeHtml(name)}を削除"
            ${disabled}
          >削除</button>
        </li>
      `;
    }).join("");
  }

  function updateSelectedFileStatus(itemId, status, statusKey) {
    const entry = state.selectedFiles.find(function (candidate) {
      return candidate.itemId === itemId;
    });
    if (!entry) {
      return;
    }
    entry.status = status;
    if (statusKey) {
      entry.statusKey = statusKey;
    }
    renderSelectedFiles();
  }

  async function collectDroppedFiles(dataTransfer) {
    const items = Array.from(dataTransfer.items || []);
    const entryItems = items
      .map(function (item) {
        return typeof item.webkitGetAsEntry === "function" ? item.webkitGetAsEntry() : null;
      })
      .filter(Boolean);

    if (entryItems.length === 0) {
      return Array.from(dataTransfer.files || []);
    }

    const nested = await Promise.all(entryItems.map(function (entry) {
      return collectEntryFiles(entry, "");
    }));
    return nested.flat();
  }

  async function collectEntryFiles(entry, parentPath) {
    if (entry.isFile) {
      return new Promise(function (resolve, reject) {
        entry.file(function (file) {
          file.relativePath = `${parentPath}${file.name}`;
          resolve([file]);
        }, reject);
      });
    }

    if (!entry.isDirectory) {
      return [];
    }

    const reader = entry.createReader();
    const children = [];
    let batch = [];

    do {
      batch = await new Promise(function (resolve, reject) {
        reader.readEntries(resolve, reject);
      });
      children.push(...batch);
    } while (batch.length > 0);

    const nested = await Promise.all(children.map(function (child) {
      return collectEntryFiles(child, `${parentPath}${entry.name}/`);
    }));
    return nested.flat();
  }

  function isSupportedFile(file) {
    const ext = extensionOf(file.name);
    return ext === "pdf" || IMAGE_TYPES.has(ext);
  }

  function displayName(file) {
    return file.webkitRelativePath || file.relativePath || file.name;
  }

  function extensionOf(fileName) {
    const parts = fileName.toLowerCase().split(".");
    return parts.length > 1 ? parts.pop() : "";
  }

  function normalizeText(value) {
    return String(value)
      .replace(/[ \t\f\v]+/g, " ")
      .replace(/\s*\n\s*/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  function formatBytes(bytes) {
    if (!bytes) {
      return "0 B";
    }
    const units = ["B", "KB", "MB", "GB"];
    const exponent = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
    const value = bytes / 1024 ** exponent;
    return `${value.toFixed(value >= 10 || exponent === 0 ? 0 : 1)} ${units[exponent]}`;
  }

  function apiUrl(path) {
    if (!API_BASE_URL) {
      return path;
    }
    return `${API_BASE_URL}${path}`;
  }

  function normalizeApiBaseUrl(value) {
    return String(value || "").replace(/\/+$/, "");
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function setLibraryStatus(message, mode) {
    els.libraryStatus.textContent = message;
    els.libraryStatus.className = `status-pill dev-only ${mode}`;
  }
})();
