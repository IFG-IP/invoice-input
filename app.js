(function () {
  const DIGITAL_TEXT_THRESHOLD = 20;
  const PDF_RENDER_SCALE = 1.8;
  const DEFAULT_TEMPLATE_URL = "/api/template/extraction-fields";
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
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
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
      normalizedBytes: 0,
    };
  }

  async function analyzePdf(file) {
    if (!window.pdfjsLib) {
      throw new Error("PDF.js を読み込めないためPDF解析をスキップしました。");
    }

    const bytes = new Uint8Array(await file.arrayBuffer());
    const loadingTask = pdfjsLib.getDocument({ data: bytes });
    const pdf = await loadingTask.promise;
    const text = await extractPdfText(pdf);
    const normalizedText = normalizeText(text);

    if (normalizedText.length >= DIGITAL_TEXT_THRESHOLD) {
      state.pdf += 1;
      const previewUrl = await renderPdfPageToImage(pdf, 1);
      await pdf.destroy();
      return {
        file,
        route: "pdf",
        routeLabel: "デジタルPDF",
        note: "PDF内部テキストをそのまま利用できます。",
        textLength: normalizedText.length,
        textSample: normalizedText,
        previewUrl,
        normalizedBytes: 0,
      };
    }

    state.ocr += 1;
    state.image += 1;
    const previewUrl = await renderPdfPageToImage(pdf, 1);
    const normalizedBytes = await estimateImageBytes(previewUrl);
    await pdf.destroy();
    return {
      file,
      route: "ocr",
      routeLabel: "画像PDF",
      note: "PDFページを画像化して画像ルートへ合流できます。",
      textLength: normalizedText.length,
      textSample: normalizedText,
      previewUrl,
      normalizedBytes,
    };
  }

  async function extractPdfText(pdf) {
    const chunks = [];
    const pageLimit = Math.min(pdf.numPages, 3);

    for (let pageNumber = 1; pageNumber <= pageLimit; pageNumber += 1) {
      const page = await pdf.getPage(pageNumber);
      const textContent = await page.getTextContent();
      chunks.push(textContent.items.map(function (item) {
        return item.str || "";
      }).join(" "));
      page.cleanup();
    }

    return chunks.join("\n");
  }

  async function renderPdfPageToImage(pdf, pageNumber) {
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale: PDF_RENDER_SCALE });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d", { alpha: false });
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);

    await page.render({ canvasContext: context, viewport }).promise;
    page.cleanup();
    return resizeCanvasToJpeg(canvas);
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

  function resizeCanvasToJpeg(canvas) {
    const maxEdge = Number(els.maxEdge.value);
    const ratio = Math.min(1, maxEdge / Math.max(canvas.width, canvas.height));
    const output = document.createElement("canvas");
    output.width = Math.max(1, Math.round(canvas.width * ratio));
    output.height = Math.max(1, Math.round(canvas.height * ratio));
    const outputContext = output.getContext("2d", { alpha: false });

    outputContext.fillStyle = "#ffffff";
    outputContext.fillRect(0, 0, output.width, output.height);
    outputContext.drawImage(canvas, 0, 0, output.width, output.height);

    return output.toDataURL("image/jpeg", Number(els.jpegQuality.value));
  }

  async function estimateImageBytes(dataUrl) {
    if (!dataUrl) {
      return 0;
    }
    const response = await fetch(dataUrl);
    const blob = await response.blob();
    return blob.size;
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
      ? `<span class="text-sample">${escapeHtml(result.textSample)}</span>`
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
    const response = await fetch("/api/gemini/generate", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model,
        contents: [
          {
            role: "user",
            parts: [
              {
                text: buildExtractionPrompt(item),
              },
              dataUrlToGeminiInlineData(item.previewUrl),
            ],
          },
        ],
        generationConfig: {
          responseMimeType: "application/json",
          responseSchema: activeExtractionSchema,
          maxOutputTokens: 1200,
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
    return JSON.parse(text);
  }

  async function ensureApiProxyReady() {
    let response;
    try {
      response = await fetch("/api/health", { cache: "no-store" });
    } catch (error) {
      throw new Error("ローカルAPIプロキシに接続できません。open-local.cmdで起動し直してください。");
    }

    if (!response.ok) {
      throw new Error("Gemini APIプロキシが見つかりません。open-local.cmdでサーバーを起動し直してください。");
    }
  }

  async function loadApiDefaults() {
    try {
      const response = await fetch("/api/health", { cache: "no-store" });
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

      const data = await response.arrayBuffer();
      const templateSource = response.headers.get("X-Template-Source");
      const templateName = response.headers.get("X-Template-Name");
      const source = templateSource === "local-workbook"
        ? `ローカルExcel (${decodeURIComponent(templateName || "")})`
        : templateSource === "sharepoint"
          ? "SharePointテンプレート"
          : "ローカル既定テンプレート";
      applyTemplateWorkbook(data, source);
    } catch (error) {
      setTemplateStatus(error.message || "既定テンプレートの読込に失敗しました。", "error");
      sendClientLog("template.failed", {
        error: error.message || "既定テンプレートの読込に失敗しました。",
      });
    }
  }

  function applyTemplateWorkbook(data, label) {
    const workbook = XLSX.read(data);
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
      [/^支払日$/, "payment_due_date"],
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
      [/^口座種別/, "bank_account_type"],
      [/^口座番号$/, "bank_account_number"],
      [/^口座名義フリガナ$/, "bank_account_kana"],
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
      [/^(取引先|会社名|店舗名|店名|宛先|仕入先|請求元)$/, "vendor"],
      [/^(合計金額|合計|総額|金額|請求金額|税込金額|領収金額)$/, "total_amount"],
      [/^(小計|税抜金額|本体金額)$/, "subtotal_amount"],
      [/^(税額|消費税|消費税額|内税|外税)$/, "tax_amount"],
      [/^(請求書番号|請求番号|領収書番号|伝票番号|管理番号|番号)$/, "invoice_number"],
      [/^(支払期日|支払期限|振込期限|期限|入金期日)$/, "payment_due_date"],
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

  function keyMatches(key, baseKey) {
    return key === baseKey || key.startsWith(`${baseKey}_`);
  }

  function fieldMeaning(key, label) {
    if (keyMatches(key, "vendor")) {
      return "支払先・請求元・領収書発行者・店舗/会社名。";
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
      return "支払内容・摘要・備考。品名、但し書き、利用内容、請求内容が該当する。";
    }
    if (keyMatches(key, "payment_method") || keyMatches(key, "receipt_method")) {
      return "支払方法または受領方法。銀行振込、クレジット、現金、口座引落など。";
    }
    if (keyMatches(key, "payment_due_date")) {
      return "支払日・支払期日・振込期限。請求書発行日とは区別する。";
    }
    if (keyMatches(key, "invoice_issue_date") || keyMatches(key, "date")) {
      return "請求書・領収書の発行日、請求日、領収日、取引日。支払期日とは区別する。";
    }
    if (keyMatches(key, "invoice_number")) {
      return "請求書番号・領収書番号・伝票番号・No.。";
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
      add(["支払先", "請求元", "発行者", "領収者", "販売者", "店舗名", "店名", "会社名", "宛先", "加盟店", "取引先", "仕入先"]);
    }
    if (key.startsWith("subtotal_amount")) {
      add(["税抜金額", "小計", "本体価格", "本体金額", "税抜合計", "支払金額(税抜)", "金額(税抜)"]);
    }
    if (key.startsWith("total_amount")) {
      add(["合計", "合計金額", "総合計", "税込金額", "請求金額", "領収金額", "お支払金額", "ご請求額", "支払額", "税込合計"]);
    }
    if (key.startsWith("tax_amount")) {
      add(["消費税", "消費税額", "税額", "内消費税", "内税", "外税", "10%対象税額"]);
    }
    if (keyMatches(key, "payment_description") || keyMatches(key, "summary") || keyMatches(key, "notes")) {
      add(["支払内容", "摘要", "備考", "但し書き", "品名", "内容", "利用内容", "明細", "内訳", "件名", "請求内容"]);
    }
    if (keyMatches(key, "payment_method")) {
      add(["支払方法", "決済方法", "お支払方法", "支払い方法", "決済種別", "現金", "クレジット", "銀行振込", "口座振替"]);
    }
    if (keyMatches(key, "receipt_method")) {
      add(["受領方法", "受取方法", "受領区分", "紙", "データ", "電子", "原本"]);
    }
    if (keyMatches(key, "payment_due_date")) {
      add(["支払日", "支払期日", "支払期限", "振込期限", "入金期日", "お支払期限", "引落日"]);
    }
    if (keyMatches(key, "invoice_issue_date") || keyMatches(key, "date")) {
      add(["発行日", "請求日", "領収日", "取引日", "利用日", "日付", "年月日", "発行年月日"]);
    }
    if (keyMatches(key, "invoice_number")) {
      add(["請求書番号", "請求番号", "領収書番号", "伝票番号", "管理番号", "No.", "No", "番号", "Invoice No"]);
    }
    if (keyMatches(key, "currency") || key.includes("_jpy") || key.includes("_usd") || key.includes("_other")) {
      add(["通貨", "通貨コード", "JPY", "円", "USD", "ドル", "$", "外貨"]);
    }
    if (key === "bank_name") {
      add(["銀行名", "金融機関名", "振込先銀行", "お振込先"]);
    }
    if (key === "bank_branch_name") {
      add(["支店名", "支店", "支店コード", "店番"]);
    }
    if (key === "bank_account_type") {
      add(["口座種別", "預金種目", "普通", "当座"]);
    }
    if (key === "bank_account_number") {
      add(["口座番号", "口座No", "口座"]);
    }
    if (key === "bank_account_kana") {
      add(["口座名義", "口座名義カナ", "口座名義フリガナ", "受取人名"]);
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
      "あなたは日本語の請求書・領収書の項目抽出エンジンです。",
      "目的は、Excelの各列に入れる値を、書類画像またはPDF本文から抽出することです。",
      "最重要ルール: 書類に明記されていない項目は必ずnullにしてください。推測、補完、一般常識、ファイル名からの推定は禁止です。",
      "社内申請用のコード、稟議No、経理CD、科目、補助、税区、コピー/原本/添付管理欄は、書類上に同じ項目名と値が明記されていない限りnullです。",
      "列名が似ていても別項目へ値を流用しないでください。例: 請求書発行日と支払日は別、税抜金額と合計金額は別、支払先名と取引先コードは別です。",
      "Excel列名と書類上の項目名は一致しない前提です。列ごとの意味と表記ゆれ候補を見て、意味が一致する値だけを抽出してください。",
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
      item.textSample ? `PDF内部テキスト候補: ${item.textSample.slice(0, 1200)}` : "",
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
    els.reviewFileName.textContent = displayName(item.file);
    els.reviewRouteLabel.textContent = "内容を確認して、必要に応じて修正してください。";
    els.reviewImage.src = item.previewUrl || "";
    els.reviewImage.parentElement.classList.toggle("has-image", Boolean(item.previewUrl));
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
      } else {
        updated[key] = rawValue;
      }
    });

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

      const data = await response.arrayBuffer();
      const workbook = XLSX.read(data, { cellDates: true });
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
    if (key === "row_no" && isBlank(extraction[key])) {
      return rowIndex + 1;
    }
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
    fetch("/api/log", {
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
    return value.replace(/\s+/g, " ").trim();
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
