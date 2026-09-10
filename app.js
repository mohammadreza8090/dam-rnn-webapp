/**
 * Dam Displacement Prediction — fully client-side RNN app.
 *
 * Everything below (Excel parsing, scaling, sequence building, model
 * training and inference, plotting) runs entirely in the user's browser.
 * There is no backend: this file, together with index.html and style.css,
 * can be hosted as a static site (e.g. GitHub Pages) and will keep
 * working forever without any server to maintain.
 *
 * Libraries used (all loaded from CDNs in index.html):
 *   - SheetJS (xlsx.js)   -> reads .xlsx files in the browser
 *   - TensorFlow.js        -> builds/trains/runs the RNN in the browser
 *   - Plotly.js             -> interactive charts
 */

// ---------------------------------------------------------------------
// Application state
// ---------------------------------------------------------------------
const state = {
  headers: [],   // column names found in the uploaded file
  rows: [],      // parsed rows, as an array of plain objects {colName: value}
};

// ---------------------------------------------------------------------
// DOM references
// ---------------------------------------------------------------------
const el = {
  dropzone: document.getElementById('dropzone'),
  dropzoneLabel: document.getElementById('dropzone-label'),
  fileInput: document.getElementById('file-input'),
  fileStatus: document.getElementById('file-status'),

  columnSection: document.getElementById('column-section'),
  dateCol: document.getElementById('date-col'),
  targetCol: document.getElementById('target-col'),
  modeRadios: document.querySelectorAll('input[name="mode"]'),
  multivariateCols: document.getElementById('multivariate-cols'),
  waterCol: document.getElementById('water-col'),
  tempCol: document.getElementById('temp-col'),

  splitSection: document.getElementById('split-section'),
  trainPct: document.getElementById('train-pct'),
  trainPctOut: document.getElementById('train-pct-out'),
  valPct: document.getElementById('val-pct'),
  valPctOut: document.getElementById('val-pct-out'),

  modelSection: document.getElementById('model-section'),
  seqLength: document.getElementById('seq-length'),
  rnnUnits1: document.getElementById('rnn-units-1'),
  rnnUnits2: document.getElementById('rnn-units-2'),
  dropoutRate: document.getElementById('dropout-rate'),
  learningRate: document.getElementById('learning-rate'),

  trainingSection: document.getElementById('training-section'),
  epochs: document.getElementById('epochs'),
  batchSize: document.getElementById('batch-size'),
  patience: document.getElementById('patience'),

  runButton: document.getElementById('run-button'),

  rawPreview: document.getElementById('raw-preview'),
  rawPlot: document.getElementById('raw-plot'),

  trainingProgress: document.getElementById('training-progress'),
  progressFill: document.getElementById('progress-fill'),
  progressLabel: document.getElementById('progress-label'),

  metricsCard: document.getElementById('metrics-card'),
  metricsTableBody: document.querySelector('#metrics-table tbody'),

  mainPlotCard: document.getElementById('main-plot-card'),
  mainPlot: document.getElementById('main-plot'),

  residualPlotCard: document.getElementById('residual-plot-card'),
  residualPlot: document.getElementById('residual-plot'),

  scatterPlotCard: document.getElementById('scatter-plot-card'),
  scatterPlot: document.getElementById('scatter-plot'),

  emptyState: document.getElementById('empty-state'),
};

// Sanity-check that the required CDN libraries actually loaded. This is the
// most common cause of "nothing happens" reports: an ad-blocker, browser
// extension, or a restricted network silently blocks the CDN request, so
// XLSX/tf/Plotly stay undefined and every action fails silently.
(function checkLibraries() {
  const missing = [];
  if (typeof XLSX === 'undefined') missing.push('SheetJS (xlsx.js)');
  if (typeof tf === 'undefined') missing.push('TensorFlow.js');
  if (typeof Plotly === 'undefined') missing.push('Plotly.js');

  if (missing.length) {
    const status = document.getElementById('file-status');
    if (status) {
      status.textContent =
        `⚠️ Failed to load: ${missing.join(', ')}. This is usually caused by an ` +
        `ad-blocker/extension or a network that blocks CDN requests (cdnjs.cloudflare.com, ` +
        `cdn.jsdelivr.net, cdn.plot.ly). Try disabling extensions, switching networks, or using a VPN.`;
      status.style.color = '#b3261e';
    }
  }
})();

// Colorblind-safe pair (Okabe-Ito): blue for actual data, amber for predictions.
const COLOR_ACTUAL = '#0072B2';
const COLOR_PREDICTED = '#E69F00';
const COLOR_SPLIT_LINE = '#7A8C94';
const SPLIT_COLORS = { Train: '#0072B2', Validation: '#E69F00', Test: '#2F7D5A' };

// =======================================================================
// 1. File loading
// =======================================================================

el.dropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  el.dropzone.classList.add('dragover');
});
el.dropzone.addEventListener('dragleave', () => el.dropzone.classList.remove('dragover'));
el.dropzone.addEventListener('drop', (e) => {
  e.preventDefault();
  el.dropzone.classList.remove('dragover');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
el.fileInput.addEventListener('change', (e) => {
  if (e.target.files.length) handleFile(e.target.files[0]);
});

function handleFile(file) {
  el.fileStatus.textContent = `Reading "${file.name}"…`;

  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const data = new Uint8Array(event.target.result);
      // cellDates: true makes SheetJS return JS Date objects for date cells
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { raw: true, defval: null });

      if (!rows.length) {
        el.fileStatus.textContent = 'The file appears to be empty.';
        return;
      }

      state.headers = Object.keys(rows[0]);
      state.rows = rows;

      el.fileStatus.textContent = `Loaded "${file.name}" — ${rows.length} rows, ${state.headers.length} columns.`;
      el.dropzoneLabel.textContent = file.name;

      populateColumnSelectors();
      revealConfigSections();
    } catch (err) {
      console.error(err);
      el.fileStatus.textContent = `Could not read this file: ${err.message}`;
    }
  };
  reader.readAsArrayBuffer(file);
}

function populateColumnSelectors() {
  const buildOptions = (selectEl, guessFn) => {
    selectEl.innerHTML = '';
    state.headers.forEach((h) => {
      const opt = document.createElement('option');
      opt.value = h;
      opt.textContent = h;
      selectEl.appendChild(opt);
    });
    const guess = guessFn(state.headers);
    if (guess) selectEl.value = guess;
  };

  buildOptions(el.dateCol, (headers) => headers.find((h) => /date/i.test(h)));
  buildOptions(el.targetCol, (headers) => headers.find((h) => !/date|water|level|temp/i.test(h)));
  buildOptions(el.waterCol, (headers) => headers.find((h) => /water|level|reservoir/i.test(h)));
  buildOptions(el.tempCol, (headers) => headers.find((h) => /temp/i.test(h)));

  plotRawSeries();
}

function revealConfigSections() {
  el.columnSection.hidden = false;
  el.splitSection.hidden = false;
  el.modelSection.hidden = false;
  el.trainingSection.hidden = false;
  el.runButton.hidden = false;
  el.rawPreview.hidden = false;
  el.emptyState.hidden = true;
}

// =======================================================================
// 2. UI wiring (mode toggle, collapsibles, sliders, re-plot on change)
// =======================================================================

el.modeRadios.forEach((radio) => {
  radio.addEventListener('change', () => {
    const isMultivariate = document.querySelector('input[name="mode"]:checked').value === 'multivariate';
    el.multivariateCols.hidden = !isMultivariate;
  });
});

document.querySelectorAll('.collapsible__toggle').forEach((btn) => {
  btn.addEventListener('click', () => {
    const section = btn.closest('.collapsible');
    const isOpen = section.getAttribute('data-open') === 'true';
    section.setAttribute('data-open', String(!isOpen));
    btn.setAttribute('aria-expanded', String(!isOpen));
  });
});

el.trainPct.addEventListener('input', () => { el.trainPctOut.textContent = el.trainPct.value; });
el.valPct.addEventListener('input', () => { el.valPctOut.textContent = el.valPct.value; });

el.targetCol.addEventListener('change', plotRawSeries);
el.dateCol.addEventListener('change', plotRawSeries);

function plotRawSeries() {
  const dateKey = el.dateCol.value;
  const targetKey = el.targetCol.value;
  if (!dateKey || !targetKey) return;

  const sorted = [...state.rows].sort((a, b) => toDate(a[dateKey]) - toDate(b[dateKey]));
  const dates = sorted.map((r) => toDate(r[dateKey]));
  const values = sorted.map((r) => Number(r[targetKey]));

  Plotly.newPlot(el.rawPlot, [{
    x: dates, y: values, type: 'scatter', mode: 'lines',
    line: { color: COLOR_ACTUAL, width: 1.5 },
    name: targetKey,
  }], baseLayout(`${targetKey} over time`, 'Date', `${targetKey} (mm)`), { displayModeBar: false, responsive: true });
}

function toDate(value) {
  if (value instanceof Date) return value;
  if (typeof value === 'number') {
    // Fallback for Excel serial dates that SheetJS did not auto-convert
    return XLSX.SSF.parse_date_code(value) ? new Date(Date.UTC(1899, 11, 30) + value * 86400000) : new Date(NaN);
  }
  return new Date(value);
}

function baseLayout(title, xTitle, yTitle) {
  return {
    title: { text: title, font: { family: 'Inter, sans-serif', size: 14 } },
    xaxis: { title: xTitle, gridcolor: '#e6ebec' },
    yaxis: { title: yTitle, gridcolor: '#e6ebec' },
    margin: { t: 40, r: 20, b: 45, l: 55 },
    paper_bgcolor: 'white',
    plot_bgcolor: 'white',
    font: { family: 'Inter, sans-serif', size: 12, color: '#1b2b33' },
    legend: { orientation: 'h', y: -0.2 },
  };
}

// =======================================================================
// 3. Min-max scaler (mirrors sklearn.preprocessing.MinMaxScaler)
// =======================================================================

class MinMaxScaler {
  fit(values) {
    this.dataMin = Math.min(...values);
    this.dataMax = Math.max(...values);
    this.range = (this.dataMax - this.dataMin) || 1e-8; // avoid div-by-zero on constant columns
    return this;
  }
  transform(values) {
    return values.map((v) => (v - this.dataMin) / this.range);
  }
  inverseTransform(values) {
    return values.map((v) => v * this.range + this.dataMin);
  }
}

// =======================================================================
// 4. Sequence building (mirrors the original create_sequences() helper)
// =======================================================================

function createSequences(matrix, window, targetIndex) {
  const X = [];
  const y = [];
  for (let i = 0; i <= matrix.length - window - 1; i++) {
    X.push(matrix.slice(i, i + window));
    y.push([matrix[i + window][targetIndex]]);
  }
  return { X, y };
}

// =======================================================================
// 5. Evaluation metrics (mirrors sklearn.metrics)
// =======================================================================

function evaluateMetrics(yTrue, yPred) {
  const n = yTrue.length;
  const meanY = yTrue.reduce((a, b) => a + b, 0) / n;

  let sumSqErr = 0;
  let sumAbsErr = 0;
  let sumSqTotal = 0;

  for (let i = 0; i < n; i++) {
    const err = yTrue[i] - yPred[i];
    sumSqErr += err * err;
    sumAbsErr += Math.abs(err);
    sumSqTotal += (yTrue[i] - meanY) ** 2;
  }

  const mse = sumSqErr / n;
  return {
    r2: 1 - sumSqErr / (sumSqTotal || 1e-8),
    mae: sumAbsErr / n,
    mse,
    rmse: Math.sqrt(mse),
  };
}

// =======================================================================
// 6. Main pipeline: run on "Run model" click
// =======================================================================

el.runButton.addEventListener('click', runPipeline);

async function runPipeline() {
  el.runButton.disabled = true;
  el.trainingProgress.hidden = false;
  el.progressFill.style.width = '0%';
  el.progressLabel.textContent = 'Preparing data…';

  try {
    // --- Read configuration from the UI ---------------------------------
    const dateKey = el.dateCol.value;
    const targetKey = el.targetCol.value;
    const mode = document.querySelector('input[name="mode"]:checked').value;
    const waterKey = el.waterCol.value;
    const tempKey = el.tempCol.value;

    const trainPct = Number(el.trainPct.value);
    const valPct = Number(el.valPct.value);
    const seqLen = Number(el.seqLength.value);
    const units1 = Number(el.rnnUnits1.value);
    const units2 = Number(el.rnnUnits2.value);
    const dropoutRate = Number(el.dropoutRate.value);
    const learningRate = Number(el.learningRate.value);
    const epochs = Number(el.epochs.value);
    const batchSize = Number(el.batchSize.value);
    const patience = Number(el.patience.value);

    // --- Build the feature list and target index -------------------------
    // Univariate mode:   features = [target]
    // Multivariate mode: features = [target, water level, temperature]
    const featureKeys = mode === 'multivariate' ? [targetKey, waterKey, tempKey] : [targetKey];
    const targetIndex = 0; // target is always the first feature column

    // --- Sort rows chronologically and build the raw numeric matrix ------
    const sorted = [...state.rows].sort((a, b) => toDate(a[dateKey]) - toDate(b[dateKey]));
    const dates = sorted.map((r) => toDate(r[dateKey]));
    const matrix = sorted.map((r) => featureKeys.map((k) => Number(r[k])));

    const n = matrix.length;
    const trainSize = Math.floor(n * (trainPct / 100));
    const valSize = Math.floor(n * (valPct / 100));

    const trainRaw = matrix.slice(0, trainSize);
    const valRaw = matrix.slice(trainSize, trainSize + valSize);
    const testRaw = matrix.slice(trainSize + valSize);

    // --- Fit one MinMaxScaler per feature, using the training split only -
    const scalers = featureKeys.map((_, f) => new MinMaxScaler().fit(trainRaw.map((row) => row[f])));
    const scaleMatrix = (mat) => mat.map((row) => row.map((v, f) => scalers[f].transform([v])[0]));

    const trainScaled = scaleMatrix(trainRaw);
    const valScaled = scaleMatrix(valRaw);
    const testScaled = scaleMatrix(testRaw);

    // --- Build sliding-window sequences, carrying the trailing window ----
    // from the previous split forward (same approach as the original
    // Python pipeline) so validation/test predictions are not starved of
    // context right at the split boundary.
    const { X: XTrainArr, y: yTrainArr } = createSequences(trainScaled, seqLen, targetIndex);
    const valInput = trainScaled.slice(-seqLen).concat(valScaled);
    const { X: XValArr, y: yValArr } = createSequences(valInput, seqLen, targetIndex);
    const testInput = valScaled.slice(-seqLen).concat(testScaled);
    const { X: XTestArr, y: yTestArr } = createSequences(testInput, seqLen, targetIndex);

    const numFeatures = featureKeys.length;
    const XTrain = tf.tensor3d(XTrainArr, [XTrainArr.length, seqLen, numFeatures]);
    const yTrain = tf.tensor2d(yTrainArr, [yTrainArr.length, 1]);
    const XVal = tf.tensor3d(XValArr, [XValArr.length, seqLen, numFeatures]);
    const yVal = tf.tensor2d(yValArr, [yValArr.length, 1]);
    const XTest = tf.tensor3d(XTestArr, [XTestArr.length, seqLen, numFeatures]);
    const yTest = tf.tensor2d(yTestArr, [yTestArr.length, 1]);

    // --- Build the model ---------------------------------------------------
    const model = tf.sequential();
    model.add(tf.layers.simpleRNN({ units: units1, returnSequences: true, inputShape: [seqLen, numFeatures] }));
    model.add(tf.layers.dropout({ rate: dropoutRate }));
    model.add(tf.layers.simpleRNN({ units: units2, returnSequences: false }));
    model.add(tf.layers.dense({ units: 64, activation: 'relu' }));
    model.add(tf.layers.dense({ units: 32, activation: 'relu' }));
    model.add(tf.layers.dense({ units: 1 }));
    model.compile({ optimizer: tf.train.adam(learningRate), loss: 'meanSquaredError', metrics: ['mae'] });

    // --- Train, with early stopping + "reduce LR on plateau" -----------------
    // NOTE: we deliberately do NOT use tf.callbacks.earlyStopping() combined
    // with another callback in the same array — mixing a built-in Callback
    // instance with a plain-object callback triggers a known TensorFlow.js
    // bug ("setParams is not a function", see tensorflow/tfjs#1792).
    // Instead we drive training epoch-by-epoch ourselves, which also lets us
    // faithfully replicate Keras' restore_best_weights=True behaviour.
    let bestValLoss = Infinity;
    let bestWeights = null;
    let earlyStopWait = 0;
    let plateauWait = 0;
    const PLATEAU_PATIENCE = 8;
    const PLATEAU_FACTOR = 0.5;
    const MIN_DELTA = 1e-6;

    for (let epoch = 0; epoch < epochs; epoch++) {
      const history = await model.fit(XTrain, yTrain, {
        epochs: 1,
        batchSize,
        validationData: [XVal, yVal],
        verbose: 0,
      });
      const valLoss = history.history.val_loss[0];

      const pct = Math.min(100, Math.round(((epoch + 1) / epochs) * 100));
      el.progressFill.style.width = `${pct}%`;
      el.progressLabel.textContent = `Epoch ${epoch + 1}/${epochs} — val_loss: ${valLoss.toFixed(5)}`;

      if (valLoss < bestValLoss - MIN_DELTA) {
        bestValLoss = valLoss;
        earlyStopWait = 0;
        plateauWait = 0;
        // Keep a snapshot of the best weights so far (restore_best_weights)
        if (bestWeights) bestWeights.forEach((t) => t.dispose());
        bestWeights = model.getWeights().map((w) => w.clone());
      } else {
        earlyStopWait++;
        plateauWait++;

        if (plateauWait >= PLATEAU_PATIENCE) {
          model.optimizer.learningRate = model.optimizer.learningRate * PLATEAU_FACTOR;
          plateauWait = 0;
        }
        if (earlyStopWait >= patience) {
          el.progressLabel.textContent =
            `Early stopping at epoch ${epoch + 1} (best val_loss: ${bestValLoss.toFixed(5)})`;
          break;
        }
      }

      await tf.nextFrame(); // keep the UI responsive during training
    }

    // Restore the best-performing weights, mirroring restore_best_weights=True
    if (bestWeights) {
      model.setWeights(bestWeights);
      bestWeights.forEach((t) => t.dispose());
    }

    el.progressLabel.textContent = 'Training complete. Generating predictions…';

    // --- Predict and inverse-transform back to physical units -------------
    const inverseTarget = (arr2d) => arr2d.map((row) => scalers[targetIndex].inverseTransform(row)[0]);

    const trainPredScaled = await model.predict(XTrain).array();
    const valPredScaled = await model.predict(XVal).array();
    const testPredScaled = await model.predict(XTest).array();

    const trainPred = inverseTarget(trainPredScaled);
    const valPred = inverseTarget(valPredScaled);
    const testPred = inverseTarget(testPredScaled);

    const yTrainReal = inverseTarget(yTrainArr);
    const yValReal = inverseTarget(yValArr);
    const yTestReal = inverseTarget(yTestArr);

    // --- Align each split with its dates -----------------------------------
    const dateTrain = dates.slice(seqLen, seqLen + yTrainReal.length);
    const dateVal = dates.slice(trainSize, trainSize + yValReal.length);
    const dateTest = dates.slice(trainSize + valSize, trainSize + valSize + yTestReal.length);
    const splitDates = [dates[trainSize], dates[trainSize + valSize]];

    // --- Metrics -------------------------------------------------------------
    const metrics = {
      Train: evaluateMetrics(yTrainReal, trainPred),
      Validation: evaluateMetrics(yValReal, valPred),
      Test: evaluateMetrics(yTestReal, testPred),
    };

    renderMetrics(metrics);
    renderMainPlot({ dateTrain, yTrainReal, trainPred, dateVal, yValReal, valPred, dateTest, yTestReal, testPred, splitDates }, targetKey);
    renderResidualPlot({ dateTrain, yTrainReal, trainPred, dateVal, yValReal, valPred, dateTest, yTestReal, testPred });
    renderScatterPlot({ yTrainReal, trainPred, yValReal, valPred, yTestReal, testPred });

    el.trainingProgress.hidden = true;

    // Free up GPU/CPU memory held by tensors
    [XTrain, yTrain, XVal, yVal, XTest, yTest].forEach((t) => t.dispose());
  } catch (err) {
    console.error(err);
    el.progressLabel.textContent = `Error: ${err.message}`;
  } finally {
    el.runButton.disabled = false;
  }
}

// =======================================================================
// 7. Rendering
// =======================================================================

function renderMetrics(metrics) {
  el.metricsCard.hidden = false;
  el.metricsTableBody.innerHTML = '';
  Object.entries(metrics).forEach(([split, m]) => {
    const row = document.createElement('tr');
    row.innerHTML = `
      <td>${split}</td>
      <td>${m.r2.toFixed(4)}</td>
      <td>${m.mae.toFixed(4)}</td>
      <td>${m.mse.toFixed(4)}</td>
      <td>${m.rmse.toFixed(4)}</td>`;
    el.metricsTableBody.appendChild(row);
  });
}

function renderMainPlot(d, targetKey) {
  el.mainPlotCard.hidden = false;

  const traces = [
    { x: d.dateTrain, y: d.yTrainReal, name: 'Actual (Train)', mode: 'lines', line: { color: COLOR_ACTUAL } },
    { x: d.dateTrain, y: d.trainPred, name: 'Predicted (Train)', mode: 'lines', line: { color: COLOR_PREDICTED, dash: 'dash' } },
    { x: d.dateVal, y: d.yValReal, name: 'Actual (Val)', mode: 'lines', line: { color: COLOR_ACTUAL }, showlegend: false },
    { x: d.dateVal, y: d.valPred, name: 'Predicted (Val)', mode: 'lines', line: { color: COLOR_PREDICTED, dash: 'dash' }, showlegend: false },
    { x: d.dateTest, y: d.yTestReal, name: 'Actual (Test)', mode: 'lines', line: { color: COLOR_ACTUAL }, showlegend: false },
    { x: d.dateTest, y: d.testPred, name: 'Predicted (Test)', mode: 'lines', line: { color: COLOR_PREDICTED, dash: 'dash' }, showlegend: false },
  ];

  const layout = baseLayout('Actual vs. predicted', 'Date', `${targetKey} (mm)`);
  layout.shapes = d.splitDates.map((sd) => ({
    type: 'line', x0: sd, x1: sd, y0: 0, y1: 1, yref: 'paper',
    line: { color: COLOR_SPLIT_LINE, dash: 'dot', width: 1 },
  }));

  Plotly.newPlot(el.mainPlot, traces, layout, { displayModeBar: false, responsive: true });
}

function renderResidualPlot(d) {
  el.residualPlotCard.hidden = false;

  // All three splits are shown on one time axis, colored by split — a
  // simpler alternative to three separate subplots that still makes
  // any drift or bias in the residuals easy to spot over time.
  const residual = (real, pred) => real.map((v, i) => v - pred[i]);

  const traces = [
    { x: d.dateTrain, y: residual(d.yTrainReal, d.trainPred), name: 'Train', mode: 'lines', line: { color: SPLIT_COLORS.Train } },
    { x: d.dateVal, y: residual(d.yValReal, d.valPred), name: 'Validation', mode: 'lines', line: { color: SPLIT_COLORS.Validation } },
    { x: d.dateTest, y: residual(d.yTestReal, d.testPred), name: 'Test', mode: 'lines', line: { color: SPLIT_COLORS.Test } },
  ];

  const layout = baseLayout('Residuals over time', 'Date', 'Residual (mm)');
  layout.shapes = [{ type: 'line', x0: d.dateTrain[0], x1: d.dateTest[d.dateTest.length - 1], y0: 0, y1: 0, line: { color: '#1b2b33', dash: 'dash', width: 1 } }];

  Plotly.newPlot(el.residualPlot, traces, layout, { displayModeBar: false, responsive: true });
}

function renderScatterPlot(d) {
  el.scatterPlotCard.hidden = false;

  const allActual = [...d.yTrainReal, ...d.yValReal, ...d.yTestReal];
  const allPred = [...d.trainPred, ...d.valPred, ...d.testPred];
  const lo = Math.min(...allActual, ...allPred);
  const hi = Math.max(...allActual, ...allPred);

  const traces = [
    { x: d.yTrainReal, y: d.trainPred, name: 'Train', mode: 'markers', marker: { color: SPLIT_COLORS.Train, size: 6, opacity: 0.65 } },
    { x: d.yValReal, y: d.valPred, name: 'Validation', mode: 'markers', marker: { color: SPLIT_COLORS.Validation, size: 6, opacity: 0.65 } },
    { x: d.yTestReal, y: d.testPred, name: 'Test', mode: 'markers', marker: { color: SPLIT_COLORS.Test, size: 6, opacity: 0.65 } },
    { x: [lo, hi], y: [lo, hi], name: '1:1 line', mode: 'lines', line: { color: '#1b2b33', dash: 'dash', width: 1 } },
  ];

  Plotly.newPlot(el.scatterPlot, traces, baseLayout('Actual vs. predicted', 'Actual', 'Predicted'), { displayModeBar: false, responsive: true });
}
