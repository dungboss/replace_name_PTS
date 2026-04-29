#target photoshop

'use strict';

app.preferences.rulerUnits = Units.PIXELS;
app.bringToFront();

var baseFolder = (new File($.fileName)).parent;
var outputFolder = new Folder(baseFolder.fsName + "/Result");
var debugLines = [];
var debugEnabled = true;

var exportOptions = new ExportOptionsSaveForWeb();
exportOptions.quality = 100;
exportOptions.PNG8 = false;
exportOptions.format = SaveDocumentType.PNG;

main();

function main() {
  debugLines = [];

  try {
    if (!outputFolder.exists && !outputFolder.create()) {
      alert("Result folder cannot be created");
      return;
    }

    var selection = getFixedSelection();
    if (!selection) {
      return;
    }

    var records = loadInputRecords(selection.inputFile);
    if (records.length === 0) {
      alert("Không tìm thấy dữ liệu hợp lệ trong input.csv.");
      return;
    }

    debugLog("Loaded records: " + records.length);

    sortRecordsByLength(records);

    var skippedCount = 0;
    var failedCount = 0;
    for (var recordIndex = 0; recordIndex < records.length; recordIndex++) {
      var record = records[recordIndex];
      var template = getTemplateForName(record.sourceName, selection.rules);
      if (!template) {
        skippedCount++;
        debugLog("SKIP: " + record.sourceName + " (no template match)");
        continue;
      }
      debugLog("PROCESS: " + record.sourceName + " -> " + record.outputName + " | template=" + template.fsName);
      try {
        processName(record.sourceName, record.outputName, template);
      } catch (recordError) {
        failedCount++;
        debugLog("RECORD FAILED: " + record.sourceName + " | error=" + recordError);
      }
    }

    if (skippedCount > 0 || failedCount > 0) {
      alert("Skipped: " + skippedCount + ", failed: " + failedCount + ". See Result/debug-log.txt");
    }
  } catch (e) {
    debugLog("FATAL: " + e);
    alert("Script failed. See Result/debug-log.txt");
  } finally {
    flushDebugLog();
  }
}

function getFixedSelection() {
  var inputFile = new File(baseFolder.fsName + "/input.csv");
  var templateFile = new File(baseFolder.fsName + "/PTS/HUONG004.1.psd");

  if (!inputFile.exists) {
    alert("Không tìm thấy input.csv trong cùng folder với script.");
    debugLog("Missing input file: " + inputFile.fsName);
    return null;
  }

  if (!templateFile.exists) {
    alert("Không tìm thấy template PSD: PTS/HUONG004.1.psd");
    debugLog("Missing template file: " + templateFile.fsName);
    return null;
  }

  debugLog("Fixed selection | input=" + inputFile.fsName + " | template=" + templateFile.fsName);

  return {
    inputFile: inputFile,
    rules: [
      {
        min: 1,
        max: null,
        template: templateFile
      }
    ]
  };
}

function loadInputRecords(fileObj) {
  fileObj.encoding = "UTF-8";
  if (!fileObj.open("r")) {
    alert("Không thể mở file " + fileObj.name);
    return [];
  }
  var content = fileObj.read();
  fileObj.close();
  if (content.length > 0 && content.charCodeAt(0) === 0xFEFF) {
    content = content.substring(1);
  }
  var rows = parseCsvRows(content);
  var items = [];
  var headerSkipped = false;

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (isRowEmpty(row)) {
      continue;
    }

    if (!headerSkipped) {
      headerSkipped = true;
      continue;
    }

    var outputName = row.length > 0 ? trimString(row[0]) : "";

    var sourceName = "";
    if (row.length > 1) {
      sourceName = trimString(row[1]);
    }
    if (sourceName.length === 0) {
      continue;
    }

    if (outputName.length === 0) {
      outputName = sourceName;
    }

    items.push({
      sourceName: sourceName,
      outputName: outputName
    });
  }

  return items;
}

function parseCsvRows(content) {
  var rows = [];
  var row = [];
  var field = "";
  var inQuotes = false;

  for (var i = 0; i < content.length; i++) {
    var ch = content.charAt(i);

    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < content.length && content.charAt(i + 1) === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
      continue;
    }

    if (ch === ',') {
      row.push(field);
      field = "";
      continue;
    }

    if (ch === '\r' || ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      if (ch === '\r' && i + 1 < content.length && content.charAt(i + 1) === '\n') {
        i++;
      }
      continue;
    }

    field += ch;
  }

  row.push(field);
  rows.push(row);
  return rows;
}

function isRowEmpty(row) {
  if (!row) {
    return true;
  }

  for (var i = 0; i < row.length; i++) {
    if (trimString(row[i]).length > 0) {
      return false;
    }
  }
  return true;
}

function sortRecordsByLength(records) {
  records.sort(function (a, b) {
    var aLength = getNameLength(a.sourceName);
    var bLength = getNameLength(b.sourceName);

    if (aLength < bLength) {
      return -1;
    }
    if (aLength > bLength) {
      return 1;
    }

    var aName = String(a.sourceName).toLowerCase();
    var bName = String(b.sourceName).toLowerCase();
    if (aName < bName) {
      return -1;
    }
    if (aName > bName) {
      return 1;
    }

    var aOutputName = String(a.outputName).toLowerCase();
    var bOutputName = String(b.outputName).toLowerCase();
    if (aOutputName < bOutputName) {
      return -1;
    }
    if (aOutputName > bOutputName) {
      return 1;
    }

    return 0;
  });
}

function showSelectionDialog() {
  var dialog = new Window("dialog", "Select input CSV and template folder");
  dialog.orientation = "column";
  dialog.alignChildren = "fill";
  dialog.spacing = 10;
  dialog.margins = 16;

  var message = dialog.add("statictext", undefined, "Choose the input CSV, the template folder, and any number of length rules.");
  message.maximumSize.width = 420;

  var inputRow = addPathPickerRow(dialog, "input.csv", "csv");
  var templateFolderRow = addPathPickerRow(dialog, "Template folder", "folder");

  var rulePanel = dialog.add("panel", undefined, "Length rules");
  rulePanel.orientation = "column";
  rulePanel.alignChildren = "fill";
  rulePanel.spacing = 8;
  rulePanel.margins = 10;

  var ruleNote = rulePanel.add("statictext", undefined, "Use Min/Max. Leave Max blank for open-ended.");
  ruleNote.maximumSize.width = 420;

  var ruleList = rulePanel.add("group");
  ruleList.orientation = "column";
  ruleList.alignChildren = "fill";
  ruleList.spacing = 6;

  var ruleRows = [];
  var currentTemplateFiles = [];

  ruleRows.push(addRuleRow(ruleList, ruleRows, { min: 1, max: 3 }, currentTemplateFiles));
  ruleRows.push(addRuleRow(ruleList, ruleRows, { min: 4, max: 6 }, currentTemplateFiles));
  ruleRows.push(addRuleRow(ruleList, ruleRows, { min: 7, max: 10 }, currentTemplateFiles));
  ruleRows.push(addRuleRow(ruleList, ruleRows, { min: 11, max: null }, currentTemplateFiles));

  var ruleButtonRow = rulePanel.add("group");
  ruleButtonRow.alignment = "right";
  var addRuleButton = ruleButtonRow.add("button", undefined, "Add rule");

  addRuleButton.onClick = function () {
    ruleRows.push(addRuleRow(ruleList, ruleRows, null, currentTemplateFiles));
    dialog.layout.layout(true);
  };

  templateFolderRow.browseButton.onClick = function () {
    var folder = Folder.selectDialog("Select the template folder");
    if (!folder) {
      return;
    }
    templateFolderRow.value = folder;
    templateFolderRow.pathField.text = folder.fsName;

    currentTemplateFiles = loadTemplateFiles(folder);
    if (currentTemplateFiles.length === 0) {
      alert("No PSD files were found in the selected folder.");
      templateFolderRow.value = null;
      templateFolderRow.pathField.text = "";
      populateRuleRowsTemplateFiles(ruleRows, currentTemplateFiles);
      dialog.layout.layout(true);
      return;
    }
    populateRuleRowsTemplateFiles(ruleRows, currentTemplateFiles);
    dialog.layout.layout(true);
  };

  var buttonGroup = dialog.add("group");
  buttonGroup.alignment = "right";
  var okButton = buttonGroup.add("button", undefined, "OK", { name: "ok" });
  var cancelButton = buttonGroup.add("button", undefined, "Cancel", { name: "cancel" });

  okButton.onClick = function () {
    dialog.close(1);
  };
  cancelButton.onClick = function () {
    dialog.close(0);
  };

  if (dialog.show() != 1) {
    return null;
  }

  if (!inputRow.value || !templateFolderRow.value) {
    alert("Vui lòng chọn file input.csv và thư mục template.");
    return null;
  }

  if (!inputRow.value.exists) {
    alert("Không tìm thấy file input.csv.");
    return null;
  }

  if (!isCsvFile(inputRow.value)) {
    alert("File input phải là định dạng .csv.");
    return null;
  }

  if (!templateFolderRow.value.exists) {
    alert("Template folder is not found.");
    return null;
  }

  var rules = buildRulesFromRows(ruleRows);
  if (!rules) {
    return null;
  }

  return {
    inputFile: inputRow.value,
    rules: rules
  };
}

function addPathPickerRow(parent, label, kind) {
  var row = parent.add("group");
  row.orientation = "row";
  row.alignChildren = ["left", "center"];
  row.spacing = 10;

  row.add("statictext", undefined, label);

  var pathField = row.add("edittext", undefined, "");
  pathField.preferredSize.width = 300;
  pathField.enabled = false;

  var browseButton = row.add("button", undefined, "Browse...");
  row.pathField = pathField;
  row.browseButton = browseButton;
  row.value = null;

  browseButton.onClick = function () {
    if (kind === "folder") {
      var folder = Folder.selectDialog("Select " + label);
      if (!folder) {
        return;
      }
      row.value = folder;
      pathField.text = folder.fsName;
      return;
    }

    var file = File.openDialog("Select " + label);
    if (!file) {
      return;
    }
    if (kind === "csv" && !isCsvFile(file)) {
      row.value = null;
      pathField.text = "";
      alert("Vui lòng chọn file input có định dạng .csv.");
      return;
    }
    row.value = file;
    pathField.text = file.fsName;
  };

  return row;
}

function addRuleRow(parent, ruleRows, defaults, templateFiles) {
  var row = parent.add("group");
  row.orientation = "row";
  row.alignChildren = ["left", "center"];
  row.spacing = 10;

  row.add("statictext", undefined, "Min");

  row.minField = row.add("edittext", undefined, defaults && defaults.min !== null && typeof defaults.min !== "undefined" ? String(defaults.min) : "");
  row.minField.preferredSize.width = 36;

  row.add("statictext", undefined, "Max");

  row.maxField = row.add("edittext", undefined, defaults && defaults.max !== null && typeof defaults.max !== "undefined" ? String(defaults.max) : "");
  row.maxField.preferredSize.width = 36;

  row.templateDropdown = row.add("dropdownlist", undefined, []);
  row.templateDropdown.preferredSize.width = 300;

  row.removeButton = row.add("button", undefined, "Remove");
  row.templateFiles = [];

  populateRuleRowTemplates(row, templateFiles || []);

  row.removeButton.onClick = function () {
    if (ruleRows.length <= 1) {
      alert("Keep at least one length rule.");
      return;
    }

    var rowIndex = -1;
    for (var i = 0; i < ruleRows.length; i++) {
      if (ruleRows[i] === row) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      return;
    }

    ruleRows.splice(rowIndex, 1);
    parent.remove(row);
    parent.layout.layout(true);
  };

  return row;
}

function populateRuleRowsTemplateFiles(ruleRows, templateFiles) {
  for (var i = 0; i < ruleRows.length; i++) {
    populateRuleRowTemplates(ruleRows[i], templateFiles);
  }
}

function populateRuleRowTemplates(row, templateFiles) {
  var previousTemplateName = getSelectedTemplateName(row);
  var dropdown = row.templateDropdown;

  while (dropdown.items.length > 0) {
    dropdown.remove(0);
  }

  row.templateFiles = templateFiles || [];

  if (!templateFiles || templateFiles.length === 0) {
    dropdown.add("item", "Select template folder first");
    dropdown.selection = 0;
    dropdown.enabled = false;
    return null;
  }

  dropdown.add("item", "Select template");
  for (var i = 0; i < templateFiles.length; i++) {
    dropdown.add("item", templateFiles[i].name);
  }
  dropdown.enabled = true;

  var selectedIndex = 0;
  if (previousTemplateName) {
    for (var j = 0; j < templateFiles.length; j++) {
      if (templateFiles[j].name === previousTemplateName) {
        selectedIndex = j + 1;
        break;
      }
    }
  }

  dropdown.selection = selectedIndex;
}

function getSelectedTemplateName(row) {
  if (!row.templateDropdown.selection || row.templateDropdown.selection.index === 0) {
    return null;
  }
  return row.templateDropdown.selection.text;
}

function getSelectedTemplateFile(row) {
  if (!row.templateDropdown.selection || row.templateDropdown.selection.index === 0) {
    return null;
  }
  var index = row.templateDropdown.selection.index - 1;
  if (!row.templateFiles || index < 0 || index >= row.templateFiles.length) {
    return null;
  }
  return row.templateFiles[index];
}

function buildRulesFromRows(ruleRows) {
  var rules = [];
  for (var i = 0; i < ruleRows.length; i++) {
    var row = ruleRows[i];
    var minValue = parseLengthValue(row.minField.text, false);
    var maxValue = parseLengthValue(row.maxField.text, true);
    var templateFile = getSelectedTemplateFile(row);

    if (isNaN(minValue)) {
      alert("Invalid Min value in rule " + (i + 1) + ".");
      return null;
    }

    if (maxValue !== null && isNaN(maxValue)) {
      alert("Invalid Max value in rule " + (i + 1) + ". Leave it blank for open-ended.");
      return null;
    }

    if (maxValue !== null && maxValue < minValue) {
      alert("Max must be greater than or equal to Min in rule " + (i + 1) + ".");
      return null;
    }

    if (!templateFile) {
      alert("Please choose a template in rule " + (i + 1) + ".");
      return null;
    }

    if (!isPsdFile(templateFile)) {
      alert("Template in rule " + (i + 1) + " must be a .psd file.");
      return null;
    }

    rules.push({
      min: minValue,
      max: maxValue,
      template: templateFile
    });
  }

  sortRulesByRange(rules);
  if (!validateRules(rules)) {
    return null;
  }
  return rules;
}

function parseLengthValue(text, allowBlank) {
  text = trimString(text);
  if (text.length === 0) {
    return allowBlank ? null : NaN;
  }
  var normalized = text.replace(/\s+/g, "");

  if (allowBlank) {
    if (/^\+$/.test(normalized) || /^\d+\+$/.test(normalized)) {
      return null;
    }
    if (/^\d+$/.test(normalized)) {
      return parseInt(normalized, 10);
    }
    return NaN;
  }

  if (/^\d+$/.test(normalized)) {
    return parseInt(normalized, 10);
  }
  return NaN;
}

function trimString(value) {
  return String(value).replace(/^\s+|\s+$/g, "");
}

function isCsvFile(fileObj) {
  return !!fileObj && /\.csv$/i.test(fileObj.name);
}

function isPsdFile(fileObj) {
  return !!fileObj && /\.psd$/i.test(fileObj.name);
}

function sortRulesByRange(rules) {
  rules.sort(function (a, b) {
    if (a.min < b.min) {
      return -1;
    }
    if (a.min > b.min) {
      return 1;
    }

    var aMax = a.max === null ? 999999999 : a.max;
    var bMax = b.max === null ? 999999999 : b.max;
    if (aMax < bMax) {
      return -1;
    }
    if (aMax > bMax) {
      return 1;
    }
    return 0;
  });
}

function validateRules(rules) {
  for (var i = 0; i < rules.length - 1; i++) {
    var currentRule = rules[i];
    var nextRule = rules[i + 1];

    if (currentRule.max === null) {
      alert("Open-ended rules must be the last rule.");
      return false;
    }

    if (nextRule.min <= currentRule.max) {
      alert("Length rules overlap between rule " + (i + 1) + " and rule " + (i + 2) + ".");
      return false;
    }
  }
  return true;
}

function loadTemplateFiles(folderObj) {
  var files = folderObj.getFiles(/\.(psd)$/i);
  sortFilesByName(files);
  return files;
}

function sortFilesByName(files) {
  files.sort(function (a, b) {
    var an = a.name.toLowerCase();
    var bn = b.name.toLowerCase();
    if (an < bn) {
      return -1;
    }
    if (an > bn) {
      return 1;
    }
    return 0;
  });
}

function getTemplateForName(name, rules) {
  var length = getNameLength(name);
  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    if (length < rule.min) {
      continue;
    }
    if (rule.max !== null && length > rule.max) {
      continue;
    }
    return rule.template;
  }
  return null;
}

function getNameLength(name) {
  name = String(name);
  name = name.replace(/^\s+|\s+$/g, "");
  return name.length;
}

function processName(sourceName, outputName, template) {
  var doc = null;

  try {
    open(template);
    doc = app.activeDocument;

    var nameLayers = findLayers(doc, true, {
      typename: "ArtLayer",
      kind: LayerKind.TEXT,
      name: "name"
    });

    if (nameLayers.length === 0) {
      debugLog("No layer named \"name\" was found in template: " + template.fsName);
      throw new Error("name layer is not found");
    }

    var nameLayer = nameLayers[0];

    changeTextLayerContent(nameLayer, String(sourceName));

    var outputFileName = buildOutputFileName(outputName);
    app.activeDocument.exportDocument(
      new File(outputFolder.fsName + "/" + outputFileName),
      ExportType.SAVEFORWEB,
      exportOptions
    );
  } catch (e) {
    debugLog("processName failed | sourceName=" + sourceName + " | error=" + e);
    throw e;
  } finally {
    try {
      doc.close(SaveOptions.DONOTSAVECHANGES);
    } catch (closeError) {
      debugLog("Could not close document: " + closeError);
    }
  }
}

function changeTextLayerContent(textLayer, textContent) {
  if (!textLayer) {
    return;
  }

  app.activeDocument.activeLayer = textLayer;
  var originalBounds = getLayerBoundsPx(textLayer);
  var originalText = "";
  try {
    originalText = String(textLayer.textItem.contents);
  } catch (e) {
    originalText = "";
  }
  textContent = String(textContent).replace(/\n/g, " ");
  var isParagraphText = isParagraphTextLayer(textLayer);
  var targetBox = getTextBoxSizePx(textLayer, originalBounds);
  var originalSizePt = getTextSizeInPoints(textLayer);
  var fitBox = estimateAdaptiveFitBox(targetBox, originalText, textContent, isParagraphText);

  debugLog(
    "Layer debug | kind=" + (isParagraphText ? "paragraph" : "point") +
    " | originalText=\"" + originalText + "\"" +
    " | newText=\"" + textContent + "\"" +
    " | originalBounds=" + formatBox(originalBounds) +
    " | targetBox=" + formatBox(targetBox) +
    " | fitBox=" + formatBox(fitBox) +
    " | originalSizePt=" + originalSizePt
  );
  debugLog(
    "Layer frame | itemKind=" + safeTextItemProp(textLayer, "kind") +
    " | isParagraphText=" + safeTextItemProp(textLayer, "isParagraphText") +
    " | frameW=" + safeTextItemPx(textLayer, "width") +
    " | frameH=" + safeTextItemPx(textLayer, "height")
  );

  if (isParagraphText) {
    setTextLayerBoxSize(textLayer, fitBox.width, Math.max(targetBox.height, originalBounds.height));
    debugLog(
      "Applied frame | frameW=" + safeTextItemPx(textLayer, "width") +
      " | frameH=" + safeTextItemPx(textLayer, "height")
    );
  }

  textLayer.textItem.contents = textContent;
  var preFitBounds = getLayerBoundsPx(textLayer);
  debugLog("Pre-fit bounds | bounds=" + formatBox(preFitBounds));
  var centerJustification = getCenterJustification();
  if (centerJustification !== null) {
    setTextLayerJustification(textLayer, centerJustification);
  }

  if (fitBox.width > 0 && originalSizePt > 0) {
    fitTextLayerToWidth(textLayer, fitBox.width, originalSizePt);
  }

  var currentBounds = getLayerBoundsPx(textLayer);
  var targetCenterX = originalBounds.left + originalBounds.width / 2;
  var targetCenterY = originalBounds.top + originalBounds.height / 2;
  if (isParagraphText) {
    var targetAnchor = getTextAnchorPx(textLayer, originalBounds);
    targetCenterX = targetAnchor.left + targetBox.width / 2;
    targetCenterY = targetAnchor.top + targetBox.height / 2;
  }
  var currentCenterX = currentBounds.left + currentBounds.width / 2;
  var currentCenterY = currentBounds.top + currentBounds.height / 2;
  var deltaX = targetCenterX - currentCenterX;
  var deltaY = targetCenterY - currentCenterY;
  textLayer.translate(deltaX, deltaY);

  var finalBounds = getLayerBoundsPx(textLayer);
  debugLog("Layer result | currentBounds=" + formatBox(finalBounds) + " | finalSizePt=" + getTextSizeInPoints(textLayer));
}

function estimateAdaptiveFitBox(referenceBox, originalText, replacementText, isParagraphText) {
  var fitBox = {
    width: referenceBox.width,
    height: referenceBox.height
  };

  var originalLength = getNameLength(originalText);
  var replacementLength = getNameLength(replacementText);

  if (referenceBox.width > 0) {
    var widthScale = 1;
    if (originalLength > 0 && replacementLength > originalLength) {
      widthScale = replacementLength / originalLength;
    }
    if (isParagraphText) {
      widthScale = Math.max(widthScale * 1.75, 2.5);
    } else {
      widthScale = Math.max(widthScale * 1.25, 1.25);
    }
    fitBox.width = referenceBox.width * widthScale;
  }

  return fitBox;
}

function fitTextLayerToWidth(textLayer, targetWidth, originalSizePt) {
  if (textLayerFitsWidth(textLayer, targetWidth)) {
    return;
  }

  var low = 1;
  var high = originalSizePt;
  var best = low;

  for (var i = 0; i < 12; i++) {
    var mid = (low + high) / 2;
    setTextLayerSizePt(textLayer, mid);
    if (textLayerFitsWidth(textLayer, targetWidth)) {
      best = mid;
      low = mid;
    } else {
      high = mid;
    }
  }

  setTextLayerSizePt(textLayer, best);
}

function textLayerFitsWidth(textLayer, targetWidth) {
  var bounds = getLayerBoundsPx(textLayer);
  var tolerance = 1;
  return bounds.width <= targetWidth + tolerance;
}

function setTextLayerBoxSize(textLayer, widthPx, heightPx) {
  try {
    if (widthPx > 0) {
      textLayer.textItem.width = new UnitValue(widthPx, "px");
    }
    if (heightPx > 0) {
      textLayer.textItem.height = new UnitValue(heightPx, "px");
    }
  } catch (e) {
    debugLog("Could not set text box size: " + e);
  }
}

function setTextLayerSizePt(textLayer, sizePt) {
  textLayer.textItem.size = new UnitValue(sizePt, "pt");
}

function setTextLayerJustification(textLayer, justification) {
  try {
    textLayer.textItem.justification = justification;
  } catch (e) {
    // Some text layers or Photoshop versions may not expose justification.
  }
}

function getCenterJustification() {
  try {
    if (typeof Justification !== "undefined" && typeof Justification.CENTER !== "undefined") {
      return Justification.CENTER;
    }
  } catch (e) {
    // Ignore and fall back to no explicit justification.
  }
  return null;
}

function getTextSizeInPoints(textLayer) {
  var size = textLayer.textItem.size;
  if (size && typeof size.as === "function") {
    return size.as("pt");
  }
  if (size && typeof size.value !== "undefined") {
    return size.value;
  }
  return Number(size);
}

function getTextBoxSizePx(textLayer, fallbackBounds) {
  var item = textLayer.textItem;
  var width = 0;
  var height = 0;

  try {
    if (item.width && item.height) {
      width = unitValueToPx(item.width);
      height = unitValueToPx(item.height);
    }
  } catch (e) {
    width = 0;
    height = 0;
  }

  if (width > 0 && height > 0) {
    return {
      width: width,
      height: height
    };
  }

  return {
    width: fallbackBounds.width,
    height: fallbackBounds.height
  };
}

function getTextAnchorPx(textLayer, fallbackBounds) {
  var item = textLayer.textItem;
  try {
    if (isParagraphTextLayer(textLayer) && item.position && item.position.length >= 2) {
      return {
        left: unitValueToPx(item.position[0]),
        top: unitValueToPx(item.position[1])
      };
    }
  } catch (e) {
    // Fall back to the visible bounds when the position is unavailable.
  }

  return {
    left: fallbackBounds.left,
    top: fallbackBounds.top
  };
}

function isParagraphTextLayer(textLayer) {
  try {
    if (typeof textLayer.textItem.isParagraphText !== "undefined") {
      return textLayer.textItem.isParagraphText === true;
    }
  } catch (e) {
    // Ignore and fall through to the older kind check.
  }

  try {
    if (typeof TextType !== "undefined" && typeof textLayer.textItem.kind !== "undefined") {
      return textLayer.textItem.kind === TextType.PARAGRAPHTEXT;
    }
  } catch (e2) {
    // Ignore and use the fallback bounds.
  }

  return false;
}

function getLayerBoundsPx(layer) {
  var bounds = layer.bounds;
  var left = unitValueToPx(bounds[0]);
  var top = unitValueToPx(bounds[1]);
  var right = unitValueToPx(bounds[2]);
  var bottom = unitValueToPx(bounds[3]);

  return {
    left: left,
    top: top,
    right: right,
    bottom: bottom,
    width: right - left,
    height: bottom - top
  };
}

function unitValueToPx(value) {
  if (value && typeof value.as === "function") {
    return value.as("px");
  }
  if (value && typeof value.value !== "undefined") {
    return value.value;
  }
  return Number(value);
}

function sanitizeFileName(name) {
  return name.replace(/[\\\/:\*\?"<>\|]/g, "_");
}

function buildOutputFileName(outputName) {
  var fileName = sanitizeFileName(trimString(outputName)).replace(/[\. ]+$/g, "");
  if (fileName.length === 0) {
    fileName = "output";
  }
  if (!/\.png$/i.test(fileName)) {
    fileName += ".png";
  }
  return fileName;
}

function findLayers(searchFolder, recursion, userData, items) {
  items = items || [];
  var folderItem;
  for (var i = 0; i < searchFolder.layers.length; i++) {
    folderItem = searchFolder.layers[i];
    if (propertiesMatch(folderItem, userData)) {
      items.push(folderItem);
    }
    if (recursion === true && folderItem.typename === "LayerSet") {
      findLayers(folderItem, recursion, userData, items);
    }
  }
  return items;
}

function propertiesMatch(projectItem, userData) {
  if (typeof userData === "undefined") return true;
  for (var propertyName in userData) {
    if (!userData.hasOwnProperty(propertyName)) continue;
    if (!projectItem.hasOwnProperty(propertyName)) return false;
    if (projectItem[propertyName].toString() !== userData[propertyName].toString()) {
      return false;
    }
  }
  return true;
}

function formatBox(box) {
  if (!box) {
    return "{null}";
  }
  return "{l:" + roundNumber(box.left) + ", t:" + roundNumber(box.top) + ", w:" + roundNumber(box.width) + ", h:" + roundNumber(box.height) + "}";
}

function roundNumber(value) {
  if (value === null || typeof value === "undefined" || isNaN(value)) {
    return "na";
  }
  return Math.round(Number(value) * 100) / 100;
}

function safeTextItemProp(textLayer, propName) {
  try {
    var value = textLayer.textItem[propName];
    if (value === null || typeof value === "undefined") {
      return "na";
    }
    if (propName === "kind") {
      return describeTextKind(value);
    }
    return String(value);
  } catch (e) {
    return "na";
  }
}

function safeTextItemPx(textLayer, propName) {
  try {
    var value = textLayer.textItem[propName];
    var px = unitValueToPx(value);
    return roundNumber(px);
  } catch (e) {
    return "na";
  }
}

function describeTextKind(kindValue) {
  try {
    if (typeof TextType !== "undefined") {
      if (kindValue === TextType.PARAGRAPHTEXT) {
        return "PARAGRAPHTEXT";
      }
      if (kindValue === TextType.POINTTEXT) {
        return "POINTTEXT";
      }
    }
  } catch (e) {
    // Fall through to stringification below.
  }

  try {
    return String(kindValue);
  } catch (e2) {
    return "na";
  }
}

function debugLog(message) {
  if (!debugEnabled) {
    return;
  }
  debugLines.push(String(message));
}

function flushDebugLog() {
  if (!debugEnabled) {
    return;
  }

  try {
    var logFile = new File(baseFolder.fsName + "/Result/debug-log.txt");
    if (!logFile.open("w")) {
      return;
    }
    logFile.encoding = "UTF-8";
    logFile.write(debugLines.join("\r\n"));
    logFile.close();
  } catch (e) {
    // Best-effort debug output only.
  }
}
