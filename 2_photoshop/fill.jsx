#target photoshop

/*
 * KOTRA newsletter PSD automation.
 * Run with Photoshop: File > Scripts > Browse...
 *
 * Expected CSV columns:
 * section,slot,title,period,deadline,url,image_path,newsletter_date
 */

(function () {
    var oldRulerUnits = app.preferences.rulerUnits;
    app.preferences.rulerUnits = Units.PIXELS;
    var autoConfig = typeof NEWSLETTER_AUTOMATION_CONFIG !== "undefined" ? NEWSLETTER_AUTOMATION_CONFIG : null;
    var automationError = null;
    var oldDisplayDialogs = app.displayDialogs;
    if (autoConfig) {
        app.displayDialogs = DialogModes.NO;
    }

    /*
     * PSD PNG 내보내기 대상 이름을 바꾸는 곳.
     * 실제 PSD의 레이어/그룹 이름이 다르면 아래 별칭 배열만 수정하세요.
     * 배너 2는 BANNER_2_ALIASES를 "배너" 루트 그룹 안에서 찾습니다.
     */
    var PNG_EXPORT_LAYER_NAMES = {
        BG_ALIASES: ["bg"],
        BANNER_2_ALIASES: ["배너2", "배너02", "2", "02", "2번", "02번"],
        RIGHT_BANNER_ALIASES: ["우측배너", "우측 배너"]
    };

    /*
     * 배너 1·3의 둥근 모서리를 만드는 실제 마스크 레이어 이름.
     * PSD에서 이름이 바뀌면 아래 두 배열만 수정하세요.
     */
    var BANNER_MASK_LAYER_NAMES = {
        "1": ["03 사본"],
        "3": ["03 사본 3"]
    };

    function trim(value) {
        return String(value === undefined || value === null ? "" : value).replace(/^\s+|\s+$/g, "");
    }

    function parseCSV(text) {
        text = text.replace(/^\uFEFF/, "");
        var rows = [];
        var row = [];
        var field = "";
        var quoted = false;
        var i;
        for (i = 0; i < text.length; i++) {
            var ch = text.charAt(i);
            if (quoted) {
                if (ch === '"') {
                    if (i + 1 < text.length && text.charAt(i + 1) === '"') {
                        field += '"';
                        i++;
                    } else {
                        quoted = false;
                    }
                } else {
                    field += ch;
                }
            } else if (ch === '"') {
                quoted = true;
            } else if (ch === ',') {
                row.push(field);
                field = "";
            } else if (ch === '\r' || ch === '\n') {
                if (ch === '\r' && i + 1 < text.length && text.charAt(i + 1) === '\n') {
                    i++;
                }
                row.push(field);
                rows.push(row);
                row = [];
                field = "";
            } else {
                field += ch;
            }
        }
        if (field.length || row.length) {
            row.push(field);
            rows.push(row);
        }
        if (!rows.length) {
            throw new Error("CSV가 비어 있습니다.");
        }
        var headers = rows.shift();
        var result = [];
        for (i = 0; i < rows.length; i++) {
            if (rows[i].length === 1 && trim(rows[i][0]) === "") {
                continue;
            }
            var item = {};
            var j;
            for (j = 0; j < headers.length; j++) {
                item[trim(headers[j])] = rows[i][j] === undefined ? "" : rows[i][j];
            }
            result.push(item);
        }
        return result;
    }

    function readCSV(file) {
        file.encoding = "UTF8";
        if (!file.open("r")) {
            throw new Error("CSV를 열 수 없습니다: " + file.fsName);
        }
        var text = file.read();
        file.close();
        return parseCSV(text);
    }

    function px(value) {
        return value.as("px");
    }

    function boundsOf(layer) {
        var bounds = layer.bounds;
        return {
            left: px(bounds[0]),
            top: px(bounds[1]),
            right: px(bounds[2]),
            bottom: px(bounds[3]),
            width: px(bounds[2]) - px(bounds[0]),
            height: px(bounds[3]) - px(bounds[1])
        };
    }

    function normalizeGroupName(value) {
        return trim(value).toLowerCase().replace(/[\s_\-~\u2013\u2014·]+/g, "");
    }

    function nameMatches(name, aliases) {
        var normalized = normalizeGroupName(name);
        var i;
        for (i = 0; i < aliases.length; i++) {
            if (normalized === normalizeGroupName(aliases[i])) {
                return true;
            }
        }
        return false;
    }

    function directGroupAliases(parent, aliases) {
        var i;
        for (i = 0; i < parent.layerSets.length; i++) {
            if (nameMatches(parent.layerSets[i].name, aliases)) {
                return parent.layerSets[i];
            }
        }
        return null;
    }

    function findGroupAliases(parent, aliases, recursive) {
        var group = directGroupAliases(parent, aliases);
        if (group || !recursive) {
            return group;
        }
        var i;
        for (i = 0; i < parent.layerSets.length; i++) {
            group = findGroupAliases(parent.layerSets[i], aliases, true);
            if (group) {
                return group;
            }
        }
        return null;
    }

    function requiredGroupAliases(parent, aliases, label) {
        var group = findGroupAliases(parent, aliases, true);
        if (!group) {
            throw new Error("레이어 그룹을 찾을 수 없습니다: " + label + " (허용 이름: " + aliases.join(", ") + ")");
        }
        return group;
    }

    function directLayerAliases(parent, aliases) {
        var i;
        for (i = 0; i < parent.layerSets.length; i++) {
            if (nameMatches(parent.layerSets[i].name, aliases)) {
                return parent.layerSets[i];
            }
        }
        for (i = 0; i < parent.artLayers.length; i++) {
            if (nameMatches(parent.artLayers[i].name, aliases)) {
                return parent.artLayers[i];
            }
        }
        return null;
    }

    function findLayerAliases(parent, aliases, recursive) {
        var layer = directLayerAliases(parent, aliases);
        if (layer || !recursive) {
            return layer;
        }
        var i;
        for (i = 0; i < parent.layerSets.length; i++) {
            layer = findLayerAliases(parent.layerSets[i], aliases, true);
            if (layer) {
                return layer;
            }
        }
        return null;
    }

    function requiredLayerAliases(parent, aliases, label) {
        var layer = findLayerAliases(parent, aliases, true);
        if (!layer) {
            throw new Error("PNG 내보내기 대상 레이어/그룹을 찾을 수 없습니다: " + label + " (허용 이름: " + aliases.join(", ") + ")");
        }
        return layer;
    }

    function groupTree(parent, prefix, depth) {
        var lines = [];
        if (depth > 12) {
            lines.push(prefix + "[depth limit]");
            return lines.join("\n");
        }
        var i;
        for (i = 0; i < parent.layerSets.length; i++) {
            var group = parent.layerSets[i];
            lines.push(prefix + "- " + group.name);
            var children = groupTree(group, prefix + "  ", depth + 1);
            if (children) {
                lines.push(children);
            }
        }
        return lines.join("\n");
    }

    function writeAutomationLog(path, text) {
        if (!path) {
            return;
        }
        var file = new File(path);
        file.encoding = "UTF8";
        if (file.open("w")) {
            file.write(text);
            file.close();
        }
    }

    function ensureFolder(path) {
        var folder = path instanceof Folder ? path : new Folder(path);
        if (!folder.exists && !folder.create()) {
            throw new Error("PNG 저장 폴더를 만들 수 없습니다: " + folder.fsName);
        }
        return folder;
    }

    function exportLayerPNG(sourceDocument, sourceLayer, outputFolder, fileName) {
        var sourceBounds = boundsOf(sourceLayer);
        var width = Math.max(1, Math.ceil(sourceBounds.width));
        var height = Math.max(1, Math.ceil(sourceBounds.height));
        if (sourceBounds.width <= 0 || sourceBounds.height <= 0) {
            throw new Error(sourceLayer.name + " 레이어/그룹의 크기가 0이라 PNG로 저장할 수 없습니다.");
        }

        var exportDocument = app.documents.add(
            width,
            height,
            sourceDocument.resolution,
            fileName.replace(/\.png$/i, ""),
            NewDocumentMode.RGB,
            DocumentFill.TRANSPARENT
        );
        try {
            app.activeDocument = sourceDocument;
            var copiedLayer = sourceLayer.duplicate(exportDocument, ElementPlacement.PLACEATBEGINNING);
            app.activeDocument = exportDocument;
            copiedLayer.visible = true;
            var copiedBounds = boundsOf(copiedLayer);
            copiedLayer.translate(-copiedBounds.left, -copiedBounds.top);
            exportDocument.trim(TrimType.TRANSPARENT, true, true, true, true);

            var outputFile = new File(outputFolder.fsName + "/" + fileName);
            var options = new PNGSaveOptions();
            options.interlaced = false;
            exportDocument.saveAs(outputFile, options, true, Extension.LOWERCASE);
            return outputFile;
        } finally {
            try {
                exportDocument.close(SaveOptions.DONOTSAVECHANGES);
            } catch (ignore) {}
            app.activeDocument = sourceDocument;
        }
    }

    function exportDocumentForWeb(document, outputFile) {
        var options = new ExportOptionsSaveForWeb();
        options.format = SaveDocumentType.PNG;
        options.PNG8 = false;
        options.transparency = false;
        options.interlaced = false;
        options.includeProfile = false;
        options.optimized = true;
        document.exportDocument(outputFile, ExportType.SAVEFORWEB, options);
    }

    function exportRequiredPNGs(document, bannerRoot, outputFolder, warnings) {
        var folder = ensureFolder(outputFolder);
        var marketNewsRoot = findGroupAliases(document, ["해외시장뉴스"], true);
        var targets = [
            {
                label: "bg",
                fileName: "bg.png",
                outputFolder: folder,
                // 전체 배경인 "BG"가 아니라 해외시장뉴스 안의 정확한 "bg"를 찾습니다.
                parent: marketNewsRoot,
                aliases: PNG_EXPORT_LAYER_NAMES.BG_ALIASES
            },
            {
                label: "배너2",
                fileName: "banner_02.png",
                outputFolder: folder,
                parent: bannerRoot,
                aliases: PNG_EXPORT_LAYER_NAMES.BANNER_2_ALIASES
            },
            {
                label: "우측배너",
                fileName: "우측배너.png",
                outputFolder: folder,
                parent: document,
                aliases: PNG_EXPORT_LAYER_NAMES.RIGHT_BANNER_ALIASES
            }
        ];
        var results = [];
        var i;
        for (i = 0; i < targets.length; i++) {
            try {
                if (!targets[i].parent) {
                    throw new Error(targets[i].label + "의 상위 그룹을 찾지 못했습니다.");
                }
                var layer = findLayerAliases(targets[i].parent, targets[i].aliases, true);
                if (!layer) {
                    throw new Error(targets[i].label + " 레이어/그룹을 찾지 못했습니다.");
                }
                results.push(exportLayerPNG(document, layer, targets[i].outputFolder, targets[i].fileName));
            } catch (error) {
                warnings.push("PNG " + targets[i].label + ": " + error.message);
            }
        }
        return results;
    }

    function textLayers(group) {
        var found = [];
        var i;
        for (i = 0; i < group.artLayers.length; i++) {
            if (group.artLayers[i].kind === LayerKind.TEXT) {
                found.push(group.artLayers[i]);
            }
        }
        for (i = 0; i < group.layerSets.length; i++) {
            found = found.concat(textLayers(group.layerSets[i]));
        }
        return found;
    }

    function findTextLayer(group, prefix, titleMode) {
        var layers = textLayers(group);
        var i;
        for (i = 0; i < layers.length; i++) {
            var contents = trim(layers[i].textItem.contents).replace(/\r/g, " ");
            if (titleMode) {
                if (contents && contents.indexOf("사업기간") !== 0 && contents.indexOf("교육기간") !== 0 && contents.indexOf("신청마감") !== 0) {
                    return layers[i];
                }
            } else if (contents.indexOf(prefix) === 0) {
                return layers[i];
            }
        }
        return null;
    }

    function renameFromContents(layer) {
        var value = layer.textItem.contents.replace(/\r/g, " ");
        layer.name = trim(value).substring(0, 80);
    }

    function layerWidth(layer) {
        return boundsOf(layer).width;
    }

    function fitFont(layer, maxWidth, maxHeight, minimumPoints) {
        var original = layer.textItem.size.as("pt");
        var current = original;
        var guard = 0;
        while (guard < 80) {
            var b = boundsOf(layer);
            if (b.width <= maxWidth && b.height <= maxHeight) {
                break;
            }
            current -= 0.25;
            if (current < minimumPoints) {
                current = minimumPoints;
                layer.textItem.size = UnitValue(current, "pt");
                break;
            }
            layer.textItem.size = UnitValue(current, "pt");
            guard++;
        }
    }

    function wrapTitle(layer, text, maxWidth, maxLines) {
        text = trim(text).replace(/\s+/g, " ");
        layer.textItem.contents = text;
        if (maxLines <= 1 || layerWidth(layer) <= maxWidth) {
            return;
        }
        var words = text.split(" ");
        var lines = [];
        var current = "";
        var i;
        for (i = 0; i < words.length; i++) {
            var candidate = current ? current + " " + words[i] : words[i];
            layer.textItem.contents = candidate;
            if (current && layerWidth(layer) > maxWidth && lines.length < maxLines - 1) {
                lines.push(current);
                current = words[i];
            } else {
                current = candidate;
            }
        }
        if (current) {
            lines.push(current);
        }
        layer.textItem.contents = lines.join("\r");
    }

    function updateSlot(group, record, periodPrefix, maxTitleLines) {
        var titleLayer = findTextLayer(group, "", true);
        var periodLayer = findTextLayer(group, periodPrefix, false);
        var deadlineLayer = findTextLayer(group, "신청마감", false);
        if (!titleLayer || !periodLayer || !deadlineLayer) {
            throw new Error(group.name + " 그룹의 제목/기간/신청마감 텍스트 레이어를 확인하세요.");
        }
        group.visible = true;
        wrapTitle(titleLayer, record.title, 300, maxTitleLines);
        fitFont(titleLayer, 305, maxTitleLines === 1 ? 20 : 45, maxTitleLines === 1 ? 7.5 : 9);
        periodLayer.textItem.contents = periodPrefix + "   " + trim(record.period);
        deadlineLayer.textItem.contents = "신청마감  " + trim(record.deadline);
        renameFromContents(titleLayer);
        renameFromContents(periodLayer);
        renameFromContents(deadlineLayer);
    }

    function recordsBySection(records, section) {
        var result = {};
        var i;
        for (i = 0; i < records.length; i++) {
            if (trim(records[i].section).toLowerCase() === section) {
                result[String(parseInt(records[i].slot, 10))] = records[i];
            }
        }
        return result;
    }

    function recordCount(records) {
        var count = 0;
        var key;
        for (key in records) {
            if (records.hasOwnProperty(key)) {
                count++;
            }
        }
        return count;
    }

    function updateTextSection(root, records, count, periodPrefix, maxTitleLines, warnings) {
        var i;
        for (i = 1; i <= count; i++) {
            try {
                var slotName = i < 10 ? "0" + i : String(i);
                var sectionPrefix = periodPrefix === "사업기간" ? "사업" : "교육";
                var slotGroup = findGroupAliases(
                    root,
                    [
                        slotName,
                        String(i),
                        slotName + "번",
                        String(i) + "번",
                        sectionPrefix + slotName,
                        sectionPrefix + String(i),
                        "리스트" + slotName,
                        "리스트" + String(i)
                    ],
                    true
                );
                if (!slotGroup) {
                    throw new Error(root.name + "/" + slotName + " 그룹을 찾을 수 없습니다.");
                }
                if (records[String(i)]) {
                    updateSlot(slotGroup, records[String(i)], periodPrefix, maxTitleLines);
                } else {
                    slotGroup.visible = false;
                }
            } catch (error) {
                warnings.push(periodPrefix + " " + i + "번: " + error.message);
            }
        }
    }

    function findShapeLayer(group) {
        var i;
        for (i = 0; i < group.artLayers.length; i++) {
            if (group.artLayers[i].kind === LayerKind.SOLIDFILL) {
                return group.artLayers[i];
            }
        }
        for (i = 0; i < group.layerSets.length; i++) {
            try {
                return findShapeLayer(group.layerSets[i]);
            } catch (ignore) {}
        }
        throw new Error(group.name + " 배너 그룹의 기준 도형을 찾지 못했습니다.");
    }

    function findArtLayerAliases(parent, aliases, recursive) {
        var i;
        for (i = 0; i < parent.artLayers.length; i++) {
            if (nameMatches(parent.artLayers[i].name, aliases)) {
                return parent.artLayers[i];
            }
        }
        if (!recursive) {
            return null;
        }
        for (i = 0; i < parent.layerSets.length; i++) {
            var layer = findArtLayerAliases(parent.layerSets[i], aliases, true);
            if (layer) {
                return layer;
            }
        }
        return null;
    }

    function findBannerMaskLayer(group, bannerRoot, slotNumber) {
        var aliases = BANNER_MASK_LAYER_NAMES[String(slotNumber)];
        if (!aliases || !aliases.length) {
            return findShapeLayer(group);
        }
        var maskLayer = findArtLayerAliases(group, aliases, true);
        if (!maskLayer) {
            maskLayer = findArtLayerAliases(bannerRoot, aliases, true);
        }
        if (!maskLayer) {
            throw new Error(
                "배너 " + slotNumber + "번의 둥근 마스크 레이어를 찾을 수 없습니다. " +
                "필요한 레이어 이름: " + aliases.join(", ")
            );
        }
        return maskLayer;
    }

    function hideOriginalBannerImages(group, shapeLayer) {
        var i;
        for (i = 0; i < group.artLayers.length; i++) {
            var candidate = group.artLayers[i];
            if (candidate === shapeLayer || candidate.name.indexOf("AUTO_BANNER_") === 0) {
                continue;
            }
            // 둥근 도형에 클리핑되어 있던 기존 래스터/고급 개체 이미지만 숨깁니다.
            // 텍스트와 도형 장식은 그대로 유지합니다.
            var isClipped = false;
            try {
                isClipped = candidate.grouped === true;
            } catch (ignore) {}
            if (isClipped && candidate.kind !== LayerKind.TEXT && candidate.kind !== LayerKind.SOLIDFILL) {
                candidate.visible = false;
            }
        }
    }

    function removeOldAutoBanner(group) {
        var i;
        for (i = group.artLayers.length - 1; i >= 0; i--) {
            if (group.artLayers[i].name.indexOf("AUTO_BANNER_") === 0) {
                group.artLayers[i].remove();
            }
        }
        for (i = group.layerSets.length - 1; i >= 0; i--) {
            removeOldAutoBanner(group.layerSets[i]);
        }
    }

    function pasteImage(targetDocument, imageFile) {
        var source = app.open(imageFile);
        source.selection.selectAll();
        source.selection.copy(true);
        source.close(SaveOptions.DONOTSAVECHANGES);
        app.activeDocument = targetDocument;
        targetDocument.paste();
        return targetDocument.activeLayer;
    }

    function createClippingMask(document, imageLayer, maskLayer) {
        // 새 이미지를 실제 둥근 마스크 바로 위에 둡니다.
        imageLayer.move(maskLayer, ElementPlacement.PLACEBEFORE);
        document.activeLayer = imageLayer;

        // Photoshop 메뉴의 "레이어 > 클리핑 마스크 만들기"와 같은 명령입니다.
        // DOM의 grouped 속성만 설정하면 일부 Photoshop 버전에서 적용되지 않아
        // Action Manager 명령을 먼저 실행합니다.
        var clipped = false;
        try {
            executeAction(charIDToTypeID("GrpL"), undefined, DialogModes.NO);
            clipped = imageLayer.grouped === true;
        } catch (actionError) {}

        if (!clipped) {
            try {
                imageLayer.grouped = true;
                clipped = imageLayer.grouped === true;
            } catch (propertyError) {}
        }

        if (!clipped) {
            throw new Error(
                "배너 이미지에 클리핑 마스크를 만들지 못했습니다. 기준 레이어: " + maskLayer.name
            );
        }
    }

    function placeBanner(document, bannerRoot, record, csvFolder) {
        var slotNumber = parseInt(record.slot, 10);
        var slotName = slotNumber < 10 ? "0" + slotNumber : String(slotNumber);
        var group = findGroupAliases(
            bannerRoot,
            [
                String(slotNumber),
                slotName,
                String(slotNumber) + "번",
                slotName + "번",
                "배너" + String(slotNumber),
                "배너" + slotName
            ],
            true
        );
        if (!group) {
            throw new Error("배너/" + record.slot + " 그룹을 찾을 수 없습니다.");
        }
        var relativePath = trim(record.image_path).replace(/\\/g, "/");
        if (!relativePath) {
            throw new Error("배너 " + record.slot + "번의 image_path가 비어 있습니다.");
        }
        var imageFile = new File(csvFolder.fsName + "/" + relativePath);
        if (!imageFile.exists) {
            throw new Error("배너 이미지를 찾을 수 없습니다: " + imageFile.fsName);
        }
        removeOldAutoBanner(group);
        var maskLayer = findBannerMaskLayer(group, bannerRoot, slotNumber);
        var imageGroup = maskLayer.parent;
        var target = boundsOf(maskLayer);
        hideOriginalBannerImages(imageGroup, maskLayer);
        maskLayer.visible = true;
        var layer = pasteImage(document, imageFile);
        layer.name = "AUTO_BANNER_" + record.slot;
        var current = boundsOf(layer);
        layer.resize((target.width / current.width) * 100, (target.height / current.height) * 100, AnchorPosition.MIDDLECENTER);
        current = boundsOf(layer);
        layer.translate(
            target.left + target.width / 2 - (current.left + current.width / 2),
            target.top + target.height / 2 - (current.top + current.height / 2)
        );
        // 마스크가 들어 있는 가장 안쪽 그룹으로 이동한 다음,
        // 실제 Photoshop 클리핑 마스크 명령을 실행합니다.
        layer.move(imageGroup, ElementPlacement.INSIDE);
        createClippingMask(document, layer, maskLayer);
        group.visible = true;
    }

    function updateBanners(document, bannerRoot, records, csvFolder, warnings) {
        var i;
        for (i = 1; i <= 3; i++) {
            // 배너 2는 고정 영역이므로 CSV 유무와 관계없이 원본을 보존합니다.
            if (i === 2) {
                continue;
            }
            try {
                if (records[String(i)]) {
                    placeBanner(document, bannerRoot, records[String(i)], csvFolder);
                } else {
                    var slotName = i < 10 ? "0" + i : String(i);
                    var group = findGroupAliases(
                        bannerRoot,
                        [
                            String(i),
                            slotName,
                            String(i) + "번",
                            slotName + "번",
                            "배너" + String(i),
                            "배너" + slotName
                        ],
                        true
                    );
                    if (group) {
                        group.visible = false;
                    }
                }
            } catch (error) {
                warnings.push("배너 " + i + "번: " + error.message);
            }
        }
    }

    function updateDate(document, isoDate) {
        isoDate = trim(isoDate);
        if (!isoDate) {
            return;
        }
        var match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(isoDate);
        if (!match) {
            throw new Error("newsletter_date는 YYYY-MM-DD 형식이어야 합니다: " + isoDate);
        }
        var titleRoot = requiredGroupAliases(document, ["타이틀", "제목", "헤더"], "타이틀");
        var layers = textLayers(titleRoot);
        var i;
        for (i = 0; i < layers.length; i++) {
            if (/^\d{4}년\s+\d{1,2}월\s+\d{1,2}일/.test(trim(layers[i].textItem.contents))) {
                layers[i].textItem.contents = parseInt(match[1], 10) + "년 " + parseInt(match[2], 10) + "월 " + parseInt(match[3], 10) + "일";
                renameFromContents(layers[i]);
                return;
            }
        }
        throw new Error("타이틀 그룹에서 발행일 텍스트 레이어를 찾지 못했습니다.");
    }

    function normalizedFilePath(value) {
        return String(value || "").replace(/\\/g, "/").toLowerCase();
    }

    function openOrReuseDocument(file) {
        var targetPath = normalizedFilePath(file.fsName);
        var i;
        for (i = 0; i < app.documents.length; i++) {
            try {
                if (normalizedFilePath(app.documents[i].fullName.fsName) === targetPath) {
                    app.activeDocument = app.documents[i];
                    return app.documents[i];
                }
            } catch (ignore) {}
        }
        return app.open(file);
    }

    function automationConfigFromTemplate(templateFile) {
        if (!templateFile || !templateFile.exists) {
            return null;
        }
        var dateMatch = /^newsletter_(20\d{6})\.psd$/i.exec(templateFile.name);
        if (!dateMatch) {
            return null;
        }

        var issueDate = dateMatch[1];
        var letterFolder = templateFile.parent;
        var inputFolder = letterFolder.parent;
        var workFolder = inputFolder.parent;
        if (
            normalizeGroupName(letterFolder.name) !== normalizeGroupName(issueDate + "_letter") ||
            normalizeGroupName(inputFolder.name) !== "input" ||
            workFolder.name !== issueDate
        ) {
            return null;
        }

        var outputFolder = new Folder(workFolder.fsName + "/output");
        var imageFolder = new Folder(outputFolder.fsName + "/images");
        return {
            csv: outputFolder.fsName + "/newsletter.csv",
            template: templateFile.fsName,
            output: outputFolder.fsName + "/newsletter_" + issueDate + ".psd",
            exportsDir: imageFolder.fsName,
            webPng: imageFolder.fsName + "/newsletter_" + issueDate + ".png",
            log: outputFolder.fsName + "/photoshop_automation_error.log"
        };
    }

    function templateFromCurrentContext() {
        if (typeof NEWSLETTER_TEMPLATE_PATH !== "undefined" && trim(NEWSLETTER_TEMPLATE_PATH)) {
            return new File(NEWSLETTER_TEMPLATE_PATH);
        }
        if (app.documents.length > 0) {
            try {
                return new File(app.activeDocument.fullName.fsName);
            } catch (ignore) {}
        }
        return null;
    }

    var document = null;
    var initialGroupTree = "";
    var warnings = [];
    try {
        var templateFile = autoConfig ? new File(autoConfig.template) : templateFromCurrentContext();
        if (!autoConfig && templateFile) {
            autoConfig = automationConfigFromTemplate(templateFile);
            if (autoConfig) {
                app.displayDialogs = DialogModes.NO;
            }
        }
        var csvFile = autoConfig ? new File(autoConfig.csv) : File.openDialog("1_excel_to_csv/export.py가 만든 newsletter.csv를 선택하세요", "CSV:*.csv");
        if (!csvFile) {
            return;
        }
        if (!csvFile.exists) {
            throw new Error("CSV를 찾을 수 없습니다: " + csvFile.fsName);
        }
        if (!templateFile) {
            templateFile = File.openDialog("원본 뉴스레터 PSD를 선택하세요", "Photoshop:*.psd");
        }
        if (!templateFile) {
            return;
        }
        if (!templateFile.exists) {
            throw new Error("원본 PSD를 찾을 수 없습니다: " + templateFile.fsName);
        }
        var records = readCSV(csvFile);
        if (!records.length) {
            warnings.push("CSV에 데이터 행이 없어 원본 PSD 내용을 유지합니다.");
        }

        document = openOrReuseDocument(templateFile);
        initialGroupTree = groupTree(document, "", 0);
        var business = recordsBySection(records, "business");
        var academy = recordsBySection(records, "academy");
        var banners = recordsBySection(records, "banner");

        var businessRoot = findGroupAliases(
            document,
            ["사업안내", "사업1-8", "사업 1-8", "사업1~8", "사업 1~8"],
            true
        );
        var academyRoot = findGroupAliases(
            document,
            ["아카데미안내", "아카데미", "아카데미 교육·연수 과정", "아카데미1-4", "아카데미 1-4"],
            true
        );
        var academyFour = null;
        if (academyRoot) {
            academyFour = findGroupAliases(academyRoot, ["4개 리스트", "4개리스트", "1-4", "01-04"], true);
            if (!academyFour) {
                academyFour = academyRoot;
            }
        }
        var bannerRoot = findGroupAliases(
            document,
            ["배너", "홍보배너", "배너1-3", "배너 1-3"],
            true
        );

        if (businessRoot && recordCount(business) > 0) {
            updateTextSection(businessRoot, business, 8, "사업기간", 2, warnings);
        } else if (businessRoot) {
            warnings.push("사업안내 CSV 데이터가 없어 원본 사업안내 레이어를 유지했습니다.");
        } else {
            warnings.push("사업안내/사업1-8 그룹을 찾지 못해 사업안내 입력을 건너뛰었습니다.");
        }
        if (academyFour && recordCount(academy) > 0) {
            updateTextSection(academyFour, academy, 4, "교육기간", 1, warnings);
        } else if (academyFour) {
            warnings.push("아카데미 CSV 데이터가 없어 원본 아카데미 레이어를 유지했습니다.");
        } else {
            warnings.push("아카데미 그룹을 찾지 못해 아카데미 입력을 건너뛰었습니다.");
        }

        if (bannerRoot && recordCount(banners) > 0) {
            updateBanners(document, bannerRoot, banners, csvFile.parent, warnings);
        } else if (bannerRoot) {
            warnings.push("배너 CSV 데이터가 없어 원본 배너 레이어를 유지했습니다.");
        } else {
            warnings.push("배너 그룹을 찾지 못해 배너 입력을 건너뛰었습니다.");
        }

        var newsletterDate = "";
        var i;
        for (i = 0; i < records.length; i++) {
            if (trim(records[i].newsletter_date)) {
                newsletterDate = records[i].newsletter_date;
                break;
            }
        }
        try {
            updateDate(document, newsletterDate);
        } catch (dateError) {
            warnings.push("발행일: " + dateError.message);
        }

        var outputFile;
        if (autoConfig) {
            outputFile = new File(autoConfig.output);
        } else {
            var outputDateMatch = /^(\d{4})-(\d{2})-(\d{2})$/.exec(trim(newsletterDate));
            var outputDateText = "";
            if (outputDateMatch) {
                outputDateText = outputDateMatch[1] + outputDateMatch[2] + outputDateMatch[3];
            } else {
                var templateDateMatch = /(20\d{6})/.exec(templateFile.name);
                if (templateDateMatch) {
                    outputDateText = templateDateMatch[1];
                }
            }
            if (!outputDateText) {
                throw new Error("CSV와 원본 PSD 파일명에서 발행일을 찾지 못해 결과 PSD 이름을 만들 수 없습니다.");
            }
            outputFile = new File(csvFile.parent.fsName + "/newsletter_" + outputDateText + ".psd");
        }
        if (!/\.psd$/i.test(outputFile.name)) {
            outputFile = new File(outputFile.fsName + ".psd");
        }
        if (outputFile.exists && !outputFile.remove()) {
            throw new Error("기존 결과 PSD를 덮어쓸 수 없습니다: " + outputFile.fsName);
        }
        var options = new PhotoshopSaveOptions();
        options.layers = true;
        options.embedColorProfile = true;
        options.maximizeCompatibility = true;
        document.saveAs(outputFile, options, true, Extension.LOWERCASE);

        // 완성된 뉴스레터 전체 캔버스를 PNG-24 "웹용으로 저장" 방식으로 내보냅니다.
        var pngOutputFolder;
        if (autoConfig) {
            pngOutputFolder = autoConfig.exportsDir ? new Folder(autoConfig.exportsDir) : outputFile.parent;
        } else {
            pngOutputFolder = ensureFolder(new Folder(outputFile.parent.fsName + "/images"));
        }
        var webPngFile;
        if (autoConfig && autoConfig.webPng) {
            webPngFile = new File(autoConfig.webPng);
        } else {
            var webPngName = outputFile.name.replace(/\.psd$/i, ".png");
            webPngFile = new File(pngOutputFolder.fsName + "/" + webPngName);
        }
        try {
            exportDocumentForWeb(document, webPngFile);
        } catch (webError) {
            warnings.push("전체 PNG: " + webError.message);
        }
        var pngOutputs = exportRequiredPNGs(document, bannerRoot, pngOutputFolder, warnings);
        if (autoConfig) {
            var status = warnings.length ? "PARTIAL" : "SUCCESS";
            var pngLines = "";
            for (i = 0; i < pngOutputs.length; i++) {
                pngLines += "\nPNG: " + pngOutputs[i].fsName;
            }
            writeAutomationLog(
                autoConfig.log,
                status + "\nOutput: " + outputFile.fsName +
                "\nWeb PNG: " + webPngFile.fsName +
                pngLines +
                (warnings.length ? "\n\nWarnings:\n- " + warnings.join("\n- ") : "") +
                "\n\nPSD layer group tree:\n" + initialGroupTree
            );
            document.close(SaveOptions.DONOTSAVECHANGES);
            document = null;
        } else {
            alert("완료했습니다.\n\n" + outputFile.fsName + "\n\n원본 PSD는 변경하지 않았습니다.");
        }
    } catch (error) {
        var message = "뉴스레터 자동화 오류\n\n" + error.message + (error.line ? "\n줄: " + error.line : "");
        if (initialGroupTree) {
            message += "\n\nPSD 레이어 그룹 구조:\n" + initialGroupTree;
        }
        if (autoConfig) {
            automationError = message;
            writeAutomationLog(autoConfig.log, "FAILED\n" + message);
            if (document) {
                try {
                    document.close(SaveOptions.DONOTSAVECHANGES);
                } catch (closeError) {}
                document = null;
            }
        } else {
            alert(message);
        }
    } finally {
        app.preferences.rulerUnits = oldRulerUnits;
        app.displayDialogs = oldDisplayDialogs;
    }
    if (automationError) {
        throw new Error(automationError);
    }
}());
