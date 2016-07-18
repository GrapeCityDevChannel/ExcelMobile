/*jshint -W030 */   // Expected an assignment or function call and instead saw an expression (a && a.fun1())
/*jshint -W004 */   // {a} is already defined (can use let instead of var in es6)
/* jshint laxbreak: true */

var app = app || {};
var spreadNS = GcSpread.Sheets;
var Calc = spreadNS.Calc;
var spread;
var ribbon;
var settingCache = {backColor: "white", foreColor: "black", tableStyles: 'medium2'};
var tableIndex = 1, pictureIndex = 1;
var fbx, isShiftKey = false;
var resourceMap = {},
    conditionalFormatTexts = {};
var $tableStyleDropdown;
var colorPicker, $colorPickerContainer;

function toggleGroupDisplay() {
    var $element = $(this),
        $content = $element.siblings(".pane-group-content"),
        $target = $('>span>span:first', $element),
        collapsed = $target.hasClass("glyphicon-triangle-right");

    // TODO: slideToggle makes the display of content a mess, need time to find a solution to resolve it
    if (collapsed) {
        $target.removeClass("glyphicon-triangle-right").addClass("glyphicon-triangle-bottom");
        //$content.slideToggle();
        $content.show();
    } else {
        $target.addClass("glyphicon-triangle-right").removeClass("glyphicon-triangle-bottom");
        //$content.slideToggle();
        $content.hide();
    }
}

var _mergeState = {};
function updateMergeState() {
    var sheet = spread.getActiveSheet();
    var sels = sheet.getSelections(),
        mergable = false,
        unmergable = false,
        hasMergeCells = false;

    sels.forEach(function (range) {
        var ranges = sheet.getSpans(range),
            spanCount = ranges.length;

        if (!hasMergeCells) {
            hasMergeCells = spanCount > 0;
        }

        if (!mergable) {
            if (spanCount > 1 || (spanCount === 0 && (range.rowCount > 1 || range.colCount > 1))) {
                mergable = true;
            } else if (spanCount === 1) {
                var range2 = ranges[0];
                if (range2.row !== range.row || range2.col !== range.col ||
                    range2.rowCount !== range2.rowCount || range2.colCount !== range.colCount) {
                    mergable = true;
                }
            }
        }
        if (!unmergable) {
            unmergable = spanCount > 0;
        }
    });

    _mergeState.mergable = mergable;
    _mergeState.unmergable = unmergable;

    // use active cell's merge state
    ribbon.setToggleButton('cellmerge', !!sheet.getSpan(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex()));
}

function updateCellStyleState(sheet, row, column) {
    var style = sheet.getActualStyle(row, column);

    if (style) {
        var sfont = style.font;

        // Font
        if (sfont) {
            var font = parseFont(sfont);

            ribbon.setToggleButton('bold', ["bold", "bolder", "700", "800", "900"].indexOf(font.fontWeight) !== -1);
            ribbon.setToggleButton("italic", font.fontStyle !== 'normal');
            ribbon.setInputValue('fontFamily', font.fontFamily.replace(/'/g, ""));
            ribbon.setInputValue('fontSize', parseFloat(font.fontSize));
        }

        var underline = spreadNS.TextDecorationType.Underline,
            linethrough = spreadNS.TextDecorationType.LineThrough,
            overline = spreadNS.TextDecorationType.Overline,
            textDecoration = style.textDecoration;

        ribbon.setToggleButton("underline", textDecoration && ((textDecoration & underline) === underline));
        ribbon.setToggleButton("strikethrough", textDecoration && ((textDecoration & linethrough) === linethrough));
        ribbon.setToggleButton("overline", textDecoration && ((textDecoration & overline) === overline));

        // Alignment
        ribbon.setToggleButton('wordwrap', style.wordwrap);

        // general (3, auto detect) without setting button just like Excel
        var alignType = ['left', 'center', 'right'][style.hAlign];
        ribbon.setToggleButton('halign-' + (alignType || 'left'), !!alignType);

        alignType = ['top', 'middle', 'bottom'][style.vAlign];
        ribbon.setToggleButton('valign-' + (alignType || 'top'), !!alignType);

        // locked
        ribbon.setToggleButton('unlockcells', style.locked);
    }
}

function markFontStyleButtonActive(name, active) {
    var $target = $("#setting-pane button[data-name='" + name + "']");

    if (active) {
        $target.addClass("toggle on");
    } else {
        $target.removeClass("toggle on");
    }
}

function parseFont(font) {
    var fontFamily = null,
        fontSize = null,
        fontStyle = "normal",
        fontWeight = "normal",
        fontVariant = "normal",
        lineHeight = "normal";

    var elements = font.split(/\s+/);
    var element;
    while ((element = elements.shift())) {
        switch (element) {
            case "normal":
                break;

            case "italic":
            case "oblique":
                fontStyle = element;
                break;

            case "small-caps":
                fontVariant = element;
                break;

            case "bold":
            case "bolder":
            case "lighter":
            case "100":
            case "200":
            case "300":
            case "400":
            case "500":
            case "600":
            case "700":
            case "800":
            case "900":
                fontWeight = element;
                break;

            default:
                if (!fontSize) {
                    var parts = element.split("/");
                    fontSize = parts[0];
                    if (fontSize.indexOf("px") !== -1) {
                        fontSize = px2pt(parseFloat(fontSize)) + 'pt';
                    }
                    if (parts.length > 1) {
                        lineHeight = parts[1];
                        if (lineHeight.indexOf("px") !== -1) {
                            lineHeight = px2pt(parseFloat(lineHeight)) + 'pt';
                        }
                    }
                    break;
                }

                fontFamily = element;
                if (elements.length)
                    fontFamily += " " + elements.join(" ");

                return {
                    "fontStyle": fontStyle,
                    "fontVariant": fontVariant,
                    "fontWeight": fontWeight,
                    "fontSize": fontSize,
                    "lineHeight": lineHeight,
                    "fontFamily": fontFamily
                };
        }
    }

    return {
        "fontStyle": fontStyle,
        "fontVariant": fontVariant,
        "fontWeight": fontWeight,
        "fontSize": fontSize,
        "lineHeight": lineHeight,
        "fontFamily": fontFamily
    };
}

var tempSpan = $("<span></span>");
function px2pt(pxValue) {
    tempSpan.css({
        "font-size": "96pt",
        "display": "none"
    });
    tempSpan.appendTo($(document.body));
    var tempPx = tempSpan.css("font-size");
    if (tempPx.indexOf("px") !== -1) {
        var tempPxValue = parseFloat(tempPx);
        return Math.round(pxValue * 96 / tempPxValue);
    }
    else {  // when browser have not convert pt to px, use 96 DPI.
        return Math.round(pxValue * 72 / 96);
    }
}

function setReferenceStyle(name) {
    var referenceStyle, columnHeaderAutoText;

    if (name === "a1style") {
        referenceStyle = spreadNS.ReferenceStyle.A1;
        columnHeaderAutoText = spreadNS.HeaderAutoText.letters;
    } else {
        referenceStyle = spreadNS.ReferenceStyle.R1C1;
        columnHeaderAutoText = spreadNS.HeaderAutoText.numbers;
    }

    spread.referenceStyle(referenceStyle);
    spread.sheets.forEach(function (sheet) {
        sheet.setColumnHeaderAutoText(columnHeaderAutoText);
    });
    updatePositionBox(spread.getActiveSheet());
}

function radioButtonClicked() {
    var $this = $(this), name = $this.attr('name'), value = $this.data('value');
    
    switch(name) {
        case "referenceStyle":
            setReferenceStyle(value);
            break;
            
        case "slicerMoveAndSize":
            setSlicerSetting("moveSize", value);    
            break;
    } 
}

function checkedChanged() {
    var $element = $(this),
        name = $element.parent().data("name");

    if ($element.hasClass("disabled")) {
        return;
    }

    var sheet = spread.getActiveSheet();

    var value = $element.prop('checked');

    spread.isPaintSuspended(true);

    switch (name) {
        case "cutCopyIndicatorVisible":
            spread.cutCopyIndicatorVisible(value);
            break;

        case "showVerticalScrollbar":
            spread.showVerticalScrollbar(value);
            break;

        case "showHorizontalScrollbar":
            spread.showHorizontalScrollbar(value);
            break;

        case "scrollIgnoreHidden":
            spread.scrollIgnoreHidden(value);
            break;

        case "scrollbarMaxAlign":
            spread.scrollbarMaxAlign(value);
            break;

        case "scrollbarShowMax":
            spread.scrollbarShowMax(value);
            break;

        case "tabStripVisible":
            spread.tabStripVisible(value);
            break;

        case "newTabVisible":
            spread.newTabVisible(value);
            break;

        case "tabEditable":
            spread.tabEditable(value);
            break;

        case "showTabNavigation":
            spread.tabNavigationVisible(value);
            break;

        case "showDragDropTip":
            spread.showDragDropTip(value);
            break;

        case "showDragFillTip":
            spread.showDragFillTip(value);
            break;

        case "sheetVisible":
            var sheetIndex = $element.data("sheetIndex"),
                sheetName = $element.data("sheetName"),
                selectedSheet = spread.sheets[sheetIndex];

            // be sure related sheet not changed (such add / remove sheet, rename sheet)
            if (selectedSheet && selectedSheet.getName() === sheetName) {
                selectedSheet.visible(value);
            } else {
                console.log("selected sheet' info was changed, please select the sheet and set visible again.");
            }
            break;

        case "canUserDragDrop":
            spread.canUserDragDrop(value);
            break;

        case "canUserDragFill":
            spread.canUserDragFill(value);
            break;

        case "allowZoom":
            spread.allowUserZoom(value);
            break;

        case "allowOverflow":
            spread.sheets.forEach(function (sheet) {
                sheet.allowCellOverflow(value);
            });
            break;

        case "showDragFillSmartTag":
            spread.showDragFillSmartTag(value);
            break;

        case "showRowRangeGroup":
            sheet.showRowRangeGroup(value);
            break;

        case "showColumnRangeGroup":
            sheet.showColumnRangeGroup(value);
            break;

        case "highlightInvalidData":
            spread.highlightInvalidData(value);
            break;

        /* table realted items */
        case "tableFilterButton":
            _activeTable && _activeTable.filterButtonVisible(value);
            break;

        case "tableHeaderRow":
            _activeTable && _activeTable.showHeader(value);
            break;

        case "tableTotalRow":
            _activeTable && _activeTable.showFooter(value);
            break;

        case "tableBandedRows":
            _activeTable && _activeTable.bandRows(value);
            break;

        case "tableBandedColumns":
            _activeTable && _activeTable.bandColumns(value);
            break;

        case "tableFirstColumn":
            _activeTable && _activeTable.highlightFirstColumn(value);
            break;

        case "tableLastColumn":
            _activeTable && _activeTable.highlightLastColumn(value);
            break;
        /* table realted items (end) */

        /* comment related items */
        case "commentDynamicSize":
            Actions.setCommentDynamicSize(spread, {comment: _activeComment, value: value});
            break;

        case "commentDynamicMove":
            Actions.setCommentDynamicMove(spread, {comment: _activeComment, value: value});
            break;

        case "commentLockText":
            Actions.setCommentLockText(spread, {comment: _activeComment, value: value});
            break;

        case "commentShowShadow":
            Actions.setCommentShowShadow(spread, {comment: _activeComment, value: value});
            break;
        /* comment related items (end) */

        /* slicer related items */
        case "displaySlicerHeader":
            setSlicerSetting("showHeader", value);
            break;

        case "lockSlicer":
            setSlicerSetting("lock", value);
            break;
            
        case "hide-no-data":
            enableRelatedItems(["mark-no-data", "show-no-data-last"], !value);
            setSlicerSetting("showNoDataItems", !value);
            break;
            
        case "mark-no-data":
            enableRelatedItems(["show-no-data-last"], value);
            setSlicerSetting("visuallyNoDataItems", value);
            break;
            
        case "show-no-data-last":
            setSlicerSetting("showNoDataItemsInLast", value);
            break;
        /* slicer related items (end) */
    }
    spread.isPaintSuspended(false);
}

function enableRelatedItems(names, enable) {
    names.forEach(function(name) {
        var $target = $("#setting-pane label[data-name='" + name + "']"),
            $input = $("input", $target);
            $target.attr("disabled", !enable);
            $input.prop("disabled", !enable);
    });
}

function updateNumberProperty() {
    var $element = $(this),
        name = $element.data("name"),
        value = parseInt($element.val(), 10);

    if (isNaN(value)) {
        return;
    }

    var sheet = spread.getActiveSheet();

    spread.isPaintSuspended(true);
    switch (name) {
        case "commentBorderWidth":
            Actions.setCommentBorderWidth(spread, {comment: _activeComment, value: value});
            break;

        case "commentOpacity":
            Actions.setCommentOpacity(spread, {comment: _activeComment, value: value / 100});
            break;

        case "slicerColumnNumber":
            setSlicerSetting("columnCount", value);
            break;

        case "slicerButtonHeight":
            setSlicerSetting("itemHeight", value);
            break;

        case "slicerButtonWidth":
            setSlicerSetting("itemWidth", value);
            break;

        default:
            console.log("updateNumberProperty need add for", name);
            break;
    }
    spread.isPaintSuspended(false);
}

function updateStringProperty() {
    var $element = $(this),
        name = $element.data("name"),
        value = $element.val();

    var sheet = spread.getActiveSheet();

    switch (name) {
        case "commentPadding":
            Actions.setCommentPadding(spread, {comment: _activeComment, value: value});
            break;

        case "slicerName":
            setSlicerSetting("name", value);
            break;

        case "slicerCaptionName":
            setSlicerSetting("captionName", value);
            break;

        default:
            console.log("updateStringProperty w/o process of ", name);
            break;
    }
}

app.fillSheetNameList = function() {
    var $ul = $("#sheetNameList");
    $ul.empty();
    spread.sheets.forEach(function (sheet, index) {
        $('<li><a class="text" data-value="' + index + '">' + sheet.getName() + '</a></li>').appendTo($ul);
    });
    $("a", $("li", $ul).first()).click();    
};

function syncSpreadPropertyValues() {
    // General
    setCheckValue("canUserDragDrop", spread.canUserDragDrop());
    setCheckValue("canUserDragFill", spread.canUserDragFill());
    setCheckValue("allowZoom", spread.allowUserZoom());
    setCheckValue("allowOverflow", spread.getActiveSheet().allowCellOverflow());
    setCheckValue("showDragFillSmartTag", spread.showDragFillSmartTag());

    // Calculation
    setRadioItemChecked("referenceStyle", spread.referenceStyle() === spreadNS.ReferenceStyle.R1C1 ? "r1c1style" : "a1style");

    // Scroll Bar
    setCheckValue("showVerticalScrollbar", spread.showVerticalScrollbar());
    setCheckValue("showHorizontalScrollbar", spread.showHorizontalScrollbar());
    setCheckValue("scrollbarMaxAlign", spread.scrollbarMaxAlign());
    setCheckValue("scrollbarShowMax", spread.scrollbarShowMax());
    setCheckValue("scrollIgnoreHidden", spread.scrollIgnoreHidden());

    // TabStrip
    setCheckValue("tabStripVisible", spread.tabStripVisible());
    setCheckValue("newTabVisible", spread.newTabVisible());
    setCheckValue("tabEditable", spread.tabEditable());
    setCheckValue("showTabNavigation", spread.tabNavigationVisible());

    // Color
    setColorValue("spreadBackcolor", spread.backColor());
    setColorValue("grayAreaBackcolor", spread.grayAreaBackColor());

    // Tip
    setDropDown("scrollTip", spread.showScrollTip());
    setDropDown("resizeTip", spread.showResizeTip());
    setCheckValue("showDragDropTip", spread.showDragDropTip());
    setCheckValue("showDragFillTip", spread.showDragFillTip());

    // Cut / Copy Indicator
    setCheckValue("cutCopyIndicatorVisible", spread.cutCopyIndicatorVisible());
    setColorValue("cutCopyIndicatorBorderColor", spread.cutCopyIndicatorBorderColor());

    // Data validation
    ribbon.setToggleButton('circleInvalidData', spread.highlightInvalidData());
    ribbon.setCheckboxButton('showSheetTabs', spread.tabStripVisible());
}

function syncSheetPropertyValues() {
    var sheet = spread.getActiveSheet();

    // Grid Line
    var options = sheet.getGridlineOptions();
    ribbon.setCheckboxButton('showGridlines', options.showHorizontalGridline && options.showVerticalGridline);

    // Header
    ribbon.setCheckboxButton('showHeadings', sheet.getRowHeaderVisible() && sheet.getColumnHeaderVisible());

    // Protection
    var isProtected = sheet.getIsProtected();
    
    ribbon.setToggleButton("protectsheet", isProtected);

    updateCellStyleState(sheet, sheet.getActiveRowIndex(), sheet.getActiveColumnIndex());

    if (!$(sheet).data("bind")) {
        $(sheet).data("bind", true);
        sheet.bind(spreadNS.Events.RangeChanged, function (event, args) {
            if (args.action === spreadNS.RangeChangedAction.Clear) {
                // check special type items and switch to cell tab (laze process)
                if (isSpecialTabSelected()) {
                    onCellSelected();
                }
            }
        });

        sheet.bind(spreadNS.Events.CommentRemoved, function (event, args) {
            // check special type items and switch to cell tab (laze process)
            if (isSpecialTabSelected()) {
                onCellSelected();
            }
        });
    }
}

function setNumberValue(name, value) {
    $(".setting-container input[data-name='" + name + "']").val(value);
}

function getNumberValue(name) {
    return +$("input[data-name='" + name + "']").val();
}

function setTextValue(name, value) {
    $(".setting-container input[data-name='" + name + "']").val(value);
}

function getTextValue(name) {
    return $(".setting-container input[data-name='" + name + "']").val();
}

function setCheckValue(name, value, options) {
    var $input = $(".setting-container label[data-name='" + name + "'] input");

    $input.prop('checked', value);

    if (options) {
        $input.data(options);
     }
}

function getCheckValue(name) {
    return $(".setting-container label[data-name='" + name + "']>input").prop('checked');
}

function setColorValue(name, value) {
    $("div[data-name='" + name + "'] div.color-picker").css("background-color", value || "");
}

function processDropDownClicked(name, numberValue, value, nameValue, $element, $group) {
    switch (name) {
        case "scrollTip":
            spread.showScrollTip(numberValue);
            break;

        case "resizeTip":
            spread.showResizeTip(numberValue);
            break;

        case "selectionPolicy":
            sheet.selectionPolicy(numberValue);
            break;

        case "selectionUnit":
            sheet.selectionUnit(numberValue);
            break;

        case "sheetName":
            var selectedSheet = spread.sheets[numberValue];
            setCheckValue("sheetVisible", selectedSheet.visible(), {
                sheetIndex: numberValue,
                sheetName: selectedSheet.getName()
            });
            break;

        case "commentFontFamily":
            Actions.setCommentFontFamily(spread, {comment: _activeComment, value: value});
            break;

        case "commentFontSize":
            value += "pt";
            Actions.setCommentFontSize(spread, {comment: _activeComment, value: value});
            break;

        case "commentDisplayMode":
            Actions.setCommentDisplayMode(spread, {comment: _activeComment, value: numberValue});
            break;

        case "commentFontStyle":
            Actions.setCommentFontStyle(spread, {comment: _activeComment, value: nameValue});
            break;

        case "commentFontWeight":
            Actions.setCommentFontWeight(spread, {comment: _activeComment, value: nameValue});
            break;

        case "commentBorderStyle":
            Actions.setCommentBorderStyle(spread, {comment: _activeComment, value: nameValue});
            break;

        case "commentHorizontalAlign":
            Actions.setCommentHorizontalAlign(spread, {comment: _activeComment, value: numberValue});
            break;

        case "zoomSpread":
            processZoomSetting(nameValue, value);
            break;

        case "commonFormat":
            if (nameValue === "custom") {
                _customFormatInput.focus();
            } else {
                var formatter = nameValue === 'nullValue' ? null : nameValue;
                Actions.setCellFormat(spread, {name: name, value: formatter});
                _customFormatInput.val(formatter || '');
            }
            break;

        case "minAxisType":
            updateManual(nameValue, "manualMin");
            break;

        case "maxAxisType":
            updateManual(nameValue, "manualMax");
            break;

        case "slicerItemSorting":
            processSlicerItemSorting(numberValue);
            break;

        case "spreadTheme":
            processChangeSpreadTheme(nameValue);
            break;

        case "resizeZeroIndicator":
            spread.resizeZeroIndicator(numberValue);
            break;
            
        case "printHeader":
            updatePreview({ header: $element.data("sections") });
            break;
            
        case "printFooter":
            updatePreview({ footer: $element.data("sections") });
            break;
            
        case "image-content":
            if (!app.isInit) {
                if (!$element.data("value")) {
                    // clear data to avoid wrong result if user canceled
                    $group.data("value", null);
                    
                    // browse to add image
                    $("#fileSelector").data("action", "addPrintImage");
                    $("#fileSelector").attr("accept", "image/*");
                    $("#fileSelector").click();
                }
            }
            break;

        case "sparklineExType":
            processSparklineSetting(nameValue);
            break;
            
        case "validatorType":
            processDataValidationSetting(nameValue);
            break;

        default:
            console.log("TODO add processDropDownClicked for ", name, value);
            break;
    }
}

function getDropDownValue(name) {
    var value = $("div[data-name='" + name + "']").data("value");

    /*jshint eqnull:true */
    if (value == null) {
        console.log('get dropdown value null / undefined', name);
    }

    return value;
}

function colorSelected(event, eventData) {
    var themeColor = eventData.themeColor,
        value = eventData.color,
        type = eventData.type;

    var data = $(colorPicker).data(),
        name = data.name,
        $target = data.target;

    var sheet = spread.getActiveSheet();

    // No Fills need special process
    if (name === "backColor" && type === "nofill") {
        value = undefined;
    }
    
    if ($target) {
        $target.css("background-color", value);
        // save to help selected corresponding color in color picker
        $target.data('themeColor', themeColor);
    }

    spread.isPaintSuspended(true);
    switch (name) {
        case "spreadBackcolor":
            spread.backColor(value);
            break;

        case "grayAreaBackcolor":
            spread.grayAreaBackColor(value);
            break;

        case "cutCopyIndicatorBorderColor":
            spread.cutCopyIndicatorBorderColor(value);
            break;

        case "foreColor":
            Actions.setTextForeColor(spread, {name: name, value: themeColor || value});

            // save to cache for reuse by quick set with one-click
            settingCache[name] = themeColor || value;
            break;

        case "backColor":
            Actions.setTextBackColor(spread, {name: name, value: themeColor || value});

            // save to cache for reuse by quick set with one-click
            settingCache[name] = themeColor || value;
            break;

        case "commentBorderColor":
            Actions.setCommentBorderColor(spread, {comment: _activeComment, value: value});
            break;

        case "commentForeColor":
            Actions.setCommentForeColor(spread, {comment: _activeComment, value: value});
            break;

        case "commentBackColor":
            Actions.setCommentBackColor(spread, {comment: _activeComment, value: value});
            break;

        default:
            if (!$target) { 
                console.log("TODO colorSelected", name);
            }
            break;
    }
    spread.isPaintSuspended(false);

    colorPicker.hide();
    var $dropdown = data.dropdown;
    if ($dropdown) {
        $dropdown.removeClass('open');
        $dropdown.parents("div.dropdown.open").data("keepopen", false);
    }
}

function sortData(sheet, ascending) {
    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    for (var n = 0; n < sels.length; n++) {
        var sel = getActualCellRange(sels[n], rowCount, columnCount);
        sheet.sortRange(sel.row, sel.col, sel.rowCount, sel.colCount, true,
            [
                {index: sel.col, ascending: ascending}
            ]);
    }
}

function updateFilter(sheet) {
    if (sheet.rowFilter()) {
        sheet.rowFilter(null);
    } else {
        var sels = sheet.getSelections();
        if (sels.length > 0) {
            var sel = sels[0];
            sheet.rowFilter(new spreadNS.HideRowFilter(sel));
        }
    }
}

function addGroup(sheet) {
    var sels = sheet.getSelections();
    var sel = sels[0];

    if (!sel) return;

    if (sel.col === -1) // row selection
    {
        var groupExtent = new GcSpread.Sheets.UndoRedo.GroupExtent(sel.row, sel.rowCount);
        var action = new GcSpread.Sheets.UndoRedo.RowGroupUndoAction(sheet, groupExtent);
        spread.doCommand(action);
    }
    else if (sel.row === -1) // column selection
    {
        var groupExtent = new GcSpread.Sheets.UndoRedo.GroupExtent(sel.col, sel.colCount);
        var action = new GcSpread.Sheets.UndoRedo.ColumnGroupUndoAction(sheet, groupExtent);
        spread.doCommand(action);
    }
    else // cell range selection
    {
        alert(getResource("messages.rowColumnRangeRequired"));
    }
}

function removeGroup(sheet) {
    var sels = sheet.getSelections();
    var sel = sels[0];

    if (!sel) return;

    if (sel.col === -1 && sel.row === -1) // sheet selection
    {
        sheet.rowRangeGroup.ungroup(0, sheet.getRowCount());
        sheet.colRangeGroup.ungroup(0, sheet.getColumnCount());
    }
    else if (sel.col === -1) // row selection
    {
        var groupExtent = new GcSpread.Sheets.UndoRedo.GroupExtent(sel.row, sel.rowCount);
        var action = new GcSpread.Sheets.UndoRedo.RowUngroupUndoAction(sheet, groupExtent);
        spread.doCommand(action);
    }
    else if (sel.row === -1) // column selection
    {
        var groupExtent = new GcSpread.Sheets.UndoRedo.GroupExtent(sel.col, sel.colCount);
        var action = new GcSpread.Sheets.UndoRedo.ColumnUngroupUndoAction(sheet, groupExtent);
        spread.doCommand(action);
    }
    else // cell range selection
    {
        alert(getResource("messages.rowColumnRangeRequired"));
    }
}

function toggleGroupDetail(sheet, expand) {
    var sels = sheet.getSelections();
    var sel = sels[0];

    if (!sel) return;

    if (sel.col === -1 && sel.row === -1) // sheet selection
    {
    }
    else if (sel.col === -1) // row selection
    {
        for (var i = 0; i < sel.rowCount; i++) {
            var rgi = sheet.rowRangeGroup.find(sel.row + i, 0);
            if (rgi) {
                sheet.rowRangeGroup.expand(rgi.level, expand);
            }
        }
    }
    else if (sel.row === -1) // column selection
    {
        for (var i = 0; i < sel.colCount; i++) {
            var rgi = sheet.colRangeGroup.find(sel.col + i, 0);
            if (rgi) {
                sheet.colRangeGroup.expand(rgi.level, expand);
            }
        }
    }
    else // cell range selection
    {
    }
}

var MARGIN_BOTTOM = 4;

function adjustSpreadSize() {
    var height = $("#inner-content-container").height() - ($("#formulaBar:visible").length > 0 ? $("#formulaBar").height() : 0) - MARGIN_BOTTOM,
        spreadHeight = $("#ss").height();

    if (spreadHeight !== height) {
        $("#controlPanel").height(height);
        $("#ss").height(height);
        $("#ss").data("spread").refresh();
    }
}

function screenAdoption() {
    function adjustSettingPaneContentHeight() {
        var $container = $("#setting-pane");
        var height = $container.innerHeight() - $('.pane-header', $container).outerHeight();

        $('.pane-content').outerHeight(height);
    }

    function setTableStyleDropdownWidth() {
        var width = $("#inner-content-container").width();

        // set drop down width to make sure 7 or 6 style items shown one row
        $tableStyleDropdown.css({width: width > 640 ? '510px' : '444px'});
    }

    // close popup
    $(".dropdown-menu:visible").parent().removeClass('open');
    $("#colorpicker").hide();

    hideSpreadContextMenu();
    adjustSpreadSize();

    // explicit set formula box' width instead of 100% because it's contained in table
    var width = $("#inner-content-container").width() - $("#positionbox").outerWidth() - 1; // 1: border' width of td contains formulabox (left only)
    $("#formulabox").css({width: width});

    ribbon.resize();

    setTableStyleDropdownWidth();

    //adjustSettingPaneContentHeight();
}

function doPrepareWork() {
    function addEventHandlers() {
        $("div.pane-group-title").click(toggleGroupDisplay);
        $(document).on("colorSelected", colorSelected);
    }

    addEventHandlers();

    $("input[type='number']:not('.not-min-zero')").attr("min", 0);
    $("input[data-name='commentOpacity']").attr("max", 100);

    // set default values
    var value = "anyvalue-validator";
    setDropDown("validatorType", value);
    processDataValidationSetting(value);         // Data Validation Setting
    value = GcSpread.Sheets.ComparisonOperator.Between;
    setDropDown("numberValidatorComparisonOperator", value);        // NumberValidator Comparison Operator
    processNumberValidatorComparisonOperatorSetting(value);
    setDropDown("dateValidatorComparisonOperator", value);          // DateValidator Comparison Operator
    processDateValidatorComparisonOperatorSetting(value);
    setDropDown("textLengthValidatorComparisonOperator", value);    // TextLengthValidator Comparison Operator
    processTextLengthValidatorComparisonOperatorSetting(value);

    conditionalFormatTexts = uiResource.conditionalFormat.texts;
}

function initSpread() {
    $.get( "files/excel-demo-intro.json", function( data ) {
        importSpreadFromJSON(null, data);

        //Change default allowCellOverflow the same with Excel.
        spread.sheets.forEach(function (sheet) {
            sheet.allowCellOverflow(true);
        });
    },
    "text"  /* dataType, the file is spread export format (ssjson, renamed to .json to avoid server not support access .ssjson file) */);
}

function getCellInfo(sheet, row, column) {
    var result = {type: ""}, object;

    if ((object = sheet.getComment(row, column))) {
        result.type = "comment";
    } else if ((object = sheet.findTable(row, column))) {
        result.type = "table";
    }

    result.object = object;

    return result;
}

var specialTabNames = ["table"];
var specialTabRefs = specialTabNames.map(function (name) {
    return "#" + name;
});
var $specialTabs;

function isSpecialTabSelected() {
    var href = $(".toolbar>ul>li.active a").attr("href");

    return specialTabRefs.indexOf(href) !== -1;
}

function getSpecialTabs() {
    $specialTabs = $specialTabs || $(".toolbar>ul>li").filter(function () {
            return specialTabRefs.indexOf($('a', this).attr('href')) !== -1;
        });

    return $specialTabs;
}

function getSpecialDropdownItems() {
    return $("#tabDropdown li").filter(function () {
        return specialTabRefs.indexOf($('a', this).attr('data-href')) !== -1;
    });
}

function getVisibleSpecialItems() {
    var result = [];

    result.push(getSpecialTabs().filter(function () {
        return $(this).is(":visible");
    }));
    result.push(getSpecialDropdownItems());

    return result;
}

function getTabItem(tabName) {
    return $(".toolbar>ul>li>a[href='#" + tabName + "']").parent();
}

function setActiveTab(tabName, needProcessSpecialItems) {
    //TODO, rename and don't call it if not required
    if (needProcessSpecialItems === undefined) {
        return;
    }

    // show / hide tabs
    var $target = getTabItem(tabName);

    if (specialTabNames.indexOf(tabName) >= 0) {
        if ($target.hasClass("hidden")) {
            hideSpecialTabs(false);

            ribbon.showTab($target, true);
        }
    } else {
        var $active = $(".toolbar>ul>li.active");
        if (isSpecialTabSelected()) {
            ribbon.hideTab([$active], true, $target);
        } else {
            if (needProcessSpecialItems) {
                var $items = getVisibleSpecialItems();
                ribbon.hideTab($items, true);
            }
        }
        var oldComment = _activeComment;
        hideSpecialTabs(true);
        if(tabName == "comment"){
            _activeComment = oldComment;
        }
    }
}

// TODO: update according the requirement, current only table need process with
function shouldProcessSpecialTab() {
    /*jshint eqnull:true */
    return _activeTable != null;
}


function hideSpecialTabs(clearCache) {
    specialTabNames.forEach(function (name) {
        getTabItem(name).addClass("hidden");
    });

    if (clearCache) {
        clearCachedItems();
    }
}

function getActualRange(range, maxRowCount, maxColCount) {
    var row = range.row < 0 ? 0 : range.row;
    var col = range.col < 0 ? 0 : range.col;
    var rowCount = range.rowCount < 0 ? maxRowCount : range.rowCount;
    var colCount = range.colCount < 0 ? maxColCount : range.colCount;

    return new spreadNS.Range(row, col, rowCount, colCount);
}

function getActualCellRange(cellRange, rowCount, columnCount) {
    if (cellRange.row === -1 && cellRange.col === -1) {
        return new spreadNS.Range(0, 0, rowCount, columnCount);
    }
    else if (cellRange.row === -1) {
        return new spreadNS.Range(0, cellRange.col, rowCount, cellRange.colCount);
    }
    else if (cellRange.col === -1) {
        return new spreadNS.Range(cellRange.row, 0, cellRange.rowCount, columnCount);
    }

    return cellRange;
}

function attachEvents() {
    attachSpreadEvents();
    attachDataValidationEvents();
    attachOtherEvents();
    attachCellTypeEvents();
    attachSparklineSettingEvents();
    attachFindEvents();
    attachSettingPaneEvents();

    $(document).on('click', "li.item label", function (e) {
        // avoid dropdown to be closed when click 
        e.stopPropagation();
    });
    
    var $positionbox = $("#positionbox");
    var insertFormulaTitle = $("#formulas button[data-name=insertFormula]").data("header");
    $positionbox.click(function(e) {
        displaySettingPane(insertFormulaTitle, $('#functionBuiilder'));
    });    

    $("#borderSetting .border-line-style ul.dropdown-menu>li").on('click', function () {
        var $li = $(this),
            value = $li.attr('data-value'),
            $target = $('>button>span:first', $li.parents('div.btn-group'));

        if (value === 'none') {
            $target.text($li.text());
        } else {
            $target.text('');
        }
        $target.removeClass().addClass(($('>a>div', $li).attr('class') || '') + ' border-line-style ');
        $li.siblings('.selected').removeClass('selected');
        $li.addClass('selected');
    });
    
    $("#borderSetting .border-type-item").on('click', function() {
        var borderType = $(this).data("name");
        setCellsBorder(borderType);
    });

    // used to avoid dropdown been closed by default
    $("#tableStyles").click(ignoreEvent);

    $(document).on('click', ".pane-color-picker>button", function () {
        if (colorPicker.isVisible()) {
            colorPicker.hide();
            return;
        }

        if ($colorPickerContainer.parent()[0] !== document.body) {
            $colorPickerContainer.appendTo($(document.body));
        }

        var $this = $(this),
            name = $this.parents('div.btn-group').attr('data-name'),
            $target = $('.color-picker', $this);
        $(colorPicker).data({name: name, target: $target});
        
        var offset, position;
        var $settingPane = $('#setting-pane'); 
        if ($.contains($settingPane[0], this)) {
            offset = $settingPane.offset();
            position = { left: offset.left - $('#colorPicker').innerWidth(), top: offset.top };
        } else {
            offset = $this.offset();
            // TODO: adjust position to avoid out of screen
            position = { left: offset.left + $this.outerWidth() + 20, top: offset.top - ( $('#colorPicker').height() - $this.outerHeight() ) / 2 };
        }
 
        colorPicker.show(position, {}, $target.data('themeColor') || $target.css("background-color"));
    });
}

// Border Type related items
function getBorderLineType(className) {
    switch (className) {
        case "line-style-none":
            return GcSpread.Sheets.LineStyle.empty;

        case "line-style-hair":
            return GcSpread.Sheets.LineStyle.hair;

        case "line-style-dotted":
            return GcSpread.Sheets.LineStyle.dotted;

        case "line-style-dash-dot-dot":
            return GcSpread.Sheets.LineStyle.dashDotDot;

        case "line-style-dash-dot":
            return GcSpread.Sheets.LineStyle.dashDot;

        case "line-style-dashed":
            return GcSpread.Sheets.LineStyle.dashed;

        case "line-style-thin":
            return GcSpread.Sheets.LineStyle.thin;

        case "line-style-medium-dash-dot-dot":
            return GcSpread.Sheets.LineStyle.mediumDashDotDot;

        case "line-style-slanted-dash-dot":
            return GcSpread.Sheets.LineStyle.slantedDashDot;

        case "line-style-medium-dash-dot":
            return GcSpread.Sheets.LineStyle.mediumDashDot;

        case "line-style-medium-dashed":
            return GcSpread.Sheets.LineStyle.mediumDashed;

        case "line-style-medium":
            return GcSpread.Sheets.LineStyle.medium;

        case "line-style-thick":
            return GcSpread.Sheets.LineStyle.thick;

        case "line-style-double":
            return GcSpread.Sheets.LineStyle.double;
    }
}
// Border Type related items (end)

function attachOtherEvents() {
    $("div.table-format-item").click(changeTableStyle);
    $("div.slicer-format-item").click(changeSlicerStyle);
    $("#spreadContextMenu a").click(processContextMenuClicked);
    $("#fileSelector").change(processFileSelected);
}

function processFileSelected() {
    var file = this.files[0],
        action = $(this).data("action");

    if (!file) return false;

    // clear to make sure change event occures even when same file selected again
    $("#fileSelector").val("");
    $("#fileSelector").attr("accept", null);

    if (action === "doImport") {
        return importSpreadFromJSON(file);
    }

    if (!/image\/\w+/.test(file.type)) {
        alert(getResource("messages.imageFileRequired"));
        return false;
    }
    var reader = new FileReader();
    reader.onload = function () {
        switch (action) {
            case "addpicture":
                Actions.addPicture(spread, {name: "Picture" + pictureIndex++, url: this.result});
                break;
                
            case "addPrintImage":
                var name = file.name, dataUrl = this.result;
                
                addUploadedImage(name, dataUrl);
                
                addImageListItem(name);
                
                console.log("addPrintImage ....", file.name);
                break;
        }
    };
    reader.readAsDataURL(file);
}

function addUploadedImage(name, dataUrl) {
    if (!app.uploadImages) {
        app.uploadImages = {};
    }
    app.uploadImages[name] = dataUrl;
}

function addImageListItem(name) {
    var $li = $("<li></li>"), $a = $("<a></a>");
            
    $a.text(name).data({ value: name }).appendTo($li);
    
    $("#imageContentList").append($li);
    $a.click();
}

function updatePositionBox(sheet) {
    var selection = sheet.getSelections().slice(-1)[0];
    if (selection) {
        var position;
        if (!isShiftKey) {
            position = getCellPositionString(sheet,
                sheet.getActiveRowIndex() + 1,
                sheet.getActiveColumnIndex() + 1, selection);
        }
        else {
            position = getSelectedRangeString(sheet, selection);
        }

        $("#positionbox").text(position);
    }
}

function syncCellRelatedItems() {
    updateMergeState();

    // sync cell type related information
    syncCellTypeInfo();
}

function syncCellTypeInfo() {
    function updateButtonCellTypeInfo(cellType) {
        setNumberValue("buttonCellTypeMarginTop", cellType.marginTop());
        setNumberValue("buttonCellTypeMarginRight", cellType.marginRight());
        setNumberValue("buttonCellTypeMarginBottom", cellType.marginBottom());
        setNumberValue("buttonCellTypeMarginLeft", cellType.marginLeft());
        setTextValue("buttonCellTypeText", cellType.text());
        setColorValue("buttonCellTypeBackColor", cellType.buttonBackColor());
    }

    function updateCheckBoxCellTypeInfo(cellType) {
        setTextValue("checkboxCellTypeCaption", cellType.caption());
        setTextValue("checkboxCellTypeTextTrue", cellType.textTrue());
        setTextValue("checkboxCellTypeTextIndeterminate", cellType.textIndeterminate());
        setTextValue("checkboxCellTypeTextFalse", cellType.textFalse());
        setDropDown("checkboxCellTypeTextAlign", cellType.textAlign());
        setCheckValue("checkboxCellTypeIsThreeState", cellType.isThreeState());
    }

    function updateComboBoxCellTypeInfo(cellType) {
        setDropDown("comboboxCellTypeEditorValueType", cellType.editorValueType());
        var items = cellType.items(),
            texts = items.map(function (item) {
                return item.text || item;
            }).join(","),
            values = items.map(function (item) {
                return item.value || item;
            }).join(",");

        setTextValue("comboboxCellTypeItemsText", texts);
        setTextValue("comboboxCellTypeItemsValue", values);
    }

    function updateHyperLinkCellTypeInfo(cellType) {
        setColorValue("hyperlinkCellTypeLinkColor", cellType.linkColor());
        setColorValue("hyperlinkCellTypeVisitedLinkColor", cellType.visitedLinkColor());
        setTextValue("hyperlinkCellTypeText", cellType.text());
        setTextValue("hyperlinkCellTypeLinkToolTip", cellType.linkToolTip());
    }

    var sheet = spread.getActiveSheet(),
        index,
        cellType = sheet.getCell(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex()).cellType();

    if (cellType instanceof spreadNS.ButtonCellType) {
        index = 0;
        updateButtonCellTypeInfo(cellType);
    } else if (cellType instanceof spreadNS.CheckBoxCellType) {
        index = 1;
        updateCheckBoxCellTypeInfo(cellType);
    } else if (cellType instanceof spreadNS.ComboBoxCellType) {
        index = 2;
        updateComboBoxCellTypeInfo(cellType);
    } else if (cellType instanceof spreadNS.HyperLinkCellType) {
        index = 3;
        updateHyperLinkCellTypeInfo(cellType);
    } else {
        index = -1;
    }

    if (index >= 0) {
        var $groups = $('#cellTypeSetting .group-celltype');
            $groups.addClass('hidden');
            $($groups[index]).removeClass('hidden');
            
        var title = $('div[data-name="celltype"] ul.dropdown-menu>li:eq(' + index + ')>a').text();

        displaySettingPane(title, $('#cellTypeSetting'));
    }
}

function onCellSelected() {
    var sheet = spread.getActiveSheet(),
        row = sheet.getActiveRowIndex(),
        column = sheet.getActiveColumnIndex();
    if (showSparklineSetting(row, column)) {
        displaySettingPane(uiResource.sparklineDialog.detail, $("#sparklineDetailSetting"));
        return;
    }
    var cellInfo = getCellInfo(sheet, row, column),
        cellType = cellInfo.type;

    syncCellRelatedItems();
    updatePositionBox(sheet);
    updateCellStyleState(sheet, row, column);

    var tabType = "home";

    // check with special tab before clear cache which it depends on
    var needProcessSpecialItems = shouldProcessSpecialTab();
    clearCachedItems();

    // add map from cell type to tab type here
    if (cellType === "table") {
        tabType = "table";
        syncTablePropertyValues(sheet, cellInfo.object);
        //$("#addslicer").removeClass("hidden");
    } else if (cellType === "comment") {
        tabType = "comment";
        syncCommentPropertyValues(sheet, cellInfo.object);
    }

    setActiveTab(tabType, needProcessSpecialItems);
}

var _activeComment;

function syncCommentPropertyValues(sheet, comment) {
    _activeComment = comment;
    // General
    setCheckValue("commentDynamicSize", comment.dynamicSize());
    setCheckValue("commentDynamicMove", comment.dynamicMove());
    setCheckValue("commentLockText", comment.lockText());
    setCheckValue("commentShowShadow", comment.showShadow());

    // Font
    setDropDown("commentFontFamily", comment.fontFamily());
    setDropDown("commentFontSize", parseFloat(comment.fontSize()));
    setDropDown("commentFontStyle", comment.fontStyle());
    setDropDown("commentFontWeight", comment.fontWeight());
    var textDecoration = comment.textDecoration();
    var TextDecorationType = spreadNS.TextDecorationType;
    markFontStyleButtonActive("comment-underline", (textDecoration & TextDecorationType.Underline) === TextDecorationType.Underline);
    markFontStyleButtonActive("comment-overline", (textDecoration & TextDecorationType.Overline) === TextDecorationType.Overline);
    markFontStyleButtonActive("comment-strikethrough", (textDecoration & TextDecorationType.LineThrough) === TextDecorationType.LineThrough);

    // Border
    setNumberValue("commentBorderWidth", comment.borderWidth());
    setDropDown("commentBorderStyle", comment.borderStyle());
    setColorValue("commentBorderColor", comment.borderColor());

    // Appearance
    setDropDown("commentHorizontalAlign", spreadNS.HorizontalAlign[comment.horizontalAlign()]);
    setDropDown("commentDisplayMode", comment.displayMode());
    setColorValue("commentForeColor", comment.foreColor());
    setColorValue("commentBackColor", comment.backColor());
    setTextValue("commentPadding", getPaddingString(comment.padding()));
    setNumberValue("commentOpacity", comment.opacity() * 100);

    displaySettingPane(uiResource.settingPane.title.comment, $('#commentSetting'));
}

function getPaddingString(padding) {
    if (!padding) return "";

    return [padding.top, padding.right, padding.bottom, padding.left].join(", ");
}

function clearCachedItems() {
    _activePicture = null;
    _activeComment = null;
    _activeTable = null;
}

var _activeTable;
function syncTablePropertyValues(sheet, table) {
    _activeTable = table;

    ribbon.setToggleButton("tableHeaderRow", table.showHeader());
    ribbon.setToggleButton("tableTotalRow", table.showFooter());

    ribbon.setToggleButton("tableFirstColumn", table.highlightFirstColumn());
    ribbon.setToggleButton("tableLastColumn", table.highlightLastColumn());
    ribbon.setToggleButton("tableBandedRows", table.bandRows());
    ribbon.setToggleButton("tableBandedColumns", table.bandColumns());
    var tableStyle = table.style(),
        styleName = tableStyle && table.style().name();

    $("#tableStyles .table-format-item").removeClass("table-format-item-selected");
    if (styleName) {
        $("#tableStyles .table-format-item div[data-name='" + styleName.toLowerCase() + "']").parent().addClass("table-format-item-selected");
    }
    ribbon.setInputValue("tableName", table.name());
}

function setTableStyle(table, styleName) {
    if (table) {
        var tableStyle = styleName === "none" ? new spreadNS.TableStyle() : spreadNS.TableStyles[styleName]();

        table.style(tableStyle);
    }
}

function changeTableStyle() {
    if (_activeTable) {
        spread.isPaintSuspended(true);

        var $item = $(".table-format-icon", this),
            styleName = $item.data("name"),
            tooltip = $item.parent().attr("title");

        setTableStyle(_activeTable, styleName);

        $("#tableStyles .table-format-item").removeClass("table-format-item-selected");
        $(this).addClass("table-format-item-selected");

        spread.isPaintSuspended(false);

        // close drop down
        $tableStyleDropdown.parent().removeClass('open');

        // Save to cache as the quick set by left icon
        settingCache.tableStyles = styleName;
        // update tooltip for icon
        $tableStyleDropdown.parent().prev().attr("title", tooltip);
    }
}

var _activePicture;

function processSelectionChanged() {
    syncCellRelatedItems();

    updatePositionBox(spread.getActiveSheet());
}

function attachSpreadEvents(rebind) {
    spread.bind(spreadNS.Events.EnterCell, onCellSelected);

    spread.bind(spreadNS.Events.ValueChanged, function (sender, args) {
        var row = args.row, col = args.col, sheet = args.sheet;

        if (sheet.getCell(row, col).wordWrap()) {
            sheet.autoFitRow(row);
        }
    });

    spread.bind(spreadNS.Events.RangeChanged, function (sender, args) {
        var sheet = args.sheet, row = args.row, rowCount = args.rowCount;
        if (args.action === spreadNS.RangeChangedAction.Paste) {
            for (var i = 0; i < rowCount; i++) {
                sheet.autoFitRow(row + i);
            }
        }
    });

    spread.bind(spreadNS.Events.ActiveSheetChanged, function () {
        syncSheetPropertyValues();
        syncCellRelatedItems();
        hideSpreadContextMenu();

        var sheet = spread.getActiveSheet(),
            picture;
        var slicers = sheet.getSlicers();
        for (var item in slicers) {
            slicers[item].isSelected(false);
        }
        sheet.setActiveCell(0, 0);
        
        onCellSelected();

        var value = getCheckValue("allowOverflow");
        if (sheet.allowCellOverflow() !== value) {
            sheet.allowCellOverflow(value);
        }
    });

    spread.bind(spreadNS.Events.SelectionChanging, function () {
        var sheet = spread.getActiveSheet();
        var selection = sheet.getSelections().slice(-1)[0];
        if (selection) {
            var position = getSelectedRangeString(sheet, selection);
            $("#positionbox").text(position);
        }
    });

    spread.bind(spreadNS.Events.SelectionChanged, function () {
        processSelectionChanged();
    });

    spread.bind(spreadNS.Events.CommentChanged, function (event, args) {
        var sheet = args.sheet, comment = args.comment, propertyName = args.propertyName;

        if (propertyName === "commentState" && comment) {
            if (comment.commentState() === spreadNS.CommentState.Edit) {
                syncCommentPropertyValues(sheet, comment);
            }
        }
    });
    
    spread.bind(spreadNS.Events.ValidationError, function (event, data) {
        var dv = data.validator;
        if (dv) {
            alert(dv.errorMessage);
        }
    });

    spread.bind(spreadNS.Events.SlicerChanged, function (event, args) {
        var sheet = args.sheet, slicer = args.slicer, propertyName = args.propertyName;

        if (!slicer) return;

        if (propertyName === "isSelected") {
            syncSlicerPropertyValues(sheet);
            if (slicer.isSelected()) {
                // display setting pane
                displaySettingPane(uiResource.settingPane.title.slicer, $("#slicerSetting"));
            }
        } else {
            changeSlicerInfo(slicer, propertyName);
        }
    });

    if (!rebind) {
        $(document).bind("keydown", function (event) {
            if (event.shiftKey) {
                isShiftKey = true;
            }
        });
        $(document).bind("keyup", function (event) {
            if (!event.shiftKey) {
                isShiftKey = false;

                var sheet = spread.getActiveSheet(),
                    position = getCellPositionString(sheet, sheet.getActiveRowIndex() + 1, sheet.getActiveColumnIndex() + 1);
                $("#positionbox").text(position);
            }
        });

        $("#ss").bind("contextmenu", processSpreadContextMenu);
        $("#ss").mouseup(function (e) {
            // hide context menu when the mouse down on SpreadJS
            if (e.button !== 2) {
                hideSpreadContextMenu();
            }
        });
    }
}

function getBackgroundColor(name) {
    return $("div[data-name='" + name + "'] div.color-picker").css("background-color");
}

// Cell Type related items
function attachCellTypeEvents() {
    $("#setCellTypeButton").click(function () {
        var currentCellType = $('#cellTypeSetting .group-celltype:visible').data('name') + '-celltype'; //getDropDownValue("cellTypes");
        var cellType = applyCellType(currentCellType);

        Actions.setCellType(spread, cellType);
    });
}

function applyCellType(name) {
    function linkify(address) {
        if (address && address.length) {
            if (address.startsWith("http://") || address.startsWith("https://")) {
                return address;
            }
            return "http://" + address;
        }

        return null;
    }

    var sheet = spread.getActiveSheet();
    var cellType, cellValue;
    switch (name) {
        case "button-celltype":
            cellType = new GcSpread.Sheets.ButtonCellType();
            cellType.marginTop(getNumberValue("buttonCellTypeMarginTop"));
            cellType.marginRight(getNumberValue("buttonCellTypeMarginRight"));
            cellType.marginBottom(getNumberValue("buttonCellTypeMarginBottom"));
            cellType.marginLeft(getNumberValue("buttonCellTypeMarginLeft"));
            cellType.text(getTextValue("buttonCellTypeText"));
            cellType.buttonBackColor(getBackgroundColor("buttonCellTypeBackColor"));
            break;

        case "checkbox-celltype":
            cellType = new GcSpread.Sheets.CheckBoxCellType();
            cellType.caption(getTextValue("checkboxCellTypeCaption"));
            cellType.textTrue(getTextValue("checkboxCellTypeTextTrue"));
            cellType.textIndeterminate(getTextValue("checkboxCellTypeTextIndeterminate"));
            cellType.textFalse(getTextValue("checkboxCellTypeTextFalse"));
            cellType.textAlign(getDropDownValue("checkboxCellTypeTextAlign"));
            cellType.isThreeState(getCheckValue("checkboxCellTypeIsThreeState"));
            break;

        case "combobox-celltype":
            cellType = new GcSpread.Sheets.ComboBoxCellType();
            cellType.editorValueType(getDropDownValue("comboboxCellTypeEditorValueType"));
            var comboboxItemsText = getTextValue("comboboxCellTypeItemsText");
            var comboboxItemsValue = getTextValue("comboboxCellTypeItemsValue");
            var itemsText = comboboxItemsText.split(",");
            var itemsValue = comboboxItemsValue.split(",");
            var itemsLength = itemsText.length > itemsValue.length ? itemsText.length : itemsValue.length;
            var items = [];
            for (var count = 0; count < itemsLength; count++) {
                var t = itemsText.length > count && itemsText[0] !== "" ? itemsText[count] : undefined;
                var v = itemsValue.length > count && itemsValue[0] !== "" ? itemsValue[count] : undefined;
                if (t !== undefined && v !== undefined) {
                    items[count] = {text: t, value: v};
                }
                else if (t !== undefined) {
                    items[count] = {text: t};
                } else if (v !== undefined) {
                    items[count] = {value: v};
                }
            }
            cellType.items(items);
            break;

        case "hyperlink-celltype":
            cellType = new GcSpread.Sheets.HyperLinkCellType();
            cellType.linkColor(getBackgroundColor("hyperlinkCellTypeLinkColor"));
            cellType.visitedLinkColor(getBackgroundColor("hyperlinkCellTypeVisitedLinkColor"));
            cellType.text(getTextValue("hyperlinkCellTypeText"));
            cellType.linkToolTip(getTextValue("hyperlinkCellTypeLinkToolTip"));
            cellValue = linkify(getTextValue("hyperlinkCellTypeAddress"));
            if (!cellValue) {
                alert("Please provides the address for the link.");
                return;
            }
            break;
    }

    return cellType;
}

function clearCellType() {
    var sheet = spread.getActiveSheet();
    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();
    sheet.isPaintSuspended(true);
    for (var i = 0; i < sels.length; i++) {
        var sel = getActualCellRange(sels[i], rowCount, columnCount);
        sheet.clear(sel.row, sel.col, sel.rowCount, sel.colCount, GcSpread.Sheets.SheetArea.viewport, GcSpread.Sheets.StorageType.Style);
    }
    sheet.isPaintSuspended(false);
}
// Cell Type related items (end)

// Data Validation related items
function processDataValidationSetting(name) {
    $("#dataValidationErrorAlertMessage").val("");
    $("#dataValidationErrorAlertTitle").val("");
    $("#dataValidationInputTitle").val("");
    $("#dataValidationInputMessage").val("");
    switch (name) {
        case "anyvalue-validator":
            $("#validatorNumberType").hide();
            $("#validatorListType").hide();
            $("#validatorFormulaListType").hide();
            $("#validatorDateType").hide();
            $("#validatorTextLengthType").hide();
            break;

        case "number-validator":
            $("#validatorNumberType").show();
            $("#validatorListType").hide();
            $("#validatorFormulaListType").hide();
            $("#validatorDateType").hide();
            $("#validatorTextLengthType").hide();
            processNumberValidatorComparisonOperatorSetting(getDropDownValue("numberValidatorComparisonOperator"));

            setTextValue("numberMinimum", 0);
            setTextValue("numberMaximum", 0);
            setTextValue("numberValue", 0);
            break;

        case "list-validator":
            $("#validatorNumberType").hide();
            $("#validatorListType").show();
            $("#validatorFormulaListType").hide();
            $("#validatorDateType").hide();
            $("#validatorTextLengthType").hide();

            setTextValue("listSource", "1,2,3");
            break;

        case "formulalist-validator":
            $("#validatorNumberType").hide();
            $("#validatorListType").hide();
            $("#validatorFormulaListType").show();
            $("#validatorDateType").hide();
            $("#validatorTextLengthType").hide();

            setTextValue("formulaListFormula", "=ISERROR(FIND(\" \",A1))");
            break;

        case "date-validator":
            $("#validatorNumberType").hide();
            $("#validatorListType").hide();
            $("#validatorFormulaListType").hide();
            $("#validatorDateType").show();
            $("#validatorTextLengthType").hide();
            processDateValidatorComparisonOperatorSetting(getDropDownValue("dateValidatorComparisonOperator"));

            var date = getCurrentTime();
            setTextValue("startDate", date);
            setTextValue("endDate", date);
            setTextValue("dateValue", date);
            break;

        case "textlength-validator":
            $("#validatorNumberType").hide();
            $("#validatorListType").hide();
            $("#validatorFormulaListType").hide();
            $("#validatorDateType").hide();
            $("#validatorTextLengthType").show();
            processTextLengthValidatorComparisonOperatorSetting(getDropDownValue("textLengthValidatorComparisonOperator"));

            setNumberValue("textLengthMinimum", 0);
            setNumberValue("textLengthMaximum", 0);
            setNumberValue("textLengthValue", 0);
            break;

        case "formula-validator":
            $("#validatorNumberType").hide();
            $("#validatorListType").hide();
            $("#validatorFormulaListType").show();
            $("#validatorDateType").hide();
            $("#validatorTextLengthType").hide();

            setTextValue("formulaListFormula", "E5:I5");
            break;

        default:
            console.log("processDataValidationSetting not process with ", name);
            break;
    }
}

function processNumberValidatorComparisonOperatorSetting(value) {
    if (value === GcSpread.Sheets.ComparisonOperator.Between || value === GcSpread.Sheets.ComparisonOperator.NotBetween) {
        $("#numberValue").hide();
        $("#numberBetweenOperator").show();
    }
    else {
        $("#numberBetweenOperator").hide();
        $("#numberValue").show();
    }
}

function processDateValidatorComparisonOperatorSetting(value) {
    if (value === GcSpread.Sheets.ComparisonOperator.Between || value === GcSpread.Sheets.ComparisonOperator.NotBetween) {
        $("#dateValue").hide();
        $("#dateBetweenOperator").show();
    }
    else {
        $("#dateBetweenOperator").hide();
        $("#dateValue").show();
    }
}

function processTextLengthValidatorComparisonOperatorSetting(value) {
    if (value === GcSpread.Sheets.ComparisonOperator.Between || value === GcSpread.Sheets.ComparisonOperator.NotBetween) {
        $("#textLengthValue").hide();
        $("#textLengthBetweenOperator").show();
    }
    else {
        $("#textLengthBetweenOperator").hide();
        $("#textLengthValue").show();
    }
}

function setDataValidator() {
    var validatorType = getDropDownValue("validatorType");
    var defaultDataValidator = GcSpread.Sheets.DefaultDataValidator;
    var currentDataValidator = null;
    var dropDownValue;

    var formulaListFormula = getTextValue("formulaListFormula");

    switch (validatorType) {
        case "anyvalue-validator":
            currentDataValidator = new GcSpread.Sheets.DefaultDataValidator();
            break;
        case "number-validator":
            var numberMinimum = getTextValue("numberMinimum");
            var numberMaximum = getTextValue("numberMaximum");
            var numberValue = getTextValue("numberValue");
            var isInteger = getCheckValue("isInteger");
            dropDownValue = getDropDownValue("numberValidatorComparisonOperator");
            if (dropDownValue !== GcSpread.Sheets.ComparisonOperator.Between && dropDownValue !== GcSpread.Sheets.ComparisonOperator.NotBetween) {
                numberMinimum = numberValue;
            }
            if (isInteger) {
                currentDataValidator = defaultDataValidator.createNumberValidator(dropDownValue,
                    isNaN(numberMinimum) ? numberMinimum : parseInt(numberMinimum, 10),
                    isNaN(numberMaximum) ? numberMaximum : parseInt(numberMaximum, 10),
                    true);
            } else {
                currentDataValidator = defaultDataValidator.createNumberValidator(dropDownValue,
                    isNaN(numberMinimum) ? numberMinimum : parseFloat(numberMinimum, 10),
                    isNaN(numberMaximum) ? numberMaximum : parseFloat(numberMaximum, 10),
                    false);
            }
            break;
        case "list-validator":
            var listSource = getTextValue("listSource");
            currentDataValidator = defaultDataValidator.createListValidator(listSource);
            break;
        case "formulalist-validator":
            currentDataValidator = defaultDataValidator.createFormulaListValidator(formulaListFormula);
            break;
        case "date-validator":
            var startDate = getTextValue("startDate");
            var endDate = getTextValue("endDate");
            var dateValue = getTextValue("dateValue");
            var isTime = getCheckValue("isTime");
            dropDownValue = getDropDownValue("dateValidatorComparisonOperator");
            if (dropDownValue !== GcSpread.Sheets.ComparisonOperator.Between && dropDownValue !== GcSpread.Sheets.ComparisonOperator.NotBetween) {
                startDate = dateValue;
            }
            if (isTime) {
                currentDataValidator = defaultDataValidator.createDateValidator(dropDownValue,
                    isNaN(startDate) ? startDate : new Date(startDate),
                    isNaN(endDate) ? endDate : new Date(endDate),
                    true);
            } else {
                currentDataValidator = defaultDataValidator.createDateValidator(dropDownValue,
                    isNaN(startDate) ? startDate : new Date(startDate),
                    isNaN(endDate) ? endDate : new Date(endDate),
                    false);
            }
            break;
        case "textlength-validator":
            var textLengthMinimum = getNumberValue("textLengthMinimum");
            var textLengthMaximum = getNumberValue("textLengthMaximum");
            var textLengthValue = getNumberValue("textLengthValue");
            dropDownValue = getDropDownValue("textLengthValidatorComparisonOperator");
            if (dropDownValue !== GcSpread.Sheets.ComparisonOperator.Between && dropDownValue !== GcSpread.Sheets.ComparisonOperator.NotBetween) {
                textLengthMinimum = textLengthValue;
            }
            currentDataValidator = defaultDataValidator.createTextLengthValidator(dropDownValue, textLengthMinimum, textLengthMaximum);
            break;
        case "formula-validator":
            currentDataValidator = defaultDataValidator.createFormulaValidator(formulaListFormula);
            break;
    }

    if (currentDataValidator) {
        currentDataValidator.errorMessage = $("#dataValidationErrorAlertMessage").val();
        currentDataValidator.errorStyle = getDropDownValue("errorAlert");
        currentDataValidator.errorTitle = $("#dataValidationErrorAlertTitle").val();
        currentDataValidator.showErrorMessage = getCheckValue("showErrorAlert");
        currentDataValidator.ignoreBlank = getCheckValue("ignoreBlank");
        var showInputMessage = getCheckValue("showInputMessage");
        if (showInputMessage) {
            currentDataValidator.inputTitle = $("#dataValidationInputTitle").val();
            currentDataValidator.inputMessage = $("#dataValidationInputMessage").val();
        }

        setDataValidatorInRange(currentDataValidator);
    }
}

function setDataValidatorInRange(dataValidator) {
    var sheet = spread.getActiveSheet();
    sheet.isPaintSuspended(true);
    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    for (var i = 0; i < sels.length; i++) {
        var sel = getActualCellRange(sels[i], rowCount, columnCount);
        for (var r = 0; r < sel.rowCount; r++) {
            for (var c = 0; c < sel.colCount; c++) {
                sheet.setDataValidator(sel.row + r, sel.col + c, dataValidator);
            }
        }
    }
    sheet.isPaintSuspended(false);
}

function getCurrentTime() {
    var date = new Date();
    var year = date.getFullYear();
    var month = date.getMonth() + 1;
    var day = date.getDate();

    var strDate = year + "-";
    if (month < 10)
        strDate += "0";
    strDate += month + "-";
    if (day < 10)
        strDate += "0";
    strDate += day;

    return strDate;
}

function attachDataValidationEvents() {
    $("#setDataValidator").click(function () {
        var currentValidatorType = getDropDownValue("validatorType");
        setDataValidator(currentValidatorType);
    });
    $("#clearDataValidatorSettings").click(function () {
        // reset to default
        var value = "anyvalue-validator";
        setDropDown("validatorType", value);
        processDataValidationSetting(value);
        setDropDown("errorAlert", 0);
        setCheckValue("showInputMessage", true);
        setCheckValue("showErrorAlert", true);
    });
}
// Data Validation related items (end)

// Sparkline related items
function processAddSparklineEx(sparklineType) {
    var sheet = spread.getActiveSheet();
    var selection = sheet.getSelections()[0];
    if (!selection) {
        return;
    }

    var $typeInfo = $("#sparklineSetting ul.dropdown-menu.sparklineExType a[data-value='" + sparklineType + "']");
    if ($typeInfo.length > 0) {
        setDropDown("sparklineExType", sparklineType);
        processSparklineSetting(sparklineType);
    }
    else {
        processSparklineSetting(getDropDownValue("sparklineExType"));
    }
    setTextValue("txtLineDataRange", parseRangeToExpString(selection));
    setTextValue("txtLineLocationRange", "");

    displaySettingPane(uiResource.sparklineDialog.title, $('#sparklineSetting'));
}

function unParseFormula(expr, row, col) {
    var sheet = spread.getActiveSheet();
    if (!sheet) {
        return null;
    }
    var calcService = sheet.getCalcService();
    return calcService.unparse(null, expr, row, col);
}

function processSparklineSetting(name, title) {
    //Show only when data range is illegal.
    $("#dataRangeError").parent().hide();
    $("#singleDataRangeError").parent().hide();
    //Show only when location range is illegal.
    $("#locationRangeError").parent().hide();

    switch (name) {
        case "line":
        case "column":
        case "winloss":
        case "pie":
        case "area":
        case "scatter":
        case "spread":
        case "stacked":
        case "boxplot":
        case "cascade":
        case "pareto":
            $("#lineContainer").show();
            $("#bulletContainer").hide();
            $("#hbarContainer").hide();
            $("#varianceContainer").hide();
            break;

        case "bullet":
            $("#lineContainer").hide();
            $("#bulletContainer").show();
            $("#hbarContainer").hide();
            $("#varianceContainer").hide();

            setTextValue("txtBulletMeasure", "");
            setTextValue("txtBulletTarget", "");
            setTextValue("txtBulletMaxi", "");
            setTextValue("txtBulletGood", "");
            setTextValue("txtBulletBad", "");
            setTextValue("txtBulletForecast", "");
            setTextValue("txtBulletTickunit", "");
            setCheckValue("checkboxBulletVertial", false);
            break;

        case "hbar":
        case "vbar":
            $("#lineContainer").hide();
            $("#bulletContainer").hide();
            $("#hbarContainer").show();
            $("#varianceContainer").hide();

            setTextValue("txtHbarValue", "");
            break;

        case "vari":
            $("#lineContainer").hide();
            $("#bulletContainer").hide();
            $("#hbarContainer").hide();
            $("#varianceContainer").show();

            setTextValue("txtVariance", "");
            setTextValue("txtVarianceReference", "");
            setTextValue("txtVarianceMini", "");
            setTextValue("txtVarianceMaxi", "");
            setTextValue("txtVarianceMark", "");
            setTextValue("txtVarianceTickUnit", "");
            setCheckValue("checkboxVarianceLegend", false);
            setCheckValue("checkboxVarianceVertical", false);
            break;

        default:
            console.log("processSparklineSetting not process with ", name, title);
            break;
    }
}

function addSparklineEvent() {
    var sheet = spread.getActiveSheet(),
        selection = sheet.getSelections()[0],
        isValid = true;

    var name = getDropDownValue("sparklineExType"),
        sparklineExType = name.toUpperCase() + "SPARKLINE";
    if (selection) {
        var range = getActualRange(selection, sheet.getRowCount(), sheet.getColumnCount());
        var formulaStr = '', row = range.row, col = range.col, direction = 0;
        switch (name) {
            case "bullet":
                var measure = getTextValue("txtBulletMeasure"),
                    target = getTextValue("txtBulletTarget"),
                    maxi = getTextValue("txtBulletMaxi"),
                    good = getTextValue("txtBulletGood"),
                    bad = getTextValue("txtBulletBad"),
                    forecast = getTextValue("txtBulletForecast"),
                    tickunit = getTextValue("txtBulletTickunit"),
                    colorScheme = getBackgroundColor("colorBulletColorScheme"),
                    vertical = getCheckValue("checkboxBulletVertial");
                formulaStr = '=' + sparklineExType + '(' + measure + ',' + target + ',' + maxi + ',' + good + ',' + bad + ',' + forecast + ',' + tickunit + ',' + '"' + colorScheme + '"' + ',' + vertical + ')';
                
                Actions.setFormulaSparkline(spread, {row: row, col: col, formula: formulaStr});

                break;
            case "hbar":
                var value = getTextValue("txtHbarValue"),
                    colorScheme = getBackgroundColor("colorHbarColorScheme");

                formulaStr = '=' + sparklineExType + '(' + value + ',' + '"' + colorScheme + '"' + ')';

                Actions.setFormulaSparkline(spread, {row: row, col: col, formula: formulaStr});
                break;
            case "vbar":
                var value = getTextValue("txtHbarValue"),
                    colorScheme = getBackgroundColor("colorHbarColorScheme");

                formulaStr = '=' + sparklineExType + '(' + value + ',' + '"' + colorScheme + '"' + ')';

                Actions.setFormulaSparkline(spread, {row: row, col: col, formula: formulaStr});
                break;
            case "vari":
                var variance = getTextValue("txtVariance"),
                    reference = getTextValue("txtVarianceReference"),
                    mini = getTextValue("txtVarianceMini"),
                    maxi = getTextValue("txtVarianceMaxi"),
                    mark = getTextValue("txtVarianceMark"),
                    tickunit = getTextValue("txtVarianceTickunit"),
                    colorPositive = getBackgroundColor("colorVariancePositive"),
                    colorNegative = getBackgroundColor("colorVarianceNegative"),
                    legend = getCheckValue("checkboxVarianceLegend"),
                    vertical = getCheckValue("checkboxVarianceVertical");

                formulaStr = '=' + sparklineExType + '(' + variance + ',' + reference + ',' + mini + ',' + maxi + ',' + mark + ',' + tickunit + ',' + legend + ',' + '"' + colorPositive + '"' + ',' + '"' + colorNegative + '"' + ',' + vertical + ')';

                Actions.setFormulaSparkline(spread, {row: row, col: col, formula: formulaStr});
                break;
            case "cascade":
            case "pareto":
                var dataRangeStr = getTextValue("txtLineDataRange"),
                    locationRangeStr = getTextValue("txtLineLocationRange"),
                    dataRangeObj = parseStringToExternalRanges(dataRangeStr, sheet),
                    locationRangeObj = parseStringToExternalRanges(locationRangeStr, sheet),
                    vertical = false,
                    dataRange, locationRange;
                if (dataRangeObj && dataRangeObj.length > 0 && dataRangeObj[0].range) {
                    dataRange = dataRangeObj[0].range;
                }
                if (locationRangeObj && locationRangeObj.length > 0 && locationRangeObj[0].range) {
                    locationRange = locationRangeObj[0].range;
                }
                if (locationRange && locationRange.rowCount < locationRange.colCount) {
                    vertical = true;
                }
                if (!dataRange) {
                    isValid = false;
                    $("#dataRangeError").parent().show();
                }
                if (!locationRange) {
                    isValid = false;
                    $("#locationRangeError").parent().show();
                }
                if (isValid) {
                    var pointCount = dataRange.rowCount * dataRange.colCount,
                        i = 1;
                        
                        // TODO: should in one action instead of multi-actions
                    for (var r = locationRange.row; r < locationRange.row + locationRange.rowCount; r++) {
                        for (var c = locationRange.col; c < locationRange.col + locationRange.colCount; c++) {
                            if (i <= pointCount) {
                                formulaStr = '=' + sparklineExType + '(' + dataRangeStr + ',' + i + ',,,,,,' + vertical + ')';

                                Actions.setFormulaSparkline(spread, {row: r, col: c, formula: formulaStr});
                                sheet.setActiveCell(r, c);
                                i++;
                            }
                        }
                    }
                }
                break;

            default:
                var dataRangeStr = getTextValue("txtLineDataRange"),
                    locationRangeStr = getTextValue("txtLineLocationRange"),
                    dataRangeObj = parseStringToExternalRanges(dataRangeStr, sheet),
                    locationRangeObj = parseStringToExternalRanges(locationRangeStr, sheet),
                    dataRange, locationRange;
                if (dataRangeObj && dataRangeObj.length > 0 && dataRangeObj[0].range) {
                    dataRange = dataRangeObj[0].range;
                }
                if (locationRangeObj && locationRangeObj.length > 0 && locationRangeObj[0].range) {
                    locationRange = locationRangeObj[0].range;
                }
                if (!dataRange) {
                    isValid = false;
                    $("#dataRangeError").parent().show();
                }
                if (!locationRange) {
                    isValid = false;
                    $("#locationRangeError".parent()).show();
                }
                if (isValid) {
                    if (["line", "column", "winloss"].indexOf(name) >= 0) {
                        if (dataRange.rowCount === 1) {
                            direction = 1;
                        }
                        else if (dataRange.colCount === 1) {
                            direction = 0;
                        }
                        else {
                            $("#singleDataRangeError").parent().show();
                            isValid = false;
                        }
                        if (isValid) {
                            formulaStr = '=' + sparklineExType + '(' + dataRangeStr + ',' + direction + ')';
                        }
                    }
                    else {
                        formulaStr = '=' + sparklineExType + '(' + dataRangeStr + ')';
                    }
                    if (isValid) {
                        row = locationRange.row;
                        col = locationRange.col;

                        Actions.setFormulaSparkline(spread, {row: row, col: col, formula: formulaStr});
                        sheet.setActiveCell(row, col);
                    }
                }
                break;
        }
    }
    if (!isValid) {
        return {canceled: true};
    }
    else {
        if (showSparklineSetting(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex())) {
            updateFormulaBar();
            setActiveTab("sparklineEx");
            displaySettingPane(uiResource.sparklineDialog.detail, $("#sparklineDetailSetting"));
            return;
        }
        console.log("Added sparkline", sparklineExType);
    }
}

function parseRangeToExpString(range) {
    var Calc = GcSpread.Sheets.Calc;
    return Calc.rangeToFormula(range, 0, 0, Calc.RangeReferenceRelative.allRelative);
}

function parseStringToExternalRanges(expString, sheet) {
    var Calc = GcSpread.Sheets.Calc;
    var results = [];
    var exps = expString.split(",");
    try {
        for (var i = 0; i < exps.length; i++) {
            var range = Calc.formulaToRange(sheet, exps[i]);
            results.push({"range": range});
        }
    }
    catch (e) {
        return null;
    }
    return results;
}

function parseFormulaSparkline(row, col) {
    var sheet = spread.getActiveSheet();
    if (!sheet) {
        return null;
    }
    var formula = sheet.getFormula(row, col);
    if (!formula) {
        return null;
    }
    var calcService = sheet.getCalcService();
    try {
        var expr = calcService.parse(null, formula, row, col);
        if (expr instanceof spreadNS.Calc.Expressions.FunctionExpression) {
            var fnName = expr.getFunctionName();
            if (fnName && spread.getSparklineEx(fnName)) {
                return expr;
            }
        }
    }
    catch (ex) {
    }
    return null;
}

function parseColorExpression(colorExpression, row, col) {
    if (!colorExpression) {
        return null;
    }
    var sheet = spread.getActiveSheet();
    if (colorExpression instanceof spreadNS.Calc.Expressions.StringExpression) {
        return colorExpression.value;
    }
    else if (colorExpression instanceof spreadNS.Calc.Expressions.MissingArgumentExpression) {
        return null;
    }
    else {
        var formula = null;
        try {
            formula = unParseFormula(colorExpression, row, col);
        }
        catch (ex) {
        }
        return spreadNS.Calc.evaluateFormula(sheet, formula, row, col);
    }
}

function getAreaSparklineSetting(formulaArgs, row, col) {
    var defaultValue = {colorPositive: "#787878", colorNegative: "#CB0000"};
    if (formulaArgs[0]) {
        setTextValue("areaSparklinePoints", unParseFormula(formulaArgs[0], row, col));
    }
    else {
        setTextValue("areaSparklinePoints", "");
    }
    var inputList = ["areaSparklineMinimumValue", "areaSparklineMaximumValue", "areaSparklineLine1", "areaSparklineLine2"];
    var len = inputList.length;
    for (var i = 1; i <= len; i++) {
        if (formulaArgs[i]) {
            setNumberValue(inputList[i - 1], unParseFormula(formulaArgs[i], row, col));
        }
        else {
            setNumberValue(inputList[i - 1], "");
        }
    }
    var positiveColor = parseColorExpression(formulaArgs[5], row, col);
    if (positiveColor) {
        setColorValue("areaSparklinePositiveColor", positiveColor);
    }
    else {
        setColorValue("areaSparklinePositiveColor", defaultValue.colorPositive);
    }
    var negativeColor = parseColorExpression(formulaArgs[6], row, col);
    if (negativeColor) {
        setColorValue("areaSparklineNegativeColor", negativeColor);
    }
    else {
        setColorValue("areaSparklineNegativeColor", defaultValue.colorNegative);
    }
}

function getBoxPlotSparklineSetting(formulaArgs, row, col) {
    var Calc = spreadNS.Calc;
    var defaultValue = {boxplotClass: "5ns", style: 0, colorScheme: "#D2D2D2", vertical: false, showAverage: false};
    if (formulaArgs && formulaArgs.length > 0) {
        var pointsValue = unParseFormula(formulaArgs[0], row, col);
        var boxPlotClassValue = formulaArgs[1] instanceof Calc.Expressions.StringExpression ? formulaArgs[1].value : null;
        var showAverageValue = formulaArgs[2] instanceof Calc.Expressions.BooleanExpression ? formulaArgs[2].value : null;
        var scaleStartValue = unParseFormula(formulaArgs[3], row, col);
        var scaleEndValue = unParseFormula(formulaArgs[4], row, col);
        var acceptableStartValue = unParseFormula(formulaArgs[5], row, col);
        var acceptableEndValue = unParseFormula(formulaArgs[6], row, col);
        var colorValue = parseColorExpression(formulaArgs[7], row, col);
        var styleValue = formulaArgs[8] ? unParseFormula(formulaArgs[8], row, col) : null;
        var verticalValue = formulaArgs[9] instanceof Calc.Expressions.BooleanExpression ? formulaArgs[9].value : null;

        setTextValue("boxplotSparklinePoints", pointsValue);
        setDropDown("boxplotClassType", boxPlotClassValue === null ? defaultValue.boxplotClass : boxPlotClassValue);
        setTextValue("boxplotSparklineScaleStart", scaleStartValue);
        setTextValue("boxplotSparklineScaleEnd", scaleEndValue);
        setTextValue("boxplotSparklineAcceptableStart", acceptableStartValue);
        setTextValue("boxplotSparklineAcceptableEnd", acceptableEndValue);
        setColorValue("boxplotSparklineColorScheme", colorValue === null ? defaultValue.colorScheme : colorValue);
        setDropDown("boxplotSparklineStyleType", styleValue === null ? defaultValue.style : styleValue);
        setCheckValue("boxplotSparklineShowAverage", showAverageValue === null ? defaultValue.showAverage : showAverageValue);
        setCheckValue("boxplotSparklineVertical", verticalValue === null ? defaultValue.vertical : verticalValue);
    }
    else {
        setTextValue("boxplotSparklinePoints", "");
        setDropDown("boxplotClassType", defaultValue.boxplotClass);
        setTextValue("boxplotSparklineScaleStart", "");
        setTextValue("boxplotSparklineScaleEnd", "");
        setTextValue("boxplotSparklineAcceptableStart", "");
        setTextValue("boxplotSparklineAcceptableEnd", "");
        setColorValue("boxplotSparklineColorScheme", defaultValue.colorScheme);
        setDropDown("boxplotSparklineStyleType", defaultValue.style);
        setCheckValue("boxplotSparklineShowAverage", defaultValue.showAverage);
        setCheckValue("boxplotSparklineVertical", defaultValue.vertical);
    }
}

function getBulletSparklineSetting(formulaArgs, row, col) {
    var Calc = spreadNS.Calc;
    var defaultValue = {vertical: false, colorScheme: "#A0A0A0"};
    if (formulaArgs && formulaArgs.length > 0) {
        var measureValue = unParseFormula(formulaArgs[0], row, col);
        var targetValue = unParseFormula(formulaArgs[1], row, col);
        var maxiValue = unParseFormula(formulaArgs[2], row, col);
        var goodValue = unParseFormula(formulaArgs[3], row, col);
        var badValue = unParseFormula(formulaArgs[4], row, col);
        var forecastValue = unParseFormula(formulaArgs[5], row, col);
        var tickunitValue = unParseFormula(formulaArgs[6], row, col);
        var colorSchemeValue = parseColorExpression(formulaArgs[7], row, col);
        var verticalValue = formulaArgs[8] instanceof Calc.Expressions.BooleanExpression ? formulaArgs[8].value : null;

        setTextValue("bulletSparklineMeasure", measureValue);
        setTextValue("bulletSparklineTarget", targetValue);
        setTextValue("bulletSparklineMaxi", maxiValue);
        setTextValue("bulletSparklineForecast", forecastValue);
        setTextValue("bulletSparklineGood", goodValue);
        setTextValue("bulletSparklineBad", badValue);
        setTextValue("bulletSparklineTickUnit", tickunitValue);
        setColorValue("bulletSparklineColorScheme", colorSchemeValue ? colorSchemeValue : defaultValue.colorScheme);
        setCheckValue("bulletSparklineVertical", verticalValue ? verticalValue : defaultValue.vertical);
    }
    else {
        setTextValue("bulletSparklineMeasure", "");
        setTextValue("bulletSparklineTarget", "");
        setTextValue("bulletSparklineMaxi", "");
        setTextValue("bulletSparklineForecast", "");
        setTextValue("bulletSparklineGood", "");
        setTextValue("bulletSparklineBad", "");
        setTextValue("bulletSparklineTickUnit", "");
        setColorValue("bulletSparklineColorScheme", defaultValue.colorScheme);
        setCheckValue("bulletSparklineVertical", defaultValue.vertical);
    }
}

function getCascadeSparklineSetting(formulaArgs, row, col) {
    var Calc = spreadNS.Calc;
    var defaultValue = {colorPositive: "#8CBF64", colorNegative: "#D6604D", vertical: false};

    if (formulaArgs && formulaArgs.length > 0) {
        var pointsRangeValue = unParseFormula(formulaArgs[0], row, col);
        var pointIndexValue = unParseFormula(formulaArgs[1], row, col);
        var labelsRangeValue = unParseFormula(formulaArgs[2], row, col);
        var minimumValue = unParseFormula(formulaArgs[3], row, col);
        var maximumValue = unParseFormula(formulaArgs[4], row, col);
        var colorPositiveValue = parseColorExpression(formulaArgs[5], row, col);
        var colorNegativeValue = parseColorExpression(formulaArgs[6], row, col);
        var verticalValue = formulaArgs[7] instanceof Calc.Expressions.BooleanExpression ? formulaArgs[7].value : null;

        setTextValue("cascadeSparklinePointsRange", pointsRangeValue);
        setTextValue("cascadeSparklinePointIndex", pointIndexValue);
        setTextValue("cascadeSparklineLabelsRange", labelsRangeValue);
        setTextValue("cascadeSparklineMinimum", minimumValue);
        setTextValue("cascadeSparklineMaximum", maximumValue);
        setColorValue("cascadeSparklinePositiveColor", colorPositiveValue ? colorPositiveValue : defaultValue.colorPositive);
        setColorValue("cascadeSparklineNegativeColor", colorNegativeValue ? colorNegativeValue : defaultValue.colorNegative);
        setCheckValue("cascadeSparklineVertical", verticalValue ? verticalValue : defaultValue.vertical);
    }
    else {
        setTextValue("cascadeSparklinePointsRange", "");
        setTextValue("cascadeSparklinePointIndex", "");
        setTextValue("cascadeSparklineLabelsRange", "");
        setTextValue("cascadeSparklineMinimum", "");
        setTextValue("cascadeSparklineMaximum", "");
        setColorValue("cascadeSparklinePositiveColor", defaultValue.colorPositive);
        setColorValue("cascadeSparklineNegativeColor", defaultValue.colorNegative);
        setCheckValue("cascadeSparklineVertical", defaultValue.vertical);
    }
}

function parseSetting(jsonSetting) {
    var setting = {}, inBracket = false, inProperty = true, property = "", value = "";
    if (jsonSetting) {
        jsonSetting = jsonSetting.substr(1, jsonSetting.length - 2);
        for (var i = 0, len = jsonSetting.length; i < len; i++) {
            var char = jsonSetting.charAt(i);
            if (char === ":") {
                inProperty = false;
            }
            else if (char === "," && !inBracket) {
                setting[property] = value;
                property = "";
                value = "";
                inProperty = true;
            }
            else if (char === "\'" || char === "\"") {
                // discard
            }
            else {
                if (char === "(") {
                    inBracket = true;
                }
                else if (char === ")") {
                    inBracket = false;
                }
                if (inProperty) {
                    property += char;
                }
                else {
                    value += char;
                }
            }
        }
        if (property) {
            setting[property] = value;
        }
        for (var p in setting) {
            var v = setting[p];
            if (v !== null && typeof (v) !== "undefined") {
                if (v.toUpperCase() === "TRUE") {
                    setting[p] = true;
                } else if (v.toUpperCase() === "FALSE") {
                    setting[p] = false;
                } else if (!isNaN(v) && isFinite(v)) {
                    setting[p] = parseFloat(v);
                }
            }
        }
    }
    return setting;
}

function updateManual(type, inputDataName) {
    var $manualInput = $("input[data-name='" + inputDataName + "']");
    var $manualDiv = $manualInput.parent();
    if (type !== "custom") {
        $manualInput.attr("disabled", "disabled");
        $manualDiv.addClass("manual-disable");
    }
    else {
        $manualInput.removeAttr("disabled");
        $manualDiv.removeClass("manual-disable");
    }
}

function updateStyleSetting(settings) {
    var defaultValue = {
        negativePoints: "#A52A2A", markers: "#244062", highPoint: "#0000FF",
        lowPoint: "#0000FF", firstPoint: "#95B3D7", lastPoint: "#95B3D7",
        series: "#244062", axis: "#000000"
    };
    setColorValue("compatibleSparklineNegativeColor", settings.negativeColor ? settings.negativeColor : defaultValue.negativePoints);
    setColorValue("compatibleSparklineMarkersColor", settings.markersColor ? settings.markersColor : defaultValue.markers);
    setColorValue("compatibleSparklineAxisColor", settings.axisColor ? settings.axisColor : defaultValue.axis);
    setColorValue("compatibleSparklineSeriesColor", settings.seriesColor ? settings.seriesColor : defaultValue.series);
    setColorValue("compatibleSparklineHighMarkerColor", settings.highMarkerColor ? settings.highMarkerColor : defaultValue.highPoint);
    setColorValue("compatibleSparklineLowMarkerColor", settings.lowMarkerColor ? settings.lowMarkerColor : defaultValue.lowPoint);
    setColorValue("compatibleSparklineFirstMarkerColor", settings.firstMarkerColor ? settings.firstMarkerColor : defaultValue.firstPoint);
    setColorValue("compatibleSparklineLastMarkerColor", settings.lastMarkerColor ? settings.lastMarkerColor : defaultValue.lastPoint);
    setTextValue("compatibleSparklineLastLineWeight", settings.lineWeight || settings.lw);
}

function updateSparklineSetting(setting) {
    if (!setting) {
        return;
    }
    var defaultSetting = {
        rightToLeft: false,
        displayHidden: false,
        displayXAxis: false,
        showFirst: false,
        showHigh: false,
        showLast: false,
        showLow: false,
        showNegative: false,
        showMarkers: false
    };

    setDropDown("emptyCellDisplayType", setting.displayEmptyCellsAs ? setting.displayEmptyCellsAs : -1);
    setCheckValue("showDataInHiddenRowOrColumn", setting.displayHidden ? setting.displayHidden : defaultSetting.displayHidden);
    setCheckValue("compatibleSparklineShowFirst", setting.showFirst ? setting.showFirst : defaultSetting.showFirst);
    setCheckValue("compatibleSparklineShowLast", setting.showLast ? setting.showLast : defaultSetting.showLast);
    setCheckValue("compatibleSparklineShowHigh", setting.showHigh ? setting.showHigh : defaultSetting.showHigh);
    setCheckValue("compatibleSparklineShowLow", setting.showLow ? setting.showLow : defaultSetting.showLow);
    setCheckValue("compatibleSparklineShowNegative", setting.showNegative ? setting.showNegative : defaultSetting.showNegative);
    setCheckValue("compatibleSparklineShowMarkers", setting.showMarkers ? setting.showMarkers : defaultSetting.showMarkers);
    var minAxisType = spreadNS.SparklineAxisMinMax[setting.minAxisType];
    setDropDown("minAxisType", minAxisType ? minAxisType : "");
    setTextValue("manualMin", setting.manualMin ? setting.manualMin : "");
    var maxAxisType = spreadNS.SparklineAxisMinMax[setting.maxAxisType];
    setDropDown("maxAxisType", maxAxisType ? maxAxisType : "");
    setTextValue("manualMax", setting.manualMax ? setting.manualMax : "");
    setCheckValue("rightToLeft", setting.rightToLeft ? setting.rightToLeft : defaultSetting.rightToLeft);
    setCheckValue("displayXAxis", setting.displayXAxis ? setting.displayXAxis : defaultSetting.displayXAxis);

    var type = getDropDownValue("minAxisType");
    updateManual(type, "manualMin");
    type = getDropDownValue("maxAxisType");
    updateManual(type, "manualMax");
}

function getCompatibleSparklineSetting(formulaArgs, row, col) {
    var Calc = spreadNS.Calc;
    var sparklineSetting = {};

    setTextValue("compatibleSparklineData", unParseFormula(formulaArgs[0], row, col));
    setDropDown("dataOrientationType", formulaArgs[1].value);
    if (formulaArgs[2]) {
        setTextValue("compatibleSparklineDateAxisData", unParseFormula(formulaArgs[2], row, col));
    }
    else {
        setTextValue("compatibleSparklineDateAxisData", "");
    }
    if (formulaArgs[3]) {
        setDropDown("dateAxisOrientationType", formulaArgs[3].value);
    }
    else {
        setDropDown("dateAxisOrientationType", -1);
    }
    var colorExpression = parseColorExpression(formulaArgs[4], row, col);
    if (colorExpression) {
        sparklineSetting = parseSetting(colorExpression);
    }
    updateSparklineSetting(sparklineSetting);
    updateStyleSetting(sparklineSetting);
}

function getScatterSparklineSetting(formulaArgs, row, col) {
    var Calc = spreadNS.Calc;
    var defaultValue = {
        tags: false,
        drawSymbol: true,
        drawLines: false,
        dash: false,
        color1: "#969696",
        color2: "#CB0000"
    };
    var inputList = ["scatterSparklinePoints1", "scatterSparklinePoints2", "scatterSparklineMinX", "scatterSparklineMaxX",
        "scatterSparklineMinY", "scatterSparklineMaxY", "scatterSparklineHLine", "scatterSparklineVLine",
        "scatterSparklineXMinZone", "scatterSparklineXMaxZone", "scatterSparklineYMinZone", "scatterSparklineYMaxZone"];
    for (var i = 0; i < inputList.length; i++) {
        var formula = "";
        if (formulaArgs[i]) {
            formula = unParseFormula(formulaArgs[i], row, col);
        }
        setTextValue(inputList[i], formula);
    }

    var color1 = parseColorExpression(formulaArgs[15], row, col);
    var color2 = parseColorExpression(formulaArgs[16], row, col);
    var tags = formulaArgs[12] instanceof Calc.Expressions.BooleanExpression ? formulaArgs[12].value : null;
    var drawSymbol = formulaArgs[13] instanceof Calc.Expressions.BooleanExpression ? formulaArgs[13].value : null;
    var drawLines = formulaArgs[14] instanceof Calc.Expressions.BooleanExpression ? formulaArgs[14].value : null;
    var dashLine = formulaArgs[17] instanceof Calc.Expressions.BooleanExpression ? formulaArgs[17].value : null;

    setColorValue("scatterSparklineColor1", (color1 !== null) ? color1 : defaultValue.color1);
    setColorValue("scatterSparklineColor2", (color2 !== null) ? color2 : defaultValue.color2);
    setCheckValue("scatterSparklineTags", tags !== null ? tags : defaultValue.tags);
    setCheckValue("scatterSparklineDrawSymbol", drawSymbol !== null ? drawSymbol : defaultValue.drawSymbol);
    setCheckValue("scatterSparklineDrawLines", drawLines !== null ? drawLines : defaultValue.drawLines);
    setCheckValue("scatterSparklineDashLine", dashLine !== null ? dashLine : defaultValue.dash);
}

function getHBarSparklineSetting(formulaArgs, row, col) {
    var defaultValue = {colorScheme: "#969696"};

    var value = unParseFormula(formulaArgs[0], row, col);
    var colorScheme = parseColorExpression(formulaArgs[1], row, col);

    setTextValue("hbarSparklineValue", value);
    setColorValue("hbarSparklineColorScheme", (colorScheme !== null) ? colorScheme : defaultValue.colorScheme);
}

function getVBarSparklineSetting(formulaArgs, row, col) {
    var defaultValue = {colorScheme: "#969696"};

    var value = unParseFormula(formulaArgs[0], row, col);
    var colorScheme = parseColorExpression(formulaArgs[1], row, col);

    setTextValue("vbarSparklineValue", value);
    setColorValue("vbarSparklineColorScheme", (colorScheme !== null) ? colorScheme : defaultValue.colorScheme);
}

function getParetoSparklineSetting(formulaArgs, row, col) {
    var Calc = spreadNS.Calc;
    var defaultValue = {label: 0, vertical: false};

    if (formulaArgs && formulaArgs.length > 0) {
        var pointsRangeValue = unParseFormula(formulaArgs[0], row, col);
        var pointIndexValue = unParseFormula(formulaArgs[1], row, col);
        var colorRangeValue = unParseFormula(formulaArgs[2], row, col);
        var targetValue = unParseFormula(formulaArgs[3], row, col);
        var target2Value = unParseFormula(formulaArgs[4], row, col);
        var highlightPositionValue = unParseFormula(formulaArgs[5], row, col);
        var labelValue = formulaArgs[6] instanceof Calc.Expressions.DoubleExpression ? formulaArgs[6].value : null;
        var verticalValue = formulaArgs[7] instanceof Calc.Expressions.BooleanExpression ? formulaArgs[7].value : null;

        setTextValue("paretoSparklinePoints", pointsRangeValue);
        setTextValue("paretoSparklinePointIndex", pointIndexValue);
        setTextValue("paretoSparklineColorRange", colorRangeValue);
        setTextValue("paretoSparklineHighlightPosition", highlightPositionValue);
        setTextValue("paretoSparklineTarget", targetValue);
        setTextValue("paretoSparklineTarget2", target2Value);
        setDropDown("paretoLabelType", labelValue === null ? defaultValue.label : labelValue);
        setCheckValue("paretoSparklineVertical", verticalValue === null ? defaultValue.vertical : verticalValue);
    }
    else {
        setTextValue("paretoSparklinePoints", "");
        setTextValue("paretoSparklinePointIndex", "");
        setTextValue("paretoSparklineColorRange", "");
        setTextValue("paretoSparklineHighlightPosition", "");
        setTextValue("paretoSparklineTarget", "");
        setTextValue("paretoSparklineTarget2", "");
        setDropDown("paretoLabelType", defaultValue.label);
        setCheckValue("paretoSparklineVertical", defaultValue.vertical);
    }
}

function getSpreadSparklineSetting(formulaArgs, row, col) {
    var Calc = spreadNS.Calc;
    var defaultValue = {showAverage: false, style: 4, colorScheme: "#646464", vertical: false};

    if (formulaArgs && formulaArgs.length > 0) {
        var pointsValue = unParseFormula(formulaArgs[0], row, col);
        var showAverageValue = formulaArgs[1] instanceof Calc.Expressions.BooleanExpression ? formulaArgs[1].value : null;
        var scaleStartValue = unParseFormula(formulaArgs[2], row, col);
        var scaleEndValue = unParseFormula(formulaArgs[3], row, col);
        var styleValue = formulaArgs[4] ? unParseFormula(formulaArgs[4], row, col) : null;
        var colorSchemeValue = parseColorExpression(formulaArgs[5], row, col);
        var verticalValue = formulaArgs[6] instanceof Calc.Expressions.BooleanExpression ? formulaArgs[6].value : null;

        setTextValue("spreadSparklinePoints", pointsValue);
        setCheckValue("spreadSparklineShowAverage", showAverageValue ? showAverageValue : defaultValue.showAverage);
        setTextValue("spreadSparklineScaleStart", scaleStartValue);
        setTextValue("spreadSparklineScaleEnd", scaleEndValue);
        setDropDown("spreadSparklineStyleType", styleValue ? styleValue : defaultValue.style);
        setColorValue("spreadSparklineColorScheme", colorSchemeValue ? colorSchemeValue : defaultValue.colorScheme);
        setCheckValue("spreadSparklineVertical", verticalValue ? verticalValue : defaultValue.vertical);
    }
    else {
        setTextValue("spreadSparklinePoints", "");
        setCheckValue("spreadSparklineShowAverage", defaultValue.showAverage);
        setTextValue("spreadSparklineScaleStart", "");
        setTextValue("spreadSparklineScaleEnd", "");
        setDropDown("spreadSparklineStyleType", defaultValue.style);
        setColorValue("spreadSparklineColorScheme", defaultValue.colorScheme);
        setCheckValue("spreadSparklineVertical", defaultValue.vertical);
    }
}

function getStackedSparklineSetting(formulaArgs, row, col) {
    var Calc = spreadNS.Calc;
    var defaultValue = {color: "#646464", vertical: false, textOrientation: 0};

    if (formulaArgs && formulaArgs.length > 0) {
        var pointsValue = unParseFormula(formulaArgs[0], row, col);
        var colorRangeValue = unParseFormula(formulaArgs[1], row, col);
        var labelRangeValue = unParseFormula(formulaArgs[2], row, col);
        var maximumValue = unParseFormula(formulaArgs[3], row, col);
        var targetRedValue = unParseFormula(formulaArgs[4], row, col);
        var targetGreenValue = unParseFormula(formulaArgs[5], row, col);
        var targetBlueValue = unParseFormula(formulaArgs[6], row, col);
        var targetYellowValue = unParseFormula(formulaArgs[7], row, col);
        var colorValue = parseColorExpression(formulaArgs[8], row, col);
        var highlightPositionValue = unParseFormula(formulaArgs[9], row, col);
        var verticalValue = formulaArgs[10] instanceof Calc.Expressions.BooleanExpression ? formulaArgs[10].value : null;
        var textOrientationValue = unParseFormula(formulaArgs[11], row, col);
        var textSizeValue = unParseFormula(formulaArgs[12], row, col);

        setTextValue("stackedSparklinePoints", pointsValue);
        setTextValue("stackedSparklineColorRange", colorRangeValue);
        setTextValue("stackedSparklineLabelRange", labelRangeValue);
        setNumberValue("stackedSparklineMaximum", maximumValue);
        setNumberValue("stackedSparklineTargetRed", targetRedValue);
        setNumberValue("stackedSparklineTargetGreen", targetGreenValue);
        setNumberValue("stackedSparklineTargetBlue", targetBlueValue);
        setNumberValue("stackedSparklineTargetYellow", targetYellowValue);
        setColorValue("stackedSparklineColor", "stacked-color-span", colorValue ? colorValue : defaultValue.color);
        setNumberValue("stackedSparklineHighlightPosition", highlightPositionValue);
        setCheckValue("stackedSparklineVertical", verticalValue ? verticalValue : defaultValue.vertical);
        setDropDown("stackedSparklineTextOrientation", textOrientationValue ? textOrientationValue : defaultValue.textOrientation);
        setNumberValue("stackedSparklineTextSize", textSizeValue);
    }
    else {
        setTextValue("stackedSparklinePoints", "");
        setTextValue("stackedSparklineColorRange", "");
        setTextValue("stackedSparklineLabelRange", "");
        setNumberValue("stackedSparklineMaximum", "");
        setNumberValue("stackedSparklineTargetRed", "");
        setNumberValue("stackedSparklineTargetGreen", "");
        setNumberValue("stackedSparklineTargetBlue", "");
        setNumberValue("stackedSparklineTargetYellow", "");
        setColorValue("stackedSparklineColor", "stacked-color-span", defaultValue.color);
        setNumberValue("stackedSparklineHighlightPosition", "");
        setCheckValue("stackedSparklineVertical", defaultValue.vertical);
        setDropDown("stackedSparklineTextOrientation", defaultValue.textOrientation);
        setNumberValue("stackedSparklineTextSize", "");
    }
}

function getVariSparklineSetting(formulaArgs, row, col) {
    var Calc = spreadNS.Calc;
    var defaultValue = {legend: false, colorPositive: "green", colorNegative: "red", vertical: false};

    if (formulaArgs && formulaArgs.length > 0) {
        var varianceValue = unParseFormula(formulaArgs[0], row, col);
        var referenceValue = unParseFormula(formulaArgs[1], row, col);
        var miniValue = unParseFormula(formulaArgs[2], row, col);
        var maxiValue = unParseFormula(formulaArgs[3], row, col);
        var markValue = unParseFormula(formulaArgs[4], row, col);
        var tickunitValue = unParseFormula(formulaArgs[5], row, col);
        var legendValue = formulaArgs[6] instanceof Calc.Expressions.BooleanExpression ? formulaArgs[6].value : null;
        var colorPositiveValue = parseColorExpression(formulaArgs[7], row, col);
        var colorNegativeValue = parseColorExpression(formulaArgs[8], row, col);
        var verticalValue = formulaArgs[9] instanceof Calc.Expressions.BooleanExpression ? formulaArgs[9].value : null;

        setTextValue("variSparklineVariance", varianceValue);
        setTextValue("variSparklineReference", referenceValue);
        setTextValue("variSparklineMini", miniValue);
        setTextValue("variSparklineMaxi", maxiValue);
        setTextValue("variSparklineMark", markValue);
        setTextValue("variSparklineTickUnit", tickunitValue);
        setColorValue("variSparklineColorPositive", colorPositiveValue ? colorPositiveValue : defaultValue.colorPositive);
        setColorValue("variSparklineColorNegative", colorNegativeValue ? colorNegativeValue : defaultValue.colorNegative);
        setCheckValue("variSparklineLegend", legendValue);
        setCheckValue("variSparklineVertical", verticalValue);
    }
    else {
        setTextValue("variSparklineVariance", "");
        setTextValue("variSparklineReference", "");
        setTextValue("variSparklineMini", "");
        setTextValue("variSparklineMaxi", "");
        setTextValue("variSparklineMark", "");
        setTextValue("variSparklineTickUnit", "");
        setColorValue("variSparklineColorPositive", defaultValue.colorPositive);
        setColorValue("variSparklineColorNegative", defaultValue.colorNegative);
        setCheckValue("variSparklineLegend", defaultValue.legend);
        setCheckValue("variSparklineVertical", defaultValue.vertical);
    }
}

function addPieSparklineColor(count, color, isMinusSymbol) {
    var defaultColor = "rgb(237, 237, 237)";
    color = color ? color : defaultColor;
    var symbolFunClass, symbolClass;
    if (isMinusSymbol) {
        symbolFunClass = "remove-pie-color";
        symbolClass = "ui-pie-sparkline-icon-minus";
    }
    else {
        symbolFunClass = "add-pie-color";
        symbolClass = "ui-pie-sparkline-icon-plus";
    }
    var $pieSparklineColorContainer = $("#pieSparklineColorContainer");
    var pieColorDataName = "pieColorName";

    var $colorDiv = $("#pie-color-item-template").children().clone();
    $('label', $colorDiv).text(uiResource.sparklineExTab.pieSparkline.values.color + count);
    $('.pane-color-picker', $colorDiv).attr('data-name', pieColorDataName + count);
    $('.color-picker', $colorDiv).css('background-color', color);
    $('.insp-inline-row-item', $colorDiv).addClass(symbolFunClass);
    $('.ui-pie-sparkline-icon', $colorDiv).addClass(symbolClass);

    $colorDiv.appendTo($pieSparklineColorContainer);
}

function addPieColor(count, color, isMinusSymbol) {
    var $colorSpanDiv = $(".add-pie-color");
    $colorSpanDiv.addClass("remove-pie-color").removeClass("add-pie-color");
    $colorSpanDiv.find("span").addClass("ui-pie-sparkline-icon-minus").removeClass("ui-pie-sparkline-icon-plus");
    addPieSparklineColor(count, color, isMinusSymbol);
    $(".add-pie-color").unbind("click");
    $(".add-pie-color").bind("click", function (evt) {
        var count = $("#pieSparklineColorContainer").find("span").length;
        addPieColor(count + 1);
    });
    $(".remove-pie-color").unbind("click");
    $(".remove-pie-color").bind("click", function (evt) {
        resetPieColor($(evt.target));
    });
}

function resetPieColor($colorSpanDiv) {
    if (!$colorSpanDiv.hasClass("ui-pie-sparkline-icon")) {
        return;
    }
    var $colorDiv = $colorSpanDiv.closest('div.pane-row');
    $colorDiv.remove();
    var $pieSparklineColorContainer = $("#pieSparklineColorContainer");
    var colorArray = [];
    $pieSparklineColorContainer.find(".color-picker").each(function () {
        colorArray.push($(this).css("background-color"));
    });
    $pieSparklineColorContainer.empty();
    addMultiPieColor(colorArray);
}

function addMultiPieColor(colorArray) {
    if (!colorArray || colorArray.length === 0) {
        return;
    }
    var length = colorArray.length;
    var i = 0;
    for (i; i < length - 1; i++) {
        addPieSparklineColor(i + 1, colorArray[i], true);
    }
    addPieColor(i + 1, colorArray[i]);
}

function getPieSparklineSetting(formulaArgs, row, col) {
    var Calc = spreadNS.Calc;
    var defaultValue = {legend: false, colorPositive: "green", colorNegative: "red", vertical: false};

    var agrsLength = formulaArgs.length;
    if (formulaArgs && agrsLength > 0) {
        var range = unParseFormula(formulaArgs[0], row, col);
        setTextValue("pieSparklinePercentage", range);

        var actualLen = agrsLength - 1;
        if (actualLen === 0) {
            addPieColor(1);
        }
        else {
            var colorArray = [];
            for (var i = 1; i <= actualLen; i++) {
                var colorItem = null;
                var color = parseColorExpression(formulaArgs[i], row, col);
                colorArray.push(color);
            }
            addMultiPieColor(colorArray);
        }
    }
}

var sparklineName;
function showSparklineSetting(row, col) {
    var expr = parseFormulaSparkline(row, col);
    if (!expr || !expr.args) {
        return false;
    }
    var formulaSparkline = spread.getSparklineEx(expr.getFunctionName());

    if (formulaSparkline) {
        var $sparklineSettingDiv = $("#sparklineDetailSetting>div");
        var formulaArgs = expr.args;
        $sparklineSettingDiv.hide();
        if (formulaSparkline instanceof spreadNS.PieSparkline) {
            $("#pieSparklineSetting").show();
            $("#pieSparklineColorContainer").empty();
            getPieSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof spreadNS.AreaSparkline) {
            $("#areaSparklineSetting").show();
            getAreaSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof spreadNS.BoxPlotSparkline) {
            $("#boxplotSparklineSetting").show();
            getBoxPlotSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof spreadNS.BulletSparkline) {
            $("#bulletSparklineSetting").show();
            getBulletSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof spreadNS.CascadeSparkline) {
            $("#cascadeSparklineSetting").show();
            getCascadeSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof spreadNS.CompatibleSparkline) {
            $("#compatibleSparklineSetting").show();
            if (expr.fn.name) {
                sparklineName = expr.fn.name;
            }
            getCompatibleSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof spreadNS.ScatterSparkline) {
            $("#scatterSparklineSetting").show();
            getScatterSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof spreadNS.HBarSparkline) {
            $("#hbarSparklineSetting").show();
            getHBarSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof spreadNS.VBarSparkline) {
            $("#vbarSparklineSetting").show();
            getVBarSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof spreadNS.ParetoSparkline) {
            $("#paretoSparklineSetting").show();
            getParetoSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof spreadNS.SpreadSparkline) {
            $("#spreadSparklineSetting").show();
            getSpreadSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof spreadNS.StackedSparkline) {
            $("#stackedSparklineSetting").show();
            getStackedSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof spreadNS.VariSparkline) {
            $("#variSparklineSetting").show();
            getVariSparklineSetting(formulaArgs, row, col);
            return true;
        }
    }
    return false;
}

function attachSparklineSettingEvents() {
    $("#setAreaSparkline").click(applyAreaSparklineSetting);
    $("#setBoxPlotSparkline").click(applyBoxPlotSparklineSetting);
    $("#setBulletSparkline").click(applyBulletSparklineSetting);
    $("#setCascadeSparkline").click(applyCascadeSparklineSetting);
    $("#setCompatibleSparkline").click(applyCompatibleSparklineSetting);
    $("#setScatterSparkline").click(applyScatterSparklineSetting);
    $("#setHbarSparkline").click(applyHbarSparklineSetting);
    $("#setVbarSparkline").click(applyVbarSparklineSetting);
    $("#setParetoSparkline").click(applyParetoSparklineSetting);
    $("#setSpreadSparkline").click(applySpreadSparklineSetting);
    $("#setStackedSparkline").click(applyStackedSparklineSetting);
    $("#setVariSparkline").click(applyVariSparklineSetting);
    $("#setPieSparkline").click(applyPieSparklineSetting);
}

function updateFormulaBar() {
    var sheet = spread.getActiveSheet();
    var formulaBar = $("#formulabox");
    if (formulaBar.length > 0) {
        var formula = sheet.getFormula(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex());
        if (formula) {
            formula = "=" + formula;
            formulaBar.text(formula);
        }
    }
}

function removeContinuousComma(parameter) {
    var len = parameter.length;
    while (len > 0 && parameter[len - 1] === ",") {
        len--;
    }
    return parameter.substr(0, len);
}

function formatFormula(paraArray) {
    var params = "";
    for (var i = 0; i < paraArray.length; i++) {
        var item = paraArray[i];
        if (item !== undefined && item !== null) {
            params += item + ",";
        }
        else {
            params += ",";
        }
    }
    params = removeContinuousComma(params);
    return params;
}

function getFormula(params) {
    var len = params.length;
    while (len > 0 && params[len - 1] === "") {
        len--;
    }
    var temp = "";
    for (var i = 0; i < len; i++) {
        temp += params[i];
        if (i !== len - 1) {
            temp += ",";
        }
    }
    return "=AREASPARKLINE(" + temp + ")";
}

function setFormulaSparkline(formula) {
    var sheet = spread.getActiveSheet();
    var row = sheet.getActiveRowIndex();
    var col = sheet.getActiveColumnIndex();
    if (formula) {
        Actions.setFormulaSparkline(spread, {row: row, col: col, formula: formula});
    }
}

function applyAreaSparklineSetting() {
    var points = getTextValue("areaSparklinePoints");
    var mini = getNumberValue("areaSparklineMinimumValue");
    var maxi = getNumberValue("areaSparklineMaximumValue");
    var line1 = getNumberValue("areaSparklineLine1");
    var line2 = getNumberValue("areaSparklineLine2");
    var colorPositive = "\"" + getBackgroundColor("areaSparklinePositiveColor") + "\"";
    var colorNegative = "\"" + getBackgroundColor("areaSparklineNegativeColor") + "\"";
    var paramArr = [points, mini, maxi, line1, line2, colorPositive, colorNegative];
    var formula = getFormula(paramArr);

    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyBoxPlotSparklineSetting() {
    var pointsValue = getTextValue("boxplotSparklinePoints");
    var boxPlotClassValue = getDropDownValue("boxplotClassType");
    var showAverageValue = getCheckValue("boxplotSparklineShowAverage");
    var scaleStartValue = getTextValue("boxplotSparklineScaleStart");
    var scaleEndValue = getTextValue("boxplotSparklineScaleEnd");
    var acceptableStartValue = getTextValue("boxplotSparklineAcceptableStart");
    var acceptableEndValue = getTextValue("boxplotSparklineAcceptableEnd");
    var colorValue = getBackgroundColor("boxplotSparklineColorScheme");
    var styleValue = getDropDownValue("boxplotSparklineStyleType");
    var verticalValue = getCheckValue("boxplotSparklineVertical");

    var boxplotClassStr = boxPlotClassValue ? "\"" + boxPlotClassValue + "\"" : null;
    var colorStr = colorValue ? "\"" + colorValue + "\"" : null;
    var paraPool = [
        pointsValue,
        boxplotClassStr,
        showAverageValue,
        scaleStartValue,
        scaleEndValue,
        acceptableStartValue,
        acceptableEndValue,
        colorStr,
        styleValue,
        verticalValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=BOXPLOTSPARKLINE(" + params + ")";

    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyBulletSparklineSetting() {
    var measureValue = getTextValue("bulletSparklineMeasure");
    var targetValue = getTextValue("bulletSparklineTarget");
    var maxiValue = getTextValue("bulletSparklineMaxi");
    var goodValue = getTextValue("bulletSparklineGood");
    var badValue = getTextValue("bulletSparklineBad");
    var forecastValue = getTextValue("bulletSparklineForecast");
    var tickunitValue = getTextValue("bulletSparklineTickUnit");
    var colorSchemeValue = getBackgroundColor("bulletSparklineColorScheme");
    var verticalValue = getCheckValue("bulletSparklineVertical");

    var colorSchemeString = colorSchemeValue ? "\"" + colorSchemeValue + "\"" : null;
    var paraPool = [
        measureValue,
        targetValue,
        maxiValue,
        goodValue,
        badValue,
        forecastValue,
        tickunitValue,
        colorSchemeString,
        verticalValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=BULLETSPARKLINE(" + params + ")";
    var sheet = spread.getActiveSheet();
    var sels = sheet.getSelections();

    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyCascadeSparklineSetting() {
    var pointsRangeValue = getTextValue("cascadeSparklinePointsRange");
    var pointIndexValue = getTextValue("cascadeSparklinePointIndex");
    var labelsRangeValue = getTextValue("cascadeSparklineLabelsRange");
    var minimumValue = getTextValue("cascadeSparklineMinimum");
    var maximumValue = getTextValue("cascadeSparklineMaximum");
    var colorPositiveValue = getBackgroundColor("cascadeSparklinePositiveColor");
    var colorNegativeValue = getBackgroundColor("cascadeSparklineNegativeColor");
    var verticalValue = getCheckValue("cascadeSparklineVertical");
    var colorPositiveStr = colorPositiveValue ? "\"" + colorPositiveValue + "\"" : null;
    var colorNegativeStr = colorNegativeValue ? "\"" + colorNegativeValue + "\"" : null;
    var paraPool = [
        pointsRangeValue,
        pointIndexValue,
        labelsRangeValue,
        minimumValue,
        maximumValue,
        colorPositiveStr,
        colorNegativeStr,
        verticalValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=CASCADESPARKLINE(" + params + ")";

    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyCompatibleSparklineSetting() {
    var data = getTextValue("compatibleSparklineData");
    var dataOrientation = getDropDownValue("dataOrientationType");
    var dateAxisData = getTextValue("compatibleSparklineDateAxisData");
    var dateAxisOrientation = getDropDownValue("dateAxisOrientationType");
    if (dateAxisOrientation === undefined) {
        dateAxisOrientation = "";
    }
    var sparklineSetting = {}, minAxisType, maxAxisType;
    
    sparklineSetting.displayEmptyCellsAs = getDropDownValue("emptyCellDisplayType");
    sparklineSetting.displayHidden = getCheckValue("showDataInHiddenRowOrColumn");
    sparklineSetting.showFirst = getCheckValue("compatibleSparklineShowFirst");
    sparklineSetting.showLast = getCheckValue("compatibleSparklineShowLast");
    sparklineSetting.showHigh = getCheckValue("compatibleSparklineShowHigh");
    sparklineSetting.showLow = getCheckValue("compatibleSparklineShowLow");
    sparklineSetting.showNegative = getCheckValue("compatibleSparklineShowNegative");
    sparklineSetting.showMarkers = getCheckValue("compatibleSparklineShowMarkers");
    minAxisType = getDropDownValue("minAxisType");
    sparklineSetting.minAxisType = spreadNS.SparklineAxisMinMax[minAxisType];
    sparklineSetting.manualMin = getTextValue("manualMin");
    maxAxisType = getDropDownValue("maxAxisType");
    sparklineSetting.maxAxisType = spreadNS.SparklineAxisMinMax[maxAxisType];
    sparklineSetting.manualMax = getTextValue("manualMax");
    sparklineSetting.rightToLeft = getCheckValue("rightToLeft");
    sparklineSetting.displayXAxis = getCheckValue("displayXAxis");

    sparklineSetting.negativeColor = getBackgroundColor("compatibleSparklineNegativeColor");
    sparklineSetting.markersColor = getBackgroundColor("compatibleSparklineMarkersColor");
    sparklineSetting.axisColor = getBackgroundColor("compatibleSparklineAxisColor");
    sparklineSetting.seriesColor = getBackgroundColor("compatibleSparklineSeriesColor");
    sparklineSetting.highMarkerColor = getBackgroundColor("compatibleSparklineHighMarkerColor");
    sparklineSetting.lowMarkerColor = getBackgroundColor("compatibleSparklineLowMarkerColor");
    sparklineSetting.firstMarkerColor = getBackgroundColor("compatibleSparklineFirstMarkerColor");
    sparklineSetting.lastMarkerColor = getBackgroundColor("compatibleSparklineLastMarkerColor");
    sparklineSetting.lineWeight = getTextValue("compatibleSparklineLastLineWeight");

    var settingArray = [];
    for (var item in sparklineSetting) {
        if (sparklineSetting[item] !== undefined && sparklineSetting[item] !== "") {
            settingArray.push(item + ":" + sparklineSetting[item]);
        }
    }
    var settingString = "";
    if (settingArray.length > 0) {
        settingString = "\"{" + settingArray.join(",") + "}\"";
    }

    var formula = "";
    if (settingString !== "") {
        formula = "=" + sparklineName + "(" + data + "," + dataOrientation +
            "," + dateAxisData + "," + dateAxisOrientation + "," + settingString + ")";
    }
    else {
        if (dateAxisOrientation !== "") {
            formula = "=" + sparklineName + "(" + data + "," + dataOrientation +
                "," + dateAxisData + "," + dateAxisOrientation + ")";
        }
        else {
            if (dateAxisData !== "") {
                formula = "=" + sparklineName + "(" + data + "," + dataOrientation +
                    "," + dateAxisData + ")";
            }
            else {
                formula = "=" + sparklineName + "(" + data + "," + dataOrientation + ")";
            }
        }
    }

    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyScatterSparklineSetting() {
    var paraPool = [];
    var inputList = ["scatterSparklinePoints1", "scatterSparklinePoints2", "scatterSparklineMinX", "scatterSparklineMaxX",
        "scatterSparklineMinY", "scatterSparklineMaxY", "scatterSparklineHLine", "scatterSparklineVLine",
        "scatterSparklineXMinZone", "scatterSparklineXMaxZone", "scatterSparklineYMinZone", "scatterSparklineYMaxZone"];
    for (var i = 0; i < inputList.length; i++) {
        var textValue = getTextValue(inputList[i]);
        paraPool.push(textValue);
    }
    var tags = getCheckValue("scatterSparklineTags");
    var drawSymbol = getCheckValue("scatterSparklineDrawSymbol");
    var drawLines = getCheckValue("scatterSparklineDrawLines");
    var color1 = getBackgroundColor("scatterSparklineColor1");
    var color2 = getBackgroundColor("scatterSparklineColor2");
    var dashLine = getCheckValue("scatterSparklineDashLine");

    color1 = color1 ? "\"" + color1 + "\"" : null;
    color2 = color2 ? "\"" + color2 + "\"" : null;

    paraPool.push(tags);
    paraPool.push(drawSymbol);
    paraPool.push(drawLines);
    paraPool.push(color1);
    paraPool.push(color2);
    paraPool.push(dashLine);

    var params = formatFormula(paraPool);
    var formula = "=SCATTERSPARKLINE(" + params + ")";

    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyHbarSparklineSetting() {
    var paraPool = [];
    var value = getTextValue("hbarSparklineValue");
    var colorScheme = getBackgroundColor("hbarSparklineColorScheme");

    colorScheme = "\"" + colorScheme + "\"";
    paraPool.push(value);
    paraPool.push(colorScheme);

    var params = formatFormula(paraPool);
    var formula = "=HBARSPARKLINE(" + params + ")";

    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyVbarSparklineSetting() {
    var paraPool = [];
    var value = getTextValue("vbarSparklineValue");
    var colorScheme = getBackgroundColor("vbarSparklineColorScheme");

    colorScheme = "\"" + colorScheme + "\"";
    paraPool.push(value);
    paraPool.push(colorScheme);

    var params = formatFormula(paraPool);
    var formula = "=VBARSPARKLINE(" + params + ")";

    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyParetoSparklineSetting() {
    var pointsRangeValue = getTextValue("paretoSparklinePoints");
    var pointIndexValue = getTextValue("paretoSparklinePointIndex");
    var colorRangeValue = getTextValue("paretoSparklineColorRange");
    var targetValue = getTextValue("paretoSparklineTarget");
    var target2Value = getTextValue("paretoSparklineTarget2");
    var highlightPositionValue = getTextValue("paretoSparklineHighlightPosition");
    var labelValue = getDropDownValue("paretoLabelType");
    var verticalValue = getCheckValue("paretoSparklineVertical");
    var paraPool = [
        pointsRangeValue,
        pointIndexValue,
        colorRangeValue,
        targetValue,
        target2Value,
        highlightPositionValue,
        labelValue,
        verticalValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=PARETOSPARKLINE(" + params + ")";

    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applySpreadSparklineSetting() {
    var pointsValue = getTextValue("spreadSparklinePoints");
    var showAverageValue = getCheckValue("spreadSparklineShowAverage");
    var scaleStartValue = getTextValue("spreadSparklineScaleStart");
    var scaleEndValue = getTextValue("spreadSparklineScaleEnd");
    var styleValue = getDropDownValue("spreadSparklineStyleType");
    var colorSchemeValue = getBackgroundColor("spreadSparklineColorScheme");
    var verticalValue = getCheckValue("spreadSparklineVertical");
    var colorSchemeString = colorSchemeValue ? "\"" + colorSchemeValue + "\"" : null;
    var paraPool = [
        pointsValue,
        showAverageValue,
        scaleStartValue,
        scaleEndValue,
        styleValue,
        colorSchemeString,
        verticalValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=SPREADSPARKLINE(" + params + ")";

    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyStackedSparklineSetting() {
    var pointsValue = getTextValue("stackedSparklinePoints");
    var colorRangeValue = getTextValue("stackedSparklineColorRange");
    var labelRangeValue = getTextValue("stackedSparklineLabelRange");
    var maximumValue = getNumberValue("stackedSparklineMaximum");
    var targetRedValue = getNumberValue("stackedSparklineTargetRed");
    var targetGreenValue = getNumberValue("stackedSparklineTargetGreen");
    var targetBlueValue = getNumberValue("stackedSparklineTargetBlue");
    var targetYellowValue = getNumberValue("stackedSparklineTargetYellow");
    var colorValue = getBackgroundColor("stackedSparklineColor");
    var highlightPositionValue = getNumberValue("stackedSparklineHighlightPosition");
    var verticalValue = getCheckValue("stackedSparklineVertical");
    var textOrientationValue = getDropDownValue("stackedSparklineTextOrientation");
    var textSizeValue = getNumberValue("stackedSparklineTextSize");
    var colorString = colorValue ? "\"" + colorValue + "\"" : null;
    var paraPool = [
        pointsValue,
        colorRangeValue,
        labelRangeValue,
        maximumValue,
        targetRedValue,
        targetGreenValue,
        targetBlueValue,
        targetYellowValue,
        colorString,
        highlightPositionValue,
        verticalValue,
        textOrientationValue,
        textSizeValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=STACKEDSPARKLINE(" + params + ")";

    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyVariSparklineSetting() {
    var varianceValue = getTextValue("variSparklineVariance");
    var referenceValue = getTextValue("variSparklineReference");
    var miniValue = getTextValue("variSparklineMini");
    var maxiValue = getTextValue("variSparklineMaxi");
    var markValue = getTextValue("variSparklineMark");
    var tickunitValue = getTextValue("variSparklineTickUnit");
    var colorPositiveValue = getBackgroundColor("variSparklineColorPositive");
    var colorNegativeValue = getBackgroundColor("variSparklineColorNegative");
    var legendValue = getCheckValue("variSparklineLegend");
    var verticalValue = getCheckValue("variSparklineVertical");
    var colorPositiveStr = colorPositiveValue ? "\"" + colorPositiveValue + "\"" : null;
    var colorNegativeStr = colorNegativeValue ? "\"" + colorNegativeValue + "\"" : null;
    var paraPool = [
        varianceValue,
        referenceValue,
        miniValue,
        maxiValue,
        markValue,
        tickunitValue,
        legendValue,
        colorPositiveStr,
        colorNegativeStr,
        verticalValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=VARISPARKLINE(" + params + ")";

    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyPieSparklineSetting() {
    var paraPool = [];
    var range = getTextValue("pieSparklinePercentage");
    paraPool.push(range);

    $("#pieSparklineColorContainer").find(".color-picker").each(function () {
        var color = "\"" + $(this).css("background-color") + "\"";
        paraPool.push(color);
    });

    var params = formatFormula(paraPool);
    var formula = "=PIESPARKLINE(" + params + ")";

    setFormulaSparkline(formula);
    updateFormulaBar();
}
// Sparkline related items (end)

// Zoom related items
function processZoomSetting(value, title) {
    if (typeof value === 'number') {
        spread.getActiveSheet().zoom(value);
    }
    else {
        console.log("processZoomSetting not process with ", value, title);
    }
}
// Zoom related items (end)

function getResource(key) {
    key = key.replace(/\./g, "_");

    return resourceMap[key];
}

function getResourceMap(src) {
    function isObject(item) {
        return typeof item === "object";
    }

    function addResourceMap(map, obj, keys) {
        if (isObject(obj)) {
            for (var p in obj) {
                var cur = obj[p];

                addResourceMap(map, cur, keys.concat(p));
            }
        } else {
            var key = keys.join("_");
            map[key] = obj;
        }
    }

    addResourceMap(resourceMap, src, []);
}

$(document).ready(function () {
    if (!String.prototype.startsWith) {
        String.prototype.startsWith = function (searchString, position) {
            position = position || 0;
            return this.substr(position, searchString.length) === searchString;
        };
    }

    function localizeUI() {
        function getLocalizeString(text) {
            var matchs = text.match(/(?:(@[\w\d\.]*@))/g);

            if (matchs) {
                matchs.forEach(function (item) {
                    var s = getResource(item.replace(/[@]/g, ""));
                    text = text.replace(item, s);
                });
            }

            return text;
        }

        $(".localize").each(function () {
            var text = $(this).text();

            $(this).text(getLocalizeString(text));
        });

        $(".localize-tooltip").each(function () {
            var text = $(this).prop("title");

            $(this).prop("title", getLocalizeString(text));
        });

        $(".localize-value").each(function () {
            var text = $(this).attr("value");

            $(this).attr("value", getLocalizeString(text));
        });
    }
    
    function setDropdownDefault() {
        $('.btn-group ul.dropdown-menu li.default').each(function(){
            var $a = $('a', this),
                $group = $(this).parents('.btn-group'); 
                $span = $('span.content', $group);

            $group.attr('data-value', $a.attr('data-value'));                
            $span.text($a.text());
            $(this).addClass('selected');
        });
    }

    getResourceMap(uiResource);

    localizeUI();
    $colorPickerContainer = $("#colorPicker");
    $(".themeColorsLabel", $colorPickerContainer).text(uiResource.colorPicker.themeColors);
    $(".standardColorsLabel", $colorPickerContainer).text(uiResource.colorPicker.standardColors);

    $colorPickerContainer.on("click", ignoreEvent);

    colorPicker = new ColorPicker($colorPickerContainer);
    initRibbon(colorPicker);

    prepareFunctionBuilder($("#functionBuiilder"));
    prepareCellFormatSetting($("#cellFormatSetting"));

    setDropdownDefault();

    spread = new spreadNS.Spread($("#ss")[0]);
    //formulabox
    fbx = new spreadNS.FormulaTextBox(document.getElementById('formulabox'));
    var themeColors = getThemeColor();
    colorPicker.setThemeColors(themeColors);
    setTimeout(function () {
        initSpread();
    }, 50);

    //window resize adjust
    var resizeTimeout = null;
    $(window).bind("resize", function () {
        if (resizeTimeout === null) {
            resizeTimeout = setTimeout(function () {
                screenAdoption();
                clearTimeout(resizeTimeout);
                resizeTimeout = null;
            }, 100);
        }
    });

    doPrepareWork();

    var toolbarHeight = $("#toolbar").height(),
        formulaboxDefaultHeight = $("#formulabox").outerHeight(true),
        verticalSplitterOriginalTop = formulaboxDefaultHeight - $("#verticalSplitter").height();
    $("#verticalSplitter").draggable({
        axis: "y",              // vertical only
        containment: "#inner-content-container",  // limit in specified range
        scroll: false,          // not allow container scroll
        zIndex: 100,            // set to move on top
        stop: function (event, ui) {
            var $this = $(this),
                top = $this.offset().top,
                offset = top - toolbarHeight - verticalSplitterOriginalTop;

            // limit min size
            if (offset < 0) {
                offset = 0;
            }
            // adjust size of related items
            $("#formulabox").css({height: formulaboxDefaultHeight + offset});
            var height = $("div.insp-container").height() - $("#formulabox").outerHeight(true);
            $("#controlPanel").height(height);
            $("#ss").height(height);
            spread.refresh();
            // reset
            $(this).css({top: 0});
        }
    });

    attachEvents();

    $(document).on("contextmenu", ignoreEvent);
    
    $(window).resize();

    spread.focus();

    syncSheetPropertyValues();

    onCellSelected();

    updatePositionBox(spread.getActiveSheet());
});

function ignoreEvent(event) {
    event.preventDefault();
    return false;
}

// context menu related items
function getCellInSelections(selections, row, col) {
    var count = selections.length, range;
    for (var i = 0; i < count; i++) {
        range = selections[i];
        if (range.contains(row, col)) {
            return range;
        }
    }
    return null;
}
function getHitTest(pageX, pageY, sheet) {
    var offset = $("#ss").offset(),
        x = pageX - offset.left,
        y = pageY - offset.top;
    return sheet.hitTest(x, y);
}
function showMergeContextMenu() {
    // use the result of updateMergeButtonsState
    if (_mergeState) {
        if (!_mergeState.mergable) {
            $(".context-merge").hide();
        } else {
            $(".context-cell.divider").show();
            $(".context-merge").show();
        }

        if (!_mergeState.unmergable) {
            $(".context-unmerge").hide();
        } else {
            $(".context-cell.divider").show();
            $(".context-unmerge").show();
        }
    }
}

function processSpreadContextMenu(e) {
    // move the context menu to the position of the mouse point
    var sheet = spread.getActiveSheet(),
        target = getHitTest(e.pageX, e.pageY, sheet),
        hitTestType = target.hitTestType,
        row = target.row,
        col = target.col,
        selections = sheet.getSelections();

    var isHideContextMenu = false;

    if (hitTestType === GcSpread.Sheets.SheetArea.colHeader) {
        if (getCellInSelections(selections, row, col) === null) {
            sheet.setSelection(-1, col, sheet.getRowCount(), 1);
        }
        if (row !== undefined && col !== undefined) {
            $(".context-header").show();
            $(".context-cell").hide();
        }
    } else if (hitTestType === GcSpread.Sheets.SheetArea.rowHeader) {
        if (getCellInSelections(selections, row, col) === null) {
            sheet.setSelection(row, -1, 1, sheet.getColumnCount());
        }
        if (row !== undefined && col !== undefined) {
            $(".context-header").show();
            $(".context-cell").hide();
        }
    } else if (hitTestType === GcSpread.Sheets.SheetArea.viewport) {
        if (getCellInSelections(selections, row, col) === null) {
            sheet.clearSelection();
            sheet.endEdit();
            sheet.setActiveCell(row, col);
            updateMergeState();
        }
        if (row !== undefined && col !== undefined) {
            $(".context-header").hide();
            $(".context-cell").hide();
            showMergeContextMenu();
        } else {
            isHideContextMenu = true;
        }
    } else if (hitTestType === GcSpread.Sheets.SheetArea.corner) {
        sheet.setSelection(-1, -1, sheet.getRowCount(), sheet.getColumnCount());
        if (row !== undefined && col !== undefined) {
            $(".context-header").hide();
            $(".context-cell").show();
        }
    }

    var $contextMenu = $("#spreadContextMenu");
    $contextMenu.data("sheetArea", hitTestType);
    if (isHideContextMenu) {
        hideSpreadContextMenu();
    } else {
        $contextMenu.css({left: e.pageX, top: e.pageY});
        $contextMenu.show();

        $(document).on("mousedown.contextmenu", function () {
            if ($(event.target).parents("#spreadContextMenu").length === 0) {
                hideSpreadContextMenu();
            }
        });
    }
}

function hideSpreadContextMenu() {
    $("#spreadContextMenu").hide();
    $(document).off("mousedown.contextmenu");
}

function processContextMenuClicked() {
    var action = $(this).data("action");
    var sheet = spread.getActiveSheet();
    var sheetArea = $("#spreadContextMenu").data("sheetArea");

    hideSpreadContextMenu();

    switch (action) {
        case "cut":
            GcSpread.Sheets.SpreadActions.cut.call(sheet);
            break;
        case "copy":
            GcSpread.Sheets.SpreadActions.copy.call(sheet);
            break;
        case "paste":
            GcSpread.Sheets.SpreadActions.paste.call(sheet);
            break;
        case "insert":
            if (sheetArea === GcSpread.Sheets.SheetArea.colHeader) {
                sheet.addColumns(sheet.getActiveColumnIndex(), sheet.getSelections()[0].colCount);
            } else if (sheetArea === GcSpread.Sheets.SheetArea.rowHeader) {
                sheet.addRows(sheet.getActiveRowIndex(), sheet.getSelections()[0].rowCount);
            }
            break;
        case "delete":
            if (sheetArea === GcSpread.Sheets.SheetArea.colHeader) {
                sheet.deleteColumns(sheet.getActiveColumnIndex(), sheet.getSelections()[0].colCount);
            } else if (sheetArea === GcSpread.Sheets.SheetArea.rowHeader) {
                sheet.deleteRows(sheet.getActiveRowIndex(), sheet.getSelections()[0].rowCount);
            }
            break;
        case "merge":
            var sel = sheet.getSelections();
            if (sel.length > 0) {
                sel = sel[sel.length - 1];
                sheet.addSpan(sel.row, sel.col, sel.rowCount, sel.colCount, GcSpread.Sheets.SheetArea.viewport);
            }
            updateMergeState();
            break;
        case "unmerge":
            var sels = sheet.getSelections();
            for (var i = 0; i < sels.length; i++) {
                var sel = getActualCellRange(sels[i], sheet.getRowCount(), sheet.getColumnCount());
                for (var r = 0; r < sel.rowCount; r++) {
                    for (var c = 0; c < sel.colCount; c++) {
                        var span = sheet.getSpan(r + sel.row, c + sel.col, GcSpread.Sheets.SheetArea.viewport);
                        if (span) {
                            sheet.removeSpan(span.row, span.col, GcSpread.Sheets.SheetArea.viewport);
                        }
                    }
                }
            }
            updateMergeState();
            break;
        default:
            break;
    }
}
// context menu related items (end)

// import / export related items
function importSpreadFromJSON(file, fileContent) {
    function updateActiveCells() {
        for (var i = 0; i < spread.getSheetCount(); i++) {
            var sheet = spread.getSheet(i);
            columnIndex = sheet.getActiveColumnIndex(),
                rowIndex = sheet.getActiveRowIndex();
            if (columnIndex !== undefined && rowIndex !== undefined) {
                spread.getSheet(i).setActiveCell(rowIndex, columnIndex);
            } else {
                spread.getSheet(i).setActiveCell(0, 0);
            }
        }
    }

    function importSuccessCallback(responseText) {
        var spreadJson = JSON.parse(responseText);
        if (spreadJson.version && spreadJson.sheets) {
            spread.unbindAll();
            spread.fromJSON(spreadJson);
            app.reset(true);
            updateActiveCells();
            spread.focus();
            onCellSelected();
            syncSpreadPropertyValues();
            syncSheetPropertyValues();
            
            app.fileMenu.closeFileScreen();
        } else {
            alert(getResource("messages.invalidImportFile"));
        }
    }

    if (file) {
        var reader = new FileReader();
        reader.onload = function () {
            importSuccessCallback(this.result);
        };
        reader.readAsText(file);
    } else if (fileContent) {
        importSuccessCallback(fileContent);
    }
    return true;
}

function exportToJSON(fileName) {
    function getFileName() {
        function to2DigitsString(num) {
            return ("0" + num).substr(-2);
        }

        var date = new Date();
        return [
            "export",
            date.getFullYear(), to2DigitsString(date.getMonth() + 1), to2DigitsString(date.getDate()),
            to2DigitsString(date.getHours()), to2DigitsString(date.getMinutes()), to2DigitsString(date.getSeconds())
        ].join("");
    }

    var json = spread.toJSON({includeBindingSource: true}),
        text = JSON.stringify(json);

    fileName = fileName || getFileName();

    saveAs(new Blob([text], {type: "text/plain;charset=utf-8"}), fileName + ".ssjson");
}
app.saveBrowserFile = exportToJSON;

// import / export related items (end)

// positionbox related items
function getSelectedRangeString(sheet, range) {
    var selectionInfo = "",
        rowCount = range.rowCount,
        columnCount = range.colCount,
        startRow = range.row + 1,
        startColumn = range.col + 1;

    if (rowCount == 1 && columnCount == 1) {
        selectionInfo = getCellPositionString(sheet, startRow, startColumn);
    }
    else {
        if (rowCount < 0 && columnCount > 0) {
            selectionInfo = columnCount + "C";
        }
        else if (columnCount < 0 && rowCount > 0) {
            selectionInfo = rowCount + "R";
        }
        else if (rowCount < 0 && columnCount < 0) {
            selectionInfo = sheet.getRowCount() + "R x " + sheet.getColumnCount() + "C";
        }
        else {
            selectionInfo = rowCount + "R x " + columnCount + "C";
        }
    }
    return selectionInfo;
}

function getColumnName(column) {
    var letters = "";
    while (column > 0) {
        var num = column % 26;
        if (num === 0) {
            letters = "Z" + letters;
            column--;
        }
        else {
            letters = String.fromCharCode('A'.charCodeAt(0) + num - 1) + letters;
        }
        column = parseInt((column / 26).toString());
    }
    return letters;
}

function getCellPositionString(sheet, row, column) {
    if (row < 1 || column < 1) {
        return null;
    }
    else {
        var letters = "";
        switch (spread.referenceStyle()) {
            case spreadNS.ReferenceStyle.A1: // 0
                letters = getColumnName(column) + row.toString();
                break;
            case spreadNS.ReferenceStyle.R1C1: // 1
                letters = "R" + row.toString() + "C" + column.toString();
                break;
            default:
                break;
        }
        return letters;
    }
}
// positionbox related items (end)

// theme color related items
function setThemeColorToSheet(sheet) {
    sheet.isPaintSuspended(true);

    sheet.getCell(2, 3).text("Background 1").themeFont("Body");
    sheet.getCell(2, 4).text("Text 1").themeFont("Body");
    sheet.getCell(2, 5).text("Background 2").themeFont("Body");
    sheet.getCell(2, 6).text("Text 2").themeFont("Body");
    sheet.getCell(2, 7).text("Accent 1").themeFont("Body");
    sheet.getCell(2, 8).text("Accent 2").themeFont("Body");
    sheet.getCell(2, 9).text("Accent 3").themeFont("Body");
    sheet.getCell(2, 10).text("Accent 4").themeFont("Body");
    sheet.getCell(2, 11).text("Accent 5").themeFont("Body");
    sheet.getCell(2, 12).text("Accent 6").themeFont("Body");

    sheet.getCell(4, 1).value("100").themeFont("Body");

    sheet.getCell(4, 3).backColor("Background 1");
    sheet.getCell(4, 4).backColor("Text 1");
    sheet.getCell(4, 5).backColor("Background 2");
    sheet.getCell(4, 6).backColor("Text 2");
    sheet.getCell(4, 7).backColor("Accent 1");
    sheet.getCell(4, 8).backColor("Accent 2");
    sheet.getCell(4, 9).backColor("Accent 3");
    sheet.getCell(4, 10).backColor("Accent 4");
    sheet.getCell(4, 11).backColor("Accent 5");
    sheet.getCell(4, 12).backColor("Accent 6");

    sheet.getCell(5, 1).value("80").themeFont("Body");

    sheet.getCell(5, 3).backColor("Background 1 80");
    sheet.getCell(5, 4).backColor("Text 1 80");
    sheet.getCell(5, 5).backColor("Background 2 80");
    sheet.getCell(5, 6).backColor("Text 2 80");
    sheet.getCell(5, 7).backColor("Accent 1 80");
    sheet.getCell(5, 8).backColor("Accent 2 80");
    sheet.getCell(5, 9).backColor("Accent 3 80");
    sheet.getCell(5, 10).backColor("Accent 4 80");
    sheet.getCell(5, 11).backColor("Accent 5 80");
    sheet.getCell(5, 12).backColor("Accent 6 80");

    sheet.getCell(6, 1).value("60").themeFont("Body");

    sheet.getCell(6, 3).backColor("Background 1 60");
    sheet.getCell(6, 4).backColor("Text 1 60");
    sheet.getCell(6, 5).backColor("Background 2 60");
    sheet.getCell(6, 6).backColor("Text 2 60");
    sheet.getCell(6, 7).backColor("Accent 1 60");
    sheet.getCell(6, 8).backColor("Accent 2 60");
    sheet.getCell(6, 9).backColor("Accent 3 60");
    sheet.getCell(6, 10).backColor("Accent 4 60");
    sheet.getCell(6, 11).backColor("Accent 5 60");
    sheet.getCell(6, 12).backColor("Accent 6 60");

    sheet.getCell(7, 1).value("40").themeFont("Body");

    sheet.getCell(7, 3).backColor("Background 1 40");
    sheet.getCell(7, 4).backColor("Text 1 40");
    sheet.getCell(7, 5).backColor("Background 2 40");
    sheet.getCell(7, 6).backColor("Text 2 40");
    sheet.getCell(7, 7).backColor("Accent 1 40");
    sheet.getCell(7, 8).backColor("Accent 2 40");
    sheet.getCell(7, 9).backColor("Accent 3 40");
    sheet.getCell(7, 10).backColor("Accent 4 40");
    sheet.getCell(7, 11).backColor("Accent 5 40");
    sheet.getCell(7, 12).backColor("Accent 6 40");

    sheet.getCell(8, 1).value("-25").themeFont("Body");

    sheet.getCell(8, 3).backColor("Background 1 -25");
    sheet.getCell(8, 4).backColor("Text 1 -25");
    sheet.getCell(8, 5).backColor("Background 2 -25");
    sheet.getCell(8, 6).backColor("Text 2 -25");
    sheet.getCell(8, 7).backColor("Accent 1 -25");
    sheet.getCell(8, 8).backColor("Accent 2 -25");
    sheet.getCell(8, 9).backColor("Accent 3 -25");
    sheet.getCell(8, 10).backColor("Accent 4 -25");
    sheet.getCell(8, 11).backColor("Accent 5 -25");
    sheet.getCell(8, 12).backColor("Accent 6 -25");

    sheet.getCell(9, 1).value("-50").themeFont("Body");

    sheet.getCell(9, 3).backColor("Background 1 -50");
    sheet.getCell(9, 4).backColor("Text 1 -50");
    sheet.getCell(9, 5).backColor("Background 2 -50");
    sheet.getCell(9, 6).backColor("Text 2 -50");
    sheet.getCell(9, 7).backColor("Accent 1 -50");
    sheet.getCell(9, 8).backColor("Accent 2 -50");
    sheet.getCell(9, 9).backColor("Accent 3 -50");
    sheet.getCell(9, 10).backColor("Accent 4 -50");
    sheet.getCell(9, 11).backColor("Accent 5 -50");
    sheet.getCell(9, 12).backColor("Accent 6 -50");
    sheet.isPaintSuspended(false);
}

function getColorName(sheet, row, col) {
    var colName = sheet.getCell(2, col).text();
    var rowName = sheet.getCell(row, 1).text();
    return colName + " " + rowName;
}

function getThemeColor() {
    var sheet = spread.getActiveSheet();
    setThemeColorToSheet(sheet);                                            // Set current theme color to sheet
    var themeColors = [];
    var $colorUl = $("#default-theme-color");
    var $themeColorLi, cellBackColor;
    for (var col = 3; col < 13; col++) {
        var row = 4;
        cellBackColor = sheet.getActualStyle(row, col).backColor;
        $themeColorLi = $("<li class=\"color-cell seed-color-column\"></li>");
        $themeColorLi.css("background-color", cellBackColor).attr("data-name", sheet.getCell(2, col).text()).appendTo($colorUl);
        for (row = 5; row < 10; row++) {
            cellBackColor = sheet.getActualStyle(row, col).backColor;
            $themeColorLi = $("<li class=\"color-cell\"></li>");
            $themeColorLi.css("background-color", cellBackColor).attr("data-name", getColorName(sheet, row, col)).appendTo($colorUl);
        }
    }
    
    for (var i = 4; i < 10; i++) {
        for (var j = 3; j < 13; j++) {
            cellBackColor = sheet.getActualStyle(i, j).backColor;
            themeColors[themeColors.length] = cellBackColor;
        }
    }

    sheet.clear(2, 1, 8, 12, GcSpread.Sheets.SheetArea.viewport, 255);      // Clear sheet theme color
    
    return themeColors;
}
// theme color related items (end)

// slicer related items
function syncSlicerPropertyValues(sheet) {
    var selectedSlicers = getSelectedSlicers(sheet);
    if (!selectedSlicers || selectedSlicers.length === 0) {
        return false;
    }
    else if (selectedSlicers.length > 1) {
        getMultiSlicerSetting(selectedSlicers);
        setTextDisabled("slicerName", true);
    }
    else if (selectedSlicers.length === 1) {
        getSingleSlicerSetting(selectedSlicers[0]);
        setTextDisabled("slicerName", false);
    }
    
    return true;
}

function getSingleSlicerSetting(slicer) {
    if (!slicer) {
        return;
    }
    setTextValue("slicerName", slicer.name());
    setTextValue("slicerCaptionName", slicer.captionName());
    setDropDown("slicerItemSorting", slicer.sortState());
    setCheckValue("displaySlicerHeader", slicer.showHeader());
    setNumberValue("slicerColumnNumber", slicer.columnCount());
    setNumberValue("slicerButtonWidth", getSlicerItemWidth(slicer.columnCount(), slicer.width()));
    setNumberValue("slicerButtonHeight", slicer.itemHeight());
    if (slicer.dynamicMove()) {
        if (slicer.dynamicSize()) {
            setRadioItemChecked("slicerMoveAndSize", "slicer-move-size");
        }
        else {
            setRadioItemChecked("slicerMoveAndSize", "slicer-move-nosize");
        }
    }
    else {
        setRadioItemChecked("slicerMoveAndSize", "slicer-nomove-size");
    }
    setCheckValue("lockSlicer", slicer.isLocked());
    
    var styleName = getSlicerStyleName(slicer);
    markSlicerStyleSelected(styleName);
        
    setNoDataRelatedValue(!slicer.showNoDataItems(), slicer.visuallyNoDataItems(), slicer.showNoDataItemsInLast());
}

function getMultiSlicerSetting(selectedSlicers) {
    if (!selectedSlicers || selectedSlicers.length === 0) {
        return;
    }
    var slicer = selectedSlicers[0];
    var isDisplayHeader = false,
        isSameSortState = true,
        isSameCaptionName = true,
        isSameColumnCount = true,
        isSameItemHeight = true,
        isSameItemWidth = true,
        isSameLocked = true,
        isSameDynamicMove = true,
        isSameDynamicSize = true,
        isHideItem = false,
        isVisuallyItem = false,
        isShowItemLast = false;

    var sortState = slicer.sortState(),
        captionName = slicer.captionName(),
        columnCount = slicer.columnCount(),
        itemHeight = slicer.itemHeight(),
        itemWidth = getSlicerItemWidth(columnCount, slicer.width()),
        dynamicMove = slicer.dynamicMove(),
        dynamicSize = slicer.dynamicSize();
        
    var styleName = getSlicerStyleName(slicer);

    for (var item in selectedSlicers) {
        var slicer = selectedSlicers[item];
        isDisplayHeader = isDisplayHeader || slicer.showHeader();
        isSameLocked = isSameLocked && slicer.isLocked();
        if (slicer.sortState() !== sortState) {
            isSameSortState = false;
        }
        if (slicer.captionName() !== captionName) {
            isSameCaptionName = false;
        }
        if (slicer.columnCount() !== columnCount) {
            isSameColumnCount = false;
        }
        if (slicer.itemHeight() !== itemHeight) {
            isSameItemHeight = false;
        }
        if (getSlicerItemWidth(slicer.columnCount(), slicer.width()) !== itemWidth) {
            isSameItemWidth = false;
        }
        if (slicer.dynamicMove() !== dynamicMove) {
            isSameDynamicMove = false;
        }
        if (slicer.dynamicSize() !== dynamicSize) {
            isSameDynamicSize = false;
        }
        if (styleName && (styleName !== getSlicerStyleName(slicer))) {
            styleName = "";
        }
        isHideItem = isHideItem || !slicer.showNoDataItems();
        isVisuallyItem = isVisuallyItem || slicer.visuallyNoDataItems();
        isShowItemLast = isShowItemLast || slicer.showNoDataItemsInLast();
    }
    
    markSlicerStyleSelected(styleName);

    setTextValue("slicerName", "");
    if (isSameCaptionName) {
        setTextValue("slicerCaptionName", captionName);
    }
    else {
        setTextValue("slicerCaptionName", "");
    }
    if (isSameSortState) {
        setDropDown("slicerItemSorting", sortState);
    }
    else {
        setDropDown("slicerItemSorting", "");
    }
    setCheckValue("displaySlicerHeader", isDisplayHeader);
    if (isSameDynamicMove && isSameDynamicSize && dynamicMove) {
        if (dynamicSize) {
            setRadioItemChecked("slicerMoveAndSize", "slicer-move-size");
        }
        else {
            setRadioItemChecked("slicerMoveAndSize", "slicer-move-nosize");
        }
    }
    else {
        setRadioItemChecked("slicerMoveAndSize", "slicer-nomove-size");
    }
    if (isSameColumnCount) {
        setNumberValue("slicerColumnNumber", columnCount);
    }
    else {
        setNumberValue("slicerColumnNumber", "");
    }
    if (isSameItemHeight) {
        setNumberValue("slicerButtonHeight", Math.round(itemHeight));
    }
    else {
        setNumberValue("slicerButtonHeight", "");
    }
    if (isSameItemWidth) {
        setNumberValue("slicerButtonWidth", itemWidth);
    }
    else {
        setNumberValue("slicerButtonWidth", "");
    }
    setCheckValue("lockSlicer", isSameLocked);
    setNoDataRelatedValue(isHideItem, isVisuallyItem, isShowItemLast);
}

function getSlicerStyleName(slicer) {
    var slicerStyle = slicer.style(),
        styleName = slicerStyle && slicerStyle.name();
        
    return styleName;    
}

var selectedStyleClassName = "slicer-format-item-selected";

function markSlicerStyleSelected(styleName) {
    $(".slicer-format-2013 div." + selectedStyleClassName).removeClass(selectedStyleClassName);
    if (styleName) {
        var name = styleName.split("SlicerStyle")[1];
        if (name) {    
            $(".slicer-format-2013 div[data-name='" + name.toLowerCase() + "']").parent().addClass(selectedStyleClassName);
        }
    }
}

function setNoDataRelatedValue(isHideItem, isVisuallyItem, isShowItemLast) {
    setCheckValue("hide-no-data", isHideItem);
    setCheckValue("mark-no-data", isVisuallyItem);
    setCheckValue("show-no-data-last", isShowItemLast);
    enableRelatedItems(["mark-no-data"], !isHideItem);
    enableRelatedItems(["show-no-data-last"], !isHideItem || isVisuallyItem);
}

function changeSlicerInfo(slicer, propertyName) {
    if (!slicer) {
        return;
    }
    switch (propertyName) {
        case "width":
            setNumberValue("slicerButtonWidth", getSlicerItemWidth(slicer.columnCount(), slicer.width()));
            break;
    }
}

function setSlicerSetting(property, value) {
    var sheet = spread.getActiveSheet();
    var selectedSlicers = getSelectedSlicers(sheet);
    if (!selectedSlicers || selectedSlicers.length === 0) {
        return;
    }
    else {
        for (var item in selectedSlicers) {
            setSlicerProperty(selectedSlicers[item], property, value);
        }
    }
}

function setSlicerProperty(slicer, property, value) {
    switch (property) {
        case "name":
            var sheet = spread.getActiveSheet();
            var slicerPreName = slicer.name();
            if (!value) {
                alert(getResource("messages.invalidSlicerName"));
                setTextValue("slicerName", slicerPreName);
            }
            else if (value && value !== slicerPreName) {
                if (sheet.getSlicer(value)) {
                    alert(getResource("messages.duplicatedSlicerName"));
                    setTextValue("slicerName", slicerPreName);
                }
                else {
                    slicer.name(value);
                }
            }
            break;
        case "captionName":
            slicer.captionName(value);
            break;
        case "sortState":
            slicer.sortState(value);
            break;
        case "showHeader":
            slicer.showHeader(value);
            break;
        case "columnCount":
            slicer.columnCount(value);
            break;
        case "itemHeight":
            slicer.itemHeight(value);
            break;
        case "itemWidth":
            slicer.width(getSlicerWidthFromItem(slicer.columnCount(), value));
            break;
        case "moveSize":
            if (value === "slicer-move-size") {
                slicer.dynamicMove(true);
                slicer.dynamicSize(true);
            }
            if (value === "slicer-move-nosize") {
                slicer.dynamicMove(true);
                slicer.dynamicSize(false);
            }
            if (value === "slicer-nomove-size") {
                slicer.dynamicMove(false);
                slicer.dynamicSize(false);
            }
            break;
        case "lock":
            slicer.isLocked(value);
            break;
        case "style":
            slicer.style(value);
            break;
            
        case "showNoDataItems":
            slicer.showNoDataItems(value);
            break;
            
        case "visuallyNoDataItems":
            slicer.visuallyNoDataItems(value);
            break;
            
        case "showNoDataItemsInLast":
            slicer.showNoDataItemsInLast(value);
            break;
            
        default:
            console.log("Slicer doesn't have property:", property);
            break;
    }
}

function setTextDisabled(name, isDisabled) {
    var $input = $("inpu[data-name='" + name + "']");
    $input.attr("disabled", isDisabled);
}

function setRadioItemChecked(groupName, value) {
    var groupSelector = "input[name='" + groupName + "']"; 
    $(groupSelector).removeClass("checked");
    $(groupSelector + "[data-value='" + value + "']").addClass("checked");
}

function getSlicerItemWidth(count, slicerWidth) {
    if (count <= 0) {
        count = 1; //Column count will be converted to 1 if it is set to 0 or negative number.
    }
    var SLICER_PADDING = 6;
    var SLICER_ITEM_SPACE = 2;
    var itemWidth = Math.round((slicerWidth - SLICER_PADDING * 2 - (count - 1) * SLICER_ITEM_SPACE) / count);
    if (itemWidth < 0) {
        return 0;
    }
    else {
        return itemWidth;
    }
}

function getSlicerWidthFromItem(count, itemWidth) {
    if (count <= 0) {
        count = 1; //Column count will be converted to 1 if it is set to 0 or negative number.
    }
    var SLICER_PADDING = 6;
    var SLICER_ITEM_SPACE = 2;
    return Math.round(itemWidth * count + (count - 1) * SLICER_ITEM_SPACE + SLICER_PADDING * 2);
}

function getSelectedSlicers(sheet) {
    if (!sheet) {
        return null;
    }
    var slicers = sheet.getSlicers();
    if (!slicers || slicers.length === 0) {
        return null;
    }
    var selectedSlicers = [];
    for (var item in slicers) {
        if (slicers[item].isSelected()) {
            selectedSlicers.push(slicers[item]);
        }
    }
    return selectedSlicers;
}

function processSlicerItemSorting(sortValue) {
    switch (sortValue) {
        case 0:
        case 1:
        case 2:
            setSlicerSetting("sortState", sortValue);
            break;

        default:
            console.log("processSlicerItemSorting not process with ", name);
            return;
    }
}

function changeSlicerStyle() {
    spread.isPaintSuspended(true);

    var styleName = $(">div", this).data("name");
    setSlicerSetting("style", spreadNS.SlicerStyles[styleName]());
    markSlicerStyleSelected("SlicerStyle" + styleName);

    spread.isPaintSuspended(false);
}
// slicer related items (end)

// spread theme related items
function processChangeSpreadTheme(value) {
    $("link[title='spread-theme']").attr("href", value);

    setTimeout(
        function () {
            spread.refresh();
        }, 300);
}
// spread theme related items (end)

// ribbon related items
function initRibbon(colorPicker) {
    ribbon = new Ribbon(data,
        document.getElementById('toolbar'),
        colorPicker);

    $tableStyleDropdown = $("#table div[data-name='tableStyles'] > ul");

    // prepare table style drop down
    $("<li></li>").append($("#tableStyles")).appendTo($tableStyleDropdown);

    // set tooltips for dropdown and icon
    var tooltip = $tableStyleDropdown.parent().parent().attr("title");
    $tableStyleDropdown.parent().attr("title", tooltip);
    // icon set to default table style (medium2), will update to last selected table style for quick set
    $tableStyleDropdown.parent().prev().attr("title", uiResource.tableTab.tableStyle.medium.medium2);

    $(ribbon).on('click', processRibbonClick);
    $(ribbon).on('dropdown', processRibbonDropDown);
    $(ribbon).on('dropdownShown', processRibbonDropDownShown);
    $(ribbon).on('textChanged', processTextChanged);

    $('a[href="#file"]').on('click', function (e) {
        e.stopPropagation();
        e.preventDefault();
        app.fileMenu.showFileScreen();
    });
}

function processRibbonClick(e, data) {
    var name = data.name, value = data.value, title = data.header || data.text || name;
    var sheet = spread.getActiveSheet();

    spread.isPaintSuspended(true);

    switch (name) {
        // Font and Formatting Group (Home Tab)
        case "fontFamily":
            Actions.setTextFontFamily(spread, {name: name, value: value});
            break;

        case "fontSize":
            Actions.setTextFontSize(spread, {name: name, value: value + "pt"});
            break;

        case "bold":
            Actions.setTextBold(spread, {name: name, value: value});
            break;

        case "italic":
            Actions.setTextItalic(spread, {name: name, value: value});
            break;

        case "underline":
            Actions.setTextUnderline(spread, {name: name, value: value});
            break;

        case "overline":
            Actions.setTextOverline(spread, {name: name, value: value});
            break;

        case "strikethrough":
            Actions.setTextStrikethrough(spread, {name: name, value: value});
            break;

        case "border":
            if (value === "SET" || value === "more") {
                displaySettingPane(title, $('#borderSetting'));
            } else {
                setCellsBorder(value);
            }
            break;

        case "backColor":
            Actions.setTextBackColor(spread, {name: name, value: settingCache[name]});
            break;

        case "foreColor":
            Actions.setTextForeColor(spread, {name: name, value: settingCache[name]});
            break;

        // Alignment Group (Home Tab)
        case "valign-top":
        case "valign-middle":
        case "valign-bottom":
        case "halign-left":
        case "halign-center":
        case "halign-right":
            var aligns = name.split('-');
            Actions.setTextAligns(spread, {
                name: name,
                value: {align: aligns[0] === 'halign' ? 'hAlign' : 'vAlign', value: aligns[1]}
            });
            break;

        case "cellmerge":
            Actions.setMergeCenter(spread, value);
            break;

        case "wordwrap":
            Actions.setWrapText(spread, {name: name, value: value});
            break;

        case "indent":
        case "outdent":
            Actions.indent(spread, {name: name, value: name === 'indent' ? 1 : -1});
            break;

        // CellType Group (Home Tab)
        case "celltype":
            var $groups = $('#cellTypeSetting .group-celltype');
            $groups.addClass('hidden');
            $groups.filter('[data-name="' + value + '"]').removeClass('hidden');

            displaySettingPane(title, $('#cellTypeSetting'));
            break;

        // Cell Format Group (Home Tab)
        case "cellformat":
            if (value === "custom") {
                displaySettingPane(title, $('#cellFormatSetting'), syncCellFormat);
            } else {
                Actions.setCellFormat(spread, {name: name, value: value === 'nullValue' ? null : value});
            }
            break;

        case "numberformat":
            switch (value) {
                case "percentStyle":
                case "commaStyle":
                    Actions.setCellFormat(spread, {
                        name: name,
                        value: value === 'percentStyle' ? uiResource.cellTab.format.percentValue : uiResource.cellTab.format.commaValue
                    });
                    break;

                case "increaseDecimal":
                    Actions.increaseDecimal(spread, {name: name});
                    break;

                case "decreaseDecimal":
                    Actions.decreaseDecimal(spread, {name: name});
                    break;
            }
            break;

        // Protect Group (Home Tab)
        case "protectsheet":
            Actions.setIsProtected(spread, {name: name, value: value});
            break;

        case "unlockcells":
            Actions.setCellLock(spread, {name: name, value: value});
            break;

        // Cells Group (Home Tab)
        case "cellsgroup":
            processInsertDeleteCells(value);
            break;

        case "clearformat":
            processClearCells(value);
            break;

        // Condition Format (Home Tab)
        case "conditionalformat":
            alert('TODO items, please wait.');
            break;

        case "find":
            var $container = $("#findOptions");
            displaySettingPane(title, $container, function () {
                // clear input and set default checked items
                setTextValue("findwhat", "");
                $("input[name='findin']", $container).first().prop("checked", true);
                $("input[name='searchby']", $container).first().prop("checked", true);
                $("input[name='lookin']", $container).first().prop("checked", true);
                setCheckValue("findMatchCase", false);
                setCheckValue("findMatchExactly", false);
                setCheckValue("findUseWildcards", false);
                $(".resultcount", $container).text("");
                $(".findoutput", $container).hide();
                findCache = {
                    activeSheetIndex: spread.getActiveSheetIndex(),
                    activeCellRowIndex: sheet.getActiveRowIndex(),
                    activeCellColumnIndex: sheet.getActiveColumnIndex()
                };
            });
            break;

        // Insert Tab
        case "insertTable":
            Actions.addTable(spread);
            onCellSelected();
            break;

        case "insertPicture":
            if (app && app.isNative) {
                app.showOpenFileDialog(
                    [{name: 'Images', extensions: ['jpg', 'png', 'gif']}],    // filters
                    function (err, data) {
                        if (!err && data) {
                            var dataUrl = 'data:image/jpeg;base64,' + new Buffer(data).toString('base64');
                            Actions.addPicture(spread, {name: "Picture" + pictureIndex++, url: dataUrl});
                        }
                    }
                );
            } else {
                $("#fileSelector").data("action", "addpicture");
                $("#fileSelector").attr("accept", "image/*");
                $("#fileSelector").click();
            }
            break;

        case "insertLink":
            var typeName = 'hyperlink';
            var $groups = $('#cellTypeSetting .group-celltype');
            $groups.addClass('hidden');
            $groups.filter('[data-name="' + typeName + '"]').removeClass('hidden');

            displaySettingPane(title, $('#cellTypeSetting'));
            break;

        case "insertComment":
            Actions.insertComment(spread);
            break;

        case "insertSparkline":
            processAddSparklineEx(value);
            break;

        // Formulas Tab
        case "autoSum":
            // SET works like dropdown' sum item
            processAutoSum(sheet, value === "SET" ? "sum" : value);
            break;

        case "insertFormula":
            displaySettingPane(title, $('#functionBuiilder'));
            break;

        case "calculateNow":
            sheet.recalcAll();
            break;

        // Sort & Filter
        case "sortAZ":
        case "sortZA":
            sortData(sheet, name === "sortAZ");
            break;

        case "filter":
            updateFilter(sheet);
            break;

        // Group
        case "group":
            addGroup(sheet);
            break;

        case "ungroup":
            removeGroup(sheet);
            break;

        case "showDetail":
            toggleGroupDetail(sheet, true);
            break;

        case "hideDetail":
            toggleGroupDetail(sheet, false);
            break;

        case "summaryBelow":
            sheet.rowRangeGroup.setDirection(value ? 1 /* Forward */ : 0 /* Backward */);
            break;

        case "summaryRight":
            sheet.colRangeGroup.setDirection(value ? 1 /* Forward */ : 0 /* Backward */);
            break;

        case "circleInvalidData":
            spread.highlightInvalidData(value);
            break;

        case "selectValidator":
            displaySettingPane(title, $('#dataValidationSetting'));
            break;

        // Show / Hide Group (View Tab)
        case "showFormulaBar":
            $("#formulaBar")[value ? "show" : "hide"]();
            adjustSpreadSize();
            break;

        case "showGridlines":
            sheet.gridline.showVerticalGridline = value;
            sheet.gridline.showHorizontalGridline = value;
            break;

        case "showHeadings":
            sheet.setRowHeaderVisible(value);
            sheet.setColumnHeaderVisible(value);
            break;

        case "showSheetTabs":
            spread.tabStripVisible(value);
            break;

        // Freeze Panes Group (View Tab)
        case "freezePanes":
            processFreezePane(sheet, value);
            break;

        // Table Tab
        case "tableOption":
            setTableOption(value, data.checked);
            break;

        case "tableStyles":
            setTableStyle(_activeTable, settingCache.tableStyles);
            break;

        default:
            console.log("ribbon-click", name, value);
            break;
    }
    spread.isPaintSuspended(false);
    spread.focus();
}

function processFreezePane(sheet, typeName) {
    var freezeOptionMap = {
        row: "setFrozenRowCount",
        col: "setFrozenColumnCount",
        trailingRow: "setFrozenTrailingRowCount",
        trailingCol: "setFrozenTrailingColumnCount"
    };

    function freeze(sheet, options) {
        for (var propertyName in options) {
            var value = +options[propertyName];
            if (!isNaN(value)) {
                sheet[freezeOptionMap[propertyName]](value);
            }
        }
    }

    var options;
    switch (typeName) {
        case "freezePanes":
            options = {row: sheet.getActiveRowIndex(), col: sheet.getActiveColumnIndex()};
            break;

        case "freezeTopRow":
            options = {row: 1, col: 0, trailingRow: 0, trailingCol: 0};
            break;

        case "freezeFirstColumn":
            options = {row: 0, col: 1, trailingRow: 0, trailingCol: 0};
            break;

        case "freezeBottomRow":
            options = {row: 0, col: 0, trailingRow: 1, trailingCol: 0};
            break;

        case "freezeLastColumn":
            options = {row: 0, col: 0, trailingRow: 0, trailingCol: 1};
            break;

        case "unfreeze":
            options = {row: 0, col: 0, trailingRow: 0, trailingCol: 0};
            break;

        default:
            return;
    }

    if (options) {
        freeze(sheet, options);
    }
}

function setTableOption(name, value) {
    switch (name) {
        case "tableFilterButton":
            _activeTable && _activeTable.filterButtonVisible(value);
            break;

        case "tableHeaderRow":
            _activeTable && _activeTable.showHeader(value);
            break;

        case "tableTotalRow":
            _activeTable && _activeTable.showFooter(value);
            break;

        case "tableBandedRows":
            _activeTable && _activeTable.bandRows(value);
            break;

        case "tableBandedColumns":
            _activeTable && _activeTable.bandColumns(value);
            break;

        case "tableFirstColumn":
            _activeTable && _activeTable.highlightFirstColumn(value);
            break;

        case "tableLastColumn":
            _activeTable && _activeTable.highlightLastColumn(value);
            break;
    }
}

function setCellsBorder(borderType) {
    var res = selectionBorderType(borderType,
        $('#borderSetting .border-line-color .color-picker').css('background-color'),
        getBorderLineType('line-style-' + $('#borderSetting .border-line-style li.selected').data('value'))
    );
    Actions.setBorder(spread, res);
}

function selectionBorderType(borderType, color, lineStyle) {
    var lineBorder = new GcSpread.Sheets.LineBorder(color, lineStyle);
    var result = [];
    switch (borderType) {
        case 'inside':
            result.push({lineBorder: lineBorder, value: {innerHorizontal: true, innerVertical: true}});
            break;
        case 'innerVertical':
            result.push({lineBorder: lineBorder, value: {innerVertical: true}});
            break;
        case 'innerHorizontal':
            result.push({lineBorder: lineBorder, value: {innerHorizontal: true}});
            break;
        case 'bottom':
            result.push({lineBorder: lineBorder, value: {bottom: true}});
            break;
        case 'top':
            result.push({lineBorder: lineBorder, value: {top: true}});
            break;
        case 'left':
            result.push({lineBorder: lineBorder, value: {left: true}});
            break;
        case 'right':
            result.push({lineBorder: lineBorder, value: {right: true}});
            break;
        case 'all':
            result.push({lineBorder: lineBorder, value: {all: true}});
            break;
        case 'none':
            result.push({
                lineBorder: new GcSpread.Sheets.LineBorder(color, GcSpread.Sheets.LineStyle.empty),
                value: {all: true}
            });
            break;
        case 'outside':
            result.push({lineBorder: lineBorder, value: {outline: true}});
            break;
        case 'thick':
            result.push({
                lineBorder: new GcSpread.Sheets.LineBorder(color, GcSpread.Sheets.LineStyle.thick),
                value: {outline: true}
            });
            break;
        case "doublebottom":
            result.push({
                lineBorder: new GcSpread.Sheets.LineBorder(color, GcSpread.Sheets.LineStyle.double),
                value: {bottom: true}
            });
            break;
        case "thickbottom":
            result.push({
                lineBorder: new GcSpread.Sheets.LineBorder(color, GcSpread.Sheets.LineStyle.thick),
                value: {bottom: true}
            });
            break;
        case "top-bottom":
            result.push({
                lineBorder: new GcSpread.Sheets.LineBorder(color, GcSpread.Sheets.LineStyle.thin),
                value: {bottom: true, top: true}
            });
            break;
        case "top-thickbottom":
            result.push({
                lineBorder: new GcSpread.Sheets.LineBorder(color, GcSpread.Sheets.LineStyle.thin),
                value: {top: true}
            });
            result.push({
                lineBorder: new GcSpread.Sheets.LineBorder(color, GcSpread.Sheets.LineStyle.thick),
                value: {bottom: true}
            });
            break;
        case "top-doublebottom":
            result.push({
                lineBorder: new GcSpread.Sheets.LineBorder(color, GcSpread.Sheets.LineStyle.thin),
                value: {top: true}
            });
            result.push({
                lineBorder: new GcSpread.Sheets.LineBorder(color, GcSpread.Sheets.LineStyle.double),
                value: {bottom: true}
            });
            break;
    }
    return result;
}

function processInsertDeleteCells(name) {
    switch (name) {
        case "insertRows":
            Actions.insertRows(spread);
            break;

        case "insertColumns":
            Actions.insertColumns(spread);
            break;

        case "insert-shiftCellsRight":
            Actions.insertRightCells(spread);
            break;

        case "insert-shiftCellsDown":
            Actions.insertDownCells(spread);
            break;

        case "deleteRows":
            Actions.deleteRows(spread);
            break;

        case "deleteColumns":
            Actions.deleteColumns(spread);
            break;

        case "delete-shiftCellsLeft":
            Actions.deleteLeftCells(spread);
            break;

        case "delete-shiftCellsUp":
            Actions.deleteUpCells(spread);
            break;
    }
}

function processClearCells(name) {
    switch (name) {
        case "clearAll":
            Actions.clearAll(spread);
            break;

        case "clearFormatting":
            Actions.clearFormatting(spread);
            break;

        case "clear":
            Actions.clearData(spread);
            break;
    }
}


function getCellFormatter() {
    var sheet = spread.getActiveSheet();
    var style = sheet.getActualStyle(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex());

    return style && style.formatter || 'nullValue';
}


function processAutoSum(sheet, funcName) {
    /*jshint eqnull:true */

    var ResultPosition = {Bottom: 1, Right: 2};

    // check range to decide where to add:
    //      position:   bottom (1), right (2) or both (3)
    //      rowOffset:  -1 if last row is empty, otherwise 0
    //      colOffset:  -1 if last column is empty, otherwise 0
    function getFormulaPosition(range) {
        var r, c, result = {rowOffset: 0, colOffset: 0}, position = 0,
            rowCount = range.rowCount, colCount = range.colCount,
            startRow = range.row, startCol = range.col,
            endRow = startRow + rowCount - 1, endCol = startCol + colCount - 1;

        if (rowCount === 1 || colCount === 1) {
            position = (rowCount > 1) ? ResultPosition.Bottom : ((colCount > 1) ? ResultPosition.Right : 0);
        }
        // check last row for empty cells
        var hasValue = false;
        if (rowCount > 1) {
            for (c = startCol; c <= endCol; c++) {
                if ((sheet.getValue(endRow, c) != null)) {
                    hasValue = true;
                    break;
                }
            }
            if (!hasValue) {
                position |= ResultPosition.Bottom;
                result.rowOffset = -1;
            }
        }

        if (colCount > 1) {
            hasValue = false;
            // check last col for empty cells
            for (r = startRow; r <= endRow; r++) {
                if ((sheet.getValue(r, endCol) != null)) {
                    hasValue = true;
                    break;
                }
            }
            if (!hasValue) {
                position |= ResultPosition.Right;
                result.colOffset = -1;
            }
        }

        // both the last row / column of range is not empty, return bottom
        if (!position) {
            position = ResultPosition.Bottom;
        }

        result.position = position;

        return result;
    }

    function getFormulaTemplateForRow(range) {
        // TODO: process w/ A1 & R1C1 style
        // funcName(@1:@2)
        var row = range.row;
        return [funcName.toUpperCase(), "(@", row + 1, ":@", row + range.rowCount, ")"].join("");
    }

    function getFormulaTemplateForColumn(range) {
        // TODO: process w/ A1 & R1C1 style
        // funcName(1@:2@)
        var col = range.col;
        return [funcName.toUpperCase(), "(", getColumnName(col + 1), "@:", getColumnName(col + range.colCount), "@)"].join("");
    }

    function setFormulas(range, value, isRow) {
        var i, r, c, count, row, col, template;

        if (isRow) {
            count = sel.colCount;
            r = value;
            col = sel.col;
            template = getFormulaTemplateForRow(sel);
            for (i = 0; i < count; i++) {
                c = col + i;
                // TODO: skip empty column
                sheet.setFormula(r, c, template.replace(/@/g, getColumnName(c + 1)));
            }
        } else {
            count = sel.rowCount;
            c = value;
            row = sel.row;
            template = getFormulaTemplateForColumn(sel);
            for (i = 0; i < count; i++) {
                r = row + i;
                // TODO: skip empty row
                sheet.setFormula(r, c, template.replace(/@/g, r + 1));
            }
        }
    }

    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    sheet.isPaintSuspended(true);
    for (var n = 0; n < sels.length; n++) {
        var sel = getActualCellRange(sels[n], rowCount, columnCount),
            formulaPosition = getFormulaPosition(sel), resultPosition = formulaPosition.position,
            targetRow, targetCol;

        if (resultPosition) {
            if (resultPosition & ResultPosition.Bottom) {
                targetRow = sel.row + sel.rowCount + formulaPosition.rowOffset;
                // TODO: skip row with value
                setFormulas(sel, targetRow, true);
            }
            if (resultPosition & ResultPosition.Right) {
                targetCol = sel.col + sel.colCount + formulaPosition.colOffset;
                // TODO: skip col with value
                setFormulas(sel, targetCol, false);
            }
        }
    }
    sheet.isPaintSuspended(false);
}

function processRibbonDropDown(e, data) {
    function selectCellFormatDropDownItem(dropdown) {
        var $ul = $('ul.dropdown-menu', dropdown);

        if ($ul.length > 0) {
            $('li.selected', $ul).removeClass('selected');
            var formatter = getCellFormatter();
            $('li', $ul)
                .filter(function() { return $(this).data('value') == Ribbon.getDataAttributeString(formatter); })
                .addClass('selected');
        }
    }
    
    function insertTableSlicer(tableName, columnNames) {
        function getAutoSlicerName() {
            return "slicer" + Date.now();
        }
        
        spread.isPaintSuspended(true);
        var sheet = spread.getActiveSheet();
        var posX = 100, posY = 200;
        var slicer;
        columnNames.forEach(function(columnName) {
            var slicerName = getAutoSlicerName();
            slicer = sheet.addSlicer(slicerName, tableName, columnName);
            slicer.position(new GcSpread.Sheets.Point(posX, posY));
            posX += 30;
            posY += 30;
        });
        spread.isPaintSuspended(false);
        if (slicer) {
            slicer.isSelected(true);
        }
    }
    
    function fillInsertSlicerDropdownList($ul) {
        function processDropdownButtonClick(element) {
            var $this = $(element), tableName = _activeTable.name();

            if ($this.data("name") === "insert-slicer") {
                var columnNames = [];
                $this.closest('ul.dropdown-menu')
                    .find('li.item input:checked')
                    .each(function () {
                        columnNames.push($(this).data("name"));
                    });

                insertTableSlicer(tableName, columnNames);
            }
        }
        
        if (_activeTable && $ul.length) {
            // remove old items
            $("li.item", $ul).remove();
            
            // prepare submit buttons & bind event
            if ($ul.children().length === 0) {
                $(document).on('click', '.btn-dropdown', function() {
                    processDropdownButtonClick(this);
                });
                $ul.append($('<li class="divider"></li>'));
                $ul.append($('<li><div class="text-right"><button data-name="insert-slicer" class="btn btn-default btn-dropdown">OK</button><button class="btn btn-default btn-dropdown">Cancel</button></div></li>'));
            }
            // add new items
            var table = _activeTable, items = [];
            for(var i = 0, count = table.range().colCount; i < count; i++) {
                var name = table.getColumnName(i);
                items.push($('<li class="item"'  + 
                    '><label class="checkbox-inline"> <input type="checkbox" data-name="' + name + '"><span>' +
                    name + '</span></label></li>'));
            }
            $ul.prepend(items);
        }
    }

    var name = data.name, open = data.open;

    switch (name) {
        case 'cellformat':
            if (open && data.dropdown) {
                selectCellFormatDropDownItem(data.dropdown);
            }
            break;

        case 'cellsgroup':
            var selectionType = getSelectionType(spread.getActiveSheet().getSelections());
            //console.log('selectionType', selectionType);
            switch (selectionType) {
                case SelectionRangeType.Mixture:
                    alert('Mixed selection not supported!');

                    if (open && data.dropdown && data.originalEvent) {
                        //$(data.dropdown).removeClass('open');
                        data.originalEvent.preventDefault();
                        data.originalEvent.stopPropagation();
                    }
                    break;

                // TODO: disable items (add 'disabled' class to li) base on selectiontype

                default:
                    var found;
                    for (var key in SelectionRangeType) {
                        if (SelectionRangeType[key] === selectionType) {
                            found = key;
                            break;
                        }
                    }
                    //console.log('others', found || selectionType);
                    break;
            }
            break;

        case 'tableStyles':
            if (open) {
                var left = $("#inner-content-container").width() - ($tableStyleDropdown.parent().offset().left + $tableStyleDropdown.width()) - 10;

                // adjust drop down position to avoid out of view
                $tableStyleDropdown.css({left: left < 0 ? left : 0});
            }
            break;
            
        case 'insertSlicer':
            if (open) {
                // fill sclier
                fillInsertSlicerDropdownList($("ul.dropdown-menu", data.dropdown));
            }                
            break;
    }
}

function processRibbonDropDownShown(e, data) {
    var name = data.name;

    switch (name) {
        case 'tableStyles':
            // make selected items visible by scroll
            $("#tableStyles .table-format-item-selected")[0].scrollIntoView();
            break;
    }
}

function processTextChanged(e, data) {
    var name = data.name, value = data.value, $element = data.$element;
    var sheet = spread.getActiveSheet();

    switch (name) {
        case "tableName":
            if (value && _activeTable && value !== _activeTable.name()) {
                if (!sheet.findTableByName(value)) {
                    _activeTable.name(value);
                } else {
                    alert(getResource("messages.duplicatedTableName"));
                    $element.val(_activeTable.name());
                }
            }
            break;

        default:
            console.log("processTextChanged not process with ", name, value);
            break;
    }
}
// ribbon related items (end)

// setting pane related items
function displaySettingPane(title, $element, showCallback) {
    var $active = $("#setting-pane .pane-content > div:not(.hidden)");

    // same content do not need switch seeting pane
    if ($active[0] !== $element[0]) {
        $("#setting-pane .pane-title").text(title);
        $active.addClass("hidden");
        $element.removeClass("hidden");
    } else {
        // always update title since one setting content may have multi sub type such as celltype
        $("#setting-pane .pane-title").text(title);
    }

    if (showCallback) {
        showCallback();
    }

    showSettingPane();
}

function showSettingPane() {
    $("#setting-pane").show();
}

function hideSettingPane() {
    colorPicker.hide();
    $("#setting-pane").hide();
}

function attachSettingPaneEvents() {
    // checkbox
    $(".setting-container input[type='checkbox']")
        .click(checkedChanged);
        
    $(".setting-container input[type='radio']")
        .click(radioButtonClicked);

    // number input
    $(".setting-container input[type='number']")
        .blur(updateNumberProperty);

    // text input
    $(".setting-container input[type='text']")
        .blur(updateStringProperty);

    $("#setting-pane button.close").on('click', function () {
        hideSettingPane();
    });

    $(document).on('click', "#setting-pane button.dropdown-toggle", function () {
        colorPicker.hide();
    });

    // dropdown
    $(document).on('click', ".pane-row .dropdown-menu>li>a", function () {
        var $this = $(this), $group = $this.closest('.btn-group');

        var name = $group.data("name"),
            $text = $(this),
            dataValue = $text.data("value"),    // data-value includes both number value and string value, should pay attention when use it
            numberValue = +dataValue,
            text = $text.text(),
            value = text,
            nameValue = dataValue === 0 ? dataValue : ( dataValue || text );

            var $li = $this.parent(), $ul = $li.parent();

        // change undefined to null for value to be replaced, otherwise old value will be kept without change!!!
        $group.data({ value: dataValue || null, index: $(">li", $ul).index($li), $element: $this });
        $('span.content', $group).text(text);

        $('li', $ul).removeClass('selected');
        $li.addClass('selected');

        processDropDownClicked(name, numberValue, value, nameValue, $this, $group);
    });

    // comment decoration buttons
    $("#setting-pane button.btn-comment-decoration").on('click', function () {
        var $element = $(this);

        $element.toggleClass('toggle on');

        Actions.setCommentTextDecoration(spread, {comment: _activeComment, value: $element.data("value")});
    });

    $("#setSparklineButton").on('click', addSparklineEvent);
    $("#setCustomFormat").on('click', setCustomFormat);
}

function setDropDown(name, value) {
    var $container = $(".setting-container div[data-name='" + name + "']"),
        $items = $("ul>li>a", $container),
        $item = $items.filter(function () {
            // use == since string and number should be considered 
            return $(this).data("value") == value || $(this).text() == value;
        });
        
    var text = $item.text() || value; 

    $("span.content", $container).text(text);
    $container.data('value', value);
    
    $('li.selected', $container).removeClass('selected');
    if ($item.length > 0) {
        $($item[0]).parent().addClass('selected');        
    }
}

function prepareFunctionBuilder($container) {
    function filterFunctions($ul, s) {
        // restore all
        $ul.children().removeClass("hidden");
        // do filter when filter string provided
        if (s) {
            // to upper for match with function name
            s = s.toUpperCase();
            // filter functions by name, hide not matched items
            $("li.function-item", $ul).filter(function () {
                var funcName = $(this).text();
                return !funcName.startsWith(s);  // or don't use startsWith:  funcName.indexOf(s) != 0;
            }).addClass("hidden");

            // hide category without visible funtion items
            $("li.function-category", $ul).filter(function () {
                var name = $(this).data("name");
                return $("li.function-item[data-category='" + name + "']:visible", $ul).length === 0;
            }).addClass("hidden");
        }
    }

    function showFunctionDesctiption(name) {
        var item = $("#function-description function[name='" + name + "']")[0];

        if (item) {
            var $item = $(item), params = $item.attr('param'), description = $item.attr('description');
            var html = "<div class='function-description-container'><strong>" + [name, "(", params, ")"].join("") + "</strong></div> " + description;

            $("#functionBuiilder .function-description").html(html);
        }
    }

    function startEditWithFormula(sheet, rowIndex, colIndex, formula) {
        var oldGetFormulaInfoFun = GcSpread.Sheets.Sheet.prototype.getFormulaInformation;
        GcSpread.Sheets.Sheet.prototype.getFormulaInformation = function (row, col) {
            return {hasFormula: true};
        };
        sheet.startEdit(false, formula);
        GcSpread.Sheets.Sheet.prototype.getFormulaInformation = oldGetFormulaInfoFun;
        var editor = sheet.getCellType(rowIndex, colIndex).getEditingElement();
        editor.selectionStart = formula.length - 1;
        editor.selectionEnd = formula.length - 1;
    }

    function insertFormulaForEdit(funcName) {
        var sheet = spread.getActiveSheet();
        var activeRowIndex = sheet.getActiveRowIndex();
        var activeColumnIndex = sheet.getActiveColumnIndex();

        startEditWithFormula(sheet, activeRowIndex, activeColumnIndex, "=" + funcName + "()");
    }

    function getFunctionItem(funcName, index, array) {
        var $li = $("<li class='function-item'></li>");
        $li.attr('data-category', array.name);
        $li.text(funcName);

        return $li;
    }

    $ul = $(".function-list", $container);
    var categories = uiResource.functions.categories;

    for (var name in categories) {
        var cat = categories[name], text = cat.text, items = cat.items;
        var $li = $("<li class='function-category'></li>");
        $li.attr('data-name', name);
        $li.text(text);
        $ul.append($li);
        items.name = name;  // used to provide addtional information for array operator later
        $ul.append(items.map(getFunctionItem));
    }

    $("li.function-item", $ul).click(function () {
        var name = $(this).text();
        $('li.selected', $ul).removeClass("selected");
        $(this).addClass("selected");
        showFunctionDesctiption(name);
    });

    $("li.function-item", $ul).dblclick(function () {
        var name = $(this).text();

        insertFormulaForEdit(name);
    });

    $("input", $container)
        .attr("placeholder", uiResource.functions.setting.filterPlaceHolder)
        .keyup(function () {
            filterFunctions($ul, $(this).val());
        });
}

var _commonFormatters = [], _customFormatInput;
function prepareCellFormatSetting($container) {
    var $ul = $('div[data-name=commonFormat] ul.dropdown-menu', $container),
        $items = $("#home div[data-name=cellformat] ul.dropdown-menu>li");
    
    // copy dropdown builtin items
    var skip = false;
    $items.each(function(){
        if (skip) return;
        
        var $src = $(this),
            value = $src.attr("data-value"),
            text = $('button>label', $src).text();
            
        if (!value) { 
            skip = true;
            return;
        }
        _commonFormatters.push(value);
        var $li = $("<li><a></a></li>"),
            $a = $("a", $li);
        $a.attr("data-value", value).text(text);

        $ul.append($li);
    });
    // add custom
    $ul.append($('<li><a data-value="custom">'+ uiResource.cellTab.format.custom +'</a></li>'));
    $('li:first', $ul).addClass('default');

    _customFormatInput = $("#cellFormatSetting input[data-name=customFormat]");    
}

function syncCellFormat() {
    var formatter = getCellFormatter(), value = Ribbon.getDataAttributeString(formatter);
    
    if (formatter === "nullValue") {
        formatter = "";
    } else {
        if (_commonFormatters.indexOf(value) === -1) {
            value = 'custom';
        }
    }
    setDropDown('commonFormat', value);
    _customFormatInput.val(formatter);
}

function setCustomFormat() {
    var format = $('#cellFormatSetting input[data-name=customFormat]').val();

    Actions.setCellFormat(spread, {name: 'custom', value: format || null });    // map empty string to null (General)     
}

// find related items
var findCache;

function attachFindEvents() {
    var SearchFlags = spreadNS.SearchFlags;
    var SearchOrder = spreadNS.SearchOrder;
    var SearchFoundFlags = spreadNS.SearchFoundFlags;
    var SearchCondition = spreadNS.SearchCondition;
    var VerticalPosition = spreadNS.VerticalPosition;
    var HorizontalPosition = spreadNS.HorizontalPosition;

    function getSearchInformation() {
        var searchString = getTextValue("findwhat");
        if (searchString === "")
            return null;

        var withinWorksheet = $findin.first().is(":checked");
        var searchOrder = $searchby.first().is(":checked") ? SearchOrder.ZOrder : SearchOrder.NOrder;
        var searchFoundFlags;
        var searchFlags = 0;

        if (!getCheckValue("findMatchCase")) {
            searchFlags |= SearchFlags.IgnoreCase;
        }
        if (getCheckValue("findMatchExactly")) {
            searchFlags |= SearchFlags.ExactMatch;
        }
        if (getCheckValue("findUseWildcards")) {
            searchFlags |= SearchFlags.UseWildCards;
        }

        if ($lookin.first().is(":checked")) {
            searchFoundFlags = SearchFoundFlags.CellText;
        }
        else {
            searchFoundFlags = SearchFoundFlags.CellFormula;
            searchString = searchString.charAt(0) === "=" ? searchString.substr(1, searchString.length) : searchString;
        }

        return {
            WithinWorksheet: withinWorksheet,
            SearchString: searchString,
            SearchFlags: searchFlags,
            SearchOrder: searchOrder,
            SearchFoundFlags: searchFoundFlags
        };
    }

    function unparse(source, expr, row, col) {
        return spread.calcService.unparse(source, expr, row, col);
    }

    // modify specified SearchCondition for next search instead of recreate and reduce duplicted code
    // Modify as:
    //      1. reset findBegin* as the default value for new instance of SearchCondition;  
    //      2. Use options to provide key-value pair for need updated items 
    function modifySearchCondition(searchCondition, options) {
        if (searchCondition) {
            // reset findBegin*
            searchCondition.findBeginColumn = -1;
            searchCondition.findBeginRow = -1;
            // modify use options provided key-value pair when provided
            if (options) {
                for (var key in options) {
                    searchCondition[key] = options[key];
                }
            }
        }
    }

    function doFindAll(searchInformation) {
        function findAllInSheet(searchInformation, sheet) {
            var result = [];

            var startRow = 0, startColumn = 0,
                rowCount = sheet.getRowCount(), columnCount = sheet.getColumnCount(),
                endRow = rowCount - 1, endColumn = columnCount - 1;

            var searchCondition = new SearchCondition();

            searchCondition.searchString = searchInformation.SearchString;
            searchCondition.searchFlags = searchInformation.SearchFlags;
            searchCondition.searchOrder = searchInformation.SearchOrder;
            searchCondition.searchTarget = searchInformation.SearchFoundFlags;
            searchCondition.sheetArea = spreadNS.SheetArea.viewport;
            searchCondition.rowStart = startRow;
            searchCondition.columnStart = startColumn;
            searchCondition.rowEnd = endRow;
            searchCondition.columnEnd = endColumn;

            var res = sheet.search(searchCondition);

            var findRow = res.foundRowIndex,
                findColumn = res.foundColumnIndex;

            while (findRow != -1 && findColumn != -1) {
                var cell = sheet.getCell(findRow, findColumn);
                var expr = new Calc.Expressions.CellExpression(cell.row, cell.col, false, false);
                var cellName = unparse(null, expr, 0, 0);
                var item = {
                    sheetName: sheet.getName(),
                    cellName: cellName,
                    value: cell.text(),
                    formula: cell.formula(),
                    cellData: cell
                };

                result.push(item);

                if (searchInformation.SearchOrder == SearchOrder.ZOrder) {
                    startRow = findRow;
                    startColumn = findColumn + 1;
                    if (startColumn >= columnCount && startRow < rowCount) {
                        startRow = findRow + 1;
                        startColumn = 0;
                    }
                }
                else {
                    startRow = findRow + 1;
                    startColumn = findColumn;
                    if (startRow >= rowCount && startColumn < sheet.columnCount) {
                        startRow = 0;
                        startColumn = findColumn + 1;
                    }
                }

                modifySearchCondition(searchCondition, {rowStart: startRow, columnStart: startColumn});

                res = sheet.search(searchCondition);

                findRow = res.foundRowIndex;
                findColumn = res.foundColumnIndex;
            }
            return result;
        }

        function findAllInWorkbook(searchInformation, workbook) {
            var result = [];

            var startRow = 0,
                startColumn = 0,
                sheetCount = workbook.getSheetCount();

            var searchCondition = new SearchCondition();

            searchCondition.startSheetIndex = 0;
            searchCondition.endSheetIndex = sheetCount - 1;
            searchCondition.searchString = searchInformation.SearchString;
            searchCondition.searchFlags = searchInformation.SearchFlags;
            searchCondition.searchOrder = searchInformation.SearchOrder;
            searchCondition.searchTarget = searchInformation.SearchFoundFlags;
            searchCondition.sheetArea = spreadNS.SheetArea.viewport;

            var res = workbook.search(searchCondition);

            var findRow = res.foundRowIndex,
                findColumn = res.foundColumnIndex,
                findSheet = res.foundSheetIndex;

            while (findRow != -1 && findColumn != -1 ||
                    (findSheet < sheetCount && findSheet != -1)) {
                modifySearchCondition(searchCondition);

                var sheet = workbook.sheets[findSheet];
                var rowCount = sheet.getRowCount(), columnCount = sheet.getColumnCount();
                    
                if (findRow != -1 && findColumn != -1) {
                    

                    var cell = sheet.getCell(findRow, findColumn);

                    var expr = new Calc.Expressions.CellExpression(cell.row, cell.col, false, false);
                    var cellName = unparse(null, expr, 0, 0);

                    var item = {
                        sheetName: sheet.getName(),
                        cellName: cellName,
                        value: cell.text(),
                        formula: cell.formula(),
                        cellData: cell
                    };

                    result.push(item);

                    if (searchInformation.SearchOrder == SearchOrder.ZOrder) {
                        startRow = findRow;
                        startColumn = findColumn + 1;
                        if (startColumn >= columnCount &&
                            startRow < sheet.rowCount) {
                            startRow = findRow + 1;
                            startColumn = 0;
                        }
                    }
                    else {
                        startRow = findRow + 1;
                        startColumn = findColumn;
                        if (startRow >= rowCount &&
                            startColumn < columnCount) {
                            startRow = 0;
                            startColumn = findColumn + 1;
                        }
                    }
                    searchCondition.rowStart = startRow;
                    searchCondition.columnStart = startColumn;
                }
                else {
                    searchCondition.rowStart = 0;
                    searchCondition.columnStart = 0;
                    searchCondition.rowEnd = rowCount - 1;
                    searchCondition.columnEnd = columnCount - 1;
                }
                searchCondition.startSheetIndex = findSheet;
                searchCondition.endSheetIndex = findSheet;
                res = workbook.search(searchCondition);

                findRow = res.foundRowIndex;
                findColumn = res.foundColumnIndex;
                if (res.foundSheetIndex != -1) {
                    findSheet = res.foundSheetIndex;
                }
                else {
                    findSheet++;
                }
            }
            return result;
        }

        if (searchInformation.WithinWorksheet) {
            return findAllInSheet(searchInformation, spread.getActiveSheet());
        } else {
            return findAllInWorkbook(searchInformation, spread);
        }
    }

    function fillFindResult(results) {
        $resultContainer.empty();
        results.forEach(function (item) {
            var formula = item.formula;
            // add prefix "=" for formula
            if (formula) {
               item.formula = "=" + formula;  
            }
            $("<tr></tr>")
                .html(["sheetName", "cellName", "value", "formula"].map(function (name) {
                    return ["<td>", item[name], "</td>"].join("");
                }).join(""))
                .data("cell", item.cellData)
                .appendTo($resultContainer);
        });
        $resultCount.text(results.length + uiResource.find.result.countssuffix);
        if (results.length === 0) {
            // add an empty row to table when no matched items found
            $resultContainer.append($("<tr />"));
        }
        $('.findoutput', $container).show();
    }

    function findall() {
        var searchInformation = getSearchInformation();

        if (!searchInformation || !searchInformation.SearchString) {
            return;
        }

        var results = doFindAll(searchInformation);

        fillFindResult(results);
    }

    function doFindNext(searchInformation) {
        function getStartPosition(searchOrder, cellRange) {
            if (!cellRange) {
                return;
            }
            var row = cellRange.row, firstRow = row,
                col = cellRange.col, firstColumn = col,
                lastRow = row + cellRange.rowCount - 1;
            lastColummn = col + cellRange.colCount - 1;

            if (searchOrder == SearchOrder.ZOrder) {
                if (findCache.activeCellColumnIndex == -1 && findCache.activeCellRowIndex == -1) {
                    findCache.rowStart = 0;
                    findCache.columnStart = 0;
                }
                else if (findCache.activeCellColumnIndex < lastColummn) {
                    findCache.rowStart = findCache.activeCellRowIndex;
                    findCache.columnStart = findCache.activeCellColumnIndex + 1;//to do
                }
                else if (findCache.activeCellColumnIndex == lastColummn) {
                    findCache.rowStart = findCache.activeCellRowIndex + 1;
                    findCache.columnStart = 0;
                }
                else {
                    findCache.rowStart = firstRow;
                    findCache.columnStart = firstColumn;
                }
            }
            else // by columns
            {
                if (findCache.activeCellColumnIndex == -1 && findCache.activeCellRowIndex == -1) {
                    findCache.rowStart = 0;
                    findCache.columnStart = 0;
                }
                else if (findCache.activeCellRowIndex < lastRow) {
                    findCache.rowStart = findCache.activeCellRowIndex + 1;
                    findCache.columnStart = findCache.activeCellColumnIndex;
                }
                else if (findCache.activeCellRowIndex == lastRow) {
                    findCache.rowStart = 0;
                    findCache.columnStart = findCache.activeCellColumnIndex + 1;
                }
                else {
                    findCache.rowStart = firstRow;
                    findCache.columnStart = firstColumn;
                }
            }
        }

        function findWithinWorksheet(searchInformation, sheet) {
            var rowCount = sheet.getRowCount(), columnCount = sheet.getColumnCount(),
                endRow = rowCount - 1, endColumn = columnCount - 1;

            getStartPosition(searchInformation.SearchOrder, new spreadNS.Range(0, 0, rowCount, columnCount));

            var searchCondition = new SearchCondition();

            searchCondition.searchString = searchInformation.SearchString;
            searchCondition.searchFlags = searchInformation.SearchFlags;
            searchCondition.searchOrder = searchInformation.SearchOrder;
            searchCondition.searchTarget = searchInformation.SearchFoundFlags;
            searchCondition.sheetArea = spreadNS.SheetArea.viewport;
            searchCondition.rowStart = findCache.rowStart;
            searchCondition.columnStart = findCache.columnStart;
            searchCondition.rowEnd = endRow;
            searchCondition.columnEnd = endColumn;

            var result = sheet.search(searchCondition);

            var row = result.foundRowIndex,
                col = result.foundColumnIndex;

            findCache.findRowIndex = row;
            findCache.findColumnIndex = col;

            return row != -1 && col != -1;
        }

        function isWorksheetContains(searchInformation, sheet) {
            var findRow, findColumn;
            var searchCondition = new SearchCondition();

            searchCondition.searchString = searchInformation.SearchString;
            searchCondition.searchFlags = searchInformation.SearchFlags | SearchFlags.BlockRange;
            searchCondition.searchOrder = searchInformation.SearchOrder;
            searchCondition.searchTarget = searchInformation.SearchFoundFlags;
            searchCondition.sheetArea = spreadNS.SheetArea.viewport;

            var result = sheet.search(searchCondition);

            findRow = result.foundRowIndex;
            findColumn = result.foundColumnIndex;
            if (findRow != -1 && findColumn != -1) {
                return true;
            }

            return false;
        }

        function findNextWithinWorksheet(searchInformation, sheet) {
            findCache.findRowIndex = -1;
            findCache.findColumnIndex = -1;

            var found = findWithinWorksheet(searchInformation, sheet);

            if (found) {
                findCache.activeCellRowIndex = findCache.findRowIndex;
                findCache.activeCellColumnIndex = findCache.findColumnIndex;

                spread.getActiveSheet().addSelection(findCache.findRowIndex, findCache.findColumnIndex, 1, 1);
                spread.showActiveCell(VerticalPosition.nearest, HorizontalPosition.nearest);

                return true;
            }
            else {
                findCache.activeCellRowIndex = -1;
                findCache.activeCellColumnIndex = -1;

                return findWithinWorksheet(searchInformation, sheet);
            }
        }

        function getFindWorksheetList(withWorksheet) {
            var worksheetList = [];

            var startFindSheetIndex = findCache.activeSheetIndex,
                sheets = spread.sheets, sheetCount = sheets.length;
            for (var i = startFindSheetIndex; i < sheetCount; i++) {
                worksheetList.push(spread.sheets[i]);
            }

            for (var j = 0; j < startFindSheetIndex; j++) {
                worksheetList.push(sheets[j]);
            }

            return worksheetList;
        }

        function markFindCell(sheet, row, col) {
            sheet.setSelection(row, col, 1, 1);
            sheet.setActiveCell(row, col);
            spread.showActiveCell(VerticalPosition.nearest, HorizontalPosition.nearest);
            processSelectionChanged();
            fbx.text(sheet.getValue(row, col));     // TODO: need product support sync of formulabox
        }

        function findNextWithinWorksheets(searchInformation) {
            var worksheetList = getFindWorksheetList(searchInformation.WithinWorksheet);

            findCache.findRowIndex = -1;
            findCache.findColumnIndex = -1;
            findCache.findSheetIndex = -1;

            for (var i = 0; i < worksheetList.length; i++) {
                var worksheet = worksheetList[i];

                var sheetIndex = spread.sheets.indexOf(worksheet);

                if (sheetIndex != spread.getActiveSheetIndex()) {
                    findCache.activeCellRowIndex = -1;
                    findCache.activeCellColumnIndex = -1;
                }

                var found = findWithinWorksheet(searchInformation, worksheet);

                if (found) {
                    findCache.findSheetIndex = sheetIndex;
                    break;
                }
            }

            if (findCache.findSheetIndex != -1) {
                findCache.activeSheetIndex = findCache.findSheetIndex;
                var row = findCache.activeCellRowIndex = findCache.findRowIndex,
                    col = findCache.activeCellColumnIndex = findCache.findColumnIndex;

                spread.setActiveSheetIndex(findCache.findSheetIndex);
                markFindCell(spread.getActiveSheet(), row, col);
                
                return true;
            }
            else {
                return false;
            }
        }

        var found;
        if (searchInformation.WithinWorksheet) {
            var sheet = spread.getActiveSheet();

            if (!isWorksheetContains(searchInformation, sheet)) {
                findCache.findRowIndex = -1;
                findCache.findColumnIndex = -1;
                findCache.findSheetIndex = -1;

                return false;
            }

            found = findNextWithinWorksheet(searchInformation, sheet);
            if (found) {
                var col = findCache.activeCellColumnIndex = findCache.findColumnIndex,
                    row = findCache.activeCellRowIndex = findCache.findRowIndex;
                
                markFindCell(sheet, row, col);
            }
            findCache.findSheetIndex = spread.getActiveSheetIndex();
        }
        else {
            found = findNextWithinWorksheets(searchInformation);
        }

        return found;
    }

    function findnext() {
        var searchInformation = getSearchInformation();

        if (!searchInformation || !searchInformation.SearchString) {
            return;
        }

        var found = doFindNext(searchInformation);

        if (!found) {
            alert(uiResource.find.result.nomatch);
        }
    }

    var $container = $("#findOptions"),
        $findin = $("input[name='findin']", $container),
        $searchby = $("input[name='searchby']", $container),
        $lookin = $("input[name='lookin']", $container),
        $resultContainer = $(".resultlist tbody", $container),
        $resultCount = $(".resultcount", $container);

    $("#findall").click(findall);
    $("#findnext").click(findnext);
    $(document).on("click", ".resultlist tbody tr", function () {
        var cell = $(this).data("cell");
        var sheet = spread.getSheetFromName(cell.sheet.getName());
        spread.setActiveSheetIndex(spread.sheets.indexOf(sheet));
        spread.getActiveSheet().setActiveCell(cell.row, cell.col);
        spread.showActiveCell(VerticalPosition.nearest, HorizontalPosition.nearest);
    });
}
// find related items (end)

// print setting related items
var printFormatStringMapping = {
    currentPage: { value: "&P", text: "1" },
    totalPage: { value: "&N", text: "?" },
    currentDate: { value: "&D", getText: function() { var now = new Date(), text = now.getFullYear() + "/" + (now.getMonth() + 1) + "/" + now.getDate(); return text; } },
    currentTime: { value: "&T", getText: function() { var now = new Date(), text = now.getHours() + ":" + now.getMinutes() + ":" + now.getSeconds(); return text; } },
    workbookName: { value: "&F", getText: function() { return spread && spread.name || document && document.title; }},
    sheetName: { value: "&A", getText: function() { return spread && spread.getActiveSheet() && spread.getActiveSheet().getName() || "Sheet1"; }}
};
var printPairFormatPatterns = ["&B", "&I", "&U"];
var printSettingSectionNames = ["left", "center", "right"];
var printImageFormat = "&G";

app.printSettingSectionNames = printSettingSectionNames;

function getPreviewHTML(formatedText, imageName) {
    function addImageData(formatedText, imageName) {
        var ss = formatedText.split(printImageFormat);
        if (ss.length > 1) {
            if(imageName) {
                var url = app.uploadImages[imageName];
                ss[0] += '<img src="' + url + '" class="preview" />';
            }
            
            formatedText = ss.join("");
        }
        
        return formatedText;
    }
    
    if (!formatedText) {
        return "";    
    }
    
    var htmlClasses = ["text-bold", "text-italic", "text-underline"];

    printPairFormatPatterns.forEach(function (pattern, index) {
        var ss = formatedText.split(pattern), length = ss.length;
        if (length > 1) {
            for (var i = 1, process = true; i < length - 1; i++) {
                if (process) {
                    ss[i] = '<span class="' + htmlClasses[index] + '">' + ss[i] + '</span>';
                }
                process = !process;
            }
            formatedText = ss.join("");
        }
    });

    // add image
    formatedText = addImageData(formatedText, imageName);

    return "<span>" + formatedText + "</span>";
}

function insertText(input, text) {
    if (input.setRangeText) {
        input.setRangeText(text);
    }
    else if (input.selectionStart || input.selectionStart === 0) {
        var startPos = input.selectionStart, endPos = input.selectionEnd, value = input.value;
        
        input.value = value.substring(0, startPos) + text + value.substring(endPos, value.length);
        input.focus(); 
        input.selectionStart = startPos + text.length; 
        input.selectionEnd = input.selectionStart;
    } else {
        input.value += text; 
        input.focus();
    }
    input.blur();
}

app.insertText = insertText;

function updatePreview(options) {
    function updateCellsContent(cells, item) {
        printSettingSectionNames.forEach(function(name, index) {
            $(cells[index]).html(getPreviewHTML(item[name], item[getWithImageSuffix(name)]));
        });
    }
    if (options) {
        var item, cells,
            map = { header: "first", footer: "last" };
        
        for(var propertyName in map) {
            item = options[propertyName];    
            if (item) {
                cells = $("#print-setting-preview table>tbody>tr:" + map[propertyName] +">td");
                updateCellsContent(cells, item);
            }
        }
    }
}

app.updatePreview = updatePreview;

function getPreviewDisplayString(text) {
    function getDisplayText(map, value) {
        if (!value || !map.value) {
            return "";
        }

        if (!map.reg) {
            map.reg = new RegExp(map.value, "g");
        }

        return value.replace(map.reg, map.text || map.getText());
    }

    if (!text) {
        return "";
    }
    for (var propertyName in printFormatStringMapping) {
        var map = printFormatStringMapping[propertyName];
        text = getDisplayText(map, text);
    }

    return text;
}

app.getPreviewDisplayString = getPreviewDisplayString;

function getWithImageSuffix(name) {
    return name + "Image";
}

function createListItem(item, isBuiltin) {
    function removePairFormat(text) {
        printPairFormatPatterns.forEach(function (pattern) {
            text = text.split(pattern).join("");
        });

        return text;
    }

    var value = item.value, text, sections = {};

    if (value) {
        var values = [];
        printSettingSectionNames.forEach(function (name) {
            var format = value[name];
            if (format) {
                format = getPreviewDisplayString(format);

                sections[name] = format;
                var imageName = getWithImageSuffix(name);
                sections[imageName] = value[imageName];
                values.push(removePairFormat(format));
            }
        });
        text = values.join(", ");
    } else {
        text = item.text;
    }

    var $li = $("<li></li>"), $a = $("<a></a>");

    $a.text(text).data({ value: value, sections: sections }).appendTo($li);

    if (isBuiltin) {
        $li.addClass("builtin");
    }

    return $li;
}

app.initPrintSetting = function() {
    if ($("#printHeaderList").children().length === 0) {
        app.isInit = true;
        fillItems($("#printHeaderList"), uiResource.printSetting.options.headerAndFooter.header.items);
        fillItems($("#printFooterList"), uiResource.printSetting.options.headerAndFooter.footer.items);
        fillDropdownList($("#textContentList"), uiResource.printSetting.options.headerAndFooter.custom.items);
        fillDropdownList($("#imageContentList"), [ { text: uiResource.printSetting.options.headerAndFooter.custom.image.emptylist } ]);
        app.isInit = false;
    }
   
    function fillItems($ul, items) {
        items.forEach(function(item){
            $ul.append(createListItem(item, true));
        });
        $("li:first>a", $ul).click();
    }
    
    function fillDropdownList($ul, items) {
        function createListItem(item) {
            var $li = $("<li></li>"), $a = $("<a></a>");
            
            $a.text(item.text).data("value", item.value).appendTo($li);
            
            return $li;
        }
        
        items.forEach(function(item) {
            $ul.append(createListItem(item));
        });
        $("li:first>a", $ul).click();
    }    
};

app.printSpread = function(activeSheetOnly, printInfo) {
    if (activeSheetOnly) {
        var sheet = spread.getActiveSheet();
        sheet.printInfo(printInfo);
        spread.print(spread.getActiveSheetIndex());
    } else {
        // same printInfo for all sheets, you can set one by one with different setting
        spread.sheets.forEach(function(sheet){
            sheet.printInfo(printInfo);
        });
        spread.print();
    }
};
// print setting related items (end)

// setting pane related items (end)

app.reset = function(noDestroy) {
    if(!noDestroy && spread) {
        spread.destroy();
        spread = null;
        
        spread = new GcSpread.Sheets.Spread($("#ss")[0]);
    }
    
    attachSpreadEvents(true);
    
    fbx.spread(spread);
    
    hideSettingPane();
};
