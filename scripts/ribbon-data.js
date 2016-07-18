var data = {
    "selected": 1,
    "tabs": [
        {
            "title": "File",
            "name": "file"
        },
        {
            "title": "Home",
            "name": "home",
            "collapse": ["groupStyle2", "groupFont", "groupAlignment", "*groupStyle3", "groupStyle1"],
            "groupCollapseItems": ["groupStyle1"],
            "groups": [
              {
                  "tools": [
                    {
                        "type": "group",
                        "collapseGroup": "groupFont",
                        "centerAlign": true,
                        "items": [
                            {
                                "type": "input-group",
                                "name": "fontFamily",
                                "width": 70,
                                "dropdown": ["Arial", "Arial Black", "Calibri", "Cambria", "Century", "Courier New", "Comic Sans MS", "Garamond", "Georgia", "Malgun Gothic", "Mangal", "Meiryo", "MS Gothic", "MS Mincho", "MS PGothic", "MS PMincho", "Roboto", "Tahoma", "Times", "Times New Roman", "Trebuchet MS", "Verdana", "Wingdings"]
                            },
                            {
                                "type": "input-group",
                                "name": "fontSize",
                                "width": 40,
                                "dropdown": [8, 9, 10, 11, 12, 13, 14, 15, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
                            },
                        ]
                    },
                    {
                        "type": "icon",
                        "iconClass": "sprite Bold",
                        "name": "bold",
                        "text": "Bold",
                        "tooltip": "Bold",
                        "collapseGroup": "groupStyle1",
                        "toggle": true,
                        "nolabel": true
                    },
                    {
                        "type": "icon",
                        "iconClass": "sprite Italic",
                        "name": "italic",
                        "text": "Italic",
                        "tooltip": "Italic",
                        "collapseGroup": "groupStyle1",
                        "toggle": true,
                        "nolabel": true
                    },
                    {
                        "type": "icon",
                        "iconClass": "sprite Underline",
                        "name": "underline",
                        "text": "Underline",
                        "tooltip": "Underline",
                        "collapseGroup": "groupStyle1",
                        "toggle": true,
                        "nolabel": true
                    },
                    {
                        "type": "group",
                        "collapseGroup": "groupStyle3",
                        "tooltip": "Borders",
                        "items": [
                            {
                                "type": "icon-group",
                                "iconClass": "sprite BorderBottomNoToggle",
                                "name": "border",
                                "text": "Borders",
                                "header": "Borders",
                                "minWidth": 280,
                                "dropdown": [
                                    { "value": "bottom", "text": "Bottom Border", "iconClass": "sprite BorderBottom" },
                                    { "value": "top", "text": "Top Border", "iconClass": "sprite BorderTop" },
                                    { "value": "left", "text": "Left Border", "iconClass": "sprite BorderLeft" },
                                    { "value": "right", "text": "Right Border", "iconClass": "sprite BorderRight" },
                                    "",
                                    { "value": "none", "text": "No Border", "iconClass": "sprite BorderNone" },
                                    { "value": "all", "text": "All Borders", "iconClass": "sprite BordersAll" },
                                    { "value": "outside", "text": "Outside Borders", "iconClass": "sprite BorderOutside" },
                                    { "value": "thick", "text": "Thick Box Border", "iconClass": "sprite BorderThickOutside" },
                                    "",
                                    { "value": "doublebottom", "text": "Bottom Double Border", "iconClass": "sprite BorderDoubleBottom" },
                                    { "value": "thickbottom", "text": "Thick Bottom Border", "iconClass": "sprite BorderThickBottom" },
                                    { "value": "top-bottom", "text": "Top and Bottom Border", "iconClass": "sprite BorderTopAndBottom" },
                                    { "value": "top-thickbottom", "text": "Top and Thick Bottom Border", "iconClass": "sprite BorderTopAndThickBottom" },
                                    { "value": "top-doublebottom", "text": "Top and Double Bottom Border", "iconClass": "sprite BorderTopAndDoubleBottom" },
                                    "",
                                    { "value": "more", "text": "More Borders...", "iconClass": "sprite BordersMoreDialog"}
                                ]
                            }
                        ]
                    },
                    {
                        "type": "group",
                        "collapseGroup": "groupStyle3",
                        "tooltip": "Fill Color",
                        "items": [
                            {
                                "type": "setcolor-group",
                                "iconClass": "sprite FillBackColorSplitDropdown",
                                "name": "backColor",
                                "text": "Fill Color",
                                "colorPickerOptions": { "nofill": { "show": true, "text": "No Fill", "color": "white"}, "header": "Fill Color" }
                            }
                        ]
                    },
                    {
                        "type": "group",
                        "collapseGroup": "groupStyle3",
                        "tooltip": "Font Color",
                        "items": [
                            {
                                "type": "setcolor-group",
                                "iconClass": "sprite GroupBasicText",
                                "name": "foreColor",
                                "text": "Font Color",
                                "colorPickerOptions": { "autocolor": { "show": false, "text": "Automatic", "color": "black"}, "header": "Font Color" }
                            }
                        ]
                    },
                    {
                        "type": "icon",
                        "iconClass": "sprite Overline",
                        "name": "overline",
                        "text": "Overline",
                        "tooltip": "Overline",
                        "collapseGroup": "groupStyle2",
                        "toggle": true
                    },
                    {
                        "type": "icon",
                        "iconClass": "sprite Strikethrough",
                        "name": "strikethrough",
                        "text": "Strikethrough",
                        "tooltip": "Strikethrough",
                        "collapseGroup": "groupStyle2",
                        "toggle": true
                    },
                    {
                        "type": "dropdown",
                        "minWidth": 210,
                        "iconClass": "glyphicon glyphicon-menu-down",
                        "name": "font",
                        "header": "Font",
                        "items": []
                    }
                  ]
              },
              {
                  "tools": [
                    {
                        "type": "dropdown",
                        "iconClass": "sprite AlignCenter",
                        "name": "align",
                        "tooltip": "Alignment",
                        "header": "Alignment",
                        "items": ["indent", "outdent"],
                        "rows": [
                            {
                                "type": "icon-group",
                                "items": [
                                    {
                                        "iconClass": "sprite AlignTopExcel",
                                        "name": "valign-top",
                                        "text": "Top",
                                        "toggle": true,
                                        "nolabel": true,
                                        "toggleGroup": "valign"
                                    },
                                    {
                                        "iconClass": "sprite AlignMiddleExcel",
                                        "name": "valign-middle",
                                        "text": "Middle",
                                        "toggle": true,
                                        "nolabel": true,
                                        "toggleGroup": "valign"
                                    },
                                    {
                                        "iconClass": "sprite AlignBottomExcel",
                                        "name": "valign-bottom",
                                        "text": "Bottom",
                                        "toggle": true,
                                        "nolabel": true,
                                        "toggleGroup": "valign"
                                    }
                                ]
                            },
                            {
                                "type": "icon-group",
                                "items": [
                                    {
                                        "iconClass": "sprite AlignLeft",
                                        "name": "halign-left",
                                        "text": "Left",
                                        "toggle": true,
                                        "nolabel": true,
                                        "toggleGroup": "halign"
                                    },
                                    {
                                        "iconClass": "sprite AlignCenter",
                                        "name": "halign-center",
                                        "text": "Center",
                                        "toggle": true,
                                        "nolabel": true,
                                        "toggleGroup": "halign"
                                    },
                                    {
                                        "iconClass": "sprite AlignRight",
                                        "name": "halign-right",
                                        "text": "Right",
                                        "toggle": true,
                                        "nolabel": true,
                                        "toggleGroup": "halign"
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        "type": "icon",
                        "iconClass": "sprite MergeCenter",
                        "name": "cellmerge",
                        "text": "Merge & Center",
                        "collapseGroup": "groupAlignment",
                        "toggle": true
                    },
                    {
                        "type": "icon",
                        "iconClass": "sprite WrapText",
                        "name": "wordwrap",
                        "text": "Wrap Text",
                        "collapseGroup": "groupAlignment",
                        "toggle": true
                    },
                    {
                        "type": "icon",
                        "iconClass": "sprite IndentIncrease",
                        "name": "indent",
                        "text": "Increase Indent",
                        "collapseGroup": "groupIndent"
                    },
                    {
                        "type": "icon",
                        "iconClass": "sprite IndentDecrease",
                        "name": "outdent",
                        "text": "Decrease Indent",
                        "collapseGroup": "groupIndent"
                    }
                  ]
              },
              {
                  "tools": [
                    {
                        "type": "group",
                        "tooltip": "Cell Format",
                        "items": [
                            {
                                "type": "dropdown-only",
                                "iconClass": "sprite DollarSign",
                                "name": "cellformat",
                                "text": "Cell Format",
                                "dropdown": [
                                    { "value": "nullValue", "text": "General", "iconClass": "sprite DataTypeGeneral" },
                                    { "value": "0.00", "text": "Number", "iconClass": "sprite DataTypeNumber" },
                                    { "value": "$#,##0.00", "text": "Currency", "iconClass": "sprite DataTypeCurrency" },
                                    { "value": "$ #,##0.00;$ (#,##0.00);$ \"-\"??;@", "text": "Accounting", "iconClass": "sprite DataTypeCurrencyBasic" },
                                    { "value": "m/d/yyyy", "text": "Short Date", "iconClass": "sprite DataTypeShortDate" },
                                    { "value": "dddd, mmmm dd, yyyy", "text": "Long Date", "iconClass": "sprite DataTypeLongDate" },
                                    { "value": "h:mm:ss AM/PM", "text": "Time", "iconClass": "sprite DataTypeTime" },
                                    { "value": "0%", "text": "Percentage", "iconClass": "sprite PercentStyle" },
                                    { "value": "# ?/?", "text": "Fraction", "iconClass": "sprite DataTypeStandard" },
                                    { "value": "0.00E+00", "text": "Scientific", "iconClass": "sprite DataTypeScientific" },
                                    { "value": "@", "text": "Text", "iconClass": "sprite DataTypeText" },
                                    "",
                                    {"value":"custom", "text": "Custom Format"}
                                ]
                            }
                        ]
                    },
                    {
                        "type": "group",
                        "tooltip": "Number Format",
                        "items": [
                            {
                                "type": "dropdown-only",
                                "iconClass": "sprite PercentStyle",
                                "name": "numberformat",
                                "text": "Number Format",
                                "dropdown": [
                                    { "value": "percentStyle", "text": "Percent Style", "iconClass": "sprite PercentStyle" },
                                    { "value": "commaStyle", "text": "Comma Style", "iconClass": "sprite CommaStyle" },
                                    { "value": "increaseDecimal", "text": "Increase Decimal", "iconClass": "sprite DecimalsIncrease" },
                                    { "value": "decreaseDecimal", "text": "Decrease Decimal", "iconClass": "sprite DecimalsDecrease" }
                                ]
                            }
                        ]
                    }
                  ]
              },
              {
                  "tools": [
                    {
                        "type": "group",
                        "tooltip": "Cell Type",
                        "items": [
                            {
                                "type": "dropdown-only",
                                "iconClass": "sprite CellProperties",
                                "name": "celltype",
                                "text": "Cell Type",
                                "dropdown": [
                                    { "value": "button", "text": "Button CellType" },
                                    { "value": "checkbox", "text": "Checkbox CellType" },
                                    { "value": "combobox", "text": "Combobox CellType" },
                                    { "value": "hyperlink", "text": "Hyperlink CellType" }
                                ]
                            }
                        ]
                    }
                  ]
              },
              {
                  "tools": [
                    {
                        "type": "icon",
                        "iconClass": "sprite Lock",
                        "name": "protectsheet",
                        "text": "Protect Sheet",
                        "tooltips": ["Protect Sheet", "Unprotect Sheet"],
                        "toggle": true
                    },
                    {
                        "type": "icon",
                        "iconClass": "sprite Lock",
                        "name": "unlockcells",
                        "text": "Unlock Cells",
                        "tooltips": ["Lock Cells", "Unlock Cells"],
                        "toggle": true,
                        "checked": true
                    }
                  ]
              },
              {
                  "tools": [
                    {
                        "type": "group",
                        "tooltip": "Insert & Delete",
                        "items": [
                            {
                                "type": "dropdown-only",
                                "iconClass": "sprite InsertCellMenu",
                                "name": "cellsgroup",
                                "text": "Insert & Delete",
                                "dropdown": [
                                    { "value": "insertRows", "text": "Insert Rows", "iconClass": "sprite InsertRow" },
                                    { "value": "insertColumns", "text": "Insert Columns", "iconClass": "sprite InsertColumns" },
                                    { "value": "insert-shiftCellsRight", "text": "Insert Cells and Shift Right", "iconClass": "sprite InsertColumnLeft" },
                                    { "value": "insert-shiftCellsDown", "text": "Insert Cells and Shift Down", "iconClass": "sprite InsertRowAbove" },
                                    { "value": "deleteRows", "text": "Delete Rows", "iconClass": "sprite CellsDelete" },
                                    { "value": "deleteColumns", "text": "Delete Columns", "iconClass": "sprite CellsDeleteSmart" },
                                    { "value": "delete-shiftCellsLeft", "text": "Delete Cells and Shift Left", "iconClass": "sprite CellsDelete" },
                                    { "value": "delete-shiftCellsUp", "text": "Delete Cells and Shift Up", "iconClass": "sprite CellsDeleteSmart" }
                                ]
                            }
                        ]
                    },
                    {
                        "type": "group",
                        "tooltip": "Clear",
                        "items": [
                            {
                                "type": "dropdown-only",
                                "iconClass": "sprite ClearMenu",
                                "name": "clearformat",
                                "text": "Clear",
                                "dropdown": [
                                    { "value": "clearAll", "text": "Clear All", "iconClass": "sprite ClearAll" },
                                    { "value": "clearFormatting", "text": "Clear Formatting", "iconClass": "sprite ClearAllFormatting" },
                                    { "value": "clear", "text": "Clear", "iconClass": "sprite Clear" }
                                ]
                            }
                        ]
                    }
                  ]
              },
              {
                  "tools": [
                    {
                        "ignore": true,
                        "type": "icon",
                        "iconClass": "sprite Clear",
                        "name": "conditionalformat",
                        "text": "Conditional Formatting"
                    }
                  ]
              },
              {
                  "tools": [
                    {
                        "type": "icon",
                        "iconClass": "sprite SearchUI",
                        "name": "find",
                        "text": "Find"
                    }
                  ]
              }
            ]
        },
        {
            "title": "Insert",
            "name": "insert",
            "collapse": ["*groupInsert"],
            "groups": [
                {
                    "tooltip": "Insert",
                    "tools": [
                        {
                            "type": "icon",
                            "iconClass": "sprite InsertTable icon-text",
                            "name": "insertTable",
                            "text": "Table",
                            "collapseGroup": "groupInsert"
                        },
                        {
                            "type": "icon",
                            "iconClass": "sprite InsertPictureDialog icon-text",
                            "name": "insertPicture",
                            "text": "Pictures",
                            "collapseGroup": "groupInsert"
                        },
                        {
                            "type": "icon",
                            "iconClass": "sprite InsertLink icon-text",
                            "name": "insertLink",
                            "text": "Link",
                            "header": "Insert Link",
                            "collapseGroup": "groupInsert"
                        },
                        {
                            "type": "icon",
                            "iconClass": "sprite InsertNewComment icon-text",
                            "name": "insertComment",
                            "text": "Comment",
                            "collapseGroup": "groupInsert"
                        },
                        {
                            "type": "group",
                            "tooltip": "Sparklines",
                            "collapseGroup": "groupInsert",
                            "items": [
                                {
                                    "type": "dropdown-only",
                                    "iconClass": "sprite SparklineLineInsert icon-text",
                                    "name": "insertSparkline",
                                    "haslabel": true,
                                    "text": "Sparklines",
                                    "dropdown": [
                                        { "value": "line", "text": "Line Sparkline" },
                                        { "value": "column", "text": "Column Sparkline" },
                                        { "value": "winloss", "text": "Win/Loss Sparkline" },
                                        { "value": "pie", "text": "Pie Sparkline" },
                                        { "value": "area", "text": "Area Sparkline" },
                                        { "value": "scatter", "text": "Scatter Sparkline" },
                                        { "value": "spread", "text": "Spread Sparkline" },
                                        { "value": "stacked", "text": "Stacked Sparkline" },
                                        { "value": "boxplot", "text": "BoxPlot Sparkline" },
                                        { "value": "cascade", "text": "Cascade Sparkline" },
                                        { "value": "pareto", "text": "Pareto Sparkline" },
                                        { "value": "bullet", "text": "Bullet Sparkline" },
                                        { "value": "hbar", "text": "Hbar Sparkline" },
                                        { "value": "vbar", "text": "Vbar Sparkline" },
                                        { "value": "vari", "text": "Variance Sparkline" }
                                    ]
                                }
                            ]
                        },
                        {
                            "type": "dropdown",
                            "iconClass": "glyphicon glyphicon-menu-down",
                            "name": "insertDropdown",
                            "header": "Insert",
                            "items": []
                        }
                    ]
                }
            ]
        },
        {
            "title": "Formulas",
            "name": "formulas",
            "collapse": ["*groupFormula"],
            "groups": [
              {
                  "tooltip": "Formulas",
                  "tools": [
                    {
                        "ignore": true,
                        "type": "group",
                        "tooltip": "Sum",
                        "collapseGroup": "groupFormula",
                        "items": [
                                {
                                    "type": "icon-group",
                                    "iconClass": "glyphicon glyphicon-usd icon-text",
                                    "name": "autoSum",
                                    "haslabel": true,
                                    "text": "AutoSum",
                                    "header": "AutoSum",
                                    "dropdown": [
                                        { "value": "sum", "text": "Sum" },
                                        { "value": "average", "text": "Average" },
                                        { "value": "count", "text": "Count Numbers" },
                                        { "value": "max", "text": "Max" },
                                        { "value": "min", "text": "Min" }
                                    ]
                                }
                        ]
                    },
                    {
                        "type": "icon",
                        "iconClass": "sprite Formula icon-text",
                        "name": "insertFormula",
                        "text": "Insert Formula",
                        "collapseGroup": "groupFormula"
                    },
                    {
                        "ignore": true,
                        "type": "icon",
                        "iconClass": "glyphicon glyphicon-paperclip icon-text",
                        "name": "nameManager",
                        "text": "Name Manager",
                        "collapseGroup": "groupFormula"
                    },
                    {
                        "type": "dropdown",
                        "iconClass": "glyphicon glyphicon-menu-down",
                        "name": "setFormula",
                        "header": "Formulas",
                        "items": []
                    }
                  ]
              }, {
                  "tooltip": "Calculate Now",
                  "tools": [
                    {
                        "type": "icon",
                        "iconClass": "sprite CalculationOptionsMenu icon-text",
                        "name": "calculateNow",
                        "text": "Calculate Now"
                    }
                  ]
              }
            ]
        },
        {
            "title": "Data",
            "name": "data",
            "collapse": ["groupDetail", "groupSort", "groupGroup"],
            "groups": [
              {
                  "tooltip": "Sort / Filter",
                  "tools": [
                      {
                          "type": "icon",
                          "iconClass": "sprite SortAscendingWord",
                          "name": "sortAZ",
                          "text": "Sort Ascending",
                          "collapseGroup": "groupSort"
                      },
                      {
                          "type": "icon",
                          "iconClass": "sprite SortDescendingWord",
                          "name": "sortZA",
                          "text": "Sort Descending",
                          "collapseGroup": "groupSort"
                      },
                      {
                          "type": "icon",
                          "iconClass": "sprite FiltersMenu",
                          "name": "filter",
                          "text": "Show Filter Buttons",
                          "collapseGroup": "groupSort"
                      },
                      {
                        "type": "dropdown",
                        "iconClass": "sprite SortFilterMenu",
                        "name": "sortAndFilter",
                        "text": "Sort & Fiter",
                        "header": "Sort & Fiter",
                        "items": []
                    }
                  ]
              },
              {
                  "tooltip": "Group",
                  "tools": [
                      {
                          "type": "icon",
                          "iconClass": "sprite OutlineGroup",
                          "name": "group",
                          "text": "Group",
                          "collapseGroup": "groupGroup"
                      },
                      {
                          "type": "icon",
                          "iconClass": "sprite OutlineUngroup",
                          "name": "ungroup",
                          "text": "Ungroup",
                          "collapseGroup": "groupGroup"
                      },
                      {
                          "type": "icon",
                          "iconClass": "sprite ShowDetailsPage",
                          "name": "showDetail",
                          "text": "Show Detail",
                          "collapseGroup": "groupDetail"
                      },
                      {
                          "type": "icon",
                          "iconClass": "sprite HideDetails",
                          "name": "hideDetail",
                          "text": "Hide Detail",
                          "collapseGroup": "groupDetail"
                      },
                      {
                          "type": "checkbox",
                          "checked": true,
                          "text": "Summary rows below detail",
                          "name": "summaryBelow",
                          "collapseGroup": "groupDetailSummary"
                      },
                      {
                          "type": "checkbox",
                          "checked": true,
                          "text": "Summary columns to the right of detail",
                          "name": "summaryRight",
                          "collapseGroup": "groupDetailSummary"
                      },
                    {
                        "type": "dropdown",
                        "iconClass": "sprite GroupTableStyleOptions",
                        "altIconClass": "glyphicon glyphicon-menu-right",
                        "name": "groupSetting",
                        "text": "Group Setting",
                        "header": "Group Setting",
                        "minWidth": "300px",
                        "items": ["summaryBelow", "summaryRight"]
                    }
                  ]
              },
              {
                  "tooltip": "Data Validation",
                  "tools": [
                      {
                          "type": "icon",
                          "iconClass": "sprite DataValidationCircleInvalid icon-text",
                          "name": "circleInvalidData",
                          "text": "Circle Invalid Data",
                          "toggle": true
                      },
                      {
                          "type": "icon",
                          "iconClass": "sprite DataValidation icon-text",
                          "name": "selectValidator",
                          "text": "Select Validator",
                          "header": "Data Validator"
                      }
                  ]
              }
            ]
        },
        {
            "title": "View",
            "name": "view",
            "collapse": ["*groupShow"],
            "groups": [
              {
                  "tooltip": "Show / Hide",
                  "tools": [
                    {
                        "type": "checkbox",
                        "checked": true,
                        "text": "Formula Bar",
                        "tooltip": "Show Formula Bar",
                        "collapseGroup": "groupShow",
                        "name": "showFormulaBar"
                    },
                    {
                        "type": "checkbox",
                        "checked": true,
                        "text": "Gridlines",
                        "tooltip": "View Gridlines",
                        "collapseGroup": "groupShow",
                        "name": "showGridlines"
                    },
                    {
                        "type": "checkbox",
                        "checked": true,
                        "text": "Headings",
                        "tooltip": "View Headings",
                        "collapseGroup": "groupShow",
                        "name": "showHeadings"
                    },
                    {
                        "type": "checkbox",
                        "checked": true,
                        "text": "Sheet Tabs",
                        "tooltip": "Show Sheet Tabs",
                        "collapseGroup": "groupShow",
                        "name": "showSheetTabs"
                    },
                    {
                        "type": "dropdown",
                        "iconClass": "glyphicon glyphicon-menu-down",
                        "name": "showHideDropdown",
                        "header": "Show / Hide",
                        "items": []
                    }
                  ]
              },
              {
                  "tools": [
                      {
                        "type": "group",
                        "items": [
                            {
                                "type": "dropdown-only",
                                "iconClass": "sprite FreezePanes icon-text",
                                "name": "freezePanes",
                                "haslabel": true,
                                "text": "Freeze Panes",
                                "dropdown": [
                                    { "value": "freezePanes", "text": "Freeze Panes" },
                                    { "value": "freezeTopRow", "text": "Freeze Top Row" },
                                    { "value": "freezeFirstColumn", "text": "Freeze First Column" },
                                    { "value": "freezeBottomRow", "text": "Freeze Bottom Row" },
                                    { "value": "freezeLastColumn", "text": "Freeze Last Column" },
                                    "",
                                    { "value": "unfreeze", "text": "Unfreeze Panes" }
                                ]
                            }
                        ]
                      }
                  ]
              }
            ]
        },
        {
            "title": "Table",
            "hidden": false,
            "name": "table",
            "groups": [
                {
                    tools: [
                        {
                            "type": "input",
                            "name": "tableName",
                            "text": "Name"
                        }
                    ]
                },
                {
                    "tools": [
                        {
                            "type": "group",
                            "tooltip": "Insert Slicer",
                            "items": [
                                {
                                    "type": "dropdown-only",
                                    "iconClass": "glyphicon glyphicon-filter icon-text",
                                    "name": "insertSlicer",
                                    "text": "Insert Slicer",
                                    "haslabel": true,
                                    "dropdown": [
                                    ]
                                }
                            ]
                        }
                    ]
                },
                {
                    "tools": [
                        {
                            "type": "group",
                            "tooltip": "Style Options",
                            "items": [
                                {
                                    "type": "dropdown-only",
                                    "iconClass": "glyphicon glyphicon-info-sign icon-text",
                                    "name": "tableOption",
                                    "text": "Style Options",
                                    "haslabel": true,
                                    "dropdown": [
                                        { "value": "tableHeaderRow", "text": "Header Row", "toggle": true, "checked": true },
                                        { "value": "tableTotalRow", "text": "Total Row", "toggle": true },
                                        { "value": "tableFirstColumn", "text": "First Column", "toggle": true },
                                        { "value": "tableLastColumn", "text": "Last Column", "toggle": true },
                                        { "value": "tableBandedRows", "text": "Banded Rows", "toggle": true, "checked": true },
                                        { "value": "tableBandedColumns", "text": "Banded Columns", "toggle": true },
                                        { "value": "tableFilterButton", "text": "Filter Button", "toggle": true }
                                    ]
                                }
                            ]
                        }
                    ]
                },
                {
                    "tools": [
                        {
                            "type": "group",
                            "tooltip": "Table Styles",
                            "items": [
                                {
                                    "type": "icon-group",
                                    "iconClass": "glyphicon glyphicon-th",
                                    "name": "tableStyles",
                                    "text": "Table Styles",
                                    "header": "Table Styles",
                                    "dropdown": [
                                    ]
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    ]
};
