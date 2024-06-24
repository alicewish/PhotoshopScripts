/*****************************************************************
 *
 * 感叹号转斜体脚本
 *
 * 作者：墨问非名
 *
 *****************************************************************
 *
 * 使用说明：
 *
 * 1、脚本位置
 * # Windows
 * "C:\Program Files\Adobe\Adobe Photoshop CC 2024\Presets\Scripts"
 *
 * # Mac
 * "/Applications/Adobe Photoshop CC 2024/Presets/Scripts"
 *
 * 2、使用方法
 * 对当前psd生效
 *
 *****************************************************************/

//================初始化标尺、字体单位设置================
var originalUnit = preferences.rulerUnits;
var originalTypeUnit = preferences.typeUnits;
preferences.rulerUnits = Units.PIXELS;
preferences.typeUnits = TypeUnits.PIXELS;
app.preferences.rulerUnits = Units.PIXELS;
app.preferences.typeUnits = TypeUnits.PIXELS;
app.preferences.pointSize = PointType.POSTSCRIPT;

const idlayer = stringIDToTypeID("layer");
const idtextLayer = stringIDToTypeID("textLayer");
const idordinal = stringIDToTypeID("ordinal");
const idtargetEnum = stringIDToTypeID("targetEnum");
const idallEnum = stringIDToTypeID("allEnum");

const idnull = stringIDToTypeID("null");

const idfrom = charIDToTypeID("From");
const idto = stringIDToTypeID("to");

const idsize = stringIDToTypeID("size");
const idpixelsUnit = stringIDToTypeID("pixelsUnit");

const idfontPostScriptName = stringIDToTypeID("fontPostScriptName");

const idcolor = stringIDToTypeID("color");
const idRd = stringIDToTypeID("red");
const idGrn = stringIDToTypeID("grain");
const idBl = stringIDToTypeID("blue");
const idRGBColor = stringIDToTypeID("RGBColor");

const idsyntheticBold = stringIDToTypeID("syntheticBold");
const idsyntheticItalic = stringIDToTypeID("syntheticItalic");
const idunderline = stringIDToTypeID("underline");
const idunderlineOnLeftInVertical = stringIDToTypeID("underlineOnLeftInVertical");

const idtextStyle = stringIDToTypeID("textStyle");
const idtextStyleRange = stringIDToTypeID("textStyleRange");

const idset = stringIDToTypeID("set");

const idfind = stringIDToTypeID("find");
const idreplace = stringIDToTypeID("replace");
const idcheckAll = stringIDToTypeID("checkAll");
const idforward = stringIDToTypeID("forward");
const idcaseSensitive = stringIDToTypeID("caseSensitive");
const idwholeWord = stringIDToTypeID("wholeWord");
const idignoreAccents = stringIDToTypeID("ignoreAccents");
const idproperty = stringIDToTypeID("property");
const idusing = stringIDToTypeID("using");
const idfindReplace = stringIDToTypeID("findReplace");
const idparagraphStyle = stringIDToTypeID("paragraphStyle");

const idselectAllLayers = stringIDToTypeID("selectAllLayers");
const idselectNoLayers = stringIDToTypeID("selectNoLayers");

const idtextOverrideFeatureName = stringIDToTypeID("textOverrideFeatureName");
const idtypeStyleOperationType = stringIDToTypeID("typeStyleOperationType");
const idbaselineDirection = stringIDToTypeID("baselineDirection");
const idcross = stringIDToTypeID("cross");
const idwithStream = stringIDToTypeID("withStream");

const idtracking = stringIDToTypeID("tracking");

const idantiAlias = stringIDToTypeID("antiAlias");
const idantiAliasType = stringIDToTypeID("antiAliasType");
const idantiAliasStrong = stringIDToTypeID("antiAliasStrong");
const idantiAliasSmooth = stringIDToTypeID("antiAliasSmooth");

const idautoKern = stringIDToTypeID("autoKern");
const idopticalKern = stringIDToTypeID("opticalKern");

const idfontName = stringIDToTypeID("fontName");
const idfontStyleName = stringIDToTypeID("fontStyleName");

const eventWait = charIDToTypeID("Wait")
const enumRedrawComplete = charIDToTypeID("RdCm")
const typeState = charIDToTypeID("Stte")
const keyState = charIDToTypeID("Stte")


// 处理文本图层
function processTextLayer(textLayer) {
    var textItem = textLayer.textItem;
    var content = textItem.contents;
    if (content.indexOf('!') !== -1 || content.indexOf('！') !== -1) {
        // alert(content);
        // 将 textLayer 设为当前活动图层
        app.activeDocument.activeLayer = textLayer;
        // 获取原始格式设置
        var originalAntiAliasMethod = textItem.antiAliasMethod;
        var hasKerning = false;
        try {
            var originalAutoKerning = textItem.autoKerning;
            hasKerning = true;
        } catch (e) {
            hasKerning = false;
        }
        var originalFontSize = textItem.size;      // 原始字号
        var originalFont = textItem.font;          // 原始字体

        var hasColor = false;
        try {
            // 原始颜色
            var originalColor = textItem.color;
            hasColor = true;
        } catch (e) {
            hasColor = false;
        }

        for (var i = 0; i < content.length; i++) {
            var chara = content[i];
            if ((chara === '!' || chara === '！') &&
                ((i > 0 && (content[i - 1] !== '?' && content[i - 1] !== '？')) || i === 0) &&
                ((i < content.length - 1 && (content[i + 1] !== '?' && content[i + 1] !== '？')) || i === content.length - 1)) {
                // 处理符合条件的感叹号
                var start = i;
                var end = i + 1;

                //动作参考指向活动的文本图层
                // The action reference specifies the active text layer.
                var reference = new ActionReference();
                reference.putEnumerated(idtextLayer, idordinal, idtargetEnum);

                //第一层是action
                var action = new ActionDescriptor();
                action.putReference(idnull, reference);

                //第二层是textAction
                var textAction = new ActionDescriptor();

                //第三层是actionList
                //actionList包含了一连串格式化动作actionList contains the sequence of formatting actions.
                var actionList = new ActionList();

                //第四层是textRange
                //选定文本段
                // textRange sets the range of characters to format.
                var textRange = new ActionDescriptor();
                textRange.putInteger(idfrom, start);
                textRange.putInteger(idto, end);

                //第五层是formatting
                //The "formatting" ActionDescriptor holds the formatting. It should be clear that you can
                //add other attributes here--just get the relevant lines (usually 2) from the Script Listener
                //output and graft them into this section.
                var formatting = new ActionDescriptor();

                formatting.putBoolean(idsyntheticItalic, true);

                // 保持原始字体
                formatting.putString(idfontPostScriptName, originalFont);
                // alert(originalFont);
                // 保持原始字号
                formatting.putUnitDouble(idsize, idpixelsUnit, originalFontSize);
                if (hasColor) {
                    // 保持原始文字颜色
                    var red = originalColor.rgb.red
                    var green = originalColor.rgb.green
                    var blue = originalColor.rgb.blue
                    var colorAction = new ActionDescriptor();
                    colorAction.putDouble(idRd, red);
                    colorAction.putDouble(idGrn, green);
                    colorAction.putDouble(idBl, blue);
                    formatting.putObject(idcolor, idRGBColor, colorAction);
                }

                textRange.putObject(idtextStyle, idtextStyle, formatting);
                actionList.putObject(idtextStyleRange, textRange);
                textAction.putList(idtextStyleRange, actionList);
                action.putObject(idto, idtextLayer, textAction);
                executeAction(idset, action, DialogModes.NO);

                // 标准垂直罗马对齐方式
                var actionRoman = new ActionDescriptor();
                var refRoman = new ActionReference();
                var descRoman = new ActionDescriptor();
                refRoman.putProperty(idproperty, idtextStyle);
                refRoman.putEnumerated(idtextLayer, idordinal, idtargetEnum); // 设置动作引用指向当前目标文本图层
                descRoman.putInteger(idtextOverrideFeatureName, 808465976);
                descRoman.putInteger(idtypeStyleOperationType, 3);
                descRoman.putEnumerated(idbaselineDirection, idbaselineDirection, idwithStream);
                actionRoman.putReference(idnull, refRoman);
                actionRoman.putObject(idto, idtextStyle, descRoman);
                executeAction(idset, actionRoman, DialogModes.NO);
            }
        }
        // 恢复原始抗锯齿方法和自动字距设置
        textLayer.textItem.antiAliasMethod = originalAntiAliasMethod;
        if (hasKerning) {
            textLayer.textItem.autoKerning = originalAutoKerning;
        }
    }
}


// 遍历所有图层
function traverseLayers(layers) {
    for (var i = 0; i < layers.length; i++) {
        var layer = layers[i];
        if (layer.typename === 'ArtLayer' && layer.kind === LayerKind.TEXT) {
            // 如果是文本图层，处理文本内容
            processTextLayer(layer);
        } else if (layer.typename === 'LayerSet') {
            // 如果是图层组，递归遍历子图层
            traverseLayers(layer.layers);
        }
    }
}


//================判断系统================
var systemOS, mainInputPath;

//在Windows下，路径也必须都用“/”
if ($.os.search(/windows/i) != -1) {
    systemOS = "windows";
} else {
    systemOS = "macintosh";
}

// 获取当前打开的文档
var doc = app.activeDocument;

// 遍历文档的所有图层
traverseLayers(doc.layers);

// alert('特定符号已被修改为斜体！');
