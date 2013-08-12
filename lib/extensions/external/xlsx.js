//----------------------------------------------------------
// Copyright (C) Microsoft Corporation. All rights reserved.
// Released under the Microsoft Office Extensible File License
// https://raw.github.com/stephen-hardy/xlsx.js/master/LICENSE.txt
//----------------------------------------------------------
function xlsx(file) {
    var result, zip = new JSZip(), zipTime, processTime, s, i, j, k, l, t, w, sharedStrings, styles, index, data, val, style,
    docProps, xl, xlWorksheets, worksheet, contentTypes = [[], []], props = [], xlRels = [], worksheets = [], id, cell, row,
    numFmts = ['General', '0', '0.00', '#,##0', '#,##0.00',,,,, '0%', '0.00%', '0.00E+00', '# ?/?', '# ??/??', 'mm-dd-yy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'h:mm AM/PM', 'h:mm:ss AM/PM',
    'h:mm', 'h:mm:ss', 'm/d/yy h:mm',,,,,,,,,,,,,,, '#,##0 ;(#,##0)', '#,##0 ;[Red](#,##0)', '#,##0.00;(#,##0.00)', '#,##0.00;[Red](#,##0.00)',,,,, 'mm:ss', '[h]:mm:ss', 'mmss.0', '##0.0E+0', '@'],
    alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', worksheetXML, stylesheetXML, column, styleIndex;
    function numAlpha(i) {
        var t = Math.floor(i / 26) - 1;
        return (t > -1 ? numAlpha(t) : '') + alphabet.charAt(i % 26);
    }
    function alphaNum(s) {
        var t = 0;
        if (s.length === 2) {
            t = alphaNum(s.charAt(0)) + 1;
        }
        return t * 26 + alphabet.indexOf(s.substr(-1));
    }
    function convertDate(input) {
        return typeof input === 'object' ? ((input - new Date(1900, 0, 0)) / 86400000) + 1 : new Date(+new Date(1900, 0, 0) + (input - 1) * 86400000);
    }
    function typeOf(obj) {
        return ({}).toString.call(obj).match(/\s([a-zA-Z]+)/)[1].toLowerCase();
    }
    function getAttr(s, n) {
        s = s.substr(s.indexOf(n + '="') + n.length + 2);
        return s.substring(0, s.indexOf('"'));
    }
    function escapeXML(s) {
        return (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#x27;');
    } // see http://www.w3.org/TR/xml/#syntax
    function unescapeXML(s) {
        return (s || '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#x27;/g, '\'');
    }
    function createNumFmt(srcNumFmt){
        var numFmt = srcNumFmt || {};
        var o = {
            numFmt:{
                '@numFmtId' : numFmt.numFmtId || 0,
                '@formatCode' : numFmt.formatCode
            }
        };
        return o;
    }
    function createFont(srcFont) {
        var font = srcFont || {};
        var o = {
            sz:{
                '@val': font.size || 11
            },
            name:{
                '@val': font.name || 'Calibri'
            },
            family:{
                '@val': 2
            },
            scheme:{
                '@val': 'minor'
            }
        };
        if( font.color ) o.color = {'@rgb': rgb2string(font.color)};
        else o.color = {'@theme': 1};
        if( font.bold===true ) o.b = null;
        if( font.italic===true ) o.i = null;
        if( font.underline===true ) o.u = null;
        return o;
    }
    function createBorderComponent(srcBorderComponent){
        var borderComponent = srcBorderComponent || {};
        var o = {
            "@style": borderComponent.style || 'thin',
            color: {
                '@rgb': rgb2string(borderComponent.color)
            }
        };
        return o;
    }
    function createBorder(srcBorder) {
        var border = srcBorder || {};
        
        var o = {
            left: border.left ? createBorderComponent(border.left) : null,
            right: border.right ? createBorderComponent(border.right) : null,
            top: border.top ? createBorderComponent(border.top) : null,
            bottom: border.bottom ? createBorderComponent(border.bottom) : null,
            diagonal: border.diagonal ? createBorderComponent(border.diagonal) : null
        };
        return o;
    }
    function createFill(srcFill) {
        var fill = srcFill || {};
        var o = {
            patternFill:{
                '@patternType': fill.type || 'none'
            }
        };
        if( fill.fgColor ){
            o.patternFill.fgColor = {'@rgb': rgb2string(fill.fgColor, 0xFFFFFF)};
        }
        if( fill.bgColor ){
            o.patternFill.bgColor = {'@rgb': rgb2string(fill.bgColor, 0xFFFFFF)};
        }
        return o;
    }
    function createStyle(srcStyle) {
        var style = srcStyle || {};
        var o = {
            '@numFmtId' : style.numFmtId || 0,
            '@fontId' : style.fontId || 0,
            '@fillId' : style.fillId || 0,
            '@borderId' : style.borderId || 0,
            '@xfId' : 0
        };
        if( style.numFmtId ) o['@applyNumberFormat'] = 1;
        if( style.fontId ) o['@applyFont'] = 1;
        if( style.fillId ) o['@applyFill'] = 1;
        if( style.borderId ) o['@applyBorder'] = 1;
        return o;
    }
    function rgb2int( obj ){
        return (obj.r || 0)<<16 + (obj.g || 0)<<8 + (obj.b || 0);
    }
    function rgb2string( obj, defaultColor ){
        var retVal = 0xFF000000;
        switch( typeof(obj) ){
            case 'number':
                retVal += obj;
                break;

            case 'object':
                retVal += rgb2int(obj);
                break;
                
            default:
                retVal += (defaultColor&0xFFFFFF)||0;
        }
        return retVal.toString(16);
    }
    function indexOf( obj, array ) {
        var json_obj = JSON.stringify(obj);
        var json_elem;
        var index = array.length;
        while( index-- ){
            json_elem = JSON.stringify(array[index]);
            if( json_elem === json_obj ) break;
        }
        return index;
    }
    function addOnce( obj, array ) {
        var index;
        index = indexOf( obj, array );
        if( index<0 ){
            index = array.push( obj )-1;
        }
        return index;
    }
    if (typeof file === 'object') {
        processTime = Date.now();
        sharedStrings = [[], 0];
        //{ Fully static
        zip.folder('_rels').file('.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
        docProps = zip.folder('docProps');

        xl = zip.folder('xl');
        xl.folder('theme').file('theme1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>');
        xlWorksheets = xl.folder('worksheets');
        //}
        //{ Not content dependent
        docProps.file('core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>'
            + (file.creator || 'XLSX.js') + '</dc:creator><cp:lastModifiedBy>' + (file.lastModifiedBy || 'XLSX.js') + '</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">'
            + (file.created || new Date()).toISOString() + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + (file.modified || new Date()).toISOString() + '</dcterms:modified></cp:coreProperties>');
        //}
        //{ Content dependent
        
        stylesheetXML = {
            styleSheet:{
                "@xmlns":"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                "@xmlns:mc":"http://schemas.openxmlformats.org/markup-compatibility/2006",
                "@xmlns:x14ac":"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
                "@mc:Ignorable":"x14ac",
                numFmts:{
                    "@count": 0,
                    numFmt:[]
                },
                fonts:{
                    "@count": 0,
                    //"@x14ac:knownFonts":"1",
                    font:[ 
                    createFont() 
                    ]
                },
                fills:{
                    "@count": 0,
                    fill:[ 
                    createFill(), 
                    createFill({
                        type:'gray125'
                    }) 
                    ]
                },
                borders:{
                    "@count": 0,
                    border:[ 
                    createBorder() 
                    ]
                },
                cellStyleXfs:{
                    "@count":"1",
                    xf:{
                        "@numFmtId":"0",
                        "@fontId":"0",
                        "@fillId":"0",
                        "@borderId":"0"
                    }
                },
                cellXfs:{
                    "@count": 0,
                    xf:[
                    createStyle()
                    ]
                },
                cellStyles:{
                    "@count":"1",
                    cellStyle:{
                        "@name":"Normal",
                        "@xfId":"0",
                        "@builtinId":"0"
                    }
                },
                dxfs:{
                    "@count":"0"
                },
                tableStyles:{
                    "@count":"0",
                    "@defaultTableStyle":"TableStyleMedium2",
                    "@defaultPivotStyle":"PivotStyleLight16"
                },
                extLst:{
                    ext:{
                        "@xmlns:x14":"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                        "@uri":"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}",
                        "x14:slicerStyles":{
                            "@defaultSlicerStyle":"SlicerStyleLight1"
                        }
                    }
                }
            }
        };        
        
        w = file.worksheets.length;
        while (w--) { // Generate worksheet (gather sharedStrings), and possibly table files, then generate entries for constant files below
            id = w + 1;
            //{ Generate sheetX.xml in var s
            worksheet = file.worksheets[w];
            data = worksheet.data;
            s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
            worksheetXML = {
                worksheet:{
                    "@xmlns":"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                    "@xmlns:r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    "@xmlns:mc":"http://schemas.openxmlformats.org/markup-compatibility/2006",
                    "@xmlns:x14ac":"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
                    "@mc:Ignorable":"x14ac",
                    dimension:{
                        "@ref":"A1:" + numAlpha(data[0].length - 1) + data.length
                    },
                    sheetViews:{
                        sheetView:{
                            "@workbookViewId" : "0"
                        }
                    },
                    sheetFormatPr:{
                        "@defaultRowHeight":"15",
                        "@x14ac:dyDescent":"0.25"
                    },
                    sheetData:{
                        row:[]
                    },
                    pageMargins:{
                        "@left":"0.7",
                        "@right":"0.7",
                        "@top":"0.75",
                        "@bottom":"0.75",
                        "@header":"0.3",
                        "@footer":"0.3"
                    }
                }
            };
            if( w=== file.activeWorksheet ){
                worksheetXML.worksheet.sheetViews.sheetView["@tabSelected"] = "1";
            }
            
            styles = [];
            i = -1;
            l = data.length;
            while (++i < l) {
                j = -1;
                k = data[i].length;
                worksheetXML.worksheet.sheetData.row.push({
                    "@r":(i+1),
                    "@x14ac:dyDescent":"0.25",
                    c:[]
                });
                row = worksheetXML.worksheet.sheetData.row[i];
                
                while (++j < k) {
                    cell = data[i][j] || '';
                    val = cell.value ? cell.value : cell;
                    t = '';
                    style = null;
                    if(cell.formatCode || cell.font || cell.border || cell.fill ){
                        style = {
                            numFmtId: cell.formatCode && cell.formatCode !== 'General' ? (index=numFmts.indexOf(cell.formatCode)<0 ? index=164+stylesheetXML.styleSheet.numFmts.numFmt.length : index) : 0,
                            fontId: cell.font ? addOnce(createFont(cell.font), stylesheetXML.styleSheet.fonts.font) : 0,
                            borderId: cell.border ? addOnce(createBorder(cell.border), stylesheetXML.styleSheet.borders.border) : 0,
                            fillId: cell.fill ? addOnce(createFill(cell.fill), stylesheetXML.styleSheet.fills.fill) : 0
                        };
                        if( style.numFmtId>=164 ){
                            addOnce(createNumFmt({
                                numFmtId: index, 
                                formatCode: cell.formatCode
                            }), stylesheetXML.styleSheet.numFmts.numFmt);
                        }
                    }
                    if (val && typeof val === 'string' && !isFinite(val)) { // If value is string, and not string of just a number, place a sharedString reference instead of the value
                        val = escapeXML(val);
                        sharedStrings[1]++; // Increment total count, unique count derived from sharedStrings[0].length
                        index = sharedStrings[0].indexOf(val);
                        if (index < 0) {
                            index = sharedStrings[0].push(val) - 1;
                        }
                        val = index;
                        t = 's';
                    }
                    else if (typeof val === 'boolean') {
                        val = (val ? 1 : 0);
                        t = 'b';
                    }
                    else if (typeOf(val) === 'date') {
                        val = convertDate(val);
                        style.numFmtId = style.numFmtId || addOnce('mm-dd-yy', stylesheetXML.styleSheet.numFormats.numFormat);
                    }
                    if( style ){
                        styleIndex = addOnce( createStyle(style), stylesheetXML.styleSheet.cellXfs.xf );
                    }
                    
                    row.c.push({
                        "@r":numAlpha(j) + (i + 1),
                        "v":val
                    });
                    column = row.c[j];
                    if( t ) column["@t"] = t;
                    if( style ) column["@s"] = styleIndex;
                    if( cell.formula ) column.f = cell.formula;
                }
            }

            if (worksheet.table) {
                worksheetXML.worksheet.tableParts = {
                    "@count":"1",
                    tablePart:{
                        "@r:id":"rId1"
                    }
                };
            }
            xlWorksheets.file('sheet' + id + '.xml', s + json2xml(worksheetXML));
            //}

            if (worksheet.table) {
                i = -1;
                l = data[0].length;
                t = numAlpha(data[0].length - 1) + data.length;
                s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="' + id
                + '" name="Table' + id + '" displayName="Table' + id + '" ref="A1:' + t + '" totalsRowShown="0"><autoFilter ref="A1:' + t + '"/><tableColumns count="' + data[0].length + '">';
                while (++i < l) {
                    s += '<tableColumn id="' + (i + 1) + '" name="' + data[0][i] + '"/>';
                }
                s += '</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';

                xl.folder('tables').file('table' + id + '.xml', s);
                xlWorksheets.folder('_rels').file('sheet' + id + '.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table' + id + '.xml"/></Relationships>');
                contentTypes[1].unshift('<Override PartName="/xl/tables/table' + id + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>');
            }

            contentTypes[0].unshift('<Override PartName="/xl/worksheets/sheet' + id + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
            props.unshift(escapeXML(worksheet.name) || 'Sheet' + id);
            xlRels.unshift('<Relationship Id="rId' + id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + id + '.xml"/>');
            worksheets.unshift('<sheet name="' + (escapeXML(worksheet.name) || 'Sheet' + id) + '" sheetId="' + id + '" r:id="rId' + id + '"/>');
        }

        //{ xl/styles.xml
        // Updates element count
        stylesheetXML.styleSheet.numFmts['@count'] = stylesheetXML.styleSheet.numFmts.numFmt.length;
        stylesheetXML.styleSheet.fonts['@count'] = stylesheetXML.styleSheet.fonts.font.length;
        stylesheetXML.styleSheet.borders['@count'] = stylesheetXML.styleSheet.borders.border.length;
        stylesheetXML.styleSheet.fills['@count'] = stylesheetXML.styleSheet.fills.fill.length;
        stylesheetXML.styleSheet.cellXfs['@count'] = stylesheetXML.styleSheet.cellXfs.xf.length;
        
        s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        xl.file('styles.xml', s + json2xml(stylesheetXML));
        //}
        //{ [Content_Types].xml
        zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            + contentTypes[0].join('') + '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            + contentTypes[1].join('') + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');
        //}
        //{ docProps/app.xml
        docProps.file('app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>XLSX.js</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>'
            + file.worksheets.length + '</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="' + props.length + '" baseType="lpstr"><vt:lpstr>' + props.join('</vt:lpstr><vt:lpstr>')
            + '</vt:lpstr></vt:vector></TitlesOfParts><Manager></Manager><Company>Microsoft Corporation</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>1.0</AppVersion></Properties>');
        //}
        //{ xl/_rels/workbook.xml.rels
        xl.folder('_rels').file('workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + xlRels.join('') + '<Relationship Id="rId' + (xlRels.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
            + '<Relationship Id="rId' + (xlRels.length + 2) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            + '<Relationship Id="rId' + (xlRels.length + 3) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>');
        //}
        //{ xl/sharedStrings.xml
        xl.file('sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'
            + sharedStrings[1] + '" uniqueCount="' + sharedStrings[0].length + '"><si><t>' + sharedStrings[0].join('</t></si><si><t>') + '</t></si></sst>');
        //}
        //{ xl/workbook.xml
        xl.file('workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            + '<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9303"/><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView '
            + (file.activeWorksheet ? 'activeTab="' + file.activeWorksheet + '" ' : '') + 'xWindow="480" yWindow="60" windowWidth="18195" windowHeight="8505"/></bookViews><sheets>'
            + worksheets.join('') + '</sheets><calcPr calcId="145621"/></workbook>');
        //}
        //}
        processTime = Date.now() - processTime;
        zipTime = Date.now();
        result = {
            base64: zip.generate({
                compression: 'DEFLATE'
            }), 
            zipTime: Date.now() - zipTime, 
            processTime: processTime,
            href: function() {
                return 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + this.base64;
            }
        };
    }
    return result;
}
