/***************************************************************************

Assembled from parts of Power Tools for Open XML (license and author below),
with the intention of producing a minimal binary, with the ability to be
compiled and run in Mono (on Linux).

/***************************************************************************


Copyright (c) Microsoft Corporation 2012-2013.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************/


using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace RevisionAccepterProgram
{
    public static class M
    {
        public static XNamespace m =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        public static XName acc = m + "acc";
        public static XName accPr = m + "accPr";
        public static XName aln = m + "aln";
        public static XName alnAt = m + "alnAt";
        public static XName alnScr = m + "alnScr";
        public static XName argPr = m + "argPr";
        public static XName argSz = m + "argSz";
        public static XName bar = m + "bar";
        public static XName barPr = m + "barPr";
        public static XName baseJc = m + "baseJc";
        public static XName begChr = m + "begChr";
        public static XName borderBox = m + "borderBox";
        public static XName borderBoxPr = m + "borderBoxPr";
        public static XName box = m + "box";
        public static XName boxPr = m + "boxPr";
        public static XName brk = m + "brk";
        public static XName brkBin = m + "brkBin";
        public static XName brkBinSub = m + "brkBinSub";
        public static XName cGp = m + "cGp";
        public static XName cGpRule = m + "cGpRule";
        public static XName chr = m + "chr";
        public static XName count = m + "count";
        public static XName cSp = m + "cSp";
        public static XName ctrlPr = m + "ctrlPr";
        public static XName d = m + "d";
        public static XName defJc = m + "defJc";
        public static XName deg = m + "deg";
        public static XName degHide = m + "degHide";
        public static XName den = m + "den";
        public static XName diff = m + "diff";
        public static XName dispDef = m + "dispDef";
        public static XName dPr = m + "dPr";
        public static XName e = m + "e";
        public static XName endChr = m + "endChr";
        public static XName eqArr = m + "eqArr";
        public static XName eqArrPr = m + "eqArrPr";
        public static XName f = m + "f";
        public static XName fName = m + "fName";
        public static XName fPr = m + "fPr";
        public static XName func = m + "func";
        public static XName funcPr = m + "funcPr";
        public static XName groupChr = m + "groupChr";
        public static XName groupChrPr = m + "groupChrPr";
        public static XName grow = m + "grow";
        public static XName hideBot = m + "hideBot";
        public static XName hideLeft = m + "hideLeft";
        public static XName hideRight = m + "hideRight";
        public static XName hideTop = m + "hideTop";
        public static XName interSp = m + "interSp";
        public static XName intLim = m + "intLim";
        public static XName intraSp = m + "intraSp";
        public static XName jc = m + "jc";
        public static XName lim = m + "lim";
        public static XName limLoc = m + "limLoc";
        public static XName limLow = m + "limLow";
        public static XName limLowPr = m + "limLowPr";
        public static XName limUpp = m + "limUpp";
        public static XName limUppPr = m + "limUppPr";
        public static XName lit = m + "lit";
        public static XName lMargin = m + "lMargin";
        public static XName _m = m + "m";
        public static XName mathFont = m + "mathFont";
        public static XName mathPr = m + "mathPr";
        public static XName maxDist = m + "maxDist";
        public static XName mc = m + "mc";
        public static XName mcJc = m + "mcJc";
        public static XName mcPr = m + "mcPr";
        public static XName mcs = m + "mcs";
        public static XName mPr = m + "mPr";
        public static XName mr = m + "mr";
        public static XName nary = m + "nary";
        public static XName naryLim = m + "naryLim";
        public static XName naryPr = m + "naryPr";
        public static XName noBreak = m + "noBreak";
        public static XName nor = m + "nor";
        public static XName num = m + "num";
        public static XName objDist = m + "objDist";
        public static XName oMath = m + "oMath";
        public static XName oMathPara = m + "oMathPara";
        public static XName oMathParaPr = m + "oMathParaPr";
        public static XName opEmu = m + "opEmu";
        public static XName phant = m + "phant";
        public static XName phantPr = m + "phantPr";
        public static XName plcHide = m + "plcHide";
        public static XName pos = m + "pos";
        public static XName postSp = m + "postSp";
        public static XName preSp = m + "preSp";
        public static XName r = m + "r";
        public static XName rad = m + "rad";
        public static XName radPr = m + "radPr";
        public static XName rMargin = m + "rMargin";
        public static XName rPr = m + "rPr";
        public static XName rSp = m + "rSp";
        public static XName rSpRule = m + "rSpRule";
        public static XName scr = m + "scr";
        public static XName sepChr = m + "sepChr";
        public static XName show = m + "show";
        public static XName shp = m + "shp";
        public static XName smallFrac = m + "smallFrac";
        public static XName sPre = m + "sPre";
        public static XName sPrePr = m + "sPrePr";
        public static XName sSub = m + "sSub";
        public static XName sSubPr = m + "sSubPr";
        public static XName sSubSup = m + "sSubSup";
        public static XName sSubSupPr = m + "sSubSupPr";
        public static XName sSup = m + "sSup";
        public static XName sSupPr = m + "sSupPr";
        public static XName strikeBLTR = m + "strikeBLTR";
        public static XName strikeH = m + "strikeH";
        public static XName strikeTLBR = m + "strikeTLBR";
        public static XName strikeV = m + "strikeV";
        public static XName sty = m + "sty";
        public static XName sub = m + "sub";
        public static XName subHide = m + "subHide";
        public static XName sup = m + "sup";
        public static XName supHide = m + "supHide";
        public static XName t = m + "t";
        public static XName transp = m + "transp";
        public static XName type = m + "type";
        public static XName val = m + "val";
        public static XName vertJc = m + "vertJc";
        public static XName wrapIndent = m + "wrapIndent";
        public static XName wrapRight = m + "wrapRight";
        public static XName zeroAsc = m + "zeroAsc";
        public static XName zeroDesc = m + "zeroDesc";
        public static XName zeroWid = m + "zeroWid";
    }

    public static class W
    {
        public static XNamespace w =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static XName abstractNum = w + "abstractNum";
        public static XName abstractNumId = w + "abstractNumId";
        public static XName accent1 = w + "accent1";
        public static XName accent2 = w + "accent2";
        public static XName accent3 = w + "accent3";
        public static XName accent4 = w + "accent4";
        public static XName accent5 = w + "accent5";
        public static XName accent6 = w + "accent6";
        public static XName activeRecord = w + "activeRecord";
        public static XName activeWritingStyle = w + "activeWritingStyle";
        public static XName actualPg = w + "actualPg";
        public static XName addressFieldName = w + "addressFieldName";
        public static XName adjustLineHeightInTable = w + "adjustLineHeightInTable";
        public static XName adjustRightInd = w + "adjustRightInd";
        public static XName after = w + "after";
        public static XName afterAutospacing = w + "afterAutospacing";
        public static XName afterLines = w + "afterLines";
        public static XName algIdExt = w + "algIdExt";
        public static XName algIdExtSource = w + "algIdExtSource";
        public static XName alias = w + "alias";
        public static XName aliases = w + "aliases";
        public static XName alignBordersAndEdges = w + "alignBordersAndEdges";
        public static XName alignment = w + "alignment";
        public static XName alignTablesRowByRow = w + "alignTablesRowByRow";
        public static XName allowPNG = w + "allowPNG";
        public static XName allowSpaceOfSameStyleInTable = w + "allowSpaceOfSameStyleInTable";
        public static XName altChunk = w + "altChunk";
        public static XName altChunkPr = w + "altChunkPr";
        public static XName altName = w + "altName";
        public static XName alwaysMergeEmptyNamespace = w + "alwaysMergeEmptyNamespace";
        public static XName alwaysShowPlaceholderText = w + "alwaysShowPlaceholderText";
        public static XName anchor = w + "anchor";
        public static XName anchorLock = w + "anchorLock";
        public static XName annotationRef = w + "annotationRef";
        public static XName applyBreakingRules = w + "applyBreakingRules";
        public static XName appName = w + "appName";
        public static XName ascii = w + "ascii";
        public static XName asciiTheme = w + "asciiTheme";
        public static XName attachedSchema = w + "attachedSchema";
        public static XName attachedTemplate = w + "attachedTemplate";
        public static XName attr = w + "attr";
        public static XName author = w + "author";
        public static XName autofitToFirstFixedWidthCell = w + "autofitToFirstFixedWidthCell";
        public static XName autoFormatOverride = w + "autoFormatOverride";
        public static XName autoHyphenation = w + "autoHyphenation";
        public static XName autoRedefine = w + "autoRedefine";
        public static XName autoSpaceDE = w + "autoSpaceDE";
        public static XName autoSpaceDN = w + "autoSpaceDN";
        public static XName autoSpaceLikeWord95 = w + "autoSpaceLikeWord95";
        public static XName b = w + "b";
        public static XName background = w + "background";
        public static XName balanceSingleByteDoubleByteWidth = w + "balanceSingleByteDoubleByteWidth";
        public static XName bar = w + "bar";
        public static XName basedOn = w + "basedOn";
        public static XName bCs = w + "bCs";
        public static XName bdr = w + "bdr";
        public static XName before = w + "before";
        public static XName beforeAutospacing = w + "beforeAutospacing";
        public static XName beforeLines = w + "beforeLines";
        public static XName behavior = w + "behavior";
        public static XName behaviors = w + "behaviors";
        public static XName between = w + "between";
        public static XName bg1 = w + "bg1";
        public static XName bg2 = w + "bg2";
        public static XName bibliography = w + "bibliography";
        public static XName bidi = w + "bidi";
        public static XName bidiVisual = w + "bidiVisual";
        public static XName blockQuote = w + "blockQuote";
        public static XName body = w + "body";
        public static XName bodyDiv = w + "bodyDiv";
        public static XName bookFoldPrinting = w + "bookFoldPrinting";
        public static XName bookFoldPrintingSheets = w + "bookFoldPrintingSheets";
        public static XName bookFoldRevPrinting = w + "bookFoldRevPrinting";
        public static XName bookmarkEnd = w + "bookmarkEnd";
        public static XName bookmarkStart = w + "bookmarkStart";
        public static XName bordersDoNotSurroundFooter = w + "bordersDoNotSurroundFooter";
        public static XName bordersDoNotSurroundHeader = w + "bordersDoNotSurroundHeader";
        public static XName bottom = w + "bottom";
        public static XName bottomFromText = w + "bottomFromText";
        public static XName br = w + "br";
        public static XName cachedColBalance = w + "cachedColBalance";
        public static XName calcOnExit = w + "calcOnExit";
        public static XName calendar = w + "calendar";
        public static XName cantSplit = w + "cantSplit";
        public static XName caps = w + "caps";
        public static XName category = w + "category";
        public static XName cellDel = w + "cellDel";
        public static XName cellIns = w + "cellIns";
        public static XName cellMerge = w + "cellMerge";
        public static XName chapSep = w + "chapSep";
        public static XName chapStyle = w + "chapStyle";
        public static XName _char = w + "char";
        public static XName characterSpacingControl = w + "characterSpacingControl";
        public static XName charset = w + "charset";
        public static XName charSpace = w + "charSpace";
        public static XName checkBox = w + "checkBox";
        public static XName _checked = w + "checked";
        public static XName checkErrors = w + "checkErrors";
        public static XName checkStyle = w + "checkStyle";
        public static XName citation = w + "citation";
        public static XName clear = w + "clear";
        public static XName clickAndTypeStyle = w + "clickAndTypeStyle";
        public static XName clrSchemeMapping = w + "clrSchemeMapping";
        public static XName cnfStyle = w + "cnfStyle";
        public static XName code = w + "code";
        public static XName col = w + "col";
        public static XName colDelim = w + "colDelim";
        public static XName colFirst = w + "colFirst";
        public static XName colLast = w + "colLast";
        public static XName color = w + "color";
        public static XName cols = w + "cols";
        public static XName column = w + "column";
        public static XName combine = w + "combine";
        public static XName combineBrackets = w + "combineBrackets";
        public static XName comboBox = w + "comboBox";
        public static XName comment = w + "comment";
        public static XName commentRangeEnd = w + "commentRangeEnd";
        public static XName commentRangeStart = w + "commentRangeStart";
        public static XName commentReference = w + "commentReference";
        public static XName comments = w + "comments";
        public static XName compat = w + "compat";
        public static XName compatSetting = w + "compatSetting";
        public static XName connectString = w + "connectString";
        public static XName consecutiveHyphenLimit = w + "consecutiveHyphenLimit";
        public static XName contextualSpacing = w + "contextualSpacing";
        public static XName continuationSeparator = w + "continuationSeparator";
        public static XName control = w + "control";
        public static XName convMailMergeEsc = w + "convMailMergeEsc";
        public static XName count = w + "count";
        public static XName countBy = w + "countBy";
        public static XName cr = w + "cr";
        public static XName cryptAlgorithmClass = w + "cryptAlgorithmClass";
        public static XName cryptAlgorithmSid = w + "cryptAlgorithmSid";
        public static XName cryptAlgorithmType = w + "cryptAlgorithmType";
        public static XName cryptProvider = w + "cryptProvider";
        public static XName cryptProviderType = w + "cryptProviderType";
        public static XName cryptProviderTypeExt = w + "cryptProviderTypeExt";
        public static XName cryptProviderTypeExtSource = w + "cryptProviderTypeExtSource";
        public static XName cryptSpinCount = w + "cryptSpinCount";
        public static XName cs = w + "cs";
        public static XName csb0 = w + "csb0";
        public static XName csb1 = w + "csb1";
        public static XName cstheme = w + "cstheme";
        public static XName customMarkFollows = w + "customMarkFollows";
        public static XName customStyle = w + "customStyle";
        public static XName customXml = w + "customXml";
        public static XName customXmlDelRangeEnd = w + "customXmlDelRangeEnd";
        public static XName customXmlDelRangeStart = w + "customXmlDelRangeStart";
        public static XName customXmlInsRangeEnd = w + "customXmlInsRangeEnd";
        public static XName customXmlInsRangeStart = w + "customXmlInsRangeStart";
        public static XName customXmlMoveFromRangeEnd = w + "customXmlMoveFromRangeEnd";
        public static XName customXmlMoveFromRangeStart = w + "customXmlMoveFromRangeStart";
        public static XName customXmlMoveToRangeEnd = w + "customXmlMoveToRangeEnd";
        public static XName customXmlMoveToRangeStart = w + "customXmlMoveToRangeStart";
        public static XName customXmlPr = w + "customXmlPr";
        public static XName dataBinding = w + "dataBinding";
        public static XName dataSource = w + "dataSource";
        public static XName dataType = w + "dataType";
        public static XName date = w + "date";
        public static XName dateFormat = w + "dateFormat";
        public static XName dayLong = w + "dayLong";
        public static XName dayShort = w + "dayShort";
        public static XName ddList = w + "ddList";
        public static XName decimalSymbol = w + "decimalSymbol";
        public static XName _default = w + "default";
        public static XName defaultTableStyle = w + "defaultTableStyle";
        public static XName defaultTabStop = w + "defaultTabStop";
        public static XName defLockedState = w + "defLockedState";
        public static XName defQFormat = w + "defQFormat";
        public static XName defSemiHidden = w + "defSemiHidden";
        public static XName defUIPriority = w + "defUIPriority";
        public static XName defUnhideWhenUsed = w + "defUnhideWhenUsed";
        public static XName del = w + "del";
        public static XName delInstrText = w + "delInstrText";
        public static XName delText = w + "delText";
        public static XName description = w + "description";
        public static XName destination = w + "destination";
        public static XName dir = w + "dir";
        public static XName dirty = w + "dirty";
        public static XName displacedByCustomXml = w + "displacedByCustomXml";
        public static XName display = w + "display";
        public static XName displayBackgroundShape = w + "displayBackgroundShape";
        public static XName displayHangulFixedWidth = w + "displayHangulFixedWidth";
        public static XName displayHorizontalDrawingGridEvery = w + "displayHorizontalDrawingGridEvery";
        public static XName displayText = w + "displayText";
        public static XName displayVerticalDrawingGridEvery = w + "displayVerticalDrawingGridEvery";
        public static XName distance = w + "distance";
        public static XName div = w + "div";
        public static XName divBdr = w + "divBdr";
        public static XName divId = w + "divId";
        public static XName divs = w + "divs";
        public static XName divsChild = w + "divsChild";
        public static XName dllVersion = w + "dllVersion";
        public static XName docDefaults = w + "docDefaults";
        public static XName docGrid = w + "docGrid";
        public static XName docLocation = w + "docLocation";
        public static XName docPart = w + "docPart";
        public static XName docPartBody = w + "docPartBody";
        public static XName docPartCategory = w + "docPartCategory";
        public static XName docPartGallery = w + "docPartGallery";
        public static XName docPartList = w + "docPartList";
        public static XName docPartObj = w + "docPartObj";
        public static XName docPartPr = w + "docPartPr";
        public static XName docParts = w + "docParts";
        public static XName docPartUnique = w + "docPartUnique";
        public static XName document = w + "document";
        public static XName documentProtection = w + "documentProtection";
        public static XName documentType = w + "documentType";
        public static XName docVar = w + "docVar";
        public static XName docVars = w + "docVars";
        public static XName doNotAutoCompressPictures = w + "doNotAutoCompressPictures";
        public static XName doNotAutofitConstrainedTables = w + "doNotAutofitConstrainedTables";
        public static XName doNotBreakConstrainedForcedTable = w + "doNotBreakConstrainedForcedTable";
        public static XName doNotBreakWrappedTables = w + "doNotBreakWrappedTables";
        public static XName doNotDemarcateInvalidXml = w + "doNotDemarcateInvalidXml";
        public static XName doNotDisplayPageBoundaries = w + "doNotDisplayPageBoundaries";
        public static XName doNotEmbedSmartTags = w + "doNotEmbedSmartTags";
        public static XName doNotExpandShiftReturn = w + "doNotExpandShiftReturn";
        public static XName doNotHyphenateCaps = w + "doNotHyphenateCaps";
        public static XName doNotIncludeSubdocsInStats = w + "doNotIncludeSubdocsInStats";
        public static XName doNotLeaveBackslashAlone = w + "doNotLeaveBackslashAlone";
        public static XName doNotOrganizeInFolder = w + "doNotOrganizeInFolder";
        public static XName doNotRelyOnCSS = w + "doNotRelyOnCSS";
        public static XName doNotSaveAsSingleFile = w + "doNotSaveAsSingleFile";
        public static XName doNotShadeFormData = w + "doNotShadeFormData";
        public static XName doNotSnapToGridInCell = w + "doNotSnapToGridInCell";
        public static XName doNotSuppressBlankLines = w + "doNotSuppressBlankLines";
        public static XName doNotSuppressIndentation = w + "doNotSuppressIndentation";
        public static XName doNotSuppressParagraphBorders = w + "doNotSuppressParagraphBorders";
        public static XName doNotTrackFormatting = w + "doNotTrackFormatting";
        public static XName doNotTrackMoves = w + "doNotTrackMoves";
        public static XName doNotUseEastAsianBreakRules = w + "doNotUseEastAsianBreakRules";
        public static XName doNotUseHTMLParagraphAutoSpacing = w + "doNotUseHTMLParagraphAutoSpacing";
        public static XName doNotUseIndentAsNumberingTabStop = w + "doNotUseIndentAsNumberingTabStop";
        public static XName doNotUseLongFileNames = w + "doNotUseLongFileNames";
        public static XName doNotUseMarginsForDrawingGridOrigin = w + "doNotUseMarginsForDrawingGridOrigin";
        public static XName doNotValidateAgainstSchema = w + "doNotValidateAgainstSchema";
        public static XName doNotVertAlignCellWithSp = w + "doNotVertAlignCellWithSp";
        public static XName doNotVertAlignInTxbx = w + "doNotVertAlignInTxbx";
        public static XName doNotWrapTextWithPunct = w + "doNotWrapTextWithPunct";
        public static XName drawing = w + "drawing";
        public static XName drawingGridHorizontalOrigin = w + "drawingGridHorizontalOrigin";
        public static XName drawingGridHorizontalSpacing = w + "drawingGridHorizontalSpacing";
        public static XName drawingGridVerticalOrigin = w + "drawingGridVerticalOrigin";
        public static XName drawingGridVerticalSpacing = w + "drawingGridVerticalSpacing";
        public static XName dropCap = w + "dropCap";
        public static XName dropDownList = w + "dropDownList";
        public static XName dstrike = w + "dstrike";
        public static XName dxaOrig = w + "dxaOrig";
        public static XName dyaOrig = w + "dyaOrig";
        public static XName dynamicAddress = w + "dynamicAddress";
        public static XName eastAsia = w + "eastAsia";
        public static XName eastAsianLayout = w + "eastAsianLayout";
        public static XName eastAsiaTheme = w + "eastAsiaTheme";
        public static XName ed = w + "ed";
        public static XName edGrp = w + "edGrp";
        public static XName edit = w + "edit";
        public static XName effect = w + "effect";
        public static XName element = w + "element";
        public static XName em = w + "em";
        public static XName embedBold = w + "embedBold";
        public static XName embedBoldItalic = w + "embedBoldItalic";
        public static XName embedItalic = w + "embedItalic";
        public static XName embedRegular = w + "embedRegular";
        public static XName embedSystemFonts = w + "embedSystemFonts";
        public static XName embedTrueTypeFonts = w + "embedTrueTypeFonts";
        public static XName emboss = w + "emboss";
        public static XName enabled = w + "enabled";
        public static XName encoding = w + "encoding";
        public static XName endnote = w + "endnote";
        public static XName endnotePr = w + "endnotePr";
        public static XName endnoteRef = w + "endnoteRef";
        public static XName endnoteReference = w + "endnoteReference";
        public static XName endnotes = w + "endnotes";
        public static XName enforcement = w + "enforcement";
        public static XName entryMacro = w + "entryMacro";
        public static XName equalWidth = w + "equalWidth";
        public static XName equation = w + "equation";
        public static XName evenAndOddHeaders = w + "evenAndOddHeaders";
        public static XName exitMacro = w + "exitMacro";
        public static XName family = w + "family";
        public static XName ffData = w + "ffData";
        public static XName fHdr = w + "fHdr";
        public static XName fieldMapData = w + "fieldMapData";
        public static XName fill = w + "fill";
        public static XName first = w + "first";
        public static XName firstColumn = w + "firstColumn";
        public static XName firstLine = w + "firstLine";
        public static XName firstLineChars = w + "firstLineChars";
        public static XName firstRow = w + "firstRow";
        public static XName fitText = w + "fitText";
        public static XName flatBorders = w + "flatBorders";
        public static XName fldChar = w + "fldChar";
        public static XName fldCharType = w + "fldCharType";
        public static XName fldData = w + "fldData";
        public static XName fldLock = w + "fldLock";
        public static XName fldSimple = w + "fldSimple";
        public static XName fmt = w + "fmt";
        public static XName followedHyperlink = w + "followedHyperlink";
        public static XName font = w + "font";
        public static XName fontKey = w + "fontKey";
        public static XName fonts = w + "fonts";
        public static XName fontSz = w + "fontSz";
        public static XName footer = w + "footer";
        public static XName footerReference = w + "footerReference";
        public static XName footnote = w + "footnote";
        public static XName footnoteLayoutLikeWW8 = w + "footnoteLayoutLikeWW8";
        public static XName footnotePr = w + "footnotePr";
        public static XName footnoteRef = w + "footnoteRef";
        public static XName footnoteReference = w + "footnoteReference";
        public static XName footnotes = w + "footnotes";
        public static XName forceUpgrade = w + "forceUpgrade";
        public static XName forgetLastTabAlignment = w + "forgetLastTabAlignment";
        public static XName format = w + "format";
        public static XName formatting = w + "formatting";
        public static XName formProt = w + "formProt";
        public static XName formsDesign = w + "formsDesign";
        public static XName frame = w + "frame";
        public static XName frameLayout = w + "frameLayout";
        public static XName framePr = w + "framePr";
        public static XName frameset = w + "frameset";
        public static XName framesetSplitbar = w + "framesetSplitbar";
        public static XName ftr = w + "ftr";
        public static XName fullDate = w + "fullDate";
        public static XName gallery = w + "gallery";
        public static XName glossaryDocument = w + "glossaryDocument";
        public static XName grammar = w + "grammar";
        public static XName gridAfter = w + "gridAfter";
        public static XName gridBefore = w + "gridBefore";
        public static XName gridCol = w + "gridCol";
        public static XName gridSpan = w + "gridSpan";
        public static XName group = w + "group";
        public static XName growAutofit = w + "growAutofit";
        public static XName guid = w + "guid";
        public static XName gutter = w + "gutter";
        public static XName gutterAtTop = w + "gutterAtTop";
        public static XName h = w + "h";
        public static XName hAnchor = w + "hAnchor";
        public static XName hanging = w + "hanging";
        public static XName hangingChars = w + "hangingChars";
        public static XName hAnsi = w + "hAnsi";
        public static XName hAnsiTheme = w + "hAnsiTheme";
        public static XName hash = w + "hash";
        public static XName hdr = w + "hdr";
        public static XName hdrShapeDefaults = w + "hdrShapeDefaults";
        public static XName header = w + "header";
        public static XName headerReference = w + "headerReference";
        public static XName headerSource = w + "headerSource";
        public static XName helpText = w + "helpText";
        public static XName hidden = w + "hidden";
        public static XName hideGrammaticalErrors = w + "hideGrammaticalErrors";
        public static XName hideMark = w + "hideMark";
        public static XName hideSpellingErrors = w + "hideSpellingErrors";
        public static XName highlight = w + "highlight";
        public static XName hint = w + "hint";
        public static XName history = w + "history";
        public static XName hMerge = w + "hMerge";
        public static XName horzAnchor = w + "horzAnchor";
        public static XName hps = w + "hps";
        public static XName hpsBaseText = w + "hpsBaseText";
        public static XName hpsRaise = w + "hpsRaise";
        public static XName hRule = w + "hRule";
        public static XName hSpace = w + "hSpace";
        public static XName hyperlink = w + "hyperlink";
        public static XName hyphenationZone = w + "hyphenationZone";
        public static XName i = w + "i";
        public static XName iCs = w + "iCs";
        public static XName id = w + "id";
        public static XName ignoreMixedContent = w + "ignoreMixedContent";
        public static XName ilvl = w + "ilvl";
        public static XName imprint = w + "imprint";
        public static XName ind = w + "ind";
        public static XName initials = w + "initials";
        public static XName inkAnnotations = w + "inkAnnotations";
        public static XName ins = w + "ins";
        public static XName insDel = w + "insDel";
        public static XName insideH = w + "insideH";
        public static XName insideV = w + "insideV";
        public static XName instr = w + "instr";
        public static XName instrText = w + "instrText";
        public static XName isLgl = w + "isLgl";
        public static XName jc = w + "jc";
        public static XName keepLines = w + "keepLines";
        public static XName keepNext = w + "keepNext";
        public static XName kern = w + "kern";
        public static XName kinsoku = w + "kinsoku";
        public static XName lang = w + "lang";
        public static XName lastColumn = w + "lastColumn";
        public static XName lastRenderedPageBreak = w + "lastRenderedPageBreak";
        public static XName lastValue = w + "lastValue";
        public static XName lastRow = w + "lastRow";
        public static XName latentStyles = w + "latentStyles";
        public static XName layoutRawTableWidth = w + "layoutRawTableWidth";
        public static XName layoutTableRowsApart = w + "layoutTableRowsApart";
        public static XName leader = w + "leader";
        public static XName left = w + "left";
        public static XName leftChars = w + "leftChars";
        public static XName leftFromText = w + "leftFromText";
        public static XName legacy = w + "legacy";
        public static XName legacyIndent = w + "legacyIndent";
        public static XName legacySpace = w + "legacySpace";
        public static XName lid = w + "lid";
        public static XName line = w + "line";
        public static XName linePitch = w + "linePitch";
        public static XName lineRule = w + "lineRule";
        public static XName lines = w + "lines";
        public static XName lineWrapLikeWord6 = w + "lineWrapLikeWord6";
        public static XName link = w + "link";
        public static XName linkedToFile = w + "linkedToFile";
        public static XName linkStyles = w + "linkStyles";
        public static XName linkToQuery = w + "linkToQuery";
        public static XName listEntry = w + "listEntry";
        public static XName listItem = w + "listItem";
        public static XName listSeparator = w + "listSeparator";
        public static XName lnNumType = w + "lnNumType";
        public static XName _lock = w + "lock";
        public static XName locked = w + "locked";
        public static XName lsdException = w + "lsdException";
        public static XName lvl = w + "lvl";
        public static XName lvlJc = w + "lvlJc";
        public static XName lvlOverride = w + "lvlOverride";
        public static XName lvlPicBulletId = w + "lvlPicBulletId";
        public static XName lvlRestart = w + "lvlRestart";
        public static XName lvlText = w + "lvlText";
        public static XName mailAsAttachment = w + "mailAsAttachment";
        public static XName mailMerge = w + "mailMerge";
        public static XName mailSubject = w + "mailSubject";
        public static XName mainDocumentType = w + "mainDocumentType";
        public static XName mappedName = w + "mappedName";
        public static XName marBottom = w + "marBottom";
        public static XName marH = w + "marH";
        public static XName markup = w + "markup";
        public static XName marLeft = w + "marLeft";
        public static XName marRight = w + "marRight";
        public static XName marTop = w + "marTop";
        public static XName marW = w + "marW";
        public static XName matchSrc = w + "matchSrc";
        public static XName maxLength = w + "maxLength";
        public static XName mirrorIndents = w + "mirrorIndents";
        public static XName mirrorMargins = w + "mirrorMargins";
        public static XName monthLong = w + "monthLong";
        public static XName monthShort = w + "monthShort";
        public static XName moveFrom = w + "moveFrom";
        public static XName moveFromRangeEnd = w + "moveFromRangeEnd";
        public static XName moveFromRangeStart = w + "moveFromRangeStart";
        public static XName moveTo = w + "moveTo";
        public static XName moveToRangeEnd = w + "moveToRangeEnd";
        public static XName moveToRangeStart = w + "moveToRangeStart";
        public static XName multiLevelType = w + "multiLevelType";
        public static XName multiLine = w + "multiLine";
        public static XName mwSmallCaps = w + "mwSmallCaps";
        public static XName name = w + "name";
        public static XName namespaceuri = w + "namespaceuri";
        public static XName next = w + "next";
        public static XName nlCheck = w + "nlCheck";
        public static XName noBorder = w + "noBorder";
        public static XName noBreakHyphen = w + "noBreakHyphen";
        public static XName noColumnBalance = w + "noColumnBalance";
        public static XName noEndnote = w + "noEndnote";
        public static XName noExtraLineSpacing = w + "noExtraLineSpacing";
        public static XName noHBand = w + "noHBand";
        public static XName noLeading = w + "noLeading";
        public static XName noLineBreaksAfter = w + "noLineBreaksAfter";
        public static XName noLineBreaksBefore = w + "noLineBreaksBefore";
        public static XName noProof = w + "noProof";
        public static XName noPunctuationKerning = w + "noPunctuationKerning";
        public static XName noResizeAllowed = w + "noResizeAllowed";
        public static XName noSpaceRaiseLower = w + "noSpaceRaiseLower";
        public static XName noTabHangInd = w + "noTabHangInd";
        public static XName notTrueType = w + "notTrueType";
        public static XName noVBand = w + "noVBand";
        public static XName noWrap = w + "noWrap";
        public static XName nsid = w + "nsid";
        public static XName _null = w + "null";
        public static XName num = w + "num";
        public static XName numbering = w + "numbering";
        public static XName numberingChange = w + "numberingChange";
        public static XName numFmt = w + "numFmt";
        public static XName numId = w + "numId";
        public static XName numIdMacAtCleanup = w + "numIdMacAtCleanup";
        public static XName numPicBullet = w + "numPicBullet";
        public static XName numPicBulletId = w + "numPicBulletId";
        public static XName numPr = w + "numPr";
        public static XName numRestart = w + "numRestart";
        public static XName numStart = w + "numStart";
        public static XName numStyleLink = w + "numStyleLink";
        public static XName _object = w + "object";
        public static XName odso = w + "odso";
        public static XName offsetFrom = w + "offsetFrom";
        public static XName oMath = w + "oMath";
        public static XName optimizeForBrowser = w + "optimizeForBrowser";
        public static XName orient = w + "orient";
        public static XName original = w + "original";
        public static XName other = w + "other";
        public static XName outline = w + "outline";
        public static XName outlineLvl = w + "outlineLvl";
        public static XName overflowPunct = w + "overflowPunct";
        public static XName p = w + "p";
        public static XName pageBreakBefore = w + "pageBreakBefore";
        public static XName panose1 = w + "panose1";
        public static XName paperSrc = w + "paperSrc";
        public static XName pBdr = w + "pBdr";
        public static XName percent = w + "percent";
        public static XName permEnd = w + "permEnd";
        public static XName permStart = w + "permStart";
        public static XName personal = w + "personal";
        public static XName personalCompose = w + "personalCompose";
        public static XName personalReply = w + "personalReply";
        public static XName pgBorders = w + "pgBorders";
        public static XName pgMar = w + "pgMar";
        public static XName pgNum = w + "pgNum";
        public static XName pgNumType = w + "pgNumType";
        public static XName pgSz = w + "pgSz";
        public static XName pict = w + "pict";
        public static XName picture = w + "picture";
        public static XName pitch = w + "pitch";
        public static XName pixelsPerInch = w + "pixelsPerInch";
        public static XName placeholder = w + "placeholder";
        public static XName pos = w + "pos";
        public static XName position = w + "position";
        public static XName pPr = w + "pPr";
        public static XName pPrChange = w + "pPrChange";
        public static XName pPrDefault = w + "pPrDefault";
        public static XName prefixMappings = w + "prefixMappings";
        public static XName printBodyTextBeforeHeader = w + "printBodyTextBeforeHeader";
        public static XName printColBlack = w + "printColBlack";
        public static XName printerSettings = w + "printerSettings";
        public static XName printFormsData = w + "printFormsData";
        public static XName printFractionalCharacterWidth = w + "printFractionalCharacterWidth";
        public static XName printPostScriptOverText = w + "printPostScriptOverText";
        public static XName printTwoOnOne = w + "printTwoOnOne";
        public static XName proofErr = w + "proofErr";
        public static XName proofState = w + "proofState";
        public static XName pStyle = w + "pStyle";
        public static XName ptab = w + "ptab";
        public static XName qFormat = w + "qFormat";
        public static XName query = w + "query";
        public static XName r = w + "r";
        public static XName readModeInkLockDown = w + "readModeInkLockDown";
        public static XName recipientData = w + "recipientData";
        public static XName recommended = w + "recommended";
        public static XName relativeTo = w + "relativeTo";
        public static XName relyOnVML = w + "relyOnVML";
        public static XName removeDateAndTime = w + "removeDateAndTime";
        public static XName removePersonalInformation = w + "removePersonalInformation";
        public static XName restart = w + "restart";
        public static XName result = w + "result";
        public static XName revisionView = w + "revisionView";
        public static XName rFonts = w + "rFonts";
        public static XName richText = w + "richText";
        public static XName right = w + "right";
        public static XName rightChars = w + "rightChars";
        public static XName rightFromText = w + "rightFromText";
        public static XName rPr = w + "rPr";
        public static XName rPrChange = w + "rPrChange";
        public static XName rPrDefault = w + "rPrDefault";
        public static XName rsid = w + "rsid";
        public static XName rsidDel = w + "rsidDel";
        public static XName rsidP = w + "rsidP";
        public static XName rsidR = w + "rsidR";
        public static XName rsidRDefault = w + "rsidRDefault";
        public static XName rsidRoot = w + "rsidRoot";
        public static XName rsidRPr = w + "rsidRPr";
        public static XName rsids = w + "rsids";
        public static XName rsidSect = w + "rsidSect";
        public static XName rsidTr = w + "rsidTr";
        public static XName rStyle = w + "rStyle";
        public static XName rt = w + "rt";
        public static XName rtl = w + "rtl";
        public static XName rtlGutter = w + "rtlGutter";
        public static XName ruby = w + "ruby";
        public static XName rubyAlign = w + "rubyAlign";
        public static XName rubyBase = w + "rubyBase";
        public static XName rubyPr = w + "rubyPr";
        public static XName salt = w + "salt";
        public static XName saveFormsData = w + "saveFormsData";
        public static XName saveInvalidXml = w + "saveInvalidXml";
        public static XName savePreviewPicture = w + "savePreviewPicture";
        public static XName saveSmartTagsAsXml = w + "saveSmartTagsAsXml";
        public static XName saveSubsetFonts = w + "saveSubsetFonts";
        public static XName saveThroughXslt = w + "saveThroughXslt";
        public static XName saveXmlDataOnly = w + "saveXmlDataOnly";
        public static XName scrollbar = w + "scrollbar";
        public static XName sdt = w + "sdt";
        public static XName sdtContent = w + "sdtContent";
        public static XName sdtEndPr = w + "sdtEndPr";
        public static XName sdtPr = w + "sdtPr";
        public static XName sectPr = w + "sectPr";
        public static XName sectPrChange = w + "sectPrChange";
        public static XName selectFldWithFirstOrLastChar = w + "selectFldWithFirstOrLastChar";
        public static XName semiHidden = w + "semiHidden";
        public static XName sep = w + "sep";
        public static XName separator = w + "separator";
        public static XName settings = w + "settings";
        public static XName shadow = w + "shadow";
        public static XName shapeDefaults = w + "shapeDefaults";
        public static XName shapeid = w + "shapeid";
        public static XName shapeLayoutLikeWW8 = w + "shapeLayoutLikeWW8";
        public static XName shd = w + "shd";
        public static XName showBreaksInFrames = w + "showBreaksInFrames";
        public static XName showEnvelope = w + "showEnvelope";
        public static XName showingPlcHdr = w + "showingPlcHdr";
        public static XName showXMLTags = w + "showXMLTags";
        public static XName sig = w + "sig";
        public static XName size = w + "size";
        public static XName sizeAuto = w + "sizeAuto";
        public static XName smallCaps = w + "smallCaps";
        public static XName smartTag = w + "smartTag";
        public static XName smartTagPr = w + "smartTagPr";
        public static XName smartTagType = w + "smartTagType";
        public static XName snapToGrid = w + "snapToGrid";
        public static XName softHyphen = w + "softHyphen";
        public static XName solutionID = w + "solutionID";
        public static XName sourceFileName = w + "sourceFileName";
        public static XName space = w + "space";
        public static XName spaceForUL = w + "spaceForUL";
        public static XName spacing = w + "spacing";
        public static XName spacingInWholePoints = w + "spacingInWholePoints";
        public static XName specVanish = w + "specVanish";
        public static XName spelling = w + "spelling";
        public static XName splitPgBreakAndParaMark = w + "splitPgBreakAndParaMark";
        public static XName src = w + "src";
        public static XName start = w + "start";
        public static XName startOverride = w + "startOverride";
        public static XName statusText = w + "statusText";
        public static XName storeItemID = w + "storeItemID";
        public static XName storeMappedDataAs = w + "storeMappedDataAs";
        public static XName strictFirstAndLastChars = w + "strictFirstAndLastChars";
        public static XName strike = w + "strike";
        public static XName style = w + "style";
        public static XName styleId = w + "styleId";
        public static XName styleLink = w + "styleLink";
        public static XName styleLockQFSet = w + "styleLockQFSet";
        public static XName styleLockTheme = w + "styleLockTheme";
        public static XName stylePaneFormatFilter = w + "stylePaneFormatFilter";
        public static XName stylePaneSortMethod = w + "stylePaneSortMethod";
        public static XName styles = w + "styles";
        public static XName subDoc = w + "subDoc";
        public static XName subFontBySize = w + "subFontBySize";
        public static XName subsetted = w + "subsetted";
        public static XName suff = w + "suff";
        public static XName summaryLength = w + "summaryLength";
        public static XName suppressAutoHyphens = w + "suppressAutoHyphens";
        public static XName suppressBottomSpacing = w + "suppressBottomSpacing";
        public static XName suppressLineNumbers = w + "suppressLineNumbers";
        public static XName suppressOverlap = w + "suppressOverlap";
        public static XName suppressSpacingAtTopOfPage = w + "suppressSpacingAtTopOfPage";
        public static XName suppressSpBfAfterPgBrk = w + "suppressSpBfAfterPgBrk";
        public static XName suppressTopSpacing = w + "suppressTopSpacing";
        public static XName suppressTopSpacingWP = w + "suppressTopSpacingWP";
        public static XName swapBordersFacingPages = w + "swapBordersFacingPages";
        public static XName sym = w + "sym";
        public static XName sz = w + "sz";
        public static XName szCs = w + "szCs";
        public static XName t = w + "t";
        public static XName t1 = w + "t1";
        public static XName t2 = w + "t2";
        public static XName tab = w + "tab";
        public static XName table = w + "table";
        public static XName tabs = w + "tabs";
        public static XName tag = w + "tag";
        public static XName targetScreenSz = w + "targetScreenSz";
        public static XName tbl = w + "tbl";
        public static XName tblBorders = w + "tblBorders";
        public static XName tblCellMar = w + "tblCellMar";
        public static XName tblCellSpacing = w + "tblCellSpacing";
        public static XName tblGrid = w + "tblGrid";
        public static XName tblGridChange = w + "tblGridChange";
        public static XName tblHeader = w + "tblHeader";
        public static XName tblInd = w + "tblInd";
        public static XName tblLayout = w + "tblLayout";
        public static XName tblLook = w + "tblLook";
        public static XName tblOverlap = w + "tblOverlap";
        public static XName tblpPr = w + "tblpPr";
        public static XName tblPr = w + "tblPr";
        public static XName tblPrChange = w + "tblPrChange";
        public static XName tblPrEx = w + "tblPrEx";
        public static XName tblPrExChange = w + "tblPrExChange";
        public static XName tblpX = w + "tblpX";
        public static XName tblpXSpec = w + "tblpXSpec";
        public static XName tblpY = w + "tblpY";
        public static XName tblpYSpec = w + "tblpYSpec";
        public static XName tblStyle = w + "tblStyle";
        public static XName tblStyleColBandSize = w + "tblStyleColBandSize";
        public static XName tblStylePr = w + "tblStylePr";
        public static XName tblStyleRowBandSize = w + "tblStyleRowBandSize";
        public static XName tblW = w + "tblW";
        public static XName tc = w + "tc";
        public static XName tcBorders = w + "tcBorders";
        public static XName tcFitText = w + "tcFitText";
        public static XName tcMar = w + "tcMar";
        public static XName tcPr = w + "tcPr";
        public static XName tcPrChange = w + "tcPrChange";
        public static XName tcW = w + "tcW";
        public static XName temporary = w + "temporary";
        public static XName tentative = w + "tentative";
        public static XName text = w + "text";
        public static XName textAlignment = w + "textAlignment";
        public static XName textboxTightWrap = w + "textboxTightWrap";
        public static XName textDirection = w + "textDirection";
        public static XName textInput = w + "textInput";
        public static XName tgtFrame = w + "tgtFrame";
        public static XName themeColor = w + "themeColor";
        public static XName themeFill = w + "themeFill";
        public static XName themeFillShade = w + "themeFillShade";
        public static XName themeFillTint = w + "themeFillTint";
        public static XName themeFontLang = w + "themeFontLang";
        public static XName themeShade = w + "themeShade";
        public static XName themeTint = w + "themeTint";
        public static XName titlePg = w + "titlePg";
        public static XName tl2br = w + "tl2br";
        public static XName tmpl = w + "tmpl";
        public static XName tooltip = w + "tooltip";
        public static XName top = w + "top";
        public static XName topFromText = w + "topFromText";
        public static XName topLinePunct = w + "topLinePunct";
        public static XName tplc = w + "tplc";
        public static XName tr = w + "tr";
        public static XName tr2bl = w + "tr2bl";
        public static XName trackRevisions = w + "trackRevisions";
        public static XName trHeight = w + "trHeight";
        public static XName trPr = w + "trPr";
        public static XName trPrChange = w + "trPrChange";
        public static XName truncateFontHeightsLikeWP6 = w + "truncateFontHeightsLikeWP6";
        public static XName txbxContent = w + "txbxContent";
        public static XName type = w + "type";
        public static XName types = w + "types";
        public static XName u = w + "u";
        public static XName udl = w + "udl";
        public static XName uiCompat97To2003 = w + "uiCompat97To2003";
        public static XName uiPriority = w + "uiPriority";
        public static XName ulTrailSpace = w + "ulTrailSpace";
        public static XName underlineTabInNumList = w + "underlineTabInNumList";
        public static XName unhideWhenUsed = w + "unhideWhenUsed";
        public static XName updateFields = w + "updateFields";
        public static XName uri = w + "uri";
        public static XName url = w + "url";
        public static XName usb0 = w + "usb0";
        public static XName usb1 = w + "usb1";
        public static XName usb2 = w + "usb2";
        public static XName usb3 = w + "usb3";
        public static XName useAltKinsokuLineBreakRules = w + "useAltKinsokuLineBreakRules";
        public static XName useAnsiKerningPairs = w + "useAnsiKerningPairs";
        public static XName useFELayout = w + "useFELayout";
        public static XName useNormalStyleForList = w + "useNormalStyleForList";
        public static XName usePrinterMetrics = w + "usePrinterMetrics";
        public static XName useSingleBorderforContiguousCells = w + "useSingleBorderforContiguousCells";
        public static XName useWord2002TableStyleRules = w + "useWord2002TableStyleRules";
        public static XName useWord97LineBreakRules = w + "useWord97LineBreakRules";
        public static XName useXSLTWhenSaving = w + "useXSLTWhenSaving";
        public static XName val = w + "val";
        public static XName vAlign = w + "vAlign";
        public static XName value = w + "value";
        public static XName vAnchor = w + "vAnchor";
        public static XName vanish = w + "vanish";
        public static XName vendorID = w + "vendorID";
        public static XName vert = w + "vert";
        public static XName vertAlign = w + "vertAlign";
        public static XName vertAnchor = w + "vertAnchor";
        public static XName vertCompress = w + "vertCompress";
        public static XName view = w + "view";
        public static XName viewMergedData = w + "viewMergedData";
        public static XName vMerge = w + "vMerge";
        public static XName vMergeOrig = w + "vMergeOrig";
        public static XName vSpace = w + "vSpace";
        public static XName _w = w + "w";
        public static XName wAfter = w + "wAfter";
        public static XName wBefore = w + "wBefore";
        public static XName webHidden = w + "webHidden";
        public static XName webSettings = w + "webSettings";
        public static XName widowControl = w + "widowControl";
        public static XName wordWrap = w + "wordWrap";
        public static XName wpJustification = w + "wpJustification";
        public static XName wpSpaceWidth = w + "wpSpaceWidth";
        public static XName wrap = w + "wrap";
        public static XName wrapTrailSpaces = w + "wrapTrailSpaces";
        public static XName writeProtection = w + "writeProtection";
        public static XName x = w + "x";
        public static XName xAlign = w + "xAlign";
        public static XName xpath = w + "xpath";
        public static XName y = w + "y";
        public static XName yAlign = w + "yAlign";
        public static XName yearLong = w + "yearLong";
        public static XName yearShort = w + "yearShort";
        public static XName zoom = w + "zoom";
        public static XName zOrder = w + "zOrder";
        public static XName tblCaption = w + "tblCaption";
        public static XName tblDescription = w + "tblDescription";
        public static XName startChars = w + "startChars";
        public static XName end = w + "end";
        public static XName endChars = w + "endChars";
        public static XName evenHBand = w + "evenHBand";
        public static XName evenVBand = w + "evenVBand";
        public static XName firstRowFirstColumn = w + "firstRowFirstColumn";
        public static XName firstRowLastColumn = w + "firstRowLastColumn";
        public static XName lastRowFirstColumn = w + "lastRowFirstColumn";
        public static XName lastRowLastColumn = w + "lastRowLastColumn";
        public static XName oddHBand = w + "oddHBand";
        public static XName oddVBand = w + "oddVBand";
        public static XName headers = w + "headers";

        public static XName[] BlockLevelContentContainers =
        {
            W.body,
            W.tc,
            W.txbxContent,
            W.hdr,
            W.ftr,
            W.endnote,
            W.footnote
        };

        public static XName[] SubRunLevelContent =
        {
            W.br,
            W.cr,
            W.dayLong,
            W.dayShort,
            W.drawing,
            W.drawing,
            W.monthLong,
            W.monthShort,
            W.noBreakHyphen,
            W.ptab,
            W.pgNum,
            W.pict,
            W.softHyphen,
            W.sym,
            W.t,
            W.tab,
            W.yearLong,
            W.yearShort,
            MC.AlternateContent,
        };
    }

    public static class MC
    {
        public static XNamespace mc =
            "http://schemas.openxmlformats.org/markup-compatibility/2006";
        public static XName AlternateContent = mc + "AlternateContent";
        public static XName Choice = mc + "Choice";
        public static XName Fallback = mc + "Fallback";
        public static XName Ignorable = mc + "Ignorable";
        public static XName PreserveAttributes = mc + "PreserveAttributes";
    }

    public static class PtOpenXmlExtensions
    {
        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            try
            {
                XDocument partXDocument = part.Annotation<XDocument>();
                if (partXDocument != null)
                    return partXDocument;
                using (Stream partStream = part.GetStream())
                {
                    if (partStream.Length == 0)
                    {
                        partXDocument = new XDocument();
                        partXDocument.Declaration = new XDeclaration("1.0", "UTF-8", "yes");
                    }
                    else
                        using (XmlReader partXmlReader = XmlReader.Create(partStream))
                            partXDocument = XDocument.Load(partXmlReader);
                }
                part.AddAnnotation(partXDocument);
                return partXDocument;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private static XmlNamespaceManager GetManagerFromXDocument(XDocument xDocument)
        {
            XmlReader reader = xDocument.CreateReader();
            XDocument newXDoc = XDocument.Load(reader);
            XElement rootElement = xDocument.Elements().FirstOrDefault();
            rootElement.ReplaceWith(newXDoc.Root);
            XmlNameTable nameTable = reader.NameTable;
            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(nameTable);
            return namespaceManager;
        }

        public static XDocument GetXDocument(this OpenXmlPart part, out XmlNamespaceManager namespaceManager)
        {

            XDocument partXDocument = part.Annotation<XDocument>();
            namespaceManager = part.Annotation<XmlNamespaceManager>();
            if (partXDocument != null)
            {
                if (namespaceManager != null)
                    return partXDocument;
                namespaceManager = GetManagerFromXDocument(partXDocument);
                part.AddAnnotation(namespaceManager);
                return partXDocument;
            }

            using (Stream partStream = part.GetStream())
            {
                if (partStream.Length == 0)
                {
                    partXDocument = new XDocument();
                    partXDocument.Declaration = new XDeclaration("1.0", "UTF-8", "yes");
                    part.AddAnnotation(partXDocument);
                    return partXDocument;
                }
                else
                {
                    using (XmlReader partXmlReader = XmlReader.Create(partStream))
                    {
                        partXDocument = XDocument.Load(partXmlReader);
                        XmlNameTable nameTable = partXmlReader.NameTable;
                        namespaceManager = new XmlNamespaceManager(nameTable);
                        part.AddAnnotation(partXDocument);
                        part.AddAnnotation(namespaceManager);
                        return partXDocument;
                    }
                }
            }
        }

        public static void PutXDocument(this OpenXmlPart part)
        {
            XDocument partXDocument = part.GetXDocument();
            if (partXDocument != null)
            {
                using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
                using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                    partXDocument.Save(partXmlWriter);
            }
        }

        public static void PutXDocumentWithFormatting(this OpenXmlPart part)
        {
            XDocument partXDocument = part.GetXDocument();
            if (partXDocument != null)
            {
                using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
                {
                    XmlWriterSettings settings = new XmlWriterSettings();
                    settings.Indent = true;
                    settings.OmitXmlDeclaration = true;
                    settings.NewLineOnAttributes = true;
                    using (XmlWriter partXmlWriter = XmlWriter.Create(partStream, settings))
                        partXDocument.Save(partXmlWriter);
                }
            }
        }

        public static void PutXDocument(this OpenXmlPart part, XDocument document)
        {
            using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                document.Save(partXmlWriter);
            part.RemoveAnnotations<XDocument>();
            part.AddAnnotation(document);
        }
    }

    public static class PtExtensions
    {
        public class SiblingsReverseDocumentOrderInfo
        {
            public XElement PreviousSibling;
        }

        public static string StringConcatenate(this IEnumerable<string> source)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string s in source)
                sb.Append(s);
            return sb.ToString();
        }

        public static string StringConcatenate<T>(
            this IEnumerable<T> source,
            Func<T, string> projectionFunc)
        {
            return source.Aggregate(
                new StringBuilder(),
                (s, i) => s.Append(projectionFunc(i)),
                s => s.ToString());
        }

        public static IEnumerable<TResult> Rollup<TSource, TResult>(
            this IEnumerable<TSource> source,
            TResult seed,
            Func<TSource, TResult, TResult> projection)
        {
            TResult nextSeed = seed;
            foreach (TSource src in source)
            {
                TResult projectedValue = projection(src, nextSeed);
                nextSeed = projectedValue;
                yield return projectedValue;
            }
        }
        // public static XElement GetXElement(this XmlNode node)
        // {
        //     XDocument xDoc = new XDocument();
        //     using (XmlWriter xmlWriter = xDoc.CreateWriter())
        //         node.WriteTo(xmlWriter);
        //     return xDoc.Root;
        // }

        // public static XmlNode GetXmlNode(this XElement element)
        // {
        //     using (XmlReader xmlReader = element.CreateReader())
        //     {
        //         XmlDocument xmlDoc = new XmlDocument();
        //         xmlDoc.Load(xmlReader);
        //         return xmlDoc;
        //     }
        // }

        public static XDocument GetXDocument(this XmlDocument document)
        {
            XDocument xDoc = new XDocument();
            using (XmlWriter xmlWriter = xDoc.CreateWriter())
                document.WriteTo(xmlWriter);
            XmlDeclaration decl =
                document.ChildNodes.OfType<XmlDeclaration>().FirstOrDefault();
            if (decl != null)
                xDoc.Declaration = new XDeclaration(decl.Version, decl.Encoding,
                    decl.Standalone);
            return xDoc;
        }

        // public static XmlDocument GetXmlDocument(this XDocument document)
        // {
        //     using (XmlReader xmlReader = document.CreateReader())
        //     {
        //         XmlDocument xmlDoc = new XmlDocument();
        //         xmlDoc.Load(xmlReader);
        //         if (document.Declaration != null)
        //         {
        //             XmlDeclaration dec = xmlDoc.CreateXmlDeclaration(document.Declaration.Version,
        //                 document.Declaration.Encoding, document.Declaration.Standalone);
        //             xmlDoc.InsertBefore(dec, xmlDoc.FirstChild);
        //         }
        //         return xmlDoc;
        //     }
        // }

        // public static string StringConcatenate(this IEnumerable<string> source)
        // {
        //     StringBuilder sb = new StringBuilder();
        //     foreach (string s in source)
        //         sb.Append(s);
        //     return sb.ToString();
        // }

        // public static string StringConcatenate<T>(
        //     this IEnumerable<T> source,
        //     Func<T, string> projectionFunc)
        // {
        //     return source.Aggregate(
        //         new StringBuilder(),
        //         (s, i) => s.Append(projectionFunc(i)),
        //         s => s.ToString());
        // }

        // public static IEnumerable<TResult> PtZip<TFirst, TSecond, TResult>(
        //     this IEnumerable<TFirst> first,
        //     IEnumerable<TSecond> second,
        //     Func<TFirst, TSecond, TResult> func)
        // {
        //     var ie1 = first.GetEnumerator();
        //     var ie2 = second.GetEnumerator();
        //     while (ie1.MoveNext() && ie2.MoveNext())
        //         yield return func(ie1.Current, ie2.Current);
        // }

        public static IEnumerable<IGrouping<TKey, TSource>> GroupAdjacent<TSource, TKey>(
            this IEnumerable<TSource> source,
            Func<TSource, TKey> keySelector)
        {
            TKey last = default(TKey);
            bool haveLast = false;
            List<TSource> list = new List<TSource>();

            foreach (TSource s in source)
            {
                TKey k = keySelector(s);
                if (haveLast)
                {
                    if (!k.Equals(last))
                    {
                        yield return new GroupOfAdjacent<TSource, TKey>(list, last);
                        list = new List<TSource>();
                        list.Add(s);
                        last = k;
                    }
                    else
                    {
                        list.Add(s);
                        last = k;
                    }
                }
                else
                {
                    list.Add(s);
                    last = k;
                    haveLast = true;
                }
            }
            if (haveLast)
                yield return new GroupOfAdjacent<TSource, TKey>(list, last);
        }

        private static void InitializeSiblingsReverseDocumentOrder(XElement element)
        {
            XElement prev = null;
            foreach (XElement e in element.Elements())
            {
                e.AddAnnotation(new SiblingsReverseDocumentOrderInfo { PreviousSibling = prev });
                prev = e;
            }
        }

        public static IEnumerable<XElement> SiblingsBeforeSelfReverseDocumentOrder(
            this XElement element)
        {
            if (element.Annotation<SiblingsReverseDocumentOrderInfo>() == null)
                InitializeSiblingsReverseDocumentOrder(element.Parent);
            XElement current = element;
            while (true)
            {
                XElement previousElement = current
                    .Annotation<SiblingsReverseDocumentOrderInfo>()
                    .PreviousSibling;
                if (previousElement == null)
                    yield break;
                yield return previousElement;
                current = previousElement;
            }
        }

        // private static void InitializeDescendantsReverseDocumentOrder(XElement element)
        // {
        //     XElement prev = null;
        //     foreach (XElement e in element.Descendants())
        //     {
        //         e.AddAnnotation(new DescendantsReverseDocumentOrderInfo { PreviousElement = prev });
        //         prev = e;
        //     }
        // }

        // public static IEnumerable<XElement> DescendantsBeforeSelfReverseDocumentOrder(
        //     this XElement element)
        // {
        //     if (element.Annotation<DescendantsReverseDocumentOrderInfo>() == null)
        //         InitializeDescendantsReverseDocumentOrder(element.AncestorsAndSelf().Last());
        //     XElement current = element;
        //     while (true)
        //     {
        //         XElement previousElement = current
        //             .Annotation<DescendantsReverseDocumentOrderInfo>()
        //             .PreviousElement;
        //         if (previousElement == null)
        //             yield break;
        //         yield return previousElement;
        //         current = previousElement;
        //     }
        // }

        // private static void InitializeDescendantsTrimmedReverseDocumentOrder(XElement element, XName trimName)
        // {
        //     XElement prev = null;
        //     foreach (XElement e in element.DescendantsTrimmed(trimName))
        //     {
        //         e.AddAnnotation(new DescendantsTrimmedReverseDocumentOrderInfo { PreviousElement = prev });
        //         prev = e;
        //     }
        // }

        // public static IEnumerable<XElement> DescendantsTrimmedBeforeSelfReverseDocumentOrder(
        //     this XElement element, XName trimName)
        // {
        //     if (element.Annotation<DescendantsTrimmedReverseDocumentOrderInfo>() == null)
        //     {
        //         var ances = element.AncestorsAndSelf(W.txbxContent).FirstOrDefault();
        //         if (ances == null)
        //             ances = element.AncestorsAndSelf().Last();
        //         InitializeDescendantsTrimmedReverseDocumentOrder(ances, trimName);
        //     }

        //     XElement current = element;
        //     while (true)
        //     {
        //         XElement previousElement = current
        //             .Annotation<DescendantsTrimmedReverseDocumentOrderInfo>()
        //             .PreviousElement;
        //         if (previousElement == null)
        //             yield break;
        //         yield return previousElement;
        //         current = previousElement;
        //     }
        // }

        // public static string ToStringNewLineOnAttributes(this XElement element)
        // {
        //     XmlWriterSettings settings = new XmlWriterSettings();
        //     settings.Indent = true;
        //     settings.OmitXmlDeclaration = true;
        //     settings.NewLineOnAttributes = true;
        //     StringBuilder stringBuilder = new StringBuilder();
        //     using (StringWriter stringWriter = new StringWriter(stringBuilder))
        //     using (XmlWriter xmlWriter = XmlWriter.Create(stringWriter, settings))
        //         element.WriteTo(xmlWriter);
        //     return stringBuilder.ToString();
        // }

        public static IEnumerable<XElement> DescendantsTrimmed(this XElement element,
            XName trimName)
        {
            return DescendantsTrimmed(element, e => e.Name == trimName);
        }

        public static IEnumerable<XElement> DescendantsTrimmed(this XElement element,
            Func<XElement, bool> predicate)
        {
            Stack<IEnumerator<XElement>> iteratorStack = new Stack<IEnumerator<XElement>>();
            iteratorStack.Push(element.Elements().GetEnumerator());
            while (iteratorStack.Count > 0)
            {
                while (iteratorStack.Peek().MoveNext())
                {
                    XElement currentXElement = iteratorStack.Peek().Current;
                    if (predicate(currentXElement))
                    {
                        yield return currentXElement;
                        continue;
                    }
                    yield return currentXElement;
                    iteratorStack.Push(currentXElement.Elements().GetEnumerator());
                }
                iteratorStack.Pop();
            }
        }

        // public static IEnumerable<TResult> Rollup<TSource, TResult>(
        //     this IEnumerable<TSource> source,
        //     TResult seed,
        //     Func<TSource, TResult, TResult> projection)
        // {
        //     TResult nextSeed = seed;
        //     foreach (TSource src in source)
        //     {
        //         TResult projectedValue = projection(src, nextSeed);
        //         nextSeed = projectedValue;
        //         yield return projectedValue;
        //     }
        // }

        // public static IEnumerable<TSource> SequenceAt<TSource>(this TSource[] source, int index)
        // {
        //     int i = index;
        //     while (i < source.Length)
        //         yield return source[i++];
        // }

        // public static IEnumerable<T> SkipLast<T>(this IEnumerable<T> source,
        //     int count)
        // {
        //     Queue<T> saveList = new Queue<T>();
        //     int saved = 0;
        //     foreach (T item in source)
        //     {
        //         if (saved < count)
        //         {
        //             saveList.Enqueue(item);
        //             ++saved;
        //             continue;
        //         }
        //         saveList.Enqueue(item);
        //         yield return saveList.Dequeue();
        //     }
        //     yield break;
        // }

        // public static bool? ToBoolean(this XAttribute a)
        // {
        //     if (a == null)
        //         return null;
        //     string s = ((string)a).ToLower();
        //     if (s == "1") return true;
        //     if (s == "0") return false;
        //     if (s == "true") return true;
        //     if (s == "false") return false;
        //     if (s == "on") return true;
        //     if (s == "off") return false;
        //     return (bool)a;
        // }

        // private static string GetQName(XElement xe)
        // {
        //     string prefix = xe.GetPrefixOfNamespace(xe.Name.Namespace);
        //     if (xe.Name.Namespace == XNamespace.None || prefix == null)
        //         return xe.Name.LocalName.ToString();
        //     else
        //         return prefix + ":" + xe.Name.LocalName.ToString();
        // }

        // private static string GetQName(XAttribute xa)
        // {
        //     string prefix =
        //         xa.Parent.GetPrefixOfNamespace(xa.Name.Namespace);
        //     if (xa.Name.Namespace == XNamespace.None || prefix == null)
        //         return xa.Name.ToString();
        //     else
        //         return prefix + ":" + xa.Name.LocalName;
        // }

        // private static string NameWithPredicate(XElement el)
        // {
        //     if (el.Parent != null && el.Parent.Elements(el.Name).Count() != 1)
        //         return GetQName(el) + "[" +
        //             (el.ElementsBeforeSelf(el.Name).Count() + 1) + "]";
        //     else
        //         return GetQName(el);
        // }

        // public static string StrCat<T>(this IEnumerable<T> source,
        //     string separator)
        // {
        //     return source.Aggregate(new StringBuilder(),
        //                (sb, i) => sb
        //                    .Append(i.ToString())
        //                    .Append(separator),
        //                s => s.ToString());
        // }

        public class GroupOfAdjacent<TSource, TKey> : IEnumerable<TSource>, IGrouping<TKey, TSource>
        {
            public TKey Key { get; set; }
            private List<TSource> GroupList { get; set; }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return ((System.Collections.Generic.IEnumerable<TSource>)this).GetEnumerator();
            }

            System.Collections.Generic.IEnumerator<TSource>
                System.Collections.Generic.IEnumerable<TSource>.GetEnumerator()
            {
                foreach (var s in GroupList)
                    yield return s;
            }

            public GroupOfAdjacent(List<TSource> source, TKey key)
            {
                GroupList = source;
                Key = key;
            }
        }
    }

    public class RevisionAccepter
    {

        public static void AcceptRevisions(WordprocessingDocument doc)
        {
            AcceptRevisionsForPart(doc.MainDocumentPart);
            foreach (var part in doc.MainDocumentPart.HeaderParts)
                AcceptRevisionsForPart(part);
            foreach (var part in doc.MainDocumentPart.FooterParts)
                AcceptRevisionsForPart(part);
            if (doc.MainDocumentPart.EndnotesPart != null)
                AcceptRevisionsForPart(doc.MainDocumentPart.EndnotesPart);
            if (doc.MainDocumentPart.FootnotesPart != null)
                AcceptRevisionsForPart(doc.MainDocumentPart.FootnotesPart);
            if (doc.MainDocumentPart.StyleDefinitionsPart != null)
                AcceptRevisionsForStylesDefinitionPart(doc.MainDocumentPart.StyleDefinitionsPart);
        }

        private static void AcceptRevisionsForStylesDefinitionPart(StyleDefinitionsPart stylesDefinitionsPart)
        {
            var xDoc = stylesDefinitionsPart.GetXDocument();
            var newRoot = AcceptRevisionsForStylesTransform(xDoc.Root);
            xDoc.Root.ReplaceWith(newRoot);
            stylesDefinitionsPart.PutXDocument();
        }

        private static object AcceptRevisionsForStylesTransform(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.pPrChange || element.Name == W.rPrChange)
                    return null;
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AcceptRevisionsForStylesTransform(n)));
            }
            return node;
        }

        public static void AcceptRevisionsForPart(OpenXmlPart part)
        {
            XElement documentElement = part.GetXDocument().Root;
            documentElement = (XElement)AcceptMoveFromMoveToTransform(documentElement);
            documentElement = AcceptMoveFromRanges(documentElement);
            // AcceptParagraphEndTagsInMoveFromTransform needs rewritten similar to AcceptDeletedAndMoveFromParagraphMarks
            documentElement = (XElement)AcceptParagraphEndTagsInMoveFromTransform(documentElement);
            documentElement = AcceptDeletedAndMovedFromContentControls(documentElement);
            documentElement = AcceptDeletedAndMoveFromParagraphMarks(documentElement);
            documentElement = (XElement)RemoveRowsLeftEmptyByMoveFrom(documentElement);
            documentElement = (XElement)AcceptAllOtherRevisionsTransform(documentElement);
            documentElement = (XElement)AcceptDeletedCellsTransform(documentElement);
            documentElement = (XElement)MergeAdjacentTablesTransform(documentElement);
            documentElement.Descendants().Attributes().Where(a => a.Name == PT.UniqueId || a.Name == PT.RunIds).Remove();
            documentElement.Descendants(W.numPr).Where(np => !np.HasElements).Remove();
            XDocument newXDoc = new XDocument(documentElement);
            part.PutXDocument(newXDoc);
        }

        private static object MergeAdjacentTablesTransform(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Element(W.tbl) != null)
                {
                    var grouped = element
                        .Elements()
                        .GroupAdjacent(e =>
                        {
                            if (e.Name != W.tbl)
                                return "";
                            var bidiVisual = e.Elements(W.tblPr).Elements(W.bidiVisual).FirstOrDefault();
                            var bidiVisString = bidiVisual == null ? "" : "|bidiVisual";
                            var key = "tbl" + bidiVisString;
                            return key;
                        });

                    var newContent = grouped
                        .Select(g =>
                        {
                            if (g.Key == "" || g.Count() == 1)
                                return (object)g;
                            var rolled = g
                                .Select(tbl =>
                                {
                                    var gridCols = tbl
                                        .Elements(W.tblGrid)
                                        .Elements(W.gridCol)
                                        .Attributes(W._w)
                                        .Select(a => (int)a)
                                        .Rollup(0, (s, i) => s + i);
                                    return gridCols;
                                })
                                .SelectMany(m => m)
                                .Distinct()
                                .OrderBy(w => w)
                                .ToArray();
                            var newTable = new XElement(W.tbl,
                                g.First().Elements(W.tblPr),
                                new XElement(W.tblGrid,
                                    rolled.Select((r, i) =>
                                    {
                                        int v;
                                        if (i == 0)
                                            v = r;
                                        else
                                            v = r - rolled[i - 1];
                                        return new XElement(W.gridCol,
                                            new XAttribute(W._w, v));
                                    })),
                                g.Select(tbl =>
                                {
                                    var fixedWidthsTbl = FixWidths(tbl);
                                    var newRows = fixedWidthsTbl.Elements(W.tr)
                                        .Select(tr =>
                                        {
                                            XElement newRow = new XElement(W.tr,
                                                tr.Attributes(),
                                                tr.Elements().Where(e => e.Name != W.tc),
                                                tr.Elements(W.tc).Select(tc =>
                                                {
                                                    int? w = (int?)tc
                                                        .Elements(W.tcPr)
                                                        .Elements(W.tcW)
                                                        .Attributes(W._w)
                                                        .FirstOrDefault();
                                                    if (w == null)
                                                        return tc;
                                                    var cellsToLeft = tc
                                                        .Parent
                                                        .Elements(W.tc)
                                                        .TakeWhile(btc => btc != tc);
                                                    int widthToLeft = 0;
                                                    if (cellsToLeft.Any())
                                                        widthToLeft = cellsToLeft
                                                        .Elements(W.tcPr)
                                                        .Elements(W.tcW)
                                                        .Attributes(W._w)
                                                        .Select(wi => (int)wi)
                                                        .Sum();
                                                    var rolledPairs = new [] { new
                                                        {
                                                            GridValue = 0,
                                                            Index = 0,
                                                        }}
                                                        .Concat(
                                                            rolled
                                                            .Select((r, i) => new
                                                            {
                                                                GridValue = r,
                                                                Index = i + 1,
                                                            }));
                                                    var start = rolledPairs
                                                        .FirstOrDefault(t => t.GridValue >= widthToLeft);
                                                    if (start != null)
                                                    {
                                                        var gridsRequired = rolledPairs
                                                            .Skip(start.Index)
                                                            .TakeWhile(rp => rp.GridValue - start.GridValue < w)
                                                            .Count();
                                                        var tcPr = new XElement(W.tcPr,
                                                                tc.Elements(W.tcPr).Elements().Where(e => e.Name != W.gridSpan),
                                                                gridsRequired != 1 ?
                                                                    new XElement(W.gridSpan,
                                                                        new XAttribute(W.val, gridsRequired)) :
                                                                    null);
                                                        var orderedTcPr = new XElement(W.tcPr,
                                                            tcPr.Elements().OrderBy(e =>
                                                            {
                                                                if (Order_tcPr.ContainsKey(e.Name))
                                                                    return Order_tcPr[e.Name];
                                                                return 999;
                                                            }));
                                                        var newCell = new XElement(W.tc,
                                                            orderedTcPr,
                                                            tc.Elements().Where(e => e.Name != W.tcPr));
                                                        return newCell;
                                                    }
                                                    return tc;
                                                }));
                                            return newRow;
                                        });
                                    return newRows;
                                }));
                            return newTable;
                        });
                    return new XElement(element.Name,
                        element.Attributes(),
                        newContent);
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => MergeAdjacentTablesTransform(n)));
            }
            return node;
        }

        private static Dictionary<XName, int> Order_tcPr = new Dictionary<XName, int>
        {
            { W.cnfStyle, 10 },
            { W.tcW, 20 },
            { W.gridSpan, 30 },
            { W.hMerge, 40 },
            { W.vMerge, 50 },
            { W.tcBorders, 60 },
            { W.shd, 70 },
            { W.noWrap, 80 },
            { W.tcMar, 90 },
            { W.textDirection, 100 },
            { W.tcFitText, 110 },
            { W.vAlign, 120 },
            { W.hideMark, 130 },
            { W.headers, 140 },
        };

        private static XElement FixWidths(XElement tbl)
        {
            var newTbl = new XElement(tbl);
            var gridLines = tbl.Elements(W.tblGrid).Elements(W.gridCol).Attributes(W._w).Select(w => (int)w).ToArray();
            foreach (var tr in newTbl.Elements(W.tr))
            {
                int used = 0;
                int lastUsed = -1;
                foreach (var tc in tr.Elements(W.tc))
                {
                    var tcW = tc.Elements(W.tcPr).Elements(W.tcW).Attributes(W._w).FirstOrDefault();
                    if (tcW != null)
                    {
                        int? gridSpan = (int?)tc.Elements(W.tcPr).Elements(W.gridSpan).Attributes(W.val).FirstOrDefault();

                        if (gridSpan == null)
                            gridSpan = 1;

                        var z = Math.Min(gridLines.Length - 1, lastUsed + (int)gridSpan);
                        int w = gridLines.Where((g, i) => i > lastUsed && i <= z).Sum();
                        tcW.Value = w.ToString();

                        lastUsed += (int)gridSpan;
                        used += (int)gridSpan;
                    }
                }
            }
            return newTbl;
        }

        private static object AcceptMoveFromMoveToTransform(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.moveTo)
                    return element.Nodes().Select(n => AcceptMoveFromMoveToTransform(n));
                if (element.Name == W.moveFrom)
                    return null;
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AcceptMoveFromMoveToTransform(n)));
            }
            return node;
        }

        private static XElement AcceptMoveFromRanges(XElement document)
        {
            string wordProcessingNamespacePrefix = document.GetPrefixOfNamespace(W.w);

            // The following lists contain the elements that are between start/end elements.
            List<XElement> startElementTagsInMoveFromRange = new List<XElement>();
            List<XElement> endElementTagsInMoveFromRange = new List<XElement>();

            // Following are the elements that *may* be in a range that has both start and end
            // elements.
            Dictionary<string, PotentialInRangeElements> potentialDeletedElements =
                new Dictionary<string, PotentialInRangeElements>();

            foreach (var tag in DescendantAndSelfTags(document))
            {
                if (tag.Element.Name == W.moveFromRangeStart)
                {
                    string id = tag.Element.Attribute(W.id).Value;
                    potentialDeletedElements.Add(id, new PotentialInRangeElements());
                    continue;
                }
                if (tag.Element.Name == W.moveFromRangeEnd)
                {
                    string id = tag.Element.Attribute(W.id).Value;
                    if (potentialDeletedElements.ContainsKey(id))
                    {
                        startElementTagsInMoveFromRange.AddRange(
                            potentialDeletedElements[id].PotentialStartElementTagsInRange);
                        endElementTagsInMoveFromRange.AddRange(
                            potentialDeletedElements[id].PotentialEndElementTagsInRange);
                        potentialDeletedElements.Remove(id);
                    }
                    continue;
                }
                if (potentialDeletedElements.Count > 0)
                {
                    if (tag.TagType == TagTypeEnum.Element &&
                        (tag.Element.Name != W.moveFromRangeStart &&
                         tag.Element.Name != W.moveFromRangeEnd))
                    {
                        foreach (var id in potentialDeletedElements)
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                        continue;
                    }
                    if (tag.TagType == TagTypeEnum.EmptyElement &&
                        (tag.Element.Name != W.moveFromRangeStart &&
                         tag.Element.Name != W.moveFromRangeEnd))
                    {
                        foreach (var id in potentialDeletedElements)
                        {
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        }
                        continue;
                    }
                    if (tag.TagType == TagTypeEnum.EndElement &&
                        (tag.Element.Name != W.moveFromRangeStart &&
                        tag.Element.Name != W.moveFromRangeEnd))
                    {
                        foreach (var id in potentialDeletedElements)
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        continue;
                    }
                }
            }
            var moveFromElementsToDelete = startElementTagsInMoveFromRange
                .Intersect(endElementTagsInMoveFromRange)
                .ToArray();
            if (moveFromElementsToDelete.Count() > 0)
                return (XElement)AcceptMoveFromRangesTransform(
                    document, moveFromElementsToDelete);
            return document;
        }

        private enum MoveFromCollectionType
        {
            ParagraphEndTagInMoveFromRange,
            Other
        };

        private static object AcceptParagraphEndTagsInMoveFromTransform(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (W.BlockLevelContentContainers.Contains(element.Name))
                {
                    var groupedBodyChildren = element
                        .Elements()
                        .GroupAdjacent(c =>
                        {
                            BlockContentInfo pi = c.GetParagraphInfo();
                            if (pi.ThisBlockContentElement != null)
                            {
                                bool paragraphMarkIsInMoveFromRange =
                                    pi.ThisBlockContentElement.Elements(W.moveFromRangeStart).Any() &&
                                    !pi.ThisBlockContentElement.Elements(W.moveFromRangeEnd).Any();
                                if (paragraphMarkIsInMoveFromRange)
                                    return MoveFromCollectionType.ParagraphEndTagInMoveFromRange;
                            }
                            XElement previousContentElement = c.ContentElementsBeforeSelf()
                                .Where(e => e.GetParagraphInfo().ThisBlockContentElement != null)
                                .FirstOrDefault();
                            if (previousContentElement != null)
                            {
                                BlockContentInfo pi2 = previousContentElement.GetParagraphInfo();
                                if (c.Name == W.p &&
                                    pi2.ThisBlockContentElement.Elements(W.moveFromRangeStart).Any() &&
                                    !pi2.ThisBlockContentElement.Elements(W.moveFromRangeEnd).Any())
                                    return MoveFromCollectionType.ParagraphEndTagInMoveFromRange;
                            }
                            return MoveFromCollectionType.Other;
                        })
                        .ToList();

                    // If there is only one group, and it's key is MoveFromCollectionType.Other
                    // then there is nothing to do.
                    if (groupedBodyChildren.Count() == 1 &&
                        groupedBodyChildren.First().Key == MoveFromCollectionType.Other)
                    {
                        XElement newElement = new XElement(element.Name,
                            element.Attributes(),
                            groupedBodyChildren.Select(g =>
                            {
                                if (g.Key == MoveFromCollectionType.Other)
                                    return (object)g;

                                // This is a transform that produces the first element in the
                                // collection, except that the paragraph in the descendents is
                                // replaced with a new paragraph that contains all contents of the
                                // existing paragraph, plus subsequent elements in the group
                                // collection, where the paragraph in each of those groups is
                                // collapsed.
                                return CoalesqueParagraphEndTagsInMoveFromTransform(g.First(), g);
                            }));
                        return newElement;
                    }
                    else
                        return new XElement(element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n =>
                                AcceptParagraphEndTagsInMoveFromTransform(n)));
                }
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AcceptParagraphEndTagsInMoveFromTransform(n)));
            }
            return node;
        }

        private static object AcceptAllOtherRevisionsTransform(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                /// Accept inserted text, inserted paragraph marks, etc.
                /// Collapse all w:ins elements.

                if (element.Name == W.ins)
                    return element
                        .Nodes()
                        .Select(n => AcceptAllOtherRevisionsTransform(n));

                /// Remove all of the following elements.  These elements are processed in:
                ///   AcceptDeletedAndMovedFromContentControls
                ///   AcceptMoveFromMoveToTransform
                ///   AcceptDeletedAndMoveFromParagraphMarksTransform
                ///   AcceptParagraphEndTagsInMoveFromTransform
                ///   AcceptMoveFromRanges

                if (element.Name == W.customXmlDelRangeStart ||
                    element.Name == W.customXmlDelRangeEnd ||
                    element.Name == W.customXmlInsRangeStart ||
                    element.Name == W.customXmlInsRangeEnd ||
                    element.Name == W.customXmlMoveFromRangeStart ||
                    element.Name == W.customXmlMoveFromRangeEnd ||
                    element.Name == W.customXmlMoveToRangeStart ||
                    element.Name == W.customXmlMoveToRangeEnd ||
                    element.Name == W.moveFromRangeStart ||
                    element.Name == W.moveFromRangeEnd ||
                    element.Name == W.moveToRangeStart ||
                    element.Name == W.moveToRangeEnd)
                    return null;

                /// Accept revisions in formatting on paragraphs.
                /// Accept revisions in formatting on runs.
                /// Accept revisions for applied styles to a table.
                /// Accept revisions for grid revisions to a table.
                /// Accept revisions for column properties.
                /// Accept revisions for row properties.
                /// Accept revisions for table level property exceptions.
                /// Accept revisions for section properties.
                /// Accept numbering revision in fields.
                /// Accept deleted field code text.
                /// Accept deleted literal text.
                /// Accept inserted cell.

                if (element.Name == W.pPrChange ||
                    element.Name == W.rPrChange ||
                    element.Name == W.tblPrChange ||
                    element.Name == W.tblGridChange ||
                    element.Name == W.tcPrChange ||
                    element.Name == W.trPrChange ||
                    element.Name == W.tblPrExChange ||
                    element.Name == W.sectPrChange ||
                    element.Name == W.numberingChange ||
                    element.Name == W.delInstrText ||
                    element.Name == W.delText ||
                    element.Name == W.cellIns)
                    return null;

                // Accept revisions for deleted math control character.
                // Match m:f/m:fPr/m:ctrlPr/w:del, remove m:f.

                if (element.Name == M.f &&
                    element.Elements(M.fPr).Elements(M.ctrlPr).Elements(W.del).Any())
                    return null;

                // Accept revisions for deleted rows in tables.
                // Match w:tr/w:trPr/w:del, remove w:tr.

                if (element.Name == W.tr &&
                    element.Elements(W.trPr).Elements(W.del).Any())
                    return null;

                // Accept deleted text in paragraphs.

                if (element.Name == W.del)
                    return null;

                // Accept revisions for vertically merged cells.
                //   cellMerge with a parent of tcPr, with attribute w:vMerge="rest" transformed
                //     to <w:vMerge w:val="restart"/>
                //   cellMerge with a parent of tcPr, with attribute w:vMerge="cont" transformed
                //     to <w:vMerge w:val="continue"/>

                if (element.Name == W.cellMerge &&
                    element.Parent.Name == W.tcPr &&
                    (string)element.Attribute(W.vMerge) == "rest")
                    return new XElement(W.vMerge,
                        new XAttribute(W.val, "restart"));
                if (element.Name == W.cellMerge &&
                    element.Parent.Name == W.tcPr &&
                    (string)element.Attribute(W.vMerge) == "cont")
                    return new XElement(W.vMerge,
                        new XAttribute(W.val, "continue"));

                // Otherwise do identity clone.
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AcceptAllOtherRevisionsTransform(n)));
            }
            return node;
        }

        private static object CollapseParagraphTransform(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.p)
                    return element.Elements().Where(e => e.Name != W.pPr);
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => CollapseParagraphTransform(n)));
            }
            return node;
        }

        private enum DeletedParagraphCollectionType
        {
            DeletedParagraphMarkContent,
            ParagraphFollowing,
            Other
        };

        private static XElement CoalesqueParagraphDeletedAndMoveFromParagraphMarksTransform(
            IGrouping<DeletedParagraphCollectionType, BlockContentInfo> g,
            IGrouping<DeletedParagraphCollectionType, BlockContentInfo> nextGroup)
        {
            // This function constructs a paragraph.
            XElement newParagraph = new XElement(W.p,
                nextGroup.First().ThisBlockContentElement.Elements(W.pPr),
                g.Select(z => CollapseParagraphTransform(z.ThisBlockContentElement)),
                nextGroup.Select(z => CollapseParagraphTransform(z.ThisBlockContentElement)));

            return newParagraph;
        }

        private static XElement AssembleWithBlockLevelContentControlTransform(XElement element)
        {
            return new XElement(element.Name,
                element.Attributes(),
                element.Elements().Select(e => AcceptDeletedAndMoveFromParagraphMarksTransform(e)));
        }

        /// Accept deleted paragraphs.
        ///
        /// Group together all paragraphs that contain w:p/w:pPr/w:rPr/w:del elements.  Make a
        /// second group for the content element immediately following a paragraph that contains
        /// a w:del element.  The code uses the approach of dealing with paragraph content at
        /// 'levels', ignoring paragraph content at other levels.  Form a new paragraph that
        /// contains the content of the grouped paragraphs with deleted paragraph marks, and the
        /// content of the paragraph immediately following a paragraph that contains a deleted
        /// paragraph mark.  Include in the new paragraph the paragraph properties from the
        /// paragraph following.  When assembling the new paragraph, use a transform that collapses
        /// the paragraph nodes when adding content, thereby preserving custom XML and content
        /// controls.

        private static void AnnotateBlockContentElements(XElement contentContainer)
        {
            // For convenience, there is a ParagraphInfo annotation on the contentContainer.
            // It contains the same information as the ParagraphInfo annotation on the first
            //   paragraph.
            if (contentContainer.Annotation<BlockContentInfo>() != null)
                return;
            XElement firstContentElement = contentContainer
                .Elements()
                .DescendantsAndSelf()
                .FirstOrDefault(e => e.Name == W.p || e.Name == W.tbl);
            if (firstContentElement == null)
                return;

            // Add the annotation on the contentContainer.
            BlockContentInfo currentContentInfo = new BlockContentInfo()
            {
                PreviousBlockContentElement = null,
                ThisBlockContentElement = firstContentElement,
                NextBlockContentElement = null
            };
            // Add as annotation even though NextParagraph is not set yet.
            contentContainer.AddAnnotation(currentContentInfo);
            while (true)
            {
                currentContentInfo.ThisBlockContentElement.AddAnnotation(currentContentInfo);
                // Find next sibling content element.
                XElement nextContentElement = null;
                XElement current = currentContentInfo.ThisBlockContentElement;
                while (true)
                {
                    nextContentElement = current
                        .ElementsAfterSelf()
                        .DescendantsAndSelf()
                        .FirstOrDefault(e => e.Name == W.p || e.Name == W.tbl);
                    if (nextContentElement != null)
                    {
                        currentContentInfo.NextBlockContentElement = nextContentElement;
                        break;
                    }
                    current = current.Parent;
                    // When we've backed up the tree to the contentContainer, we're done.
                    if (current == contentContainer)
                        return;
                }
                currentContentInfo = new BlockContentInfo()
                {
                    PreviousBlockContentElement = currentContentInfo.ThisBlockContentElement,
                    ThisBlockContentElement = nextContentElement,
                    NextBlockContentElement = null
                };
            }
        }

        private static IEnumerable<BlockContentInfo> IterateBlockContentElements(XElement element)
        {
            XElement current = element.Elements().FirstOrDefault();
            if (current == null)
                yield break;
            AnnotateBlockContentElements(element);
            BlockContentInfo currentBlockContentInfo = element.Annotation<BlockContentInfo>();
            if (currentBlockContentInfo != null)
            {
                while (true)
                {
                    yield return currentBlockContentInfo;
                    if (currentBlockContentInfo.NextBlockContentElement == null)
                        yield break;
                    currentBlockContentInfo = currentBlockContentInfo.NextBlockContentElement.Annotation<BlockContentInfo>();
                }
            }
        }

        public static class PT
        {
            public static XNamespace pt = "http://www.codeplex.com/PowerTools/2009/RevisionAccepter";
            public static XName UniqueId = pt + "UniqueId";
            public static XName RunIds = pt + "RunIds";
        }

        private static void AnnotateRunElementsWithId(XElement element)
        {
            int runId = 0;
            foreach (XElement e in element.Descendants().Where(e => e.Name == W.r))
            {
                if (e.Name == W.r)
                    e.Add(new XAttribute(PT.UniqueId, runId++));
            }
        }

        // TODO - Need to test, need to trim so that Descendants doesn't see runs under text boxes.
        private static void AnnotateContentControlsWithRunIds(XElement element)
        {
            int sdtId = 0;
            foreach (XElement e in element.Descendants(W.sdt))
            {
                // old version
                //e.Add(new XAttribute(PT.RunIds,
                //    e.Descendants(W.r).Select(r => r.Attribute(PT.UniqueId).Value).StringConcatenate(s => s + ",").Trim(',')),
                //    new XAttribute(PT.UniqueId, sdtId++));
                e.Add(new XAttribute(PT.RunIds,
                    e.DescendantsTrimmed(W.txbxContent)
                     .Where(d => d.Name == W.r)
                     .Select(r => r.Attribute(PT.UniqueId).Value)
                     .StringConcatenate(s => s + ",")
                     .Trim(',')),
                    new XAttribute(PT.UniqueId, sdtId++));
            }
        }

        private static XElement AddBlockLevelContentControls(XElement newDocument, XElement original)
        {
            var originalContentControls = original.Descendants(W.sdt).ToList();
            var existingContentControls = newDocument.Descendants(W.sdt).ToList();
            var contentControlsToAdd = originalContentControls
                .Select(occ => occ.Attribute(PT.UniqueId).Value)
                .Except(existingContentControls
                    .Select(ecc => ecc.Attribute(PT.UniqueId).Value));
            foreach (var contentControl in originalContentControls
                .Where(occ => contentControlsToAdd.Contains(occ.Attribute(PT.UniqueId).Value)))
            {
                // TODO - Need a slight modification here.  If there is a paragraph
                // in the content control that contains no runs, then the paragraph isn't included in the
                // content control, because the following triggers off of runs.
                // To see an example of this, see example document "NumberingParagraphPropertiesChange.docxs"

                // find list of runs to surround
                var runIds = contentControl.Attribute(PT.RunIds).Value.Split(',');
                var runs = contentControl.Descendants(W.r).Where(r => runIds.Contains(r.Attribute(PT.UniqueId).Value));
                // find the runs in the new document

                var runsInNewDocument = runs.Select(r => newDocument.Descendants(W.r).First(z => z.Attribute(PT.UniqueId).Value == r.Attribute(PT.UniqueId).Value)).ToList();

                // find common ancestor
                List<XElement> runAncestorIntersection = null;
                foreach (var run in runsInNewDocument)
                {
                    if (runAncestorIntersection == null)
                        runAncestorIntersection = run.Ancestors().ToList();
                    else
                        runAncestorIntersection = run.Ancestors().Intersect(runAncestorIntersection).ToList();
                }
                if (runAncestorIntersection == null)
                    continue;
                XElement commonAncestor = runAncestorIntersection.InDocumentOrder().Last();
                // find child of common ancestor that contains first run
                // find child of common ancestor that contains last run
                // create new common ancestor:
                //   elements before first run child
                //   add content control, and runs from first run child to last run child
                //   elements after last run child
                var firstRunChild = commonAncestor
                    .Elements()
                    .First(c => c.DescendantsAndSelf()
                        .Any(z => z.Name == W.r &&
                             z.Attribute(PT.UniqueId).Value == runsInNewDocument.First().Attribute(PT.UniqueId).Value));
                var lastRunChild = commonAncestor
                    .Elements()
                    .First(c => c.DescendantsAndSelf()
                        .Any(z => z.Name == W.r &&
                             z.Attribute(PT.UniqueId).Value == runsInNewDocument.Last().Attribute(PT.UniqueId).Value));

                /// If the list of runs for the content control is exactly the list of runs for the paragraph, then
                /// create the content control surrounding the paragraph, not surrounding the runs.

                if (commonAncestor.Name == W.p &&
                    commonAncestor.Elements()
                        .Where(e => e.Name != W.pPr &&
                            e.Name != W.commentRangeStart &&
                            e.Name != W.commentRangeEnd)
                        .FirstOrDefault() == firstRunChild &&
                    commonAncestor.Elements()
                        .Where(e => e.Name != W.pPr &&
                            e.Name != W.commentRangeStart &&
                            e.Name != W.commentRangeEnd)
                        .LastOrDefault() == lastRunChild)
                {
                    // replace commonAncestor with content control containing commonAncestor
                    XElement newContentControl = new XElement(contentControl.Name,
                        contentControl.Attributes(),
                        contentControl.Elements().Where(e => e.Name != W.sdtContent),
                        new XElement(W.sdtContent, commonAncestor));

                    XElement newContentControlOrdered = new XElement(contentControl.Name,
                        contentControl.Attributes(),
                        contentControl.Elements().OrderBy(e =>
                        {
                            if (Order_sdt.ContainsKey(e.Name))
                                return Order_sdt[e.Name];
                            return 999;
                        }));

                    commonAncestor.ReplaceWith(newContentControlOrdered);
                    continue;
                }

                List<XElement> elementsBeforeRange = commonAncestor
                    .Elements()
                    .TakeWhile(e => e != firstRunChild)
                    .ToList();
                List<XElement> elementsInRange = commonAncestor
                    .Elements()
                    .SkipWhile(e => e != firstRunChild)
                    .TakeWhile(e => e != lastRunChild.ElementsAfterSelf().FirstOrDefault())
                    .ToList();
                List<XElement> elementsAfterRange = commonAncestor
                    .Elements()
                    .SkipWhile(e => e != lastRunChild.ElementsAfterSelf().FirstOrDefault())
                    .ToList();

                // detatch from current parent
                commonAncestor.Elements().Remove();

                XElement newContentControl2 = new XElement(contentControl.Name,
                    contentControl.Attributes(),
                    contentControl.Elements().Where(e => e.Name != W.sdtContent),
                    new XElement(W.sdtContent, elementsInRange));

                XElement newContentControlOrdered2 = new XElement(newContentControl2.Name,
                    newContentControl2.Attributes(),
                    newContentControl2.Elements().OrderBy(e =>
                    {
                        if (Order_sdt.ContainsKey(e.Name))
                            return Order_sdt[e.Name];
                        return 999;
                    }));

                commonAncestor.Add(
                    elementsBeforeRange,
                    newContentControlOrdered2,
                    elementsAfterRange);
            }
            return newDocument;
        }

        private static Dictionary<XName, int> Order_sdt = new Dictionary<XName, int>
        {
            { W.sdtPr, 10 },
            { W.sdtEndPr, 20 },
            { W.sdtContent, 30 },
            { W.bookmarkStart, 40 },
            { W.bookmarkEnd, 50 },
        };

        private static XElement AcceptDeletedAndMoveFromParagraphMarks(XElement element)
        {
            AnnotateRunElementsWithId(element);
            AnnotateContentControlsWithRunIds(element);
            XElement newElement = (XElement)AcceptDeletedAndMoveFromParagraphMarksTransform(element);
            XElement withBlockLevelContentControls = AddBlockLevelContentControls(newElement, element);
            return withBlockLevelContentControls;
        }

        private static object AcceptDeletedAndMoveFromParagraphMarksTransform(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (W.BlockLevelContentContainers.Contains(element.Name))
                {
                    var groupedParagraphSiblings = IterateBlockContentElements(element)
                        .GroupAdjacent(c =>
                        {
                            bool paragraphMarkIsDeletedOrMovedFrom = c
                                .ThisBlockContentElement
                                .Elements(W.pPr)
                                .Elements(W.rPr)
                                .Elements()
                                .Where(e => e.Name == W.del || e.Name == W.moveFrom)
                                .Any();
                            if (paragraphMarkIsDeletedOrMovedFrom)
                                return DeletedParagraphCollectionType.DeletedParagraphMarkContent;
                            if (c.PreviousBlockContentElement != null)
                            {
                                paragraphMarkIsDeletedOrMovedFrom = c
                                    .PreviousBlockContentElement
                                    .Elements(W.pPr)
                                    .Elements(W.rPr)
                                    .Elements()
                                    .Where(e => e.Name == W.del || e.Name == W.moveFrom)
                                    .Any();
                                if (c.ThisBlockContentElement.Name == W.p && paragraphMarkIsDeletedOrMovedFrom)
                                    return DeletedParagraphCollectionType.ParagraphFollowing;
                            }
                            return DeletedParagraphCollectionType.Other;
                        })
                        .ToList();

                    // Create a new block level content container.
                    XElement newParagraphParentElement = new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Where(e => e.Name == W.tcPr),
                        groupedParagraphSiblings.Select((g, i) =>
                        {
                            // THIS IS THE OLD CODE
                            //if (g.Key == DeletedParagraphCollectionType.Other)
                            //    return g.Select(n =>
                            //        AcceptDeletedAndMoveFromParagraphMarksTransform(n.ThisBlockContentElement));
                            if (g.Key == DeletedParagraphCollectionType.Other)
                                return g.Select(n =>
                                    AssembleWithBlockLevelContentControlTransform(n.ThisBlockContentElement));

                            // This is a transform that produces the first element in the
                            // collection; the paragraph in the descendents is replaced with a
                            // new paragraph that contains all contents of the existing paragraph,
                            // plus subsequent elements in the group collection, where the
                            // paragraph in each of those groups is collapsed.

                            if (g.Key ==
                                DeletedParagraphCollectionType.DeletedParagraphMarkContent)
                            {
                                if (i < groupedParagraphSiblings.Count() - 1)
                                {
                                    var nextG = groupedParagraphSiblings.ElementAt(i + 1);
                                    if (nextG.Key ==
                                        DeletedParagraphCollectionType.ParagraphFollowing)
                                    {
                                        XElement newParagraph = (XElement)
                                            CoalesqueParagraphDeletedAndMoveFromParagraphMarksTransform(
                                                g, nextG);
                                        return (object)newParagraph;
                                    }
                                    if (AllContentDeleted(g))
                                        return null;
                                    // todo need to handle if not all content is deleted
                                }
                                return g.Select(n =>
                                    AcceptDeletedAndMoveFromParagraphMarksTransform(n.ThisBlockContentElement));
                            }

                            // Groups with DeletedParagraphCollectionType.ParagraphFollowing
                            // have their content incorporated when processing
                            // DeletedParagraphCollectionType.DeletedParagraphMarkContent.

                            return null;
                        }),
                        element.Elements().Where(e => e.Name == W.sectPr)
                    );
                    return newParagraphParentElement;
                }

                // Otherwise, identity clone.
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AcceptDeletedAndMoveFromParagraphMarksTransform(n)));
            }
            return node;
        }

        private static bool AllContentDeleted(IGrouping<DeletedParagraphCollectionType, BlockContentInfo> g)
        {
            var someNotDeleted = g.Any(b => ! AllParaContentIsDeleted(b.ThisBlockContentElement));
            return ! someNotDeleted;
        }

        // Determine if the paragraph contains any content that is not deleted.
        private static bool AllParaContentIsDeleted(XElement p)
        {
            // needs collapse
            // dir, bdo, sdt, ins, moveTo, smartTag
            var testP = (XElement)CollapseTransform(p);

            var childElements = testP.Elements();
            var contentElements = childElements
                .Where(ce =>
                {
                    var b = IsRunContent(ce.Name);
                    if (b != null)
                        return (bool)b;
                    throw new Exception("Internal error 20, found element " + ce.Name.ToString());
                });
            if (contentElements.Any())
                return false;
            return true;
        }

        // dir, bdo, sdt, ins, moveTo, smartTag
        private static object CollapseTransform(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.dir ||
                    element.Name == W.bdr ||
                    element.Name == W.ins ||
                    element.Name == W.moveTo ||
                    element.Name == W.smartTag)
                    return element.Elements();

                if (element.Name == W.sdt)
                    return element.Elements(W.sdtContent).Elements();

                if (element.Name == W.pPr)
                    return null;

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => CollapseTransform(n)));
            }
            return node;
        }

        private static bool? IsRunContent(XName ceName)
        {
            // is content
            // r, fldSimple, hyperlink, oMath, oMathPara, subDoc
            if (ceName == W.r ||
                ceName == W.fldSimple ||
                ceName == W.hyperlink ||
                ceName == W.oMath ||
                ceName == W.subDoc)
                return true;

            // not content
            // bookmarkStart, bookmarkEnd, commentRangeStart, commentRangeEnd, del, moveFrom, proofErr
            if (ceName == W.bookmarkStart ||
                ceName == W.bookmarkEnd ||
                ceName == W.commentRangeStart ||
                ceName == W.commentRangeEnd ||
                ceName == W.del ||
                ceName == W.moveFrom ||
                ceName == W.proofErr)
                return false;

            return null;
        }

        private static IEnumerable<Tag> DescendantAndSelfTags(XElement element)
        {
            yield return new Tag
            {
                Element = element,
                TagType = TagTypeEnum.Element
            };
            Stack<IEnumerator<XElement>> iteratorStack = new Stack<IEnumerator<XElement>>();
            iteratorStack.Push(element.Elements().GetEnumerator());
            while (iteratorStack.Count > 0)
            {
                if (iteratorStack.Peek().MoveNext())
                {
                    XElement currentXElement = iteratorStack.Peek().Current;
                    if (!currentXElement.Nodes().Any())
                    {
                        yield return new Tag()
                        {
                            Element = currentXElement,
                            TagType = TagTypeEnum.EmptyElement
                        };
                        continue;
                    }
                    yield return new Tag()
                    {
                        Element = currentXElement,
                        TagType = TagTypeEnum.Element
                    };
                    iteratorStack.Push(currentXElement.Elements().GetEnumerator());
                    continue;
                }
                iteratorStack.Pop();
                if (iteratorStack.Count > 0)
                    yield return new Tag()
                    {
                        Element = iteratorStack.Peek().Current,
                        TagType = TagTypeEnum.EndElement
                    };
            }
            yield return new Tag
            {
                Element = element,
                TagType = TagTypeEnum.EndElement
            };
        }

        private class PotentialInRangeElements
        {
            public List<XElement> PotentialStartElementTagsInRange;
            public List<XElement> PotentialEndElementTagsInRange;

            public PotentialInRangeElements()
            {
                PotentialStartElementTagsInRange = new List<XElement>();
                PotentialEndElementTagsInRange = new List<XElement>();
            }
        }

        private enum TagTypeEnum
        {
            Element,
            EndElement,
            EmptyElement
        }

        private class Tag
        {
            public XElement Element;
            public TagTypeEnum TagType;
        }

        private static object AcceptDeletedAndMovedFromContentControlsTransform(XNode node,
            XElement[] contentControlElementsToCollapse,
            XElement[] moveFromElementsToDelete)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.sdt && contentControlElementsToCollapse.Contains(element))
                    return element
                        .Element(W.sdtContent)
                        .Nodes()
                        .Select(n => AcceptDeletedAndMovedFromContentControlsTransform(
                            n, contentControlElementsToCollapse, moveFromElementsToDelete));
                if (moveFromElementsToDelete.Contains(element))
                    return null;
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AcceptDeletedAndMovedFromContentControlsTransform(
                        n, contentControlElementsToCollapse, moveFromElementsToDelete)));
            }
            return node;
        }

        private static XElement AcceptDeletedAndMovedFromContentControls(XElement documentRootElement)
        {
            string wordProcessingNamespacePrefix = documentRootElement.GetPrefixOfNamespace(W.w);

            // The following lists contain the elements that are between start/end elements.
            List<XElement> startElementTagsInDeleteRange = new List<XElement>();
            List<XElement> endElementTagsInDeleteRange = new List<XElement>();
            List<XElement> startElementTagsInMoveFromRange = new List<XElement>();
            List<XElement> endElementTagsInMoveFromRange = new List<XElement>();

            // Following are the elements that *may* be in a range that has both start and end
            // elements.
            Dictionary<string, PotentialInRangeElements> potentialDeletedElements =
                new Dictionary<string, PotentialInRangeElements>();
            Dictionary<string, PotentialInRangeElements> potentialMoveFromElements =
                new Dictionary<string, PotentialInRangeElements>();

            foreach (var tag in DescendantAndSelfTags(documentRootElement))
            {
                if (tag.Element.Name == W.customXmlDelRangeStart)
                {
                    string id = tag.Element.Attribute(W.id).Value;
                    potentialDeletedElements.Add(id, new PotentialInRangeElements());
                    continue;
                }
                if (tag.Element.Name == W.customXmlDelRangeEnd)
                {
                    string id = tag.Element.Attribute(W.id).Value;
                    if (potentialDeletedElements.ContainsKey(id))
                    {
                        startElementTagsInDeleteRange.AddRange(
                            potentialDeletedElements[id].PotentialStartElementTagsInRange);
                        endElementTagsInDeleteRange.AddRange(
                            potentialDeletedElements[id].PotentialEndElementTagsInRange);
                        potentialDeletedElements.Remove(id);
                    }
                    continue;
                }
                if (tag.Element.Name == W.customXmlMoveFromRangeStart)
                {
                    string id = tag.Element.Attribute(W.id).Value;
                    potentialMoveFromElements.Add(id, new PotentialInRangeElements());
                    continue;
                }
                if (tag.Element.Name == W.customXmlMoveFromRangeEnd)
                {
                    string id = tag.Element.Attribute(W.id).Value;
                    if (potentialMoveFromElements.ContainsKey(id))
                    {
                        startElementTagsInMoveFromRange.AddRange(
                            potentialMoveFromElements[id].PotentialStartElementTagsInRange);
                        endElementTagsInMoveFromRange.AddRange(
                            potentialMoveFromElements[id].PotentialEndElementTagsInRange);
                        potentialMoveFromElements.Remove(id);
                    }
                    continue;
                }
                if (tag.Element.Name == W.sdt)
                {
                    if (tag.TagType == TagTypeEnum.Element)
                    {
                        foreach (var id in potentialDeletedElements)
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                        foreach (var id in potentialMoveFromElements)
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                        continue;
                    }
                    if (tag.TagType == TagTypeEnum.EmptyElement)
                    {
                        foreach (var id in potentialDeletedElements)
                        {
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        }
                        foreach (var id in potentialMoveFromElements)
                        {
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        }
                        continue;
                    }
                    if (tag.TagType == TagTypeEnum.EndElement)
                    {
                        foreach (var id in potentialDeletedElements)
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        foreach (var id in potentialMoveFromElements)
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        continue;
                    }
                    throw new Exception("Should not have reached this point.");
                }
                if (potentialMoveFromElements.Count() > 0 &&
                    tag.Element.Name != W.moveFromRangeStart &&
                    tag.Element.Name != W.moveFromRangeEnd &&
                    tag.Element.Name != W.customXmlMoveFromRangeStart &&
                    tag.Element.Name != W.customXmlMoveFromRangeEnd)
                {
                    if (tag.TagType == TagTypeEnum.Element)
                    {
                        foreach (var id in potentialMoveFromElements)
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                        continue;
                    }
                    if (tag.TagType == TagTypeEnum.EmptyElement)
                    {
                        foreach (var id in potentialMoveFromElements)
                        {
                            id.Value.PotentialStartElementTagsInRange.Add(tag.Element);
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        }
                        continue;
                    }
                    if (tag.TagType == TagTypeEnum.EndElement)
                    {
                        foreach (var id in potentialMoveFromElements)
                            id.Value.PotentialEndElementTagsInRange.Add(tag.Element);
                        continue;
                    }
                }
            }

            var contentControlElementsToCollapse = startElementTagsInDeleteRange
                .Intersect(endElementTagsInDeleteRange)
                .ToArray();
            var elementsToDeleteBecauseMovedFrom = startElementTagsInMoveFromRange
                .Intersect(endElementTagsInMoveFromRange)
                .ToArray();
            if (contentControlElementsToCollapse.Length > 0 ||
                elementsToDeleteBecauseMovedFrom.Length > 0)
            {
                var newDoc = AcceptDeletedAndMovedFromContentControlsTransform(documentRootElement,
                    contentControlElementsToCollapse, elementsToDeleteBecauseMovedFrom);
                return newDoc as XElement;
            }
            else
                return documentRootElement;
        }

        private static object AcceptMoveFromRangesTransform(XNode node,
            XElement[] elementsToDelete)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (elementsToDelete.Contains(element))
                    return null;
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n =>
                        AcceptMoveFromRangesTransform(n, elementsToDelete)));
            }
            return node;
        }

        private static object CoalesqueParagraphEndTagsInMoveFromTransform(XNode node,
            IGrouping<MoveFromCollectionType, XElement> g)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.p)
                    return new XElement(W.p,
                        element.Attributes(),
                        element.Elements(),
                        g.Skip(1).Select(p => CollapseParagraphTransform(p)));
                else
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(n =>
                            CoalesqueParagraphEndTagsInMoveFromTransform(n, g)));
            }
            return node;
        }

        private enum DeletedCellCollectionType
        {
            DeletedCell,
            Other
        };

        // For each table row, group deleted cells plus the cell before any deleted cell.
        // Produce a new cell that has gridSpan set appropriately for group, and clone everything
        // else.
        private static object AcceptDeletedCellsTransform(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.tr)
                {
                    var groupedCells = element
                        .Elements()
                        .GroupAdjacent(e =>
                        {
                            XElement cellAfter = e.ElementsAfterSelf(W.tc).FirstOrDefault();
                            bool cellAfterIsDeleted = cellAfter != null &&
                                cellAfter.Descendants(W.cellDel).Any();
                            if (e.Name == W.tc &&
                                (cellAfterIsDeleted || e.Descendants(W.cellDel).Any()))
                            {
                                var a = new
                                {
                                    CollectionType = DeletedCellCollectionType.DeletedCell,
                                    Disambiguator = new[] { e }
                                        .Concat(e.SiblingsBeforeSelfReverseDocumentOrder())
                                        .Where(z => z.Name == W.tc &&
                                            !z.Descendants(W.cellDel).Any())
                                        .FirstOrDefault()
                                };
                                return a;
                            }
                            var a2 = new
                            {
                                CollectionType = DeletedCellCollectionType.Other,
                                Disambiguator = e
                            };
                            return a2;
                        });
                    var tr = new XElement(W.tr,
                        groupedCells.Select(g =>
                        {
                            if (g.Key.CollectionType == DeletedCellCollectionType.DeletedCell
                                && g.First().Descendants(W.cellDel).Any())
                                return null;
                            if (g.Key.CollectionType == DeletedCellCollectionType.Other)
                                return (object)g;
                            XElement gridSpanElement = g
                                .First()
                                .Elements(W.tcPr)
                                .Elements(W.gridSpan)
                                .FirstOrDefault();
                            int gridSpan = gridSpanElement != null ?
                                (int)gridSpanElement.Attribute(W.val) :
                                1;
                            int newGridSpan = gridSpan + g.Count() - 1;
                            XElement currentTcPr = g.First().Elements(W.tcPr).FirstOrDefault();
                            XElement newTcPr = new XElement(W.tcPr,
                                currentTcPr != null ? currentTcPr.Attributes() : null,
                                new XElement(W.gridSpan,
                                    new XAttribute(W.val, newGridSpan)),
                                currentTcPr.Elements().Where(e => e.Name != W.gridSpan));
                            var orderedTcPr = new XElement(W.tcPr,
                                newTcPr.Elements().OrderBy(e =>
                                {
                                    if (Order_tcPr.ContainsKey(e.Name))
                                        return Order_tcPr[e.Name];
                                    return 999;
                                }));
                            XElement newTc = new XElement(W.tc,
                                orderedTcPr,
                                g.First().Elements().Where(e => e.Name != W.tcPr));
                            return (object)newTc;
                        }));
                    return tr;
                }

                // Identity clone
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => AcceptDeletedCellsTransform(n)));
            }
            return node;
        }

        private static XName[] BlockLevelElements = new[] {
            W.p,
            W.tbl,
            W.sdt,
            W.del,
            W.ins,
            M.oMath,
            M.oMathPara,
            W.moveTo,
        };

        private static object RemoveRowsLeftEmptyByMoveFrom(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.tr)
                {
                    var nonEmptyCells = element.Elements(W.tc).Any(tc => tc.Elements().Any(tcc => BlockLevelElements.Contains(tcc.Name)));
                    if (nonEmptyCells)
                    {
                        return new XElement(element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n => RemoveRowsLeftEmptyByMoveFrom(n)));
                    }
                    return null;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => RemoveRowsLeftEmptyByMoveFrom(n)));
            }
            return node;
        }

        public static XName[] TrackedRevisionsElements = new[]
        {
            W.cellDel,
            W.cellIns,
            W.cellMerge,
            W.customXmlDelRangeEnd,
            W.customXmlDelRangeStart,
            W.customXmlInsRangeEnd,
            W.customXmlInsRangeStart,
            W.del,
            W.delInstrText,
            W.delText,
            W.ins,
            W.moveFrom,
            W.moveFromRangeEnd,
            W.moveFromRangeStart,
            W.moveTo,
            W.moveToRangeEnd,
            W.moveToRangeStart,
            W.numberingChange,
            W.pPrChange,
            W.rPrChange,
            W.sectPrChange,
            W.tblGridChange,
            W.tblPrChange,
            W.tblPrExChange,
            W.tcPrChange,
            W.trPrChange,
        };

        public static bool PartHasTrackedRevisions(OpenXmlPart part)
        {
            return part.GetXDocument()
                .Descendants()
                .Any(e => TrackedRevisionsElements.Contains(e.Name));
        }


        public static bool HasTrackedRevisions(WordprocessingDocument doc)
        {
            if (PartHasTrackedRevisions(doc.MainDocumentPart))
                return true;
            foreach (var part in doc.MainDocumentPart.HeaderParts)
                if (PartHasTrackedRevisions(part))
                    return true;
            foreach (var part in doc.MainDocumentPart.FooterParts)
                if (PartHasTrackedRevisions(part))
                    return true;
            if (doc.MainDocumentPart.EndnotesPart != null)
                if (PartHasTrackedRevisions(doc.MainDocumentPart.EndnotesPart))
                    return true;
            if (doc.MainDocumentPart.FootnotesPart != null)
                if (PartHasTrackedRevisions(doc.MainDocumentPart.FootnotesPart))
                    return true;
            return false;
        }

        public static void DoAcceptRevisions(string fileName)
        {
            // Given a document name and an author name, accept revisions.
            using (WordprocessingDocument wdDoc =
                WordprocessingDocument.Open(fileName, true))
            {
                AcceptRevisions(wdDoc);
            }

        }


        public static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                Console.WriteLine(args[0]);
                RevisionAccepter.DoAcceptRevisions(args[0]);
            }
        }
    }

    public class BlockContentInfo
    {
        public XElement PreviousBlockContentElement;
        public XElement ThisBlockContentElement;
        public XElement NextBlockContentElement;
    }

    public static class RevisionAccepterExtension
    {
        private static void InitializeParagraphInfo(XElement contentContext)
        {
            if (!(W.BlockLevelContentContainers.Contains(contentContext.Name)))
                throw new ArgumentException(
                    "GetParagraphInfo called for element that is not child of content container");
            XElement prev = null;
            foreach (var content in contentContext.Elements())
            {
                // This may return null, indicating that there is no descendant paragraph.  For
                // example, comment elements have no descendant elements.
                XElement paragraph = content
                    .DescendantsAndSelf()
                    .Where(e => e.Name == W.p || e.Name == W.tc || e.Name == W.txbxContent)
                    .FirstOrDefault();
                if (paragraph != null &&
                    (paragraph.Name == W.tc || paragraph.Name == W.txbxContent))
                    paragraph = null;
                BlockContentInfo pi = new BlockContentInfo()
                {
                    PreviousBlockContentElement = prev,
                    ThisBlockContentElement = paragraph
                };
                content.AddAnnotation(pi);
                prev = content;
            }
        }

        public static BlockContentInfo GetParagraphInfo(this XElement contentElement)
        {
            BlockContentInfo paragraphInfo = contentElement.Annotation<BlockContentInfo>();
            if (paragraphInfo != null)
                return paragraphInfo;
            InitializeParagraphInfo(contentElement.Parent);
            return contentElement.Annotation<BlockContentInfo>();
        }

        public static IEnumerable<XElement> ContentElementsBeforeSelf(this XElement element)
        {
            XElement current = element;
            while (true)
            {
                BlockContentInfo pi = current.GetParagraphInfo();
                if (pi.PreviousBlockContentElement == null)
                    yield break;
                yield return pi.PreviousBlockContentElement;
                current = pi.PreviousBlockContentElement;
            }
        }
    }
}






