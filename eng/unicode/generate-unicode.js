/*
const { parseUnicodeText } = require('./unicode/parser');
const { extractCategories } = require('./unicode/categories');
const { codepointSequenceToRanges, rangesToPrettyRangesDump, rangesToTextBitmapDump, dumpUnicodeCategories } = require('./unicode/util');
const { readFileUtf8, writeFileUtf8, readFileJson, writeFileJsonPretty, writeFileYamlPretty, mkdir } = require('./lib/util/fs');
const { pathJoin } = require('./lib/util/fs');
const { createConversionMaps } = require('./unicode/case_conversion');
const { jsonDeepClone } = require('../util/clone');
*/
    
const { generateUnicodeFiles } = require('./lib/configure/configure_sources');

generateUnicodeFiles("./UnicodeData.txt", "./SpecialCasing.txt", "./temp")


// no longer needed
// Parse Unicode data and generate useful intermediate outputs.
function xxgenerateUnicodeFiles(unicodeDataFile, specialCasingFile, tempDirectory, srcGenDirectory) {
  // Parse UnicodeData.txt and SpecialCasing.txt into a master codepoint map.
  var cpMap = parseUnicodeText(readFileUtf8(unicodeDataFile), readFileUtf8(specialCasingFile));
  writeFileJsonPretty(pathJoin(tempDirectory, 'codepoint_map.json'), cpMap);

  // Unicode categories.
  var cats = extractCategories(cpMap);
  writeFileJsonPretty(pathJoin(tempDirectory, 'unicode_categories.json'), cats);
  writeFileUtf8(pathJoin(tempDirectory, 'unicode_category_dump.txt'), dumpUnicodeCategories(cpMap, cats));

  // Case conversion maps.
  var convMaps = createConversionMaps(cpMap);
  writeFileJsonPretty(pathJoin(tempDirectory, 'conversion_maps.json'), convMaps);
  var convUcMap = jsonDeepClone(convMaps.uc);
  removeConversionMapAscii(convUcMap);
  var convLcMap = jsonDeepClone(convMaps.lc);
  removeConversionMapAscii(convLcMap);
  var { data: convUcNoa } = generateCaseconvTables(convUcMap);
  var { data: convLcNoa } = generateCaseconvTables(convLcMap);
return;
  // RegExp canonicalization tables.
  var reCanonTab = generateReCanonDirectLookup(convMaps.uc);
  writeFileJsonPretty(pathJoin(tempDirectory, 're_canon_tab.json'), reCanonTab);
  var reCanonBitmap = generateReCanonBitmap(reCanonTab);
  writeFileJsonPretty(pathJoin(tempDirectory, 're_canon_bitmap.json'), reCanonBitmap);

  //var dontcare = require('./lib/unicode/regexp_canon').generateReCanonDontCare(canontab);
  //console.log(dontcare);
  //var ranges = require('./lib/unicode/regexp_canon').generateReCanonRanges(canontab);
  //console.log(ranges);
  //var needcheck = require('./lib/unicode/regexp_canon').generateReCanonNeedCheck(canontab);
  //console.log(needcheck);

  // Category helpers for matchers.
  var catsWs = ['Zs'];
  var catsLetter = ['Lu', 'Ll', 'Lt', 'Lm', 'Lo'];
  var catsIdStart = ['Lu', 'Ll', 'Lt', 'Lm', 'Lo', 'Nl', 0x0024, 0x005f];
  var catsIdPart = ['Lu', 'Ll', 'Lt', 'Lm', 'Lo', 'Nl', 0x0024, 0x005f, 'Mn', 'Mc', 'Nd', 'Pc', 0x200c, 0x200d];

  // Matchers for various codepoint sets.
  var matchWs = extractChars(cpMap, catsWs, []);
  //var matchLetter = extractChars(cpMap, catsLetter, []);
  //var matchLetterNoa = extractChars(cpMap, catsLetter, [ 'ASCII' ]);
  //var matchLetterNoabmp = extractChars(cpMap, catsLetter, [ 'ASCII', 'NONBMP' ]);
  //var matchIdStart = extractChars(cpMap, catsIdStart, []);
  var matchIdStartNoa = extractChars(cpMap, catsIdStart, ['ASCII']);
  var matchIdStartNoabmp = extractChars(cpMap, catsIdStart, ['ASCII', 'NONBMP']);
  //var matchIdStartMinusLetter = extractChars(cpMap, catsIdStart, catsLetter);
  var matchIdStartMinusLetterNoa = extractChars(cpMap, catsIdStart, catsLetter.concat(['ASCII']));
  var matchIdStartMinusLetterNoabmp = extractChars(cpMap, catsIdStart, catsLetter.concat(['ASCII', 'NONBMP']));
  //var matchIdPartMinusIdStart = extractChars(cpMap, catsIdPart, catsIdStart);
  var matchIdPartMinusIdStartNoa = extractChars(cpMap, catsIdPart, catsIdStart.concat(['ASCII']));
  var matchIdPartMinusIdStartNoabmp = extractChars(cpMap, catsIdPart, catsIdStart.concat(['ASCII', 'NONBMP']));

  // Generate C/H files.

  function emitReCanonLookup() {
      var genc;

      genc = new GenerateC();
      genc.emitArray(reCanonTab, {
          tableName: 'duk_unicode_re_canon_lookup',
          typeName: 'duk_uint16_t',
          useConst: true,
          useCast: false,
          visibility: 'DUK_INTERNAL'
      });
      writeFileUtf8(pathJoin(srcGenDirectory, 'duk_unicode_re_canon_lookup.c'), genc.getString());

      genc = new GenerateC();
      genc.emitLine('#if !defined(DUK_SINGLE_FILE)');
      genc.emitLine('DUK_INTERNAL_DECL const duk_uint16_t duk_unicode_re_canon_lookup[' + reCanonTab.length + '];');
      genc.emitLine('#endif');
      writeFileUtf8(pathJoin(srcGenDirectory, 'duk_unicode_re_canon_lookup.h'), genc.getString());
  }

  function emitReCanonBitmap() {
      var genc;

      genc = new GenerateC();
      genc.emitArray(reCanonBitmap.bitmapContinuity, {
          tableName: 'duk_unicode_re_canon_bitmap',
          typeName: 'duk_uint8_t',
          useConst: true,
          useCast: false,
          visibility: 'DUK_INTERNAL'
      });
      writeFileUtf8(pathJoin(srcGenDirectory, 'duk_unicode_re_canon_bitmap.c'), genc.getString());

      genc = new GenerateC();
      genc.emitDefine('DUK_CANON_BITMAP_BLKSIZE', reCanonBitmap.blockSize);
      genc.emitDefine('DUK_CANON_BITMAP_BLKSHIFT', reCanonBitmap.blockShift);
      genc.emitDefine('DUK_CANON_BITMAP_BLKMASK', reCanonBitmap.blockMask);
      genc.emitLine('#if !defined(DUK_SINGLE_FILE)');
      genc.emitLine('DUK_INTERNAL_DECL const duk_uint8_t duk_unicode_re_canon_bitmap[' + reCanonBitmap.bitmapContinuity.length + '];');
      genc.emitLine('#endif');
      writeFileUtf8(pathJoin(srcGenDirectory, 'duk_unicode_re_canon_bitmap.h'), genc.getString());
  }

  function emitMatchTable(arg, tableName) {
      var genc;
      var data = arg.data;
      var ranges = arg.ranges;
      var filename = tableName;

      console.debug(tableName, data.length);

      genc = new GenerateC();
      genc.emitArray(data, {
          tableName: tableName,
          typeName: 'duk_uint8_t',
          useConst: true,
          useCast: false,
          visibility: 'DUK_INTERNAL'
      });
      writeFileUtf8(pathJoin(srcGenDirectory, filename + '.c'), genc.getString());

      genc = new GenerateC();
      genc.emitLine('#if !defined(DUK_SINGLE_FILE)');
      genc.emitLine('DUK_INTERNAL_DECL const duk_uint8_t ' + tableName + '[' + data.length + '];');
      genc.emitLine('#endif');
      writeFileUtf8(pathJoin(srcGenDirectory, filename + '.h'), genc.getString());

      writeFileUtf8(pathJoin(tempDirectory, filename + '_ranges.txt'), rangesToPrettyRangesDump(ranges));
      writeFileUtf8(pathJoin(tempDirectory, filename + '_bitmap.txt'), rangesToTextBitmapDump(ranges));
  }

  function emitCaseconvTables(ucData, lcData) {
      var genc;

      genc = new GenerateC();
      genc.emitArray(ucData, {
          tableName: 'duk_unicode_caseconv_uc',
          typeName: 'duk_uint8_t',
          useConst: true,
          useCast: false,
          visibility: 'DUK_INTERNAL'
      });
      genc.emitArray(lcData, {
          tableName: 'duk_unicode_caseconv_lc',
          typeName: 'duk_uint8_t',
          useConst: true,
          useCast: false,
          visibility: 'DUK_INTERNAL'
      });
      writeFileUtf8(pathJoin(srcGenDirectory, 'duk_unicode_caseconv.c'), genc.getString());

      genc = new GenerateC();
      genc.emitLine('#if !defined(DUK_SINGLE_FILE)');
      genc.emitLine('DUK_INTERNAL_DECL const duk_uint8_t duk_unicode_caseconv_uc[' + ucData.length + '];');
      genc.emitLine('DUK_INTERNAL_DECL const duk_uint8_t duk_unicode_caseconv_lc[' + lcData.length + '];');
      genc.emitLine('#endif');
      writeFileUtf8(pathJoin(srcGenDirectory, 'duk_unicode_caseconv.h'), genc.getString());
  }

  emitCaseconvTables(convUcNoa, convLcNoa);
  emitMatchTable(matchWs, 'duk_unicode_ws'); // not used runtime, but dump is useful
  emitMatchTable(matchIdStartNoa, 'duk_unicode_ids_noa');
  emitMatchTable(matchIdStartNoabmp, 'duk_unicode_ids_noabmp');
  emitMatchTable(matchIdStartMinusLetterNoa, 'duk_unicode_ids_m_let_noa');
  emitMatchTable(matchIdStartMinusLetterNoabmp, 'duk_unicode_ids_m_let_noabmp');
  emitMatchTable(matchIdPartMinusIdStartNoa, 'duk_unicode_idp_m_ids_noa');
  emitMatchTable(matchIdPartMinusIdStartNoabmp, 'duk_unicode_idp_m_ids_noabmp');
  emitReCanonLookup();
  emitReCanonBitmap();
}
