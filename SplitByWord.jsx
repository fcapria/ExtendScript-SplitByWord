/**
 * Split selected live TextFrames into POINTTEXT words on a new layer,
 * attempting to preserve original word spacing:
 *  - Authored whitespace (spaces, NBSPs, tabs) widths are measured and used.
 *  - Newlines (\n / paragraph breaks) advance by the source leading (or ~1.2x).
 *  - Soft line wraps from area text are not preserved (Illustrator doesn't expose them).
 *  - Forces NO STROKE on outputs.
 *
 * Layer name: "Split Text"
 */

(function () {
  if (app.documents.length === 0) { alert("No document open."); return; }
  var doc = app.activeDocument;
  if (!doc.selection || doc.selection.length === 0) { alert("Nothing is selected."); return; }

  // ---- config ----
  var ARTBOARD_MARGIN_X = 40;     // px from artboard left
  var ARTBOARD_MARGIN_Y = 80;     // px from artboard top
  var TAB_SPACES = 4;             // how many spaces a tab approximates (for width); used only if measuring "\t" fails
  var ADVANCE_CHAR_FACTOR = 0.55; // fallback width multiplier per char when Illustrator hasn't computed bounds yet
  var GAP_MIN_PX = 0;             // minimal extra gap between tokens (set 0 since we now use authored spaces)
  var SAFETY_MIN_WORD_W = 1;      // px

  // ---- layer helper ----
  function getOrCreateNewLayer() {
    for (var i = 0; i < doc.layers.length; i++) {
      if (doc.layers[i].name === "Split Text") {
        try { doc.layers[i].locked = false; doc.layers[i].visible = true; } catch (e) {}
        return doc.layers[i];
      }
    }
    var lyr = doc.layers.add();
    lyr.name = "Split Text";
    try { lyr.locked = false; lyr.visible = true; } catch (e) {}
    return lyr;
  }
  var outLayer = getOrCreateNewLayer();

  // ---- utils ----
  function isTextFrame(n){ return n && n.typename === "TextFrame" && !n.locked && !n.hidden; }

  // Descend ONLY within selection; do not walk up to layer/doc.
  function collectTextFramesDescendOnly(node, out){
    if (!node || node.hidden || node.locked) return;
    if (isTextFrame(node)) { out.push(node); return; }
    var kids = [node.textFrames, node.groupItems, node.pageItems, node.compoundPathItems, node.symbolItems, node.pathItems];
    for (var k = 0; k < kids.length; k++){
      var coll = kids[k]; if (!coll) continue;
      for (var i = 0; i < coll.length; i++) collectTextFramesDescendOnly(coll[i], out);
    }
  }

  function cloneColor(c){ if (!c) return null; var n;
    switch(c.typename){
      case "RGBColor": n=new RGBColor(); n.red=c.red; n.green=c.green; n.blue=c.blue; return n;
      case "GrayColor": n=new GrayColor(); n.gray=c.gray; return n;
      case "CMYKColor": n=new CMYKColor(); n.cyan=c.cyan; n.magenta=c.magenta; n.yellow=c.yellow; n.black=c.black; return n;
      case "SpotColor": n=new SpotColor(); n.spot=c.spot; n.tint=c.tint; return n;
      default: return null;
    }
  }
  function copyCharStyle(fromRange, toRange){
    var fa=fromRange.characterAttributes, ta=toRange.characterAttributes;
    if (fa.textFont) ta.textFont = fa.textFont;
    if (fa.size != null) ta.size = fa.size;
    var fc=cloneColor(fa.fillColor);   if (fc) ta.fillColor = fc;
    var sc=cloneColor(fa.strokeColor); if (sc) ta.strokeColor = sc;
    if (fa.tracking != null) ta.tracking = fa.tracking;
    if (fa.leading  != null) ta.leading  = fa.leading;
    if (fa.horizontalScale != null) ta.horizontalScale = fa.horizontalScale;
    if (fa.verticalScale   != null) ta.verticalScale   = fa.verticalScale;
    if (fa.baselineShift   != null) ta.baselineShift   = fa.baselineShift;
    if (fa.capitalization  != null) ta.capitalization  = fa.capitalization;
    if (fa.kerning         != null) ta.kerning         = fa.kerning;
    if (fa.strokeWeight    != null) ta.strokeWeight    = fa.strokeWeight;
    if (fa.overprintFill   != null) ta.overprintFill   = fa.overprintFill;
    if (fa.overprintStroke != null) ta.overprintStroke = fa.overprintStroke;
  }

  // Tokenize preserving authored spaces and newlines.
  // Produces array of { kind: "word"|"space"|"nl", text, startIdx, endIdx (exclusive) }
  function tokenizePreserveSpaces(tf){
    var s = String(tf.contents || "");
    // Normalize CRLF to \n, preserve NBSP \u00A0 and tabs \t
    s = s.replace(/\r\n?/g, "\n");

    var tokens = [];
    var i = 0, n = s.length;
    function isSpaceCh(ch){ return ch === " " || ch === "\t" || ch === "\u00A0"; }

    while (i < n) {
      var ch = s[i];

      if (ch === "\n") { tokens.push({kind:"nl", text:"\n", startIdx:i, endIdx:i+1}); i++; continue; }

      // spaces block
      if (isSpaceCh(ch)) {
        var s0 = i;
        while (i<n && isSpaceCh(s[i])) i++;
        tokens.push({kind:"space", text: s.substring(s0,i), startIdx:s0, endIdx:i});
        continue;
      }

      // word block
      var w0 = i;
      while (i<n && s[i] !== "\n" && !isSpaceCh(s[i])) i++;
      tokens.push({kind:"word", text: s.substring(w0,i), startIdx:w0, endIdx:i});
    }
    return tokens;
  }

  // Safely get a character range (first char of a token) for style sampling
  function styleRangeFor(tf, token){
    try {
      var idx = Math.min(token.startIdx, tf.textRange.characters.length-1);
      if (idx < 0) idx = 0;
      return tf.textRange.characters[idx];
    } catch(e){ return tf.textRange.characters[0]; }
  }

  // Measure arbitrary text width for given character style by creating a temporary point text.
  function measureTextWidth(text, styleCharRange){
    var tmp = outLayer.textFrames.add();
    try {
      tmp.kind = TextType.POINTTEXT;
      tmp.position = [ -99999, 99999 ]; // off-canvas
      // set font first to avoid flash
      var ca = styleCharRange.characterAttributes;
      if (ca && ca.textFont) tmp.textRange.characterAttributes.textFont = ca.textFont;
      tmp.contents = text;
      copyCharStyle(styleCharRange, tmp.textRange);

      // FORCE no stroke (consistent with outputs)
      var ta = tmp.textRange.characterAttributes;
      ta.strokeWeight = 0;
      ta.overprintStroke = false;
      ta.strokeColor = new NoColor();

      app.redraw();
      var gb = tmp.geometricBounds; // [top, left, bottom, right]
      var w = gb[3] - gb[1];
      if (!(w > 0)) {
        // Fallback if Illustrator hasn't computed bounds yet
        var size = (ca && ca.size) ? ca.size : 24;
        w = Math.max(SAFETY_MIN_WORD_W, String(text).length * (size * ADVANCE_CHAR_FACTOR));
      }
      return w;
    } catch(e){
      // Emergency fallback
      var caf = styleCharRange.characterAttributes;
      var sz = (caf && caf.size) ? caf.size : 24;
      return Math.max(SAFETY_MIN_WORD_W, String(text).length * (sz * ADVANCE_CHAR_FACTOR));
    } finally {
      try { tmp.remove(); } catch(e){}
    }
  }

  // ---- collect frames from selection ----
  var frames = [];
  for (var s = 0; s < doc.selection.length; s++) collectTextFramesDescendOnly(doc.selection[s], frames);
  if (!frames.length) { alert("No live TextFrames inside the selection."); return; }

  // Visible anchor on active artboard
  var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
  var anchorLeft = ab[0] + ARTBOARD_MARGIN_X;
  var anchorTop  = ab[1] - ARTBOARD_MARGIN_Y;

  // ---- main ----
  var totalWords = 0;

  for (var f = 0; f < frames.length; f++) {
    var src = frames[f];
    if (!src.contents) continue;

    var defCA = src.textRange.characters[0].characterAttributes;
    var defSize = defCA.size || 24;
    var defLeading = (defCA.leading && defCA.leading > 0) ? defCA.leading : defSize * 1.2;

    // tokens preserving spaces
    var tokens = tokenizePreserveSpaces(src);
    if (!tokens || !tokens.length) continue;

    var grp = outLayer.groupItems.add();
    grp.name = "SplitWords_" + (+new Date());

    var x = anchorLeft, y = anchorTop;

    for (var t = 0; t < tokens.length; t++) {
      var tk = tokens[t];

      if (tk.kind === "nl") {
        x = anchorLeft;
        // pick leading based on next visible token's style if available
        var nextIdx = t+1;
        var leadSrc = defCA;
        while (nextIdx < tokens.length && tokens[nextIdx].kind !== "word") nextIdx++;
        if (nextIdx < tokens.length) {
          try { leadSrc = styleRangeFor(src, tokens[nextIdx]).characterAttributes; } catch(e){}
        }
        var leadVal = (leadSrc && leadSrc.leading && leadSrc.leading > 0) ? leadSrc.leading : defLeading;
        y -= leadVal;
        continue;
      }

      if (tk.kind === "space") {
        // measure whitespace width with its own style
        var sRange = styleRangeFor(src, tk);
        var textForMeasure = tk.text;

        // If Illustrator returns 0 width for tabs, approximate as N spaces
        // We'll attempt real measure first; if width <= 0, replace tabs with spaces
        var wSpace = measureTextWidth(textForMeasure, sRange);
        if (!(wSpace > 0) && /\t/.test(textForMeasure)) {
          var approx = textForMeasure.replace(/\t/g, new Array(TAB_SPACES+1).join(" "));
          wSpace = measureTextWidth(approx, sRange);
        }

        x += Math.max(0, wSpace) + GAP_MIN_PX;
        continue;
      }

      if (tk.kind === "word") {
        var wordText = tk.text;
        if (!wordText) continue;

        // style source: first character of this token
        var styleSrcRange = styleRangeFor(src, tk);
        var srcCA = styleSrcRange.characterAttributes;

        // --- create ONE word ---
        var tf = outLayer.textFrames.add();
        tf.kind = TextType.POINTTEXT;
        tf.position = [x, y];

        // set exact font first
        if (srcCA && srcCA.textFont) {
          tf.textRange.characterAttributes.textFont = srcCA.textFont;
        }

        // contents + style
        tf.contents = wordText;
        copyCharStyle(styleSrcRange, tf.textRange);

        // FORCE NO STROKE
        var ta = tf.textRange.characterAttributes;
        ta.strokeWeight = 0;
        ta.overprintStroke = false;
        ta.strokeColor = new NoColor();

        // measure placed word & advance
        app.redraw();
        var gbw = tf.geometricBounds; // [top,left,bottom,right]
        var width = gbw[3] - gbw[1];
        var sizeThis = (srcCA && srcCA.size) ? srcCA.size : defSize;
        if (!(width > 0)) width = Math.max(SAFETY_MIN_WORD_W, wordText.length * (sizeThis * ADVANCE_CHAR_FACTOR));

        x += width + GAP_MIN_PX;

        tf.move(grp, ElementPlacement.PLACEATEND);
        totalWords++;
      }
    }

    try { grp.zOrder(ZOrderMethod.BRINGTOFRONT); } catch(e){}
  }

  alert("Created " + totalWords + " word blocks on layer: Split Text");
})();
