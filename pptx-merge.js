/* Merge several .pptx files into one, entirely in the browser.
   Each extra deck keeps its own master/layouts/theme, so formatting is preserved.
   Speaker notes and comments from extra decks are dropped. */
(function (root) {
  function dirname(p) { var i = p.lastIndexOf('/'); return i < 0 ? '' : p.slice(0, i); }
  function basename(p) { return p.slice(p.lastIndexOf('/') + 1); }
  function safeDecode(s) { try { return decodeURIComponent(s); } catch (e) { return s; } }
  function xmlDecode(s) {
    return s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  }
  function resolve(dir, target) {
    target = xmlDecode(target);
    if (target.charAt(0) === '/') return safeDecode(target.slice(1));
    var parts = (dir ? dir.split('/') : []).concat(safeDecode(target).split('/'));
    var out = [];
    parts.forEach(function (s) { if (s === '..') out.pop(); else if (s !== '.' && s !== '') out.push(s); });
    return out.join('/');
  }
  function relpath(fromDir, to) {
    var a = fromDir ? fromDir.split('/') : [], b = to.split('/'), i = 0;
    while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
    var up = a.slice(i).map(function () { return '..'; });
    return up.concat(b.slice(i)).join('/');
  }
  function relsPathFor(part) {
    var d = dirname(part);
    return (d ? d + '/' : '') + '_rels/' + basename(part) + '.rels';
  }
  function parseRels(xml) {
    var out = [], re = /<Relationship\b[^>]*?\/>/g, m;
    while ((m = re.exec(xml))) {
      var attrs = {}, are = /([\w:]+)="([^"]*)"/g, a;
      while ((a = are.exec(m[0]))) attrs[a[1]] = a[2];
      out.push({ tag: m[0], attrs: attrs });
    }
    return out;
  }
  function parseCT(xml) {
    var o = {}, d = {}, m, re = /<Override\b[^>]*?\/>/g;
    while ((m = re.exec(xml))) {
      var p = /PartName="([^"]*)"/.exec(m[0]), c = /ContentType="([^"]*)"/.exec(m[0]);
      if (p && c) o[p[1].replace(/^\//, '')] = c[1];
    }
    re = /<Default\b[^>]*?\/>/g;
    while ((m = re.exec(xml))) {
      var e = /Extension="([^"]*)"/.exec(m[0]), c2 = /ContentType="([^"]*)"/.exec(m[0]);
      if (e && c2) d[e[1].toLowerCase()] = c2[1];
    }
    return { overrides: o, defaults: d };
  }
  function extOf(p) { var b = basename(p), i = b.lastIndexOf('.'); return i < 0 ? '' : b.slice(i + 1).toLowerCase(); }
  function newName(part, k) {
    var d = dirname(part), b = basename(part), i = b.lastIndexOf('.');
    var nb = i < 0 ? b + '_d' + k : b.slice(0, i) + '_d' + k + b.slice(i);
    return (d ? d + '/' : '') + nb;
  }
  var SKIP_REL = /\/(notesSlide|notesMaster|handoutMaster|comments|commentAuthors|presProps|viewProps|tableStyles|printerSettings|presentation)$/;

  async function text(zip, path) {
    var f = zip.file(path);
    return f ? f.async('string') : null;
  }

  async function mergePptx(JSZip, buffers) {
    if (buffers.length < 2) throw new Error('Pick at least two PowerPoint files.');
    var warnings = [];
    var base = await JSZip.loadAsync(buffers[0]);
    var pres = await text(base, 'ppt/presentation.xml');
    var presRels = await text(base, 'ppt/_rels/presentation.xml.rels');
    var ct = await text(base, '[Content_Types].xml');
    if (!pres || !presRels || !ct) throw new Error('First file is not a valid .pptx.');
    var baseSize = /<p:sldSz[^>]*?cx="(\d+)"[^>]*?cy="(\d+)"/.exec(pres);

    // sections would not list the new slides; drop the section list so PowerPoint does not complain
    pres = pres.replace(/<p:ext\b[^>]*uri="\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}"[^>]*>[\s\S]*?<\/p:ext>/g, '');

    var maxId = 2147483647;
    (pres.match(/<p:sldMasterId\b[^>]*\bid="(\d+)"/g) || []).forEach(function (s) {
      maxId = Math.max(maxId, +/id="(\d+)"/.exec(s)[1]);
    });
    var baseMasters = Object.keys(base.files).filter(function (n) { return /^ppt\/slideMasters\/[^/]+\.xml$/.test(n); });
    for (var bm = 0; bm < baseMasters.length; bm++) {
      var mx = await text(base, baseMasters[bm]);
      (mx.match(/<p:sldLayoutId\b[^>]*\bid="(\d+)"/g) || []).forEach(function (s) {
        maxId = Math.max(maxId, +/id="(\d+)"/.exec(s)[1]);
      });
    }
    var nextLayoutOrMasterId = maxId + 1;
    var slideId = 255;
    (pres.match(/<p:sldId\b[^>]*\bid="(\d+)"/g) || []).forEach(function (s) {
      slideId = Math.max(slideId, +/id="(\d+)"/.exec(s)[1]);
    });
    var rid = 0;
    (presRels.match(/\bId="rId(\d+)"/g) || []).forEach(function (s) { rid = Math.max(rid, +/(\d+)/.exec(s)[1]); });

    var newMasterEntries = '', newSlideEntries = '', newRelEntries = '', newCT = '';
    var baseCT = parseCT(ct);

    for (var k = 1; k < buffers.length; k++) {
      var src = await JSZip.loadAsync(buffers[k]);
      var sPres = await text(src, 'ppt/presentation.xml');
      var sPresRels = await text(src, 'ppt/_rels/presentation.xml.rels');
      var sCT = await text(src, '[Content_Types].xml');
      if (!sPres || !sPresRels || !sCT) throw new Error('File ' + (k + 1) + ' is not a valid .pptx.');
      var size = /<p:sldSz[^>]*?cx="(\d+)"[^>]*?cy="(\d+)"/.exec(sPres);
      if (baseSize && size && (size[1] !== baseSize[1] || size[2] !== baseSize[2])) {
        warnings.push('File ' + (k + 1) + ' has a different slide size from file 1; its slides may look stretched.');
      }
      var srcCT = parseCT(sCT);
      var relById = {};
      parseRels(sPresRels).forEach(function (r) { relById[r.attrs.Id] = r; });
      var slides = [];
      (sPres.match(/<p:sldId\b[^>]*?\/>/g) || []).forEach(function (s) {
        var m = /r:id="([^"]*)"/.exec(s);
        if (m && relById[m[1]]) slides.push(resolve('ppt', relById[m[1]].attrs.Target));
      });

      // find every part the slides need
      var seen = {}, order = [], relsOf = {}, queue = slides.slice();
      queue.forEach(function (s) { seen[s] = true; });
      while (queue.length) {
        var part = queue.shift();
        if (!src.file(part)) continue;
        order.push(part);
        var rx = await text(src, relsPathFor(part));
        var kept = [];
        if (rx) {
          parseRels(rx).forEach(function (r) {
            if (/External/i.test(r.attrs.TargetMode || '')) { kept.push({ r: r, ext: true }); return; }
            if (SKIP_REL.test(r.attrs.Type || '')) return;
            var t = resolve(dirname(part), r.attrs.Target);
            if (!src.file(t)) return;
            kept.push({ r: r, target: t });
            if (!seen[t]) { seen[t] = true; queue.push(t); }
          });
        }
        relsOf[part] = kept;
      }
      var map = {};
      order.forEach(function (p) { map[p] = newName(p, k + 1); });

      for (var i = 0; i < order.length; i++) {
        var p = order[i], np = map[p];
        if (/^ppt\/slideMasters\/[^/]+\.xml$/.test(p)) {
          var mxml = await src.file(p).async('string');
          mxml = mxml.replace(/(<p:sldLayoutId\b[^>]*?\bid=")(\d+)(")/g, function (_, a, b, c) { return a + (nextLayoutOrMasterId++) + c; });
          base.file(np, mxml);
        } else {
          base.file(np, await src.file(p).async('uint8array'));
        }
        // content type
        var ext = extOf(p);
        if (srcCT.overrides[p]) {
          newCT += '<Override PartName="/' + np + '" ContentType="' + srcCT.overrides[p] + '"/>';
        } else if (ext && !baseCT.defaults[ext] && srcCT.defaults[ext]) {
          baseCT.defaults[ext] = srcCT.defaults[ext];
          newCT += '<Default Extension="' + ext + '" ContentType="' + srcCT.defaults[ext] + '"/>';
        }
        // rels
        var kept2 = relsOf[p] || [];
        if (kept2.length) {
          var out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
          kept2.forEach(function (e) {
            if (e.ext) { out += e.r.tag; return; }
            var nt = encodeURI(relpath(dirname(np), map[e.target]));
            out += e.r.tag.replace(/Target="[^"]*"/, 'Target="' + nt + '"');
          });
          base.file(relsPathFor(np), out + '</Relationships>');
        }
        // register masters
        if (/^ppt\/slideMasters\/[^/]+\.xml$/.test(p)) {
          rid++;
          newMasterEntries += '<p:sldMasterId id="' + (nextLayoutOrMasterId++) + '" r:id="rId' + rid + '"/>';
          newRelEntries += '<Relationship Id="rId' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="' + encodeURI(relpath('ppt', np)) + '"/>';
        }
      }
      slides.forEach(function (s) {
        if (!map[s]) return;
        rid++;
        newSlideEntries += '<p:sldId id="' + (++slideId) + '" r:id="rId' + rid + '"/>';
        newRelEntries += '<Relationship Id="rId' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="' + encodeURI(relpath('ppt', map[s])) + '"/>';
      });
    }

    // masters must be registered before slides use them
    if (!/<\/p:sldMasterIdLst>/.test(pres)) throw new Error('First file has no slide master list.');
    pres = pres.replace('</p:sldMasterIdLst>', newMasterEntries + '</p:sldMasterIdLst>');
    if (/<\/p:sldIdLst>/.test(pres)) pres = pres.replace('</p:sldIdLst>', newSlideEntries + '</p:sldIdLst>');
    else if (/<p:sldIdLst\s*\/>/.test(pres)) pres = pres.replace(/<p:sldIdLst\s*\/>/, '<p:sldIdLst>' + newSlideEntries + '</p:sldIdLst>');
    else pres = pres.replace('<p:sldSz', '<p:sldIdLst>' + newSlideEntries + '</p:sldIdLst><p:sldSz');
    presRels = presRels.replace('</Relationships>', newRelEntries + '</Relationships>');
    ct = ct.replace('</Types>', newCT + '</Types>');

    base.file('ppt/presentation.xml', pres);
    base.file('ppt/_rels/presentation.xml.rels', presRels);
    base.file('[Content_Types].xml', ct);
    var data = await base.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return { data: data, warnings: warnings };
  }

  if (typeof module !== 'undefined' && module.exports) module.exports = mergePptx;
  else root.mergePptx = mergePptx;
})(typeof self !== 'undefined' ? self : this);
