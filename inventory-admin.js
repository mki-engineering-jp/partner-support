/* 在庫状況 → AI秘書 更新（管理者専用）
 * 幹_SAI_フォーキャスト.xlsx をブラウザ内で解析し、bot_knowledge id=107 を更新する。
 * ルールは inventory_to_ai.py と同一:
 *  - 「幹開発在庫数」行の黄色セル(FFFF00)の一番右 = 実在庫、その月が基準月
 *  - 黄色セル無し → 「在庫状況はお問い合わせください」
 *  - 丸め: 100以上は上2桁切り捨て / 1〜99「100台未満」/ 0以下「在庫なし」(マイナス=受注分の未引き当て)
 *  - ラベル固定値: 500以上 余裕あり / 1〜499 残りわずか / 0 在庫なし
 *  - フォーキャスト: 案件行の確度(G列)が「高」のものだけ、基準月から FORECAST_MONTHS か月
 *  - 入荷予定: 「幹開発入荷予定」行の基準月以降
 *  - 社名・案件名・備考は一切出さない
 */
(function (global) {
  "use strict";

  var KNOWLEDGE_ID = 107;
  var FORECAST_MONTHS = 3;
  var TITLE = "幹開発 在庫状況（概算）・入荷予定・フォーキャスト出荷予定 — 月次更新";
  var LEADTIME_OVERRIDE = {
    "MKI-P7": "約2〜3か月",
    "MKI-E4R": "約2〜3か月",
    "MKI-E4L": "約2〜3か月",
    "MKI-P3CH10-2": "約2〜3か月",
    "MKI-P1S1(3.5mm)": "約3〜4か月",
    "MKI-P6NL": "約2〜3か月",
    "SSE-1083H": "約2〜3か月",
    "MKI-P6MG": "約2〜3か月"
  };
  var SKIP_NAMES = { "その他": true };

  // ── セル値の取り出し（数式セル・リッチテキスト対応） ──
  function val(cell) {
    if (!cell) return null;
    var v = cell.value;
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return v;
    if (typeof v === "object") {
      if ("result" in v) return val({ value: v.result });
      if (v.richText) return v.richText.map(function (r) { return r.text; }).join("");
      if (v.text) return v.text;
      return null;
    }
    return v;
  }
  function num(cell) {
    var v = val(cell);
    return (typeof v === "number" && isFinite(v)) ? v : null;
  }
  function str(cell) {
    var v = val(cell);
    return (v === null || v instanceof Date) ? "" : String(v).trim();
  }
  function isYellow(cell) {
    var f = cell && cell.fill;
    if (!f || f.type !== "pattern" || f.pattern !== "solid" || !f.fgColor) return false;
    var argb = f.fgColor.argb || "";
    return argb.toUpperCase().slice(-6) === "FFFF00";
  }
  function monthKey(d) { return d.getUTCFullYear() * 100 + (d.getUTCMonth() + 1); }
  function ymLabel(k) { return Math.floor(k / 100) + "年" + (k % 100) + "月"; }
  function addMonths(k, n) {
    var y = Math.floor(k / 100), m = (k % 100) + n;
    return (y + Math.floor((m - 1) / 12)) * 100 + ((m - 1) % 12 + 1);
  }
  function roundQty(n) {
    if (n === null || n === undefined) return null;
    if (n <= 0) return "在庫なし";
    if (n < 100) return "100台未満";
    var p = Math.pow(10, Math.floor(Math.log10(n)) - 1);
    return "約" + (Math.floor(n / p) * p).toLocaleString("ja-JP") + "台";
  }
  function label(n) {
    if (n === null || n === undefined) return "お問い合わせ";
    if (n >= 500) return "余裕あり";
    if (n > 0) return "残りわずか";
    return "在庫なし";
  }

  // ── 解析 ──
  function parse(workbook) {
    var ws = workbook.getWorksheet("Forecast");
    if (!ws) throw new Error("「Forecast」シートが見つかりません");
    var months = {}; // col -> monthKey
    ws.getRow(4).eachCell(function (c, col) {
      var v = val(c);
      if (v instanceof Date) months[col] = monthKey(v);
    });
    var monthCols = Object.keys(months).map(Number);

    var leadtime = {}, leadtimeByModel = {};
    var lt = workbook.getWorksheet("定数・納期");
    if (lt) {
      lt.eachRow(function (r, rn) {
        if (rn < 5) return;
        var name = str(r.getCell(4)), model = str(r.getCell(5)), days = str(r.getCell(8));
        if (name) leadtime[name] = days;
        if (model) leadtimeByModel[model] = days;
      });
    }

    var blocks = [], cur = null, name = null, model = null;
    ws.eachRow(function (r, rn) {
      if (rn < 5) return;
      var b = str(r.getCell(2)), c = str(r.getCell(3)), d = str(r.getCell(4));
      if (d === "案件") { name = model = null; cur = null; return; }
      if (b) name = b;
      if (c) model = c;
      if ((!name || SKIP_NAMES[name]) && !model) return;
      if (!cur || cur.name !== name || cur.model !== model) {
        cur = { name: name, model: model, stock: null, stockMonth: null, need: {}, arrivals: {}, unrated: {} };
        blocks.push(cur);
      }
      if (d.indexOf("幹開発在庫数") === 0) {
        if (cur.stock === null) {
          var ys = [];
          monthCols.forEach(function (col) { if (isYellow(r.getCell(col))) ys.push([months[col], num(r.getCell(col))]); });
          if (ys.length) {
            var last = ys[ys.length - 1];
            if (last[1] !== null) { cur.stockMonth = last[0]; cur.stock = Math.trunc(last[1]); }
          }
        }
      } else if (d === "幹開発入荷予定") {
        monthCols.forEach(function (col) {
          var v = num(r.getCell(col));
          if (v !== null && v > 0) cur.arrivals[months[col]] = (cur.arrivals[months[col]] || 0) + v;
        });
      } else if (d.indexOf("必要数") === 0 || d.indexOf("必要最") === 0) {
        // 合計行は使わない
      } else {
        var conf = str(r.getCell(7));
        if (conf === "高") {
          monthCols.forEach(function (col) {
            var v = num(r.getCell(col));
            if (v !== null && v > 0) cur.need[months[col]] = (cur.need[months[col]] || 0) + v;
          });
        } else if (conf === "") {
          monthCols.forEach(function (col) {
            var v = num(r.getCell(col));
            if (v !== null && v > 0) cur.unrated[months[col]] = true;
          });
        }
      }
    });

    // 同一製品のブロック統合（機種名のみ／型番のみの行を吸収）
    var nameToKey = {};
    blocks.forEach(function (bl) { if (bl.model && bl.name) nameToKey[bl.name] = bl.model; });
    var merged = {}, order = [];
    blocks.forEach(function (bl) {
      var key = bl.model || nameToKey[bl.name] || bl.name;
      if (!key) return;
      if (!bl.name || SKIP_NAMES[bl.name]) bl.name = key;
      if (merged[key]) {
        var m = merged[key];
        if (!m.model && bl.model) m.model = bl.model;
        if (m.stock === null && bl.stock !== null) { m.stock = bl.stock; m.stockMonth = bl.stockMonth; }
        ["need", "arrivals"].forEach(function (k) {
          Object.keys(bl[k]).forEach(function (mm) { m[k][mm] = (m[k][mm] || 0) + bl[k][mm]; });
        });
        Object.keys(bl.unrated).forEach(function (mm) { m.unrated[mm] = true; });
      } else { merged[key] = bl; order.push(key); }
    });
    var items = order.map(function (k) { return merged[k]; });
    items.forEach(function (bl) {
      var l = LEADTIME_OVERRIDE[bl.name];
      if (l === undefined) {
        var base = bl.name.split("(")[0];
        l = leadtime[bl.name] || leadtime[base] || leadtimeByModel[bl.model || ""] || "";
      }
      bl.leadtime = l ? String(l).replace(/３/g, "3").replace(/～/g, "〜") : "";
    });
    return items;
  }

  function buildText(items, baseMonth) {
    var today = new Date();
    var ds = today.getFullYear() + "/" + String(today.getMonth() + 1).padStart(2, "0") + "/" + String(today.getDate()).padStart(2, "0");
    var L = [];
    L.push("【幹開発 在庫状況（概算）】基準: " + ymLabel(baseMonth) + "時点の幹開発実在庫（" + ds + "更新）");
    L.push("※在庫は日々変動します。数量は概算（丸めた値）です。確定在庫・納期は必ずSlackで幹開発にご確認ください。");
    L.push("※「フォーキャスト出荷予定」は、確度「高」として共有されている出荷見込み分です（確度「低」は含めていません）。この分は在庫から引き当てられる可能性があるため、新規案件の数量が「在庫数－フォーキャスト出荷予定」を超える場合は特に幹開発へ事前にご相談ください。");
    L.push("※納期目安は在庫が無い場合の生産リードタイムで、あくまで目安です。材料が足りない場合はそれ以上かかる恐れがあり、状況によって変わります。詳細は幹開発の担当にご確認ください。");
    L.push("");
    var ord = { "余裕あり": 0, "残りわずか": 1, "在庫なし": 2, "お問い合わせ": 3 };
    items = items.slice().sort(function (a, b) {
      var da = ord[label(a.stock)] - ord[label(b.stock)];
      return da !== 0 ? da : ((b.stock || 0) - (a.stock || 0));
    });
    var limit = addMonths(baseMonth, FORECAST_MONTHS - 1);
    items.forEach(function (bl) {
      var disp = (bl.model && bl.model !== bl.name) ? bl.name + "（" + bl.model + "）" : bl.name;
      if (bl.stock === null) {
        L.push("■ " + disp + "：在庫状況はお問い合わせください（Slackで幹開発へ）");
      } else if (bl.stock < 0) {
        L.push("■ " + disp + "：在庫なし［在庫なし］（受注済み分に対して" + roundQty(-bl.stock) + "が未引き当て＝注文に対して足りていない状態。次回入荷で充当予定。新規のご注文は入荷後の対応となるため、納期は幹開発にご確認ください）");
      } else {
        L.push("■ " + disp + "：" + roundQty(bl.stock) + "［" + label(bl.stock) + "］");
      }
      var needs = Object.keys(bl.need).map(Number).filter(function (m) { return m >= baseMonth && m <= limit; }).sort();
      if (needs.length) L.push("  ・フォーキャスト出荷予定: " + needs.map(function (m) { return ymLabel(m) + " " + roundQty(bl.need[m]); }).join("、"));
      var arr = Object.keys(bl.arrivals).map(Number).filter(function (m) { return m >= baseMonth; }).sort();
      if (arr.length) L.push("  ・入荷予定: " + arr.map(function (m) { return ymLabel(m) + " " + roundQty(bl.arrivals[m]); }).join("、"));
      if (bl.leadtime) L.push("  ・在庫切れ時の納期目安: " + bl.leadtime);
    });
    L.push("");
    L.push("【回答方針】在庫を聞かれたら、上の概算とラベルを答え、必ず「最新の確定在庫はSlackで幹開発にご確認ください」と添える。社名・案件名は一切出さない。フォーキャスト出荷予定がある製品は「うち○台は○月に出荷予定あり」として伝える。黄色セル未記入（お問い合わせ）の製品は数量を推測せず問い合わせを案内する。価格は答えない。");
    return L.join("\n");
  }

  // 入口: ArrayBuffer → { text, warnings, baseMonth }
  function analyze(arrayBuffer) {
    var wb = new ExcelJS.Workbook();
    return wb.xlsx.load(arrayBuffer).then(function () {
      var items = parse(wb);
      var months = items.map(function (b) { return b.stockMonth; }).filter(Boolean);
      if (!months.length) throw new Error("黄色セル（実在庫）が見つかりません。「幹開発在庫数」行の当月セルを黄色く塗ってから再度お試しください。");
      var baseMonth = Math.max.apply(null, months);
      items.forEach(function (b) { if (b.stockMonth && b.stockMonth !== baseMonth) b.stock = null; });
      var warnings = [];
      items.forEach(function (b) {
        var ms = Object.keys(b.unrated).map(Number).filter(function (m) { return m >= baseMonth; }).sort();
        if (ms.length) warnings.push(b.name + "：" + ms.map(ymLabel).join("、") + " に確度が空欄の案件行があります（フォーキャストに含めていません）");
      });
      var noStock = items.filter(function (b) { return b.stock === null; }).map(function (b) { return b.name; });
      if (noStock.length) warnings.push("黄色セル無し（お問い合わせ扱い）: " + noStock.join("、"));
      return { text: buildText(items, baseMonth), warnings: warnings, baseMonth: baseMonth, count: items.length };
    });
  }

  // ── UI ──
  var EXCELJS_URLS = [
    "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js",
    "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"
  ];
  function loadExcelJS() {
    if (global.ExcelJS) return Promise.resolve();
    return EXCELJS_URLS.reduce(function (p, url) {
      return p.catch(function () {
        return new Promise(function (res, rej) {
          var s = document.createElement("script");
          s.src = url; s.onload = res; s.onerror = function () { rej(new Error("load " + url)); };
          document.head.appendChild(s);
        });
      });
    }, Promise.reject()).catch(function () { throw new Error("Excel解析ライブラリの読み込みに失敗しました"); });
  }

  function mount(sb, container) {
    container.innerHTML =
      '<div class="inv-admin-head">在庫状況の更新（管理者）</div>' +
      '<div class="inv-admin-desc">SharePointからダウンロードした「幹_SAI_フォーキャスト.xlsx」を選ぶと、AI秘書の在庫知識を更新できます。' +
      '黄色く塗った「幹開発在庫数」セルが実在庫として読み取られます。<span id="inv-current"></span></div>' +
      '<div class="inv-drop" id="inv-drop">ここにxlsxをドロップ、またはクリックして選択' +
      '<input type="file" id="inv-file" accept=".xlsx" style="display:none"></div>' +
      '<div id="inv-warn" class="inv-warn" style="display:none"></div>' +
      '<textarea id="inv-preview" class="inv-preview" style="display:none" spellcheck="false"></textarea>' +
      '<div id="inv-actions" style="display:none">' +
      '<button type="button" id="inv-save" class="inv-save">この内容でAI秘書を更新</button>' +
      '<button type="button" id="inv-cancel" class="inv-cancel">取り消し</button></div>' +
      '<div id="inv-msg" class="inv-msg"></div>';

    var drop = container.querySelector("#inv-drop");
    var fileIn = container.querySelector("#inv-file");
    var warn = container.querySelector("#inv-warn");
    var preview = container.querySelector("#inv-preview");
    var actions = container.querySelector("#inv-actions");
    var msg = container.querySelector("#inv-msg");
    var currentEl = container.querySelector("#inv-current");

    function setMsg(t, cls) { msg.textContent = t || ""; msg.className = "inv-msg" + (cls ? " " + cls : ""); }
    function reset() { preview.style.display = "none"; actions.style.display = "none"; warn.style.display = "none"; preview.value = ""; }

    sb.from("bot_knowledge").select("updated_at, content").eq("id", KNOWLEDGE_ID).maybeSingle().then(function (res) {
      if (res.data && res.data.updated_at) {
        var d = new Date(res.data.updated_at);
        var head = (res.data.content || "").split("\n")[0].replace(/^【[^】]*】/, "");
        currentEl.textContent = " 現在の登録: " + head + "（更新 " + d.toLocaleString("ja-JP") + "）";
      }
    });

    function handleFile(file) {
      if (!file) return;
      if (!/\.xlsx$/i.test(file.name)) { setMsg("xlsxファイルを選んでください", "err"); return; }
      reset();
      setMsg("解析中…");
      loadExcelJS().then(function () { return file.arrayBuffer(); })
        .then(analyze)
        .then(function (r) {
          preview.value = r.text;
          preview.style.display = "block";
          actions.style.display = "flex";
          if (r.warnings.length) { warn.innerHTML = r.warnings.map(function (w) { return "⚠ " + w; }).join("<br>"); warn.style.display = "block"; }
          setMsg(ymLabel(r.baseMonth) + "基準・" + r.count + "製品を読み取りました。内容を確認して「更新」を押してください（必要なら直接編集できます）。", "ok");
        })
        .catch(function (e) { setMsg("読み取りに失敗しました: " + (e && e.message ? e.message : e), "err"); });
    }

    drop.addEventListener("click", function () { fileIn.click(); });
    fileIn.addEventListener("change", function () { handleFile(fileIn.files[0]); fileIn.value = ""; });
    ["dragenter", "dragover"].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add("over"); }); });
    ["dragleave", "drop"].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove("over"); }); });
    drop.addEventListener("drop", function (e) { handleFile(e.dataTransfer.files && e.dataTransfer.files[0]); });

    container.querySelector("#inv-cancel").addEventListener("click", function () { reset(); setMsg(""); });
    container.querySelector("#inv-save").addEventListener("click", function () {
      var text = preview.value.trim();
      if (!text) return;
      var btn = container.querySelector("#inv-save");
      btn.disabled = true; setMsg("更新中…");
      sb.from("bot_knowledge")
        .update({ content: text, title: TITLE, category: "在庫状況", audience: "science_arts", updated_at: new Date().toISOString() })
        .eq("id", KNOWLEDGE_ID).select("id")
        .then(function (res) {
          btn.disabled = false;
          if (res.error) { setMsg("更新に失敗しました: " + res.error.message, "err"); return; }
          if (!res.data || !res.data.length) { setMsg("更新できませんでした（管理者権限がない可能性があります）", "err"); return; }
          reset();
          currentEl.textContent = " 現在の登録: " + text.split("\n")[0].replace(/^【[^】]*】/, "") + "（更新 " + new Date().toLocaleString("ja-JP") + "）";
          setMsg("AI秘書の在庫知識を更新しました。チャットで「P3の在庫は？」などと聞いて確認してください。", "ok");
        });
    });
  }

  global.InventoryAdmin = { mount: mount, analyze: analyze, parse: parse, buildText: buildText };
})(typeof window !== "undefined" ? window : globalThis);
