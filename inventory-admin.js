/* 在庫状況 → AI秘書（管理者専用）  v20261026
 * 2026-10-26 から、在庫数は幹開発の社内在庫システム（inv_stock）を元に
 * Supabase 側の関数 refresh_sa_inventory_knowledge() が毎朝 6:00(JST) に
 * bot_knowledge id=107 を自動生成する方式に変更。xlsx（黄色セル）取込は廃止。
 *
 * この画面でできること:
 *  - 現在の登録内容（最終更新・製品数）の確認とプレビュー
 *  - 「今すぐ反映」: 在庫サイトの最新値で id107 を即時再生成（RPC sa_inventory_refresh_now）
 *  - 「補足メモ」: 入荷予定など一言添えたいときに書く（sa_inventory_note）。空なら何も出ない。
 *
 * 表示ルール（SQL側）: 100以上は上2桁切り捨て「約1,200台」/ 1〜99「100台未満」/ 0以下「在庫なし」
 *  ラベル: 500以上 余裕あり / 1〜499 残りわずか / 0 在庫なし
 *  対象製品は sa_inventory_map に登録した型番のみ（VEXA・デモ機・派生型番は出ない）
 */
(function (global) {
  "use strict";

  var KNOWLEDGE_ID = 107;

  function fmt(d) { return d ? new Date(d).toLocaleString("ja-JP") : "—"; }
  function countItems(text) { return (String(text || "").match(/^■ /gm) || []).length; }

  function mount(sb, container) {
    container.innerHTML =
      '<div class="inv-admin-head">在庫状況（管理者）— 社内在庫システムから自動反映</div>' +
      '<div class="inv-admin-desc">AI秘書の在庫知識は、幹開発の社内在庫サイトの数値から<b>毎朝6時に自動更新</b>されます（概算・丸めた値のみ）。' +
      'xlsxの取込は不要になりました。<span id="inv-current"></span></div>' +
      '<div id="inv-actions" style="display:flex">' +
      '<button type="button" id="inv-refresh" class="inv-save">今すぐ反映（在庫サイトの最新値で更新）</button>' +
      '<button type="button" id="inv-toggle" class="inv-cancel">登録内容を見る</button></div>' +
      '<textarea id="inv-preview" class="inv-preview" style="display:none" readonly spellcheck="false"></textarea>' +
      '<div class="inv-admin-desc" style="margin-top:10px;margin-bottom:4px"><b>補足メモ</b>（入荷予定など、在庫一覧に添えてSAに伝えたいことがあれば。空欄なら表示されません。社名・案件名は書かないでください）</div>' +
      '<textarea id="inv-note" class="inv-preview" style="height:90px" placeholder="例）MKI-P3 は 11月中旬に約3,000台の入荷予定です。" spellcheck="false"></textarea>' +
      '<div style="display:flex;gap:8px;margin-top:8px">' +
      '<button type="button" id="inv-note-save" class="inv-save">メモを保存して反映</button>' +
      '<button type="button" id="inv-note-clear" class="inv-cancel">メモを消す</button></div>' +
      '<div id="inv-msg" class="inv-msg"></div>';

    var currentEl = container.querySelector("#inv-current");
    var preview   = container.querySelector("#inv-preview");
    var toggle    = container.querySelector("#inv-toggle");
    var refresh   = container.querySelector("#inv-refresh");
    var note      = container.querySelector("#inv-note");
    var noteSave  = container.querySelector("#inv-note-save");
    var noteClear = container.querySelector("#inv-note-clear");
    var msg       = container.querySelector("#inv-msg");
    var noteOrig  = "";

    function setMsg(t, cls) { msg.textContent = t || ""; msg.className = "inv-msg" + (cls ? " " + cls : ""); }
    function busy(b) { refresh.disabled = noteSave.disabled = noteClear.disabled = !!b; }

    function showCurrent(text, updatedAt) {
      currentEl.textContent = " 現在の登録: " + countItems(text) + "製品（最終更新 " + fmt(updatedAt) + "）";
      preview.value = text || "";
    }

    function loadCurrent() {
      return sb.from("bot_knowledge").select("updated_at, content").eq("id", KNOWLEDGE_ID).maybeSingle().then(function (res) {
        if (res.error) { setMsg("現在の登録内容を読めませんでした: " + res.error.message, "err"); return; }
        if (res.data) showCurrent(res.data.content, res.data.updated_at);
      });
    }
    function loadNote() {
      return sb.from("sa_inventory_note").select("note, updated_at").eq("id", 1).maybeSingle().then(function (res) {
        if (res.error || !res.data) return;
        noteOrig = res.data.note || "";
        note.value = noteOrig;
      });
    }

    // 在庫サイトの最新値で id107 を再生成
    function doRefresh(okMsg) {
      busy(true); setMsg("在庫サイトの最新値で更新中…");
      return sb.rpc("sa_inventory_refresh_now").then(function (res) {
        busy(false);
        if (res.error) { setMsg("更新に失敗しました: " + res.error.message, "err"); return false; }
        showCurrent(res.data, new Date().toISOString());
        setMsg(okMsg || "AI秘書の在庫知識を更新しました。チャットで「P3の在庫は？」などと聞いて確認してください。", "ok");
        return true;
      });
    }

    toggle.addEventListener("click", function () {
      var open = preview.style.display !== "none";
      preview.style.display = open ? "none" : "block";
      toggle.textContent = open ? "登録内容を見る" : "閉じる";
    });
    refresh.addEventListener("click", function () { doRefresh(); });

    function saveNote(text) {
      busy(true); setMsg("メモを保存中…");
      return sb.from("sa_inventory_note")
        .update({ note: text, updated_at: new Date().toISOString() })
        .eq("id", 1).select("id")
        .then(function (res) {
          if (res.error) { busy(false); setMsg("メモの保存に失敗しました: " + res.error.message, "err"); return; }
          if (!res.data || !res.data.length) { busy(false); setMsg("保存できませんでした（管理者権限がない可能性があります）", "err"); return; }
          noteOrig = text;
          return doRefresh(text ? "補足メモを保存し、AI秘書の在庫知識に反映しました。" : "補足メモを消して、AI秘書の在庫知識を更新しました。");
        });
    }
    noteSave.addEventListener("click", function () { saveNote(note.value.trim()); });
    noteClear.addEventListener("click", function () { note.value = ""; saveNote(""); });

    loadCurrent();
    loadNote();
  }

  global.InventoryAdmin = { mount: mount };
})(typeof window !== "undefined" ? window : globalThis);
