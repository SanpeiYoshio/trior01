/* global Office, document */
/*
 * 送信前 宛先確認(新しい Outlook / Web 版 Outlook 用アドイン)
 *
 *  - 送信時イベント(OnMessageSend)で、社内ドメイン以外の宛先があれば送信を止め、
 *    「宛先を確認する」ボタンからタスクペインを開く。
 *  - タスクペインで宛先にチェックを付け、「確定」を押すと、チェックなしの宛先を
 *    メールから外し、確認済みの印を保存する。
 *  - もう一度「送信」を押すと、印が一致していれば送信される。
 *
 * このファイルは次の3か所から読み込まれる(同じコードを共有)
 *   ・イベントハンドラー用 JS ランタイム(manifest の JSRuntime.Url)
 *   ・commands.html(ブラウザーランタイム)
 *   ・taskpane.html(タスクペイン)
 *
 * クラシック Outlook for Windows の JS ランタイムでも動くよう、
 * async/await やアロー関数は使わず、ES5 の書き方にしている。
 */
(function () {
  'use strict';

  // ===== 設定 ==========================================================
  var CONFIG = {
    // 確認対象外にするドメイン(先頭の @ は省略可)
    INTERNAL_DOMAINS: ['@trior.jp'],
    // 確認済みの印を保存するキー
    STATE_KEY: 'recipientCheckConfirmed',
    // true にすると「確定」後に自動で送信を試みる(item.sendAsync が使える環境のみ)
    AUTO_SEND_AFTER_CONFIRM: false,
    // manifest.xml のタスクペインボタンの ID(変更しない)
    PANE_BUTTON_ID: 'msgComposeOpenPaneButton'
  };
  // ====================================================================

  var TYPES = ['to', 'cc', 'bcc'];
  var TYPE_LABEL = { to: 'To', cc: 'Cc', bcc: 'Bcc' };

  // ---- 基本関数 ------------------------------------------------------

  function norm(s) {
    return String(s || '').replace(/^\s+|\s+$/g, '').toLowerCase();
  }

  function isInternal(address) {
    var a = norm(address);
    for (var i = 0; i < CONFIG.INTERNAL_DOMAINS.length; i++) {
      var d = norm(CONFIG.INTERNAL_DOMAINS[i]);
      if (!d) { continue; }
      if (d.charAt(0) !== '@') { d = '@' + d; }
      if (a.length >= d.length && a.lastIndexOf(d) === a.length - d.length) {
        return true;
      }
    }
    return false;
  }

  // 宛先一覧の「印」(種別+アドレスを並べ替えて連結)
  function signature(list) {
    return list.map(function (r) {
      return r.type + ':' + norm(r.address);
    }).sort().join('|');
  }

  // To / Cc / Bcc をすべて読み取る。cb(err, list) の list は To→Cc→Bcc の順
  function getRecipients(item, cb) {
    var byType = { to: [], cc: [], bcc: [] };
    var pending = TYPES.length;
    var failed = false;

    TYPES.forEach(function (t) {
      item[t].getAsync(function (r) {
        if (failed) { return; }
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          failed = true;
          cb(r.error || new Error('getAsync failed: ' + t));
          return;
        }
        (r.value || []).forEach(function (x) {
          byType[t].push({
            type: t,
            name: x.displayName || '',
            address: x.emailAddress || ''
          });
        });
        pending--;
        if (pending === 0) {
          cb(null, byType.to.concat(byType.cc, byType.bcc));
        }
      });
    });
  }

  function setRecipients(item, type, arr, cb) {
    item[type].setAsync(arr, function (r) {
      cb(r.status === Office.AsyncResultStatus.Succeeded
        ? null
        : (r.error || new Error('setAsync failed: ' + type)));
    });
  }

  // 削除対象を除いた宛先と、種別ごとの設定値を計算する(純粋関数)
  function planChanges(all, removed) {
    var kept = all.filter(function (r) { return removed.indexOf(r) === -1; });
    var byType = { to: [], cc: [], bcc: [] };
    var changed = { to: false, cc: false, bcc: false };
    kept.forEach(function (r) {
      byType[r.type].push({ displayName: r.name, emailAddress: r.address });
    });
    removed.forEach(function (r) { changed[r.type] = true; });
    return { kept: kept, byType: byType, changed: changed };
  }

  // ---- 確認済みの印の保存(sessionData と customProperties の両方) ----

  function storeGet(item, cb) {
    function fromProps() {
      try {
        item.loadCustomPropertiesAsync(function (r) {
          if (r.status === Office.AsyncResultStatus.Succeeded) {
            cb(r.value.get(CONFIG.STATE_KEY) || null);
          } else {
            cb(null);
          }
        });
      } catch (e) { cb(null); }
    }
    try {
      item.sessionData.getAsync(CONFIG.STATE_KEY, function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded && r.value) {
          cb(r.value);
        } else {
          fromProps();
        }
      });
    } catch (e) { fromProps(); }
  }

  function storeSet(item, value, cb) {
    var done = 0;
    function fin() { done++; if (done === 2) { cb(); } }
    try {
      item.sessionData.setAsync(CONFIG.STATE_KEY, value, function () { fin(); });
    } catch (e) { fin(); }
    try {
      item.loadCustomPropertiesAsync(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          r.value.set(CONFIG.STATE_KEY, value);
          r.value.saveAsync(function () { fin(); });
        } else { fin(); }
      });
    } catch (e) { fin(); }
  }

  function storeClear(item, cb) {
    var done = 0;
    function fin() { done++; if (done === 2) { cb(); } }
    try {
      item.sessionData.removeAsync(CONFIG.STATE_KEY, function () { fin(); });
    } catch (e) { fin(); }
    try {
      item.loadCustomPropertiesAsync(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          r.value.remove(CONFIG.STATE_KEY);
          r.value.saveAsync(function () { fin(); });
        } else { fin(); }
      });
    } catch (e) { fin(); }
  }

  // ---- 送信時イベント(OnMessageSend) -------------------------------

  function onMessageSendHandler(event) {
    var item = Office.context.mailbox.item;

    function allow() { event.completed({ allowEvent: true }); }

    try {
      getRecipients(item, function (err, list) {
        // 宛先を読めない場合は確認できないため、そのまま送信を通す(フェイルオープン)
        if (err) { allow(); return; }

        var external = list.filter(function (r) { return !isInternal(r.address); });
        if (external.length === 0) { allow(); return; }

        var sig = signature(list);
        storeGet(item, function (saved) {
          if (saved && saved === sig) {
            // 確認済み: 印を消して送信を通す
            storeClear(item, allow);
            return;
          }
          event.completed({
            allowEvent: false,
            errorMessage: '社内ドメイン以外の宛先が含まれています。' +
              '「宛先を確認する」を押して、送信先を確認してください。' +
              '確認後、もう一度「送信」を押すと送信されます。',
            errorMessageMarkdown: '社内ドメイン以外の宛先が含まれています。\n\n' +
              '**「宛先を確認する」** を押して、送信先を確認してください。\n\n' +
              '確認後、もう一度 **「送信」** を押すと送信されます。',
            cancelLabel: '宛先を確認する',
            commandId: CONFIG.PANE_BUTTON_ID
          });
        });
      });
    } catch (e) {
      allow();
    }
  }

  if (typeof Office !== 'undefined' && Office.actions && Office.actions.associate) {
    Office.actions.associate('onMessageSendHandler', onMessageSendHandler);
  }

  // ---- タスクペイン --------------------------------------------------

  var pane = { all: [], boxes: [] };   // boxes: [{ rec, input }]

  function byId(id) { return document.getElementById(id); }

  function showStatus(msg, kind) {
    var el = byId('status');
    if (!el) { return; }
    el.textContent = msg || '';
    el.className = 'status' + (kind ? ' ' + kind : '');
  }

  function setBusy(busy) {
    ['btn-all', 'btn-none', 'btn-confirm', 'btn-cancel'].forEach(function (id) {
      var b = byId(id);
      if (b) { b.disabled = busy; }
    });
  }

  function setAll(checked) {
    pane.boxes.forEach(function (b) { b.input.checked = checked; });
  }

  function closePane() {
    try { Office.context.ui.closeContainer(); } catch (e) { /* 閉じられない環境では何もしない */ }
  }

  function renderList() {
    var box = byId('recipient-list');
    box.innerHTML = '';
    pane.boxes = [];

    var internalCount = 0;
    pane.all.forEach(function (rec) {
      if (isInternal(rec.address)) { internalCount++; return; }

      var label = document.createElement('label');
      label.className = 'row';

      var input = document.createElement('input');
      input.type = 'checkbox';
      input.checked = true;

      var tag = document.createElement('span');
      tag.className = 'tag tag-' + rec.type;
      tag.textContent = TYPE_LABEL[rec.type];

      var txt = document.createElement('span');
      txt.className = 'txt';
      var shown = rec.name && norm(rec.name) !== norm(rec.address)
        ? rec.name + ' <' + rec.address + '>'
        : rec.address;
      txt.textContent = shown;

      label.appendChild(input);
      label.appendChild(tag);
      label.appendChild(txt);
      box.appendChild(label);
      pane.boxes.push({ rec: rec, input: input });
    });

    if (pane.boxes.length === 0) {
      var empty = document.createElement('p');
      empty.className = 'note';
      empty.textContent = '確認対象の宛先はありません。';
      box.appendChild(empty);
    }

    var note = byId('internal-note');
    if (internalCount > 0) {
      note.textContent = '社内ドメイン宛て ' + internalCount +
        ' 件は確認対象外です(常に送信されます)。';
    } else {
      note.textContent = '';
    }
  }

  function loadList() {
    var item = Office.context.mailbox.item;
    showStatus('宛先を読み込んでいます…');
    getRecipients(item, function (err, list) {
      if (err) {
        showStatus('宛先を読み込めませんでした。', 'error');
        return;
      }
      pane.all = list;
      renderList();
      showStatus('');
    });
  }

  function onConfirm() {
    var item = Office.context.mailbox.item;
    var removed = [];
    pane.boxes.forEach(function (b) {
      if (!b.input.checked) { removed.push(b.rec); }
    });
    var plan = planChanges(pane.all, removed);

    if (plan.kept.length === 0) {
      showStatus('送信先が1件もありません。1件以上チェックしてください。', 'error');
      return;
    }

    setBusy(true);
    showStatus('宛先を反映しています…');

    var targets = TYPES.filter(function (t) { return plan.changed[t]; });

    function finalize() {
      // 実際の宛先を読み直して印を保存する(表示名の補正などを吸収)
      getRecipients(item, function (err, list) {
        if (err) {
          setBusy(false);
          showStatus('宛先の確認に失敗しました。もう一度お試しください。', 'error');
          return;
        }
        storeSet(item, signature(list), function () {
          if (CONFIG.AUTO_SEND_AFTER_CONFIRM && typeof item.sendAsync === 'function') {
            showStatus('確認しました。送信します…', 'ok');
            item.sendAsync(function (r) {
              if (r.status !== Office.AsyncResultStatus.Succeeded) {
                setBusy(false);
                showStatus('自動送信に失敗しました。「送信」を押してください。', 'error');
              }
            });
            return;
          }
          showStatus('確認しました。ウィンドウを閉じたら、もう一度「送信」を押してください。', 'ok');
          setTimeout(closePane, 1200);
        });
      });
    }

    function next(i) {
      if (i >= targets.length) { finalize(); return; }
      var t = targets[i];
      setRecipients(item, t, plan.byType[t], function (err) {
        if (err) {
          setBusy(false);
          showStatus('宛先を反映できませんでした(' + TYPE_LABEL[t] + ')。', 'error');
          return;
        }
        next(i + 1);
      });
    }
    next(0);
  }

  function initPane() {
    var item = Office.context.mailbox.item;

    byId('btn-all').onclick = function () { setAll(true); };
    byId('btn-none').onclick = function () { setAll(false); };
    byId('btn-confirm').onclick = onConfirm;
    byId('btn-cancel').onclick = closePane;

    try {
      item.subject.getAsync(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          byId('subject').textContent = '件名: ' + (r.value || '(件名なし)');
        }
      });
    } catch (e) { /* 件名が取れなくても続行 */ }

    loadList();
  }

  if (typeof document !== 'undefined' && typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady(function () {
      if (document.getElementById('recipient-list')) { initPane(); }
    });
  }

  // テスト用(ブラウザー/JSランタイムでは module が無いので何も起きない)
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
      CONFIG: CONFIG,
      isInternal: isInternal,
      signature: signature,
      getRecipients: getRecipients,
      planChanges: planChanges,
      onMessageSendHandler: onMessageSendHandler
    };
  }
})();
