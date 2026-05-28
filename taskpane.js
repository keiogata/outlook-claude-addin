/**
 * taskpane.js — Outlook Claude AIアシスタント メインロジック
 *
 * Office.js初期化 → チャットUI管理 → Claude API連携 → Graph API連携（Phase 2以降）
 */

'use strict';

// ===== グローバル状態 =====
let conversationHistory = []; // 会話履歴（Claude APIに送る形式）
let isLoading = false;        // APIリクエスト中フラグ
let lastUserMessage = null;   // リトライ用：直前のユーザーメッセージ

// Office.js 初期化完了後に実行
// Outlookの中でも、ブラウザ単体でも動作するように対応
if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady(() => {
    initApp();
  });
} else {
  // ブラウザ直接アクセス時のフォールバック
  document.addEventListener('DOMContentLoaded', initApp);
}

/**
 * アプリ初期化
 */
function initApp() {
  // DOM要素を取得
  const settingsBtn = document.getElementById('settings-btn');
  const settingsPanel = document.getElementById('settings-panel');
  const saveApiKeyBtn = document.getElementById('save-api-key-btn');
  const cancelSettingsBtn = document.getElementById('cancel-settings-btn');
  const apiKeyInput = document.getElementById('api-key-input');
  const settingsSavedMsg = document.getElementById('settings-saved-msg');
  const sendBtn = document.getElementById('send-btn');
  const messageInput = document.getElementById('message-input');
  const clearBtn = document.getElementById('clear-btn');
  const apiKeyWarning = document.getElementById('api-key-warning');
  const apiKeyWarningLink = document.getElementById('api-key-warning-link');

  // APIキーが設定されていない場合は警告表示
  if (!ClaudeAPI.hasApiKey()) {
    apiKeyWarning.classList.add('visible');
  }

  // ===== 設定ボタン =====
  settingsBtn.addEventListener('click', () => {
    const isVisible = settingsPanel.classList.contains('visible');
    if (isVisible) {
      settingsPanel.classList.remove('visible');
    } else {
      // 既存のAPIキーをマスク表示
      const existingKey = ClaudeAPI.getApiKey();
      if (existingKey) {
        apiKeyInput.value = '●'.repeat(20); // マスク表示
        apiKeyInput.dataset.hasKey = 'true';
      } else {
        apiKeyInput.value = '';
        apiKeyInput.dataset.hasKey = 'false';
      }
      settingsSavedMsg.style.display = 'none';
      settingsPanel.classList.add('visible');
    }
  });

  // APIキー入力欄をクリックしたときマスクを解除
  apiKeyInput.addEventListener('focus', () => {
    if (apiKeyInput.dataset.hasKey === 'true') {
      apiKeyInput.value = '';
      apiKeyInput.dataset.hasKey = 'false';
    }
  });

  // APIキー保存
  saveApiKeyBtn.addEventListener('click', () => {
    const key = apiKeyInput.value.trim();

    // 変更なし（マスク表示のまま）の場合はスキップ
    if (key === '' && apiKeyInput.dataset.hasKey === 'true') {
      settingsPanel.classList.remove('visible');
      return;
    }

    if (key === '') {
      apiKeyInput.style.borderColor = '#a4262c';
      apiKeyInput.placeholder = 'APIキーを入力してください';
      return;
    }

    ClaudeAPI.saveApiKey(key);
    apiKeyInput.style.borderColor = '';
    settingsSavedMsg.style.display = 'block';
    apiKeyWarning.classList.remove('visible');

    setTimeout(() => {
      settingsPanel.classList.remove('visible');
      settingsSavedMsg.style.display = 'none';
    }, 1200);
  });

  // 設定キャンセル
  cancelSettingsBtn.addEventListener('click', () => {
    settingsPanel.classList.remove('visible');
  });

  // 警告バーからの設定画面へのリンク
  apiKeyWarningLink.addEventListener('click', () => {
    apiKeyWarning.classList.remove('visible');
    settingsPanel.classList.add('visible');
    apiKeyInput.focus();
  });

  // ===== メッセージ入力 =====
  // Enterキーで送信（Shift+Enterは改行）
  messageInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });

  // テキストエリアの高さを自動調整
  messageInput.addEventListener('input', () => {
    messageInput.style.height = 'auto';
    messageInput.style.height = Math.min(messageInput.scrollHeight, 120) + 'px';
  });

  // 送信ボタン
  sendBtn.addEventListener('click', sendMessage);

  // 会話クリアボタン
  clearBtn.addEventListener('click', () => {
    conversationHistory = [];
    const chatContainer = document.getElementById('chat-container');
    chatContainer.innerHTML = '';
    renderWelcomeMessage();
  });

  // ウェルカムメッセージを表示
  renderWelcomeMessage();
}

/**
 * ウェルカムメッセージと例文クエリを表示する
 */
function renderWelcomeMessage() {
  const chatContainer = document.getElementById('chat-container');

  const exampleQueries = [
    '来週の自分の予定を教えて',
    '西尾さんと60分空いている時間は？',
    'KPITのMinamiさんからの先週のメールをまとめて',
    'CHG定例の議事録を探して',
  ];

  const welcomeDiv = document.createElement('div');
  welcomeDiv.id = 'welcome-msg';
  welcomeDiv.innerHTML = `
    <div class="welcome-icon">🤖</div>
    <h3>Claude AI アシスタント</h3>
    <p>メール・予定表・Teamsを自然言語で検索・操作できます。</p>
    <div class="example-queries">
      ${exampleQueries.map(q => `<p class="example-query">${q}</p>`).join('')}
    </div>
  `;

  chatContainer.appendChild(welcomeDiv);

  // 例文クエリをクリックしたら入力欄に設定
  welcomeDiv.querySelectorAll('.example-query').forEach((el) => {
    el.addEventListener('click', () => {
      const messageInput = document.getElementById('message-input');
      messageInput.value = el.textContent;
      messageInput.focus();
    });
  });
}

/**
 * メッセージを送信してClaudeから回答を得る
 */
async function sendMessage() {
  if (isLoading) return;

  const messageInput = document.getElementById('message-input');
  const userText = messageInput.value.trim();

  if (!userText) return;

  // APIキーチェック
  if (!ClaudeAPI.hasApiKey()) {
    document.getElementById('api-key-warning').classList.add('visible');
    document.getElementById('settings-panel').classList.add('visible');
    return;
  }

  // ウェルカムメッセージを削除（初回送信時）
  const welcomeMsg = document.getElementById('welcome-msg');
  if (welcomeMsg) welcomeMsg.remove();

  // 入力欄をリセット
  messageInput.value = '';
  messageInput.style.height = 'auto';
  lastUserMessage = userText;

  // ユーザーメッセージをUIに追加
  appendMessage('user', userText);

  // 会話履歴に追加
  conversationHistory.push({ role: 'user', content: userText });

  // 送信ボタンを無効化＆ローディング表示
  setLoading(true);

  // システムプロンプト（現在のメール・ユーザー情報を注入）
  const systemPrompt = buildSystemPrompt();

  // Claude APIにストリーミングで送信
  let assistantBubble = null;

  ClaudeAPI.sendMessage(
    conversationHistory,
    systemPrompt,
    // onChunk: テキストが届くたびにUIを更新
    (chunk) => {
      if (!assistantBubble) {
        assistantBubble = appendMessage('assistant', '');
      }
      appendChunkToBubble(assistantBubble, chunk);
    },
    // onDone: 完了時に会話履歴に追加
    (fullText) => {
      conversationHistory.push({ role: 'assistant', content: fullText });
      setLoading(false);
    },
    // onError: エラー表示
    (error) => {
      appendMessage('error', error.message, true);
      setLoading(false);
    }
  );
}

/**
 * システムプロンプトを構築する
 * @returns {string} システムプロンプト
 */
function buildSystemPrompt() {
  const now = new Date();
  const dateStr = now.toLocaleDateString('ja-JP', {
    year: 'numeric', month: 'long', day: 'numeric', weekday: 'long'
  });
  const timeStr = now.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' });

  // 現在選択中のメールアイテムの情報を取得（Office.js）
  let currentItemInfo = '';
  try {
    const item = Office.context.mailbox.item;
    if (item) {
      currentItemInfo = `\n現在開いているアイテム: ${item.subject || '（件名なし）'}`;
    }
  } catch (_) {
    // Office.jsが利用できない場合（ローカルテスト時）
  }

  return `あなたはOutlook内で動作するAIアシスタントです。
Honda SDV事業開発部 ELS部のユーザーをサポートします。

現在日時: ${dateStr} ${timeStr}${currentItemInfo}

【役割】
- ユーザーの質問に対して、日本語で簡潔・丁寧に回答してください。
- メール・予定表・Teamsに関する質問には、具体的なアドバイスを提供してください。
- Phase 1では Microsoft Graph APIとの連携はまだ実装されていません。
  その旨をユーザーに説明しつつ、Graph APIで実現可能な方法を提案してください。
- 一般的な質問（文章作成・要約・翻訳・EV充電規格など）にも対応してください。

【注意】
- 不明な点は正直に「わかりません」と伝えてください。
- 個人情報・機密情報の取り扱いに注意してください。`;
}

/**
 * チャットにメッセージバブルを追加する
 * @param {string} role - "user" | "assistant" | "error"
 * @param {string} text - メッセージテキスト
 * @param {boolean} showRetry - リトライボタンを表示するか
 * @returns {HTMLElement} メッセージバブル要素
 */
function appendMessage(role, text, showRetry = false) {
  const chatContainer = document.getElementById('chat-container');

  const messageDiv = document.createElement('div');
  messageDiv.classList.add('message', role);

  const labelDiv = document.createElement('div');
  labelDiv.classList.add('message-label');
  labelDiv.textContent = role === 'user' ? 'あなた' : role === 'error' ? 'エラー' : 'Claude AI';

  const bubbleDiv = document.createElement('div');
  bubbleDiv.classList.add('message-bubble');
  bubbleDiv.textContent = text;

  messageDiv.appendChild(labelDiv);
  messageDiv.appendChild(bubbleDiv);

  // エラー時のリトライボタン
  if (showRetry && lastUserMessage) {
    const retryBtn = document.createElement('button');
    retryBtn.classList.add('retry-btn');
    retryBtn.textContent = '再試行';
    retryBtn.addEventListener('click', () => {
      messageDiv.remove();
      const messageInput = document.getElementById('message-input');
      messageInput.value = lastUserMessage;
      sendMessage();
    });
    messageDiv.appendChild(retryBtn);
  }

  // ローディングインジケーターの前に挿入
  const loading = document.getElementById('loading');
  chatContainer.insertBefore(messageDiv, loading);

  // 最下部にスクロール
  chatContainer.scrollTop = chatContainer.scrollHeight;

  return bubbleDiv;
}

/**
 * ストリーミングチャンクをバブルに追記する
 * @param {HTMLElement} bubbleEl - バブル要素
 * @param {string} chunk - 追記するテキスト
 */
function appendChunkToBubble(bubbleEl, chunk) {
  bubbleEl.textContent += chunk;

  // マークダウン的な改行を処理（\nを<br>に）
  bubbleEl.innerHTML = bubbleEl.textContent
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');

  // 最下部にスクロール
  const chatContainer = document.getElementById('chat-container');
  chatContainer.scrollTop = chatContainer.scrollHeight;
}

/**
 * ローディング状態を設定する
 * @param {boolean} loading - ローディング中かどうか
 */
function setLoading(loading) {
  isLoading = loading;
  const sendBtn = document.getElementById('send-btn');
  const loadingEl = document.getElementById('loading');
  const messageInput = document.getElementById('message-input');

  if (loading) {
    sendBtn.disabled = true;
    loadingEl.classList.add('visible');
    messageInput.disabled = true;

    // ローディングを最下部にスクロール
    const chatContainer = document.getElementById('chat-container');
    chatContainer.scrollTop = chatContainer.scrollHeight;
  } else {
    sendBtn.disabled = false;
    loadingEl.classList.remove('visible');
    messageInput.disabled = false;
    messageInput.focus();
  }
}
