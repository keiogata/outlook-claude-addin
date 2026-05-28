/**
 * claude.js — Anthropic Claude API ラッパー
 *
 * Claude APIに対してメッセージを送信し、回答を取得する。
 * APIキーはlocalStorageに保存（個人利用前提）。
 *
 * 注意: ブラウザから直接APIを呼び出すとAPIキーが露出するため、
 *       個人利用・ローカル環境での使用を前提とする。
 */

const ClaudeAPI = (() => {
  // 使用するモデル
  const MODEL = 'claude-opus-4-6';

  // 1回の会話で保持する最大往復数
  const MAX_HISTORY = 10;

  // localStorageのキー名
  const STORAGE_KEY = 'outlook_claude_api_key';

  /**
   * APIキーをlocalStorageに保存する
   * @param {string} apiKey - AnthropicのAPIキー
   */
  function saveApiKey(apiKey) {
    localStorage.setItem(STORAGE_KEY, apiKey.trim());
  }

  /**
   * localStorageからAPIキーを取得する
   * @returns {string|null} APIキー、未設定の場合はnull
   */
  function getApiKey() {
    return localStorage.getItem(STORAGE_KEY);
  }

  /**
   * APIキーが設定されているか確認する
   * @returns {boolean}
   */
  function hasApiKey() {
    const key = getApiKey();
    return key !== null && key.trim() !== '';
  }

  /**
   * APIキーを削除する
   */
  function clearApiKey() {
    localStorage.removeItem(STORAGE_KEY);
  }

  /**
   * Claude APIにメッセージを送信する（ストリーミング対応）
   *
   * @param {Array} messages - 会話履歴（{role: "user"|"assistant", content: string}の配列）
   * @param {string} systemPrompt - システムプロンプト
   * @param {function} onChunk - ストリームのチャンクを受信するコールバック (text) => void
   * @param {function} onDone - 完了時のコールバック (fullText) => void
   * @param {function} onError - エラー時のコールバック (error) => void
   */
  async function sendMessage(messages, systemPrompt, onChunk, onDone, onError) {
    const apiKey = getApiKey();

    if (!apiKey) {
      onError(new Error('APIキーが設定されていません。設定画面からAnthropicのAPIキーを入力してください。'));
      return;
    }

    // 会話履歴を直近MAX_HISTORY往復に制限
    const trimmedMessages = messages.slice(-MAX_HISTORY * 2);

    try {
      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          // ブラウザCORSのためdangerouslyAllowBrowserフラグ相当の設定
          'anthropic-dangerous-direct-browser-access': 'true',
        },
        body: JSON.stringify({
          model: MODEL,
          max_tokens: 2048,
          stream: true,
          system: systemPrompt,
          messages: trimmedMessages,
        }),
      });

      // HTTPエラーチェック
      if (!response.ok) {
        let errorMessage = `APIエラー: HTTP ${response.status}`;
        try {
          const errorData = await response.json();
          if (errorData.error && errorData.error.message) {
            errorMessage = `APIエラー: ${errorData.error.message}`;
          }
          // 認証エラー
          if (response.status === 401) {
            errorMessage = 'APIキーが無効です。設定画面で正しいAPIキーを入力してください。';
          }
          // レート制限
          if (response.status === 429) {
            errorMessage = 'APIの使用制限に達しました。しばらく待ってから再度お試しください。';
          }
        } catch (_) {
          // JSONパース失敗はデフォルトメッセージを使用
        }
        onError(new Error(errorMessage));
        return;
      }

      // ストリーミングレスポンスを処理
      const reader = response.body.getReader();
      const decoder = new TextDecoder('utf-8');
      let fullText = '';
      let buffer = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');

        // 最後の不完全な行をバッファに残す
        buffer = lines.pop() || '';

        for (const line of lines) {
          // Server-Sent Eventsの形式: "data: {...}"
          if (line.startsWith('data: ')) {
            const data = line.slice(6).trim();

            // ストリーム終了
            if (data === '[DONE]') continue;

            try {
              const event = JSON.parse(data);

              // テキストデルタイベントを処理
              if (
                event.type === 'content_block_delta' &&
                event.delta &&
                event.delta.type === 'text_delta'
              ) {
                const chunk = event.delta.text;
                fullText += chunk;
                onChunk(chunk);
              }
            } catch (_) {
              // JSONパース失敗は無視（不完全なチャンクの場合）
            }
          }
        }
      }

      onDone(fullText);
    } catch (error) {
      // ネットワークエラーなど
      if (error.name === 'TypeError' && error.message.includes('fetch')) {
        onError(new Error(
          'APIへの接続に失敗しました。\n' +
          'ネットワーク接続を確認してください。\n' +
          '社内プロキシ環境の場合は、プロキシ設定をご確認ください。'
        ));
      } else {
        onError(error);
      }
    }
  }

  // 公開API
  return {
    saveApiKey,
    getApiKey,
    hasApiKey,
    clearApiKey,
    sendMessage,
    MODEL,
  };
})();
