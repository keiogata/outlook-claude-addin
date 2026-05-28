/**
 * auth.js — Microsoft Graph 認証処理（MSAL.js）
 *
 * Phase 2で実装予定。
 * OAuth 2.0 Authorization Code Flow with PKCEを使用して
 * Microsoft 365のメール・予定表データにアクセスする権限を取得する。
 *
 * 事前準備:
 *   1. Azure Portal でアプリを登録（Microsoft Entra ID）
 *   2. クライアントID（Application ID）を取得
 *   3. リダイレクトURIを設定（https://YOUR_USERNAME.github.io/outlook-claude-addin/auth.html）
 *   4. 下記スコープをAPIのアクセス許可に追加
 */

// ===== Azure ADアプリ設定（Phase 2で設定） =====
const AUTH_CONFIG = {
  // Azure Portalで発行されたクライアントID（アプリケーションID）
  // 例: "12345678-1234-1234-1234-123456789012"
  clientId: 'YOUR_CLIENT_ID_HERE',

  // テナントID（組織のディレクトリID）または "common"（マルチテナント）
  authority: 'https://login.microsoftonline.com/common',

  // 認証後のリダイレクト先
  redirectUri: 'https://YOUR_USERNAME.github.io/outlook-claude-addin/auth.html',

  // 要求するスコープ（権限）
  scopes: [
    'User.Read',           // ユーザー情報の読み取り
    'Mail.Read',           // メールの読み取り
    'Calendars.Read',      // 自分の予定表の読み取り
    'Calendars.Read.Shared', // 他ユーザーの共有予定表の読み取り
    'People.Read',         // 連絡先の読み取り
  ],
};

/**
 * MSALクライアントの初期化（Phase 2で実装）
 * npm install @azure/msal-browser が必要
 */
const AuthManager = (() => {
  let msalInstance = null;
  let currentAccount = null;

  /**
   * MSALを初期化する
   * Phase 2: @azure/msal-browser をロードして初期化
   */
  async function initialize() {
    // Phase 2で実装
    console.warn('[auth.js] Microsoft Graph認証はPhase 2で実装予定です。');
    return false;
  }

  /**
   * Microsoftアカウントでサインインする
   * @returns {Promise<object|null>} アカウント情報
   */
  async function signIn() {
    // Phase 2で実装
    console.warn('[auth.js] サインイン機能はPhase 2で実装予定です。');
    return null;
  }

  /**
   * サインアウトする
   */
  async function signOut() {
    // Phase 2で実装
    console.warn('[auth.js] サインアウト機能はPhase 2で実装予定です。');
  }

  /**
   * アクセストークンを取得する（サイレント更新 → ポップアップのフォールバック）
   * @returns {Promise<string|null>} アクセストークン
   */
  async function getAccessToken() {
    // Phase 2で実装
    console.warn('[auth.js] トークン取得機能はPhase 2で実装予定です。');
    return null;
  }

  /**
   * サインイン済みかどうかを確認する
   * @returns {boolean}
   */
  function isSignedIn() {
    return currentAccount !== null;
  }

  /**
   * 現在のアカウント情報を取得する
   * @returns {object|null}
   */
  function getAccount() {
    return currentAccount;
  }

  return {
    initialize,
    signIn,
    signOut,
    getAccessToken,
    isSignedIn,
    getAccount,
    AUTH_CONFIG,
  };
})();
