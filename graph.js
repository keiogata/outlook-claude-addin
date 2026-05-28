/**
 * graph.js — Microsoft Graph APIラッパー
 *
 * Phase 2・3で実装予定。
 * メール検索・予定表検索・空き時間検索などのGraph API呼び出しを管理する。
 */

const GraphAPI = (() => {
  const BASE_URL = 'https://graph.microsoft.com/v1.0';

  /**
   * Graph APIにGETリクエストを送る汎用関数
   * @param {string} endpoint - APIエンドポイント（/me/messagesなど）
   * @param {string} accessToken - アクセストークン
   * @returns {Promise<object>} レスポンスデータ
   */
  async function get(endpoint, accessToken) {
    const response = await fetch(`${BASE_URL}${endpoint}`, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
    });

    if (!response.ok) {
      const error = await response.json().catch(() => ({}));
      throw new Error(
        error.error?.message || `Graph APIエラー: HTTP ${response.status}`
      );
    }

    return response.json();
  }

  /**
   * Graph APIにPOSTリクエストを送る汎用関数
   * @param {string} endpoint - APIエンドポイント
   * @param {object} body - リクエストボディ
   * @param {string} accessToken - アクセストークン
   * @returns {Promise<object>} レスポンスデータ
   */
  async function post(endpoint, body, accessToken) {
    const response = await fetch(`${BASE_URL}${endpoint}`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const error = await response.json().catch(() => ({}));
      throw new Error(
        error.error?.message || `Graph APIエラー: HTTP ${response.status}`
      );
    }

    return response.json();
  }

  // ===== Phase 3: メール機能 =====

  /**
   * メールを検索する
   * @param {string} query - 検索キーワード
   * @param {string} accessToken - アクセストークン
   * @param {number} top - 取得件数（デフォルト10）
   * @returns {Promise<Array>} メール一覧
   */
  async function searchEmails(query, accessToken, top = 10) {
    // Phase 3で実装
    console.warn('[graph.js] メール検索はPhase 3で実装予定です。');
    return [];
    /*
    // 実装例:
    const encoded = encodeURIComponent(query);
    const data = await get(
      `/me/messages?$search="${encoded}"&$top=${top}&$select=subject,from,receivedDateTime,bodyPreview`,
      accessToken
    );
    return data.value || [];
    */
  }

  /**
   * 差出人でメールをフィルターする
   * @param {string} emailAddress - 差出人のメールアドレス
   * @param {string} accessToken - アクセストークン
   * @param {string} since - 開始日時（ISO 8601形式）
   * @returns {Promise<Array>} メール一覧
   */
  async function getEmailsFromSender(emailAddress, accessToken, since = null) {
    // Phase 3で実装
    console.warn('[graph.js] メール検索はPhase 3で実装予定です。');
    return [];
  }

  /**
   * メール本文を取得する
   * @param {string} messageId - メールID
   * @param {string} accessToken - アクセストークン
   * @returns {Promise<object>} メール詳細
   */
  async function getEmailBody(messageId, accessToken) {
    // Phase 3で実装
    console.warn('[graph.js] メール本文取得はPhase 3で実装予定です。');
    return null;
  }

  // ===== Phase 2: 予定表機能 =====

  /**
   * 自分の予定を取得する
   * @param {string} startDate - 開始日時（ISO 8601形式）
   * @param {string} endDate - 終了日時（ISO 8601形式）
   * @param {string} accessToken - アクセストークン
   * @returns {Promise<Array>} 予定一覧
   */
  async function getMyCalendar(startDate, endDate, accessToken) {
    // Phase 2で実装
    console.warn('[graph.js] 予定表機能はPhase 2で実装予定です。');
    return [];
    /*
    // 実装例:
    const start = encodeURIComponent(startDate);
    const end = encodeURIComponent(endDate);
    const data = await get(
      `/me/calendarView?startDateTime=${start}&endDateTime=${end}&$select=subject,start,end,organizer,attendees`,
      accessToken
    );
    return data.value || [];
    */
  }

  /**
   * 他ユーザーの予定を取得する（共有されている場合のみ）
   * @param {string} userId - 対象ユーザーのIDまたはメールアドレス
   * @param {string} startDate - 開始日時
   * @param {string} endDate - 終了日時
   * @param {string} accessToken - アクセストークン
   * @returns {Promise<Array>} 予定一覧
   */
  async function getUserCalendar(userId, startDate, endDate, accessToken) {
    // Phase 2で実装
    console.warn('[graph.js] 他ユーザー予定表はPhase 2で実装予定です（共有が必要）。');
    return [];
  }

  /**
   * 空き時間を検索する（/me/findMeetingTimes）
   * @param {Array} attendees - 参加者のメールアドレス配列
   * @param {string} startDate - 検索開始日時
   * @param {string} endDate - 検索終了日時
   * @param {number} durationMinutes - 会議時間（分）
   * @param {string} accessToken - アクセストークン
   * @returns {Promise<Array>} 空き時間候補
   */
  async function findMeetingTimes(attendees, startDate, endDate, durationMinutes, accessToken) {
    // Phase 2で実装
    console.warn('[graph.js] 空き時間検索はPhase 2で実装予定です。');
    return [];
    /*
    // 実装例:
    const body = {
      attendees: attendees.map(email => ({
        emailAddress: { address: email },
        type: 'required',
      })),
      timeConstraint: {
        activityDomain: 'work',
        timeslots: [{ start: { dateTime: startDate, timeZone: 'Tokyo Standard Time' }, end: { dateTime: endDate, timeZone: 'Tokyo Standard Time' } }],
      },
      meetingDuration: `PT${durationMinutes}M`,
      minimumAttendeePercentage: 100,
    };
    const data = await post('/me/findMeetingTimes', body, accessToken);
    return data.meetingTimeSuggestions || [];
    */
  }

  /**
   * ユーザー情報を取得する（名前・メールアドレス）
   * @param {string} accessToken - アクセストークン
   * @returns {Promise<object>} ユーザー情報
   */
  async function getMyProfile(accessToken) {
    return get('/me?$select=displayName,mail,userPrincipalName', accessToken);
  }

  /**
   * 社内連絡先（People）を検索する
   * @param {string} query - 検索キーワード（名前など）
   * @param {string} accessToken - アクセストークン
   * @returns {Promise<Array>} 連絡先一覧
   */
  async function searchPeople(query, accessToken) {
    // Phase 2で実装
    console.warn('[graph.js] People検索はPhase 2で実装予定です。');
    return [];
  }

  return {
    searchEmails,
    getEmailsFromSender,
    getEmailBody,
    getMyCalendar,
    getUserCalendar,
    findMeetingTimes,
    getMyProfile,
    searchPeople,
  };
})();
