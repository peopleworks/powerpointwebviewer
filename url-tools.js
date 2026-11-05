/*
 * url-tools.js — URL helpers shared by viewer.html and dialog.html.
 *
 * Live Website Viewer for PowerPoint
 * Created by Pedro Hernandez — PeopleWorks Services
 *
 * Works in the browser (window.UrlTools) and in Node (require) so it can be
 * unit-tested with `node --test`.
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) {
    module.exports = api;
  } else {
    root.UrlTools = api;
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var ALLOWED_PROTOCOLS = { 'https:': true, 'http:': true };

  /**
   * Turn user input (or a URL saved in the presentation) into a safe URL.
   * Returns the normalized URL string, or null if it is not allowed.
   *
   * Only http(s) is accepted. Anything else (javascript:, data:, file:, ...)
   * is rejected, because the result is loaded into an iframe that shares
   * permissions with the add-in.
   */
  function safeUrl(input) {
    if (typeof input !== 'string') return null;
    var value = input.trim();
    if (!value) return null;

    if (!/^https?:\/\//i.test(value)) {
      // Any other explicit scheme (javascript:, data:, file:, ...) is refused.
      // "host:port" (e.g. localhost:3000) is not a scheme, so it is allowed.
      if (/^[a-z][a-z0-9+.-]*:(?!\d)/i.test(value)) return null;
      // No protocol typed: assume https
      value = 'https://' + value;
    }

    var u;
    try {
      u = new URL(value);
    } catch (e) {
      return null;
    }

    if (!ALLOWED_PROTOCOLS[u.protocol]) return null;
    if (!u.hostname) return null;
    if (u.username || u.password) return null; // https://user:pass@host

    return u.href;
  }

  /** Parse a YouTube time value: "90", "90s", "1m30s", "1h2m3s". Returns seconds or 0. */
  function parseYouTubeTime(value) {
    if (!value) return 0;
    if (/^\d+$/.test(value)) return parseInt(value, 10);
    var m = /^(?:(\d+)h)?(?:(\d+)m)?(?:(\d+)s)?$/i.exec(value);
    if (!m || (!m[1] && !m[2] && !m[3])) return 0;
    return (parseInt(m[1] || '0', 10) * 3600) +
           (parseInt(m[2] || '0', 10) * 60) +
            parseInt(m[3] || '0', 10);
  }

  var YT_ID = /^[A-Za-z0-9_-]{6,64}$/;
  var YT_LIST = /^[A-Za-z0-9_-]+$/;

  function youTubeEmbed(videoId, startAt, listId) {
    var params = [];
    if (startAt > 0) params.push('start=' + startAt);
    if (listId && YT_LIST.test(listId)) params.push('list=' + encodeURIComponent(listId));
    return 'https://www.youtube.com/embed/' + encodeURIComponent(videoId) +
           (params.length ? '?' + params.join('&') : '');
  }

  /**
   * Convert YouTube / Vimeo page links to their embeddable player URL.
   * Regular video pages block iframe embedding; the /embed/ players do not.
   * Any other URL is returned unchanged.
   */
  function toEmbedUrl(url) {
    var u;
    try {
      u = new URL(url);
    } catch (e) {
      return url;
    }
    var host = u.hostname.replace(/^www\./i, '').toLowerCase();

    // ─── YouTube ───
    if (host === 'youtube.com' || host === 'm.youtube.com' ||
        host === 'music.youtube.com' || host === 'youtube-nocookie.com') {
      if (u.pathname.indexOf('/embed/') === 0) return url; // already embeddable

      var videoId = u.searchParams.get('v');
      var pathMatch = u.pathname.match(/^\/(?:shorts|live)\/([^/?#]+)/);
      if (pathMatch) videoId = pathMatch[1];

      var listId = u.searchParams.get('list');
      var startAt = parseYouTubeTime(u.searchParams.get('t') || u.searchParams.get('start'));

      if (videoId && YT_ID.test(videoId)) {
        return youTubeEmbed(videoId, startAt, listId);
      }
      if (listId && YT_LIST.test(listId) && u.pathname.indexOf('/playlist') === 0) {
        return 'https://www.youtube.com/embed/videoseries?list=' + encodeURIComponent(listId);
      }
      return url;
    }

    // ─── YouTube short links (youtu.be/ID) ───
    if (host === 'youtu.be') {
      var id = u.pathname.split('/').filter(Boolean)[0];
      if (id && YT_ID.test(id)) {
        return youTubeEmbed(id, parseYouTubeTime(u.searchParams.get('t')), u.searchParams.get('list'));
      }
      return url;
    }

    // ─── Vimeo ───
    // vimeo.com/123456, vimeo.com/123456/abcdef1234 (unlisted),
    // vimeo.com/channels/name/123456, vimeo.com/groups/name/videos/123456
    if (host === 'vimeo.com') {
      var segments = u.pathname.split('/').filter(Boolean);
      for (var i = segments.length - 1; i >= 0; i--) {
        if (/^\d+$/.test(segments[i])) {
          // Unlisted videos need their privacy hash, or the player refuses to play
          var hash = u.searchParams.get('h') || segments[i + 1];
          var embed = 'https://player.vimeo.com/video/' + segments[i];
          if (hash && /^[0-9a-f]{6,}$/i.test(hash)) embed += '?h=' + hash;
          return embed;
        }
      }
    }

    return url;
  }

  return {
    safeUrl: safeUrl,
    toEmbedUrl: toEmbedUrl,
    parseYouTubeTime: parseYouTubeTime
  };
});
