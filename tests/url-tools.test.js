// Run with: node --test tests/url-tools.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const { safeUrl, toEmbedUrl, parseYouTubeTime } = require('../url-tools.js');

test('safeUrl: adds https:// when no protocol is typed', () => {
  assert.equal(safeUrl('example.com'), 'https://example.com/');
  assert.equal(safeUrl('  example.com/page?a=1  '), 'https://example.com/page?a=1');
  assert.equal(safeUrl('localhost:3000'), 'https://localhost:3000/');
});

test('safeUrl: keeps http and https', () => {
  assert.equal(safeUrl('http://intranet.local/dashboard'), 'http://intranet.local/dashboard');
  assert.equal(safeUrl('HTTPS://Example.com'), 'https://example.com/');
});

test('safeUrl: rejects dangerous or unknown schemes', () => {
  for (const bad of [
    'javascript:alert(1)',
    'JavaScript:alert(document.domain)',
    ' javascript:parent.Office',
    'data:text/html,<script>alert(1)</script>',
    'file:///C:/Windows/win.ini',
    'blob:https://example.com/123',
    'vbscript:msgbox(1)',
    'about:blank',
    'mailto:someone@example.com',
  ]) {
    assert.equal(safeUrl(bad), null, bad);
  }
});

test('safeUrl: rejects empty, non-string and credential URLs', () => {
  assert.equal(safeUrl(''), null);
  assert.equal(safeUrl('   '), null);
  assert.equal(safeUrl(null), null);
  assert.equal(safeUrl(undefined), null);
  assert.equal(safeUrl(42), null);
  assert.equal(safeUrl('https://user:pass@example.com/'), null);
  assert.equal(safeUrl('https://'), null);
});

test('parseYouTubeTime: supports all YouTube formats', () => {
  assert.equal(parseYouTubeTime('90'), 90);
  assert.equal(parseYouTubeTime('90s'), 90);
  assert.equal(parseYouTubeTime('1m30s'), 90);
  assert.equal(parseYouTubeTime('2h1m'), 7260);
  assert.equal(parseYouTubeTime('1h2m3s'), 3723);
  assert.equal(parseYouTubeTime(''), 0);
  assert.equal(parseYouTubeTime(null), 0);
  assert.equal(parseYouTubeTime('abc'), 0);
});

test('toEmbedUrl: YouTube watch, short, shorts, live, music', () => {
  assert.equal(toEmbedUrl('https://www.youtube.com/watch?v=dQw4w9WgXcQ'), 'https://www.youtube.com/embed/dQw4w9WgXcQ');
  assert.equal(toEmbedUrl('https://m.youtube.com/watch?v=dQw4w9WgXcQ'), 'https://www.youtube.com/embed/dQw4w9WgXcQ');
  assert.equal(toEmbedUrl('https://music.youtube.com/watch?v=dQw4w9WgXcQ'), 'https://www.youtube.com/embed/dQw4w9WgXcQ');
  assert.equal(toEmbedUrl('https://youtu.be/dQw4w9WgXcQ'), 'https://www.youtube.com/embed/dQw4w9WgXcQ');
  assert.equal(toEmbedUrl('https://youtu.be/dQw4w9WgXcQ?si=abc123'), 'https://www.youtube.com/embed/dQw4w9WgXcQ');
  assert.equal(toEmbedUrl('https://www.youtube.com/shorts/dQw4w9WgXcQ'), 'https://www.youtube.com/embed/dQw4w9WgXcQ');
  assert.equal(toEmbedUrl('https://www.youtube.com/live/dQw4w9WgXcQ?feature=share'), 'https://www.youtube.com/embed/dQw4w9WgXcQ');
});

test('toEmbedUrl: YouTube start times', () => {
  assert.equal(toEmbedUrl('https://youtu.be/dQw4w9WgXcQ?t=1m30s'), 'https://www.youtube.com/embed/dQw4w9WgXcQ?start=90');
  assert.equal(toEmbedUrl('https://www.youtube.com/watch?v=dQw4w9WgXcQ&t=2h1m'), 'https://www.youtube.com/embed/dQw4w9WgXcQ?start=7260');
  assert.equal(toEmbedUrl('https://www.youtube.com/watch?v=dQw4w9WgXcQ&t=90s'), 'https://www.youtube.com/embed/dQw4w9WgXcQ?start=90');
  assert.equal(toEmbedUrl('https://www.youtube.com/watch?v=dQw4w9WgXcQ&t=42'), 'https://www.youtube.com/embed/dQw4w9WgXcQ?start=42');
});

test('toEmbedUrl: YouTube playlists', () => {
  assert.equal(toEmbedUrl('https://www.youtube.com/watch?v=dQw4w9WgXcQ&list=PLabc_123-x'),
    'https://www.youtube.com/embed/dQw4w9WgXcQ?list=PLabc_123-x');
  assert.equal(toEmbedUrl('https://www.youtube.com/playlist?list=PLabc_123-x'),
    'https://www.youtube.com/embed/videoseries?list=PLabc_123-x');
});

test('toEmbedUrl: YouTube parameters cannot be injected', () => {
  // list=PL1&autoplay=1 (encoded) must not add an autoplay parameter
  assert.equal(toEmbedUrl('https://www.youtube.com/watch?v=dQw4w9WgXcQ&list=PL1%26autoplay%3D1'),
    'https://www.youtube.com/embed/dQw4w9WgXcQ');
  // A video id with odd characters is left untouched
  const odd = 'https://www.youtube.com/watch?v=abc%22%3E%3Cx';
  assert.equal(toEmbedUrl(odd), odd);
});

test('toEmbedUrl: lookalike domains are not converted', () => {
  for (const url of [
    'https://youtube.com.example.invalid/watch?v=dQw4w9WgXcQ',
    'https://notyoutube.com/watch?v=dQw4w9WgXcQ',
    'https://vimeo.com.example.invalid/123456',
  ]) {
    assert.equal(toEmbedUrl(url), url);
  }
});

test('toEmbedUrl: YouTube embed URLs stay as they are', () => {
  const url = 'https://www.youtube.com/embed/dQw4w9WgXcQ?start=10';
  assert.equal(toEmbedUrl(url), url);
});

test('toEmbedUrl: Vimeo public, unlisted, channel and group videos', () => {
  assert.equal(toEmbedUrl('https://vimeo.com/123456'), 'https://player.vimeo.com/video/123456');
  assert.equal(toEmbedUrl('https://vimeo.com/123456/abcdef1234'), 'https://player.vimeo.com/video/123456?h=abcdef1234');
  assert.equal(toEmbedUrl('https://vimeo.com/123456?h=abcdef1234'), 'https://player.vimeo.com/video/123456?h=abcdef1234');
  assert.equal(toEmbedUrl('https://vimeo.com/channels/staffpicks/123456'), 'https://player.vimeo.com/video/123456');
  assert.equal(toEmbedUrl('https://vimeo.com/groups/name/videos/123456'), 'https://player.vimeo.com/video/123456');
  const player = 'https://player.vimeo.com/video/123456';
  assert.equal(toEmbedUrl(player), player);
});

test('toEmbedUrl: other sites are unchanged', () => {
  assert.equal(toEmbedUrl('https://example.com/page'), 'https://example.com/page');
  assert.equal(toEmbedUrl('not a url'), 'not a url');
});
