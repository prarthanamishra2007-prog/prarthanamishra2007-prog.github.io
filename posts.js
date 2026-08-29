// ============================================================
// posts.js — the little reader that fetches YOUR writing
// from the Google notebook and hands it to the room pages,
// the post pages, and the homepage peeks.
//
// The URL below must match the one in stories.html (SHELF_URL).
//
// SPEED TRICK: your posts are remembered in the browser for a
// few minutes, so repeat visits paint instantly — the page shows
// what it already knows, then quietly refreshes in the background.
// ============================================================

const POSTS_URL = "https://script.google.com/macros/s/AKfycbwiGVGcNqYDgmXg2njFVMOrpuYKmQBbqbvV-kzTvsfZh2kF560SqXaDUJyG9CM37mMW/exec";

const POSTS_CACHE_KEY = 'postsCache_v1';

// ---- tiny cache helpers ----
function loadCachedPosts() {
  try {
    const raw = JSON.parse(localStorage.getItem(POSTS_CACHE_KEY));
    return Array.isArray(raw.posts) ? raw.posts : null;
  } catch (e) { return null; }
}
function saveCachedPosts(posts) {
  try { localStorage.setItem(POSTS_CACHE_KEY, JSON.stringify({ t: Date.now(), posts })); } catch (e) {}
}

// ---- the main reader ----
// fetchPosts(room, onUpdate)
//  · room      — 'thoughts' / 'books' / 'voices' (or leave empty for all)
//  · onUpdate  — called again when the freshly-fetched list arrives
// Returns the posts right away (from memory when possible), then always
// checks the notebook behind the scenes — so new posts appear fast and
// deleted ones disappear on every browser, no long stale waits.
async function fetchPosts(room, onUpdate) {
  const filter = posts => room ? posts.filter(p => p.room === room) : posts;

  async function fromServer() {
    const r = await fetch(POSTS_URL + '?action=posts&_=' + Date.now());
    const posts = await r.json();
    saveCachedPosts(posts);
    return posts;
  }

  const cached = loadCachedPosts();
  if (cached) {
    // paint what we know instantly, then quietly fetch the true newest list
    fromServer().then(fresh => onUpdate && onUpdate(filter(fresh))).catch(() => {});
    return filter(cached);
  }

  // first time here → one honest fetch is all it takes
  return filter(await fromServer());
}

// ---- small helpers ----

// safe text → html (so visitors' words can't break a page)
function esc(s) {
  const d = document.createElement('div');
  d.textContent = s || '';
  return d.innerHTML;
}

// split a long writing into paragraphs (blank line = new paragraph)
function paraBlocks(t) {
  return (t || '').split(/\n\s*\n/).map(b => b.trim()).filter(Boolean);
}

function postHref(id) { return 'post.html?id=' + encodeURIComponent(id); }

// the nook calls this after publishing/taking-down, so the next
// page visit always shows the real, newest list (no stale memory)
function forgetPostsCache() {
  try { localStorage.removeItem(POSTS_CACHE_KEY); } catch (e) {}
}