// ===== app.js — shared data loader =====
// Loads data.json and returns a promise with the parsed data.

let _cache = null;

function loadData() {
  if (_cache) return Promise.resolve(_cache);
  return fetch('data.json')
    .then(r => {
      if (!r.ok) throw new Error('Could not load data.json');
      return r.json();
    })
    .then(data => {
      _cache = data;
      return data;
    })
    .catch(err => {
      console.error('loadData error:', err);
      // Return empty structure so pages don't crash
      return { brands: [], models: [], engines: {}, stages: {} };
    });
}
