// Self-check for the CloudStorage mirror: settings live in localStorage but are
// synced through Telegram's per-user store, so a second device starts out
// configured. Three things can go wrong quietly, and each is checked here.
// The existing local install must migrate itself upwards on first run, the
// cloud copy must win over the local one when both exist, and a write that
// fails must come back later instead of being dropped.
// Run: node test/cloud-sync.test.js
const fs = require('fs');
const vm = require('vm');

const APP = fs.readFileSync(__dirname + '/../app.js', 'utf8');
// config.js is deployment-specific and never committed; app.js just expects it.
const PRELUDE = "const GAS_URL = 'https://example.invalid/exec';\nconst APP_ID = 'rasshody';\n";

/** Storage API stand-in: only length/key/getItem/setItem/removeItem are used. */
const makeStorage = (init = {}) => {
  const data = new Map(Object.entries(init));
  return {
    data,
    get length() { return data.size; },
    key: i => [...data.keys()][i],
    getItem: k => (data.has(k) ? data.get(k) : null),
    setItem: (k, v) => data.set(k, String(v)),
    removeItem: k => { data.delete(k); },
  };
};

/** CloudStorage stand-in. `fail` makes every write report an error. */
const makeCloud = (init = {}) => {
  const data = new Map(Object.entries(init));
  const c = {
    data, fail: false, removed: [],
    getKeys: cb => cb(null, [...data.keys()]),
    getItems: (keys, cb) => cb(null, Object.fromEntries(keys.map(k => [k, data.get(k) ?? '']))),
    setItem: (k, v, cb) => (c.fail ? cb(new Error('offline')) : (data.set(k, v), cb(null, true))),
    removeItem: (k, cb) => { c.removed.push(k); data.delete(k); cb(null, true); },
  };
  return c;
};

/** Boot app.js against fake storage and a fake client, and let init() settle. */
async function boot({ local = {}, cloudData = {}, version = '6.9' } = {}) {
  const storage = makeStorage(Object.fromEntries(
    Object.entries(local).map(([k, v]) => ['rasshody:' + k, v])));
  const cloud = version ? makeCloud(cloudData) : null;
  const timers = [];
  const stubEl = { style: {}, classList: { contains: () => false, add() {}, remove() {} } };

  const sandbox = {
    window: {
      addEventListener() {},
      Telegram: version ? { WebApp: {
        isVersionAtLeast: v => version >= v,
        CloudStorage: cloud,
        ready() {}, expand() {}, initDataUnsafe: {}, BackButton: { onClick() {} },
      } } : undefined,
    },
    document: { getElementById: () => stubEl, documentElement: {}, querySelectorAll: () => [] },
    localStorage: storage,
    TextEncoder,
    console,
    setTimeout: fn => timers.push(fn),
  };
  vm.createContext(sandbox);
  vm.runInContext(PRELUDE + APP, sandbox);
  // init() only awaits callbacks the fakes answer synchronously, so one macro
  // task is enough to drain the microtasks it queued.
  await new Promise(r => setImmediate(r));

  return {
    cloud, timers, storage,
    S: vm.runInContext('S', sandbox),
    ls: k => storage.getItem('rasshody:' + k),
    set: (k, v) => vm.runInContext(`lsSet(${JSON.stringify(k)}, BIG_OR_VALUE)`,
      Object.assign(sandbox, { BIG_OR_VALUE: v })),
  };
}

let failed = 0;
const check = (label, cond) => { if (!cond) { console.error('FAIL: ' + label); failed++; } };

(async () => {
  // An install that predates the mirror pushes what it has upwards, untouched.
  {
    const { cloud, ls } = await boot({ local: { spreadsheetId: 'SHEET1', userName: 'kost' } });
    check('existing local settings migrate into the cloud',
      cloud.data.get('spreadsheetId') === 'SHEET1' && cloud.data.get('userName') === 'kost');
    check('migration leaves localStorage alone', ls('spreadsheetId') === 'SHEET1');
  }

  // Another device already configured the account. Its copy is the agreed one.
  {
    const { S, ls } = await boot({
      local: { spreadsheetId: 'STALE' },
      cloudData: { spreadsheetId: 'FRESH', userName: 'kost' },
    });
    check('cloud value wins over the local one', S.spreadsheetId === 'FRESH');
    check('cloud value is written back to localStorage', ls('spreadsheetId') === 'FRESH');
    check('cloud-only key reaches the state', S.userName === 'kost');
  }

  // A key the cloud does not answer for must not blank out the local value.
  {
    const { S } = await boot({
      local: { userName: 'kost' },
      cloudData: { spreadsheetId: 'FRESH' },
    });
    check('an empty cloud answer does not erase a local key', S.userName === 'kost');
  }

  // Past 4096 bytes Telegram refuses the write, and any older copy up there
  // would come back down over this device's newer value on the next sync.
  {
    const { cloud, set } = await boot({ cloudData: { catCache: '{"outflow":["old"]}' } });
    set('catCache', JSON.stringify({ outflow: ['x'.repeat(5000)] }));
    check('an oversized value is dropped from the cloud instead of stored',
      !cloud.data.has('catCache') && cloud.removed.includes('catCache'));
  }

  // Offline writes are queued, and the retry sends the value as it stands then,
  // not the one that happened to fail.
  {
    const { cloud, timers, set } = await boot({ local: { userName: 'kost' } });
    cloud.fail = true;
    set('userName', 'first');
    set('userName', 'second');
    check('a failed write does not reach the cloud', cloud.data.get('userName') === 'kost');
    check('failed writes share one retry timer', timers.length === 1);
    cloud.fail = false;
    timers[0]();
    check('the retry sends the current value', cloud.data.get('userName') === 'second');
  }

  // Old clients and plain browsers have no CloudStorage at all.
  {
    const { S } = await boot({ local: { userName: 'kost' }, version: '6.0' });
    check('an old client still reads its local settings', S.userName === 'kost');
  }

  console.log(failed ? `${failed} check(s) failed` : 'cloud-sync: all checks passed');
  process.exit(failed ? 1 : 0);
})();
