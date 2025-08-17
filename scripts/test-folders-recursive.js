// Quick integration probe to dump recursive folders for selected store
(async () => {
  const oc = require('../src/server/outlookConnector');
  const storeId = process.env.TARGET_STORE_ID || process.argv[2] || '';
  if (!storeId) { console.error('Usage: node scripts/test-folders-recursive.js <StoreID>'); process.exit(1); }
  const list = await oc.listFoldersRecursive(storeId, { maxDepth: Number(process.env.FOLDER_ENUM_MAX_DEPTH || -1) });
  console.log(JSON.stringify({ count: list.length, sample: list.slice(0, 10) }, null, 2));
})();
