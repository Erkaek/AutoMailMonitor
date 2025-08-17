// Inspect COM tree for a specific store
(async () => {
	const oc = require('../src/server/outlookConnector');
	const storeId = process.env.TARGET_STORE_ID || process.argv[2] || '';
	const tree = await oc.getAllStoresAndTreeViaCOM({ maxDepth: 10, forceRefresh: true });
	console.log('Stores:', tree.stores.map(s => ({ Name: s.Name, StoreID: s.StoreID })).slice(0, 10));
	const key = storeId || (tree.stores[0] && tree.stores[0].StoreID) || '';
	const map = tree.foldersByStore[key];
	if (!map) {
		console.log('No map for store key:', key.slice(0, 24));
		const alt = tree.stores.find(s => s.Name === key);
		if (alt) {
			console.log('Trying alt by name:', alt.StoreID.slice(0, 24));
			console.log('Has map?', !!tree.foldersByStore[alt.StoreID]);
		}
		process.exit(0);
	}
	const entries = Object.values(map);
	console.log('Count entries:', entries.length);
	console.log('Sample paths:', entries.slice(0, 10).map(e => e.path));
})();
