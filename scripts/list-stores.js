// Dump Outlook stores via connector to help pick a StoreID quickly
(async () => {
  const oc = require('../src/server/outlookConnector');
  try {
    const boxes = await oc.getMailboxes();
    console.log(JSON.stringify(boxes, null, 2));
  } catch (e) {
    console.error('Error listing stores:', e.message);
  }
})();
