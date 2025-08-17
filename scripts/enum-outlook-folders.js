// WSH JScript (run with cscript.exe //nologo) to enumerate Outlook stores and folders
// Usage: cscript //nologo enum-outlook-folders.js [StoreName] [MaxDepth]
// Outputs JSON to stdout

function toLower(s) { return (s || "").toLowerCase(); }
function endsWithBS(p) { return p && p.length && p.charAt(p.length-1) === '\\'; }

function echo(s) { WScript.Echo(s); }

function main() {
  var args = WScript.Arguments;
  var targetName = args.length > 0 ? args.Item(0) : "";
  var maxDepth = -1;
  try { if (args.length > 1) maxDepth = parseInt(args.Item(1), 10); } catch (e) {}
  if (!(maxDepth >= 0)) maxDepth = -1;

  var result = { success: false, error: null, folders: [] };
  try {
    var app = new ActiveXObject("Outlook.Application");
    var ns = app.Session; // GetNamespace("MAPI")
    var stores = ns.Stores;
    var store = null;
    var i;
    // Pick target store by DisplayName
    if (targetName && targetName.length) {
      for (i = 1; i <= stores.Count; i++) {
        var s = stores.Item(i);
        try {
          var name = s.DisplayName;
          if (toLower(name) === toLower(targetName) || toLower(name).indexOf(toLower(targetName)) >= 0) {
            store = s; break;
          }
        } catch (e1) {}
      }
    }
    if (!store) { store = ns.DefaultStore; }

    var storeName = ""; try { storeName = String(store.DisplayName); } catch (e2) {}
    var storeId = ""; try { storeId = String(store.StoreID); } catch (e3) {}

    var root = null;
    try { root = store.GetRootFolder(); } catch (e4) {}

    function getCountSafely(fldrs) { try { return fldrs.Count; } catch (e9) { return 0; } }

    // Some profiles expose a different visual root via Namespace.Folders
    try {
      if (root === null || getCountSafely(root.Folders) === 0) {
        var topLevel = ns.Folders;
        for (i = 1; i <= topLevel.Count; i++) {
          var tf = topLevel.Item(i);
          try { if (tf.Store && tf.Store.StoreID === storeId) { root = tf; break; } } catch (e5) {}
        }
      }
    } catch (e6) {}

    var all = [];
    function addFolder(folder, parentPath, depth) {
      if (maxDepth >= 0 && depth > maxDepth) return;
      var name = ""; try { name = String(folder.Name); } catch (e7) {}
      if (!name) return;
      var fullPath = parentPath ? (parentPath + "\\" + name) : (storeName + "\\" + name);
      var entryId = ""; try { entryId = String(folder.EntryID); } catch (e8) {}
      var childCount = getCountSafely(folder.Folders);
      all.push({ StoreDisplayName: storeName, StoreEntryID: storeId, FolderName: name, FolderEntryID: entryId, FullPath: fullPath, ChildCount: childCount });
      if (childCount > 0) {
        var subs = folder.Folders;
        for (var j = 1; j <= subs.Count; j++) {
          try { addFolder(subs.Item(j), fullPath, depth + 1); } catch (e10) {}
        }
      }
    }

    if (root) {
      // Enumerate children of root as top-level; include root itself children only
      var top = root.Folders;
      if (getCountSafely(top) > 0) {
        for (i = 1; i <= top.Count; i++) {
          try { addFolder(top.Item(i), storeName, 0); } catch (e11) {}
        }
      } else {
        // As a last resort, try one extra level deeper
        for (i = 1; i <= top.Count; i++) {
          var mid = top.Item(i);
          try {
            var mids = mid.Folders;
            for (var k = 1; k <= mids.Count; k++) {
              try { addFolder(mids.Item(k), storeName, 0); } catch (e12) {}
            }
          } catch (e13) {}
        }
      }
    }

    result.success = true;
    result.folders = all;
  } catch (err) {
    result.success = false;
    result.error = String(err && (err.message || err.description || err.toString())) || 'Unknown error';
  }
  try { echo(JSON.stringify(result)); } catch (ejson) { echo('{"success":false,"error":"json"}'); }
}

main();
