/**
 * Connecteur COM moderne utilisant FFI-NAPI
 * Remplace WinAX pour une meilleure compatibilité avec Node.js
 */

const ffi = require('ffi-napi');
const ref = require('ref-napi');
const Struct = require('ref-struct-napi');
const ArrayType = require('ref-array-napi');

// Types COM de base
const VARIANT = ref.types.void; // Simplifié pour commencer
const BSTR = ref.types.CString;
const HRESULT = ref.types.long;
const IUnknown = ref.refType(ref.types.void);

class COMConnector {
  constructor() {
    this.isInitialized = false;
    this.outlookApp = null;
    this.namespace = null;
    
    // Charger les DLLs Windows nécessaires
    this.ole32 = null;
    this.oleaut32 = null;
    
    this.initializeCOM();
  }

  initializeCOM() {
    try {
      console.log('🔧 [COM] Initialisation COM avec FFI-NAPI...');
      
      // Charger OLE32.dll pour COM
      this.ole32 = ffi.Library('ole32', {
        'CoInitializeEx': [HRESULT, [ref.types.void, ref.types.ulong]],
        'CoUninitialize': [ref.types.void, []],
        'CoCreateInstance': [HRESULT, [
          ref.refType(ref.types.void), // CLSID
          IUnknown,                    // pUnkOuter
          ref.types.ulong,             // dwClsContext
          ref.refType(ref.types.void), // IID
          ref.refType(IUnknown)        // ppv
        ]],
        'CLSIDFromProgID': [HRESULT, [BSTR, ref.refType(ref.types.void)]],
        'IIDFromString': [HRESULT, [BSTR, ref.refType(ref.types.void)]]
      });

      // Charger OLEAUT32.dll pour les types VARIANT
      this.oleaut32 = ffi.Library('oleaut32', {
        'SysAllocString': [BSTR, [ref.types.CString]],
        'SysFreeString': [ref.types.void, [BSTR]],
        'VariantInit': [ref.types.void, [ref.refType(VARIANT)]],
        'VariantClear': [HRESULT, [ref.refType(VARIANT)]]
      });

      // Initialiser COM
      const COINIT_APARTMENTTHREADED = 0x2;
      const hr = this.ole32.CoInitializeEx(ref.NULL, COINIT_APARTMENTTHREADED);
      
      if (hr < 0 && hr !== -2147417850) { // S_FALSE = déjà initialisé
        throw new Error(`Erreur initialisation COM: 0x${hr.toString(16)}`);
      }

      this.isInitialized = true;
      console.log('✅ [COM] COM initialisé avec succès');
      
    } catch (error) {
      console.error('❌ [COM] Erreur initialisation COM:', error.message);
      throw error;
    }
  }

  async connectToOutlook() {
    try {
      console.log('🔗 [COM] Connexion à Outlook...');
      
      if (!this.isInitialized) {
        throw new Error('COM non initialisé');
      }

      // CLSID pour Outlook.Application
      const outlookCLSID = Buffer.alloc(16);
      const progID = this.oleaut32.SysAllocString('Outlook.Application');
      
      let hr = this.ole32.CLSIDFromProgID(progID, outlookCLSID);
      this.oleaut32.SysFreeString(progID);
      
      if (hr < 0) {
        throw new Error(`Erreur récupération CLSID Outlook: 0x${hr.toString(16)}`);
      }

      // IID pour IDispatch
      const dispatchIID = Buffer.alloc(16);
      const iidString = this.oleaut32.SysAllocString('{00020400-0000-0000-C000-000000000046}');
      
      hr = this.ole32.IIDFromString(iidString, dispatchIID);
      this.oleaut32.SysFreeString(iidString);
      
      if (hr < 0) {
        throw new Error(`Erreur récupération IID IDispatch: 0x${hr.toString(16)}`);
      }

      // Créer l'instance Outlook
      const CLSCTX_LOCAL_SERVER = 0x4;
      const outlookPtr = ref.alloc(IUnknown);
      
      hr = this.ole32.CoCreateInstance(
        outlookCLSID,
        ref.NULL,
        CLSCTX_LOCAL_SERVER,
        dispatchIID,
        outlookPtr
      );

      if (hr < 0) {
        throw new Error(`Erreur création instance Outlook: 0x${hr.toString(16)}`);
      }

      this.outlookApp = outlookPtr.deref();
      console.log('✅ [COM] Connecté à Outlook via COM');
      
      return true;
      
    } catch (error) {
      console.error('❌ [COM] Erreur connexion Outlook:', error.message);
      throw error;
    }
  }

  async getNamespace() {
    try {
      if (!this.outlookApp) {
        throw new Error('Outlook non connecté');
      }

      // Pour l'instant, retourner un placeholder
      // L'implémentation complète nécessiterait d'appeler GetNamespace via IDispatch
      console.log('📂 [COM] Récupération namespace MAPI...');
      
      // Simulation pour test
      this.namespace = { connected: true };
      return this.namespace;
      
    } catch (error) {
      console.error('❌ [COM] Erreur récupération namespace:', error.message);
      throw error;
    }
  }

  async getFolders() {
    try {
      if (!this.namespace) {
        await this.getNamespace();
      }

      console.log('📁 [COM] Récupération dossiers...');
      
      // Pour l'instant, retourner des données de test
      // L'implémentation complète nécessiterait d'appeler les méthodes COM
      return [
        {
          Name: 'Boîte de réception',
          FolderPath: 'Boîte de réception',
          Count: 0,
          UnreadItemCount: 0,
          SubFolders: []
        }
      ];
      
    } catch (error) {
      console.error('❌ [COM] Erreur récupération dossiers:', error.message);
      throw error;
    }
  }

  async getFolderByPath(folderPath) {
    try {
      console.log(`📂 [COM] Recherche dossier: ${folderPath}`);
      
      // Implémentation simplifiée pour test
      const folders = await this.getFolders();
      return folders.find(f => f.FolderPath === folderPath) || null;
      
    } catch (error) {
      console.error('❌ [COM] Erreur recherche dossier:', error.message);
      return null;
    }
  }

  async getEmails(folderPath, maxCount = 100) {
    try {
      console.log(`📧 [COM] Récupération emails de: ${folderPath}`);
      
      const folder = await this.getFolderByPath(folderPath);
      if (!folder) {
        throw new Error(`Dossier non trouvé: ${folderPath}`);
      }

      // Pour l'instant, retourner un tableau vide
      // L'implémentation complète nécessiterait d'itérer sur les Items du dossier
      return [];
      
    } catch (error) {
      console.error('❌ [COM] Erreur récupération emails:', error.message);
      throw error;
    }
  }

  isConnected() {
    return this.isInitialized && this.outlookApp !== null;
  }

  cleanup() {
    try {
      console.log('🧹 [COM] Nettoyage COM...');
      
      if (this.outlookApp) {
        // Libérer l'interface Outlook
        this.outlookApp = null;
      }

      if (this.isInitialized && this.ole32) {
        this.ole32.CoUninitialize();
        this.isInitialized = false;
      }

      console.log('✅ [COM] Nettoyage terminé');
      
    } catch (error) {
      console.error('❌ [COM] Erreur nettoyage:', error.message);
    }
  }
}

module.exports = COMConnector;
