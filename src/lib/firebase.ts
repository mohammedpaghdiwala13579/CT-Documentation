import { initializeApp, getApps, getApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import importedConfig from "../../firebase-applet-config.json";

// Permanent Comilla Traders Firebase credentials tied to comillatraders@gmail.com
// Ensures that whenever this code is pushed to GitHub or opened from any other account,
// it stays permanently connected to the comillatraders@gmail.com database.
export const COMILLA_FIREBASE_CONFIG = {
  projectId: "gifted-gift-rdpgw",
  appId: "1:102803571663:web:b97a4b94273ec54269c8ba",
  apiKey: "AIzaSyD8QK5nJ2TVllAjCeMojOD9wZPBJAeuy4E",
  authDomain: "gifted-gift-rdpgw.firebaseapp.com",
  firestoreDatabaseId: "ai-studio-2c592343-56ab-4d40-a2ac-d15fed703e91",
  storageBucket: "gifted-gift-rdpgw.firebasestorage.app",
  messagingSenderId: "102803571663",
  measurementId: ""
};

// Always enforce the comillatraders@gmail.com project and database ID
const effectiveConfig = {
  ...COMILLA_FIREBASE_CONFIG,
  ...(importedConfig?.projectId === COMILLA_FIREBASE_CONFIG.projectId ? importedConfig : {})
};

// Initialize Firebase App
export const app = getApps().length > 0 ? getApp() : initializeApp(effectiveConfig);

// Initialize Firestore pointing directly to the comillatraders@gmail.com database
export const db = getFirestore(app, effectiveConfig.firestoreDatabaseId || "(default)");
