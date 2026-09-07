import { initializeApp } from "firebase/app";
import { getFirestore, doc, getDocFromServer } from "firebase/firestore";
import firebaseConfig from "../../firebase-applet-config.json";

// Initialize Firebase using the project applet configuration
const app = initializeApp(firebaseConfig);

// Initialize Firestore with custom database ID from configuration
export const db = getFirestore(app, firebaseConfig.firestoreDatabaseId || "(default)");

// Validate connection to Firestore on initial boot
async function testConnection() {
  try {
    await getDocFromServer(doc(db, "test", "connection"));
  } catch (error) {
    if (error instanceof Error && error.message.includes("the client is offline")) {
      console.warn("Please check your Firebase configuration.");
    }
  }
}
testConnection();
