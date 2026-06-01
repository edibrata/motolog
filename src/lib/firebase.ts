import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getStorage } from "firebase/storage";

const firebaseConfig = {
  apiKey: "AIzaSyC6IXCk6EXm0ndord8Is6cRZf_mG2MY5UM",
  authDomain: "kedinasan-d9051.firebaseapp.com",
  projectId: "kedinasan-d9051",
  storageBucket: "kedinasan-d9051.firebasestorage.app",
  messagingSenderId: "488517479224",
  appId: "1:488517479224:web:57ed244cf69e2a010c6fcd"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const storage = getStorage(app);
