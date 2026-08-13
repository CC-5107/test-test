// Firebase initialization
// ---------------------------------------------------------------------------
// 1. Go to https://console.firebase.google.com, signed in with the Google
//    account you want this new project to live under, and create a project.
// 2. In Project settings > General > Your apps, register a Web app.
// 3. Copy the config values Firebase gives you into firebaseConfig below,
//    replacing every "REPLACE_ME_..." placeholder.
// 4. Enable Authentication > Sign-in method > Email/Password.
// 5. Enable Firestore Database (start in production mode) and set security
//    rules so only signed-in users can read/write (see the setup guide for
//    the exact "usernames" collection rules this app needs).
// ---------------------------------------------------------------------------

import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyASAbcAxmI1HHo6r1-FR3N-NEWdCzrVKGo",
  authDomain: "cerpsrldashboard.firebaseapp.com",
  projectId: "cerpsrldashboard",
  storageBucket: "cerpsrldashboard.firebasestorage.app",
  messagingSenderId: "1033966789680",
  appId: "1:1033966789680:web:273bc5fe6bc096538579f5",
  measurementId: "G-DPHFHPSVN0"
};

const app = initializeApp(firebaseConfig);

export const auth = getAuth(app);
export const db = getFirestore(app);
export default app;
