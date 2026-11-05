/* Shared configuration for the whole repository.
   Edit ONLY this file to change credentials or Firebase. */
window.APP_CONFIG = {

  /* ------- Firebase (same project as the original admin panel) ------- */
  firebase: {
    apiKey: "AIzaSyAQbnBB0FxycWtF6B3gKWfG4ZqijNZU9Vw",
    authDomain: "ar-calculator-38781.firebaseapp.com",
    databaseURL: "https://ar-calculator-38781-default-rtdb.firebaseio.com",
    projectId: "ar-calculator-38781",
    storageBucket: "ar-calculator-38781.firebasestorage.app",
    messagingSenderId: "312708241196",
    appId: "1:312708241196:web:b1d73cd31dd267c6627e89"
  },

  /* ------- Admin credentials ------- */
  admin: {
    username: "admin",
    /* SHA-256 hash of "admin123".
       To change password: run  python python/hash_gen.py  and paste result here. */
    passwordHash: "240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9"
  },

  /* ------- Session ------- */
  session: {
    key: "stt_admin_session",
    ttlMs: 2 * 60 * 60 * 1000     /* 2 hours */
  },

  /* ------- Firebase paths ------- */
  paths: {
    dataSharing: "control/data_sharing_on",
    bookings:    "bookings",
    reviews:     "reviews"
  }
};
