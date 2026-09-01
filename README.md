# Daily Ledger — Android App

Local-only expense tracker + "who owes me money" tracker. No sign-in, no cloud —
everything is stored on your phone in the app's local storage.

## What's inside
- A React web app (src/App.jsx) — same Apple-style design as the web version
- Wrapped as a native Android app using Capacitor (the `android/` folder)
- Exports to Excel (.xlsx) from within the app anytime

## Building the APK

You need Android Studio installed (free): https://developer.android.com/studio

1. Open a terminal in this folder, run:
   npm install
   npm run build
   npx cap sync android

2. Open Android Studio → "Open" → select the `android` folder inside this project.

3. Let it finish indexing/syncing Gradle (first time takes a few minutes, downloads
   some Android build tools — needs internet).

4. Build menu → Build Bundle(s) / APK(s) → Build APK(s).

5. Once done, Android Studio shows a notification with "locate" — click it, or find
   the file at: android/app/build/outputs/apk/debug/app-debug.apk

6. Copy that .apk to your phone (email it to yourself, or use a USB cable/Google Drive)
   and open it on the phone to install. You'll need to allow "Install from unknown
   sources" the first time — Android will prompt you for this automatically.

## Making changes later
If you edit App.jsx again, just re-run steps 1's three commands, then in Android
Studio: Build → Build APK(s) again. No need to redo `cap add android`.

## Data & safety
- Data is stored in the Android app's private local storage — not shared with other
  apps, not sent anywhere. It's tied to this specific app install.
- It survives normal phone use, restarts, etc. It's only lost if you uninstall the
  app or manually clear its storage in Android Settings → Apps → Daily Ledger → Storage → Clear Data.
- Use the Export button regularly to back up to an Excel file if you want a copy
  outside the app.
