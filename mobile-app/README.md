# ResumeFit Pro Mobile App

This folder contains the native mobile wrapper for ResumeFit Pro.

The app loads the production web app:

```text
https://resumefit-pro.onrender.com
```

## Android Play Store Build

Requirements:

- Node.js
- Java JDK 17
- Android Studio
- Android SDK
- Google Play Developer account

Commands:

```bash
npm install
npx cap add android
npx cap sync android
npx cap open android
```

In Android Studio:

1. Open **Build** > **Generate Signed Bundle / APK**.
2. Select **Android App Bundle**.
3. Create or select a keystore.
4. Build the `.aab`.
5. Upload the `.aab` in Google Play Console.

## iOS App Store Build

Requirements:

- macOS
- Xcode
- Apple Developer account

Commands on Mac:

```bash
npm install
npx cap add ios
npx cap sync ios
npx cap open ios
```

In Xcode:

1. Set signing team and bundle ID.
2. Archive the app.
3. Upload to App Store Connect.

## Store Notes

Because this is resume software, prepare:

- Privacy policy URL
- Support email
- Screenshots
- App icon assets
- Data handling disclosure
- Gemini/API disclosure
