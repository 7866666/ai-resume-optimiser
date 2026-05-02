# ResumeFit Pro App Distribution

ResumeFit Pro is already installable as a Progressive Web App from the live website:

```text
https://resumefit-pro.onrender.com
```

## Install Without App Stores

### Android

1. Open the website in Chrome.
2. Tap **Install app** or Chrome menu.
3. Tap **Add to Home screen**.

### iPhone / iPad

1. Open the website in Safari.
2. Tap **Share**.
3. Tap **Add to Home Screen**.

### Windows / Mac

1. Open the website in Chrome or Edge.
2. Click the install icon in the address bar.
3. Confirm installation.

## Play Store

To publish on Google Play Store, wrap the PWA as an Android app using Trusted Web Activity.

Recommended tools:

- Bubblewrap
- Android Studio
- Google Play Console

Required:

- Google Play Developer account
- App name, icon, screenshots, privacy policy
- Signed Android App Bundle `.aab`

## Apple App Store

To publish on iPhone/iPad App Store, wrap the web app using Capacitor or a native WebView shell.

Required:

- Apple Developer account
- Xcode on macOS
- App Store screenshots
- Privacy policy
- App review approval

## Desktop Apps

For Windows, macOS, and Linux desktop builds, package the web app using Electron or Tauri.

Recommended:

- Electron for easiest packaging
- Tauri for smaller app size

## Important Store Requirements

Because users upload resumes, add these before public paid distribution:

- Privacy policy page
- Terms page
- Data retention statement
- Support email
- Clear API usage disclosure
- Optional login/payment layer if monetizing
