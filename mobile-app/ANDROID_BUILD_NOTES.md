# Android Build Notes

The Android native project has been generated in:

```text
mobile-app/android
```

The local debug build could not complete on this machine because Android SDK is not installed/configured:

```text
SDK location not found. Define ANDROID_HOME or android/local.properties.
```

## Fix

Install Android Studio, then install:

- Android SDK Platform
- Android SDK Build-Tools
- Android SDK Platform-Tools

Then create:

```text
mobile-app/android/local.properties
```

Example:

```properties
sdk.dir=C:\\Users\\Sumit\\AppData\\Local\\Android\\Sdk
```

Then run:

```bat
cd mobile-app\\android
.\\gradlew.bat assembleDebug
```

For Play Store:

```text
Android Studio > Build > Generate Signed Bundle / APK > Android App Bundle
```
