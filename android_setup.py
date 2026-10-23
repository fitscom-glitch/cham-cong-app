"""Cấu hình thư mục android/ do `flutter create` sinh ra (chạy trong GitHub Actions).

- Vân tay / khuôn mặt (local_auth): FlutterFragmentActivity, quyền USE_BIOMETRIC, theme AppCompat
- Thông báo nhắc giờ (flutter_local_notifications): quyền, receiver, core library desugaring
"""
import pathlib
import re
import sys

APP = pathlib.Path("app/android/app")
MAIN = APP / "src" / "main"
DESUGAR = "com.android.tools:desugar_jdk_libs:2.1.4"


def fail(msg):
    sys.exit("LOI cau hinh Android: " + msg)


# 1. MainActivity -> FlutterFragmentActivity (bắt buộc cho local_auth)
activities = list(MAIN.rglob("MainActivity.*"))
if not activities:
    fail("khong tim thay MainActivity")
for f in activities:
    f.write_text(f.read_text().replace("FlutterActivity", "FlutterFragmentActivity"))

# 2. AndroidManifest: quyền + receiver cho thông báo hẹn giờ
manifest = MAIN / "AndroidManifest.xml"
text = manifest.read_text()
permissions = [
    '<uses-permission android:name="android.permission.USE_BIOMETRIC"/>',
    '<uses-permission android:name="android.permission.POST_NOTIFICATIONS"/>',
    '<uses-permission android:name="android.permission.VIBRATE"/>',
    '<uses-permission android:name="android.permission.RECEIVE_BOOT_COMPLETED"/>',
    '<uses-permission android:name="android.permission.USE_EXACT_ALARM"/>',
    '<uses-permission android:name="android.permission.SCHEDULE_EXACT_ALARM" android:maxSdkVersion="32"/>',
]
missing = [p for p in permissions if p.split('"')[1] not in text]
if missing:
    text = text.replace("<application", "".join(p + "\n    " for p in missing) + "<application", 1)

if "ScheduledNotificationReceiver" not in text:
    receivers = (
        '<receiver android:exported="false" '
        'android:name="com.dexterous.flutterlocalnotifications.ScheduledNotificationReceiver" />\n'
        '        <receiver android:exported="false" '
        'android:name="com.dexterous.flutterlocalnotifications.ScheduledNotificationBootReceiver">\n'
        "            <intent-filter>\n"
        '                <action android:name="android.intent.action.BOOT_COMPLETED"/>\n'
        '                <action android:name="android.intent.action.MY_PACKAGE_REPLACED"/>\n'
        '                <action android:name="android.intent.action.QUICKBOOT_POWERON"/>\n'
        '                <action android:name="com.htc.intent.action.QUICKBOOT_POWERON"/>\n'
        "            </intent-filter>\n"
        "        </receiver>\n"
        "    "
    )
    if "</application>" not in text:
        fail("manifest khong co </application>")
    text = text.replace("</application>", receivers + "</application>", 1)
manifest.write_text(text)

# 3. Theme AppCompat (tránh lỗi của hộp thoại sinh trắc học trên Android cũ)
for f in MAIN.glob("res/values*/styles.xml"):
    t = f.read_text()
    t = re.sub(r'parent="@android:style/Theme\.[A-Za-z.]*"', 'parent="Theme.AppCompat.DayNight.NoActionBar"', t)
    f.write_text(t)

# 4. Core library desugaring (bắt buộc cho flutter_local_notifications)
kts = APP / "build.gradle.kts"
groovy = APP / "build.gradle"
if kts.exists():
    t = kts.read_text()
    if "isCoreLibraryDesugaringEnabled" not in t:
        t, n = re.subn(r"(compileOptions\s*\{)", r"\1\n        isCoreLibraryDesugaringEnabled = true", t, count=1)
        if n == 0:
            fail("khong tim thay compileOptions trong build.gradle.kts")
        t += '\ndependencies {\n    coreLibraryDesugaring("%s")\n}\n' % DESUGAR
        kts.write_text(t)
elif groovy.exists():
    t = groovy.read_text()
    if "coreLibraryDesugaringEnabled" not in t:
        t, n = re.subn(r"(compileOptions\s*\{)", r"\1\n        coreLibraryDesugaringEnabled true", t, count=1)
        if n == 0:
            fail("khong tim thay compileOptions trong build.gradle")
        t += "\ndependencies {\n    coreLibraryDesugaring '%s'\n}\n" % DESUGAR
        groovy.write_text(t)
else:
    fail("khong tim thay build.gradle(.kts)")

print("Da cau hinh Android: van tay, thong bao, desugaring")
