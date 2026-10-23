import 'dart:async';
import 'dart:convert';
import 'dart:io';
import 'dart:math' as math;
import 'dart:ui' show FontFeature;

import 'package:excel/excel.dart' as xl;
import 'package:file_selector/file_selector.dart' show XTypeGroup, openFile;
import 'package:flutter_local_notifications/flutter_local_notifications.dart'
    show
        AndroidFlutterLocalNotificationsPlugin,
        AndroidInitializationSettings,
        AndroidNotificationCategory,
        AndroidNotificationDetails,
        AndroidScheduleMode,
        AudioAttributesUsage,
        FlutterLocalNotificationsPlugin,
        Importance,
        InitializationSettings,
        NotificationDetails,
        NotificationVisibility,
        Priority,
        UriAndroidNotificationSound;
import 'package:flutter/material.dart' hide Text;
import 'package:flutter/material.dart' as mat show Text;
import 'package:flutter/services.dart' show Clipboard, ClipboardData, rootBundle;
import 'package:intl/intl.dart';
import 'package:local_auth/local_auth.dart' show AuthenticationOptions, LocalAuthentication;
import 'package:path/path.dart' as p;
import 'package:path_provider/path_provider.dart';
import 'package:pdf/pdf.dart' as pdf;
import 'package:pdf/widgets.dart' as pw;
import 'package:share_plus/share_plus.dart';
import 'package:shared_preferences/shared_preferences.dart';
import 'package:sqflite/sqflite.dart';
import 'package:timezone/timezone.dart' as tz;

Future<void> main() async {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(const SalaryApp());
}

// ════════════════════════════════════════════════════════════════════════════
//  Hằng số & tiện ích
// ════════════════════════════════════════════════════════════════════════════

// ── Đa ngôn ngữ (Tiếng Việt / English) ────────────────────────────────────
// Chuỗi gốc viết bằng tiếng Việt. Khi chọn English, Text và tr() tra bảng dịch
// bên dưới (khớp nguyên câu hoặc theo mẫu cho câu có số liệu).

String currentLang = 'vi';
const supportedLangs = {'vi': 'Tiếng Việt', 'en': 'English'};

List<String> get _weekdayShort => currentLang == 'en'
    ? const ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    : const ['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN'];

List<String> get _weekdayFull => currentLang == 'en'
    ? const ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    : const ['Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy', 'Chủ nhật'];

/// Tên loại ngày để hiển thị (giá trị lưu trữ vẫn là 'Thường' / 'OT' / 'Lễ').
String dayLabel(String type) {
  if (currentLang != 'en') return type;
  return switch (type) {
    'Lễ' => 'Holiday',
    'OT' => 'Weekend',
    _ => 'Weekday',
  };
}

const _enExact = <String, String>{
  'Theo dõi lương': 'Pay Tracker',
  'Xuất dữ liệu': 'Export data',
  'Cài đặt': 'Settings',
  'Thêm ca': 'Add shift',
  'Thêm ca làm': 'Add shift',
  'Sửa ca làm': 'Edit shift',
  'Bắt đầu ca làm': 'Start a shift',
  'Hôm nay': 'Today',
  'Tuần này': 'This week',
  'Tháng này': 'This month',
  'Tháng trước': 'Last month',
  'Năm nay': 'This year',
  'Tất cả': 'All',
  'Lịch sử ca làm': 'Shift history',
  'Thu nhập tuần này': 'This week’s earnings',
  'Vào ca': 'Start shift',
  'Ra ca': 'End shift',
  'Hủy ca': 'Cancel shift',
  'ĐANG TRONG CA': 'ON SHIFT',
  'Hủy': 'Cancel',
  'Lưu': 'Save',
  'Sửa': 'Edit',
  'Xóa': 'Delete',
  'Thử lại': 'Retry',
  'Đóng': 'Close',
  'Sao chép': 'Copy',
  'Xuất CSV': 'Export CSV',
  'Đã sao chép CSV.': 'CSV copied.',
  'Xuất Excel (.xlsx) và PDF chỉ có trong bản app đầy đủ (main.dart).':
      'Excel (.xlsx) and PDF export is only in the full app (main.dart).',
  'Lương theo giờ': 'Hourly rate',
  'Lương mặc định (mỗi giờ)': 'Default hourly rate',
  'Mức lương không hợp lệ.': 'Invalid hourly rate.',
  'Giờ ra phải khác giờ vào.': 'End time must differ from start time.',
  'Giờ ra phải sau giờ vào.': 'End time must be after start time.',
  'Một ca tối đa 24 giờ.': 'A shift can be at most 24 hours.',
  'Ca này trùng giờ với một ca đã có.': 'This shift overlaps an existing shift.',
  'Ca làm cần ít nhất 1 phút. Bấm "Hủy ca" nếu bạn bấm nhầm.':
      'A shift must last at least 1 minute. Tap "Cancel shift" if you tapped by mistake.',
  'Hủy ca đang làm?': 'Cancel the current shift?',
  'Ca này sẽ không được lưu.': 'This shift will not be saved.',
  'Xóa ca làm?': 'Delete this shift?',
  'Đã lưu.': 'Saved.',
  'Đã lưu ca làm.': 'Shift saved.',
  'Đã lưu. Ca dài bất thường, hãy kiểm tra lại giờ ra (chạm vào ca để sửa).':
      'Saved. This shift is unusually long. Check the end time (tap the shift to edit).',
  'Đã xóa.': 'Deleted.',
  'Chưa có dữ liệu để xuất.': 'There is no data to export.',
  'Không có ca làm nào trong khoảng này.': 'No shifts in this period.',
  'Chưa có ca làm nào': 'No shifts yet',
  'Bấm “Vào ca” để chấm công hoặc thêm một ca thủ công.':
      'Tap “Start shift” to clock in, or add a shift manually.',
  'Thử lại.': 'Retry.',
  'Khoảng thời gian': 'Period',
  'Định dạng': 'Format',
  'Chi tiết từng ca và tổng hợp theo tuần': 'Each shift, plus weekly totals',
  'Báo cáo để in hoặc gửi (chạy offline)': 'A report to print or send (works offline)',
  'Văn bản thuần, mở được bằng Excel hoặc Google Sheets':
      'Plain text; opens in Excel or Google Sheets',
  'Ngày lễ (tính x2)': 'Holidays (paid x2)',
  'Chưa đánh dấu ngày lễ nào.': 'No holidays marked.',
  'Thêm ngày lễ': 'Add holiday',
  'Chọn ngày lễ': 'Select a holiday',
  'Đã đủ 40 giờ thường. Giờ làm thêm tính OT.':
      'Full 40 regular hours reached. Extra hours count as OT.',
  'Giờ thường trong tuần': 'Regular hours this week',
  'Thường': 'Regular',
  'Tổng giờ': 'Total hours',
  'Giờ thường': 'Regular hours',
  'Tổng tiền': 'Total pay',
  'Tiền': 'Pay',
  'Số ca': 'Shifts',
  'Tuần': 'Week',
  'Theo tuần': 'By week',
  'Chi tiết ca làm': 'Shift details',
  'Ngày': 'Date',
  'Thứ': 'Day',
  'Vào': 'In',
  'Ra': 'Out',
  'Loại': 'Type',
  'Giờ vào': 'Start',
  'Giờ ra': 'End',
  'Loại ngày': 'Day type',
  'Lương/giờ': 'Rate/hr',
  'Lương/giờ (USD)': 'Rate/hr (USD)',
  'Tiền (USD)': 'Pay (USD)',
  'Giờ OT (x1.5)': 'OT hours (x1.5)',
  'Giờ x2': 'x2 hours',
  'TỔNG': 'TOTAL',
  'Ca làm': 'Shifts',
  'Bảng lương': 'Payroll',
  'Không tạo được file Excel': 'Could not create the Excel file',
  // — Khóa ứng dụng —
  'Nhập mật khẩu để mở ứng dụng': 'Enter your password to open the app',
  'Mật khẩu': 'Password',
  'Mở khóa': 'Unlock',
  'Dùng vân tay / khuôn mặt': 'Use fingerprint / face',
  'Mật khẩu không đúng.': 'Incorrect password.',
  'Mở khóa ứng dụng': 'Unlock the app',
  'Xác nhận để bật mở khóa bằng sinh trắc học': 'Confirm to enable biometric unlock',
  'Bảo vệ ứng dụng bằng mật khẩu': 'Protect the app with a password',
  'Tạo mật khẩu để người khác không mở được dữ liệu chấm công của bạn.':
      'Create a password so that others cannot open your time records.',
  'Mật khẩu mới': 'New password',
  'Nhập lại mật khẩu': 'Confirm password',
  'Mật khẩu hiện tại': 'Current password',
  'Tạo mật khẩu': 'Create password',
  'Tắt khóa ứng dụng': 'Turn off app lock',
  'Bỏ qua, bật sau trong Cài đặt': 'Skip, enable later in Settings',
  'Mật khẩu cần ít nhất 4 ký tự.': 'The password needs at least 4 characters.',
  'Hai mật khẩu không giống nhau.': 'The two passwords do not match.',
  'Mật khẩu hiện tại không đúng.': 'The current password is incorrect.',
  'Nhập mật khẩu hiện tại.': 'Enter the current password.',
  'Xác nhận': 'Confirm',
  'Bảo mật & mật khẩu': 'Security & password',
  'Khóa ứng dụng bằng mật khẩu': 'Lock the app with a password',
  'Đổi mật khẩu': 'Change password',
  'Mở khóa bằng vân tay / khuôn mặt': 'Unlock with fingerprint / face',
  'Thiết bị chưa hỗ trợ hoặc chưa đăng ký vân tay / khuôn mặt.':
      'This device does not support fingerprint / face unlock, or none is enrolled.',
  'Không xác thực được nên chưa bật mở khóa sinh trắc học.':
      'Authentication failed, so biometric unlock was not enabled.',
  'Nếu quên mật khẩu, bạn phải gỡ và cài lại ứng dụng rồi nhập dữ liệu từ file sao lưu. Hãy xuất file sao lưu thường xuyên.':
      'If you forget the password, you must uninstall and reinstall the app, then import your data from a backup file. Export backups regularly.',
  'Đã bật khóa ứng dụng.': 'App lock enabled.',
  'Đã tắt khóa ứng dụng.': 'App lock disabled.',
  'Đã đổi mật khẩu.': 'Password changed.',
  // — Sao lưu & khôi phục —
  'Sao lưu & khôi phục dữ liệu': 'Backup & restore',
  'Sao lưu dữ liệu': 'Back up data',
  'Xuất file sao lưu (.json)': 'Export backup file (.json)',
  'Lưu toàn bộ ca làm và cài đặt để nhập lại khi cài lại ứng dụng.':
      'Saves all shifts and settings so you can restore them after reinstalling.',
  'Nhập từ file sao lưu': 'Import from backup file',
  'Nhập từ nội dung đã sao chép': 'Import from copied text',
  'Dán nội dung file sao lưu vào đây': 'Paste the backup file content here',
  'Dán': 'Paste',
  'Tiếp tục': 'Continue',
  'Nhập dữ liệu': 'Import data',
  'Chưa có nội dung để nhập.': 'There is nothing to import.',
  'Tệp sao lưu không hợp lệ.': 'The backup file is not valid.',
  'Đây không phải file sao lưu của ứng dụng này.': 'This is not a backup file from this app.',
  'File sao lưu quá mới, hãy cập nhật ứng dụng.': 'This backup is from a newer version. Please update the app.',
  'Dữ liệu trong file bị lỗi.': 'The data in the file is corrupted.',
  'Gộp vào dữ liệu hiện có': 'Merge with existing data',
  'Ca trùng giờ với ca đã có sẽ bị bỏ qua.': 'Shifts that overlap existing ones are skipped.',
  'Thay thế toàn bộ': 'Replace everything',
  'Xóa hết dữ liệu hiện tại và dùng dữ liệu trong file (kể cả cài đặt).':
      'Deletes the current data and uses the data in the file (including settings).',
  'Sao chép nội dung dưới đây và lưu vào nơi an toàn (ghi chú, email...). Khi cần, dùng "Nhập từ nội dung đã sao chép".':
      'Copy the text below and keep it somewhere safe (notes, email...). When needed, use "Import from copied text".',
  'Nhập từ file chỉ có trong bản app đầy đủ. Hãy dùng "Nhập từ nội dung đã sao chép".':
      'Importing from a file is only in the full app. Use "Import from copied text" instead.',
  // — Thông tin —
  'Thông tin': 'About',
  'Phần Mềm Chấm Công': 'Time Clock Software',
  'Tác giả dự án: Trung Nguyen': 'Project author: Trung Nguyen',
  'Với sự hỗ trợ code của AI: Claude - Sonnet': 'Code assisted by AI: Claude - Sonnet',
  'Điện thoại liên hệ: 408-455-0327': 'Contact phone: 408-455-0327',
  // — Nhắc giờ vào ca / ra ca —
  'Nhắc giờ vào ca / ra ca': 'Shift reminders',
  'Chuông báo giờ vào ca và ra ca': 'Alarm for shift start and end times',
  'Nhắc ca làm': 'Shift reminder',
  'Đến giờ vào ca': 'Time to start your shift',
  'Đến giờ ra ca': 'Time to end your shift',
  'Đừng quên bấm "Vào ca" để chấm công.': 'Remember to tap "Start shift" to clock in.',
  'Đừng quên bấm "Ra ca" để lưu ca làm.': 'Remember to tap "End shift" to save your shift.',
  'Đây là thông báo thử. Nếu nghe chuông, nhắc giờ đã hoạt động.':
      'This is a test notification. If you hear the alarm, reminders are working.',
  'Nhắc vào ca': 'Start reminder',
  'Nhắc ra ca': 'End reminder',
  'Lặp lại vào các ngày': 'Repeat on',
  'Gửi thông báo thử': 'Send a test notification',
  'Chưa được cấp quyền thông báo. Hãy bật thông báo cho ứng dụng trong Cài đặt Android.':
      'Notification permission was not granted. Enable notifications for the app in Android Settings.',
  'Thông báo hiện trên thanh trạng thái và phát chuông báo thức, kể cả khi đã tắt ứng dụng. Nếu không nghe chuông, hãy kiểm tra âm lượng báo thức, cho phép thông báo và không hạn chế pin cho ứng dụng trong Cài đặt Android. Lịch nhắc được làm mới mỗi khi bạn mở ứng dụng (đặt trước tối đa 30 ngày).':
      'Reminders appear in the status bar and play an alarm sound, even when the app is closed. If you hear nothing, check the alarm volume, allow notifications and do not restrict battery for the app in Android Settings. The schedule is refreshed every time you open the app (set up to 30 days ahead).',
  'Nhắc giờ chỉ hoạt động trong app Android (bản đầy đủ), không dùng được trên DartPad.':
      'Reminders only work in the full Android app, not on DartPad.',
  // — Ảnh đại diện —
  'Ảnh đại diện': 'Profile photo',
  'Chọn ảnh từ điện thoại': 'Choose a photo from your phone',
  'Đổi ảnh': 'Change photo',
  'Xóa ảnh': 'Remove photo',
  'Xóa ảnh đại diện?': 'Remove the profile photo?',
  'Ảnh sẽ được gỡ khỏi ứng dụng.': 'The photo will be removed from the app.',
  'Chọn ảnh chỉ có trong bản app đầy đủ.': 'Choosing a photo is only in the full app.',
  // — Lịch sử & so sánh —
  'Tổng kết': 'Summary',
  'Chọn khoảng ngày': 'Select date range',
  'Áp dụng': 'Apply',
  'Xem tất cả': 'View all',
  'So sánh thu nhập': 'Compare earnings',
  'So sánh tuần, tháng hoặc năm với một kỳ khác': 'Compare a week, month or year with another period',
  'Kỳ A': 'Period A',
  'Kỳ B': 'Period B',
  'Chênh lệch': 'Difference',
  'Lương': 'Pay',
  'Tháng': 'Month',
  'Năm': 'Year',
  'Chọn một ngày trong tuần cần xem': 'Pick any day in the week',
  'Đổi chỗ hai kỳ': 'Swap the two periods',
  'Hai kỳ đang giống nhau, hãy chọn kỳ khác để so sánh.':
      'Both periods are the same. Pick a different period to compare.',
};

final _enRules = <(RegExp, String Function(RegExpMatch))>[
  (RegExp(r'^(\d+) ca$'), (RegExpMatch m) => '${m[1]} ${m[1] == '1' ? 'shift' : 'shifts'}'),
  (RegExp(r'^(.+) · (\d+) ca$'), (RegExpMatch m) => '${m[1]} · ${m[2]} ${m[2] == '1' ? 'shift' : 'shifts'}'),
  (RegExp(r'^Vào lúc (.+) · (.+)/giờ$'), (RegExpMatch m) => 'Started ${m[1]} · ${m[2]}/hr'),
  (RegExp(r'^(.+)/giờ$'), (RegExpMatch m) => '${m[1]}/hr'),
  (RegExp(r'^Thường (\S+) · OT (\S+) · x2 (\S+)$'), (RegExpMatch m) => 'Regular ${m[1]} · OT ${m[2]} · x2 ${m[3]}'),
  (RegExp(r'^Thường (\d+h\d\d)$'), (RegExpMatch m) => 'Regular ${m[1]}'),
  (RegExp(r'^Tuần (.+)$'), (RegExpMatch m) => 'Week ${m[1]}'),
  (RegExp(r'^Hôm nay: (.+)$'), (RegExpMatch m) => 'Today: ${m[1]}'),
  (RegExp(r'^Hôm nay tính là: (.+)$'), (RegExpMatch m) => 'Today counts as: ${m[1]}'),
  (RegExp(r'^Loại ngày: (.+)$'), (RegExpMatch m) => 'Day type: ${m[1]}'),
  (RegExp(r'^Vào (\d\d:\d\d)$'), (RegExpMatch m) => 'In ${m[1]}'),
  (RegExp(r'^Ra (\d\d:\d\d)$'), (RegExpMatch m) => 'Out ${m[1]}'),
  (RegExp(r'^≈ (.+) tính đến lúc này$'), (RegExpMatch m) => '≈ ${m[1]} so far'),
  (
    RegExp(r'^Thời lượng: (\S+)( · qua nửa đêm)?$'),
    (RegExpMatch m) => 'Duration: ${m[1]}${m[2] != null ? ' · overnight' : ''}'
  ),
  (
    RegExp(r'^Còn thiếu (\S+)\. Tuần có ngày lễ T2–T6[\s\S]*$'),
    (RegExpMatch m) => '${m[1]} short. This week has a Mon–Fri holiday, so the weekend cannot make up the difference.'
  ),
  (
    RegExp(r'^Còn thiếu (\S+)\. Thứ 7, CN[\s\S]*$'),
    (RegExpMatch m) => '${m[1]} short. Sat and Sun can make up at most 8 hours per day.'
  ),
  (
    RegExp(r'^Có (\d+) bản ghi cũ[\s\S]*$'),
    (RegExpMatch m) =>
        '${m[1]} old records could not be read and are not shown. Your data has not been deleted and is still in the DB file (a copy exists as work_records_v3.db.pre_v5.bak).'
  ),
  (RegExp(r'^Xóa ca (.+)\?$'), (RegExpMatch m) => 'Delete the shift on ${m[1]}?'),
  (RegExp(r'^Bảng lương – (.+)$'), (RegExpMatch m) => 'Payroll – ${m[1]}'),
  (RegExp(r'^Trang (.+)$'), (RegExpMatch m) => 'Page ${m[1]}'),
  (
    RegExp(r'^Thu nhập chỉ là ước tính\.[\s\S]*$'),
    (RegExpMatch m) =>
        'Earnings are estimates. Each shift is recorded on the day it starts (including overnight shifts). Mon–Fri: first 8 hours x1, hours 8–12 x1.5, over 12 hours x2. Sat and Sun: the first 8 hours make up missing regular hours for the week (up to 40 hours), then x1.5, and over 12 hours x2. Holidays: x2 and never count as regular hours; if a Mon–Fri holiday falls in the week, the weekend cannot make up missing hours.'
  ),
  (
    RegExp(r'^Tuần làm việc bắt đầu thứ 2\.[\s\S]*$'),
    (RegExpMatch m) =>
        'The work week starts on Monday. Mon–Fri are weekdays, Sat and Sun are OT days (automatic). The default hourly rate only applies to new shifts.'
  ),
  (RegExp(r'^Không thể tải dữ liệu\n([\s\S]*)$'), (RegExpMatch m) => 'Could not load data\n${m[1]}'),
  (RegExp(r'^Không thể lưu ca: ([\s\S]*)$'), (RegExpMatch m) => 'Could not save the shift: ${m[1]}'),
  (RegExp(r'^Không thể lưu: ([\s\S]*)$'), (RegExpMatch m) => 'Could not save: ${m[1]}'),
  (RegExp(r'^Không thể xóa: ([\s\S]*)$'), (RegExpMatch m) => 'Could not delete: ${m[1]}'),
  (
    RegExp(r'^Không thể tạo PDF \(kiểm tra đã khai báo font trong pubspec\.yaml\): ([\s\S]*)$'),
    (RegExpMatch m) => 'Could not create the PDF (check that the font is declared in pubspec.yaml): ${m[1]}'
  ),
  (RegExp(r'^Không thể tạo file: ([\s\S]*)$'), (RegExpMatch m) => 'Could not create the file: ${m[1]}'),
  (RegExp(r'^Không thể chia sẻ file: ([\s\S]*)$'), (RegExpMatch m) => 'Could not share the file: ${m[1]}'),
  (RegExp(r'^Nhập sai quá nhiều lần\. Thử lại sau (\d+) giây\.$'),
      (RegExpMatch m) => 'Too many wrong attempts. Try again in ${m[1]} seconds.'),
  (RegExp(r'^Tìm thấy (\d+) ca làm$'), (RegExpMatch m) => 'Found ${m[1]} shifts'),
  (RegExp(r'^Xuất lúc (.+)$'), (RegExpMatch m) => 'Exported on ${m[1]}'),
  (RegExp(r'^Hiện có (\d+) ca trong ứng dụng\.$'), (RegExpMatch m) => 'The app currently has ${m[1]} shifts.'),
  (RegExp(r'^Đã nhập (\d+) ca, bỏ qua (\d+) ca trùng giờ\.$'),
      (RegExpMatch m) => 'Imported ${m[1]} shifts, skipped ${m[2]} overlapping.'),
  (RegExp(r'^Không thể đọc file: ([\s\S]*)$'), (RegExpMatch m) => 'Could not read the file: ${m[1]}'),
  (RegExp(r'^Không thể nhập: ([\s\S]*)$'), (RegExpMatch m) => 'Could not import: ${m[1]}'),
  (RegExp(r'^Không thể xuất file sao lưu: ([\s\S]*)$'),
      (RegExpMatch m) => 'Could not export the backup file: ${m[1]}'),
  (RegExp(r'^Tháng (\d{2}/\d{4})$'), (RegExpMatch m) => 'Month ${m[1]}'),
  (RegExp(r'^Năm (\d{4})$'), (RegExpMatch m) => 'Year ${m[1]}'),
  (RegExp(r'^Không thể chọn ảnh: ([\s\S]*)$'), (RegExpMatch m) => 'Could not choose the photo: ${m[1]}'),
];

String tr(String s) {
  if (currentLang != 'en') return s;
  final exact = _enExact[s];
  if (exact != null) return exact;
  for (final (re, build) in _enRules) {
    final match = re.firstMatch(s);
    if (match != null) return build(match);
  }
  return s;
}

/// Thay cho Text của Flutter: tự dịch chuỗi theo ngôn ngữ hiện tại.
class Text extends StatelessWidget {
  const Text(
    this.data, {
    super.key,
    this.style,
    this.textAlign,
    this.maxLines,
    this.overflow,
    this.softWrap,
  });
  final String data;
  final TextStyle? style;
  final TextAlign? textAlign;
  final int? maxLines;
  final TextOverflow? overflow;
  final bool? softWrap;

  @override
  Widget build(BuildContext context) => mat.Text(
        tr(data),
        style: style,
        textAlign: textAlign,
        maxLines: maxLines,
        overflow: overflow,
        softWrap: softWrap,
      );
}

const dayTypes = ['Thường', 'OT', 'Lễ'];
const kDefaultRate = 25.0;

const _tealMid = Color(0xFF0F766E);
const _tealDark = Color(0xFF0B4F4A);
const _surfaceBg = Color(0xFFF3F6F8);
const _muted = Color(0xFF64748B);
String _fmtStart(DateTime d) => '${_two(d.hour)}:${_two(d.minute)} ${_two(d.day)}/${_two(d.month)}';

final _moneyFmt = NumberFormat.currency(locale: 'en_US', symbol: r'$', decimalDigits: 2);
final _rateFmt = NumberFormat('0.##', 'en_US');
final _dateFmt = DateFormat('yyyy-MM-dd');
final _dmyFmt = DateFormat('dd/MM/yyyy');
final _timeFmt = DateFormat('HH:mm');

String money(double v) => _moneyFmt.format(v);

String fmtHours(double hours) {
  final m = (hours * 60).round();
  return '${m ~/ 60}h${(m % 60).toString().padLeft(2, '0')}';
}

String _two(int n) => n.toString().padLeft(2, '0');

DateTime floorMinute(DateTime d) => DateTime(d.year, d.month, d.day, d.hour, d.minute);
DateTime dateOnly(DateTime d) => DateTime(d.year, d.month, d.day);

DateTime startOfWeek(DateTime d, int weekStart) {
  final day = dateOnly(d);
  return DateTime(day.year, day.month, day.day - (day.weekday - weekStart + 7) % 7);
}

/// Loại ngày tự suy ra: ngày lễ (đánh dấu trong cài đặt) > thứ 7/CN (OT) > ngày thường.
String dayTypeOf(DateTime d, Set<String> holidays) {
  if (holidays.contains(_dateFmt.format(d))) return 'Lễ';
  return d.weekday >= DateTime.saturday ? 'OT' : 'Thường';
}

/// Nhận cả "23.5", "23,5", "1,234.50", "1.234,50".
double? parseNumber(String input) {
  var s = input.trim().replaceAll(RegExp(r'[\s$]'), '');
  if (s.isEmpty) return null;
  final hasComma = s.contains(',');
  final hasDot = s.contains('.');
  if (hasComma && hasDot) {
    final commaIsDecimal = s.lastIndexOf(',') > s.lastIndexOf('.');
    s = commaIsDecimal ? s.replaceAll('.', '').replaceAll(',', '.') : s.replaceAll(',', '');
  } else if (hasComma) {
    s = s.replaceAll(',', '.');
  }
  final v = double.tryParse(s);
  return (v == null || !v.isFinite || v < 0) ? null : v;
}

Color typeColor(String t) => switch (t) {
      'Lễ' => Colors.deepOrange,
      'OT' => Colors.indigo,
      _ => Colors.teal,
    };

IconData typeIcon(String t) => switch (t) {
      'Lễ' => Icons.celebration,
      'OT' => Icons.more_time,
      _ => Icons.work_outline,
    };

bool _near(double a, double b) => (a - b).abs() < 1e-6;

// ════════════════════════════════════════════════════════════════════════════
//  Model
// ════════════════════════════════════════════════════════════════════════════

class WorkRecord {
  const WorkRecord({
    this.id,
    required this.start,
    required this.end,
    required this.dayType,
    required this.hourlyRate,
    this.totalHours = 0,
    this.regularHours = 0,
    this.otHours = 0,
    this.doubleHours = 0,
    this.earned = 0,
  });

  final int? id;
  final DateTime start;
  final DateTime end;
  final String dayType;
  final double hourlyRate;
  final double totalHours;
  final double regularHours;
  final double otHours; // x1.5
  final double doubleHours; // x2
  final double earned;

  int get minutes => end.difference(start).inMinutes;
  bool get overnight => dateOnly(end) != dateOnly(start);

  WorkRecord withId(int newId) => WorkRecord(
        id: newId,
        start: start,
        end: end,
        dayType: dayType,
        hourlyRate: hourlyRate,
        totalHours: totalHours,
        regularHours: regularHours,
        otHours: otHours,
        doubleHours: doubleHours,
        earned: earned,
      );

  WorkRecord withPay({
    required double total,
    required double regular,
    required double ot,
    required double dbl,
    required double earned,
    required String dayType,
  }) =>
      WorkRecord(
        id: id,
        start: start,
        end: end,
        dayType: dayType,
        hourlyRate: hourlyRate,
        totalHours: total,
        regularHours: regular,
        otHours: ot,
        doubleHours: dbl,
        earned: earned,
      );

  bool samePay(WorkRecord o) =>
      dayType == o.dayType &&
      _near(totalHours, o.totalHours) &&
      _near(regularHours, o.regularHours) &&
      _near(otHours, o.otHours) &&
      _near(doubleHours, o.doubleHours) &&
      (earned - o.earned).abs() < 0.005;

  // Vẫn ghi dateStr/checkInStr/checkOutStr để tương thích cột cũ (NOT NULL).
  Map<String, Object?> toMap() => {
        'dateStr': _dateFmt.format(start),
        'checkInStr': _timeFmt.format(start),
        'checkOutStr': _timeFmt.format(end),
        'startMs': start.millisecondsSinceEpoch,
        'endMs': end.millisecondsSinceEpoch,
        'totalHours': totalHours,
        'regularHours': regularHours,
        'otHours': otHours,
        'doubleTimeHours': doubleHours,
        'dayType': dayType,
        'hourlyRate': hourlyRate,
        'earned': earned,
      };

  static WorkRecord? fromMap(Map<String, Object?> m) {
    double n(String k) => (m[k] as num?)?.toDouble() ?? 0;
    final startMs = (m['startMs'] as num?)?.toInt() ?? 0;
    final endMs = (m['endMs'] as num?)?.toInt() ?? 0;
    if (startMs <= 0 || endMs <= startMs) return null;
    final type = m['dayType'] as String? ?? 'Thường';
    return WorkRecord(
      id: m['id'] as int?,
      start: DateTime.fromMillisecondsSinceEpoch(startMs),
      end: DateTime.fromMillisecondsSinceEpoch(endMs),
      dayType: dayTypes.contains(type) ? type : 'Thường',
      hourlyRate: n('hourlyRate'),
      totalHours: n('totalHours'),
      regularHours: n('regularHours'),
      otHours: n('otHours'),
      doubleHours: n('doubleTimeHours'),
      earned: n('earned'),
    );
  }
}

class RecordDraft {
  const RecordDraft({
    this.id,
    required this.start,
    required this.end,
    required this.rate,
  });
  final int? id;
  final DateTime start;
  final DateTime end;
  final double rate;
}

class PeriodSummary {
  const PeriodSummary({
    this.earned = 0,
    this.total = 0,
    this.regular = 0,
    this.ot = 0,
    this.dbl = 0,
    this.shifts = 0,
  });
  final double earned, total, regular, ot, dbl;
  final int shifts;
}

PeriodSummary summarize(Iterable<WorkRecord> records, DateTime from, DateTime to) {
  double e = 0, t = 0, r = 0, o = 0, d = 0;
  var n = 0;
  for (final x in records) {
    if (x.start.isBefore(from) || !x.start.isBefore(to)) continue;
    e += x.earned;
    t += x.totalHours;
    r += x.regularHours;
    o += x.otHours;
    d += x.doubleHours;
    n++;
  }
  return PeriodSummary(earned: e, total: t, regular: r, ot: o, dbl: d, shifts: n);
}

// ════════════════════════════════════════════════════════════════════════════
//  Tính lương (hàm thuần, dễ unit test)
//
//  Mỗi ca được ghi vào NGÀY BẮT ĐẦU. Tuần bắt đầu thứ 2.
//  Loại ngày tự suy ra: Lễ (đánh dấu trong cài đặt) > thứ 7/CN = OT > Thường.
//
//   • Ngày thường (T2–T6): 8 giờ đầu x1 (tối đa 40 giờ thường/tuần, vượt → x1.5);
//     giờ 8–12 x1.5; trên 12 giờ x2. Giờ OT trong ngày KHÔNG dùng để bù giờ thường.
//   • Thứ 7/CN (OT): x1.5, giờ 8–12 x1.5, trên 12 giờ/ngày x2. Nếu T2–T6 chưa đủ
//     40 giờ thường thì tối đa 8 giờ đầu của mỗi ngày cuối tuần được bù vào x1.
//   • Lễ: cả ca x2, không tính vào giờ thường. Tuần có ngày lễ rơi vào T2–T6
//     thì cuối tuần KHÔNG được bù giờ thường (vẫn tính OT bình thường).
// ════════════════════════════════════════════════════════════════════════════

abstract final class PayCalculator {
  static const _dailyRegular = 8 * 60;
  static const _dailyOt = 12 * 60;
  static const _weeklyRegular = 40 * 60;

  static List<WorkRecord> calculate(List<WorkRecord> input, Set<String> holidays) {
    final sorted = [...input]..sort((a, b) {
        final c = a.start.compareTo(b.start);
        return c != 0 ? c : (a.id ?? 0).compareTo(b.id ?? 0);
      });
    final daily = <DateTime, int>{};
    final weekly = <DateTime, int>{};
    return [for (final r in sorted) _one(r, holidays, daily, weekly)];
  }

  /// Tuần (bắt đầu thứ 2) có ngày lễ rơi vào T2–T6?
  static bool _weekdayHolidayIn(DateTime monday, Set<String> holidays) {
    for (var i = 0; i < 5; i++) {
      final d = DateTime(monday.year, monday.month, monday.day + i);
      if (holidays.contains(_dateFmt.format(d))) return true;
    }
    return false;
  }

  static WorkRecord _one(
    WorkRecord r,
    Set<String> holidays,
    Map<DateTime, int> daily,
    Map<DateTime, int> weekly,
  ) {
    final day = dateOnly(r.start);
    final week = startOfWeek(r.start, DateTime.monday);
    final type = dayTypeOf(r.start, holidays);
    var remaining = r.minutes;
    var reg = 0, ot = 0, dbl = 0;

    if (type == 'Lễ') {
      // Cả ca x2, độc lập: không cộng vào bộ đếm ngày/tuần.
      dbl = remaining;
      remaining = 0;
    } else {
      final weekend = type == 'OT';
      final canFill = weekend && !_weekdayHolidayIn(week, holidays);

      while (remaining > 0) {
        final used = daily[day] ?? 0;
        final int room;
        if (used >= _dailyOt) {
          room = remaining;
        } else if (used < _dailyRegular) {
          room = _dailyRegular - used;
        } else {
          room = _dailyOt - used;
        }
        final chunk = math.min(remaining, room);

        if (used >= _dailyOt) {
          dbl += chunk;
        } else if (used >= _dailyRegular) {
          ot += chunk; // giờ 8–12 trong ngày: x1.5
        } else if (weekend) {
          // 8 giờ đầu của ngày cuối tuần: bù vào giờ thường còn thiếu của tuần
          // (tối đa 40 giờ) trước, phần còn lại x1.5.
          var fill = 0;
          if (canFill) {
            final usedWeek = weekly[week] ?? 0;
            fill = math.min(chunk, math.max(0, _weeklyRegular - usedWeek));
            reg += fill;
            weekly[week] = usedWeek + fill;
          }
          ot += chunk - fill;
        } else {
          // Ngày thường, 8 giờ đầu: giờ thường, tối đa 40 giờ/tuần, vượt → OT.
          final usedWeek = weekly[week] ?? 0;
          final regular = math.min(chunk, math.max(0, _weeklyRegular - usedWeek));
          reg += regular;
          ot += chunk - regular;
          weekly[week] = usedWeek + regular;
        }
        daily[day] = used + chunk;
        remaining -= chunk;
      }
    }

    final raw = (reg + ot * 1.5 + dbl * 2) / 60.0 * r.hourlyRate;
    return r.withPay(
      total: r.minutes / 60.0,
      regular: reg / 60.0,
      ot: ot / 60.0,
      dbl: dbl / 60.0,
      earned: (raw * 100).roundToDouble() / 100,
      dayType: type,
    );
  }
}

/// Ước tính tiền của ca đang làm tính đến `now` (chỉ tính lại trong tuần của ca).
double estimateActive(
  List<WorkRecord> records,
  ActiveShift a,
  Set<String> holidays,
  DateTime now,
) {
  final end = floorMinute(now);
  if (!end.isAfter(a.start)) return 0;
  final week = startOfWeek(a.start, DateTime.monday);
  final next = DateTime(week.year, week.month, week.day + 7);
  const probeId = -1;
  final raw = [
    for (final r in records)
      if (!r.start.isBefore(week) && r.start.isBefore(next)) r,
    WorkRecord(id: probeId, start: a.start, end: end, dayType: 'Thường', hourlyRate: a.rate),
  ];
  return PayCalculator.calculate(raw, holidays).firstWhere((r) => r.id == probeId).earned;
}

// ════════════════════════════════════════════════════════════════════════════
//  Lưu trữ
// ════════════════════════════════════════════════════════════════════════════

class AppDatabase {
  AppDatabase._();
  static final instance = AppDatabase._();

  Future<Database>? _open;

  Future<Database> get db async {
    try {
      return await (_open ??= _openDb());
    } catch (_) {
      _open = null;
      rethrow;
    }
  }

  Future<Database> _openDb() async {
    final dir = await getDatabasesPath();
    final path = p.join(dir, 'work_records_v3.db');
    await _backupOnce(path);
    final db = await openDatabase(
      path,
      version: 5,
      onCreate: _onCreate,
      onUpgrade: _onUpgrade,
    );
    await _importLegacy(db, dir);
    return db;
  }

  /// Nhập một lần dữ liệu từ các file DB tên cũ (nếu có). Không bao giờ làm
  /// app lỗi khi mở: thất bại thì bỏ qua và thử lại ở lần mở sau.
  Future<void> _importLegacy(Database db, String dir) async {
    final prefs = await SharedPreferences.getInstance();
    final rate = prefs.getDouble('hourlyRate') ?? kDefaultRate;

    for (final name in const ['work_records.db', 'work_records_v2.db']) {
      final flag = 'legacyImported:$name';
      if (prefs.getBool(flag) == true) continue;
      try {
        final legacyPath = p.join(dir, name);
        if (!await databaseExists(legacyPath)) continue;

        final old = await openDatabase(legacyPath, readOnly: true);
        try {
          final rows = await old.query('records');
          final batch = db.batch();
          for (final m in rows) {
            final start = DateTime.tryParse('${m['dateStr']} ${m['checkInStr']}:00');
            final end0 = DateTime.tryParse('${m['dateStr']} ${m['checkOutStr']}:00');
            if (start == null || end0 == null) continue;
            var end = end0;
            if (!end.isAfter(start)) {
              end = DateTime(end.year, end.month, end.day + 1, end.hour, end.minute);
            }
            final type = m['dayType'] as String?;
            final oldRate = (m['hourlyRate'] as num?)?.toDouble();
            batch.insert(
              'records',
              WorkRecord(
                start: start,
                end: end,
                dayType: dayTypes.contains(type) ? type! : 'Thường',
                hourlyRate: (oldRate != null && oldRate > 0) ? oldRate : rate,
              ).toMap(),
            );
          }
          await batch.commit(noResult: true);
        } finally {
          await old.close();
        }
        await prefs.setBool(flag, true);
      } catch (_) {
        // bỏ qua, thử lại lần mở sau
      }
    }
  }

  /// Chép file DB hiện có sang bản .bak một lần duy nhất, trước khi nâng cấp
  /// cấu trúc. Nếu có sự cố, dữ liệu gốc vẫn còn nguyên.
  Future<void> _backupOnce(String path) async {
    try {
      final src = File(path);
      final bak = File('$path.pre_v5.bak');
      if (await src.exists() && !await bak.exists()) await src.copy(bak.path);
    } catch (_) {}
  }

  Future<void> _onCreate(Database db, int version) async {
    await db.execute('''
      CREATE TABLE records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        dateStr TEXT NOT NULL,
        checkInStr TEXT NOT NULL,
        checkOutStr TEXT NOT NULL,
        startMs INTEGER NOT NULL DEFAULT 0,
        endMs INTEGER NOT NULL DEFAULT 0,
        totalHours REAL NOT NULL DEFAULT 0,
        regularHours REAL NOT NULL DEFAULT 0,
        otHours REAL NOT NULL DEFAULT 0,
        doubleTimeHours REAL NOT NULL DEFAULT 0,
        dayType TEXT NOT NULL DEFAULT 'Thường',
        hourlyRate REAL NOT NULL DEFAULT 0,
        earned REAL NOT NULL DEFAULT 0
      )
    ''');
    await db.execute('CREATE INDEX idx_records_start ON records(startMs)');
  }

  Future<void> _onUpgrade(Database db, int oldVersion, int newVersion) async {
    final cols = (await db.rawQuery('PRAGMA table_info(records)'))
        .map((c) => c['name'] as String)
        .toSet();
    Future<void> add(String name, String type) async {
      if (!cols.contains(name)) {
        await db.execute('ALTER TABLE records ADD COLUMN $name $type NOT NULL DEFAULT 0');
      }
    }

    await add('doubleTimeHours', 'REAL');
    await add('hourlyRate', 'REAL');
    await add('startMs', 'INTEGER');
    await add('endMs', 'INTEGER');

    // Chuyển bản ghi cũ (chuỗi ngày + HH:mm) sang mốc thời gian đầy đủ.
    final rows = await db.query(
      'records',
      columns: ['id', 'dateStr', 'checkInStr', 'checkOutStr'],
      where: 'startMs = 0',
    );
    final batch = db.batch();
    for (final r in rows) {
      final start = DateTime.tryParse('${r['dateStr']} ${r['checkInStr']}:00');
      final end0 = DateTime.tryParse('${r['dateStr']} ${r['checkOutStr']}:00');
      if (start == null || end0 == null) continue;
      var end = end0;
      if (!end.isAfter(start)) {
        end = DateTime(end.year, end.month, end.day + 1, end.hour, end.minute);
      }
      batch.update(
        'records',
        {'startMs': start.millisecondsSinceEpoch, 'endMs': end.millisecondsSinceEpoch},
        where: 'id = ?',
        whereArgs: [r['id']],
      );
    }
    await batch.commit(noResult: true);
    await db.execute('CREATE INDEX IF NOT EXISTS idx_records_start ON records(startMs)');
  }

  Future<List<WorkRecord>> getAll() async {
    final rows = await (await db).query('records', orderBy: 'startMs DESC, id DESC');
    return rows.map(WorkRecord.fromMap).whereType<WorkRecord>().toList();
  }

  /// Số dòng còn trong DB nhưng không đọc được (giờ/ngày hỏng). Không bị xóa.
  Future<int> countUnreadable() async {
    final rows = await (await db).rawQuery(
      'SELECT COUNT(*) FROM records WHERE startMs <= 0 OR endMs <= startMs',
    );
    return Sqflite.firstIntValue(rows) ?? 0;
  }

  Future<int> insert(WorkRecord r) async => (await db).insert('records', r.toMap());

  Future<void> update(WorkRecord r) async {
    if (r.id == null) return;
    await (await db).update('records', r.toMap(), where: 'id = ?', whereArgs: [r.id]);
  }

  Future<void> delete(int id) async =>
      (await db).delete('records', where: 'id = ?', whereArgs: [id]);

  Future<void> updateAll(List<WorkRecord> records) async {
    if (records.isEmpty) return;
    final batch = (await db).batch();
    for (final r in records) {
      if (r.id == null) continue;
      batch.update('records', r.toMap(), where: 'id = ?', whereArgs: [r.id]);
    }
    await batch.commit(noResult: true);
  }

  Future<void> insertAll(List<WorkRecord> records) async {
    if (records.isEmpty) return;
    final batch = (await db).batch();
    for (final r in records) {
      batch.insert('records', r.toMap());
    }
    await batch.commit(noResult: true);
  }

  /// Xóa hết rồi ghi lại trong MỘT giao dịch: lỗi giữa chừng thì dữ liệu cũ vẫn còn.
  Future<void> replaceWith(List<WorkRecord> records) async {
    await (await db).transaction((txn) async {
      await txn.delete('records');
      final batch = txn.batch();
      for (final r in records) {
        batch.insert('records', r.toMap());
      }
      await batch.commit(noResult: true);
    });
  }
}

class AppSettings {
  const AppSettings({
    this.hourlyRate = kDefaultRate,
    this.holidays = const {},
    this.lang = 'vi',
    this.reminders = const ReminderConfig(),
  });
  final double hourlyRate;
  final Set<String> holidays; // 'yyyy-MM-dd'
  final String lang; // 'vi' hoặc 'en'
  final ReminderConfig reminders;

  /// Quy tắc cố định: tuần làm việc bắt đầu từ thứ 2.
  int get weekStart => DateTime.monday;

  AppSettings copyWith({
    double? hourlyRate,
    Set<String>? holidays,
    String? lang,
    ReminderConfig? reminders,
  }) =>
      AppSettings(
        hourlyRate: hourlyRate ?? this.hourlyRate,
        holidays: holidays ?? this.holidays,
        lang: lang ?? this.lang,
        reminders: reminders ?? this.reminders,
      );

  static Future<AppSettings> load() async {
    final prefs = await SharedPreferences.getInstance();
    final days = (prefs.getStringList('rem_days') ?? const <String>['1', '2', '3', '4', '5'])
        .map(int.tryParse)
        .whereType<int>()
        .where((d) => d >= 1 && d <= 7)
        .toSet();
    return AppSettings(
      hourlyRate: prefs.getDouble('hourlyRate') ?? kDefaultRate,
      holidays: (prefs.getStringList('holidays') ?? const <String>[]).toSet(),
      lang: supportedLangs.containsKey(prefs.getString('lang')) ? prefs.getString('lang')! : 'vi',
      reminders: ReminderConfig(
        startMin: prefs.getInt('rem_start'),
        endMin: prefs.getInt('rem_end'),
        days: days.isEmpty ? const {1, 2, 3, 4, 5} : days,
      ),
    );
  }

  Future<void> save() async {
    final prefs = await SharedPreferences.getInstance();
    await prefs.setDouble('hourlyRate', hourlyRate);
    await prefs.setStringList('holidays', holidays.toList()..sort());
    await prefs.setString('lang', lang);
    final st = reminders.startMin;
    final en = reminders.endMin;
    if (st == null) {
      await prefs.remove('rem_start');
    } else {
      await prefs.setInt('rem_start', st);
    }
    if (en == null) {
      await prefs.remove('rem_end');
    } else {
      await prefs.setInt('rem_end', en);
    }
    await prefs.setStringList('rem_days', (reminders.days.toList()..sort()).map((d) => '$d').toList());
  }
}

class ActiveShift {
  const ActiveShift({required this.start, required this.dayType, required this.rate});
  final DateTime start;
  final String dayType;
  final double rate;

  static Future<ActiveShift?> load(double fallbackRate) async {
    final prefs = await SharedPreferences.getInstance();
    final ms = prefs.getInt('activeCheckInMillis');
    if (ms == null) return null;
    final type = prefs.getString('activeDayType') ?? 'Thường';
    return ActiveShift(
      start: DateTime.fromMillisecondsSinceEpoch(ms),
      dayType: dayTypes.contains(type) ? type : 'Thường',
      rate: prefs.getDouble('activeRate') ?? fallbackRate,
    );
  }

  Future<void> save() async {
    final prefs = await SharedPreferences.getInstance();
    await prefs.setInt('activeCheckInMillis', start.millisecondsSinceEpoch);
    await prefs.setString('activeDayType', dayType);
    await prefs.setDouble('activeRate', rate);
  }

  static Future<void> clear() async {
    final prefs = await SharedPreferences.getInstance();
    await prefs.remove('activeCheckInMillis');
    await prefs.remove('activeDayType');
    await prefs.remove('activeRate');
  }
}

String buildCsv(List<WorkRecord> newestFirst) {
  String cell(Object? v) => '"${'$v'.replaceAll('"', '""')}"';
  final rows = <List<Object?>>[
    [
      tr('Ngày'), tr('Giờ vào'), tr('Giờ ra'), tr('Loại ngày'), tr('Tổng giờ'), tr('Giờ thường'),
      tr('Giờ OT (x1.5)'), tr('Giờ x2'), tr('Lương/giờ (USD)'), tr('Tiền (USD)')
    ],
    for (final r in newestFirst.reversed)
      [
        _dateFmt.format(r.start),
        _timeFmt.format(r.start),
        _timeFmt.format(r.end) + (r.overnight ? ' (+1)' : ''),
        dayLabel(r.dayType),
        r.totalHours.toStringAsFixed(2),
        r.regularHours.toStringAsFixed(2),
        r.otHours.toStringAsFixed(2),
        r.doubleHours.toStringAsFixed(2),
        r.hourlyRate.toStringAsFixed(2),
        r.earned.toStringAsFixed(2),
      ],
  ];
  return '\uFEFF${rows.map((row) => row.map(cell).join(',')).join('\r\n')}';
}

// ════════════════════════════════════════════════════════════════════════════
//  Khóa ứng dụng, sinh trắc học và sao lưu dữ liệu
// ════════════════════════════════════════════════════════════════════════════

// <PLATFORM:LOCK>
/// Có chọn được file sao lưu từ bộ nhớ máy hay không.
const kHasFilePicker = true;

/// Lưu cấu hình khóa (mật khẩu đã băm, bật sinh trắc học...).
class LockStore {
  Future<String?> getString(String key) async {
    final prefs = await SharedPreferences.getInstance();
    return prefs.getString(key);
  }

  Future<void> setString(String key, String? value) async {
    final prefs = await SharedPreferences.getInstance();
    if (value == null) {
      await prefs.remove(key);
    } else {
      await prefs.setString(key, value);
    }
  }

  Future<bool> getBool(String key) async {
    final prefs = await SharedPreferences.getInstance();
    return prefs.getBool(key) ?? false;
  }

  Future<void> setBool(String key, bool value) async {
    final prefs = await SharedPreferences.getInstance();
    await prefs.setBool(key, value);
  }
}

/// Xác thực vân tay / khuôn mặt (local_auth).
class Biometrics {
  final LocalAuthentication _auth = LocalAuthentication();

  Future<bool> available() async {
    try {
      if (!await _auth.isDeviceSupported()) return false;
      return (await _auth.getAvailableBiometrics()).isNotEmpty;
    } catch (_) {
      return false;
    }
  }

  Future<bool> authenticate(String reason) async {
    try {
      return await _auth.authenticate(
        localizedReason: reason,
        options: const AuthenticationOptions(biometricOnly: true, stickyAuth: true),
      );
    } catch (_) {
      return false;
    }
  }
}
// </PLATFORM:LOCK>

// ── Băm mật khẩu: PBKDF2-HMAC-SHA256 viết bằng Dart thuần (không cần package) ──

const List<int> _k256 = [
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
];

const List<int> _h256 = [
  0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
];

int _rotr32(int x, int n) => ((x >>> n) | (x << (32 - n))) & 0xFFFFFFFF;

List<int> sha256Bytes(List<int> data) {
  final len = data.length;
  final padded = ((len + 9 + 63) ~/ 64) * 64;
  final msg = List<int>.filled(padded, 0);
  for (var i = 0; i < len; i++) {
    msg[i] = data[i] & 0xFF;
  }
  msg[len] = 0x80;
  final bits = len * 8;
  msg[padded - 4] = (bits >>> 24) & 0xFF;
  msg[padded - 3] = (bits >>> 16) & 0xFF;
  msg[padded - 2] = (bits >>> 8) & 0xFF;
  msg[padded - 1] = bits & 0xFF;

  var h0 = _h256[0], h1 = _h256[1], h2 = _h256[2], h3 = _h256[3];
  var h4 = _h256[4], h5 = _h256[5], h6 = _h256[6], h7 = _h256[7];
  final w = List<int>.filled(64, 0);

  for (var off = 0; off < padded; off += 64) {
    for (var i = 0; i < 16; i++) {
      final j = off + i * 4;
      w[i] = ((msg[j] << 24) | (msg[j + 1] << 16) | (msg[j + 2] << 8) | msg[j + 3]) & 0xFFFFFFFF;
    }
    for (var i = 16; i < 64; i++) {
      final s0 = _rotr32(w[i - 15], 7) ^ _rotr32(w[i - 15], 18) ^ (w[i - 15] >>> 3);
      final s1 = _rotr32(w[i - 2], 17) ^ _rotr32(w[i - 2], 19) ^ (w[i - 2] >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) & 0xFFFFFFFF;
    }
    var a = h0, b = h1, c = h2, d = h3, e = h4, f = h5, g = h6, h = h7;
    for (var i = 0; i < 64; i++) {
      final bigS1 = _rotr32(e, 6) ^ _rotr32(e, 11) ^ _rotr32(e, 25);
      final ch = (e & f) ^ ((~e & 0xFFFFFFFF) & g);
      final t1 = (h + bigS1 + ch + _k256[i] + w[i]) & 0xFFFFFFFF;
      final bigS0 = _rotr32(a, 2) ^ _rotr32(a, 13) ^ _rotr32(a, 22);
      final maj = (a & b) ^ (a & c) ^ (b & c);
      final t2 = (bigS0 + maj) & 0xFFFFFFFF;
      h = g;
      g = f;
      f = e;
      e = (d + t1) & 0xFFFFFFFF;
      d = c;
      c = b;
      b = a;
      a = (t1 + t2) & 0xFFFFFFFF;
    }
    h0 = (h0 + a) & 0xFFFFFFFF;
    h1 = (h1 + b) & 0xFFFFFFFF;
    h2 = (h2 + c) & 0xFFFFFFFF;
    h3 = (h3 + d) & 0xFFFFFFFF;
    h4 = (h4 + e) & 0xFFFFFFFF;
    h5 = (h5 + f) & 0xFFFFFFFF;
    h6 = (h6 + g) & 0xFFFFFFFF;
    h7 = (h7 + h) & 0xFFFFFFFF;
  }

  final out = <int>[];
  for (final v in [h0, h1, h2, h3, h4, h5, h6, h7]) {
    out
      ..add((v >>> 24) & 0xFF)
      ..add((v >>> 16) & 0xFF)
      ..add((v >>> 8) & 0xFF)
      ..add(v & 0xFF);
  }
  return out;
}

List<int> hmacSha256(List<int> key, List<int> msg) {
  var k = key.length > 64 ? sha256Bytes(key) : List<int>.of(key);
  k = [...k, ...List<int>.filled(64 - k.length, 0)];
  final ipad = [for (final b in k) b ^ 0x36];
  final opad = [for (final b in k) b ^ 0x5c];
  return sha256Bytes([...opad, ...sha256Bytes([...ipad, ...msg])]);
}

List<int> _pbkdf2(String password, List<int> salt, int iterations) {
  final key = utf8.encode(password);
  var u = hmacSha256(key, [...salt, 0, 0, 0, 1]);
  final t = List<int>.of(u);
  for (var i = 1; i < iterations; i++) {
    u = hmacSha256(key, u);
    for (var j = 0; j < t.length; j++) {
      t[j] ^= u[j];
    }
  }
  return t;
}

abstract final class PasswordHasher {
  static const _iterations = 10000;

  /// Trả về chuỗi 'p1:số-vòng:muối:băm' (base64) để lưu; không lưu mật khẩu gốc.
  static String create(String password) {
    final rnd = math.Random.secure();
    final salt = List<int>.generate(16, (_) => rnd.nextInt(256));
    final hash = _pbkdf2(password, salt, _iterations);
    return 'p1:$_iterations:${base64.encode(salt)}:${base64.encode(hash)}';
  }

  static bool verify(String password, String stored) {
    final parts = stored.split(':');
    if (parts.length != 4 || parts[0] != 'p1') return false;
    final iterations = int.tryParse(parts[1]);
    if (iterations == null || iterations < 1 || iterations > 1000000) return false;
    try {
      final salt = base64.decode(parts[2]);
      final expected = base64.decode(parts[3]);
      final actual = _pbkdf2(password, salt, iterations);
      var diff = expected.length ^ actual.length;
      for (var i = 0; i < expected.length && i < actual.length; i++) {
        diff |= expected[i] ^ actual[i];
      }
      return diff == 0;
    } catch (_) {
      return false;
    }
  }
}

String? validateNewPassword(String password, String confirm) {
  if (password.length < 4) return 'Mật khẩu cần ít nhất 4 ký tự.';
  if (password != confirm) return 'Hai mật khẩu không giống nhau.';
  return null;
}

// ── Trạng thái khóa ứng dụng ────────────────────────────────────────────────

class AppLock extends ChangeNotifier {
  AppLock(this._store, this._bio);
  final LockStore _store;
  final Biometrics _bio;

  /// Khóa lại khi ứng dụng bị đưa xuống nền quá số giây này.
  static const graceSeconds = 30;
  static const _maxFails = 5;
  static const _cooldown = Duration(seconds: 30);

  bool ready = false;
  bool enabled = false;
  bool locked = false;
  bool needsSetup = false;
  bool biometricOn = false;
  bool biometricAvailable = false;

  String? _hash;
  DateTime? _pausedAt;
  bool _busy = false; // đang hiện hộp thoại hệ thống (sinh trắc học, chọn file...)
  int _fails = 0;
  DateTime? _blockedUntil;

  Future<void> load() async {
    try {
      _hash = await _store.getString('lock_hash');
      biometricOn = await _store.getBool('lock_bio');
      final prompted = await _store.getBool('lock_prompted');
      enabled = _hash != null;
      locked = enabled;
      needsSetup = !enabled && !prompted;
      biometricAvailable = await _bio.available();
    } catch (_) {
      // Không đọc được cấu hình: mở ứng dụng bình thường để không khóa mất dữ liệu.
    }
    ready = true;
    notifyListeners();
  }

  /// Số giây còn bị chặn vì nhập sai nhiều lần (0 nếu không bị chặn).
  int get blockedSeconds {
    final until = _blockedUntil;
    if (until == null) return 0;
    final s = until.difference(DateTime.now()).inSeconds + 1;
    return s > 0 ? s : 0;
  }

  Future<void> skipSetup() async {
    await _store.setBool('lock_prompted', true);
    needsSetup = false;
    notifyListeners();
  }

  Future<void> setPassword(String password) async {
    final hash = PasswordHasher.create(password);
    _hash = hash;
    await _store.setString('lock_hash', hash);
    await _store.setBool('lock_prompted', true);
    enabled = true;
    locked = false;
    needsSetup = false;
    _fails = 0;
    _blockedUntil = null;
    notifyListeners();
  }

  bool checkPassword(String password) {
    final h = _hash;
    return h != null && PasswordHasher.verify(password, h);
  }

  /// Mở khóa bằng mật khẩu. Trả về null nếu đúng, hoặc thông báo lỗi.
  String? unlock(String password) {
    final blocked = blockedSeconds;
    if (blocked > 0) return 'Nhập sai quá nhiều lần. Thử lại sau $blocked giây.';
    if (checkPassword(password)) {
      locked = false;
      _fails = 0;
      _blockedUntil = null;
      notifyListeners();
      return null;
    }
    _fails++;
    if (_fails >= _maxFails) {
      _fails = 0;
      _blockedUntil = DateTime.now().add(_cooldown);
      return 'Nhập sai quá nhiều lần. Thử lại sau ${_cooldown.inSeconds} giây.';
    }
    return 'Mật khẩu không đúng.';
  }

  Future<String?> changePassword(String current, String newPassword) async {
    if (!checkPassword(current)) return 'Mật khẩu hiện tại không đúng.';
    await setPassword(newPassword);
    return null;
  }

  Future<String?> disable(String current) async {
    if (!checkPassword(current)) return 'Mật khẩu hiện tại không đúng.';
    _hash = null;
    enabled = false;
    locked = false;
    biometricOn = false;
    await _store.setString('lock_hash', null);
    await _store.setBool('lock_bio', false);
    notifyListeners();
    return null;
  }

  Future<bool> setBiometric(bool on) async {
    if (!enabled) return false;
    if (on) {
      _busy = true;
      final ok = await _bio.authenticate(tr('Xác nhận để bật mở khóa bằng sinh trắc học'));
      _busy = false;
      if (!ok) return false;
    }
    biometricOn = on;
    await _store.setBool('lock_bio', on);
    notifyListeners();
    return true;
  }

  Future<bool> unlockWithBiometric() async {
    if (!enabled || !biometricOn || !biometricAvailable || _busy) return false;
    _busy = true;
    final ok = await _bio.authenticate(tr('Mở khóa ứng dụng'));
    _busy = false;
    if (ok) {
      locked = false;
      _fails = 0;
      notifyListeners();
    }
    return ok;
  }

  /// Chạy một việc cần mở màn hình hệ thống (chia sẻ file, chọn file...)
  /// mà không bị tính là rời ứng dụng.
  Future<T> runExternal<T>(Future<T> Function() action) async {
    _busy = true;
    try {
      return await action();
    } finally {
      _busy = false;
    }
  }

  void onPaused() {
    if (enabled && !locked && !_busy) _pausedAt = DateTime.now();
  }

  void onResumed() {
    final t = _pausedAt;
    _pausedAt = null;
    if (enabled && !locked && !_busy && t != null && DateTime.now().difference(t).inSeconds >= graceSeconds) {
      locked = true;
      notifyListeners();
    }
  }
}

// ── Sao lưu / khôi phục ─────────────────────────────────────────────────────

enum ImportMode { merge, replace }

class BackupData {
  const BackupData({
    required this.records,
    required this.hourlyRate,
    required this.holidays,
    required this.lang,
    required this.exportedAt,
    this.reminders = const ReminderConfig(),
  });
  final List<WorkRecord> records;
  final double hourlyRate;
  final Set<String> holidays;
  final String lang;
  final DateTime exportedAt;
  final ReminderConfig reminders;

  static const _appId = 'theo_doi_luong';
  static const _format = 1;

  String toJson() => const JsonEncoder.withIndent(' ').convert({
        'app': _appId,
        'format': _format,
        'exportedAt': exportedAt.toIso8601String(),
        'settings': {
          'hourlyRate': hourlyRate,
          'holidays': holidays.toList()..sort(),
          'lang': lang,
          'reminders': reminders.toJson(),
        },
        'records': [
          for (final r in records)
            {
              'start': r.start.millisecondsSinceEpoch,
              'end': r.end.millisecondsSinceEpoch,
              'rate': r.hourlyRate,
              'startText': '${_dateFmt.format(r.start)} ${_timeFmt.format(r.start)}',
              'endText': '${_dateFmt.format(r.end)} ${_timeFmt.format(r.end)}',
            },
        ],
      });

  /// Đọc và kiểm tra nội dung file sao lưu; ném FormatException kèm thông báo nếu sai.
  static BackupData parse(String raw) {
    Object? decoded;
    try {
      decoded = jsonDecode(raw.replaceFirst('\uFEFF', '').trim());
    } catch (_) {
      throw const FormatException('Tệp sao lưu không hợp lệ.');
    }
    if (decoded is! Map<String, dynamic> || decoded['app'] != _appId) {
      throw const FormatException('Đây không phải file sao lưu của ứng dụng này.');
    }
    final fmt = decoded['format'];
    if (fmt is! int || fmt < 1) throw const FormatException('Tệp sao lưu không hợp lệ.');
    if (fmt > _format) throw const FormatException('File sao lưu quá mới, hãy cập nhật ứng dụng.');

    try {
      final st = decoded['settings'];
      final settings = st is Map<String, dynamic> ? st : const <String, dynamic>{};
      final rateRaw = settings['hourlyRate'];
      final rate = (rateRaw is num && rateRaw >= 0 && rateRaw.isFinite) ? rateRaw.toDouble() : kDefaultRate;
      final lang = supportedLangs.containsKey(settings['lang']) ? settings['lang'] as String : 'vi';
      final holidays = <String>{};
      final hl = settings['holidays'];
      if (hl is List) {
        for (final h in hl) {
          if (h is String && RegExp(r'^\d{4}-\d{2}-\d{2}$').hasMatch(h)) holidays.add(h);
        }
      }

      final list = decoded['records'];
      if (list is! List || list.length > 200000) {
        throw const FormatException('Dữ liệu trong file bị lỗi.');
      }
      final out = <WorkRecord>[];
      for (final item in list) {
        if (item is! Map) throw const FormatException('Dữ liệu trong file bị lỗi.');
        final s = item['start'];
        final e = item['end'];
        final rt = item['rate'];
        if (s is! int || e is! int || s <= 0 || e <= s || e - s > 72 * 3600 * 1000) {
          throw const FormatException('Dữ liệu trong file bị lỗi.');
        }
        if (rt is! num || rt < 0 || !rt.isFinite) {
          throw const FormatException('Dữ liệu trong file bị lỗi.');
        }
        out.add(WorkRecord(
          start: DateTime.fromMillisecondsSinceEpoch(s),
          end: DateTime.fromMillisecondsSinceEpoch(e),
          dayType: 'Thường', // loại ngày thật do PayCalculator suy ra
          hourlyRate: rt.toDouble(),
        ));
      }

      final ea = decoded['exportedAt'];
      return BackupData(
        records: out,
        hourlyRate: rate,
        holidays: holidays,
        lang: lang,
        reminders: ReminderConfig.fromJson(settings['reminders']),
        exportedAt: (ea is String ? DateTime.tryParse(ea) : null) ?? DateTime.now(),
      );
    } on FormatException {
      rethrow;
    } catch (_) {
      throw const FormatException('Dữ liệu trong file bị lỗi.');
    }
  }
}

/// Chọn các ca sẽ được thêm vào (bỏ qua ca trùng giờ với ca đã có hoặc đã chọn).
({List<WorkRecord> add, int skipped}) planImport(List<WorkRecord> existing, List<WorkRecord> incoming) {
  final occupied = List<WorkRecord>.of(existing);
  final add = <WorkRecord>[];
  var skipped = 0;
  final sorted = [...incoming]..sort((a, b) => a.start.compareTo(b.start));
  for (final r in sorted) {
    final clash = occupied.any((x) => x.start.isBefore(r.end) && x.end.isAfter(r.start));
    if (clash) {
      skipped++;
      continue;
    }
    occupied.add(r);
    add.add(r);
  }
  return (add: add, skipped: skipped);
}

// ════════════════════════════════════════════════════════════════════════════
//  Nhắc giờ, ảnh đại diện, kỳ so sánh
// ════════════════════════════════════════════════════════════════════════════

/// Cấu hình nhắc giờ: giờ vào ca / ra ca (phút trong ngày, null = tắt) và các thứ lặp lại.
class ReminderConfig {
  const ReminderConfig({this.startMin, this.endMin, this.days = const {1, 2, 3, 4, 5}});
  final int? startMin;
  final int? endMin;
  final Set<int> days; // 1 = thứ 2 ... 7 = chủ nhật

  bool get anyEnabled => (startMin != null || endMin != null) && days.isNotEmpty;

  /// Dùng để biết khi nào cần đặt lại lịch thông báo.
  String get signature => '$startMin|$endMin|${(days.toList()..sort()).join(',')}';

  ReminderConfig copyWith({
    int? startMin,
    int? endMin,
    Set<int>? days,
    bool clearStart = false,
    bool clearEnd = false,
  }) =>
      ReminderConfig(
        startMin: clearStart ? null : (startMin ?? this.startMin),
        endMin: clearEnd ? null : (endMin ?? this.endMin),
        days: days ?? this.days,
      );

  Map<String, Object?> toJson() => {
        'start': startMin,
        'end': endMin,
        'days': (days.toList()..sort()),
      };

  static ReminderConfig fromJson(Object? raw) {
    if (raw is! Map) return const ReminderConfig();
    int? minutes(Object? v) => (v is int && v >= 0 && v < 1440) ? v : null;
    final d = raw['days'];
    final days = <int>{
      if (d is List)
        for (final x in d)
          if (x is int && x >= 1 && x <= 7) x,
    };
    return ReminderConfig(
      startMin: minutes(raw['start']),
      endMin: minutes(raw['end']),
      days: days.isEmpty ? const {1, 2, 3, 4, 5} : days,
    );
  }
}

// <PLATFORM:NOTIFY>
/// Đặt được thông báo nhắc giờ hay không (chỉ có trong app Android đầy đủ).
const kHasReminders = true;

/// Đặt lịch thông báo có chuông (flutter_local_notifications).
class ReminderService {
  final FlutterLocalNotificationsPlugin _plugin = FlutterLocalNotificationsPlugin();
  bool _ready = false;

  Future<void> _init() async {
    if (_ready) return;
    await _plugin.initialize(
      const InitializationSettings(android: AndroidInitializationSettings('@mipmap/ic_launcher')),
    );
    _ready = true;
  }

  AndroidFlutterLocalNotificationsPlugin? get _android =>
      _plugin.resolvePlatformSpecificImplementation<AndroidFlutterLocalNotificationsPlugin>();

  /// Xin quyền hiện thông báo (Android 13 trở lên). Trả về true nếu được phép.
  Future<bool> requestPermission() async {
    try {
      await _init();
      return await _android?.requestNotificationsPermission() ?? true;
    } catch (_) {
      return false;
    }
  }

  NotificationDetails _details() => NotificationDetails(
        android: AndroidNotificationDetails(
          'shift_reminders',
          tr('Nhắc giờ vào ca / ra ca'),
          channelDescription: tr('Chuông báo giờ vào ca và ra ca'),
          importance: Importance.max,
          priority: Priority.high,
          category: AndroidNotificationCategory.alarm,
          audioAttributesUsage: AudioAttributesUsage.alarm,
          sound: const UriAndroidNotificationSound('content://settings/system/alarm_alert'),
          playSound: true,
          enableVibration: true,
          visibility: NotificationVisibility.public,
        ),
      );

  Future<void> showTest() async {
    await _init();
    await _plugin.show(
      999,
      tr('Nhắc ca làm'),
      tr('Đây là thông báo thử. Nếu nghe chuông, nhắc giờ đã hoạt động.'),
      _details(),
    );
  }

  Future<void> _schedule(int id, String title, String body, DateTime at, AndroidScheduleMode mode) async {
    // Đổi giờ địa phương sang cùng thời điểm theo UTC: không cần cơ sở dữ liệu múi giờ.
    final when = tz.TZDateTime.from(at, tz.UTC);
    try {
      await _plugin.zonedSchedule(id, title, body, when, _details(), androidScheduleMode: mode);
    } catch (_) {
      // Thiếu quyền báo thức chính xác: dùng chế độ gần đúng thay vì bỏ lỡ.
      await _plugin.zonedSchedule(
        id,
        title,
        body,
        when,
        _details(),
        androidScheduleMode: AndroidScheduleMode.inexactAllowWhileIdle,
      );
    }
  }

  /// Đặt lại toàn bộ lịch nhắc cho 30 ngày tới. Trả về số thông báo đã đặt.
  Future<int> sync(ReminderConfig cfg) async {
    await _init();
    await _plugin.cancelAll();
    if (!cfg.anyEnabled) return 0;

    var mode = AndroidScheduleMode.exactAllowWhileIdle;
    try {
      if (await _android?.canScheduleExactNotifications() == false) {
        mode = AndroidScheduleMode.inexactAllowWhileIdle;
      }
    } catch (_) {}

    final slots = [
      (cfg.startMin, tr('Đến giờ vào ca'), tr('Đừng quên bấm "Vào ca" để chấm công.')),
      (cfg.endMin, tr('Đến giờ ra ca'), tr('Đừng quên bấm "Ra ca" để lưu ca làm.')),
    ];
    final now = DateTime.now();
    var id = 1;
    for (var i = 0; i < 30; i++) {
      final day = DateTime(now.year, now.month, now.day + i);
      if (!cfg.days.contains(day.weekday)) continue;
      for (final (minutes, title, body) in slots) {
        if (minutes == null) continue;
        final at = DateTime(day.year, day.month, day.day, minutes ~/ 60, minutes % 60);
        if (!at.isAfter(now)) continue;
        await _schedule(id++, title, body, at, mode);
      }
    }
    return id - 1;
  }
}
// </PLATFORM:NOTIFY>

// <PLATFORM:AVATAR>
/// Chọn được ảnh đại diện từ bộ nhớ máy hay không.
const kHasAvatarPicker = true;

/// Lưu ảnh đại diện: chép ảnh vào thư mục riêng của ứng dụng và nhớ đường dẫn.
class AvatarStore {
  Future<ImageProvider?> load() async {
    final prefs = await SharedPreferences.getInstance();
    final path = prefs.getString('avatar_path');
    if (path == null) return null;
    final file = File(path);
    if (!await file.exists()) return null;
    return ResizeImage(FileImage(file), width: 256);
  }

  /// Mở bộ chọn ảnh của hệ thống. Trả về null nếu người dùng hủy.
  Future<ImageProvider?> pick() async {
    final picked = await openFile(
      acceptedTypeGroups: const [XTypeGroup(label: 'images', mimeTypes: ['image/*'])],
    );
    if (picked == null) return null;
    final bytes = await picked.readAsBytes();
    if (bytes.isEmpty) return null;
    final dir = await getApplicationDocumentsDirectory();
    final file = File(p.join(dir.path, 'avatar_${DateTime.now().millisecondsSinceEpoch}.jpg'));
    await file.writeAsBytes(bytes, flush: true);
    final prefs = await SharedPreferences.getInstance();
    final old = prefs.getString('avatar_path');
    await prefs.setString('avatar_path', file.path);
    if (old != null && old != file.path) {
      try {
        await File(old).delete();
      } catch (_) {}
    }
    return ResizeImage(FileImage(file), width: 256);
  }

  Future<void> remove() async {
    final prefs = await SharedPreferences.getInstance();
    final old = prefs.getString('avatar_path');
    await prefs.remove('avatar_path');
    if (old != null) {
      try {
        await File(old).delete();
      } catch (_) {}
    }
  }
}
// </PLATFORM:AVATAR>

class AvatarController extends ChangeNotifier {
  AvatarController(this._store);
  final AvatarStore _store;
  ImageProvider? image;

  Future<void> load() async {
    try {
      image = await _store.load();
    } catch (_) {}
    notifyListeners();
  }

  /// Trả về thông báo lỗi, hoặc null nếu thành công / người dùng hủy.
  Future<String?> pick() async {
    try {
      final img = await _store.pick();
      if (img != null) {
        image = img;
        notifyListeners();
      }
      return null;
    } catch (e) {
      return 'Không thể chọn ảnh: $e';
    }
  }

  Future<void> remove() async {
    try {
      await _store.remove();
    } catch (_) {}
    image = null;
    notifyListeners();
  }
}

// ── Kỳ thời gian (tuần / tháng / năm) cho phần so sánh ──────────────────────

enum CompareKind { week, month, year }

/// (từ, đến) với đến là mốc không bao gồm.
(DateTime, DateTime) periodBounds(CompareKind k, DateTime anchor) {
  switch (k) {
    case CompareKind.week:
      final w = startOfWeek(anchor, DateTime.monday);
      return (w, DateTime(w.year, w.month, w.day + 7));
    case CompareKind.month:
      return (DateTime(anchor.year, anchor.month), DateTime(anchor.year, anchor.month + 1));
    case CompareKind.year:
      return (DateTime(anchor.year), DateTime(anchor.year + 1));
  }
}

DateTime shiftPeriod(CompareKind k, DateTime anchor, int delta) {
  switch (k) {
    case CompareKind.week:
      return DateTime(anchor.year, anchor.month, anchor.day + 7 * delta);
    case CompareKind.month:
      return DateTime(anchor.year, anchor.month + delta);
    case CompareKind.year:
      return DateTime(anchor.year + delta);
  }
}

String periodTitle(CompareKind k, DateTime anchor) {
  final b = periodBounds(k, anchor);
  switch (k) {
    case CompareKind.week:
      final sun = DateTime(b.$1.year, b.$1.month, b.$1.day + 6);
      return 'Tuần ${_two(b.$1.day)}/${_two(b.$1.month)} – ${_two(sun.day)}/${_two(sun.month)}/${sun.year}';
    case CompareKind.month:
      return 'Tháng ${_two(anchor.month)}/${anchor.year}';
    case CompareKind.year:
      return 'Năm ${anchor.year}';
  }
}

List<String> get _monthShort => currentLang == 'en'
    ? const ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    : const ['Th1', 'Th2', 'Th3', 'Th4', 'Th5', 'Th6', 'Th7', 'Th8', 'Th9', 'Th10', 'Th11', 'Th12'];

// ════════════════════════════════════════════════════════════════════════════
//  Controller (toàn bộ logic nghiệp vụ, UI chỉ lắng nghe)
// ════════════════════════════════════════════════════════════════════════════

class AppController extends ChangeNotifier {
  AppController(this._db);
  final AppDatabase _db;

  List<WorkRecord> records = const []; // mới nhất trước
  AppSettings settings = const AppSettings();
  ActiveShift? active;
  bool loading = true;
  String? loadError;
  int skipped = 0; // bản ghi cũ không đọc được (vẫn còn trong file DB)

  Future<void> load() async {
    loading = true;
    loadError = null;
    notifyListeners();
    try {
      settings = await AppSettings.load();
      active = await ActiveShift.load(settings.hourlyRate);
      skipped = await _db.countUnreadable();
      await _recalculate(await _db.getAll());
    } catch (e) {
      loadError = '$e';
    }
    loading = false;
    notifyListeners();
  }

  Future<void> refresh() async {
    try {
      await _recalculate(await _db.getAll());
    } catch (e) {
      loadError = '$e';
      notifyListeners();
    }
  }

  /// Tính lại toàn bộ và chỉ ghi những bản ghi có thay đổi.
  Future<void> _recalculate(List<WorkRecord> raw) async {
    final calculated = PayCalculator.calculate(raw, settings.holidays);
    final before = {for (final r in raw) r.id: r};
    final changed = [
      for (final r in calculated)
        if (before[r.id]?.samePay(r) != true) r
    ];
    await _db.updateAll(changed);
    records = calculated.reversed.toList();
    notifyListeners();
  }

  String? validate(RecordDraft d) {
    if (!d.end.isAfter(d.start)) return 'Giờ ra phải sau giờ vào.';
    if (d.end.difference(d.start) > const Duration(hours: 24)) {
      return 'Một ca tối đa 24 giờ.';
    }
    final clash = records.any(
      (r) => r.id != d.id && r.start.isBefore(d.end) && r.end.isAfter(d.start),
    );
    return clash ? 'Ca này trùng giờ với một ca đã có.' : null;
  }

  Future<String?> saveDraft(RecordDraft d) async {
    final err = validate(d);
    if (err != null) return err;
    try {
      final base = WorkRecord(
        id: d.id,
        start: d.start,
        end: d.end,
        dayType: 'Thường', // loại ngày thật do PayCalculator suy ra
        hourlyRate: d.rate,
      );
      final int id;
      if (d.id == null) {
        id = await _db.insert(base);
      } else {
        await _db.update(base);
        id = d.id!;
      }
      await _recalculate([...records.where((r) => r.id != id), base.withId(id)]);
      return null;
    } catch (e) {
      return 'Không thể lưu: $e';
    }
  }

  Future<String?> delete(WorkRecord r) async {
    try {
      await _db.delete(r.id!);
      await _recalculate(records.where((x) => x.id != r.id).toList());
      return null;
    } catch (e) {
      return 'Không thể xóa: $e';
    }
  }

  Future<void> updateSettings(AppSettings s) async {
    settings = s;
    await s.save();
    await _recalculate(List.of(records)); // đổi ngày lễ → tính lại
  }

  String dayTypeFor(DateTime d) => dayTypeOf(d, settings.holidays);

  Future<void> checkIn(double rate) async {
    final start = floorMinute(DateTime.now());
    final shift = ActiveShift(start: start, dayType: dayTypeFor(start), rate: rate);
    await shift.save();
    active = shift;
    notifyListeners();
  }

  Future<String?> checkOut() async {
    final shift = active;
    if (shift == null) return null;
    final end = floorMinute(DateTime.now());
    if (!end.isAfter(shift.start)) {
      return 'Ca làm cần ít nhất 1 phút. Bấm "Hủy ca" nếu bạn bấm nhầm.';
    }
    try {
      final base = WorkRecord(
        start: shift.start,
        end: end,
        dayType: shift.dayType,
        hourlyRate: shift.rate,
      );
      final id = await _db.insert(base);
      await ActiveShift.clear();
      active = null;
      await _recalculate([...records, base.withId(id)]);
      return null;
    } catch (e) {
      return 'Không thể lưu ca: $e';
    }
  }

  Future<void> discardActive() async {
    await ActiveShift.clear();
    active = null;
    notifyListeners();
  }

  BackupData buildBackup() => BackupData(
        records: records.reversed.toList(),
        hourlyRate: settings.hourlyRate,
        holidays: Set<String>.of(settings.holidays),
        lang: settings.lang,
        reminders: settings.reminders,
        exportedAt: DateTime.now(),
      );

  Future<({int added, int skipped})> importBackup(BackupData d, {required bool replace}) async {
    final plan = planImport(replace ? const <WorkRecord>[] : records, d.records);
    if (replace) {
      await _db.replaceWith(plan.add);
    } else {
      await _db.insertAll(plan.add);
    }
    settings = replace
        ? settings.copyWith(
            hourlyRate: d.hourlyRate,
            holidays: Set<String>.of(d.holidays),
            lang: d.lang,
            reminders: d.reminders,
          )
        : settings.copyWith(holidays: {...settings.holidays, ...d.holidays});
    await settings.save();
    await _recalculate(await _db.getAll());
    return (added: plan.add.length, skipped: plan.skipped);
  }
}

// ════════════════════════════════════════════════════════════════════════════
//  Xuất file: CSV, Excel (.xlsx), PDF
// ════════════════════════════════════════════════════════════════════════════

enum ExportRange { week, month, lastMonth, year, all }

enum ExportFormat { csv, xlsx, pdf }

const _rangeLabels = {
  ExportRange.week: 'Tuần này',
  ExportRange.month: 'Tháng này',
  ExportRange.lastMonth: 'Tháng trước',
  ExportRange.year: 'Năm nay',
  ExportRange.all: 'Tất cả',
};

/// (từ, đến) với đến là mốc không bao gồm; null nghĩa là tất cả.
(DateTime, DateTime)? exportBounds(ExportRange r, DateTime now) {
  switch (r) {
    case ExportRange.week:
      final w = startOfWeek(now, DateTime.monday);
      return (w, DateTime(w.year, w.month, w.day + 7));
    case ExportRange.month:
      return (DateTime(now.year, now.month), DateTime(now.year, now.month + 1));
    case ExportRange.lastMonth:
      return (DateTime(now.year, now.month - 1), DateTime(now.year, now.month));
    case ExportRange.year:
      return (DateTime(now.year), DateTime(now.year + 1));
    case ExportRange.all:
      return null;
  }
}

List<WorkRecord> recordsInRange(List<WorkRecord> all, ExportRange r) {
  final b = exportBounds(r, DateTime.now());
  if (b == null) return all;
  return [
    for (final x in all)
      if (!x.start.isBefore(b.$1) && x.start.isBefore(b.$2)) x
  ];
}

/// `newestFirst` là danh sách đã lọc, mới nhất trước.
String rangeSubtitle(ExportRange r, List<WorkRecord> newestFirst) {
  final b = exportBounds(r, DateTime.now());
  if (b != null) {
    final last = DateTime(b.$2.year, b.$2.month, b.$2.day - 1);
    return '${_dmyFmt.format(b.$1)} – ${_dmyFmt.format(last)}';
  }
  if (newestFirst.isEmpty) return '';
  return '${_dmyFmt.format(newestFirst.last.start)} – ${_dmyFmt.format(newestFirst.first.start)}';
}

/// Tổng hợp theo tuần; `asc` phải sắp xếp từ cũ đến mới.
List<(DateTime, PeriodSummary)> weeklySummaries(List<WorkRecord> asc) {
  final out = <(DateTime, PeriodSummary)>[];
  DateTime? current;
  for (final r in asc) {
    final w = startOfWeek(r.start, DateTime.monday);
    if (current != w) {
      current = w;
      out.add((w, summarize(asc, w, DateTime(w.year, w.month, w.day + 7))));
    }
  }
  return out;
}

String _weekLabel(DateTime monday) {
  final sun = DateTime(monday.year, monday.month, monday.day + 6);
  return '${_two(monday.day)}/${_two(monday.month)} – ${_two(sun.day)}/${_two(sun.month)}/${sun.year}';
}

// ── Excel ───────────────────────────────────────────────────────────────────

List<int> buildXlsx(List<WorkRecord> newestFirst) {
  final asc = newestFirst.reversed.toList();
  final book = xl.Excel.createExcel();
  book.rename('Sheet1', tr('Ca làm'));
  final shifts = book[tr('Ca làm')];
  final weeks = book[tr('Theo tuần')];

  final head = xl.CellStyle(
    bold: true,
    fontColorHex: xl.ExcelColor.white,
    backgroundColorHex: xl.ExcelColor.fromHexString('#0F766E'),
    horizontalAlign: xl.HorizontalAlign.Center,
  );
  final bold = xl.CellStyle(bold: true);

  void put(xl.Sheet sh, int col, int row, xl.CellValue v, [xl.CellStyle? st]) => sh.updateCell(
        xl.CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row),
        v,
        cellStyle: st,
      );
  xl.CellValue t(String v) => xl.TextCellValue(v);
  xl.CellValue n(double v) => xl.DoubleCellValue(double.parse(v.toStringAsFixed(2)));

  // Sheet 1: chi tiết từng ca
  final h1 = [
    tr('Ngày'), tr('Thứ'), tr('Giờ vào'), tr('Giờ ra'), tr('Loại ngày'), tr('Tổng giờ'),
    tr('Giờ thường'), tr('Giờ OT (x1.5)'), tr('Giờ x2'), tr('Lương/giờ (USD)'), tr('Tiền (USD)')
  ];
  for (var c = 0; c < h1.length; c++) {
    put(shifts, c, 0, t(h1[c]), head);
    shifts.setColumnWidth(c, c == 0 ? 14 : 13);
  }
  var row = 1;
  double sTotal = 0, sReg = 0, sOt = 0, sDbl = 0, sEarned = 0;
  for (final r in asc) {
    put(shifts, 0, row, t(_dmyFmt.format(r.start)));
    put(shifts, 1, row, t(_weekdayShort[r.start.weekday - 1]));
    put(shifts, 2, row, t(_timeFmt.format(r.start)));
    put(shifts, 3, row, t('${_timeFmt.format(r.end)}${r.overnight ? ' (+1)' : ''}'));
    put(shifts, 4, row, t(dayLabel(r.dayType)));
    put(shifts, 5, row, n(r.totalHours));
    put(shifts, 6, row, n(r.regularHours));
    put(shifts, 7, row, n(r.otHours));
    put(shifts, 8, row, n(r.doubleHours));
    put(shifts, 9, row, n(r.hourlyRate));
    put(shifts, 10, row, n(r.earned));
    sTotal += r.totalHours;
    sReg += r.regularHours;
    sOt += r.otHours;
    sDbl += r.doubleHours;
    sEarned += r.earned;
    row++;
  }
  put(shifts, 0, row, t(tr('TỔNG')), bold);
  put(shifts, 5, row, n(sTotal), bold);
  put(shifts, 6, row, n(sReg), bold);
  put(shifts, 7, row, n(sOt), bold);
  put(shifts, 8, row, n(sDbl), bold);
  put(shifts, 10, row, n(sEarned), bold);

  // Sheet 2: tổng hợp theo tuần
  final h2 = [tr('Tuần'), tr('Số ca'), tr('Tổng giờ'), tr('Giờ thường'), tr('Giờ OT (x1.5)'), tr('Giờ x2'), tr('Tiền (USD)')];
  for (var c = 0; c < h2.length; c++) {
    put(weeks, c, 0, t(h2[c]), head);
    weeks.setColumnWidth(c, c == 0 ? 24 : 14);
  }
  var wr = 1;
  for (final (monday, sum) in weeklySummaries(asc)) {
    put(weeks, 0, wr, t(_weekLabel(monday)));
    put(weeks, 1, wr, xl.IntCellValue(sum.shifts));
    put(weeks, 2, wr, n(sum.total));
    put(weeks, 3, wr, n(sum.regular));
    put(weeks, 4, wr, n(sum.ot));
    put(weeks, 5, wr, n(sum.dbl));
    put(weeks, 6, wr, n(sum.earned));
    wr++;
  }

  final bytes = book.save();
  if (bytes == null) throw StateError(tr('Không tạo được file Excel'));
  return bytes;
}

// ── PDF ─────────────────────────────────────────────────────────────────────

/// Chạy hoàn toàn offline: dùng font Liberation Sans đóng gói sẵn trong assets (có tiếng Việt).
Future<List<int>> buildPdf(List<WorkRecord> newestFirst, String title, String subtitle) async {
  final asc = newestFirst.reversed.toList();
  final regular = pw.Font.ttf(await rootBundle.load('assets/fonts/LiberationSans-Regular.ttf'));
  final boldFont = pw.Font.ttf(await rootBundle.load('assets/fonts/LiberationSans-Bold.ttf'));
  final doc = pw.Document(theme: pw.ThemeData.withFont(base: regular, bold: boldFont));

  final teal = pdf.PdfColor.fromHex('#0F766E');
  final grey = pdf.PdfColor.fromHex('#64748B');
  final light = pdf.PdfColor.fromHex('#F1F5F9');
  final sum = summarize(asc, DateTime(1970), DateTime(3000));

  pw.Widget stat(String label, String value) => pw.Expanded(
        child: pw.Container(
          padding: const pw.EdgeInsets.all(10),
          decoration: pw.BoxDecoration(color: light),
          child: pw.Column(
            crossAxisAlignment: pw.CrossAxisAlignment.start,
            children: [
              pw.Text(label, style: pw.TextStyle(fontSize: 9, color: grey)),
              pw.SizedBox(height: 3),
              pw.Text(value, style: pw.TextStyle(fontSize: 12, fontWeight: pw.FontWeight.bold)),
            ],
          ),
        ),
      );

  pw.Widget table(List<String> headers, List<List<String>> data) => pw.TableHelper.fromTextArray(
        headers: headers,
        data: data,
        headerStyle: pw.TextStyle(fontSize: 9, fontWeight: pw.FontWeight.bold, color: pdf.PdfColors.white),
        headerDecoration: pw.BoxDecoration(color: teal),
        cellStyle: const pw.TextStyle(fontSize: 9),
        oddRowDecoration: pw.BoxDecoration(color: light),
        cellPadding: const pw.EdgeInsets.symmetric(horizontal: 5, vertical: 4),
        cellAlignments: {
          0: pw.Alignment.centerLeft,
          for (var i = 1; i < headers.length; i++) i: pw.Alignment.centerRight,
        },
      );

  doc.addPage(
    pw.MultiPage(
      pageFormat: pdf.PdfPageFormat.a4,
      margin: const pw.EdgeInsets.all(32),
      footer: (ctx) => pw.Align(
        alignment: pw.Alignment.centerRight,
        child: pw.Text(
          tr('Trang ${ctx.pageNumber}/${ctx.pagesCount}'),
          style: pw.TextStyle(fontSize: 9, color: grey),
        ),
      ),
      build: (ctx) => [
        pw.Text(title, style: pw.TextStyle(fontSize: 20, fontWeight: pw.FontWeight.bold, color: teal)),
        if (subtitle.isNotEmpty) pw.Text(subtitle, style: pw.TextStyle(fontSize: 10, color: grey)),
        pw.SizedBox(height: 14),
        pw.Row(
          children: [
            stat(tr('Tổng tiền'), money(sum.earned)),
            pw.SizedBox(width: 8),
            stat(tr('Tổng giờ'), fmtHours(sum.total)),
            pw.SizedBox(width: 8),
            stat(tr('Giờ thường'), fmtHours(sum.regular)),
            pw.SizedBox(width: 8),
            stat('OT x1.5', fmtHours(sum.ot)),
            pw.SizedBox(width: 8),
            stat('x2', fmtHours(sum.dbl)),
          ],
        ),
        pw.SizedBox(height: 18),
        pw.Text(tr('Theo tuần'), style: pw.TextStyle(fontSize: 12, fontWeight: pw.FontWeight.bold)),
        pw.SizedBox(height: 6),
        table(
          [tr('Tuần'), tr('Số ca'), tr('Tổng giờ'), tr('Thường'), 'OT x1.5', 'x2', tr('Tiền')],
          [
            for (final (monday, w) in weeklySummaries(asc))
              [
                _weekLabel(monday),
                '${w.shifts}',
                fmtHours(w.total),
                fmtHours(w.regular),
                fmtHours(w.ot),
                fmtHours(w.dbl),
                money(w.earned),
              ],
          ],
        ),
        pw.SizedBox(height: 18),
        pw.Text(tr('Chi tiết ca làm'), style: pw.TextStyle(fontSize: 12, fontWeight: pw.FontWeight.bold)),
        pw.SizedBox(height: 6),
        table(
          [tr('Ngày'), tr('Vào'), tr('Ra'), tr('Loại'), tr('Thường'), 'OT x1.5', 'x2', tr('Lương/giờ'), tr('Tiền')],
          [
            for (final r in asc)
              [
                '${_weekdayShort[r.start.weekday - 1]} ${_dmyFmt.format(r.start)}',
                _timeFmt.format(r.start),
                '${_timeFmt.format(r.end)}${r.overnight ? ' (+1)' : ''}',
                dayLabel(r.dayType),
                fmtHours(r.regularHours),
                fmtHours(r.otHours),
                fmtHours(r.doubleHours),
                money(r.hourlyRate),
                money(r.earned),
              ],
          ],
        ),
      ],
    ),
  );
  return doc.save();
}

// ── Tạo file + chia sẻ ──────────────────────────────────────────────────────

Future<({List<int> bytes, String name, String mime})> buildExport(
  List<WorkRecord> newestFirst,
  ExportRange range,
  ExportFormat format,
) async {
  final stamp = _dateFmt.format(DateTime.now());
  switch (format) {
    case ExportFormat.xlsx:
      return (
        bytes: buildXlsx(newestFirst),
        name: 'bang_luong_$stamp.xlsx',
        mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      );
    case ExportFormat.pdf:
      return (
        bytes: await buildPdf(
          newestFirst,
          tr('Bảng lương – ${tr(_rangeLabels[range]!)}'),
          rangeSubtitle(range, newestFirst),
        ),
        name: 'bang_luong_$stamp.pdf',
        mime: 'application/pdf',
      );
    case ExportFormat.csv:
      return (
        bytes: utf8.encode(buildCsv(newestFirst)),
        name: 'bang_luong_$stamp.csv',
        mime: 'text/csv',
      );
  }
}

Future<void> shareBytes(List<int> bytes, String name, String mime, {String? subject}) async {
  final dir = await getTemporaryDirectory();
  final file = File(p.join(dir.path, name));
  await file.writeAsBytes(bytes, flush: true);
  // share_plus >= 11: dùng SharePlus.instance.share(ShareParams(files: [...]))
  await Share.shareXFiles([XFile(file.path, mimeType: mime)], subject: subject ?? tr('Bảng lương'));
}

class ExportSheet extends StatefulWidget {
  const ExportSheet({super.key});

  @override
  State<ExportSheet> createState() => _ExportSheetState();
}

class _ExportSheetState extends State<ExportSheet> {
  ExportRange _range = ExportRange.month;

  void _pick(ExportFormat f) => Navigator.pop(context, (range: _range, format: f));

  Widget _tile(IconData icon, Color color, String title, String subtitle, ExportFormat f) => Card(
        child: ListTile(
          leading: Container(
            padding: const EdgeInsets.all(8),
            decoration: BoxDecoration(color: color.withValues(alpha: 0.12), borderRadius: BorderRadius.circular(12)),
            child: Icon(icon, color: color),
          ),
          title: Text(title, style: const TextStyle(fontWeight: FontWeight.w700)),
          subtitle: Text(subtitle, style: const TextStyle(fontSize: 12)),
          trailing: const Icon(Icons.chevron_right_rounded),
          onTap: () => _pick(f),
        ),
      );

  @override
  Widget build(BuildContext context) => SafeArea(
        child: SingleChildScrollView(
          padding: const EdgeInsets.fromLTRB(20, 0, 20, 20),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            spacing: 12,
            children: [
              Text('Xuất dữ liệu', style: Theme.of(context).textTheme.titleLarge),
              const Text('Khoảng thời gian', style: TextStyle(color: _muted)),
              Wrap(
                spacing: 8,
                runSpacing: 4,
                children: [
                  for (final r in ExportRange.values)
                    ChoiceChip(
                      label: Text(_rangeLabels[r]!),
                      selected: _range == r,
                      onSelected: (_) => setState(() => _range = r),
                    ),
                ],
              ),
              const SizedBox(height: 4),
              const Text('Định dạng', style: TextStyle(color: _muted)),
              _tile(Icons.table_chart_rounded, Colors.green.shade700, 'Excel (.xlsx)',
                  'Chi tiết từng ca và tổng hợp theo tuần', ExportFormat.xlsx),
              _tile(Icons.picture_as_pdf_rounded, Colors.red.shade700, 'PDF',
                  'Báo cáo để in hoặc gửi (chạy offline)', ExportFormat.pdf),
              _tile(Icons.description_outlined, Colors.blueGrey, 'CSV',
                  'Văn bản thuần, mở được bằng Excel hoặc Google Sheets', ExportFormat.csv),
            ],
          ),
        ),
      );
}

// ════════════════════════════════════════════════════════════════════════════
//  UI
// ════════════════════════════════════════════════════════════════════════════

class SalaryApp extends StatefulWidget {
  const SalaryApp({super.key});
  @override
  State<SalaryApp> createState() => _SalaryAppState();
}

class _SalaryAppState extends State<SalaryApp> with WidgetsBindingObserver {
  final _controller = AppController(AppDatabase.instance);
  final _lock = AppLock(LockStore(), Biometrics());
  final _reminders = ReminderService();
  final _avatar = AvatarController(AvatarStore());
  String _reminderSig = '';

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addObserver(this);
    _controller.addListener(_onControllerChanged);
    _controller.load();
    _lock.load();
    _avatar.load();
  }

  /// Đặt lại lịch nhắc khi đổi giờ, thứ lặp lại hoặc ngôn ngữ (nội dung thông báo).
  void _onControllerChanged() {
    if (_controller.loading) return;
    final sig = '${_controller.settings.reminders.signature}|${_controller.settings.lang}';
    if (sig != _reminderSig) {
      _reminderSig = sig;
      _syncReminders();
    }
  }

  Future<void> _syncReminders() async {
    currentLang = _controller.settings.lang;
    try {
      await _reminders.sync(_controller.settings.reminders);
    } catch (_) {
      // Không cho thông báo lỗi làm hỏng ứng dụng.
    }
  }

  @override
  void didChangeAppLifecycleState(AppLifecycleState state) {
    if (state == AppLifecycleState.paused) {
      _lock.onPaused();
    } else if (state == AppLifecycleState.resumed) {
      _lock.onResumed();
      if (!_controller.loading) _syncReminders(); // làm mới lịch cho 30 ngày tới
    }
  }

  @override
  void dispose() {
    WidgetsBinding.instance.removeObserver(this);
    _controller.removeListener(_onControllerChanged);
    _controller.dispose();
    _lock.dispose();
    _avatar.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) => ListenableBuilder(
        listenable: _controller,
        builder: (context, _) {
          currentLang = _controller.settings.lang;
          return _buildApp();
        },
      );

  Widget _buildApp() => MaterialApp(
        key: ValueKey(currentLang),
        title: tr('Theo dõi lương'),
        debugShowCheckedModeBanner: false,
        theme: ThemeData(
          useMaterial3: true,
          colorScheme: ColorScheme.fromSeed(seedColor: _tealMid),
          scaffoldBackgroundColor: _surfaceBg,
          appBarTheme: const AppBarTheme(
            backgroundColor: _surfaceBg,
            surfaceTintColor: Colors.transparent,
            scrolledUnderElevation: 0,
          ),
          cardTheme: CardThemeData(
            elevation: 0,
            color: Colors.white,
            surfaceTintColor: Colors.transparent,
            margin: EdgeInsets.zero,
            shape: RoundedRectangleBorder(
              borderRadius: BorderRadius.circular(20),
              side: const BorderSide(color: Color(0xFFE6ECF0)),
            ),
          ),
        ),
        builder: (context, child) => LockGate(lock: _lock, child: child ?? const SizedBox.shrink()),
        home: HomePage(c: _controller, lock: _lock, reminders: _reminders, avatar: _avatar),
      );
}

void _toast(BuildContext context, String msg) {
  ScaffoldMessenger.of(context)
    ..hideCurrentSnackBar()
    ..showSnackBar(SnackBar(content: Text(msg), behavior: SnackBarBehavior.floating));
}

Future<bool> confirmDialog(
  BuildContext context, {
  required String title,
  required String message,
  required String action,
}) async {
  final ok = await showDialog<bool>(
    context: context,
    builder: (ctx) => AlertDialog(
      title: Text(title),
      content: Text(message),
      actions: [
        TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('Hủy')),
        FilledButton(onPressed: () => Navigator.pop(ctx, true), child: Text(action)),
      ],
    ),
  );
  return ok ?? false;
}

class HomePage extends StatelessWidget {
  const HomePage({
    super.key,
    required this.c,
    required this.lock,
    required this.reminders,
    required this.avatar,
  });
  final AppController c;
  final AppLock lock;
  final ReminderService reminders;
  final AvatarController avatar;

  @override
  Widget build(BuildContext context) {
    return ListenableBuilder(
      listenable: c,
      builder: (context, _) {
        if (c.loading) {
          return const Scaffold(body: Center(child: CircularProgressIndicator()));
        }
        if (c.loadError != null && c.records.isEmpty) {
          return Scaffold(
            appBar: AppBar(title: const Text('Theo dõi lương')),
            body: Center(
              child: Padding(
                padding: const EdgeInsets.all(24),
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  spacing: 12,
                  children: [
                    const Icon(Icons.error_outline, size: 40),
                    Text('Không thể tải dữ liệu\n${c.loadError}', textAlign: TextAlign.center),
                    FilledButton(onPressed: c.load, child: const Text('Thử lại')),
                  ],
                ),
              ),
            ),
          );
        }

        final records = c.records;
        final holidays = c.settings.holidays;
        final now = DateTime.now();
        final monday = startOfWeek(now, DateTime.monday);
        final week = summarize(records, monday, DateTime(monday.year, monday.month, monday.day + 7));
        final today = summarize(records, dateOnly(now), DateTime(now.year, now.month, now.day + 1));
        final month = summarize(records, DateTime(now.year, now.month), DateTime(now.year, now.month + 1));
        final year = summarize(records, DateTime(now.year), DateTime(now.year + 1));

        return Scaffold(
          floatingActionButton: FloatingActionButton.extended(
            onPressed: () => _editRecord(context),
            icon: const Icon(Icons.add_rounded),
            label: const Text('Thêm ca'),
          ),
          body: RefreshIndicator(
            onRefresh: c.refresh,
            child: CustomScrollView(
              physics: const AlwaysScrollableScrollPhysics(),
              slivers: [
                SliverAppBar.large(
                  leadingWidth: 68,
                  leading: Padding(
                    padding: const EdgeInsets.only(left: 16),
                    child: Center(child: AvatarButton(avatar: avatar, onTap: () => _openAvatar(context))),
                  ),
                  title: const Text('Theo dõi lương', style: TextStyle(fontWeight: FontWeight.w700)),
                  actions: [
                    IconButton(
                      tooltip: tr('Xuất dữ liệu'),
                      onPressed: () => _export(context),
                      icon: const Icon(Icons.ios_share_rounded),
                    ),
                    IconButton(
                      tooltip: tr('Cài đặt'),
                      onPressed: () => _openSettings(context),
                      icon: const Icon(Icons.tune_rounded),
                    ),
                    const SizedBox(width: 4),
                  ],
                ),
                SliverPadding(
                  padding: const EdgeInsets.fromLTRB(16, 4, 16, 0),
                  sliver: SliverToBoxAdapter(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.stretch,
                      spacing: 14,
                      children: [
                        if (c.skipped > 0)
                          Card(
                            color: Colors.amber.shade100,
                            child: Padding(
                              padding: const EdgeInsets.all(12),
                              child: Text(
                                'Có ${c.skipped} bản ghi cũ không đọc được nên chưa hiển thị. '
                                'Dữ liệu chưa bị xóa và vẫn nằm trong file DB '
                                '(có bản sao work_records_v3.db.pre_v5.bak).',
                                style: const TextStyle(fontSize: 13),
                              ),
                            ),
                          ),
                        ShiftCard(
                          active: c.active,
                          todayType: c.dayTypeFor(now),
                          weekEarned: week.earned,
                          weekHours: week.total,
                          weekShifts: week.shifts,
                          estimate: () => c.active == null
                              ? 0.0
                              : estimateActive(records, c.active!, holidays, DateTime.now()),
                          onStart: () => _startShift(context),
                          onEnd: () => _endShift(context),
                          onDiscard: () => _discardShift(context),
                        ),
                        WeekOverview(records: records, holidays: holidays),
                        Row(
                          spacing: 10,
                          children: [
                            Expanded(
                              child: StatTile(
                                icon: Icons.today_rounded,
                                label: 'Hôm nay',
                                value: money(today.earned),
                                hours: fmtHours(today.total),
                                color: _tealMid,
                              ),
                            ),
                            Expanded(
                              child: StatTile(
                                icon: Icons.calendar_month_rounded,
                                label: 'Tháng này',
                                value: money(month.earned),
                                hours: fmtHours(month.total),
                                color: Colors.indigo,
                              ),
                            ),
                            Expanded(
                              child: StatTile(
                                icon: Icons.emoji_events_outlined,
                                label: 'Năm nay',
                                value: money(year.earned),
                                hours: fmtHours(year.total),
                                color: Colors.deepOrange,
                              ),
                            ),
                          ],
                        ),
                        CompareTile(onOpen: () => _openCompare(context)),
                        HistoryPreviewCard(records: records, onOpen: () => _openHistory(context)),
                      ],
                    ),
                  ),
                ),
                const SliverToBoxAdapter(
                  child: Padding(
                    padding: EdgeInsets.fromLTRB(20, 16, 20, 100),
                    child: Text(
                      'Thu nhập chỉ là ước tính. Mỗi ca được ghi theo ngày bắt đầu '
                      '(kể cả ca qua nửa đêm). Thứ 2–6: 8 giờ đầu x1, 8–12 giờ x1.5, trên 12 giờ x2. '
                      'Thứ 7, CN: 8 giờ đầu được bù vào giờ thường còn thiếu của tuần (tối đa 40 giờ), '
                      'sau đó x1.5, trên 12 giờ x2. Ngày lễ: x2, không tính vào giờ thường; tuần có '
                      'ngày lễ rơi vào thứ 2–6 thì cuối tuần không được bù.',
                      style: TextStyle(fontSize: 12, color: _muted, height: 1.4),
                    ),
                  ),
                ),
              ],
            ),
          ),
        );
      },
    );
  }

  Future<void> _startShift(BuildContext context) async {
    final rate = await showDialog<double>(
      context: context,
      builder: (_) => StartShiftDialog(
        defaultRate: c.settings.hourlyRate,
        dayType: c.dayTypeFor(DateTime.now()),
      ),
    );
    if (rate != null) await c.checkIn(rate);
  }

  Future<void> _endShift(BuildContext context) async {
    final start = c.active?.start;
    if (start == null) return;
    final longShift = DateTime.now().difference(start).inHours >= 16;
    final err = await c.checkOut();
    if (!context.mounted) return;
    _toast(
      context,
      err ??
          (longShift
              ? 'Đã lưu. Ca dài bất thường, hãy kiểm tra lại giờ ra (chạm vào ca để sửa).'
              : 'Đã lưu ca làm.'),
    );
  }

  Future<void> _discardShift(BuildContext context) async {
    final ok = await confirmDialog(
      context,
      title: 'Hủy ca đang làm?',
      message: 'Ca này sẽ không được lưu.',
      action: 'Hủy ca',
    );
    if (ok) await c.discardActive();
  }

  Future<void> _editRecord(BuildContext context, [WorkRecord? existing]) async {
    final draft = await showModalBottomSheet<RecordDraft>(
      context: context,
      isScrollControlled: true,
      useSafeArea: true,
      showDragHandle: true,
      builder: (_) => RecordEditor(
        existing: existing,
        defaultRate: c.settings.hourlyRate,
        validate: c.validate,
        typeOf: c.dayTypeFor,
      ),
    );
    if (draft == null) return;
    final err = await c.saveDraft(draft);
    if (context.mounted) _toast(context, err ?? 'Đã lưu.');
  }

  Future<void> _deleteRecord(BuildContext context, WorkRecord r) async {
    final ok = await confirmDialog(
      context,
      title: 'Xóa ca làm?',
      message: 'Xóa ca ${_dmyFmt.format(r.start)} '
          '(${_timeFmt.format(r.start)}–${_timeFmt.format(r.end)})?',
      action: 'Xóa',
    );
    if (!ok) return;
    final err = await c.delete(r);
    if (context.mounted) _toast(context, err ?? 'Đã xóa.');
  }

  Future<void> _openSettings(BuildContext context) async {
    final r = await showDialog<({AppSettings settings, SettingsNext? then})>(
      context: context,
      builder: (_) => SettingsDialog(initial: c.settings),
    );
    if (r == null) return;
    await c.updateSettings(r.settings);
    if (!context.mounted) return;
    switch (r.then) {
      case SettingsNext.security:
        await _openSecurity(context);
      case SettingsNext.backup:
        await _openBackup(context);
      case SettingsNext.reminders:
        await _openReminders(context);
      case SettingsNext.about:
        await _openAbout(context);
      case null:
        break;
    }
  }

  Future<void> _openSecurity(BuildContext context) => showModalBottomSheet<void>(
        context: context,
        showDragHandle: true,
        isScrollControlled: true,
        builder: (_) => SecuritySheet(lock: lock),
      );

  Future<void> _openBackup(BuildContext context) => showModalBottomSheet<void>(
        context: context,
        showDragHandle: true,
        isScrollControlled: true,
        builder: (_) => BackupSheet(
          onExport: () => _exportBackup(context),
          onImportFile: kHasFilePicker ? () => _importFromFile(context) : null,
          onImportPaste: () => _importFromText(context),
        ),
      );

  Future<void> _openReminders(BuildContext context) => showModalBottomSheet<void>(
        context: context,
        showDragHandle: true,
        isScrollControlled: true,
        builder: (_) => ReminderSheet(c: c, service: reminders),
      );

  Future<void> _openAvatar(BuildContext context) => showModalBottomSheet<void>(
        context: context,
        showDragHandle: true,
        builder: (_) => AvatarSheet(
          avatar: avatar,
          onPick: () => _pickAvatar(context),
          onRemove: () => _removeAvatar(context),
        ),
      );

  Future<void> _pickAvatar(BuildContext context) async {
    if (!kHasAvatarPicker) {
      _toast(context, 'Chọn ảnh chỉ có trong bản app đầy đủ.');
      return;
    }
    final err = await lock.runExternal(avatar.pick);
    if (err != null && context.mounted) _toast(context, err);
  }

  Future<void> _removeAvatar(BuildContext context) async {
    final ok = await confirmDialog(
      context,
      title: 'Xóa ảnh đại diện?',
      message: 'Ảnh sẽ được gỡ khỏi ứng dụng.',
      action: 'Xóa ảnh',
    );
    if (ok) await avatar.remove();
  }

  Future<void> _openHistory(BuildContext context) => Navigator.of(context).push<void>(
        MaterialPageRoute(
          builder: (_) => HistoryPage(
            c: c,
            onEdit: (ctx, r) => _editRecord(ctx, r),
            onDelete: (ctx, r) => _deleteRecord(ctx, r),
          ),
        ),
      );

  Future<void> _openCompare(BuildContext context) => Navigator.of(context).push<void>(
        MaterialPageRoute(builder: (_) => ComparePage(c: c)),
      );

  Future<void> _openAbout(BuildContext context) => showDialog<void>(
        context: context,
        builder: (_) => const AboutDialogView(),
      );

  Future<void> _importFromText(BuildContext context) async {
    final text = await showDialog<String>(
      context: context,
      builder: (_) => const PasteBackupDialog(),
    );
    if (text == null || !context.mounted) return;
    await _confirmAndImport(context, text);
  }

  Future<void> _confirmAndImport(BuildContext context, String text) async {
    final BackupData data;
    try {
      data = BackupData.parse(text);
    } on FormatException catch (e) {
      _toast(context, e.message);
      return;
    }
    final mode = await showDialog<ImportMode>(
      context: context,
      builder: (_) => ImportDialog(data: data, existing: c.records.length),
    );
    if (mode == null || !context.mounted) return;
    try {
      final r = await c.importBackup(data, replace: mode == ImportMode.replace);
      if (context.mounted) _toast(context, 'Đã nhập ${r.added} ca, bỏ qua ${r.skipped} ca trùng giờ.');
    } catch (e) {
      if (context.mounted) _toast(context, 'Không thể nhập: $e');
    }
  }

  Future<void> _export(BuildContext context) async {
    if (c.records.isEmpty) {
      _toast(context, 'Chưa có dữ liệu để xuất.');
      return;
    }
    final choice = await showModalBottomSheet<({ExportRange range, ExportFormat format})>(
      context: context,
      showDragHandle: true,
      isScrollControlled: true,
      builder: (_) => const ExportSheet(),
    );
    if (choice == null || !context.mounted) return;

    final rs = recordsInRange(c.records, choice.range);
    if (rs.isEmpty) {
      _toast(context, 'Không có ca làm nào trong khoảng này.');
      return;
    }

    // Vòng quay chờ trong lúc tạo file (PDF cần tải font).
    final nav = Navigator.of(context);
    unawaited(showDialog<void>(
      context: context,
      barrierDismissible: false,
      builder: (_) => const PopScope(canPop: false, child: Center(child: CircularProgressIndicator())),
    ));
    ({List<int> bytes, String name, String mime})? file;
    String? error;
    try {
      file = await buildExport(rs, choice.range, choice.format);
    } catch (e) {
      error = '$e';
    }
    nav.pop();

    if (file == null) {
      if (context.mounted) {
        _toast(
          context,
          choice.format == ExportFormat.pdf
              ? 'Không thể tạo PDF (kiểm tra đã khai báo font trong pubspec.yaml): $error'
              : 'Không thể tạo file: $error',
        );
      }
      return;
    }
    final out = file;
    try {
      await lock.runExternal(() => shareBytes(out.bytes, out.name, out.mime));
    } catch (e) {
      if (context.mounted) _toast(context, 'Không thể chia sẻ file: $e');
    }
  }

  Future<void> _exportBackup(BuildContext context) async {
    try {
      final json = c.buildBackup().toJson();
      final name = 'sao_luu_theo_doi_luong_${_dateFmt.format(DateTime.now())}.json';
      await lock.runExternal(
        () => shareBytes(utf8.encode(json), name, 'application/json', subject: tr('Sao lưu dữ liệu')),
      );
    } catch (e) {
      if (context.mounted) _toast(context, 'Không thể xuất file sao lưu: $e');
    }
  }

  Future<void> _importFromFile(BuildContext context) async {
    try {
      final file = await lock.runExternal(() => openFile());
      if (file == null) return;
      final text = await file.readAsString();
      if (!context.mounted) return;
      await _confirmAndImport(context, text);
    } catch (e) {
      if (context.mounted) _toast(context, 'Không thể đọc file: $e');
    }
  }
}

// ── Thẻ chấm công (hero) ───────────────────────────────────────────────────

class _GlassPill extends StatelessWidget {
  const _GlassPill(this.text);
  final String text;

  @override
  Widget build(BuildContext context) => DecoratedBox(
        decoration: BoxDecoration(
          color: Colors.white.withValues(alpha: 0.18),
          borderRadius: BorderRadius.circular(20),
        ),
        child: Padding(
          padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
          child: Text(text, style: const TextStyle(color: Colors.white, fontSize: 12, fontWeight: FontWeight.w600)),
        ),
      );
}

class ShiftCard extends StatelessWidget {
  const ShiftCard({
    super.key,
    required this.active,
    required this.todayType,
    required this.weekEarned,
    required this.weekHours,
    required this.weekShifts,
    required this.estimate,
    required this.onStart,
    required this.onEnd,
    required this.onDiscard,
  });
  final ActiveShift? active;
  final String todayType;
  final double weekEarned, weekHours;
  final int weekShifts;
  final double Function() estimate;
  final VoidCallback onStart, onEnd, onDiscard;

  @override
  Widget build(BuildContext context) {
    final a = active;
    final soft = Colors.white.withValues(alpha: 0.78);
    final solid = FilledButton.styleFrom(
      backgroundColor: Colors.white,
      foregroundColor: _tealMid,
      padding: const EdgeInsets.symmetric(vertical: 16, horizontal: 24),
      textStyle: const TextStyle(fontSize: 16, fontWeight: FontWeight.w700),
    );

    return Container(
      padding: const EdgeInsets.all(20),
      decoration: BoxDecoration(
        borderRadius: BorderRadius.circular(28),
        gradient: const LinearGradient(
          begin: Alignment.topLeft,
          end: Alignment.bottomRight,
          colors: [_tealMid, _tealDark],
        ),
        boxShadow: [
          BoxShadow(
            color: _tealMid.withValues(alpha: 0.35),
            blurRadius: 24,
            offset: const Offset(0, 10),
          ),
        ],
      ),
      child: a == null
          ? Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  children: [
                    Expanded(
                      child: Text(
                        '${_weekdayFull[DateTime.now().weekday - 1]}, ${_dmyFmt.format(DateTime.now())}',
                        style: TextStyle(color: soft, fontSize: 13),
                      ),
                    ),
                    _GlassPill('Hôm nay: ${dayLabel(todayType)}'),
                  ],
                ),
                const SizedBox(height: 18),
                Text('Thu nhập tuần này', style: TextStyle(color: soft)),
                const SizedBox(height: 2),
                Text(
                  money(weekEarned),
                  style: const TextStyle(
                    color: Colors.white,
                    fontSize: 40,
                    fontWeight: FontWeight.w700,
                    letterSpacing: -0.5,
                  ),
                ),
                Text('${fmtHours(weekHours)} · $weekShifts ca', style: TextStyle(color: soft)),
                const SizedBox(height: 18),
                SizedBox(
                  width: double.infinity,
                  child: FilledButton.icon(
                    style: solid,
                    onPressed: onStart,
                    icon: const Icon(Icons.play_arrow_rounded),
                    label: const Text('Vào ca'),
                  ),
                ),
              ],
            )
          : Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  children: [
                    Container(
                      width: 10,
                      height: 10,
                      decoration: const BoxDecoration(color: Color(0xFF86EFAC), shape: BoxShape.circle),
                    ),
                    const SizedBox(width: 8),
                    const Expanded(
                      child: Text(
                        'ĐANG TRONG CA',
                        style: TextStyle(color: Colors.white, fontWeight: FontWeight.w700, letterSpacing: 1),
                      ),
                    ),
                    _GlassPill(dayLabel(a.dayType)),
                  ],
                ),
                const SizedBox(height: 12),
                ElapsedText(
                  start: a.start,
                  style: const TextStyle(color: Colors.white, fontSize: 48, fontWeight: FontWeight.w300),
                ),
                _Tick(
                  builder: (_) => Text(
                    '≈ ${money(estimate())} tính đến lúc này',
                    style: const TextStyle(color: Colors.white, fontSize: 17, fontWeight: FontWeight.w600),
                  ),
                ),
                const SizedBox(height: 4),
                Text(
                  'Vào lúc ${_fmtStart(a.start)} · ${money(a.rate)}/giờ',
                  style: TextStyle(color: soft, fontSize: 13),
                ),
                const SizedBox(height: 18),
                Row(
                  spacing: 10,
                  children: [
                    OutlinedButton(
                      style: OutlinedButton.styleFrom(
                        foregroundColor: Colors.white,
                        side: BorderSide(color: Colors.white.withValues(alpha: 0.6)),
                        padding: const EdgeInsets.symmetric(vertical: 16, horizontal: 18),
                      ),
                      onPressed: onDiscard,
                      child: const Text('Hủy ca'),
                    ),
                    Expanded(
                      child: FilledButton.icon(
                        style: solid,
                        onPressed: onEnd,
                        icon: const Icon(Icons.stop_rounded),
                        label: const Text('Ra ca'),
                      ),
                    ),
                  ],
                ),
              ],
            ),
    );
  }
}

/// Rebuild con của nó mỗi giây (không ảnh hưởng phần còn lại của màn hình).
class _Tick extends StatefulWidget {
  const _Tick({required this.builder});
  final WidgetBuilder builder;

  @override
  State<_Tick> createState() => _TickState();
}

class _TickState extends State<_Tick> {
  Timer? _timer;

  @override
  void initState() {
    super.initState();
    _timer = Timer.periodic(const Duration(seconds: 1), (_) {
      if (mounted) setState(() {});
    });
  }

  @override
  void dispose() {
    _timer?.cancel();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) => widget.builder(context);
}

class ElapsedText extends StatelessWidget {
  const ElapsedText({super.key, required this.start, this.style});
  final DateTime start;
  final TextStyle? style;

  @override
  Widget build(BuildContext context) => _Tick(
        builder: (_) {
          var d = DateTime.now().difference(start);
          if (d.isNegative) d = Duration.zero;
          return Text(
            '${_two(d.inHours)}:${_two(d.inMinutes % 60)}:${_two(d.inSeconds % 60)}',
            style: (style ?? const TextStyle())
                .copyWith(fontFeatures: const [FontFeature.tabularFigures()]),
          );
        },
      );
}

// ── Tổng quan tuần ──────────────────────────────────────────────────────────

String _shortH(double h) {
  final m = (h * 60).round();
  if (m < 60) return '${m}p';
  final hh = m / 60;
  return hh == hh.roundToDouble() ? '${hh.round()}h' : '${hh.toStringAsFixed(1)}h';
}

int _flex(double h) => math.max(1, (h * 60).round());

class _Legend extends StatelessWidget {
  const _Legend(this.color, this.text);
  final Color color;
  final String text;

  @override
  Widget build(BuildContext context) => Row(
        mainAxisSize: MainAxisSize.min,
        spacing: 6,
        children: [
          Container(width: 10, height: 10, decoration: BoxDecoration(color: color, shape: BoxShape.circle)),
          Text(text, style: const TextStyle(fontSize: 12, color: _muted)),
        ],
      );
}

class _DayBar extends StatelessWidget {
  const _DayBar({
    required this.label,
    required this.today,
    required this.holiday,
    required this.reg,
    required this.ot,
    required this.dbl,
    required this.maxH,
  });
  final String label;
  final bool today, holiday;
  final double reg, ot, dbl, maxH;

  @override
  Widget build(BuildContext context) {
    final total = reg + ot + dbl;
    const barArea = 78.0;
    final h = total <= 0 ? 4.0 : math.max(6.0, barArea * total / maxH);

    return Container(
      padding: const EdgeInsets.symmetric(vertical: 6),
      decoration: BoxDecoration(
        color: today ? _tealMid.withValues(alpha: 0.08) : null,
        borderRadius: BorderRadius.circular(14),
      ),
      child: Column(
        children: [
          SizedBox(
            height: 14,
            child: total > 0
                ? Text(_shortH(total),
                    style: const TextStyle(fontSize: 10, color: _muted, fontWeight: FontWeight.w600))
                : null,
          ),
          const SizedBox(height: 2),
          SizedBox(
            height: barArea,
            child: Align(
              alignment: Alignment.bottomCenter,
              child: AnimatedContainer(
                duration: const Duration(milliseconds: 500),
                curve: Curves.easeOutCubic,
                width: 18,
                height: h,
                clipBehavior: Clip.antiAlias,
                decoration: BoxDecoration(
                  color: total <= 0 ? const Color(0xFFE2E8F0) : null,
                  borderRadius: BorderRadius.circular(6),
                ),
                child: total <= 0
                    ? null
                    : Column(
                        children: [
                          if (dbl > 0) Expanded(flex: _flex(dbl), child: const ColoredBox(color: Colors.deepOrange)),
                          if (ot > 0) Expanded(flex: _flex(ot), child: const ColoredBox(color: Colors.indigo)),
                          if (reg > 0) Expanded(flex: _flex(reg), child: const ColoredBox(color: _tealMid)),
                        ],
                      ),
              ),
            ),
          ),
          const SizedBox(height: 6),
          Text(
            holiday ? '$label★' : label,
            style: TextStyle(
              fontSize: 12,
              fontWeight: today ? FontWeight.w700 : FontWeight.w500,
              color: holiday ? Colors.deepOrange : (today ? _tealMid : _muted),
            ),
          ),
        ],
      ),
    );
  }
}

class WeekOverview extends StatelessWidget {
  const WeekOverview({super.key, required this.records, required this.holidays});
  final List<WorkRecord> records;
  final Set<String> holidays;

  @override
  Widget build(BuildContext context) {
    final now = DateTime.now();
    final monday = startOfWeek(now, DateTime.monday);
    final sunday = DateTime(monday.year, monday.month, monday.day + 6);
    final next = DateTime(monday.year, monday.month, monday.day + 7);
    final days = [for (var i = 0; i < 7; i++) DateTime(monday.year, monday.month, monday.day + i)];

    final reg = List.filled(7, 0.0), ot = List.filled(7, 0.0), dbl = List.filled(7, 0.0);
    for (final r in records) {
      if (r.start.isBefore(monday) || !r.start.isBefore(next)) continue;
      final i = r.start.weekday - 1;
      reg[i] += r.regularHours;
      ot[i] += r.otHours;
      dbl[i] += r.doubleHours;
    }
    final maxH = math.max(
      10.0,
      [for (var i = 0; i < 7; i++) reg[i] + ot[i] + dbl[i]].fold<double>(0, (a, b) => math.max(a, b)),
    );
    final weekReg = reg.fold<double>(0, (a, b) => a + b);
    final missing = math.max(0.0, 40 - weekReg);
    final noFill = PayCalculator._weekdayHolidayIn(monday, holidays);

    final String note;
    if (missing <= 0) {
      note = 'Đã đủ 40 giờ thường. Giờ làm thêm tính OT.';
    } else if (noFill) {
      note = 'Còn thiếu ${fmtHours(missing)}. Tuần có ngày lễ T2–T6 nên cuối tuần không được bù.';
    } else {
      note = 'Còn thiếu ${fmtHours(missing)}. Thứ 7, CN có thể bù tối đa 8 giờ mỗi ngày.';
    }

    return Card(
      child: Padding(
        padding: const EdgeInsets.all(18),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                const Expanded(
                  child: Text('Tuần này', style: TextStyle(fontSize: 17, fontWeight: FontWeight.w700)),
                ),
                Text(
                  '${_two(monday.day)}/${_two(monday.month)} – ${_two(sunday.day)}/${_two(sunday.month)}',
                  style: const TextStyle(color: _muted),
                ),
              ],
            ),
            const SizedBox(height: 12),
            Row(
              spacing: 4,
              children: [
                for (var i = 0; i < 7; i++)
                  Expanded(
                    child: _DayBar(
                      label: _weekdayShort[i],
                      today: dateOnly(now) == days[i],
                      holiday: holidays.contains(_dateFmt.format(days[i])),
                      reg: reg[i],
                      ot: ot[i],
                      dbl: dbl[i],
                      maxH: maxH,
                    ),
                  ),
              ],
            ),
            const SizedBox(height: 12),
            const Wrap(
              spacing: 14,
              runSpacing: 4,
              children: [
                _Legend(_tealMid, 'Thường'),
                _Legend(Colors.indigo, 'OT x1.5'),
                _Legend(Colors.deepOrange, 'x2'),
              ],
            ),
            const SizedBox(height: 16),
            Row(
              children: [
                const Expanded(child: Text('Giờ thường trong tuần', style: TextStyle(color: _muted))),
                Text('${fmtHours(weekReg)} / 40h', style: const TextStyle(fontWeight: FontWeight.w700)),
              ],
            ),
            const SizedBox(height: 8),
            ClipRRect(
              borderRadius: BorderRadius.circular(8),
              child: TweenAnimationBuilder<double>(
                tween: Tween<double>(begin: 0, end: (weekReg / 40).clamp(0.0, 1.0)),
                duration: const Duration(milliseconds: 600),
                curve: Curves.easeOutCubic,
                builder: (_, v, __) => LinearProgressIndicator(
                  value: v,
                  minHeight: 10,
                  backgroundColor: const Color(0xFFE2E8F0),
                  color: _tealMid,
                ),
              ),
            ),
            const SizedBox(height: 8),
            Text(note, style: const TextStyle(fontSize: 12, color: _muted, height: 1.35)),
          ],
        ),
      ),
    );
  }
}

class StatTile extends StatelessWidget {
  const StatTile({
    super.key,
    required this.icon,
    required this.label,
    required this.value,
    required this.hours,
    required this.color,
  });
  final IconData icon;
  final String label, value, hours;
  final Color color;

  @override
  Widget build(BuildContext context) => Card(
        child: Padding(
          padding: const EdgeInsets.all(12),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Container(
                padding: const EdgeInsets.all(7),
                decoration: BoxDecoration(
                  color: color.withValues(alpha: 0.12),
                  borderRadius: BorderRadius.circular(10),
                ),
                child: Icon(icon, size: 18, color: color),
              ),
              const SizedBox(height: 10),
              Text(label, style: const TextStyle(fontSize: 12, color: _muted)),
              const SizedBox(height: 2),
              FittedBox(
                fit: BoxFit.scaleDown,
                alignment: Alignment.centerLeft,
                child: Text(value, style: const TextStyle(fontSize: 16, fontWeight: FontWeight.w700)),
              ),
              Text(hours, style: const TextStyle(fontSize: 12, color: _muted)),
            ],
          ),
        ),
      );
}

// ── Lịch sử ─────────────────────────────────────────────────────────────────

class WeekHeader extends StatelessWidget {
  const WeekHeader({super.key, required this.monday, required this.summary});
  final DateTime monday;
  final PeriodSummary summary;

  @override
  Widget build(BuildContext context) {
    final sunday = DateTime(monday.year, monday.month, monday.day + 6);
    return Padding(
      padding: const EdgeInsets.fromLTRB(4, 16, 4, 8),
      child: Row(
        children: [
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  'Tuần ${_two(monday.day)}/${_two(monday.month)} – ${_two(sunday.day)}/${_two(sunday.month)}',
                  style: const TextStyle(fontWeight: FontWeight.w700),
                ),
                Text(
                  'Thường ${fmtHours(summary.regular)} · OT ${fmtHours(summary.ot)} · x2 ${fmtHours(summary.dbl)}',
                  style: const TextStyle(fontSize: 12, color: _muted),
                ),
              ],
            ),
          ),
          Text(
            money(summary.earned),
            style: const TextStyle(fontWeight: FontWeight.w700, color: _tealMid, fontSize: 15),
          ),
        ],
      ),
    );
  }
}

class RecordTile extends StatelessWidget {
  const RecordTile({
    super.key,
    required this.record,
    required this.onTap,
    required this.onDelete,
  });
  final WorkRecord record;
  final VoidCallback onTap, onDelete;

  @override
  Widget build(BuildContext context) {
    final r = record;
    final color = typeColor(r.dayType);
    final range = '${_timeFmt.format(r.start)} – ${_timeFmt.format(r.end)}${r.overnight ? ' (+1)' : ''}';

    return Padding(
      padding: const EdgeInsets.only(bottom: 8),
      child: Card(
        clipBehavior: Clip.antiAlias,
        child: InkWell(
          onTap: onTap,
          child: Padding(
            padding: const EdgeInsets.fromLTRB(12, 12, 4, 12),
            child: Row(
              children: [
                Container(
                  width: 50,
                  height: 56,
                  decoration: BoxDecoration(
                    color: color.withValues(alpha: 0.12),
                    borderRadius: BorderRadius.circular(14),
                  ),
                  child: Column(
                    mainAxisAlignment: MainAxisAlignment.center,
                    children: [
                      Text(
                        _weekdayShort[r.start.weekday - 1],
                        style: TextStyle(fontSize: 11, color: color, fontWeight: FontWeight.w600),
                      ),
                      Text(
                        '${r.start.day}',
                        style: TextStyle(fontSize: 20, color: color, fontWeight: FontWeight.w700),
                      ),
                    ],
                  ),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Row(
                        children: [
                          Flexible(
                            child: Text(range, style: const TextStyle(fontWeight: FontWeight.w700)),
                          ),
                          if (r.dayType != 'Thường') ...[
                            const SizedBox(width: 6),
                            _Pill(dayLabel(r.dayType), color: color),
                          ],
                        ],
                      ),
                      const SizedBox(height: 6),
                      Wrap(
                        spacing: 6,
                        runSpacing: 4,
                        children: [
                          if (r.regularHours > 0) _Pill('Thường ${fmtHours(r.regularHours)}', color: _tealMid),
                          if (r.otHours > 0) _Pill('OT ${fmtHours(r.otHours)}', color: Colors.indigo),
                          if (r.doubleHours > 0) _Pill('x2 ${fmtHours(r.doubleHours)}', color: Colors.deepOrange),
                          _Pill('${money(r.hourlyRate)}/giờ', color: Colors.blueGrey),
                        ],
                      ),
                    ],
                  ),
                ),
                Column(
                  crossAxisAlignment: CrossAxisAlignment.end,
                  children: [
                    Text(money(r.earned), style: const TextStyle(fontWeight: FontWeight.w700, fontSize: 16)),
                    Text(fmtHours(r.totalHours), style: const TextStyle(fontSize: 12, color: _muted)),
                  ],
                ),
                PopupMenuButton<String>(
                  icon: const Icon(Icons.more_vert, size: 20, color: _muted),
                  onSelected: (v) => v == 'edit' ? onTap() : onDelete(),
                  itemBuilder: (_) => const [
                    PopupMenuItem(value: 'edit', child: Text('Sửa')),
                    PopupMenuItem(value: 'delete', child: Text('Xóa')),
                  ],
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}

class _Pill extends StatelessWidget {
  const _Pill(this.text, {required this.color});
  final String text;
  final Color color;

  @override
  Widget build(BuildContext context) => DecoratedBox(
        decoration: BoxDecoration(
          color: color.withValues(alpha: 0.1),
          borderRadius: BorderRadius.circular(20),
        ),
        child: Padding(
          padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 3),
          child: Text(text, style: TextStyle(fontSize: 12, color: color, fontWeight: FontWeight.w500)),
        ),
      );
}

// ── Form & hộp thoại ────────────────────────────────────────────────────────

class StartShiftDialog extends StatefulWidget {
  const StartShiftDialog({super.key, required this.defaultRate, required this.dayType});
  final double defaultRate;
  final String dayType;

  @override
  State<StartShiftDialog> createState() => _StartShiftDialogState();
}

class _StartShiftDialogState extends State<StartShiftDialog> {
  late final TextEditingController _rate =
      TextEditingController(text: _rateFmt.format(widget.defaultRate));
  String? _error;

  @override
  void dispose() {
    _rate.dispose();
    super.dispose();
  }

  void _submit() {
    final rate = parseNumber(_rate.text);
    if (rate == null) {
      setState(() => _error = 'Mức lương không hợp lệ.');
      return;
    }
    Navigator.pop(context, rate);
  }

  @override
  Widget build(BuildContext context) => AlertDialog(
        title: const Text('Bắt đầu ca làm'),
        content: SingleChildScrollView(
          child: Column(
            mainAxisSize: MainAxisSize.min,
            crossAxisAlignment: CrossAxisAlignment.start,
            spacing: 16,
            children: [
              _Pill('Hôm nay tính là: ${dayLabel(widget.dayType)}', color: typeColor(widget.dayType)),
              TextField(
                controller: _rate,
                keyboardType: const TextInputType.numberWithOptions(decimal: true),
                decoration: InputDecoration(
                  labelText: tr('Lương theo giờ'),
                  prefixText: r'$ ',
                  errorText: _error == null ? null : tr(_error!),
                  border: const OutlineInputBorder(),
                ),
              ),
            ],
          ),
        ),
        actions: [
          TextButton(onPressed: () => Navigator.pop(context), child: const Text('Hủy')),
          FilledButton(onPressed: _submit, child: const Text('Vào ca')),
        ],
      );
}

// ── Khóa ứng dụng: màn hình & hộp thoại ─────────────────────────────────────

enum SettingsNext { security, backup, reminders, about }

/// Phủ màn hình khóa lên trên toàn bộ ứng dụng (kể cả hộp thoại đang mở).
class LockGate extends StatelessWidget {
  const LockGate({super.key, required this.lock, required this.child});
  final AppLock lock;
  final Widget child;

  @override
  Widget build(BuildContext context) => ListenableBuilder(
        listenable: lock,
        builder: (context, _) {
          final covered = !lock.ready || lock.needsSetup || lock.locked;
          return Stack(
            fit: StackFit.expand,
            children: [
              Offstage(offstage: covered, child: child),
              if (covered) Positioned.fill(child: Overlay.wrap(child: _LockSwitcher(lock: lock))),
            ],
          );
        },
      );
}

class _LockSwitcher extends StatelessWidget {
  const _LockSwitcher({required this.lock});
  final AppLock lock;

  @override
  Widget build(BuildContext context) => ListenableBuilder(
        listenable: lock,
        builder: (context, _) {
          if (!lock.ready) {
            return const Scaffold(body: Center(child: CircularProgressIndicator()));
          }
          if (lock.needsSetup) return SetupLockScreen(lock: lock);
          return LockScreen(lock: lock);
        },
      );
}

class _PasswordField extends StatelessWidget {
  const _PasswordField({
    required this.controller,
    required this.label,
    required this.hide,
    required this.onToggle,
    this.onSubmitted,
    this.autofocus = false,
  });
  final TextEditingController controller;
  final String label;
  final bool hide;
  final VoidCallback onToggle;
  final VoidCallback? onSubmitted;
  final bool autofocus;

  @override
  Widget build(BuildContext context) => TextField(
        controller: controller,
        obscureText: hide,
        autofocus: autofocus,
        autocorrect: false,
        enableSuggestions: false,
        onSubmitted: onSubmitted == null ? null : (_) => onSubmitted!(),
        decoration: InputDecoration(
          labelText: tr(label),
          border: const OutlineInputBorder(),
          suffixIcon: IconButton(
            icon: Icon(hide ? Icons.visibility_outlined : Icons.visibility_off_outlined),
            onPressed: onToggle,
          ),
        ),
      );
}

class _LockBadge extends StatelessWidget {
  const _LockBadge({required this.icon});
  final IconData icon;

  @override
  Widget build(BuildContext context) => Center(
        child: Container(
          padding: const EdgeInsets.all(18),
          decoration: const BoxDecoration(
            shape: BoxShape.circle,
            gradient: LinearGradient(
              begin: Alignment.topLeft,
              end: Alignment.bottomRight,
              colors: [_tealMid, _tealDark],
            ),
          ),
          child: Icon(icon, size: 36, color: Colors.white),
        ),
      );
}

class LockScreen extends StatefulWidget {
  const LockScreen({super.key, required this.lock});
  final AppLock lock;

  @override
  State<LockScreen> createState() => _LockScreenState();
}

class _LockScreenState extends State<LockScreen> {
  final _pw = TextEditingController();
  bool _hide = true;
  bool _wrong = false;
  bool _wasBlocked = false;
  Timer? _timer;

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addPostFrameCallback((_) => widget.lock.unlockWithBiometric());
    _timer = Timer.periodic(const Duration(seconds: 1), (_) {
      final blocked = widget.lock.blockedSeconds > 0;
      if (mounted && (blocked || _wasBlocked)) {
        _wasBlocked = blocked;
        setState(() {});
      }
    });
  }

  @override
  void dispose() {
    _timer?.cancel();
    _pw.dispose();
    super.dispose();
  }

  void _submit() {
    final err = widget.lock.unlock(_pw.text);
    if (err == null) return; // mở khóa thành công: LockGate tự gỡ màn hình này
    _pw.clear();
    setState(() {
      _wrong = widget.lock.blockedSeconds == 0;
      _wasBlocked = widget.lock.blockedSeconds > 0;
    });
  }

  @override
  Widget build(BuildContext context) {
    final lock = widget.lock;
    final blocked = lock.blockedSeconds;
    final errorColor = Theme.of(context).colorScheme.error;
    return Scaffold(
      backgroundColor: _surfaceBg,
      body: SafeArea(
        child: Center(
          child: SingleChildScrollView(
            padding: const EdgeInsets.all(24),
            child: ConstrainedBox(
              constraints: const BoxConstraints(maxWidth: 380),
              child: Column(
                mainAxisSize: MainAxisSize.min,
                crossAxisAlignment: CrossAxisAlignment.stretch,
                spacing: 16,
                children: [
                  const _LockBadge(icon: Icons.lock_rounded),
                  const Text(
                    'Theo dõi lương',
                    textAlign: TextAlign.center,
                    style: TextStyle(fontSize: 22, fontWeight: FontWeight.w700),
                  ),
                  const Text(
                    'Nhập mật khẩu để mở ứng dụng',
                    textAlign: TextAlign.center,
                    style: TextStyle(color: _muted),
                  ),
                  _PasswordField(
                    controller: _pw,
                    label: 'Mật khẩu',
                    hide: _hide,
                    onToggle: () => setState(() => _hide = !_hide),
                    onSubmitted: _submit,
                  ),
                  if (blocked > 0)
                    Text(
                      'Nhập sai quá nhiều lần. Thử lại sau $blocked giây.',
                      textAlign: TextAlign.center,
                      style: TextStyle(color: errorColor),
                    )
                  else if (_wrong)
                    Text('Mật khẩu không đúng.', textAlign: TextAlign.center, style: TextStyle(color: errorColor)),
                  FilledButton(
                    onPressed: blocked > 0 ? null : _submit,
                    style: FilledButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 16)),
                    child: const Text('Mở khóa'),
                  ),
                  if (lock.biometricOn && lock.biometricAvailable)
                    OutlinedButton.icon(
                      onPressed: lock.unlockWithBiometric,
                      icon: const Icon(Icons.fingerprint_rounded),
                      label: const Text('Dùng vân tay / khuôn mặt'),
                    ),
                ],
              ),
            ),
          ),
        ),
      ),
    );
  }
}

class SetupLockScreen extends StatefulWidget {
  const SetupLockScreen({super.key, required this.lock});
  final AppLock lock;

  @override
  State<SetupLockScreen> createState() => _SetupLockScreenState();
}

class _SetupLockScreenState extends State<SetupLockScreen> {
  final _pw = TextEditingController();
  final _confirm = TextEditingController();
  bool _hide = true;
  bool _busy = false;
  String? _error;

  @override
  void dispose() {
    _pw.dispose();
    _confirm.dispose();
    super.dispose();
  }

  Future<void> _submit() async {
    final err = validateNewPassword(_pw.text, _confirm.text);
    if (err != null) {
      setState(() => _error = err);
      return;
    }
    setState(() {
      _busy = true;
      _error = null;
    });
    await widget.lock.setPassword(_pw.text);
  }

  @override
  Widget build(BuildContext context) => Scaffold(
        backgroundColor: _surfaceBg,
        body: SafeArea(
          child: Center(
            child: SingleChildScrollView(
              padding: const EdgeInsets.all(24),
              child: ConstrainedBox(
                constraints: const BoxConstraints(maxWidth: 380),
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  crossAxisAlignment: CrossAxisAlignment.stretch,
                  spacing: 16,
                  children: [
                    const _LockBadge(icon: Icons.shield_outlined),
                    const Text(
                      'Bảo vệ ứng dụng bằng mật khẩu',
                      textAlign: TextAlign.center,
                      style: TextStyle(fontSize: 22, fontWeight: FontWeight.w700),
                    ),
                    const Text(
                      'Tạo mật khẩu để người khác không mở được dữ liệu chấm công của bạn.',
                      textAlign: TextAlign.center,
                      style: TextStyle(color: _muted),
                    ),
                    _PasswordField(
                      controller: _pw,
                      label: 'Mật khẩu mới',
                      hide: _hide,
                      onToggle: () => setState(() => _hide = !_hide),
                    ),
                    _PasswordField(
                      controller: _confirm,
                      label: 'Nhập lại mật khẩu',
                      hide: _hide,
                      onToggle: () => setState(() => _hide = !_hide),
                      onSubmitted: _busy ? null : _submit,
                    ),
                    if (_error != null)
                      Text(
                        _error!,
                        textAlign: TextAlign.center,
                        style: TextStyle(color: Theme.of(context).colorScheme.error),
                      ),
                    FilledButton(
                      onPressed: _busy ? null : _submit,
                      style: FilledButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 16)),
                      child: const Text('Tạo mật khẩu'),
                    ),
                    TextButton(
                      onPressed: _busy ? null : widget.lock.skipSetup,
                      child: const Text('Bỏ qua, bật sau trong Cài đặt'),
                    ),
                  ],
                ),
              ),
            ),
          ),
        ),
      );
}

/// Hộp thoại nhập mật khẩu dùng chung: bật khóa, đổi mật khẩu, tắt khóa.
/// `onSubmit` trả về thông báo lỗi (hộp thoại giữ nguyên để nhập lại) hoặc null nếu thành công.
class PasswordDialog extends StatefulWidget {
  const PasswordDialog({
    super.key,
    required this.title,
    required this.onSubmit,
    this.askCurrent = false,
    this.askNew = false,
  });
  final String title;
  final bool askCurrent, askNew;
  final Future<String?> Function(String current, String newPassword) onSubmit;

  @override
  State<PasswordDialog> createState() => _PasswordDialogState();
}

class _PasswordDialogState extends State<PasswordDialog> {
  final _cur = TextEditingController();
  final _new = TextEditingController();
  final _confirm = TextEditingController();
  bool _hide = true;
  bool _busy = false;
  String? _error;

  @override
  void dispose() {
    _cur.dispose();
    _new.dispose();
    _confirm.dispose();
    super.dispose();
  }

  Future<void> _go() async {
    String? err;
    if (widget.askCurrent && _cur.text.isEmpty) err = 'Nhập mật khẩu hiện tại.';
    if (err == null && widget.askNew) err = validateNewPassword(_new.text, _confirm.text);
    if (err == null) {
      setState(() => _busy = true);
      err = await widget.onSubmit(_cur.text, _new.text);
    }
    if (!mounted) return;
    if (err == null) {
      Navigator.pop(context, true);
    } else {
      setState(() {
        _error = err;
        _busy = false;
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    void toggle() => setState(() => _hide = !_hide);
    return AlertDialog(
      title: Text(widget.title),
      content: SingleChildScrollView(
        child: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.stretch,
          spacing: 12,
          children: [
            if (widget.askCurrent)
              _PasswordField(
                controller: _cur,
                label: 'Mật khẩu hiện tại',
                hide: _hide,
                onToggle: toggle,
                autofocus: true,
                onSubmitted: widget.askNew ? null : _go,
              ),
            if (widget.askNew) ...[
              _PasswordField(
                controller: _new,
                label: 'Mật khẩu mới',
                hide: _hide,
                onToggle: toggle,
                autofocus: !widget.askCurrent,
              ),
              _PasswordField(
                controller: _confirm,
                label: 'Nhập lại mật khẩu',
                hide: _hide,
                onToggle: toggle,
                onSubmitted: _go,
              ),
            ],
            if (_error != null)
              Text(_error!, style: TextStyle(color: Theme.of(context).colorScheme.error)),
          ],
        ),
      ),
      actions: [
        TextButton(onPressed: _busy ? null : () => Navigator.pop(context), child: const Text('Hủy')),
        FilledButton(onPressed: _busy ? null : _go, child: const Text('Xác nhận')),
      ],
    );
  }
}

class _NavTile extends StatelessWidget {
  const _NavTile({required this.icon, required this.title, this.subtitle, required this.onTap});
  final IconData icon;
  final String title;
  final String? subtitle;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) => ListTile(
        contentPadding: EdgeInsets.zero,
        leading: Container(
          padding: const EdgeInsets.all(8),
          decoration: BoxDecoration(
            color: _tealMid.withValues(alpha: 0.1),
            borderRadius: BorderRadius.circular(12),
          ),
          child: Icon(icon, color: _tealMid),
        ),
        title: Text(title, style: const TextStyle(fontWeight: FontWeight.w600)),
        subtitle: subtitle == null ? null : Text(subtitle!, style: const TextStyle(fontSize: 12)),
        trailing: const Icon(Icons.chevron_right_rounded),
        onTap: onTap,
      );
}

class SecuritySheet extends StatefulWidget {
  const SecuritySheet({super.key, required this.lock});
  final AppLock lock;

  @override
  State<SecuritySheet> createState() => _SecuritySheetState();
}

class _SecuritySheetState extends State<SecuritySheet> {
  String? _msg;
  AppLock get lock => widget.lock;

  Future<void> _toggleLock(bool on) async {
    final ok = await showDialog<bool>(
      context: context,
      builder: (_) => on
          ? PasswordDialog(
              title: 'Tạo mật khẩu',
              askNew: true,
              onSubmit: (cur, nw) async {
                await lock.setPassword(nw);
                return null;
              },
            )
          : PasswordDialog(
              title: 'Tắt khóa ứng dụng',
              askCurrent: true,
              onSubmit: (cur, nw) => lock.disable(cur),
            ),
    );
    if (!mounted || ok != true) return;
    setState(() => _msg = on ? 'Đã bật khóa ứng dụng.' : 'Đã tắt khóa ứng dụng.');
  }

  Future<void> _change() async {
    final ok = await showDialog<bool>(
      context: context,
      builder: (_) => PasswordDialog(
        title: 'Đổi mật khẩu',
        askCurrent: true,
        askNew: true,
        onSubmit: (cur, nw) => lock.changePassword(cur, nw),
      ),
    );
    if (!mounted || ok != true) return;
    setState(() => _msg = 'Đã đổi mật khẩu.');
  }

  Future<void> _toggleBio(bool on) async {
    final ok = await lock.setBiometric(on);
    if (!mounted) return;
    setState(() => _msg = (!ok && on) ? 'Không xác thực được nên chưa bật mở khóa sinh trắc học.' : null);
  }

  @override
  Widget build(BuildContext context) => ListenableBuilder(
        listenable: lock,
        builder: (context, _) => SafeArea(
          child: SingleChildScrollView(
            padding: const EdgeInsets.fromLTRB(20, 0, 20, 20),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              spacing: 8,
              children: [
                Text('Bảo mật & mật khẩu', style: Theme.of(context).textTheme.titleLarge),
                SwitchListTile(
                  contentPadding: EdgeInsets.zero,
                  title: const Text('Khóa ứng dụng bằng mật khẩu'),
                  value: lock.enabled,
                  onChanged: _toggleLock,
                ),
                if (lock.enabled)
                  ListTile(
                    contentPadding: EdgeInsets.zero,
                    leading: const Icon(Icons.password_rounded),
                    title: const Text('Đổi mật khẩu'),
                    trailing: const Icon(Icons.chevron_right_rounded),
                    onTap: _change,
                  ),
                if (lock.enabled && lock.biometricAvailable)
                  SwitchListTile(
                    contentPadding: EdgeInsets.zero,
                    title: const Text('Mở khóa bằng vân tay / khuôn mặt'),
                    value: lock.biometricOn,
                    onChanged: _toggleBio,
                  ),
                if (lock.enabled && !lock.biometricAvailable)
                  const Text(
                    'Thiết bị chưa hỗ trợ hoặc chưa đăng ký vân tay / khuôn mặt.',
                    style: TextStyle(fontSize: 12, color: _muted),
                  ),
                if (_msg != null) Text(_msg!, style: const TextStyle(color: _tealMid, fontWeight: FontWeight.w600)),
                const SizedBox(height: 4),
                const Text(
                  'Nếu quên mật khẩu, bạn phải gỡ và cài lại ứng dụng rồi nhập dữ liệu từ file sao lưu. '
                  'Hãy xuất file sao lưu thường xuyên.',
                  style: TextStyle(fontSize: 12, color: _muted, height: 1.4),
                ),
              ],
            ),
          ),
        ),
      );
}

// ── Sao lưu & khôi phục: giao diện ──────────────────────────────────────────

class BackupSheet extends StatelessWidget {
  const BackupSheet({
    super.key,
    required this.onExport,
    required this.onImportFile,
    required this.onImportPaste,
  });
  final VoidCallback onExport, onImportPaste;
  final VoidCallback? onImportFile;

  @override
  Widget build(BuildContext context) {
    void run(VoidCallback action) {
      Navigator.pop(context);
      action();
    }

    return SafeArea(
      child: SingleChildScrollView(
        padding: const EdgeInsets.fromLTRB(20, 0, 20, 20),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          spacing: 4,
          children: [
            Text('Sao lưu & khôi phục dữ liệu', style: Theme.of(context).textTheme.titleLarge),
            const SizedBox(height: 4),
            _NavTile(
              icon: Icons.ios_share_rounded,
              title: 'Xuất file sao lưu (.json)',
              subtitle: 'Lưu toàn bộ ca làm và cài đặt để nhập lại khi cài lại ứng dụng.',
              onTap: () => run(onExport),
            ),
            if (onImportFile != null)
              _NavTile(
                icon: Icons.folder_open_rounded,
                title: 'Nhập từ file sao lưu',
                onTap: () => run(onImportFile!),
              ),
            _NavTile(
              icon: Icons.content_paste_rounded,
              title: 'Nhập từ nội dung đã sao chép',
              onTap: () => run(onImportPaste),
            ),
          ],
        ),
      ),
    );
  }
}

class ImportDialog extends StatelessWidget {
  const ImportDialog({super.key, required this.data, required this.existing});
  final BackupData data;
  final int existing;

  @override
  Widget build(BuildContext context) {
    final rs = data.records;
    String range = '';
    if (rs.isNotEmpty) {
      var first = rs.first.start, last = rs.first.start;
      for (final r in rs) {
        if (r.start.isBefore(first)) first = r.start;
        if (r.start.isAfter(last)) last = r.start;
      }
      range = '${_dmyFmt.format(first)} – ${_dmyFmt.format(last)}';
    }
    final exported = '${_dmyFmt.format(data.exportedAt)} ${_timeFmt.format(data.exportedAt)}';
    return AlertDialog(
      title: const Text('Nhập dữ liệu'),
      content: SingleChildScrollView(
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          mainAxisSize: MainAxisSize.min,
          spacing: 6,
          children: [
            Text('Tìm thấy ${rs.length} ca làm', style: const TextStyle(fontWeight: FontWeight.w700)),
            if (range.isNotEmpty) Text(range, style: const TextStyle(color: _muted)),
            Text('Xuất lúc $exported', style: const TextStyle(color: _muted)),
            Text('Hiện có $existing ca trong ứng dụng.'),
            const SizedBox(height: 8),
            FilledButton(
              onPressed: () => Navigator.pop(context, ImportMode.merge),
              child: const Text('Gộp vào dữ liệu hiện có'),
            ),
            const Text('Ca trùng giờ với ca đã có sẽ bị bỏ qua.', style: TextStyle(fontSize: 12, color: _muted)),
            const SizedBox(height: 6),
            OutlinedButton(
              style: OutlinedButton.styleFrom(foregroundColor: Theme.of(context).colorScheme.error),
              onPressed: () => Navigator.pop(context, ImportMode.replace),
              child: const Text('Thay thế toàn bộ'),
            ),
            const Text(
              'Xóa hết dữ liệu hiện tại và dùng dữ liệu trong file (kể cả cài đặt).',
              style: TextStyle(fontSize: 12, color: _muted),
            ),
          ],
        ),
      ),
      actions: [
        TextButton(onPressed: () => Navigator.pop(context), child: const Text('Hủy')),
      ],
    );
  }
}

class PasteBackupDialog extends StatefulWidget {
  const PasteBackupDialog({super.key});

  @override
  State<PasteBackupDialog> createState() => _PasteBackupDialogState();
}

class _PasteBackupDialogState extends State<PasteBackupDialog> {
  final _text = TextEditingController();
  String? _error;

  @override
  void dispose() {
    _text.dispose();
    super.dispose();
  }

  Future<void> _paste() async {
    final data = await Clipboard.getData(Clipboard.kTextPlain);
    if (data?.text != null && mounted) setState(() => _text.text = data!.text!);
  }

  @override
  Widget build(BuildContext context) => AlertDialog(
        title: const Text('Nhập từ nội dung đã sao chép'),
        content: SingleChildScrollView(
          child: Column(
            mainAxisSize: MainAxisSize.min,
            crossAxisAlignment: CrossAxisAlignment.stretch,
            spacing: 8,
            children: [
              TextField(
                controller: _text,
                minLines: 4,
                maxLines: 8,
                decoration: InputDecoration(
                  hintText: tr('Dán nội dung file sao lưu vào đây'),
                  errorText: _error == null ? null : tr(_error!),
                  border: const OutlineInputBorder(),
                ),
              ),
              Align(
                alignment: Alignment.centerRight,
                child: TextButton.icon(
                  onPressed: _paste,
                  icon: const Icon(Icons.content_paste_rounded),
                  label: const Text('Dán'),
                ),
              ),
            ],
          ),
        ),
        actions: [
          TextButton(onPressed: () => Navigator.pop(context), child: const Text('Hủy')),
          FilledButton(
            onPressed: () {
              if (_text.text.trim().isEmpty) {
                setState(() => _error = 'Chưa có nội dung để nhập.');
              } else {
                Navigator.pop(context, _text.text);
              }
            },
            child: const Text('Tiếp tục'),
          ),
        ],
      );
}

class AboutDialogView extends StatelessWidget {
  const AboutDialogView({super.key});

  @override
  Widget build(BuildContext context) => AlertDialog(
        title: const Text('Thông tin'),
        content: const Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.start,
          spacing: 10,
          children: [
            Text('Phần Mềm Chấm Công', style: TextStyle(fontSize: 20, fontWeight: FontWeight.w700)),
            Text('Tác giả dự án: Trung Nguyen'),
            Text('Với sự hỗ trợ code của AI: Claude - Sonnet'),
            Text('Điện thoại liên hệ: 408-455-0327'),
          ],
        ),
        actions: [
          TextButton(onPressed: () => Navigator.pop(context), child: const Text('Đóng')),
        ],
      );
}

// ── Ảnh đại diện ────────────────────────────────────────────────────────────

class AvatarButton extends StatelessWidget {
  const AvatarButton({super.key, required this.avatar, required this.onTap});
  final AvatarController avatar;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) => ListenableBuilder(
        listenable: avatar,
        builder: (context, _) => InkWell(
          customBorder: const CircleBorder(),
          onTap: onTap,
          child: Container(
            padding: const EdgeInsets.all(2),
            decoration: BoxDecoration(shape: BoxShape.circle, border: Border.all(color: _tealMid, width: 2)),
            child: CircleAvatar(
              radius: 20,
              backgroundColor: _tealMid.withValues(alpha: 0.12),
              backgroundImage: avatar.image,
              child: avatar.image == null ? const Icon(Icons.person_rounded, color: _tealMid) : null,
            ),
          ),
        ),
      );
}

class AvatarSheet extends StatelessWidget {
  const AvatarSheet({super.key, required this.avatar, required this.onPick, required this.onRemove});
  final AvatarController avatar;
  final VoidCallback onPick, onRemove;

  @override
  Widget build(BuildContext context) => ListenableBuilder(
        listenable: avatar,
        builder: (context, _) {
          final has = avatar.image != null;
          void run(VoidCallback action) {
            Navigator.pop(context);
            action();
          }

          return SafeArea(
            child: SingleChildScrollView(
              padding: const EdgeInsets.fromLTRB(20, 0, 20, 20),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.stretch,
                spacing: 8,
                children: [
                  Center(
                    child: CircleAvatar(
                      radius: 52,
                      backgroundColor: _tealMid.withValues(alpha: 0.12),
                      backgroundImage: avatar.image,
                      child: has ? null : const Icon(Icons.person_rounded, size: 52, color: _tealMid),
                    ),
                  ),
                  const SizedBox(height: 4),
                  const Center(
                    child: Text('Ảnh đại diện', style: TextStyle(fontSize: 18, fontWeight: FontWeight.w700)),
                  ),
                  _NavTile(
                    icon: Icons.photo_library_outlined,
                    title: has ? 'Đổi ảnh' : 'Chọn ảnh từ điện thoại',
                    onTap: () => run(onPick),
                  ),
                  if (has)
                    _NavTile(icon: Icons.delete_outline_rounded, title: 'Xóa ảnh', onTap: () => run(onRemove)),
                ],
              ),
            ),
          );
        },
      );
}

// ── Nhắc giờ vào ca / ra ca ─────────────────────────────────────────────────

class _AlarmRow extends StatelessWidget {
  const _AlarmRow({
    required this.icon,
    required this.title,
    required this.minutes,
    required this.onToggle,
    required this.onPickTime,
  });
  final IconData icon;
  final String title;
  final int? minutes;
  final ValueChanged<bool> onToggle;
  final VoidCallback onPickTime;

  @override
  Widget build(BuildContext context) {
    final on = minutes != null;
    final text = on ? '${_two(minutes! ~/ 60)}:${_two(minutes! % 60)}' : '--:--';
    return Card(
      child: Padding(
        padding: const EdgeInsets.fromLTRB(16, 12, 8, 12),
        child: Row(
          children: [
            Container(
              padding: const EdgeInsets.all(8),
              decoration: BoxDecoration(color: _tealMid.withValues(alpha: 0.1), borderRadius: BorderRadius.circular(12)),
              child: Icon(icon, color: _tealMid),
            ),
            const SizedBox(width: 12),
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(title, style: const TextStyle(color: _muted)),
                  InkWell(
                    onTap: onPickTime,
                    child: Text(
                      text,
                      style: TextStyle(
                        fontSize: 34,
                        fontWeight: FontWeight.w300,
                        color: on ? null : _muted,
                        fontFeatures: const [FontFeature.tabularFigures()],
                      ),
                    ),
                  ),
                ],
              ),
            ),
            Switch(value: on, onChanged: onToggle),
          ],
        ),
      ),
    );
  }
}

class ReminderSheet extends StatefulWidget {
  const ReminderSheet({super.key, required this.c, required this.service});
  final AppController c;
  final ReminderService service;

  @override
  State<ReminderSheet> createState() => _ReminderSheetState();
}

class _ReminderSheetState extends State<ReminderSheet> {
  String? _msg;
  ReminderConfig get cfg => widget.c.settings.reminders;

  Future<void> _apply(ReminderConfig next, {bool permission = false}) async {
    if (permission && kHasReminders) {
      final ok = await widget.service.requestPermission();
      if (!ok && mounted) {
        setState(() => _msg = 'Chưa được cấp quyền thông báo. Hãy bật thông báo cho ứng dụng trong Cài đặt Android.');
      }
    }
    await widget.c.updateSettings(widget.c.settings.copyWith(reminders: next));
  }

  Future<void> _pickTime(bool isStart) async {
    final current = (isStart ? cfg.startMin : cfg.endMin) ?? (isStart ? 8 * 60 : 17 * 60);
    final t = await showTimePicker(
      context: context,
      initialTime: TimeOfDay(hour: current ~/ 60, minute: current % 60),
    );
    if (t == null) return;
    final m = t.hour * 60 + t.minute;
    await _apply(isStart ? cfg.copyWith(startMin: m) : cfg.copyWith(endMin: m), permission: true);
  }

  Future<void> _toggle(bool isStart, bool on) async {
    if (on) {
      final m = (isStart ? cfg.startMin : cfg.endMin) ?? (isStart ? 8 * 60 : 17 * 60);
      await _apply(isStart ? cfg.copyWith(startMin: m) : cfg.copyWith(endMin: m), permission: true);
    } else {
      await _apply(isStart ? cfg.copyWith(clearStart: true) : cfg.copyWith(clearEnd: true));
    }
  }

  Future<void> _toggleDay(int day, bool on) {
    final d = Set<int>.of(cfg.days);
    if (on) {
      d.add(day);
    } else {
      d.remove(day);
    }
    return _apply(cfg.copyWith(days: d));
  }

  Future<void> _test() async {
    final ok = await widget.service.requestPermission();
    if (!ok) {
      if (mounted) {
        setState(() => _msg = 'Chưa được cấp quyền thông báo. Hãy bật thông báo cho ứng dụng trong Cài đặt Android.');
      }
      return;
    }
    await widget.service.showTest();
  }

  @override
  Widget build(BuildContext context) => ListenableBuilder(
        listenable: widget.c,
        builder: (context, _) => SafeArea(
          child: SingleChildScrollView(
            padding: const EdgeInsets.fromLTRB(20, 0, 20, 20),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              spacing: 10,
              children: [
                Text('Nhắc giờ vào ca / ra ca', style: Theme.of(context).textTheme.titleLarge),
                if (!kHasReminders)
                  const Text(
                    'Nhắc giờ chỉ hoạt động trong app Android (bản đầy đủ), không dùng được trên DartPad.',
                    style: TextStyle(color: _muted),
                  ),
                _AlarmRow(
                  icon: Icons.login_rounded,
                  title: 'Nhắc vào ca',
                  minutes: cfg.startMin,
                  onToggle: (v) => _toggle(true, v),
                  onPickTime: () => _pickTime(true),
                ),
                _AlarmRow(
                  icon: Icons.logout_rounded,
                  title: 'Nhắc ra ca',
                  minutes: cfg.endMin,
                  onToggle: (v) => _toggle(false, v),
                  onPickTime: () => _pickTime(false),
                ),
                const Text('Lặp lại vào các ngày', style: TextStyle(color: _muted)),
                Wrap(
                  spacing: 6,
                  children: [
                    for (var i = 1; i <= 7; i++)
                      FilterChip(
                        label: Text(_weekdayShort[i - 1]),
                        selected: cfg.days.contains(i),
                        onSelected: (v) => _toggleDay(i, v),
                      ),
                  ],
                ),
                if (_msg != null) Text(_msg!, style: TextStyle(color: Theme.of(context).colorScheme.error)),
                if (kHasReminders)
                  OutlinedButton.icon(
                    onPressed: _test,
                    icon: const Icon(Icons.notifications_active_outlined),
                    label: const Text('Gửi thông báo thử'),
                  ),
                const Text(
                  'Thông báo hiện trên thanh trạng thái và phát chuông báo thức, kể cả khi đã tắt ứng dụng. '
                  'Nếu không nghe chuông, hãy kiểm tra âm lượng báo thức, cho phép thông báo và không hạn chế '
                  'pin cho ứng dụng trong Cài đặt Android. Lịch nhắc được làm mới mỗi khi bạn mở ứng dụng '
                  '(đặt trước tối đa 30 ngày).',
                  style: TextStyle(fontSize: 12, color: _muted, height: 1.4),
                ),
              ],
            ),
          ),
        ),
      );
}

// ── Lịch sử ca làm (trang riêng, có lọc theo khoảng ngày) ──────────────────

class _MetricCell extends StatelessWidget {
  const _MetricCell(this.label, this.value);
  final String label, value;

  @override
  Widget build(BuildContext context) => Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Text(label, style: const TextStyle(color: _muted, fontSize: 12)),
          const SizedBox(height: 2),
          Text(value, style: const TextStyle(fontWeight: FontWeight.w600)),
        ],
      );
}

class RangeSummaryCard extends StatelessWidget {
  const RangeSummaryCard({super.key, required this.summary, required this.rangeText});
  final PeriodSummary summary;
  final String rangeText;

  @override
  Widget build(BuildContext context) => Card(
        child: Padding(
          padding: const EdgeInsets.all(16),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            spacing: 10,
            children: [
              Row(
                children: [
                  const Expanded(
                    child: Text('Tổng kết', style: TextStyle(fontSize: 16, fontWeight: FontWeight.w700)),
                  ),
                  Flexible(child: Text(rangeText, style: const TextStyle(color: _muted, fontSize: 12))),
                ],
              ),
              Text(money(summary.earned), style: const TextStyle(fontSize: 30, fontWeight: FontWeight.w700)),
              Wrap(
                spacing: 22,
                runSpacing: 10,
                children: [
                  _MetricCell('Số ca', '${summary.shifts}'),
                  _MetricCell('Tổng giờ', fmtHours(summary.total)),
                  _MetricCell('Thường', fmtHours(summary.regular)),
                  _MetricCell('OT x1.5', fmtHours(summary.ot)),
                  _MetricCell('x2', fmtHours(summary.dbl)),
                ],
              ),
            ],
          ),
        ),
      );
}

typedef RecordAction = void Function(BuildContext context, WorkRecord record);

enum HistoryPreset { all, week, month, lastMonth, year, custom }

class HistoryPage extends StatefulWidget {
  const HistoryPage({super.key, required this.c, required this.onEdit, required this.onDelete});
  final AppController c;
  final RecordAction onEdit, onDelete;

  @override
  State<HistoryPage> createState() => _HistoryPageState();
}

class _HistoryPageState extends State<HistoryPage> {
  HistoryPreset _preset = HistoryPreset.all;
  DateTimeRange? _custom;

  static const _presetLabels = {
    HistoryPreset.all: 'Tất cả',
    HistoryPreset.week: 'Tuần này',
    HistoryPreset.month: 'Tháng này',
    HistoryPreset.lastMonth: 'Tháng trước',
    HistoryPreset.year: 'Năm nay',
  };

  /// (từ, đến) với đến là mốc không bao gồm; null = không lọc.
  (DateTime, DateTime)? get _bounds {
    final now = DateTime.now();
    switch (_preset) {
      case HistoryPreset.all:
        return null;
      case HistoryPreset.week:
        return periodBounds(CompareKind.week, now);
      case HistoryPreset.month:
        return periodBounds(CompareKind.month, now);
      case HistoryPreset.lastMonth:
        return periodBounds(CompareKind.month, DateTime(now.year, now.month - 1));
      case HistoryPreset.year:
        return periodBounds(CompareKind.year, now);
      case HistoryPreset.custom:
        final r = _custom;
        if (r == null) return null;
        return (dateOnly(r.start), DateTime(r.end.year, r.end.month, r.end.day + 1));
    }
  }

  Future<void> _pickRange() async {
    final now = DateTime.now();
    final r = await showDateRangePicker(
      context: context,
      firstDate: DateTime(2020),
      lastDate: DateTime(now.year + 2, 12, 31),
      initialDateRange: _custom,
      helpText: tr('Chọn khoảng ngày'),
      saveText: tr('Áp dụng'),
    );
    if (r != null && mounted) {
      setState(() {
        _custom = r;
        _preset = HistoryPreset.custom;
      });
    }
  }

  String _rangeText(List<WorkRecord> rs, (DateTime, DateTime)? b) {
    if (b != null) {
      final last = DateTime(b.$2.year, b.$2.month, b.$2.day - 1);
      return '${_dmyFmt.format(b.$1)} – ${_dmyFmt.format(last)}';
    }
    if (rs.isEmpty) return '';
    return '${_dmyFmt.format(rs.last.start)} – ${_dmyFmt.format(rs.first.start)}';
  }

  @override
  Widget build(BuildContext context) => ListenableBuilder(
        listenable: widget.c,
        builder: (context, _) {
          final all = widget.c.records;
          final b = _bounds;
          final rs = b == null
              ? all
              : [
                  for (final r in all)
                    if (!r.start.isBefore(b.$1) && r.start.isBefore(b.$2)) r
                ];
          final sum = summarize(rs, DateTime(1970), DateTime(3000));

          // Gộp theo tuần: DateTime = tiêu đề tuần, WorkRecord = ca làm.
          final items = <Object>[];
          DateTime? currentWeek;
          for (final r in rs) {
            final w = startOfWeek(r.start, DateTime.monday);
            if (currentWeek != w) {
              currentWeek = w;
              items.add(w);
            }
            items.add(r);
          }

          final customLabel = (_preset == HistoryPreset.custom && _custom != null)
              ? '${_dmyFmt.format(_custom!.start)} – ${_dmyFmt.format(_custom!.end)}'
              : 'Chọn khoảng ngày';

          return Scaffold(
            appBar: AppBar(title: const Text('Lịch sử ca làm')),
            body: CustomScrollView(
              slivers: [
                SliverPadding(
                  padding: const EdgeInsets.fromLTRB(16, 4, 16, 8),
                  sliver: SliverToBoxAdapter(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.stretch,
                      spacing: 12,
                      children: [
                        Wrap(
                          spacing: 8,
                          runSpacing: 4,
                          children: [
                            for (final e in _presetLabels.entries)
                              ChoiceChip(
                                label: Text(e.value),
                                selected: _preset == e.key,
                                onSelected: (_) => setState(() => _preset = e.key),
                              ),
                            ChoiceChip(
                              avatar: const Icon(Icons.date_range_rounded, size: 18),
                              label: Text(customLabel),
                              selected: _preset == HistoryPreset.custom,
                              onSelected: (_) => _pickRange(),
                            ),
                          ],
                        ),
                        RangeSummaryCard(summary: sum, rangeText: _rangeText(rs, b)),
                      ],
                    ),
                  ),
                ),
                if (rs.isEmpty)
                  const SliverToBoxAdapter(
                    child: Padding(
                      padding: EdgeInsets.all(32),
                      child: Center(child: Text('Không có ca làm nào trong khoảng này.')),
                    ),
                  )
                else
                  SliverPadding(
                    padding: const EdgeInsets.symmetric(horizontal: 16),
                    sliver: SliverList.builder(
                      itemCount: items.length,
                      itemBuilder: (context, i) {
                        final it = items[i];
                        if (it is DateTime) {
                          return WeekHeader(
                            monday: it,
                            summary: summarize(rs, it, DateTime(it.year, it.month, it.day + 7)),
                          );
                        }
                        final r = it as WorkRecord;
                        return RecordTile(
                          key: ValueKey(r.id),
                          record: r,
                          onTap: () => widget.onEdit(context, r),
                          onDelete: () => widget.onDelete(context, r),
                        );
                      },
                    ),
                  ),
                const SliverToBoxAdapter(child: SizedBox(height: 40)),
              ],
            ),
          );
        },
      );
}

/// Thẻ trên trang chủ: bấm vào để mở trang lịch sử đầy đủ.
class HistoryPreviewCard extends StatelessWidget {
  const HistoryPreviewCard({super.key, required this.records, required this.onOpen});
  final List<WorkRecord> records;
  final VoidCallback onOpen;

  @override
  Widget build(BuildContext context) {
    final recent = records.take(3).toList();
    return Card(
      clipBehavior: Clip.antiAlias,
      child: InkWell(
        onTap: onOpen,
        child: Padding(
          padding: const EdgeInsets.all(16),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            spacing: 10,
            children: [
              Row(
                children: [
                  const Expanded(
                    child: Text('Lịch sử ca làm', style: TextStyle(fontSize: 20, fontWeight: FontWeight.w700)),
                  ),
                  Text('${records.length} ca', style: const TextStyle(color: _muted)),
                  const Icon(Icons.chevron_right_rounded, color: _muted),
                ],
              ),
              if (recent.isEmpty)
                const Text(
                  'Bấm “Vào ca” để chấm công hoặc thêm một ca thủ công.',
                  style: TextStyle(color: _muted),
                )
              else
                for (final r in recent)
                  Row(
                    children: [
                      Container(
                        width: 40,
                        height: 44,
                        decoration: BoxDecoration(
                          color: typeColor(r.dayType).withValues(alpha: 0.12),
                          borderRadius: BorderRadius.circular(12),
                        ),
                        child: Column(
                          mainAxisAlignment: MainAxisAlignment.center,
                          children: [
                            Text(
                              _weekdayShort[r.start.weekday - 1],
                              style: TextStyle(fontSize: 10, color: typeColor(r.dayType)),
                            ),
                            Text(
                              '${r.start.day}',
                              style: TextStyle(
                                fontSize: 16,
                                fontWeight: FontWeight.w700,
                                color: typeColor(r.dayType),
                              ),
                            ),
                          ],
                        ),
                      ),
                      const SizedBox(width: 12),
                      Expanded(
                        child: Text(
                          '${_timeFmt.format(r.start)} – ${_timeFmt.format(r.end)}${r.overnight ? ' (+1)' : ''}',
                          style: const TextStyle(fontWeight: FontWeight.w600),
                        ),
                      ),
                      Text(money(r.earned), style: const TextStyle(fontWeight: FontWeight.w700)),
                    ],
                  ),
              const Align(
                alignment: Alignment.centerRight,
                child: Text('Xem tất cả', style: TextStyle(color: _tealMid, fontWeight: FontWeight.w700)),
              ),
            ],
          ),
        ),
      ),
    );
  }
}

// ── So sánh thu nhập ────────────────────────────────────────────────────────

class CompareTile extends StatelessWidget {
  const CompareTile({super.key, required this.onOpen});
  final VoidCallback onOpen;

  @override
  Widget build(BuildContext context) => Card(
        clipBehavior: Clip.antiAlias,
        child: InkWell(
          onTap: onOpen,
          child: Padding(
            padding: const EdgeInsets.all(14),
            child: Row(
              spacing: 12,
              children: [
                Container(
                  padding: const EdgeInsets.all(10),
                  decoration: BoxDecoration(
                    color: Colors.indigo.withValues(alpha: 0.12),
                    borderRadius: BorderRadius.circular(12),
                  ),
                  child: const Icon(Icons.compare_arrows_rounded, color: Colors.indigo),
                ),
                const Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text('So sánh thu nhập', style: TextStyle(fontWeight: FontWeight.w700, fontSize: 16)),
                      Text(
                        'So sánh tuần, tháng hoặc năm với một kỳ khác',
                        style: TextStyle(color: _muted, fontSize: 12),
                      ),
                    ],
                  ),
                ),
                const Icon(Icons.chevron_right_rounded, color: _muted),
              ],
            ),
          ),
        ),
      );
}

class MonthPickerDialog extends StatefulWidget {
  const MonthPickerDialog({super.key, required this.initial});
  final DateTime initial;

  @override
  State<MonthPickerDialog> createState() => _MonthPickerDialogState();
}

class _MonthPickerDialogState extends State<MonthPickerDialog> {
  late int _year = widget.initial.year;

  @override
  Widget build(BuildContext context) {
    final maxYear = DateTime.now().year + 2;
    return AlertDialog(
      title: Row(
        children: [
          IconButton(
            onPressed: _year > 2020 ? () => setState(() => _year--) : null,
            icon: const Icon(Icons.chevron_left_rounded),
          ),
          Expanded(child: Center(child: Text('$_year'))),
          IconButton(
            onPressed: _year < maxYear ? () => setState(() => _year++) : null,
            icon: const Icon(Icons.chevron_right_rounded),
          ),
        ],
      ),
      content: Wrap(
        alignment: WrapAlignment.center,
        spacing: 8,
        runSpacing: 8,
        children: [
          for (var m = 1; m <= 12; m++)
            ChoiceChip(
              label: Text(_monthShort[m - 1]),
              selected: _year == widget.initial.year && m == widget.initial.month,
              onSelected: (_) => Navigator.pop(context, DateTime(_year, m)),
            ),
        ],
      ),
      actions: [
        TextButton(onPressed: () => Navigator.pop(context), child: const Text('Hủy')),
      ],
    );
  }
}

class YearPickerDialog extends StatelessWidget {
  const YearPickerDialog({super.key, required this.initial});
  final DateTime initial;

  @override
  Widget build(BuildContext context) => AlertDialog(
        title: const Text('Năm'),
        content: SizedBox(
          width: 300,
          height: 300,
          child: YearPicker(
            firstDate: DateTime(2020),
            lastDate: DateTime(DateTime.now().year + 2),
            selectedDate: initial,
            onChanged: (d) => Navigator.pop(context, DateTime(d.year)),
          ),
        ),
        actions: [
          TextButton(onPressed: () => Navigator.pop(context), child: const Text('Hủy')),
        ],
      );
}

class _PeriodButton extends StatelessWidget {
  const _PeriodButton({required this.tag, required this.color, required this.title, required this.onTap});
  final String tag, title;
  final Color color;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) => Card(
        clipBehavior: Clip.antiAlias,
        child: InkWell(
          onTap: onTap,
          child: Padding(
            padding: const EdgeInsets.all(12),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              spacing: 4,
              children: [
                Row(
                  spacing: 6,
                  children: [
                    Container(width: 10, height: 10, decoration: BoxDecoration(color: color, shape: BoxShape.circle)),
                    Text(tag, style: TextStyle(color: color, fontWeight: FontWeight.w700, fontSize: 12)),
                  ],
                ),
                Text(title, style: const TextStyle(fontWeight: FontWeight.w600, fontSize: 13)),
              ],
            ),
          ),
        ),
      );
}

String _dHours(double a, double b) {
  final m = ((a - b) * 60).round();
  if (m == 0) return '0h00';
  return '${m > 0 ? '+' : '-'}${fmtHours(m.abs() / 60)}';
}

String _dMoney(double a, double b) {
  final d = a - b;
  if (d.abs() < 0.005) return money(0);
  return '${d > 0 ? '+' : '-'}${money(d.abs())}';
}

String _dPct(double a, double b) {
  if (b.abs() < 0.005) return a.abs() < 0.005 ? '' : '—';
  return '${a >= b ? '+' : '-'}${((a - b).abs() / b * 100).toStringAsFixed(1)}%';
}

Color _dColor(double d) => d > 0.0001 ? Colors.green.shade700 : (d < -0.0001 ? Colors.red.shade700 : _muted);

class ComparePage extends StatefulWidget {
  const ComparePage({super.key, required this.c});
  final AppController c;

  @override
  State<ComparePage> createState() => _ComparePageState();
}

class _ComparePageState extends State<ComparePage> {
  CompareKind _kind = CompareKind.month;
  DateTime _a = DateTime.now();
  DateTime _b = shiftPeriod(CompareKind.month, DateTime.now(), -1);

  void _setKind(CompareKind k) {
    final now = DateTime.now();
    setState(() {
      _kind = k;
      _a = now;
      _b = shiftPeriod(k, now, -1);
    });
  }

  Future<DateTime?> _choose(DateTime current) {
    switch (_kind) {
      case CompareKind.week:
        return showDatePicker(
          context: context,
          initialDate: current,
          firstDate: DateTime(2020),
          lastDate: DateTime(DateTime.now().year + 2, 12, 31),
          helpText: tr('Chọn một ngày trong tuần cần xem'),
        );
      case CompareKind.month:
        return showDialog<DateTime>(context: context, builder: (_) => MonthPickerDialog(initial: current));
      case CompareKind.year:
        return showDialog<DateTime>(context: context, builder: (_) => YearPickerDialog(initial: current));
    }
  }

  Future<void> _pick(bool isA) async {
    final d = await _choose(isA ? _a : _b);
    if (d == null || !mounted) return;
    setState(() {
      if (isA) {
        _a = d;
      } else {
        _b = d;
      }
    });
  }

  Widget _bar(String tag, double value, double max, Color color) => Row(
        spacing: 8,
        children: [
          SizedBox(width: 18, child: Text(tag, style: TextStyle(color: color, fontWeight: FontWeight.w700))),
          Expanded(
            child: ClipRRect(
              borderRadius: BorderRadius.circular(6),
              child: LinearProgressIndicator(
                value: max <= 0 ? 0 : (value / max).clamp(0.0, 1.0),
                minHeight: 12,
                color: color,
                backgroundColor: const Color(0xFFE2E8F0),
              ),
            ),
          ),
          SizedBox(
            width: 84,
            child: Text(money(value), textAlign: TextAlign.right, style: const TextStyle(fontWeight: FontWeight.w600)),
          ),
        ],
      );

  Widget _row(String label, String a, String b, String delta, Color deltaColor, {bool strong = false, String sub = ''}) {
    final style = TextStyle(fontWeight: strong ? FontWeight.w700 : FontWeight.w500, fontSize: strong ? 14 : 13);
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 7),
      child: Row(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Expanded(flex: 5, child: Text(label, style: TextStyle(color: _muted, fontSize: strong ? 14 : 13))),
          Expanded(flex: 5, child: Text(a, textAlign: TextAlign.right, style: style)),
          Expanded(flex: 5, child: Text(b, textAlign: TextAlign.right, style: style)),
          Expanded(
            flex: 5,
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.end,
              children: [
                Text(delta, style: style.copyWith(color: deltaColor)),
                if (sub.isNotEmpty) Text(sub, style: TextStyle(fontSize: 11, color: deltaColor)),
              ],
            ),
          ),
        ],
      ),
    );
  }

  @override
  Widget build(BuildContext context) => ListenableBuilder(
        listenable: widget.c,
        builder: (context, _) {
          final rs = widget.c.records;
          final ba = periodBounds(_kind, _a);
          final bb = periodBounds(_kind, _b);
          final sa = summarize(rs, ba.$1, ba.$2);
          final sb = summarize(rs, bb.$1, bb.$2);
          const colorA = _tealMid;
          const colorB = Colors.indigo;
          final same = ba.$1 == bb.$1;
          final maxPay = math.max(sa.earned, sb.earned);

          return Scaffold(
            appBar: AppBar(title: const Text('So sánh thu nhập')),
            body: ListView(
              padding: const EdgeInsets.fromLTRB(16, 4, 16, 40),
              children: [
                SegmentedButton<CompareKind>(
                  showSelectedIcon: false,
                  segments: const [
                    ButtonSegment(value: CompareKind.week, label: Text('Tuần')),
                    ButtonSegment(value: CompareKind.month, label: Text('Tháng')),
                    ButtonSegment(value: CompareKind.year, label: Text('Năm')),
                  ],
                  selected: {_kind},
                  onSelectionChanged: (v) => _setKind(v.first),
                ),
                const SizedBox(height: 12),
                Row(
                  spacing: 8,
                  children: [
                    Expanded(
                      child: _PeriodButton(
                        tag: 'Kỳ A',
                        color: colorA,
                        title: periodTitle(_kind, _a),
                        onTap: () => _pick(true),
                      ),
                    ),
                    IconButton(
                      tooltip: tr('Đổi chỗ hai kỳ'),
                      onPressed: () => setState(() {
                        final t = _a;
                        _a = _b;
                        _b = t;
                      }),
                      icon: const Icon(Icons.swap_horiz_rounded),
                    ),
                    Expanded(
                      child: _PeriodButton(
                        tag: 'Kỳ B',
                        color: colorB,
                        title: periodTitle(_kind, _b),
                        onTap: () => _pick(false),
                      ),
                    ),
                  ],
                ),
                if (same)
                  const Padding(
                    padding: EdgeInsets.only(top: 8),
                    child: Text(
                      'Hai kỳ đang giống nhau, hãy chọn kỳ khác để so sánh.',
                      style: TextStyle(color: _muted, fontSize: 12),
                    ),
                  ),
                const SizedBox(height: 8),
                Card(
                  child: Padding(
                    padding: const EdgeInsets.all(16),
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.stretch,
                      children: [
                        _bar('A', sa.earned, maxPay, colorA),
                        const SizedBox(height: 8),
                        _bar('B', sb.earned, maxPay, colorB),
                        const Divider(height: 24),
                        Row(
                          children: [
                            const Spacer(flex: 5),
                            Expanded(
                              flex: 5,
                              child: Text('Kỳ A',
                                  textAlign: TextAlign.right,
                                  style: TextStyle(color: colorA, fontWeight: FontWeight.w700, fontSize: 12)),
                            ),
                            Expanded(
                              flex: 5,
                              child: Text('Kỳ B',
                                  textAlign: TextAlign.right,
                                  style: TextStyle(color: colorB, fontWeight: FontWeight.w700, fontSize: 12)),
                            ),
                            const Expanded(
                              flex: 5,
                              child: Text('Chênh lệch',
                                  textAlign: TextAlign.right,
                                  style: TextStyle(color: _muted, fontWeight: FontWeight.w700, fontSize: 12)),
                            ),
                          ],
                        ),
                        _row('Lương', money(sa.earned), money(sb.earned), _dMoney(sa.earned, sb.earned),
                            _dColor(sa.earned - sb.earned),
                            strong: true, sub: _dPct(sa.earned, sb.earned)),
                        _row('Tổng giờ', fmtHours(sa.total), fmtHours(sb.total), _dHours(sa.total, sb.total),
                            _dColor(sa.total - sb.total)),
                        _row('Thường', fmtHours(sa.regular), fmtHours(sb.regular), _dHours(sa.regular, sb.regular),
                            _dColor(sa.regular - sb.regular)),
                        _row('OT x1.5', fmtHours(sa.ot), fmtHours(sb.ot), _dHours(sa.ot, sb.ot),
                            _dColor(sa.ot - sb.ot)),
                        _row('x2', fmtHours(sa.dbl), fmtHours(sb.dbl), _dHours(sa.dbl, sb.dbl),
                            _dColor(sa.dbl - sb.dbl)),
                        _row('Số ca', '${sa.shifts}', '${sb.shifts}',
                            '${sa.shifts - sb.shifts > 0 ? '+' : ''}${sa.shifts - sb.shifts}',
                            _dColor((sa.shifts - sb.shifts).toDouble())),
                      ],
                    ),
                  ),
                ),
              ],
            ),
          );
        },
      );
}

class SettingsDialog extends StatefulWidget {
  const SettingsDialog({super.key, required this.initial});
  final AppSettings initial;

  @override
  State<SettingsDialog> createState() => _SettingsDialogState();
}

class _SettingsDialogState extends State<SettingsDialog> {
  late final TextEditingController _rate =
      TextEditingController(text: _rateFmt.format(widget.initial.hourlyRate));
  late final Set<String> _holidays = {...widget.initial.holidays};
  late String _lang = widget.initial.lang;
  String? _error;

  @override
  void dispose() {
    _rate.dispose();
    super.dispose();
  }

  Future<void> _addHoliday() async {
    final d = await showDatePicker(
      context: context,
      initialDate: DateTime.now(),
      firstDate: DateTime(2020),
      lastDate: DateTime(2100),
      helpText: tr('Chọn ngày lễ'),
    );
    if (d != null) setState(() => _holidays.add(_dateFmt.format(d)));
  }

  String _label(String key) {
    final d = DateTime.tryParse(key);
    return d == null ? key : '${_weekdayShort[d.weekday - 1]}, ${_dmyFmt.format(d)}';
  }

  void _finish(SettingsNext? next) {
    final rate = parseNumber(_rate.text);
    if (rate == null) {
      setState(() => _error = 'Mức lương không hợp lệ.');
      return;
    }
    Navigator.pop(
      context,
      (settings: widget.initial.copyWith(hourlyRate: rate, holidays: _holidays, lang: _lang), then: next),
    );
  }

  @override
  Widget build(BuildContext context) {
    final sorted = _holidays.toList()..sort();
    return AlertDialog(
      title: const Text('Cài đặt'),
      content: SingleChildScrollView(
        child: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.start,
          spacing: 14,
          children: [
            const Text('Ngôn ngữ / Language', style: TextStyle(fontWeight: FontWeight.w600)),
            SegmentedButton<String>(
              showSelectedIcon: false,
              segments: [
                for (final e in supportedLangs.entries) ButtonSegment(value: e.key, label: Text(e.value)),
              ],
              selected: {_lang},
              onSelectionChanged: (v) => setState(() => _lang = v.first),
            ),
            TextField(
              controller: _rate,
              keyboardType: const TextInputType.numberWithOptions(decimal: true),
              decoration: InputDecoration(
                labelText: tr('Lương mặc định (mỗi giờ)'),
                prefixText: r'$ ',
                errorText: _error == null ? null : tr(_error!),
                border: const OutlineInputBorder(),
              ),
            ),
            const Text('Ngày lễ (tính x2)', style: TextStyle(fontWeight: FontWeight.w600)),
            if (sorted.isEmpty)
              const Text('Chưa đánh dấu ngày lễ nào.',
                  style: TextStyle(color: Colors.black54))
            else
              Wrap(
                spacing: 8,
                runSpacing: 4,
                children: [
                  for (final h in sorted)
                    InputChip(
                      label: Text(_label(h)),
                      onDeleted: () => setState(() => _holidays.remove(h)),
                    ),
                ],
              ),
            OutlinedButton.icon(
              onPressed: _addHoliday,
              icon: const Icon(Icons.event_available),
              label: const Text('Thêm ngày lễ'),
            ),
            const Text(
              'Tuần làm việc bắt đầu thứ 2. Thứ 2–6 là ngày thường, thứ 7 và CN là ngày OT '
              '(tự động). Lương mặc định chỉ áp dụng cho ca mới.',
              style: TextStyle(fontSize: 12, color: Colors.black54),
            ),
            const Divider(height: 8),
            _NavTile(
              icon: Icons.lock_outline_rounded,
              title: 'Bảo mật & mật khẩu',
              onTap: () => _finish(SettingsNext.security),
            ),
            _NavTile(
              icon: Icons.alarm_rounded,
              title: 'Nhắc giờ vào ca / ra ca',
              onTap: () => _finish(SettingsNext.reminders),
            ),
            _NavTile(
              icon: Icons.backup_outlined,
              title: 'Sao lưu & khôi phục dữ liệu',
              onTap: () => _finish(SettingsNext.backup),
            ),
            _NavTile(
              icon: Icons.info_outline_rounded,
              title: 'Thông tin',
              onTap: () => _finish(SettingsNext.about),
            ),
          ],
        ),
      ),
      actions: [
        TextButton(onPressed: () => Navigator.pop(context), child: const Text('Hủy')),
        FilledButton(onPressed: () => _finish(null), child: const Text('Lưu')),
      ],
    );
  }
}

class RecordEditor extends StatefulWidget {
  const RecordEditor({
    super.key,
    required this.existing,
    required this.defaultRate,
    required this.validate,
    required this.typeOf,
  });
  final WorkRecord? existing;
  final double defaultRate;
  final String? Function(RecordDraft) validate;
  final String Function(DateTime) typeOf;

  @override
  State<RecordEditor> createState() => _RecordEditorState();
}

class _RecordEditorState extends State<RecordEditor> {
  late DateTime _date;
  late TimeOfDay _in;
  late TimeOfDay _out;
  late final TextEditingController _rate;
  String? _error;

  @override
  void initState() {
    super.initState();
    final e = widget.existing;
    _date = dateOnly(e?.start ?? DateTime.now());
    _in = e != null ? TimeOfDay.fromDateTime(e.start) : const TimeOfDay(hour: 9, minute: 0);
    _out = e != null ? TimeOfDay.fromDateTime(e.end) : const TimeOfDay(hour: 17, minute: 0);
    _rate = TextEditingController(text: _rateFmt.format(e?.hourlyRate ?? widget.defaultRate));
  }

  @override
  void dispose() {
    _rate.dispose();
    super.dispose();
  }

  DateTime get _startDt =>
      DateTime(_date.year, _date.month, _date.day, _in.hour, _in.minute);

  DateTime get _endDt {
    final e = DateTime(_date.year, _date.month, _date.day, _out.hour, _out.minute);
    return e.isAfter(_startDt)
        ? e
        : DateTime(e.year, e.month, e.day + 1, e.hour, e.minute);
  }

  String _tod(TimeOfDay t) => '${_two(t.hour)}:${_two(t.minute)}';

  Future<void> _pickDate() async {
    final d = await showDatePicker(
      context: context,
      initialDate: _date,
      firstDate: DateTime(2020),
      lastDate: DateTime(2100),
    );
    if (d != null) setState(() {
      _date = dateOnly(d);
      _error = null;
    });
  }

  Future<void> _pickTime(bool isIn) async {
    final t = await showTimePicker(context: context, initialTime: isIn ? _in : _out);
    if (t == null) return;
    setState(() {
      if (isIn) {
        _in = t;
      } else {
        _out = t;
      }
      _error = null;
    });
  }

  void _save() {
    final rate = parseNumber(_rate.text);
    if (rate == null) {
      setState(() => _error = 'Mức lương không hợp lệ.');
      return;
    }
    if (_in == _out) {
      setState(() => _error = 'Giờ ra phải khác giờ vào.');
      return;
    }
    final draft = RecordDraft(
      id: widget.existing?.id,
      start: _startDt,
      end: _endDt,
      rate: rate,
    );
    final err = widget.validate(draft);
    if (err != null) {
      setState(() => _error = err);
      return;
    }
    Navigator.pop(context, draft);
  }

  @override
  Widget build(BuildContext context) {
    final cs = Theme.of(context).colorScheme;
    final minutes = _in == _out ? 0 : _endDt.difference(_startDt).inMinutes;
    final overnight = dateOnly(_endDt) != dateOnly(_startDt);

    return Padding(
      padding: EdgeInsets.fromLTRB(20, 0, 20, 20 + MediaQuery.viewInsetsOf(context).bottom),
      child: SingleChildScrollView(
        child: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.stretch,
          spacing: 14,
          children: [
            Text(
              widget.existing == null ? 'Thêm ca làm' : 'Sửa ca làm',
              style: Theme.of(context).textTheme.titleLarge,
            ),
            OutlinedButton.icon(
              onPressed: _pickDate,
              icon: const Icon(Icons.event),
              label: Text(
                  '${_weekdayShort[_date.weekday - 1]}, ${_dmyFmt.format(_date)}'),
            ),
            Row(
              spacing: 10,
              children: [
                Expanded(
                  child: OutlinedButton.icon(
                    onPressed: () => _pickTime(true),
                    icon: const Icon(Icons.login),
                    label: Text('Vào ${_tod(_in)}'),
                  ),
                ),
                Expanded(
                  child: OutlinedButton.icon(
                    onPressed: () => _pickTime(false),
                    icon: const Icon(Icons.logout),
                    label: Text('Ra ${_tod(_out)}'),
                  ),
                ),
              ],
            ),
            Text(
              'Thời lượng: ${fmtHours(minutes / 60)}${overnight && minutes > 0 ? ' · qua nửa đêm' : ''}',
              style: const TextStyle(color: Colors.black54),
            ),
            Align(
              alignment: Alignment.centerLeft,
              child: _Pill('Loại ngày: ${dayLabel(widget.typeOf(_date))}', color: typeColor(widget.typeOf(_date))),
            ),
            TextField(
              controller: _rate,
              keyboardType: const TextInputType.numberWithOptions(decimal: true),
              decoration: InputDecoration(
                labelText: tr('Lương theo giờ'),
                prefixText: r'$ ',
                border: const OutlineInputBorder(),
              ),
            ),
            if (_error != null) Text(_error!, style: TextStyle(color: cs.error)),
            FilledButton(
              onPressed: _save,
              style: FilledButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 14)),
              child: const Text('Lưu'),
            ),
          ],
        ),
      ),
    );
  }
}
