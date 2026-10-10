import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/domain/models/target_month_info.dart';

void main() {
  group('TargetMonthInfo Domain Model Tests', () {
    const infoWithHistory1 = TargetMonthInfo(
      nextTargetMonth: '2026-12',
      lastPublishedMonth: '2026-11',
    );

    const infoWithHistory2 = TargetMonthInfo(
      nextTargetMonth: '2026-12',
      lastPublishedMonth: '2026-11',
    );

    const freshDeptInfo = TargetMonthInfo(
      nextTargetMonth: '2026-11',
      lastPublishedMonth: null,
    );

    test('supports value equality and identical hashCodes', () {
      expect(infoWithHistory1, equals(infoWithHistory2));
      expect(infoWithHistory1.hashCode, equals(infoWithHistory2.hashCode));
      expect(infoWithHistory1, isNot(equals(freshDeptInfo)));
    });

    test('hasPublishedSchedules getter indicates if department has prior history', () {
      expect(infoWithHistory1.hasPublishedSchedules, isTrue);
      expect(freshDeptInfo.hasPublishedSchedules, isFalse);
    });

    test('toString includes nextTargetMonth and lastPublishedMonth', () {
      final str = infoWithHistory1.toString();
      expect(str, contains('2026-12'));
      expect(str, contains('2026-11'));
    });
  });
}
