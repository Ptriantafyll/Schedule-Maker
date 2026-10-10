import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/data/dtos/target_month_info_dto.dart';
import 'package:frontend/features/admin/domain/models/target_month_info.dart';

void main() {
  group('TargetMonthInfoDto Tests', () {
    final jsonWithHistory = <String, dynamic>{
      'next_target_month': '2026-12',
      'last_published_month': '2026-11',
    };

    final jsonFreshDepartment = <String, dynamic>{
      'next_target_month': '2026-11',
      'last_published_month': null,
    };

    test('fromJson parses department with published schedule history', () {
      final dto = TargetMonthInfoDto.fromJson(jsonWithHistory);
      final domain = dto.toDomain();

      expect(domain, isA<TargetMonthInfo>());
      expect(domain.nextTargetMonth, '2026-12');
      expect(domain.lastPublishedMonth, '2026-11');
      expect(domain.hasPublishedSchedules, isTrue);
    });

    test('fromJson parses fresh department with null last_published_month', () {
      final dto = TargetMonthInfoDto.fromJson(jsonFreshDepartment);
      final domain = dto.toDomain();

      expect(domain.nextTargetMonth, '2026-11');
      expect(domain.lastPublishedMonth, isNull);
      expect(domain.hasPublishedSchedules, isFalse);
    });

    test('toJson serializes back into backend contract', () {
      expect(TargetMonthInfoDto.fromJson(jsonWithHistory).toJson(), jsonWithHistory);
      expect(TargetMonthInfoDto.fromJson(jsonFreshDepartment).toJson(), jsonFreshDepartment);
    });

    test('fromJson throws FormatException when next_target_month is missing', () {
      expect(() => TargetMonthInfoDto.fromJson(const {}), throwsFormatException);
      expect(() => TargetMonthInfoDto.fromJson({'next_target_month': 123}), throwsFormatException);
    });
  });
}
