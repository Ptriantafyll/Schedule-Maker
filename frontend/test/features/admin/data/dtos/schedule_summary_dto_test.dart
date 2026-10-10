import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/data/dtos/schedule_summary_dto.dart';
import 'package:frontend/features/admin/domain/models/schedule_summary.dart';

void main() {
  group('ScheduleSummaryDto Tests', () {
    final validJson = <String, dynamic>{
      'id': '550e8400-e29b-41d4-a716-446655440000',
      'department_id': '6ba7b810-9dad-11d1-80b4-00c04fd430c8',
      'target_month': '2026-11',
      'source_filename': 'november_roster.xlsx',
      'total_duties': 45,
      'solver_status': 'OPTIMAL',
      'status': 'published',
      'created_at': '2026-11-01T08:00:00.000Z',
      'updated_at': '2026-11-01T08:30:00.000Z',
    };

    test('fromJson parses snake_case JSON and toDomain produces ScheduleSummary', () {
      final dto = ScheduleSummaryDto.fromJson(validJson);
      final domain = dto.toDomain();

      expect(domain, isA<ScheduleSummary>());
      expect(domain.id, '550e8400-e29b-41d4-a716-446655440000');
      expect(domain.departmentId, '6ba7b810-9dad-11d1-80b4-00c04fd430c8');
      expect(domain.targetMonth, '2026-11');
      expect(domain.sourceFilename, 'november_roster.xlsx');
      expect(domain.totalDuties, 45);
      expect(domain.solverStatus, 'OPTIMAL');
      expect(domain.status, 'published');
      expect(domain.isPublished, isTrue);
      expect(domain.createdAt, DateTime.parse('2026-11-01T08:00:00.000Z'));
      expect(domain.updatedAt, DateTime.parse('2026-11-01T08:30:00.000Z'));
    });

    test('toJson serializes back into backend contract', () {
      final dto = ScheduleSummaryDto.fromJson(validJson);
      expect(dto.toJson(), validJson);
    });

    test('fromJson handles default total_duties and solver_status when omitted', () {
      final minimalJson = <String, dynamic>{
        'id': 'draft-1',
        'department_id': 'dept-1',
        'target_month': '2026-11',
        'source_filename': 'schedule.xlsx',
        'status': 'draft',
        'created_at': '2026-11-01T08:00:00.000Z',
        'updated_at': '2026-11-01T08:00:00.000Z',
      };

      final dto = ScheduleSummaryDto.fromJson(minimalJson);
      expect(dto.totalDuties, 0);
      expect(dto.solverStatus, 'OPTIMAL');
    });

    test('fromJson throws FormatException when required fields are missing', () {
      final missingId = Map<String, dynamic>.from(validJson)..remove('id');
      expect(() => ScheduleSummaryDto.fromJson(missingId), throwsFormatException);

      final missingDept = Map<String, dynamic>.from(validJson)..remove('department_id');
      expect(() => ScheduleSummaryDto.fromJson(missingDept), throwsFormatException);

      final missingMonth = Map<String, dynamic>.from(validJson)..remove('target_month');
      expect(() => ScheduleSummaryDto.fromJson(missingMonth), throwsFormatException);
    });

    test('fromJson throws FormatException for invalid date string', () {
      final invalidDate = Map<String, dynamic>.from(validJson)..['created_at'] = 'not-a-date';
      expect(() => ScheduleSummaryDto.fromJson(invalidDate), throwsFormatException);
    });
  });
}
