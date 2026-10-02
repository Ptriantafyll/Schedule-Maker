import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/data/dtos/schedule_draft_dto.dart';
import 'package:frontend/features/admin/domain/models/schedule_draft.dart';

void main() {
  group('ScheduleAssignmentDto Tests', () {
    final validAssignmentJson = <String, dynamic>{
      'date': '2026-11-01',
      'day_name': 'Sunday',
      'doctor_name': 'Dr. Gregory House',
      'doctor_email': 'house@hospital.org',
      'position': 'ER',
      'shift': 'Night',
    };

    test('fromJson successfully parses snake_case JSON and toDomain produces ScheduleAssignment', () {
      final dto = ScheduleAssignmentDto.fromJson(validAssignmentJson);
      final domain = dto.toDomain();

      expect(domain, isA<ScheduleAssignment>());
      expect(domain.date, '2026-11-01');
      expect(domain.dayName, 'Sunday');
      expect(domain.doctorName, 'Dr. Gregory House');
      expect(domain.doctorEmail, 'house@hospital.org');
      expect(domain.position, 'ER');
      expect(domain.shift, 'Night');
    });

    test('toJson serializes back into backend snake_case contract', () {
      final dto = ScheduleAssignmentDto.fromJson(validAssignmentJson);
      final json = dto.toJson();

      expect(json, validAssignmentJson);
    });

    test('fromJson throws FormatException when required fields are missing', () {
      final missingDate = Map<String, dynamic>.from(validAssignmentJson)..remove('date');
      expect(() => ScheduleAssignmentDto.fromJson(missingDate), throwsFormatException);

      final missingDoctor = Map<String, dynamic>.from(validAssignmentJson)..remove('doctor_name');
      expect(() => ScheduleAssignmentDto.fromJson(missingDoctor), throwsFormatException);
    });
  });

  group('ScheduleDraftDto Tests', () {
    final validDraftJson = <String, dynamic>{
      'id': '550e8400-e29b-41d4-a716-446655440000',
      'department_id': '6ba7b810-9dad-11d1-80b4-00c04fd430c8',
      'target_month': '2026-11',
      'source_filename': 'november_roster.xlsx',
      'total_duties': 28,
      'solver_status': 'OPTIMAL',
      'status': 'draft',
      'assignments': [
        {
          'date': '2026-11-01',
          'day_name': 'Sunday',
          'doctor_name': 'Dr. Gregory House',
          'doctor_email': 'house@hospital.org',
          'position': 'ER',
          'shift': 'Night',
        },
      ],
      'unavailabilities': {
        'Dr. Gregory House': [3, 14, 22],
      },
      'created_at': '2026-11-01T08:30:00.000Z',
      'updated_at': '2026-11-01T08:35:00.000Z',
    };

    test('fromJson and toDomain produce ScheduleDraft domain entity with nested models', () {
      final dto = ScheduleDraftDto.fromJson(validDraftJson);
      final domain = dto.toDomain();

      expect(domain, isA<ScheduleDraft>());
      expect(domain.id, '550e8400-e29b-41d4-a716-446655440000');
      expect(domain.departmentId, '6ba7b810-9dad-11d1-80b4-00c04fd430c8');
      expect(domain.targetMonth, '2026-11');
      expect(domain.sourceFilename, 'november_roster.xlsx');
      expect(domain.totalDuties, 28);
      expect(domain.solverStatus, 'OPTIMAL');
      expect(domain.status, 'draft');
      expect(domain.assignments.length, 1);
      expect(domain.assignments.first.doctorName, 'Dr. Gregory House');
      expect(domain.unavailabilities['Dr. Gregory House'], [3, 14, 22]);
      expect(domain.createdAt, DateTime.parse('2026-11-01T08:30:00.000Z'));
      expect(domain.updatedAt, DateTime.parse('2026-11-01T08:35:00.000Z'));
    });

    test('fromJson handles empty or omitted assignments and unavailabilities', () {
      final minimalJson = <String, dynamic>{
        'id': 'draft-1',
        'department_id': 'dept-1',
        'target_month': '2026-11',
        'source_filename': 'schedule.xlsx',
        'created_at': '2026-11-01T08:30:00.000Z',
        'updated_at': '2026-11-01T08:30:00.000Z',
      };

      final dto = ScheduleDraftDto.fromJson(minimalJson);
      final domain = dto.toDomain();

      expect(domain.assignments, isEmpty);
      expect(domain.unavailabilities, isEmpty);
      expect(domain.totalDuties, 0);
      expect(domain.solverStatus, 'OPTIMAL');
      expect(domain.status, 'draft');
    });

    test('toJson serializes nested structures back to snake_case contract', () {
      final dto = ScheduleDraftDto.fromJson(validDraftJson);
      final json = dto.toJson();

      expect(json['id'], '550e8400-e29b-41d4-a716-446655440000');
      expect(json['department_id'], '6ba7b810-9dad-11d1-80b4-00c04fd430c8');
      expect(json['target_month'], '2026-11');
      expect(json['source_filename'], 'november_roster.xlsx');
      expect(json['total_duties'], 28);
      expect(json['solver_status'], 'OPTIMAL');
      expect(json['status'], 'draft');
      expect(json['assignments'], isA<List>());
      expect((json['assignments'] as List).length, 1);
      expect(json['unavailabilities'], {
        'Dr. Gregory House': [3, 14, 22],
      });
      expect(json['created_at'], '2026-11-01T08:30:00.000Z');
      expect(json['updated_at'], '2026-11-01T08:35:00.000Z');
    });

    test('fromJson throws FormatException when required id is missing', () {
      final missingId = Map<String, dynamic>.from(validDraftJson)..remove('id');
      expect(() => ScheduleDraftDto.fromJson(missingId), throwsFormatException);
    });
  });
}
