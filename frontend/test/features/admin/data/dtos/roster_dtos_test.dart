import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/data/dtos/doctor_pre_assignment_dto.dart';
import 'package:frontend/features/admin/data/dtos/doctor_unavailability_dto.dart';
import 'package:frontend/features/admin/data/dtos/roster_doctor_dto.dart';
import 'package:frontend/features/admin/data/dtos/roster_position_dto.dart';
import 'package:frontend/features/admin/data/dtos/roster_shift_dto.dart';
import 'package:frontend/features/admin/data/dtos/roster_team_dto.dart';
import 'package:frontend/features/admin/domain/models/doctor_pre_assignment.dart';
import 'package:frontend/features/admin/domain/models/doctor_unavailability.dart';
import 'package:frontend/features/admin/domain/models/roster_doctor.dart';
import 'package:frontend/features/admin/domain/models/roster_position.dart';
import 'package:frontend/features/admin/domain/models/roster_shift.dart';
import 'package:frontend/features/admin/domain/models/roster_team.dart';

void main() {
  group('RosterTeamDto & RosterTeam', () {
    final teamJson = <String, dynamic>{
      'id': 'team-uuid-1',
      'name': 'Team Alpha',
      'department_id': 'dept-uuid-1',
      'created_at': '2026-11-01T08:00:00Z',
      'updated_at': '2026-11-01T08:00:00Z',
      'is_deleted': false,
      'sync_status': false,
    };

    test('parses JSON, converts to RosterTeam, and supports value equality', () {
      final dto = RosterTeamDto.fromJson(teamJson);
      final domain = dto.toDomain();
      final duplicate = dto.toDomain();

      expect(domain, isA<RosterTeam>());
      expect(domain.id, 'team-uuid-1');
      expect(domain.name, 'Team Alpha');
      expect(domain.departmentId, 'dept-uuid-1');
      expect(domain, equals(duplicate));
      expect(domain.hashCode, duplicate.hashCode);
    });

    test('toJson preserves backend snake_case contract', () {
      final dto = RosterTeamDto.fromJson(teamJson);
      expect(dto.toJson(), {
        'id': 'team-uuid-1',
        'name': 'Team Alpha',
        'department_id': 'dept-uuid-1',
      });
    });

    test('throws FormatException when required fields are missing', () {
      for (final key in ['id', 'name', 'department_id']) {
        final invalid = Map<String, dynamic>.from(teamJson)..remove(key);
        expect(() => RosterTeamDto.fromJson(invalid), throwsFormatException);
      }
    });
  });

  group('RosterPositionDto & RosterPosition', () {
    final positionJson = <String, dynamic>{
      'id': 'pos-uuid-1',
      'name': 'Emergency',
      'department_id': 'dept-uuid-1',
      'duty_days': [0, 2, 4],
      'created_at': '2026-11-01T08:00:00Z',
      'updated_at': '2026-11-01T08:00:00Z',
      'is_deleted': false,
      'sync_status': false,
    };

    test('parses JSON including duty_days ints and converts to RosterPosition', () {
      final dto = RosterPositionDto.fromJson(positionJson);
      final domain = dto.toDomain();

      expect(domain, isA<RosterPosition>());
      expect(domain.id, 'pos-uuid-1');
      expect(domain.name, 'Emergency');
      expect(domain.departmentId, 'dept-uuid-1');
      expect(domain.dutyDays, [0, 2, 4]);
    });

    test('toJson outputs snake_case fields', () {
      final dto = RosterPositionDto.fromJson(positionJson);
      expect(dto.toJson(), {
        'id': 'pos-uuid-1',
        'name': 'Emergency',
        'department_id': 'dept-uuid-1',
        'duty_days': [0, 2, 4],
      });
    });

    test('throws FormatException when required fields are missing or invalid', () {
      for (final key in ['id', 'name', 'department_id', 'duty_days']) {
        final invalid = Map<String, dynamic>.from(positionJson)..remove(key);
        expect(
          () => RosterPositionDto.fromJson(invalid),
          throwsFormatException,
        );
      }
    });
  });

  group('RosterShiftDto & RosterShift', () {
    final shiftJson = <String, dynamic>{
      'id': 'shift-uuid-1',
      'name': 'Night',
      'position_id': 'pos-uuid-1',
      'doctors_per_shift': 2,
      'grants_day_off': true,
      'created_at': '2026-11-01T08:00:00Z',
      'updated_at': '2026-11-01T08:00:00Z',
      'is_deleted': false,
      'sync_status': false,
    };

    test('parses JSON and converts to RosterShift domain model', () {
      final dto = RosterShiftDto.fromJson(shiftJson);
      final domain = dto.toDomain();

      expect(domain, isA<RosterShift>());
      expect(domain.id, 'shift-uuid-1');
      expect(domain.name, 'Night');
      expect(domain.positionId, 'pos-uuid-1');
      expect(domain.doctorsPerShift, 2);
      expect(domain.grantsDayOff, isTrue);
    });

    test('toJson serializes snake_case contract', () {
      final dto = RosterShiftDto.fromJson(shiftJson);
      expect(dto.toJson(), {
        'id': 'shift-uuid-1',
        'name': 'Night',
        'position_id': 'pos-uuid-1',
        'doctors_per_shift': 2,
        'grants_day_off': true,
      });
    });

    test('throws FormatException when required fields are missing', () {
      for (final key in [
        'id',
        'name',
        'position_id',
        'doctors_per_shift',
        'grants_day_off',
      ]) {
        final invalid = Map<String, dynamic>.from(shiftJson)..remove(key);
        expect(() => RosterShiftDto.fromJson(invalid), throwsFormatException);
      }
    });
  });

  group('RosterDoctorDto & RosterDoctor', () {
    final doctorJson = <String, dynamic>{
      'id': 'doc-uuid-1',
      'name': 'Dr. Panos',
      'department_id': 'dept-uuid-1',
      'team_id': 'team-uuid-1',
      'position_ids': ['pos-uuid-1', 'pos-uuid-2'],
      'created_at': '2026-11-01T08:00:00Z',
      'updated_at': '2026-11-01T08:00:00Z',
      'is_deleted': false,
      'sync_status': false,
    };

    test('parses nullable team_id and position_ids list', () {
      final assigned = RosterDoctorDto.fromJson(doctorJson).toDomain();
      expect(assigned, isA<RosterDoctor>());
      expect(assigned.id, 'doc-uuid-1');
      expect(assigned.name, 'Dr. Panos');
      expect(assigned.departmentId, 'dept-uuid-1');
      expect(assigned.teamId, 'team-uuid-1');
      expect(assigned.positionIds, ['pos-uuid-1', 'pos-uuid-2']);

      final teamlessJson = Map<String, dynamic>.from(doctorJson)
        ..['team_id'] = null
        ..remove('position_ids');
      final teamless = RosterDoctorDto.fromJson(teamlessJson).toDomain();
      expect(teamless.teamId, isNull);
      expect(teamless.positionIds, isEmpty);
    });

    test('toJson serializes snake_case fields including position_ids', () {
      final dto = RosterDoctorDto.fromJson(doctorJson);
      expect(dto.toJson(), {
        'id': 'doc-uuid-1',
        'name': 'Dr. Panos',
        'department_id': 'dept-uuid-1',
        'team_id': 'team-uuid-1',
        'position_ids': ['pos-uuid-1', 'pos-uuid-2'],
      });
    });

    test('throws FormatException when required fields are missing', () {
      for (final key in ['id', 'name', 'department_id']) {
        final invalid = Map<String, dynamic>.from(doctorJson)..remove(key);
        expect(() => RosterDoctorDto.fromJson(invalid), throwsFormatException);
      }
    });
  });

  group('DoctorUnavailabilityDto & DoctorUnavailability', () {
    final unavailJson = <String, dynamic>{
      'id': 'unavail-uuid-1',
      'doctor_id': 'doc-uuid-1',
      'date': '2026-11-15',
      'created_at': '2026-11-01T08:00:00Z',
      'updated_at': '2026-11-01T08:00:00Z',
      'is_deleted': false,
      'sync_status': false,
    };

    test('parses JSON and converts to DoctorUnavailability', () {
      final dto = DoctorUnavailabilityDto.fromJson(unavailJson);
      final domain = dto.toDomain();

      expect(domain, isA<DoctorUnavailability>());
      expect(domain.id, 'unavail-uuid-1');
      expect(domain.doctorId, 'doc-uuid-1');
      expect(domain.date, '2026-11-15');
    });

    test('toJson outputs backend fields', () {
      final dto = DoctorUnavailabilityDto.fromJson(unavailJson);
      expect(dto.toJson(), {
        'id': 'unavail-uuid-1',
        'doctor_id': 'doc-uuid-1',
        'date': '2026-11-15',
      });
    });

    test('throws FormatException when required fields are missing', () {
      for (final key in ['id', 'doctor_id', 'date']) {
        final invalid = Map<String, dynamic>.from(unavailJson)..remove(key);
        expect(
          () => DoctorUnavailabilityDto.fromJson(invalid),
          throwsFormatException,
        );
      }
    });
  });

  group('DoctorPreAssignmentDto & DoctorPreAssignment', () {
    final preAssignJson = <String, dynamic>{
      'id': 'pa-uuid-1',
      'doctor_id': 'doc-uuid-1',
      'shift_id': 'shift-uuid-1',
      'date': '2026-11-20',
      'created_at': '2026-11-01T08:00:00Z',
      'updated_at': '2026-11-01T08:00:00Z',
      'is_deleted': false,
      'sync_status': false,
    };

    test('parses JSON and converts to DoctorPreAssignment', () {
      final dto = DoctorPreAssignmentDto.fromJson(preAssignJson);
      final domain = dto.toDomain();

      expect(domain, isA<DoctorPreAssignment>());
      expect(domain.id, 'pa-uuid-1');
      expect(domain.doctorId, 'doc-uuid-1');
      expect(domain.shiftId, 'shift-uuid-1');
      expect(domain.date, '2026-11-20');
    });

    test('toJson outputs backend fields', () {
      final dto = DoctorPreAssignmentDto.fromJson(preAssignJson);
      expect(dto.toJson(), {
        'id': 'pa-uuid-1',
        'doctor_id': 'doc-uuid-1',
        'shift_id': 'shift-uuid-1',
        'date': '2026-11-20',
      });
    });

    test('throws FormatException when required fields are missing', () {
      for (final key in ['id', 'doctor_id', 'shift_id', 'date']) {
        final invalid = Map<String, dynamic>.from(preAssignJson)..remove(key);
        expect(
          () => DoctorPreAssignmentDto.fromJson(invalid),
          throwsFormatException,
        );
      }
    });
  });
}
