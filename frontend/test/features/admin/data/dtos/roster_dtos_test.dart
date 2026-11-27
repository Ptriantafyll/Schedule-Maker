import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/data/dtos/department_doctor_user_dto.dart';
import 'package:frontend/features/admin/data/dtos/doctor_pre_assignment_dto.dart';
import 'package:frontend/features/admin/data/dtos/doctor_unavailability_dto.dart';
import 'package:frontend/features/admin/data/dtos/roster_doctor_dto.dart';
import 'package:frontend/features/admin/data/dtos/roster_position_dto.dart';
import 'package:frontend/features/admin/data/dtos/roster_shift_dto.dart';
import 'package:frontend/features/admin/data/dtos/roster_team_dto.dart';
import 'package:frontend/features/admin/data/dtos/staff_invitation_dto.dart';
import 'package:frontend/features/admin/domain/models/department_doctor_user.dart';
import 'package:frontend/features/admin/domain/models/doctor_pre_assignment.dart';
import 'package:frontend/features/admin/domain/models/doctor_unavailability.dart';
import 'package:frontend/features/admin/domain/models/roster_doctor.dart';
import 'package:frontend/features/admin/domain/models/roster_position.dart';
import 'package:frontend/features/admin/domain/models/roster_shift.dart';
import 'package:frontend/features/admin/domain/models/roster_team.dart';
import 'package:frontend/features/admin/domain/models/staff_invitation.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';

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

    test('parses linked_user_id and linked_user_email when present', () {
      final linkedJson = Map<String, dynamic>.from(doctorJson)
        ..['linked_user_id'] = 'usr-uuid-1'
        ..['linked_user_email'] = 'panos@hospital.org';

      final dto = RosterDoctorDto.fromJson(linkedJson);
      final domain = dto.toDomain();

      expect(domain.linkedUserId, 'usr-uuid-1');
      expect(domain.linkedUserEmail, 'panos@hospital.org');
      expect(dto.toJson(), {
        'id': 'doc-uuid-1',
        'name': 'Dr. Panos',
        'department_id': 'dept-uuid-1',
        'team_id': 'team-uuid-1',
        'position_ids': ['pos-uuid-1', 'pos-uuid-2'],
        'linked_user_id': 'usr-uuid-1',
        'linked_user_email': 'panos@hospital.org',
      });

      final unlinked = RosterDoctorDto.fromJson(doctorJson).toDomain();
      expect(unlinked.linkedUserId, isNull);
      expect(unlinked.linkedUserEmail, isNull);
      expect(unlinked, isNot(equals(domain)));
    });

    test('throws FormatException when linked user fields have invalid types', () {
      for (final key in ['linked_user_id', 'linked_user_email']) {
        final invalid = Map<String, dynamic>.from(doctorJson)..[key] = 123;
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

  group('StaffInvitationDto & StaffInvitation', () {
    final futureExpiry = DateTime.now().toUtc().add(const Duration(days: 7));
    final pastExpiry = DateTime.now().toUtc().subtract(const Duration(days: 1));
    final invitationJson = <String, dynamic>{
      'id': 'inv-uuid-1',
      'role': 'doctor',
      'department_id': 'dept-uuid-1',
      'doctor_id': null,
      'created_by_user_id': 'admin-uuid-1',
      'expires_at': futureExpiry.toIso8601String(),
      'used_at': null,
      'revoked_at': null,
      'created_at': '2026-11-01T08:00:00Z',
      'updated_at': '2026-11-01T08:00:00Z',
      'is_deleted': false,
      'sync_status': false,
      'max_uses': 5,
      'used_count': 2,
    };

    test('parses JSON, converts to StaffInvitation, and computes active status', () {
      final domain = StaffInvitationDto.fromJson(invitationJson).toDomain();
      final duplicate = StaffInvitationDto.fromJson(invitationJson).toDomain();

      expect(domain, isA<StaffInvitation>());
      expect(domain.id, 'inv-uuid-1');
      expect(domain.role, UserRole.doctor);
      expect(domain.departmentId, 'dept-uuid-1');
      expect(domain.doctorId, isNull);
      expect(domain.expiresAt, futureExpiry);
      expect(domain.usedAt, isNull);
      expect(domain.revokedAt, isNull);
      expect(domain.maxUses, 5);
      expect(domain.usedCount, 2);
      expect(domain.remainingUses, 3);
      expect(domain.isActive, isTrue);
      expect(domain, equals(duplicate));
      expect(domain.hashCode, duplicate.hashCode);
    });

    test('isActive is false when revoked, exhausted, or expired', () {
      final revoked = StaffInvitationDto.fromJson({
        ...invitationJson,
        'revoked_at': '2026-11-02T08:00:00Z',
      }).toDomain();
      expect(revoked.isActive, isFalse);

      final exhausted = StaffInvitationDto.fromJson({
        ...invitationJson,
        'used_count': 5,
        'used_at': '2026-11-02T08:00:00Z',
      }).toDomain();
      expect(exhausted.remainingUses, 0);
      expect(exhausted.isActive, isFalse);

      final expired = StaffInvitationDto.fromJson({
        ...invitationJson,
        'expires_at': pastExpiry.toIso8601String(),
      }).toDomain();
      expect(expired.isActive, isFalse);
    });

    test('throws FormatException when required or date fields are invalid', () {
      for (final key in ['id', 'role', 'department_id', 'expires_at']) {
        final invalid = Map<String, dynamic>.from(invitationJson)..remove(key);
        expect(
          () => StaffInvitationDto.fromJson(invalid),
          throwsFormatException,
        );
      }

      expect(
        () => StaffInvitationDto.fromJson({
          ...invitationJson,
          'expires_at': 'not-a-timestamp',
        }),
        throwsFormatException,
      );
      expect(
        () => StaffInvitationDto.fromJson({...invitationJson, 'role': 'ghost'}),
        throwsFormatException,
      );
    });
  });

  group('CreatedStaffInvitationDto & CreatedStaffInvitation', () {
    final createdJson = <String, dynamic>{
      'id': 'inv-uuid-2',
      'role': 'doctor',
      'department_id': 'dept-uuid-1',
      'doctor_id': 'doc-uuid-1',
      'expires_at': '2026-12-01T12:00:00Z',
      'raw_token': 'raw-secret-token-123',
      'max_uses': 1,
      'used_count': 0,
    };

    test('parses JSON and converts to CreatedStaffInvitation', () {
      final domain = CreatedStaffInvitationDto.fromJson(createdJson).toDomain();
      final duplicate = CreatedStaffInvitationDto.fromJson(
        createdJson,
      ).toDomain();

      expect(domain, isA<CreatedStaffInvitation>());
      expect(domain.id, 'inv-uuid-2');
      expect(domain.role, UserRole.doctor);
      expect(domain.departmentId, 'dept-uuid-1');
      expect(domain.doctorId, 'doc-uuid-1');
      expect(domain.expiresAt, DateTime.parse('2026-12-01T12:00:00Z'));
      expect(domain.rawToken, 'raw-secret-token-123');
      expect(domain.maxUses, 1);
      expect(domain.usedCount, 0);
      expect(domain, equals(duplicate));
      expect(domain.hashCode, duplicate.hashCode);
    });

    test('throws FormatException when required fields are missing or invalid', () {
      for (final key in [
        'id',
        'role',
        'department_id',
        'expires_at',
        'raw_token',
      ]) {
        final invalid = Map<String, dynamic>.from(createdJson)..remove(key);
        expect(
          () => CreatedStaffInvitationDto.fromJson(invalid),
          throwsFormatException,
        );
      }
    });
  });

  group('DepartmentDoctorUserDto & DepartmentDoctorUser', () {
    final userJson = <String, dynamic>{
      'id': 'usr-uuid-1',
      'full_name': 'Dr. Gregory House',
      'email': 'house@hospital.org',
      'doctor_id': 'doc-uuid-1',
    };

    test('parses JSON and converts to DepartmentDoctorUser', () {
      final domain = DepartmentDoctorUserDto.fromJson(userJson).toDomain();
      final duplicate = DepartmentDoctorUserDto.fromJson(userJson).toDomain();

      expect(domain, isA<DepartmentDoctorUser>());
      expect(domain.id, 'usr-uuid-1');
      expect(domain.fullName, 'Dr. Gregory House');
      expect(domain.email, 'house@hospital.org');
      expect(domain.doctorId, 'doc-uuid-1');
      expect(domain, equals(duplicate));
      expect(domain.hashCode, duplicate.hashCode);
    });

    test('throws FormatException when required fields are missing', () {
      for (final key in ['id', 'full_name', 'email', 'doctor_id']) {
        final invalid = Map<String, dynamic>.from(userJson)..remove(key);
        expect(
          () => DepartmentDoctorUserDto.fromJson(invalid),
          throwsFormatException,
        );
      }
    });
  });
}
