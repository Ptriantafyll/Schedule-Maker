import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/auth/data/dtos/user_dto.dart';
import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';

void main() {
  group('UserDto', () {
    const validDoctorJson = {
      'id': 'usr-123',
      'email': 'doctor@hospital.org',
      'full_name': 'Dr. Gregory House',
      'role': 'doctor',
      'department_id': 'dept-456',
      'doctor_id': 'doc-789',
      'is_deleted': false,
      'sync_status': false,
      'created_at': '2026-01-01T00:00:00Z',
      'updated_at': '2026-01-01T00:00:00Z',
    };

    const validSuperAdminJson = {
      'id': 'usr-999',
      'email': 'admin@hospital.org',
      'full_name': 'Admin Alice',
      'role': 'super_admin',
      'department_id': null,
      'doctor_id': null,
      'is_deleted': false,
      'sync_status': false,
      'created_at': '2026-01-01T00:00:00Z',
      'updated_at': '2026-01-01T00:00:00Z',
    };

    group('fromJson', () {
      test('parses full doctor payload correctly', () {
        final dto = UserDto.fromJson(validDoctorJson);

        expect(dto.id, equals('usr-123'));
        expect(dto.email, equals('doctor@hospital.org'));
        expect(dto.fullName, equals('Dr. Gregory House'));
        expect(dto.role, equals(UserRole.doctor));
        expect(dto.departmentId, equals('dept-456'));
        expect(dto.doctorId, equals('doc-789'));
      });

      test('parses super admin payload with null department and doctor IDs', () {
        final dto = UserDto.fromJson(validSuperAdminJson);

        expect(dto.id, equals('usr-999'));
        expect(dto.email, equals('admin@hospital.org'));
        expect(dto.fullName, equals('Admin Alice'));
        expect(dto.role, equals(UserRole.superAdmin));
        expect(dto.departmentId, isNull);
        expect(dto.doctorId, isNull);
      });

      test('throws FormatException when role is unknown', () {
        final invalidRoleJson = Map<String, dynamic>.from(validDoctorJson)
          ..['role'] = 'unknown_alien_role';

        expect(
          () => UserDto.fromJson(invalidRoleJson),
          throwsA(isA<FormatException>()),
        );
      });

      test('throws FormatException when required id field is missing', () {
        final missingIdJson = Map<String, dynamic>.from(validDoctorJson)
          ..remove('id');

        expect(
          () => UserDto.fromJson(missingIdJson),
          throwsA(isA<FormatException>()),
        );
      });
    });

    group('toDomain', () {
      test('maps UserDto to a pure domain User entity', () {
        final dto = UserDto.fromJson(validDoctorJson);
        final user = dto.toDomain();

        expect(user, isA<User>());
        expect(
          user,
          equals(
            const User(
              id: 'usr-123',
              email: 'doctor@hospital.org',
              fullName: 'Dr. Gregory House',
              role: UserRole.doctor,
              departmentId: 'dept-456',
              doctorId: 'doc-789',
            ),
          ),
        );
      });
    });

    group('toJson', () {
      test('serializes back to backend snake_case format', () {
        final dto = UserDto.fromJson(validDoctorJson);
        final json = dto.toJson();

        expect(json['id'], equals('usr-123'));
        expect(json['email'], equals('doctor@hospital.org'));
        expect(json['full_name'], equals('Dr. Gregory House'));
        expect(json['role'], equals('doctor'));
        expect(json['department_id'], equals('dept-456'));
        expect(json['doctor_id'], equals('doc-789'));
      });
    });
  });
}
