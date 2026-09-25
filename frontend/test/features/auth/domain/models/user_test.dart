import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';

void main() {
  group('User', () {
    group('Instantiation', () {
      test('Creates a User instance with all fields populated', () {
        const user = User(
          id: 'usr-1',
          email: 'doctor@hospital.org',
          fullName: 'Dr. Gregory House',
          role: UserRole.doctor,
          departmentId: 'dept-101',
          doctorId: 'doc-55',
        );

        expect(user.id, equals('usr-1'));
        expect(user.email, equals('doctor@hospital.org'));
        expect(user.fullName, equals('Dr. Gregory House'));
        expect(user.role, equals(UserRole.doctor));
        expect(user.departmentId, equals('dept-101'));
        expect(user.doctorId, equals('doc-55'));
      });
      test('creates a User instance with null departmentId and doctorId', () {
        const user = User(
          id: 'usr-2',
          email: 'superadmin@hospital.org',
          fullName: 'Admin Alice',
          role: UserRole.superAdmin,
        );

        expect(user.departmentId, isNull);
        expect(user.doctorId, isNull);
      });
    });

    group('equality', () {
      const user1 = User(
        id: 'usr-1',
        email: 'doctor@hospital.org',
        fullName: 'Dr. Gregory House',
        role: UserRole.doctor,
        departmentId: 'dept-101',
        doctorId: 'doc-55',
      );

      const user2 = User(
        id: 'usr-1',
        email: 'doctor@hospital.org',
        fullName: 'Dr. Gregory House',
        role: UserRole.doctor,
        departmentId: 'dept-101',
        doctorId: 'doc-55',
      );

      const differentUser = User(
        id: 'usr-999',
        email: 'other@hospital.org',
        fullName: 'Dr. James Wilson',
        role: UserRole.doctor,
        departmentId: 'dept-101',
        doctorId: 'doc-56',
      );
      test('two instances with identical fields are equal', () {
        expect(user1, equals(user2));
      });

      test('two instances with identical fields have the same hashCode', () {
        expect(user1.hashCode, equals(user2.hashCode));
      });

      test('instances with different fields are not equal', () {
        expect(user1, isNot(equals(differentUser)));
      });
    });

    group('copyWith', () {
      const originalUser = User(
        id: 'usr-1',
        email: 'doctor@hospital.org',
        fullName: 'Dr. Gregory House',
        role: UserRole.doctor,
        departmentId: 'dept-101',
        doctorId: 'doc-55',
      );

      test('creates a copy with updated fields while preserving others', () {
        final updated = originalUser.copyWith(
          fullName: 'Dr. Greg House, M.D.',
          email: 'g.house@hospital.org',
        );

        expect(updated.id, equals('usr-1')); // preserved
        expect(updated.fullName, equals('Dr. Greg House, M.D.')); // updated
        expect(updated.email, equals('g.house@hospital.org')); // updated
        expect(updated.role, equals(UserRole.doctor)); // preserved
        expect(updated.departmentId, equals('dept-101')); // preserved
        expect(updated.doctorId, equals('doc-55')); // preserved
      });

      test('creates an identical copy when no arguments are passed', () {
        final copy = originalUser.copyWith();

        expect(copy, equals(originalUser));
      });
    });
  });
}
