import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';

void main() {
  group('UserRole', () {
    group('fromString', () {
      // 1. Parameterized Test using Dart 3 Records: (input, expected)
      const validCases = [
        ('super_admin', UserRole.superAdmin),
        ('department_admin', UserRole.departmentAdmin),
        ('doctor', UserRole.doctor),
        ('viewer', UserRole.viewer),
      ];

      for (final (input, expected) in validCases) {
        test('parses "$input" to $expected', () {
          expect(UserRole.fromString(input), equals(expected));
        });
      }

      // 2. Exception Testing using Closures
      test('throws FormatException for unknown role string', () {
        expect(
          () => UserRole.fromString('invalid_role'),
          throwsA(isA<FormatException>()),
        );
      });

      test('throws FormatException for empty string', () {
        expect(
          () => UserRole.fromString(''),
          throwsA(isA<FormatException>()),
        );
      });
    });

    // 3. Serialization / raw backend value
    group('value', () {
      test('returns correct backend string representation', () {
        expect(UserRole.superAdmin.value, equals('super_admin'));
        expect(UserRole.departmentAdmin.value, equals('department_admin'));
        expect(UserRole.doctor.value, equals('doctor'));
        expect(UserRole.viewer.value, equals('viewer'));
      });
    });

    // 4. Role Hierarchy & Permission Helpers
    group('helper getters', () {
      test('isAdmin returns true only for superAdmin and departmentAdmin', () {
        expect(UserRole.superAdmin.isAdmin, isTrue);
        expect(UserRole.departmentAdmin.isAdmin, isTrue);
        expect(UserRole.doctor.isAdmin, isFalse);
        expect(UserRole.viewer.isAdmin, isFalse);
      });

      test('requiresDepartment returns false for superAdmin and true for others', () {
        expect(UserRole.superAdmin.requiresDepartment, isFalse);
        expect(UserRole.departmentAdmin.requiresDepartment, isTrue);
        expect(UserRole.doctor.requiresDepartment, isTrue);
        expect(UserRole.viewer.requiresDepartment, isTrue);
      });

      test('requiresDoctor returns true only for doctor', () {
        expect(UserRole.superAdmin.requiresDoctor, isFalse);
        expect(UserRole.departmentAdmin.requiresDoctor, isFalse);
        expect(UserRole.doctor.requiresDoctor, isTrue);
        expect(UserRole.viewer.requiresDoctor, isFalse);
      });

      test('displayName returns human-readable label', () {
        expect(UserRole.superAdmin.displayName, equals('Super Admin'));
        expect(UserRole.departmentAdmin.displayName, equals('Department Admin'));
        expect(UserRole.doctor.displayName, equals('Doctor'));
        expect(UserRole.viewer.displayName, equals('Viewer'));
      });
    });
  });
}