import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';
import 'package:frontend/features/auth/domain/state/auth_state.dart';

void main() {
  const testUser = User(
    id: 'usr-1',
    email: 'doctor@hospital.org',
    fullName: 'Dr. Gregory House',
    role: UserRole.doctor,
    departmentId: 'dept-101',
    doctorId: 'doc-55',
  );

  const differentUser = User(
    id: 'usr-2',
    email: 'admin@hospital.org',
    fullName: 'Admin Alice',
    role: UserRole.superAdmin,
  );

  group('AuthState', () {
    group('AuthStateUnauthenticated', () {
      test('is an instance of AuthState', () {
        const state = AuthStateUnauthenticated();

        expect(state, isA<AuthState>());
        expect(state.isAuthenticated, isFalse);
        expect(state.userOrNull, isNull);
      });

      test('equality holds for separate unauthenticated instances', () {
        const state1 = AuthStateUnauthenticated();
        const state2 = AuthStateUnauthenticated();

        expect(state1, equals(state2));
        expect(state1.hashCode, equals(state2.hashCode));
      });

      test('toString formats cleanly', () {
        const state = AuthStateUnauthenticated();

        expect(state.toString(), equals('AuthStateUnauthenticated()'));
      });
    });

    group('AuthStateAuthenticated', () {
      test('is an instance of AuthState and holds user data', () {
        const state = AuthStateAuthenticated(testUser);

        expect(state, isA<AuthState>());
        expect(state.isAuthenticated, isTrue);
        expect(state.user, equals(testUser));
        expect(state.userOrNull, equals(testUser));
      });

      test('equality holds for instances with identical users', () {
        const state1 = AuthStateAuthenticated(testUser);
        const state2 = AuthStateAuthenticated(testUser);

        expect(state1, equals(state2));
        expect(state1.hashCode, equals(state2.hashCode));
      });

      test('instances with different users are not equal', () {
        const state1 = AuthStateAuthenticated(testUser);
        const state2 = AuthStateAuthenticated(differentUser);

        expect(state1, isNot(equals(state2)));
      });

      test('unauthenticated and authenticated states are not equal', () {
        const unauth = AuthStateUnauthenticated();
        const auth = AuthStateAuthenticated(testUser);

        expect(unauth, isNot(equals(auth)));
      });

      test('toString includes user details', () {
        const state = AuthStateAuthenticated(testUser);

        expect(state.toString(), contains('AuthStateAuthenticated(user:'));
      });
    });

    group('pattern matching (Dart 3 sealed class)', () {
      test('exhaustive switch cleanly distinguishes states', () {
        const AuthState unauthState = AuthStateUnauthenticated();
        const AuthState authState = AuthStateAuthenticated(testUser);

        String describeState(AuthState state) => switch (state) {
              AuthStateUnauthenticated() => 'guest',
              AuthStateAuthenticated(:final user) => 'logged_in: ${user.fullName}',
            };

        expect(describeState(unauthState), equals('guest'));
        expect(describeState(authState), equals('logged_in: Dr. Gregory House'));
      });
    });
  });
}
