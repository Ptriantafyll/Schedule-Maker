import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/auth/data/repositories/auth_repository.dart';
import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';
import 'package:frontend/features/auth/domain/state/auth_state.dart';
import 'package:frontend/features/auth/presentation/controllers/auth_controller.dart';
import 'package:frontend/shared/widgets/profile_drawer.dart';

class FakeAuthRepository implements AuthRepository {
  int logoutCallCount = 0;

  @override
  Future<User> login({required String email, required String password}) {
    throw UnimplementedError();
  }

  @override
  Future<User> signup({
    required String invitationToken,
    required String firstName,
    required String lastName,
    required String email,
    required String password,
  }) {
    throw UnimplementedError();
  }

  @override
  Future<void> logout() async {
    logoutCallCount++;
  }

  @override
  Future<User?> restoreSession() async => null;
}

void main() {
  late FakeAuthRepository fakeRepository;

  const testUser = User(
    id: 'user-1',
    email: 'admin@hospital.org',
    fullName: 'System Administrator',
    role: UserRole.superAdmin,
  );

  setUp(() {
    fakeRepository = FakeAuthRepository();
  });

  Widget createWidgetUnderTest({User? user}) {
    return ProviderScope(
      overrides: [
        authRepositoryProvider.overrideWithValue(fakeRepository),
        if (user != null)
          currentUserProvider.overrideWithValue(user)
        else
          currentUserProvider.overrideWithValue(null),
      ],
      child: const MaterialApp(
        home: Scaffold(
          body: ProfileDrawer(),
        ),
      ),
    );
  }

  group('ProfileDrawer Widget Tests', () {
    testWidgets('renders user full name, email, and role display name', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest(user: testUser));
      await tester.pumpAndSettle();

      expect(find.text('System Administrator'), findsOneWidget);
      expect(find.text('admin@hospital.org'), findsOneWidget);
      expect(find.text('Super Admin'), findsOneWidget);
    });

    testWidgets('renders fallback guest title when user is null', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest(user: null));
      await tester.pumpAndSettle();

      expect(find.text('Guest'), findsOneWidget);
      expect(find.text('Not signed in'), findsOneWidget);
    });

    testWidgets('tapping Log Out triggers authController logout', (tester) async {
      await tester.pumpWidget(
        ProviderScope(
          overrides: [
            authRepositoryProvider.overrideWithValue(fakeRepository),
            currentUserProvider.overrideWithValue(testUser),
            authControllerProvider.overrideWith(() => _TestAuthController(testUser, fakeRepository)),
          ],
          child: const MaterialApp(
            home: Scaffold(
              body: ProfileDrawer(),
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      final logoutBtn = find.text('Log Out');
      expect(logoutBtn, findsOneWidget);

      await tester.ensureVisible(logoutBtn);
      await tester.tap(logoutBtn);
      await tester.pumpAndSettle();

      expect(fakeRepository.logoutCallCount, equals(1));
    });
  });
}

class _TestAuthController extends AuthController {
  _TestAuthController(this._initialUser, this._repository);

  final User _initialUser;
  final FakeAuthRepository _repository;

  @override
  Future<AuthState> build() async {
    return AuthStateAuthenticated(_initialUser);
  }

  @override
  Future<void> logout() async {
    await _repository.logout();
    state = const AsyncValue.data(AuthStateUnauthenticated());
  }
}
