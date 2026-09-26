import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/auth/data/repositories/auth_repository.dart';
import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';
import 'package:frontend/features/auth/domain/state/auth_state.dart';
import 'package:frontend/features/auth/presentation/controllers/auth_controller.dart';

class FakeAuthRepository implements AuthRepository {
  User? userToReturnOnLogin;
  User? userToReturnOnSignup;
  User? userToReturnOnRestoreSession;

  Exception? loginException;
  Exception? signupException;
  Exception? logoutException;
  Exception? restoreSessionException;

  int loginCallCount = 0;
  int signupCallCount = 0;
  int logoutCallCount = 0;
  int restoreSessionCallCount = 0;

  String? capturedLoginEmail;
  String? capturedLoginPassword;
  String? capturedSignupInvitationToken;
  String? capturedSignupFirstName;
  String? capturedSignupLastName;
  String? capturedSignupEmail;
  String? capturedSignupPassword;

  @override
  Future<User> login({
    required String email,
    required String password,
  }) async {
    loginCallCount++;
    capturedLoginEmail = email;
    capturedLoginPassword = password;
    if (loginException != null) throw loginException!;
    return userToReturnOnLogin!;
  }

  @override
  Future<User> signup({
    required String invitationToken,
    required String firstName,
    required String lastName,
    required String email,
    required String password,
  }) async {
    signupCallCount++;
    capturedSignupInvitationToken = invitationToken;
    capturedSignupFirstName = firstName;
    capturedSignupLastName = lastName;
    capturedSignupEmail = email;
    capturedSignupPassword = password;
    if (signupException != null) throw signupException!;
    return userToReturnOnSignup!;
  }

  @override
  Future<void> logout() async {
    logoutCallCount++;
    if (logoutException != null) throw logoutException!;
  }

  @override
  Future<User?> restoreSession() async {
    restoreSessionCallCount++;
    if (restoreSessionException != null) throw restoreSessionException!;
    return userToReturnOnRestoreSession;
  }
}

void main() {
  late FakeAuthRepository fakeRepository;
  late ProviderContainer container;

  const testUser = User(
    id: 'usr-1',
    email: 'doctor@hospital.org',
    fullName: 'Dr. Gregory House',
    role: UserRole.doctor,
    departmentId: 'dept-1',
    doctorId: 'doc-1',
  );

  setUp(() {
    fakeRepository = FakeAuthRepository();
  });

  ProviderContainer createContainer() {
    final c = ProviderContainer(
      overrides: [
        authRepositoryProvider.overrideWithValue(fakeRepository),
      ],
    );
    addTearDown(c.dispose);
    return c;
  }

  group('AuthController.build (Session Restoration on Startup)', () {
    test('resolves to AuthStateUnauthenticated when restoreSession returns null', () async {
      fakeRepository.userToReturnOnRestoreSession = null;
      container = createContainer();

      final state = await container.read(authControllerProvider.future);

      expect(state, equals(const AuthStateUnauthenticated()));
      expect(container.read(authControllerProvider).value, equals(const AuthStateUnauthenticated()));
      expect(container.read(isAuthenticatedProvider), isFalse);
      expect(container.read(currentUserProvider), isNull);
      expect(container.read(userRoleProvider), isNull);
      expect(fakeRepository.restoreSessionCallCount, equals(1));
    });

    test('resolves to AuthStateAuthenticated when restoreSession returns a user', () async {
      fakeRepository.userToReturnOnRestoreSession = testUser;
      container = createContainer();

      final state = await container.read(authControllerProvider.future);

      expect(state, equals(const AuthStateAuthenticated(testUser)));
      expect(container.read(isAuthenticatedProvider), isTrue);
      expect(container.read(currentUserProvider), equals(testUser));
      expect(container.read(userRoleProvider), equals(UserRole.doctor));
      expect(fakeRepository.restoreSessionCallCount, equals(1));
    });

    test('resolves to AsyncError when restoreSession throws an unexpected error', () async {
      fakeRepository.restoreSessionException = const ApiException(
        type: ApiErrorType.server,
        message: 'Database unavailable',
        statusCode: 500,
      );
      container = createContainer();

      await expectLater(
        container.read(authControllerProvider.future),
        throwsA(isA<ApiException>()),
      );

      final state = container.read(authControllerProvider);
      expect(state, isA<AsyncError<AuthState>>());
      expect(container.read(isAuthenticatedProvider), isFalse);
      expect(container.read(currentUserProvider), isNull);
    });
  });

  group('AuthController.login', () {
    test('transitions to Authenticated on successful login and updates derived providers', () async {
      fakeRepository.userToReturnOnRestoreSession = null;
      fakeRepository.userToReturnOnLogin = testUser;
      container = createContainer();

      // Wait for initial build
      await container.read(authControllerProvider.future);

      await container.read(authControllerProvider.notifier).login(
        email: 'doctor@hospital.org',
        password: 'Password123!',
      );

      expect(fakeRepository.capturedLoginEmail, equals('doctor@hospital.org'));
      expect(fakeRepository.capturedLoginPassword, equals('Password123!'));

      final state = container.read(authControllerProvider);
      expect(state.value, equals(const AuthStateAuthenticated(testUser)));
      expect(container.read(isAuthenticatedProvider), isTrue);
      expect(container.read(currentUserProvider), equals(testUser));
      expect(container.read(userRoleProvider), equals(UserRole.doctor));
    });

    test('transitions to AsyncError and rethrows when login fails', () async {
      fakeRepository.userToReturnOnRestoreSession = null;
      fakeRepository.loginException = const ApiException(
        type: ApiErrorType.unauthorized,
        message: 'Invalid credentials',
        statusCode: 401,
      );
      container = createContainer();

      await container.read(authControllerProvider.future);

      await expectLater(
        () => container.read(authControllerProvider.notifier).login(
          email: 'doctor@hospital.org',
          password: 'WrongPassword',
        ),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.unauthorized),
        ),
      );

      final state = container.read(authControllerProvider);
      expect(state, isA<AsyncError<AuthState>>());
      expect(container.read(isAuthenticatedProvider), isFalse);
      expect(container.read(currentUserProvider), isNull);
    });
  });

  group('AuthController.signup', () {
    test('calls repository, returns created user, and keeps controller in Unauthenticated state', () async {
      fakeRepository.userToReturnOnRestoreSession = null;
      fakeRepository.userToReturnOnSignup = testUser;
      container = createContainer();

      await container.read(authControllerProvider.future);

      final createdUser = await container.read(authControllerProvider.notifier).signup(
        invitationToken: 'inv-tok-123',
        firstName: 'Gregory',
        lastName: 'House',
        email: 'doctor@hospital.org',
        password: 'Password123!',
      );

      expect(createdUser, equals(testUser));
      expect(fakeRepository.capturedSignupInvitationToken, equals('inv-tok-123'));
      expect(fakeRepository.capturedSignupFirstName, equals('Gregory'));
      expect(fakeRepository.capturedSignupLastName, equals('House'));
      expect(fakeRepository.capturedSignupEmail, equals('doctor@hospital.org'));

      // Crucial requirement: User remains unauthenticated after invitation signup
      final state = container.read(authControllerProvider);
      expect(state.value, equals(const AuthStateUnauthenticated()));
      expect(container.read(isAuthenticatedProvider), isFalse);
      expect(container.read(currentUserProvider), isNull);
    });

    test('propagates exception when signup fails', () async {
      fakeRepository.userToReturnOnRestoreSession = null;
      fakeRepository.signupException = const ApiException(
        type: ApiErrorType.validation,
        message: 'Invalid invitation',
        statusCode: 400,
      );
      container = createContainer();

      await container.read(authControllerProvider.future);

      await expectLater(
        () => container.read(authControllerProvider.notifier).signup(
          invitationToken: 'bad-token',
          firstName: 'Gregory',
          lastName: 'House',
          email: 'doctor@hospital.org',
          password: 'Password123!',
        ),
        throwsA(isA<ApiException>()),
      );
    });
  });

  group('AuthController.logout', () {
    test('calls repository logout and resets state to AuthStateUnauthenticated', () async {
      fakeRepository.userToReturnOnRestoreSession = testUser;
      container = createContainer();

      // Ensure user is authenticated initially
      await container.read(authControllerProvider.future);
      expect(container.read(isAuthenticatedProvider), isTrue);

      await container.read(authControllerProvider.notifier).logout();

      expect(fakeRepository.logoutCallCount, equals(1));
      final state = container.read(authControllerProvider);
      expect(state.value, equals(const AuthStateUnauthenticated()));
      expect(container.read(isAuthenticatedProvider), isFalse);
      expect(container.read(currentUserProvider), isNull);
      expect(container.read(userRoleProvider), isNull);
    });

    test('propagates error if logout fails but still resets unauthenticated if desired', () async {
      fakeRepository.userToReturnOnRestoreSession = testUser;
      fakeRepository.logoutException = const ApiException(
        type: ApiErrorType.server,
        message: 'Server error during logout',
        statusCode: 500,
      );
      container = createContainer();

      await container.read(authControllerProvider.future);

      await expectLater(
        () => container.read(authControllerProvider.notifier).logout(),
        throwsA(isA<ApiException>()),
      );
    });
  });
}
