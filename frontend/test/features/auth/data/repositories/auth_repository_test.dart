import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/core/storage/token_storage.dart';
import 'package:frontend/features/auth/data/datasources/auth_remote_data_source.dart';
import 'package:frontend/features/auth/data/repositories/auth_repository.dart';
import 'package:frontend/features/auth/domain/models/auth_tokens.dart';
import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';

class FakeAuthRemoteDataSource implements AuthRemoteDataSource {
  String? capturedUsername;
  String? capturedPassword;
  String? capturedInvitationToken;
  String? capturedFirstName;
  String? capturedLastName;
  String? capturedEmail;
  String? capturedSignupPassword;
  String? capturedRefreshToken;
  String? capturedCsrfToken;

  int getCurrentUserCallCount = 0;
  int refreshCallCount = 0;
  int logoutCallCount = 0;

  AuthTokens? tokensToReturn;
  User? signupUserToReturn;
  User? currentUserToReturn;

  Exception? loginException;
  Exception? signupException;
  Exception? refreshException;
  Exception? logoutException;
  Exception? currentUserException;
  final List<Exception?> currentUserExceptionsQueue = [];

  @override
  Future<AuthTokens> login({
    required String username,
    required String password,
  }) async {
    capturedUsername = username;
    capturedPassword = password;
    if (loginException != null) throw loginException!;
    return tokensToReturn!;
  }

  @override
  Future<User> signup({
    required String invitationToken,
    required String firstName,
    required String lastName,
    required String email,
    required String password,
  }) async {
    capturedInvitationToken = invitationToken;
    capturedFirstName = firstName;
    capturedLastName = lastName;
    capturedEmail = email;
    capturedSignupPassword = password;
    if (signupException != null) throw signupException!;
    return signupUserToReturn!;
  }

  @override
  Future<AuthTokens> refresh({
    String? refreshToken,
    String? csrfToken,
  }) async {
    refreshCallCount++;
    capturedRefreshToken = refreshToken;
    capturedCsrfToken = csrfToken;
    if (refreshException != null) throw refreshException!;
    return tokensToReturn!;
  }

  @override
  Future<void> logout({
    String? refreshToken,
    String? csrfToken,
  }) async {
    logoutCallCount++;
    capturedRefreshToken = refreshToken;
    capturedCsrfToken = csrfToken;
    if (logoutException != null) throw logoutException!;
  }

  @override
  Future<User> getCurrentUser() async {
    getCurrentUserCallCount++;
    if (currentUserExceptionsQueue.isNotEmpty) {
      final exc = currentUserExceptionsQueue.removeAt(0);
      if (exc != null) throw exc;
    } else if (currentUserException != null) {
      throw currentUserException!;
    }
    return currentUserToReturn!;
  }
}

class FakeTokenStorage implements TokenStorage {
  String? storedAccessToken;
  String? storedRefreshToken;

  int saveTokensCallCount = 0;
  int clearTokensCallCount = 0;

  @override
  Future<void> saveTokens({
    required String accessToken,
    String? refreshToken,
  }) async {
    saveTokensCallCount++;
    storedAccessToken = accessToken;
    storedRefreshToken = refreshToken;
  }

  @override
  Future<String?> readAccessToken() async {
    return storedAccessToken;
  }

  @override
  Future<String?> readRefreshToken() async {
    return storedRefreshToken;
  }

  @override
  Future<void> clearTokens() async {
    clearTokensCallCount++;
    storedAccessToken = null;
    storedRefreshToken = null;
  }
}

void main() {
  late FakeAuthRemoteDataSource fakeRemoteDataSource;
  late FakeTokenStorage fakeTokenStorage;
  late AuthRepository repository;

  const testUser = User(
    id: 'usr-1',
    email: 'doctor@hospital.org',
    fullName: 'Dr. Gregory House',
    role: UserRole.doctor,
    departmentId: 'dept-1',
    doctorId: 'doc-1',
  );

  const testTokens = AuthTokens(
    accessToken: 'access-123',
    refreshToken: 'refresh-456',
    tokenType: 'bearer',
  );

  setUp(() {
    fakeRemoteDataSource = FakeAuthRemoteDataSource();
    fakeTokenStorage = FakeTokenStorage();
    repository = AuthRepositoryImpl(
      remoteDataSource: fakeRemoteDataSource,
      tokenStorage: fakeTokenStorage,
    );
  });

  group('AuthRepository.login', () {
    test('authenticates, stores tokens securely, fetches user profile, and returns User', () async {
      fakeRemoteDataSource.tokensToReturn = testTokens;
      fakeRemoteDataSource.currentUserToReturn = testUser;

      final user = await repository.login(
        email: 'doctor@hospital.org',
        password: 'Password123!',
      );

      expect(user, equals(testUser));
      expect(fakeRemoteDataSource.capturedUsername, equals('doctor@hospital.org'));
      expect(fakeRemoteDataSource.capturedPassword, equals('Password123!'));
      expect(fakeTokenStorage.saveTokensCallCount, equals(1));
      expect(fakeTokenStorage.storedAccessToken, equals('access-123'));
      expect(fakeTokenStorage.storedRefreshToken, equals('refresh-456'));
      expect(fakeRemoteDataSource.getCurrentUserCallCount, equals(1));
    });

    test('propagates ApiException when remote login fails and leaves storage untouched', () async {
      fakeRemoteDataSource.loginException = const ApiException(
        type: ApiErrorType.unauthorized,
        message: 'Invalid credentials',
        statusCode: 401,
      );

      await expectLater(
        () => repository.login(
          email: 'doctor@hospital.org',
          password: 'WrongPassword',
        ),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.unauthorized),
        ),
      );

      expect(fakeTokenStorage.saveTokensCallCount, equals(0));
      expect(fakeRemoteDataSource.getCurrentUserCallCount, equals(0));
    });

    test('propagates ApiException when getCurrentUser fails after login', () async {
      fakeRemoteDataSource.tokensToReturn = testTokens;
      fakeRemoteDataSource.currentUserException = const ApiException(
        type: ApiErrorType.server,
        message: 'Server error',
        statusCode: 500,
      );

      await expectLater(
        () => repository.login(
          email: 'doctor@hospital.org',
          password: 'Password123!',
        ),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.server),
        ),
      );

      expect(fakeTokenStorage.saveTokensCallCount, equals(1));
      expect(fakeRemoteDataSource.getCurrentUserCallCount, equals(1));
    });
  });

  group('AuthRepository.signup', () {
    test('calls remote signup and returns created User without saving tokens or fetching me', () async {
      fakeRemoteDataSource.signupUserToReturn = testUser;

      final user = await repository.signup(
        invitationToken: 'inv-tok-999',
        firstName: 'Gregory',
        lastName: 'House',
        email: 'doctor@hospital.org',
        password: 'Password123!',
      );

      expect(user, equals(testUser));
      expect(fakeRemoteDataSource.capturedInvitationToken, equals('inv-tok-999'));
      expect(fakeRemoteDataSource.capturedFirstName, equals('Gregory'));
      expect(fakeRemoteDataSource.capturedLastName, equals('House'));
      expect(fakeRemoteDataSource.capturedEmail, equals('doctor@hospital.org'));
      expect(fakeRemoteDataSource.capturedSignupPassword, equals('Password123!'));

      expect(fakeTokenStorage.saveTokensCallCount, equals(0));
      expect(fakeRemoteDataSource.getCurrentUserCallCount, equals(0));
    });

    test('propagates ApiException when signup fails', () async {
      fakeRemoteDataSource.signupException = const ApiException(
        type: ApiErrorType.validation,
        message: 'Invitation expired',
        statusCode: 400,
      );

      await expectLater(
        () => repository.signup(
          invitationToken: 'expired-token',
          firstName: 'Gregory',
          lastName: 'House',
          email: 'doctor@hospital.org',
          password: 'Password123!',
        ),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.validation),
        ),
      );

      expect(fakeTokenStorage.saveTokensCallCount, equals(0));
    });
  });

  group('AuthRepository.logout', () {
    test('reads stored refresh token, revokes on backend, and clears local credentials', () async {
      fakeTokenStorage.storedAccessToken = 'access-to-clear';
      fakeTokenStorage.storedRefreshToken = 'refresh-to-revoke';

      await repository.logout();

      expect(fakeRemoteDataSource.logoutCallCount, equals(1));
      expect(fakeRemoteDataSource.capturedRefreshToken, equals('refresh-to-revoke'));
      expect(fakeTokenStorage.clearTokensCallCount, equals(1));
      expect(fakeTokenStorage.storedAccessToken, isNull);
      expect(fakeTokenStorage.storedRefreshToken, isNull);
    });

    test('clears local credentials even if remote logout fails', () async {
      fakeTokenStorage.storedAccessToken = 'access-to-clear';
      fakeTokenStorage.storedRefreshToken = 'refresh-to-revoke';
      fakeRemoteDataSource.logoutException = const ApiException(
        type: ApiErrorType.server,
        message: 'Server error',
        statusCode: 500,
      );

      await expectLater(
        () => repository.logout(),
        throwsA(isA<ApiException>()),
      );

      expect(fakeTokenStorage.clearTokensCallCount, equals(1));
      expect(fakeTokenStorage.storedAccessToken, isNull);
      expect(fakeTokenStorage.storedRefreshToken, isNull);
    });
  });

  group('AuthRepository.restoreSession', () {
    test('returns null when no stored access token exists', () async {
      fakeTokenStorage.storedAccessToken = null;
      fakeTokenStorage.storedRefreshToken = null;

      final user = await repository.restoreSession();

      expect(user, isNull);
      expect(fakeRemoteDataSource.getCurrentUserCallCount, equals(0));
      expect(fakeRemoteDataSource.refreshCallCount, equals(0));
    });

    test('returns User when stored access token is valid and getCurrentUser succeeds', () async {
      fakeTokenStorage.storedAccessToken = 'valid-access-token';
      fakeTokenStorage.storedRefreshToken = 'valid-refresh-token';
      fakeRemoteDataSource.currentUserToReturn = testUser;

      final user = await repository.restoreSession();

      expect(user, equals(testUser));
      expect(fakeRemoteDataSource.getCurrentUserCallCount, equals(1));
      expect(fakeRemoteDataSource.refreshCallCount, equals(0));
    });

    test('refreshes session, saves rotated tokens, and returns User when access token expired (401)', () async {
      fakeTokenStorage.storedAccessToken = 'expired-access-token';
      fakeTokenStorage.storedRefreshToken = 'active-refresh-token';

      // First getCurrentUser() throws 401 Unauthorized; second call after refresh succeeds
      fakeRemoteDataSource.currentUserExceptionsQueue.add(
        const ApiException(
          type: ApiErrorType.unauthorized,
          message: 'Access token expired',
          statusCode: 401,
        ),
      );

      const rotatedTokens = AuthTokens(
        accessToken: 'rotated-access-token',
        refreshToken: 'rotated-refresh-token',
        tokenType: 'bearer',
      );
      fakeRemoteDataSource.tokensToReturn = rotatedTokens;
      fakeRemoteDataSource.currentUserToReturn = testUser;

      final user = await repository.restoreSession();

      expect(user, equals(testUser));
      expect(fakeRemoteDataSource.refreshCallCount, equals(1));
      expect(fakeRemoteDataSource.capturedRefreshToken, equals('active-refresh-token'));
      expect(fakeTokenStorage.storedAccessToken, equals('rotated-access-token'));
      expect(fakeTokenStorage.storedRefreshToken, equals('rotated-refresh-token'));
      expect(fakeRemoteDataSource.getCurrentUserCallCount, equals(2));
    });

    test('clears storage and returns null when refresh token is revoked or expired (401)', () async {
      fakeTokenStorage.storedAccessToken = 'expired-access-token';
      fakeTokenStorage.storedRefreshToken = 'revoked-refresh-token';

      fakeRemoteDataSource.currentUserExceptionsQueue.add(
        const ApiException(
          type: ApiErrorType.unauthorized,
          message: 'Access token expired',
          statusCode: 401,
        ),
      );

      fakeRemoteDataSource.refreshException = const ApiException(
        type: ApiErrorType.unauthorized,
        message: 'Refresh token expired or revoked',
        statusCode: 401,
      );

      final user = await repository.restoreSession();

      expect(user, isNull);
      expect(fakeRemoteDataSource.refreshCallCount, equals(1));
      expect(fakeTokenStorage.clearTokensCallCount, equals(1));
      expect(fakeTokenStorage.storedAccessToken, isNull);
      expect(fakeTokenStorage.storedRefreshToken, isNull);
    });

    test('rethrows ApiException without clearing tokens when network error occurs during restore', () async {
      fakeTokenStorage.storedAccessToken = 'valid-access-token';
      fakeRemoteDataSource.currentUserException = const ApiException(
        type: ApiErrorType.network,
        message: 'No internet connection',
      );

      await expectLater(
        () => repository.restoreSession(),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.network),
        ),
      );

      // Do NOT log the user out due to a temporary network outage!
      expect(fakeTokenStorage.clearTokensCallCount, equals(0));
      expect(fakeTokenStorage.storedAccessToken, equals('valid-access-token'));
    });
  });
}
