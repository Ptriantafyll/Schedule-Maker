import 'package:dio/dio.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/core/network/api_client.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/auth/data/datasources/auth_remote_data_source.dart';
import 'package:frontend/features/auth/domain/models/auth_tokens.dart';
import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';

class FakeApiClient implements ApiClient {
  String? capturedPath;
  Object? capturedData;
  String? capturedContentType;
  bool? capturedRequiresAuth;

  Response<dynamic>? responseToReturn;
  Exception? exceptionToThrow;

  @override
  Future<Response<T>> post<T>(
    String path, {
    Object? data,
    Map<String, dynamic>? queryParameters,
    String? contentType,
    bool requiresAuth = true,
  }) async {
    capturedPath = path;
    capturedData = data;
    capturedContentType = contentType;
    capturedRequiresAuth = requiresAuth;

    if (exceptionToThrow != null) {
      throw exceptionToThrow!;
    }

    return responseToReturn as Response<T>;
  }

  @override
  Future<Response<T>> get<T>(
    String path, {
    Map<String, dynamic>? queryParameters,
    bool requiresAuth = true,
  }) async {
    capturedPath = path;
    capturedRequiresAuth = requiresAuth;

    if (exceptionToThrow != null) {
      throw exceptionToThrow!;
    }

    return responseToReturn as Response<T>;
  }
}

void main() {
  late FakeApiClient fakeApiClient;
  late AuthRemoteDataSource dataSource;

  setUp(() {
    fakeApiClient = FakeApiClient();
    dataSource = AuthRemoteDataSource(fakeApiClient);
  });

  group('AuthRemoteDataSource.login', () {
    const validTokenJson = {
      'access_token': 'acc-123',
      'token_type': 'bearer',
      'refresh_token': 'ref-456',
      'csrf_token': 'csrf-789',
    };

    test('sends form-urlencoded POST to /auth/login without requiring auth and returns AuthTokens', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        data: validTokenJson,
        statusCode: 200,
        requestOptions: RequestOptions(path: '/auth/login'),
      );

      final tokens = await dataSource.login(
        username: 'doctor@hospital.org',
        password: 'Password123!',
      );

      expect(tokens, isA<AuthTokens>());
      expect(tokens.accessToken, equals('acc-123'));
      expect(tokens.refreshToken, equals('ref-456'));
      expect(tokens.csrfToken, equals('csrf-789'));
      expect(tokens.tokenType, equals('bearer'));

      expect(fakeApiClient.capturedPath, equals('/auth/login'));
      expect(
        fakeApiClient.capturedData,
        equals({
          'username': 'doctor@hospital.org',
          'password': 'Password123!',
        }),
      );
      expect(fakeApiClient.capturedContentType, equals('application/x-www-form-urlencoded'));
      expect(fakeApiClient.capturedRequiresAuth, isFalse);
    });

    test('propagates ApiException when ApiClient throws unauthorized', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.unauthorized,
        message: 'Invalid credentials',
        statusCode: 401,
      );

      expect(
        () => dataSource.login(
          username: 'doctor@hospital.org',
          password: 'WrongPassword',
        ),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.unauthorized),
        ),
      );
    });

    test('propagates ApiException on validation error', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.validation,
        message: 'Validation failed',
        statusCode: 422,
      );

      expect(
        () => dataSource.login(
          username: 'invalid-email',
          password: '',
        ),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.validation),
        ),
      );
    });

    test('throws ApiException when server response data is null', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        data: null,
        statusCode: 200,
        requestOptions: RequestOptions(path: '/auth/login'),
      );

      expect(
        () => dataSource.login(
          username: 'doctor@hospital.org',
          password: 'Password123!',
        ),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.unknown),
        ),
      );
    });
  });

  group('AuthRemoteDataSource.signup', () {
    const validUserJson = {
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

    test('sends JSON POST to /auth/signup without requiring auth and returns domain User', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        data: validUserJson,
        statusCode: 201,
        requestOptions: RequestOptions(path: '/auth/signup'),
      );

      final user = await dataSource.signup(
        invitationToken: 'inv-tok-123',
        firstName: 'Gregory',
        lastName: 'House',
        email: 'doctor@hospital.org',
        password: 'Password123!',
      );

      expect(user, isA<User>());
      expect(user.id, equals('usr-123'));
      expect(user.email, equals('doctor@hospital.org'));
      expect(user.fullName, equals('Dr. Gregory House'));
      expect(user.role, equals(UserRole.doctor));
      expect(user.departmentId, equals('dept-456'));
      expect(user.doctorId, equals('doc-789'));

      expect(fakeApiClient.capturedPath, equals('/auth/signup'));
      expect(
        fakeApiClient.capturedData,
        equals({
          'invitation_token': 'inv-tok-123',
          'first_name': 'Gregory',
          'last_name': 'House',
          'email': 'doctor@hospital.org',
          'password': 'Password123!',
        }),
      );
      expect(fakeApiClient.capturedRequiresAuth, isFalse);
    });

    test('propagates ApiException when invitation is invalid or expired', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.validation,
        message: 'Invalid or expired invitation token',
        statusCode: 400,
      );

      expect(
        () => dataSource.signup(
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
    });

    test('throws ApiException when signup response data is null', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        data: null,
        statusCode: 201,
        requestOptions: RequestOptions(path: '/auth/signup'),
      );

      expect(
        () => dataSource.signup(
          invitationToken: 'inv-tok-123',
          firstName: 'Gregory',
          lastName: 'House',
          email: 'doctor@hospital.org',
          password: 'Password123!',
        ),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.unknown),
        ),
      );
    });
  });

  group('AuthRemoteDataSource.refresh', () {
    const validTokenJson = {
      'access_token': 'new-acc-123',
      'token_type': 'bearer',
      'refresh_token': 'new-ref-456',
      'csrf_token': 'new-csrf-789',
    };

    test('sends POST to /auth/refresh with body when refresh token is provided and returns AuthTokens', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        data: validTokenJson,
        statusCode: 200,
        requestOptions: RequestOptions(path: '/auth/refresh'),
      );

      final tokens = await dataSource.refresh(
        refreshToken: 'old-ref-123',
        csrfToken: 'csrf-abc',
      );

      expect(tokens, isA<AuthTokens>());
      expect(tokens.accessToken, equals('new-acc-123'));
      expect(tokens.refreshToken, equals('new-ref-456'));
      expect(tokens.csrfToken, equals('new-csrf-789'));

      expect(fakeApiClient.capturedPath, equals('/auth/refresh'));
      expect(
        fakeApiClient.capturedData,
        equals({
          'refresh_token': 'old-ref-123',
          'csrf_token': 'csrf-abc',
        }),
      );
      expect(fakeApiClient.capturedRequiresAuth, isFalse);
    });

    test('sends POST to /auth/refresh with null data when no tokens provided and returns AuthTokens', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        data: validTokenJson,
        statusCode: 200,
        requestOptions: RequestOptions(path: '/auth/refresh'),
      );

      final tokens = await dataSource.refresh();

      expect(tokens, isA<AuthTokens>());
      expect(tokens.accessToken, equals('new-acc-123'));

      expect(fakeApiClient.capturedPath, equals('/auth/refresh'));
      expect(fakeApiClient.capturedData, isNull);
      expect(fakeApiClient.capturedRequiresAuth, isFalse);
    });

    test('propagates ApiException when refresh token is invalid or revoked', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.unauthorized,
        message: 'Invalid or revoked refresh token',
        statusCode: 401,
      );

      expect(
        () => dataSource.refresh(refreshToken: 'revoked-token'),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.unauthorized),
        ),
      );
    });

    test('throws ApiException when refresh response data is null', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        data: null,
        statusCode: 200,
        requestOptions: RequestOptions(path: '/auth/refresh'),
      );

      expect(
        () => dataSource.refresh(refreshToken: 'some-token'),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.unknown),
        ),
      );
    });
  });

  group('AuthRemoteDataSource.logout', () {
    test('sends POST to /auth/logout with body when tokens provided and completes', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        data: {'detail': 'Logged out successfully'},
        statusCode: 200,
        requestOptions: RequestOptions(path: '/auth/logout'),
      );

      await dataSource.logout(
        refreshToken: 'ref-to-revoke',
        csrfToken: 'csrf-to-revoke',
      );

      expect(fakeApiClient.capturedPath, equals('/auth/logout'));
      expect(
        fakeApiClient.capturedData,
        equals({
          'refresh_token': 'ref-to-revoke',
          'csrf_token': 'csrf-to-revoke',
        }),
      );
      expect(fakeApiClient.capturedRequiresAuth, isFalse);
    });

    test('sends POST to /auth/logout with null data when no tokens provided', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        data: {'detail': 'Logged out successfully'},
        statusCode: 200,
        requestOptions: RequestOptions(path: '/auth/logout'),
      );

      await dataSource.logout();

      expect(fakeApiClient.capturedPath, equals('/auth/logout'));
      expect(fakeApiClient.capturedData, isNull);
      expect(fakeApiClient.capturedRequiresAuth, isFalse);
    });

    test('propagates ApiException when server fails during logout', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.server,
        message: 'Internal server error',
        statusCode: 500,
      );

      expect(
        () => dataSource.logout(),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.server),
        ),
      );
    });
  });

  group('AuthRemoteDataSource.getCurrentUser', () {
    const validUserJson = {
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

    test('sends GET to /auth/me with requiresAuth: true and returns domain User', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        data: validUserJson,
        statusCode: 200,
        requestOptions: RequestOptions(path: '/auth/me'),
      );

      final user = await dataSource.getCurrentUser();

      expect(user, isA<User>());
      expect(user.id, equals('usr-123'));
      expect(user.email, equals('doctor@hospital.org'));
      expect(user.fullName, equals('Dr. Gregory House'));
      expect(user.role, equals(UserRole.doctor));
      expect(user.departmentId, equals('dept-456'));
      expect(user.doctorId, equals('doc-789'));

      expect(fakeApiClient.capturedPath, equals('/auth/me'));
      expect(fakeApiClient.capturedRequiresAuth, isTrue);
    });

    test('propagates ApiException when unauthenticated', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.unauthorized,
        message: 'Not authenticated',
        statusCode: 401,
      );

      expect(
        () => dataSource.getCurrentUser(),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.unauthorized),
        ),
      );
    });

    test('throws ApiException when current user response data is null', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        data: null,
        statusCode: 200,
        requestOptions: RequestOptions(path: '/auth/me'),
      );

      expect(
        () => dataSource.getCurrentUser(),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.unknown),
        ),
      );
    });
  });
}
