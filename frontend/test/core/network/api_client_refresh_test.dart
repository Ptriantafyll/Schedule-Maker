import 'dart:convert';
import 'dart:typed_data';

import 'package:dio/dio.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/core/network/api_client.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/core/storage/token_storage.dart';

class FakeTokenStorage implements TokenStorage {
  String? storedAccessToken;
  String? storedRefreshToken;
  int clearTokensCallCount = 0;
  int saveTokensCallCount = 0;

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
  Future<String?> readAccessToken() async => storedAccessToken;

  @override
  Future<String?> readRefreshToken() async => storedRefreshToken;

  @override
  Future<void> clearTokens() async {
    clearTokensCallCount++;
    storedAccessToken = null;
    storedRefreshToken = null;
  }
}

class FakeHttpClientAdapter implements HttpClientAdapter {
  final List<RequestOptions> capturedRequests = [];
  Future<ResponseBody> Function(RequestOptions options)? handler;

  @override
  Future<ResponseBody> fetch(
    RequestOptions options,
    Stream<Uint8List>? requestStream,
    Future<void>? cancelFuture,
  ) async {
    capturedRequests.add(options);
    if (handler != null) {
      return handler!(options);
    }
    return ResponseBody.fromString(
      jsonEncode({'status': 'ok'}),
      200,
      headers: {
        Headers.contentTypeHeader: [Headers.jsonContentType],
      },
    );
  }

  @override
  void close({bool force = false}) {}
}

void main() {
  late Dio dio;
  late FakeHttpClientAdapter fakeAdapter;
  late FakeTokenStorage fakeTokenStorage;
  late ApiClient apiClient;

  setUp(() {
    fakeAdapter = FakeHttpClientAdapter();
    fakeTokenStorage = FakeTokenStorage();

    dio = Dio(
      BaseOptions(
        baseUrl: 'http://localhost:8000/api/v1',
        connectTimeout: const Duration(seconds: 15),
        receiveTimeout: const Duration(seconds: 15),
        sendTimeout: const Duration(seconds: 15),
        contentType: 'application/json',
      ),
    );
    dio.httpClientAdapter = fakeAdapter;

    apiClient = ApiClient(dio, fakeTokenStorage);
  });

  group('ApiClient Silent 401 Refresh & Single-Flight Queue', () {
    test('single protected request that gets 401 refreshes once and retries successfully', () async {
      fakeTokenStorage.storedAccessToken = 'expired-access-token';
      fakeTokenStorage.storedRefreshToken = 'valid-refresh-token';

      fakeAdapter.handler = (options) async {
        // 1. First call to /doctors with expired token -> returns 401
        if (options.path.contains('/doctors') &&
            options.headers['Authorization'] == 'Bearer expired-access-token') {
          return ResponseBody.fromString(
            jsonEncode({'detail': 'Could not validate credentials'}),
            401,
            headers: {Headers.contentTypeHeader: [Headers.jsonContentType]},
          );
        }

        // 2. Call to /auth/refresh with refresh token -> returns 200 with new tokens
        if (options.path.contains('/auth/refresh')) {
          return ResponseBody.fromString(
            jsonEncode({
              'access_token': 'new-access-token',
              'refresh_token': 'new-refresh-token',
              'token_type': 'bearer',
            }),
            200,
            headers: {Headers.contentTypeHeader: [Headers.jsonContentType]},
          );
        }

        // 3. Retried call to /doctors with new access token -> returns 200
        if (options.path.contains('/doctors') &&
            options.headers['Authorization'] == 'Bearer new-access-token') {
          return ResponseBody.fromString(
            jsonEncode({'doctors': ['Dr. House']}),
            200,
            headers: {Headers.contentTypeHeader: [Headers.jsonContentType]},
          );
        }

        throw UnimplementedError('Unexpected request to ${options.path}');
      };

      final response = await apiClient.get<Map<String, dynamic>>('/doctors', requiresAuth: true);

      expect(response.statusCode, equals(200));
      expect(response.data, equals({'doctors': ['Dr. House']}));

      // Verify tokens were updated in storage
      expect(fakeTokenStorage.storedAccessToken, equals('new-access-token'));
      expect(fakeTokenStorage.storedRefreshToken, equals('new-refresh-token'));
      expect(fakeTokenStorage.saveTokensCallCount, equals(1));

      // Verify only 1 refresh call was made
      final refreshCalls = fakeAdapter.capturedRequests.where((r) => r.path.contains('/auth/refresh'));
      expect(refreshCalls.length, equals(1));
    });

    test('concurrent protected requests sharing a single in-flight refresh (single-flight)', () async {
      fakeTokenStorage.storedAccessToken = 'expired-access-token';
      fakeTokenStorage.storedRefreshToken = 'valid-refresh-token';

      var refreshRequestCount = 0;

      fakeAdapter.handler = (options) async {
        if (options.headers['Authorization'] == 'Bearer expired-access-token') {
          return ResponseBody.fromString(
            jsonEncode({'detail': 'Token expired'}),
            401,
            headers: {Headers.contentTypeHeader: [Headers.jsonContentType]},
          );
        }

        if (options.path.contains('/auth/refresh')) {
          refreshRequestCount++;
          // Simulate network latency so concurrent calls wait on the in-flight future
          await Future.delayed(const Duration(milliseconds: 50));
          return ResponseBody.fromString(
            jsonEncode({
              'access_token': 'new-access-token',
              'refresh_token': 'new-refresh-token',
              'token_type': 'bearer',
            }),
            200,
            headers: {Headers.contentTypeHeader: [Headers.jsonContentType]},
          );
        }

        if (options.headers['Authorization'] == 'Bearer new-access-token') {
          return ResponseBody.fromString(
            jsonEncode({'data': options.path}),
            200,
            headers: {Headers.contentTypeHeader: [Headers.jsonContentType]},
          );
        }

        throw UnimplementedError('Unexpected request to ${options.path}');
      };

      // Fire 3 concurrent requests simultaneously while access token is expired
      final futureShifts = apiClient.get<Map<String, dynamic>>('/shifts', requiresAuth: true);
      final futureDoctors = apiClient.get<Map<String, dynamic>>('/doctors', requiresAuth: true);
      final futureRequests = apiClient.get<Map<String, dynamic>>('/requests', requiresAuth: true);

      final results = await Future.wait([futureShifts, futureDoctors, futureRequests]);

      for (final res in results) {
        expect(res.statusCode, equals(200));
      }

      // CRITICAL: Exactly ONE refresh network call must be made despite 3 concurrent 401s
      expect(refreshRequestCount, equals(1));
      expect(fakeTokenStorage.storedAccessToken, equals('new-access-token'));
      expect(fakeTokenStorage.storedRefreshToken, equals('new-refresh-token'));
    });

    test('public requests getting 401 do not trigger refresh', () async {
      fakeTokenStorage.storedAccessToken = 'some-access-token';
      fakeTokenStorage.storedRefreshToken = 'some-refresh-token';

      fakeAdapter.handler = (options) async {
        if (options.path.contains('/auth/login')) {
          return ResponseBody.fromString(
            jsonEncode({'detail': 'Incorrect email or password'}),
            401,
            headers: {Headers.contentTypeHeader: [Headers.jsonContentType]},
          );
        }
        throw UnimplementedError('Unexpected request to ${options.path}');
      };

      await expectLater(
        () => apiClient.post<Map<String, dynamic>>(
          '/auth/login',
          data: {'username': 'wrong@hospital.org', 'password': 'bad'},
          requiresAuth: false,
        ),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.unauthorized),
        ),
      );

      // Verify no refresh call was ever dispatched
      final refreshCalls = fakeAdapter.capturedRequests.where((r) => r.path.contains('/auth/refresh'));
      expect(refreshCalls.length, equals(0));
      // Storage should not be cleared for public login failures
      expect(fakeTokenStorage.clearTokensCallCount, equals(0));
    });

    test('when refresh token is expired or revoked (401 on /auth/refresh), clears tokens and throws unauthorized', () async {
      fakeTokenStorage.storedAccessToken = 'expired-access-token';
      fakeTokenStorage.storedRefreshToken = 'revoked-refresh-token';

      fakeAdapter.handler = (options) async {
        if (options.path.contains('/auth/refresh')) {
          return ResponseBody.fromString(
            jsonEncode({'detail': 'Refresh token expired or revoked'}),
            401,
            headers: {Headers.contentTypeHeader: [Headers.jsonContentType]},
          );
        }

        // Original request
        return ResponseBody.fromString(
          jsonEncode({'detail': 'Unauthorized'}),
          401,
          headers: {Headers.contentTypeHeader: [Headers.jsonContentType]},
        );
      };

      await expectLater(
        () => apiClient.get<Map<String, dynamic>>('/shifts', requiresAuth: true),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.unauthorized),
        ),
      );

      // Session must be cleared so the user is routed to login
      expect(fakeTokenStorage.clearTokensCallCount, equals(1));
      expect(fakeTokenStorage.storedAccessToken, isNull);
      expect(fakeTokenStorage.storedRefreshToken, isNull);
    });

    test('when network error occurs during refresh, does not clear stored tokens and throws network exception', () async {
      fakeTokenStorage.storedAccessToken = 'expired-access-token';
      fakeTokenStorage.storedRefreshToken = 'valid-refresh-token';

      fakeAdapter.handler = (options) async {
        if (options.path.contains('/auth/refresh')) {
          throw DioException(
            requestOptions: options,
            type: DioExceptionType.connectionError,
            error: 'Connection refused',
          );
        }

        // Initial 401
        return ResponseBody.fromString(
          jsonEncode({'detail': 'Token expired'}),
          401,
          headers: {Headers.contentTypeHeader: [Headers.jsonContentType]},
        );
      };

      await expectLater(
        () => apiClient.get<Map<String, dynamic>>('/shifts', requiresAuth: true),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.network),
        ),
      );

      // Must NOT clear tokens on temporary network failure
      expect(fakeTokenStorage.clearTokensCallCount, equals(0));
      expect(fakeTokenStorage.storedAccessToken, equals('expired-access-token'));
    });

    test('does not retry infinitely if 401 persists after refresh', () async {
      fakeTokenStorage.storedAccessToken = 'expired-access-token';
      fakeTokenStorage.storedRefreshToken = 'valid-refresh-token';

      fakeAdapter.handler = (options) async {
        if (options.path.contains('/auth/refresh')) {
          return ResponseBody.fromString(
            jsonEncode({
              'access_token': 'new-access-token',
              'refresh_token': 'new-refresh-token',
              'token_type': 'bearer',
            }),
            200,
            headers: {Headers.contentTypeHeader: [Headers.jsonContentType]},
          );
        }

        // Both original and retried call return 401
        return ResponseBody.fromString(
          jsonEncode({'detail': 'Forbidden or deauthorized'}),
          401,
          headers: {Headers.contentTypeHeader: [Headers.jsonContentType]},
        );
      };

      await expectLater(
        () => apiClient.get<Map<String, dynamic>>('/shifts', requiresAuth: true),
        throwsA(
          isA<ApiException>().having((e) => e.type, 'type', ApiErrorType.unauthorized),
        ),
      );

      // Verify refresh was only called ONCE, not in an infinite loop
      final refreshCalls = fakeAdapter.capturedRequests.where((r) => r.path.contains('/auth/refresh'));
      expect(refreshCalls.length, equals(1));
    });
  });
}
