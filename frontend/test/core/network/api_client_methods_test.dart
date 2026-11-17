import 'dart:convert';
import 'dart:typed_data';

import 'package:dio/dio.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/core/network/api_client.dart';
import 'package:frontend/core/storage/token_storage.dart';

class _FakeTokenStorage implements TokenStorage {
  String? storedAccessToken;
  String? storedRefreshToken;

  @override
  Future<void> saveTokens({
    required String accessToken,
    String? refreshToken,
  }) async {
    storedAccessToken = accessToken;
    storedRefreshToken = refreshToken;
  }

  @override
  Future<String?> readAccessToken() async => storedAccessToken;

  @override
  Future<String?> readRefreshToken() async => storedRefreshToken;

  @override
  Future<void> clearTokens() async {
    storedAccessToken = null;
    storedRefreshToken = null;
  }
}

class _FakeHttpClientAdapter implements HttpClientAdapter {
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
  late _FakeHttpClientAdapter fakeAdapter;
  late _FakeTokenStorage fakeTokenStorage;
  late ApiClient apiClient;

  setUp(() {
    fakeAdapter = _FakeHttpClientAdapter();
    fakeTokenStorage = _FakeTokenStorage()
      ..storedAccessToken = 'valid-access-token'
      ..storedRefreshToken = 'valid-refresh-token';

    dio = Dio(
      BaseOptions(
        baseUrl: 'http://localhost:8000',
        contentType: 'application/json',
      ),
    )..httpClientAdapter = fakeAdapter;

    apiClient = ApiClient(dio, fakeTokenStorage);
  });

  group('ApiClient.patch', () {
    test('sends PATCH request with JSON body and Authorization header', () async {
      fakeAdapter.handler = (options) async {
        return ResponseBody.fromString(
          jsonEncode({'id': 'team-1', 'name': 'Cardiology A'}),
          200,
          headers: {
            Headers.contentTypeHeader: [Headers.jsonContentType],
          },
        );
      };

      final response = await apiClient.patch<Map<String, dynamic>>(
        '/api/v1/teams/team-1',
        data: {'name': 'Cardiology A'},
      );

      expect(response.statusCode, 200);
      expect(response.data?['name'], 'Cardiology A');
      expect(fakeAdapter.capturedRequests.length, 1);

      final req = fakeAdapter.capturedRequests.single;
      expect(req.method, 'PATCH');
      expect(req.path, '/api/v1/teams/team-1');
      expect(req.data, {'name': 'Cardiology A'});
      expect(req.headers['Authorization'], 'Bearer valid-access-token');
    });

    test('refreshes token once on 401 and retries PATCH request', () async {
      fakeTokenStorage.storedAccessToken = 'expired-token';

      fakeAdapter.handler = (options) async {
        if (options.path == '/api/v1/auth/refresh') {
          return ResponseBody.fromString(
            jsonEncode({
              'access_token': 'fresh-token',
              'refresh_token': 'valid-refresh-token',
            }),
            200,
            headers: {
              Headers.contentTypeHeader: [Headers.jsonContentType],
            },
          );
        }

        if (options.headers['Authorization'] == 'Bearer expired-token') {
          return ResponseBody.fromString(
            jsonEncode({'detail': 'Expired'}),
            401,
            headers: {
              Headers.contentTypeHeader: [Headers.jsonContentType],
            },
          );
        }

        return ResponseBody.fromString(
          jsonEncode({'id': 'doc-1', 'name': 'Dr Updated'}),
          200,
          headers: {
            Headers.contentTypeHeader: [Headers.jsonContentType],
          },
        );
      };

      final response = await apiClient.patch<Map<String, dynamic>>(
        '/api/v1/doctors/doc-1',
        data: {'name': 'Dr Updated'},
      );

      expect(response.statusCode, 200);
      expect(response.data?['name'], 'Dr Updated');
      expect(fakeAdapter.capturedRequests.length, 3);
      expect(fakeAdapter.capturedRequests.last.method, 'PATCH');
      expect(
        fakeAdapter.capturedRequests.last.headers['Authorization'],
        'Bearer fresh-token',
      );
    });
  });

  group('ApiClient.delete', () {
    test('sends DELETE request and succeeds on 204 No Content', () async {
      fakeAdapter.handler = (options) async {
        return ResponseBody.fromString('', 204);
      };

      final response = await apiClient.delete<void>('/api/v1/teams/team-1');

      expect(response.statusCode, 204);
      expect(fakeAdapter.capturedRequests.length, 1);

      final req = fakeAdapter.capturedRequests.single;
      expect(req.method, 'DELETE');
      expect(req.path, '/api/v1/teams/team-1');
      expect(req.headers['Authorization'], 'Bearer valid-access-token');
    });

    test('refreshes token once on 401 and retries DELETE request', () async {
      fakeTokenStorage.storedAccessToken = 'expired-token';

      fakeAdapter.handler = (options) async {
        if (options.path == '/api/v1/auth/refresh') {
          return ResponseBody.fromString(
            jsonEncode({
              'access_token': 'fresh-token',
              'refresh_token': 'valid-refresh-token',
            }),
            200,
            headers: {
              Headers.contentTypeHeader: [Headers.jsonContentType],
            },
          );
        }

        if (options.headers['Authorization'] == 'Bearer expired-token') {
          return ResponseBody.fromString(
            jsonEncode({'detail': 'Expired'}),
            401,
            headers: {
              Headers.contentTypeHeader: [Headers.jsonContentType],
            },
          );
        }

        return ResponseBody.fromString('', 204);
      };

      final response = await apiClient.delete<void>('/api/v1/shifts/shift-1');

      expect(response.statusCode, 204);
      expect(fakeAdapter.capturedRequests.length, 3);
      expect(fakeAdapter.capturedRequests.last.method, 'DELETE');
      expect(
        fakeAdapter.capturedRequests.last.headers['Authorization'],
        'Bearer fresh-token',
      );
    });
  });
}
