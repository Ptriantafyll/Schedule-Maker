import 'package:dio/dio.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';

import '../config/app_config.dart';
import '../storage/token_storage.dart';
import 'api_exception.dart';

final apiClientProvider = Provider<ApiClient>((ref) {
  final tokenStorage = ref.watch(tokenStorageProvider);

  final dio = Dio(
    BaseOptions(
      baseUrl: AppConfig.apiBaseUrl.toString(),
      connectTimeout: const Duration(seconds: 15),
      receiveTimeout: const Duration(seconds: 15),
      sendTimeout: const Duration(seconds: 15),
      contentType: 'application/json',
      headers: const {'Accept': 'application/json'},
    ),
  );

  return ApiClient(dio, tokenStorage);
});

class ApiClient {
  ApiClient(this._dio, this._tokenStorage);

  final Dio _dio;
  final TokenStorage _tokenStorage;

  Future<Response<T>> get<T>(
    String path, {
    Map<String, dynamic>? queryParameters,
    bool requiresAuth = true,
  }) {
    return _send(() async {
      return _dio.get<T>(
        path,
        queryParameters: queryParameters,
        options: await _requestOptions(requiresAuth: requiresAuth),
      );
    });
  }

  Future<Response<T>> post<T>(
    String path, {
    Object? data,
    Map<String, dynamic>? queryParameters,
    String? contentType,
    bool requiresAuth = true,
  }) {
    return _send(() async {
      return _dio.post<T>(
        path,
        data: data,
        queryParameters: queryParameters,
        options: await _requestOptions(
          requiresAuth: requiresAuth,
          contentType: contentType,
        ),
      );
    });
  }

  Future<Options> _requestOptions({
    required bool requiresAuth,
    String? contentType,
  }) async {
    final headers = <String, dynamic>{};

    if (requiresAuth) {
      final accessToken = await _tokenStorage.readAccessToken();

      if (accessToken == null || accessToken.isEmpty) {
        throw const ApiException(
          type: ApiErrorType.unauthorized,
          message: 'You are not signed in.',
        );
      }

      headers['Authorization'] = 'Bearer $accessToken';
    }

    return Options(
      headers: headers.isNotEmpty ? headers : null,
      contentType: contentType,
    );
  }

  Future<Response<T>> _send<T>(Future<Response<T>> Function() request) async {
    try {
      return await request();
    } on DioException catch (exception) {
      throw _mapDioException(exception);
    }
  }

  ApiException _mapDioException(DioException exception) {
    final statusCode = exception.response?.statusCode;

    if (statusCode != null) {
      if (statusCode == 400 || statusCode == 422) {
        return ApiException(
          type: ApiErrorType.validation,
          message: 'Some information was not accepted. Check your entries',
          statusCode: statusCode,
        );
      }

      if (statusCode == 401) {
        return ApiException(
          type: ApiErrorType.unauthorized,
          message: 'Authentication was not accepted',
          statusCode: statusCode,
        );
      }

      if (statusCode == 403) {
        return ApiException(
          type: ApiErrorType.forbidden,
          message: 'You do not have permission to perform this action',
          statusCode: statusCode,
        );
      }

      if (statusCode == 404) {
        return ApiException(
          type: ApiErrorType.notFound,
          message: 'The requeseted information was not found',
          statusCode: statusCode,
        );
      }

      if (statusCode >= 500) {
        return ApiException(
          type: ApiErrorType.server,
          message: 'Server error. Please try again later',
          statusCode: statusCode,
        );
      }

      return ApiException(
        type: ApiErrorType.unknown,
        message: 'The request could not be completed',
        statusCode: statusCode,
      );
    }

    switch (exception.type) {
      case DioExceptionType.connectionTimeout:
      case DioExceptionType.receiveTimeout:
      case DioExceptionType.sendTimeout:
      case DioExceptionType.connectionError:
      case DioExceptionType.transformTimeout:
        return const ApiException(
          type: ApiErrorType.network,
          message: 'Unable to connect. Check your internet connection.',
        );
      case DioExceptionType.badCertificate:
        return const ApiException(
          type: ApiErrorType.network,
          message: 'Could not establish a secure connection.',
        );
      case DioExceptionType.cancel:
        return const ApiException(
          type: ApiErrorType.unknown,
          message: 'The request was cancelled.',
        );
      case DioExceptionType.badResponse:
      case DioExceptionType.unknown:
        return const ApiException(
          type: ApiErrorType.unknown,
          message: 'An unexpected error occured. Please try again.',
        );
    }
  }
}
