import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/core/network/api_client.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/auth/data/dtos/auth_tokens_dto.dart';
import 'package:frontend/features/auth/domain/models/auth_tokens.dart';
import 'package:frontend/features/auth/data/dtos/user_dto.dart';
import 'package:frontend/features/auth/domain/models/user.dart';

final authRemoteDataSourceProvider = Provider<AuthRemoteDataSource>((ref) {
  return AuthRemoteDataSource(ref.watch(apiClientProvider));
});

class AuthRemoteDataSource {
  AuthRemoteDataSource(this._apiClient);

  final ApiClient _apiClient;

  Future<AuthTokens> login({
    required String username,
    required String password,
  }) async {
    final response = await _apiClient.post<Map<String, dynamic>>(
      '/auth/login',
      data: {'username': username, 'password': password},
      contentType: 'application/x-www-form-urlencoded',
      requiresAuth: false,
    );

    final data = response.data;
    if (data == null) {
      throw const ApiException(
        type: ApiErrorType.unknown,
        message: 'No response data received from login',
      );
    }

    return AuthTokensDto.fromJson(data).toDomain();
  }

  Future<User> signup({
    required String invitationToken,
    required String firstName,
    required String lastName,
    required String email,
    required String password,
  }) async {
    final response = await _apiClient.post<Map<String, dynamic>>(
      '/auth/signup',
      data: {
        'invitation_token': invitationToken,
        'first_name': firstName,
        'last_name': lastName,
        'email': email,
        'password': password,
      },
      requiresAuth: false,
    );

    final data = response.data;
    if (data == null) {
      throw const ApiException(
        type: ApiErrorType.unknown,
        message: 'No response data received from signup',
      );
    }

    return UserDto.fromJson(data).toDomain();
  }

  Future<AuthTokens> refresh({String? refreshToken, String? csrfToken}) async {
    final body = refreshToken != null
        ? {'refresh_token': refreshToken, 'csrf_token': ?csrfToken}
        : null;

    final response = await _apiClient.post<Map<String, dynamic>>(
      '/auth/refresh',
      data: body,
      requiresAuth: false,
    );

    final data = response.data;
    if (data == null) {
      throw const ApiException(
        type: ApiErrorType.unknown,
        message: 'No response data received from refresh',
      );
    }

    return AuthTokensDto.fromJson(data).toDomain();
  }

  Future<void> logout({String? refreshToken, String? csrfToken}) async {
    final body = refreshToken != null
        ? {'refresh_token': refreshToken, 'csrf_token': ?csrfToken}
        : null;

    final response = await _apiClient.post<Map<String, dynamic>>(
      '/auth/logout',
      data: body,
      requiresAuth: false,
    );

    final data = response.data;
    if (data == null) {
      throw const ApiException(
        type: ApiErrorType.unknown,
        message: 'No response data received from logout',
      );
    }
  }

  Future<User> getCurrentUser() async {
    final response = await _apiClient.get<Map<String, dynamic>>('/auth/me', requiresAuth: true);
    
    final data = response.data;
    if (data == null) {
      throw const ApiException(
        type: ApiErrorType.unknown,
        message: 'No response data received from auth/me',
      );
    }

    return UserDto.fromJson(data).toDomain();
  }
}
