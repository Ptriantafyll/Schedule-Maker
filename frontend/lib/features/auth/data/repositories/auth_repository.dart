import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/core/storage/token_storage.dart';
import 'package:frontend/features/auth/data/datasources/auth_remote_data_source.dart';
import 'package:frontend/features/auth/domain/models/user.dart';

final authRepositoryProvider = Provider<AuthRepository>((ref) {
  return AuthRepositoryImpl(
    remoteDataSource: ref.watch(authRemoteDataSourceProvider),
    tokenStorage: ref.watch(tokenStorageProvider),
  );
});

abstract class AuthRepository {
  Future<User> login({required String email, required String password});

  Future<User> signup({
    required String email,
    required String password,
    required String firstName,
    required String lastName,
    required String invitationToken,
  });

  Future<void> logout();

  Future<User?> restoreSession();
}

class AuthRepositoryImpl implements AuthRepository {
  AuthRepositoryImpl({
    required this.remoteDataSource,
    required this.tokenStorage,
  });

  final AuthRemoteDataSource remoteDataSource;
  final TokenStorage tokenStorage;

  @override
  Future<User> login({required String email, required String password}) async {
    final authTokens = await remoteDataSource.login(
      username: email,
      password: password,
    );

    await tokenStorage.saveTokens(
      accessToken: authTokens.accessToken,
      refreshToken: authTokens.refreshToken,
    );

    final user = await remoteDataSource.getCurrentUser();
    return user;
  }

  @override
  Future<User> signup({
    required String email,
    required String password,
    required String firstName,
    required String lastName,
    required String invitationToken,
  }) async {
    final user = await remoteDataSource.signup(
      invitationToken: invitationToken,
      firstName: firstName,
      lastName: lastName,
      email: email,
      password: password,
    );
    return user;
  }

  @override
  Future<void> logout() async {
    final refreshToken = await tokenStorage.readRefreshToken();

    try {
      await remoteDataSource.logout(refreshToken: refreshToken);
    } finally {
      await tokenStorage.clearTokens();
    }
  }

  @override
  Future<User?> restoreSession() async {
    final accessToken = await tokenStorage.readAccessToken();
    if (accessToken == null || accessToken.isEmpty) {
      return null;
    }

    try {
      return await remoteDataSource.getCurrentUser();
    } on ApiException catch (e) {
      //refresh if unauthorized
      if (e.statusCode == 401 || e.type == ApiErrorType.unauthorized) {
        final refreshToken = await tokenStorage.readRefreshToken();
        try {
          final newTokens = await remoteDataSource.refresh(
            refreshToken: refreshToken,
          );
          await tokenStorage.saveTokens(
            accessToken: newTokens.accessToken,
            refreshToken: newTokens.refreshToken,
          );
          return await remoteDataSource.getCurrentUser();
        } on ApiException catch (refreshError) {
          if (refreshError.statusCode == 401 ||
              refreshError.type == ApiErrorType.unauthorized) {
            await tokenStorage.clearTokens();
            return null;
          }
          rethrow;
        }
      }
      rethrow;
    }
  }
}
