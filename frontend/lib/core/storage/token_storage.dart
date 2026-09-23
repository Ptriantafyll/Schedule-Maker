import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:flutter_secure_storage/flutter_secure_storage.dart';

final tokenStorageProvider = Provider<TokenStorage>((ref) {
  return TokenStorage(const FlutterSecureStorage());
});

final class TokenStorage {
  TokenStorage(this._storage);

  final FlutterSecureStorage _storage;
  static const _accessTokenKey = 'auth_access_token';
  static const _refreshTokenKey = 'auth_refresh_token';

  Future<void> saveTokens({
    required String accessToken,
    String? refreshToken,
  }) async {
    if (accessToken.isEmpty) {
      throw ArgumentError.value(
        accessToken,
        'accessToken'
        'Access Token must not be empty',
      );
    }

    await _storage.write(key: _accessTokenKey, value: accessToken);

    if (refreshToken == null || refreshToken.isEmpty) {
      await _storage.delete(key: _refreshTokenKey);
    }

    await _storage.write(key: _refreshTokenKey, value: refreshToken);
  }

  Future<String?> readAccessToken() {
    return _storage.read(key: _accessTokenKey);
  }

  Future<String?> readRefreshToken() {
    return _storage.read(key: _refreshTokenKey);
  }

  Future<void> clearTokens() async {
    await _storage.delete(key: _accessTokenKey);
    await _storage.delete(key: _refreshTokenKey);
  }
}
