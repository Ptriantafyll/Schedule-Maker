import 'package:frontend/features/auth/domain/models/auth_tokens.dart';

class AuthTokensDto {
  const AuthTokensDto({
    required this.accessToken,
    this.refreshToken,
    this.csrfToken,
    this.tokenType = 'bearer',
  });

  final String accessToken;
  final String tokenType;
  final String? refreshToken;
  final String? csrfToken;

  factory AuthTokensDto.fromJson(Map<String, dynamic> json) {
    final accessToken = json['access_token'];
    if (accessToken == null || accessToken is! String) {
      throw const FormatException(
        'Tokens JSON missing required "access_token"',
      );
    }

    final tokenType = (json['token_type'] as String?) ?? 'bearer';
    final refreshToken = json['refresh_token'] as String?;
    final csrfToken = json['csrf_token'] as String?;

    return AuthTokensDto(
      accessToken: accessToken,
      refreshToken: refreshToken,
      csrfToken: csrfToken,
      tokenType: tokenType,
    );
  }

  AuthTokens toDomain() {
    return AuthTokens(
      accessToken: accessToken,
      refreshToken: refreshToken,
      csrfToken: csrfToken,
      tokenType: tokenType,
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'access_token': accessToken,
      'refresh_token': refreshToken,
      'csrf_token': csrfToken,
      'token_type': tokenType,
    };
  }

  @override
  String toString() =>
      'AuthTokensDto(accessToken: [REDACTED], tokenType: $tokenType, refreshToken: ${refreshToken != null ? '[REDACTED]' : null}, csrfToken:${csrfToken != null ? '[REDACTED]' : null})';
}
