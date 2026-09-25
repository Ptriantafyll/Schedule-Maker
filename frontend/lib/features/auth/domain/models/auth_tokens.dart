class AuthTokens {
  const AuthTokens({
    required this.accessToken,
    this.refreshToken,
    this.csrfToken,
    this.tokenType = 'bearer',
  });

  final String accessToken;
  final String? refreshToken;
  final String? csrfToken;
  final String tokenType;

  AuthTokens copyWith({
    String? accessToken,
    String? refreshToken,
    String? csrfToken,
    String? tokenType,
  }) {
    return AuthTokens(
      accessToken: accessToken ?? this.accessToken,
      refreshToken: refreshToken ?? this.refreshToken,
      csrfToken: csrfToken ?? this.csrfToken,
      tokenType: tokenType ?? this.tokenType,
    );
  }

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is AuthTokens &&
        other.accessToken == accessToken &&
        other.refreshToken == refreshToken &&
        other.csrfToken == csrfToken &&
        other.tokenType == tokenType;
  }

  @override
  int get hashCode =>
      Object.hash(accessToken, refreshToken, csrfToken, tokenType);

  @override
  String toString() =>
      'AuthTokens(accessToken: [REDACTED], refreshToken: ${refreshToken != null ? '[REDACTED]' : null}, csrfToken: ${csrfToken != null ? '[REDACTED]' : null}, tokenType: $tokenType)';
}
