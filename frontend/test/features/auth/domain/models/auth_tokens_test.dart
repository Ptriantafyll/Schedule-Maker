import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/auth/domain/models/auth_tokens.dart';

void main() {
  group('AuthTokens', () {
    group('instantiation', () {
      test('creates AuthTokens with all fields populated', () {
        const tokens = AuthTokens(
          accessToken: 'acc-123',
          refreshToken: 'ref-456',
          csrfToken: 'csrf-789',
          tokenType: 'bearer',
        );

        expect(tokens.accessToken, equals('acc-123'));
        expect(tokens.refreshToken, equals('ref-456'));
        expect(tokens.csrfToken, equals('csrf-789'));
        expect(tokens.tokenType, equals('bearer'));
      });

      test('defaults tokenType to "bearer" and allows null refresh/csrf tokens', () {
        const tokens = AuthTokens(accessToken: 'acc-web');

        expect(tokens.accessToken, equals('acc-web'));
        expect(tokens.refreshToken, isNull);
        expect(tokens.csrfToken, isNull);
        expect(tokens.tokenType, equals('bearer'));
      });
    });

    group('equality', () {
      const tokens1 = AuthTokens(
        accessToken: 'acc-1',
        refreshToken: 'ref-1',
        csrfToken: 'csrf-1',
      );

      const tokens2 = AuthTokens(
        accessToken: 'acc-1',
        refreshToken: 'ref-1',
        csrfToken: 'csrf-1',
      );

      const differentTokens = AuthTokens(
        accessToken: 'acc-diff',
        refreshToken: 'ref-1',
      );

      test('two instances with identical fields are equal', () {
        expect(tokens1, equals(tokens2));
      });

      test('two instances with identical fields have same hashCode', () {
        expect(tokens1.hashCode, equals(tokens2.hashCode));
      });

      test('instances with different tokens are not equal', () {
        expect(tokens1, isNot(equals(differentTokens)));
      });
    });

    group('copyWith', () {
      const original = AuthTokens(
        accessToken: 'old-access',
        refreshToken: 'my-refresh',
        csrfToken: 'my-csrf',
      );

      test('updates accessToken while preserving refreshToken and csrfToken', () {
        final refreshed = original.copyWith(accessToken: 'new-access');

        expect(refreshed.accessToken, equals('new-access'));
        expect(refreshed.refreshToken, equals('my-refresh'));
        expect(refreshed.csrfToken, equals('my-csrf'));
        expect(refreshed.tokenType, equals('bearer'));
      });

      test('preserves all fields when no arguments passed', () {
        final copy = original.copyWith();

        expect(copy, equals(original));
      });
    });

    group('security redaction (toString)', () {
      const secretTokens = AuthTokens(
        accessToken: 'super-secret-access-token',
        refreshToken: 'super-secret-refresh-token',
        csrfToken: 'super-secret-csrf-token',
      );

      test('redacts raw tokens and never leaks secrets in toString', () {
        final stringified = secretTokens.toString();

        expect(stringified, contains('[REDACTED]'));
        expect(stringified, isNot(contains('super-secret-access-token')));
        expect(stringified, isNot(contains('super-secret-refresh-token')));
        expect(stringified, isNot(contains('super-secret-csrf-token')));
      });

      test('indicates null refresh token safely when absent', () {
        const webTokens = AuthTokens(accessToken: 'web-access-token');
        final stringified = webTokens.toString();

        expect(stringified, contains('refreshToken: null'));
        expect(stringified, isNot(contains('web-access-token')));
      });
    });
  });
}
