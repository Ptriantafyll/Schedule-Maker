import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/auth/data/dtos/auth_tokens_dto.dart';
import 'package:frontend/features/auth/domain/models/auth_tokens.dart';

void main() {
  group('AuthTokensDto', () {
    const fullTokensJson = {
      'access_token': 'acc-secret-123',
      'token_type': 'bearer',
      'refresh_token': 'ref-secret-456',
      'csrf_token': 'csrf-secret-789',
    };

    const webTokensJson = {
      'access_token': 'acc-secret-web',
      'token_type': 'bearer',
      'refresh_token': null,
      'csrf_token': null,
    };

    group('fromJson', () {
      test('parses full tokens payload (native platform) correctly', () {
        final dto = AuthTokensDto.fromJson(fullTokensJson);

        expect(dto.accessToken, equals('acc-secret-123'));
        expect(dto.tokenType, equals('bearer'));
        expect(dto.refreshToken, equals('ref-secret-456'));
        expect(dto.csrfToken, equals('csrf-secret-789'));
      });

      test('parses web payload with null refresh and csrf tokens', () {
        final dto = AuthTokensDto.fromJson(webTokensJson);

        expect(dto.accessToken, equals('acc-secret-web'));
        expect(dto.tokenType, equals('bearer'));
        expect(dto.refreshToken, isNull);
        expect(dto.csrfToken, isNull);
      });

      test('defaults token_type to "bearer" when omitted', () {
        const minimalJson = {
          'access_token': 'acc-secret-minimal',
        };

        final dto = AuthTokensDto.fromJson(minimalJson);

        expect(dto.accessToken, equals('acc-secret-minimal'));
        expect(dto.tokenType, equals('bearer'));
        expect(dto.refreshToken, isNull);
        expect(dto.csrfToken, isNull);
      });

      test('throws FormatException when access_token is missing', () {
        const missingAccessJson = {
          'token_type': 'bearer',
          'refresh_token': 'ref-123',
        };

        expect(
          () => AuthTokensDto.fromJson(missingAccessJson),
          throwsA(isA<FormatException>()),
        );
      });

      test('throws FormatException when access_token is null or not a string', () {
        const invalidAccessJson = {
          'access_token': 12345,
        };

        expect(
          () => AuthTokensDto.fromJson(invalidAccessJson),
          throwsA(isA<FormatException>()),
        );
      });
    });

    group('toDomain', () {
      test('maps AuthTokensDto to a pure domain AuthTokens entity', () {
        final dto = AuthTokensDto.fromJson(fullTokensJson);
        final tokens = dto.toDomain();

        expect(tokens, isA<AuthTokens>());
        expect(
          tokens,
          equals(
            const AuthTokens(
              accessToken: 'acc-secret-123',
              refreshToken: 'ref-secret-456',
              csrfToken: 'csrf-secret-789',
              tokenType: 'bearer',
            ),
          ),
        );
      });
    });

    group('toJson', () {
      test('serializes back to backend snake_case format', () {
        final dto = AuthTokensDto.fromJson(fullTokensJson);
        final json = dto.toJson();

        expect(json['access_token'], equals('acc-secret-123'));
        expect(json['token_type'], equals('bearer'));
        expect(json['refresh_token'], equals('ref-secret-456'));
        expect(json['csrf_token'], equals('csrf-secret-789'));
      });
    });

    group('security redaction (toString)', () {
      test('never leaks raw secrets in toString', () {
        final dto = AuthTokensDto.fromJson(fullTokensJson);
        final stringified = dto.toString();

        expect(stringified, contains('[REDACTED]'));
        expect(stringified, isNot(contains('acc-secret-123')));
        expect(stringified, isNot(contains('ref-secret-456')));
        expect(stringified, isNot(contains('csrf-secret-789')));
      });
    });
  });
}
