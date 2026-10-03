import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/core/config/app_config.dart';

void main() {
  group('AppConfig Tests', () {
    test('apiBaseUrl returns valid non-empty string', () {
      final url = AppConfig.apiBaseUrl;
      expect(url, isNotEmpty);
    });

    test('apiBaseUri parses valid URI without throwing', () {
      final uri = AppConfig.apiBaseUri;
      expect(uri.host, isNotEmpty);
      expect(uri.scheme, anyOf('http', 'https'));
    });
  });
}
