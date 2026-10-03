import 'package:flutter/foundation.dart';

abstract final class AppConfig {
  static const _envApiBaseUrl = String.fromEnvironment('API_BASE_URL');

  static String get apiBaseUrl {
    if (_envApiBaseUrl.isNotEmpty) {
      return _envApiBaseUrl;
    }

    if (kIsWeb) {
      return Uri.base.origin;
    }

    return "http://127.0.0.1:8000";
  }

  static Uri get apiBaseUri {
    final uri = Uri.tryParse(apiBaseUrl);

    if (uri == null ||
        uri.host.isEmpty ||
        (uri.scheme != 'http' && uri.scheme != 'https')) {
      throw StateError('API_BASE_URL must be valid http or https URL.');
    }

    return uri;
  }
}
