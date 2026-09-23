abstract final class AppConfig {
  static const apiBaseUrl = String.fromEnvironment('API_BASE_URL');

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
