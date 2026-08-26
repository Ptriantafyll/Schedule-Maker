import 'package:flutter/material.dart';

abstract final class AppTheme {
  static const _primaryLight = Color(0xFF005EB8);
  static const _secondaryLight = Color(0xFF00A3AD);
  static const _tertiaryLight = Color(0xFFDA291C);
  static const _neutralLight = Color(0xFFF3F9FC);
  static const _primaryDark = Color(0xFF3B82F6);
  static const _secondaryDark = Color(0xFF2DD4BF);
  static const _tertiaryDark = Color(0xFFF87171);
  static const _neutralDark = Color(0xFF0F172A);

  static final lightScheme = ColorScheme.fromSeed(
    seedColor: _primaryLight,
    brightness: Brightness.light,
    primary: _primaryLight,
    onPrimary: Colors.white,
    secondary: _secondaryLight,
    onSecondary: const Color(0xFF001F20),
    tertiary: _tertiaryLight,
    onTertiary: Colors.white,
    surface: _neutralLight,
    onSurface: Color(0xFF172B3A),
  );

  static final darkScheme = ColorScheme.fromSeed(
    seedColor: _primaryDark,
    brightness: Brightness.dark,
    primary: _primaryDark,
    onPrimary: _neutralDark,
    secondary: _secondaryDark,
    onSecondary: _neutralDark,
    tertiary: _tertiaryDark,
    onTertiary: _neutralDark,
    surface: _neutralDark,
    onSurface: const Color(0xFFE2E8F0),

    onSurfaceVariant: const Color(0xFFCBD5E1),
    outline: const Color(0xFF94A3B8),
    outlineVariant: const Color(0xFF475569),
  );

  static final dark = ThemeData(useMaterial3: true, colorScheme: darkScheme);
  static final light = ThemeData(useMaterial3: true, colorScheme: lightScheme);
}
